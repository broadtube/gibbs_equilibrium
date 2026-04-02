"""ギブズ自由エネルギー最小化による化学平衡計算器。

Canteraを使用してギブズリアクターの平衡計算を行う。
入口の化学種・組成比・温度・圧力を指定し、出口の平衡組成を計算する。
"""

from __future__ import annotations

import re

import cantera as ct
import numpy as np


# ---------------------------------------------------------------------------
# 圧力パーサー (chemflow互換)
# ---------------------------------------------------------------------------
def parse_pressure(P) -> float:
    """圧力を Pa (絶対圧) に変換する。

    Parameters
    ----------
    P : float or str
        数値 → Pa (absolute)
        文字列 → "2MPaG", "2MPa", "10atm", "500kPa", "200kPaG" 等
    """
    ATM = 101325.0
    if isinstance(P, (int, float)):
        return float(P)

    s = str(P).strip()
    patterns = [
        (r"^([0-9.]+)\s*MPaG$", 1e6, ATM),
        (r"^([0-9.]+)\s*MPa$", 1e6, 0),
        (r"^([0-9.]+)\s*kPaG$", 1e3, ATM),
        (r"^([0-9.]+)\s*kPa$", 1e3, 0),
        (r"^([0-9.]+)\s*atm$", ATM, 0),
        (r"^([0-9.]+)\s*bar$", 1e5, 0),
    ]
    for pat, mult, offset in patterns:
        m = re.match(pat, s, re.IGNORECASE)
        if m:
            return float(m.group(1)) * mult + offset
    raise ValueError(f"Cannot parse pressure: '{P}'")


# ---------------------------------------------------------------------------
# gri30.yamlに無い化学種のNASA7多項式データ (NIST/Burcat)
# ---------------------------------------------------------------------------
CUSTOM_SPECIES_DATA = {
    "CH3OCH3": {
        "composition": {"C": 2, "H": 6, "O": 1},
        "note": "Dimethyl ether - Burcat",
        "temperature-ranges": [200.0, 1000.0, 6000.0],
        "data": [
            [5.46595502, -3.78307493e-03, 5.80104975e-05, -6.61704995e-08,
             2.36675037e-11, -2.51269095e+04, -7.41883739e-01],
            [5.52256973, 1.51516598e-02, -5.31000340e-06, 8.43166522e-10,
             -4.98976297e-14, -2.63076590e+04, -5.97699961e+00],
        ],
    },
    "CH3COOH": {
        "composition": {"C": 2, "H": 4, "O": 2},
        "note": "Acetic acid - Burcat",
        "temperature-ranges": [200.0, 1000.0, 6000.0],
        "data": [
            [3.40447203, 7.63843891e-03, 4.16989270e-05, -5.97498790e-08,
             2.37498760e-11, -5.34686310e+04, 9.95028030e+00],
            [6.33684770, 1.28429920e-02, -4.62915530e-06, 7.49795670e-10,
             -4.50063450e-14, -5.47041050e+04, -7.71503300e+00],
        ],
    },
    "CH3COOCH3": {
        "composition": {"C": 3, "H": 6, "O": 2},
        "note": "Methyl acetate - Burcat",
        "temperature-ranges": [200.0, 1000.0, 6000.0],
        "data": [
            [3.23511860, 1.39174410e-02, 5.23453750e-05, -7.82097280e-08,
             3.18429160e-11, -5.14816780e+04, 1.16177310e+01],
            [8.54553130, 1.63798170e-02, -5.82757620e-06, 9.34218900e-10,
             -5.56766010e-14, -5.35253830e+04, -1.93270130e+01],
        ],
    },
    "C2H5OH": {
        "composition": {"C": 2, "H": 6, "O": 1},
        "note": "Ethanol - Burcat",
        "temperature-ranges": [200.0, 1000.0, 6000.0],
        "data": [
            [4.85869570, -3.74017260e-03, 6.95554680e-05, -8.86547960e-08,
             3.51688360e-11, -2.99961320e+04, 4.80185450e+00],
            [6.56243650, 1.52042120e-02, -5.38967830e-06, 8.62524870e-10,
             -5.12299630e-14, -3.15254870e+04, -9.47302050e+00],
        ],
    },
    "HCOOH": {
        "composition": {"C": 1, "H": 2, "O": 2},
        "note": "Formic acid - Burcat",
        "temperature-ranges": [200.0, 1000.0, 6000.0],
        "data": [
            [1.43548185, 1.63363016e-02, -7.06729710e-06, -2.29787832e-09,
             2.06826580e-12, -4.64616264e+04, 1.74427640e+01],
            [4.61383160, 6.44963710e-03, -2.29082880e-06, 3.67147470e-10,
             -2.18738860e-14, -4.74299100e+04, 6.47163020e-01],
        ],
    },
}

# gri30の名前と一般的な名前のマッピング
SPECIES_ALIASES = {
    "HCHO": "CH2O",       # ホルムアルデヒド
    "CH2O": "CH2O",       # そのまま
}


# ---------------------------------------------------------------------------
# Cantera Solutionの構築
# ---------------------------------------------------------------------------
def build_gas(species: list[str]) -> ct.Solution:
    """指定された化学種を含むCantera Solutionを構築する。

    gri30.yamlに含まれる種はそこから取得し、含まれない種は
    カスタムNASA7データから構築する。
    """
    # gri30から利用可能な種を取得
    gri30_species = {s.name: s for s in ct.Species.list_from_file("gri30.yaml")}

    selected = []
    resolved_names = []  # Cantera内部での名前

    for sp in species:
        # エイリアス解決
        cantera_name = SPECIES_ALIASES.get(sp, sp)

        if cantera_name in gri30_species:
            selected.append(gri30_species[cantera_name])
            resolved_names.append(cantera_name)
        elif sp in CUSTOM_SPECIES_DATA:
            # カスタム種を構築
            data = CUSTOM_SPECIES_DATA[sp]
            thermo = ct.NasaPoly2(
                data["temperature-ranges"][0],
                data["temperature-ranges"][2],
                ct.one_atm,
                np.concatenate([[data["temperature-ranges"][1]], data["data"][0], data["data"][1]]),
            )
            ct_sp = ct.Species(sp, data["composition"])
            ct_sp.thermo = thermo
            selected.append(ct_sp)
            resolved_names.append(sp)
        else:
            raise ValueError(
                f"化学種 '{sp}' はgri30.yamlにもカスタムデータにも見つかりません。"
            )

    gas = ct.Solution(thermo="ideal-gas", species=selected)
    return gas, resolved_names


# ---------------------------------------------------------------------------
# ギブズ平衡計算
# ---------------------------------------------------------------------------
def gibbs_equilibrium(
    species: list[str],
    inlet_moles: dict[str, float],
    T_celsius: float,
    P: float | str = "1atm",
) -> dict:
    """ギブズ自由エネルギー最小化による平衡計算を行う。

    Parameters
    ----------
    species : list[str]
        考慮する化学種のリスト
    inlet_moles : dict[str, float]
        入口のモル比 (例: {"H2": 3, "CO2": 1})
    T_celsius : float
        温度 [°C]
    P : float or str
        圧力。数値ならPa、文字列なら "1atm", "2MPaG" 等

    Returns
    -------
    dict
        {
            "species": [...],
            "inlet_moles": [...],
            "inlet_mole_fractions": [...],
            "outlet_moles": [...],
            "outlet_mole_fractions": [...],
            "T_celsius": float,
            "T_kelvin": float,
            "P_pascal": float,
            "P_atm": float,
        }
    """
    T_kelvin = T_celsius + 273.15
    P_pascal = parse_pressure(P)

    gas, resolved_names = build_gas(species)

    # 入口組成を設定
    inlet_array = np.zeros(len(species))
    for i, sp in enumerate(species):
        inlet_array[i] = inlet_moles.get(sp, 0.0)

    total_inlet = inlet_array.sum()
    if total_inlet <= 0:
        raise ValueError("入口の総モル数が0以下です。")

    inlet_fractions = inlet_array / total_inlet

    # Cantera状態設定
    composition = {rn: f for rn, f in zip(resolved_names, inlet_fractions) if f > 0}
    gas.TPX = T_kelvin, P_pascal, composition

    # Quantity で平衡計算（モル数保存を追跡）
    q = ct.Quantity(gas, moles=total_inlet / 1000.0)  # mol → kmol
    q.equilibrate("TP")

    # 平衡後のモル流量を取得
    outlet_moles = np.zeros(len(species))
    for i, rn in enumerate(resolved_names):
        idx = gas.species_index(rn)
        outlet_moles[i] = q.moles * q.X[idx] * 1000.0  # kmol → mol

    outlet_total = outlet_moles.sum()
    outlet_fractions = outlet_moles / outlet_total if outlet_total > 0 else outlet_moles

    return {
        "species": species,
        "inlet_moles": inlet_array.tolist(),
        "inlet_mole_fractions": inlet_fractions.tolist(),
        "outlet_moles": outlet_moles.tolist(),
        "outlet_mole_fractions": outlet_fractions.tolist(),
        "T_celsius": T_celsius,
        "T_kelvin": T_kelvin,
        "P_pascal": P_pascal,
        "P_atm": P_pascal / 101325.0,
    }


# ---------------------------------------------------------------------------
# 結果表示
# ---------------------------------------------------------------------------
def print_result(result: dict) -> None:
    """平衡計算結果を表形式で出力する。"""
    species = result["species"]
    inlet_mol = result["inlet_moles"]
    inlet_frac = result["inlet_mole_fractions"]
    outlet_mol = result["outlet_moles"]
    outlet_frac = result["outlet_mole_fractions"]

    print("=" * 70)
    print(f"  ギブズ平衡計算結果")
    print(f"  温度: {result['T_celsius']:.1f} °C ({result['T_kelvin']:.2f} K)")
    print(f"  圧力: {result['P_atm']:.3f} atm ({result['P_pascal']:.0f} Pa)")
    print("=" * 70)
    print(f"  {'化学種':12s} {'入口mol':>10s} {'入口mol%':>10s} {'出口mol':>12s} {'出口mol%':>10s}")
    print("-" * 70)

    for i, sp in enumerate(species):
        in_m = inlet_mol[i]
        in_f = inlet_frac[i] * 100
        out_m = outlet_mol[i]
        out_f = outlet_frac[i] * 100
        print(f"  {sp:12s} {in_m:10.4f} {in_f:10.2f} {out_m:12.6f} {out_f:10.4f}")

    print("-" * 70)
    print(f"  {'合計':12s} {sum(inlet_mol):10.4f} {100.0:10.2f} {sum(outlet_mol):12.6f} {100.0:10.4f}")
    print("=" * 70)


# ---------------------------------------------------------------------------
# メイン (使用例)
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    # 使用例: CO2メタネーション (Sabatier反応)
    # CO2 + 4H2 → CH4 + 2H2O
    print("\n【例1】CO2メタネーション (300°C, 1atm)")
    result = gibbs_equilibrium(
        species=["H2", "CO2", "CO", "H2O", "CH4", "C"],
        inlet_moles={"H2": 4, "CO2": 1},
        T_celsius=300,
        P="1atm",
    )
    print_result(result)

    # 使用例: メタン水蒸気改質
    # CH4 + H2O → CO + 3H2
    print("\n【例2】メタン水蒸気改質 (850°C, 2MPaG)")
    result = gibbs_equilibrium(
        species=["H2", "CO2", "CO", "H2O", "CH4", "C"],
        inlet_moles={"CH4": 1, "H2O": 3},
        T_celsius=850,
        P="2MPaG",
    )
    print_result(result)

    # 使用例: メタノール合成 (全化学種)
    # CO2 + 3H2 → CH3OH + H2O
    ALL_SPECIES = [
        "H2", "CO2", "CO", "C", "H2O", "CH4",
        "CH3OH", "CH3OCH3", "CH3CHO", "CH3COOH", "CH3COOCH3",
        "HCHO", "C2H6", "C2H4", "C2H5OH", "HCOOH",
        "N2",
    ]
    print("\n【例3】メタノール合成系 - 全化学種 (250°C, 5MPa)")
    result = gibbs_equilibrium(
        species=ALL_SPECIES,
        inlet_moles={"H2": 3, "CO2": 1, "N2": 0.1},
        T_celsius=250,
        P="5MPa",
    )
    print_result(result)

    # 使用例: 水性ガスシフト反応
    # CO + H2O → CO2 + H2
    print("\n【例4】水性ガスシフト (400°C, 1atm)")
    result = gibbs_equilibrium(
        species=["H2", "CO2", "CO", "H2O", "CH4", "C"],
        inlet_moles={"CO": 1, "H2O": 1},
        T_celsius=400,
        P="1atm",
    )
    print_result(result)
