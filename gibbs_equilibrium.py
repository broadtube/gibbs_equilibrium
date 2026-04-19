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
# Cantera同梱 nasa_gas.yaml / gri30.yaml にも無い化学種のNASA7多項式データ
#
# CH3OCH3, CH3COOH, C2H5OH, HCOOH は Cantera同梱の nasa_gas.yaml
# (NASA Glenn, McBride et al. NASA/TP-2002-211556) に収録されているため、
# build_gas() でそちらを優先的に読み込む。ここには収録されていない種のみを
# カスタムデータとして保持する。
# ---------------------------------------------------------------------------
CUSTOM_SPECIES_DATA = {
    # Methyl acetate: Burcat's Third Millennium Thermodynamic Database
    # エントリ "C3H6O2 Meacetate T10/07" より。
    # 出典: https://burcat.technion.ac.il/ (BURCAT.THR.txt)
    # ミラー: https://github.com/mutationpp/Mutationpp/blob/master/data/thermo/burcat_therm.dat
    #         https://github.com/OpenFOAM/OpenFOAM-2.2.x/blob/master/etc/thermoData/therm.dat
    # 参照: Felsmann et al. 2016, Proc. Combust. Inst. 36 (supplementary YAMLでも同一係数)
    "CH3COOCH3": {
        "composition": {"C": 3, "H": 6, "O": 2},
        "note": "Methyl acetate - Burcat T10/07",
        "temperature-ranges": [200.0, 1000.0, 6000.0],
        "data": [
            [7.18744749e+00, -6.29221513e-03, 8.17059377e-05, -9.82940778e-08,
             3.73744521e-11, -5.23417155e+04, -3.24161798e+00],
            [8.38776809e+00, 1.90836514e-02, -6.82197320e-06, 1.09765423e-09,
             -6.55561842e-14, -5.40805971e+04, -1.64156253e+01],
        ],
    },
}

# gri30の名前と一般的な名前のマッピング
SPECIES_ALIASES = {
    "HCHO": "CH2O",       # ホルムアルデヒド
    "CH2O": "CH2O",       # そのまま
}

# 固体炭素 (graphite) の種名。含まれていれば MultiPhase 計算に切り替える。
# thermo データは Cantera 同梱 graphite.yaml / nasa_condensed.yaml (JANAF X 4/83)。
CARBON_PHASE = "C(gr)"


# ---------------------------------------------------------------------------
# Cantera Solutionの構築
# ---------------------------------------------------------------------------
def build_gas(species: list[str]) -> ct.Solution:
    """指定された気相化学種を含むCantera Solutionを構築する。

    優先順位: gri30.yaml → nasa_gas.yaml (NASA Glenn) → CUSTOM_SPECIES_DATA。
    C(gr) は気相ではないため本関数からは除外される（呼び出し側で処理）。
    """
    gri30_species = {s.name: s for s in ct.Species.list_from_file("gri30.yaml")}
    nasa_species = {s.name: s for s in ct.Species.list_from_file("nasa_gas.yaml")}

    selected = []
    resolved_names = []  # Cantera内部での名前

    for sp in species:
        if sp == CARBON_PHASE:
            continue  # 固体相は別フェーズで扱う

        # エイリアス解決
        cantera_name = SPECIES_ALIASES.get(sp, sp)

        if cantera_name in gri30_species:
            selected.append(gri30_species[cantera_name])
            resolved_names.append(cantera_name)
        elif sp in nasa_species:
            selected.append(nasa_species[sp])
            resolved_names.append(sp)
        elif sp in CUSTOM_SPECIES_DATA:
            # カスタム種を構築。
            # Cantera NasaPoly2 の係数配列順序は [Tmid, high(7), low(7)]
            # （低温が先ではないので注意）。
            data = CUSTOM_SPECIES_DATA[sp]
            thermo = ct.NasaPoly2(
                data["temperature-ranges"][0],
                data["temperature-ranges"][2],
                ct.one_atm,
                np.concatenate([[data["temperature-ranges"][1]], data["data"][1], data["data"][0]]),
            )
            ct_sp = ct.Species(sp, data["composition"])
            ct_sp.thermo = thermo
            selected.append(ct_sp)
            resolved_names.append(sp)
        else:
            raise ValueError(
                f"化学種 '{sp}' はgri30.yaml / nasa_gas.yaml / カスタムデータに見つかりません。"
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
    has_carbon = CARBON_PHASE in species

    # 入口組成を設定 (species順)
    inlet_array = np.zeros(len(species))
    for i, sp in enumerate(species):
        inlet_array[i] = inlet_moles.get(sp, 0.0)

    total_inlet = inlet_array.sum()
    if total_inlet <= 0:
        raise ValueError("入口の総モル数が0以下です。")

    inlet_fractions = inlet_array / total_inlet

    outlet_moles = np.zeros(len(species))

    if not has_carbon:
        # 気相単一フェーズ (従来パス)
        composition = {rn: f for rn, f in zip(resolved_names, inlet_fractions) if f > 0}
        gas.TPX = T_kelvin, P_pascal, composition
        q = ct.Quantity(gas, moles=total_inlet / 1000.0)  # mol → kmol
        q.equilibrate("TP")

        for i, sp in enumerate(species):
            cantera_name = SPECIES_ALIASES.get(sp, sp)
            idx = gas.species_index(cantera_name)
            outlet_moles[i] = q.moles * q.X[idx] * 1000.0  # kmol → mol
    else:
        # 気相 + C(gr) 固体相の MultiPhase 平衡。
        # Cantera 同梱 graphite.yaml (固定化学量相、C(gr) 単独) を使用。
        carbon = ct.Solution("graphite.yaml")
        gas.TP = T_kelvin, P_pascal
        carbon.TP = T_kelvin, P_pascal
        mix = ct.Mixture([gas, carbon])
        mix.T = T_kelvin
        mix.P = P_pascal

        # 入口モル数を Mixture にセット (kmol単位)
        moles = np.zeros(mix.n_species)
        for sp, v in zip(species, inlet_array):
            if v <= 0:
                continue
            if sp == CARBON_PHASE:
                moles[mix.species_index(1, "C(gr)")] = v / 1000.0
            else:
                cantera_name = SPECIES_ALIASES.get(sp, sp)
                moles[mix.species_index(0, cantera_name)] = v / 1000.0
        mix.species_moles = moles

        # VCS algorithm は固体相の on/off を含む MultiPhase で 'gibbs' より堅牢。
        mix.equilibrate("TP", solver="vcs", max_steps=500)

        for i, sp in enumerate(species):
            if sp == CARBON_PHASE:
                idx = mix.species_index(1, "C(gr)")
            else:
                cantera_name = SPECIES_ALIASES.get(sp, sp)
                idx = mix.species_index(0, cantera_name)
            outlet_moles[i] = mix.species_moles[idx] * 1000.0  # kmol → mol

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
        species=["H2", "CO2", "CO", "H2O", "CH4", "C(gr)"],
        inlet_moles={"H2": 4, "CO2": 1},
        T_celsius=300,
        P="1atm",
    )
    print_result(result)

    # 使用例: メタン水蒸気改質
    # CH4 + H2O → CO + 3H2
    print("\n【例2】メタン水蒸気改質 (850°C, 2MPaG)")
    result = gibbs_equilibrium(
        species=["H2", "CO2", "CO", "H2O", "CH4", "C(gr)"],
        inlet_moles={"CH4": 1, "H2O": 3},
        T_celsius=850,
        P="2MPaG",
    )
    print_result(result)

    # 使用例: メタノール合成 (全化学種)
    # CO2 + 3H2 → CH3OH + H2O
    ALL_SPECIES = [
        "H2", "CO2", "CO", "C(gr)", "H2O", "CH4",
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
        species=["H2", "CO2", "CO", "H2O", "CH4", "C(gr)"],
        inlet_moles={"CO": 1, "H2O": 1},
        T_celsius=400,
        P="1atm",
    )
    print_result(result)
