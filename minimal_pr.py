"""Minimal Gibbs equilibrium (Peng-Robinson EOS) using Cantera.

Data sources (all Cantera-bundled except H2O critical props):
  - gri30.yaml       : thermo for H2, CO, CO2, H2O, CH4, CH3OH
  - nasa_gas.yaml    : thermo for CH3OCH3 (L12/92, not in gri30)
  - graphite.yaml    : thermo for C(gr) solid
  - critical-properties.yaml : PR critical params (Tc, Pc, omega) for
      species that match by name. Uses molecular-formula naming
      (methanol = CH4O, DME = C2H6O).
  - H2O critical params: NIST WebBook / IAPWS-95, hardcoded here since
      critical-properties.yaml does not include water.
"""
from __future__ import annotations

import tempfile
import os
import yaml
import cantera as ct


# Species to include, grouped by where thermo comes from
THERMO_SOURCE = {
    "H2":      "gri30.yaml",
    "CO":      "gri30.yaml",
    "CO2":     "gri30.yaml",
    "H2O":     "gri30.yaml",
    "CH4":     "gri30.yaml",
    "CH3OH":   "gri30.yaml",
    "CH3OCH3": "nasa_gas.yaml",
    "C(gr)":   "graphite.yaml",
}

# Name aliases for critical-properties.yaml lookup
#   critical-properties.yaml uses molecular-formula names.
CRIT_ALIAS = {
    "CH3OH":   "CH4O",   # methanol
    "CH3OCH3": "C2H6O",  # methyl-ether (DME)
}

# Species NOT in Cantera's critical-properties.yaml — must provide manually
# Source: NIST WebBook (https://webbook.nist.gov) / IAPWS-95 for H2O
CRIT_OVERRIDE = {
    "H2O": {
        "critical-temperature": 647.096,   # K   (IAPWS-95)
        "critical-pressure":    22.064e6,  # Pa  (IAPWS-95)
        "acentric-factor":      0.3443,    # Reid, Prausnitz, Poling 5th ed.
    },
}


def _cantera_data_path(filename: str) -> str:
    """Return absolute path of a Cantera bundled data file."""
    import cantera as _ct
    base = os.path.dirname(_ct.__file__)
    return os.path.join(base, "data", filename)


def _build_yaml() -> str:
    """Assemble a temporary YAML combining thermo + PR critical params.

    Returns the path to the temp file.
    """
    # 1. Load critical-properties database
    with open(_cantera_data_path("critical-properties.yaml")) as f:
        crit_db = {s["name"]: s.get("critical-parameters", {})
                   for s in yaml.safe_load(f)["species"]}

    # 2. Load thermo for each species from its source yaml
    species_entries = []
    for sp, src in THERMO_SOURCE.items():
        with open(_cantera_data_path(src)) as f:
            src_doc = yaml.safe_load(f)
        entry = next(s for s in src_doc["species"] if s["name"] == sp)

        # Attach PR critical parameters (gas species only, not graphite)
        if sp != "C(gr)":
            crit = CRIT_OVERRIDE.get(sp) or crit_db.get(CRIT_ALIAS.get(sp, sp))
            if crit is None:
                raise RuntimeError(f"No critical properties for {sp}")
            entry["critical-parameters"] = {
                "critical-temperature": crit["critical-temperature"],
                "critical-pressure":    crit["critical-pressure"],
                "acentric-factor":      crit["acentric-factor"],
            }
            entry["equation-of-state"] = {"model": "Peng-Robinson"}
        species_entries.append(entry)

    # 3. Build combined YAML doc with two phases: PR gas + solid graphite
    doc = {
        "phases": [
            {
                "name": "gas",
                "thermo": "Peng-Robinson",
                "elements": ["C", "H", "O"],
                "species": [sp for sp in THERMO_SOURCE if sp != "C(gr)"],
                "state": {"T": 300, "P": 101325},
            },
            {
                "name": "graphite",
                "thermo": "fixed-stoichiometry",
                "elements": ["C"],
                "species": ["C(gr)"],
            },
        ],
        "species": species_entries,
    }

    tmp = tempfile.NamedTemporaryFile("w", suffix=".yaml", delete=False)
    yaml.dump(doc, tmp)
    tmp.close()
    return tmp.name


def equilibrate_pr(
    inlet_moles: dict[str, float],
    T_celsius: float,
    P_pascal: float,
    include_graphite: bool = False,
) -> dict[str, float]:
    """Solve Gibbs equilibrium with PR EOS."""
    path = _build_yaml()
    try:
        gas = ct.Solution(path, "gas")
        T_K = T_celsius + 273.15
        # set gas-phase inlet composition
        gas_inlet = {k: v for k, v in inlet_moles.items() if k != "C(gr)"}
        gas.TPX = T_K, P_pascal, gas_inlet

        if include_graphite:
            carbon = ct.Solution(path, "graphite")
            carbon.TP = T_K, P_pascal
            mix = ct.Mixture([gas, carbon])
            mix.T = T_K
            mix.P = P_pascal
            # seed moles (in kmol for Mixture)
            n_total = sum(inlet_moles.values())
            moles = [0.0] * mix.n_species
            for sp, v in inlet_moles.items():
                if sp == "C(gr)":
                    moles[mix.species_index(1, "C(gr)")] = v / 1000.0
                else:
                    moles[mix.species_index(0, sp)] = v / 1000.0
            mix.species_moles = moles
            mix.equilibrate("TP", solver="vcs", max_steps=500)

            result = {}
            for sp in inlet_moles:
                if sp == "C(gr)":
                    idx = mix.species_index(1, "C(gr)")
                else:
                    idx = mix.species_index(0, sp)
                result[sp] = mix.species_moles[idx] * 1000.0
        else:
            # Single-phase PR equilibrium
            q = ct.Quantity(gas, moles=sum(gas_inlet.values()) / 1000.0)
            q.equilibrate("TP")
            result = {sp: q.moles * q.X[gas.species_index(sp)] * 1000.0
                      for sp in gas_inlet}

        # Compressibility factor
        Z = gas.P * gas.mean_molecular_weight / (gas.density * ct.gas_constant * gas.T)
        result["_Z"] = Z
        return result
    finally:
        os.unlink(path)


if __name__ == "__main__":
    # Methanol synthesis: CO + 2 H2 ⇌ CH3OH at 250 °C, 5 MPa
    print("=== Methanol synthesis: H2:CO = 2:1, 250°C, 5 MPa, PR EOS ===")
    res = equilibrate_pr(
        inlet_moles={"H2": 2, "CO": 1, "CO2": 0, "H2O": 0, "CH4": 0, "CH3OH": 0, "CH3OCH3": 0},
        T_celsius=250,
        P_pascal=5e6,
    )
    print(f"  Z (compressibility) = {res.pop('_Z'):.4f}")
    total = sum(res.values())
    for sp, n in res.items():
        print(f"  {sp:9s} {n:8.5f} mol   ({n/total*100:6.2f} mol%)")
    print(f"  {'total':9s} {total:8.5f} mol")

    # DME synthesis: 3 H2 + 3 CO → CH3OCH3 + CO2 (adds DME as product)
    print("\n=== DME synthesis: H2:CO = 2:1, 200°C, 5 MPa, PR EOS ===")
    res = equilibrate_pr(
        inlet_moles={"H2": 2, "CO": 1, "CO2": 0, "H2O": 0, "CH4": 0,
                     "CH3OH": 0, "CH3OCH3": 0},
        T_celsius=200,
        P_pascal=5e6,
    )
    print(f"  Z (compressibility) = {res.pop('_Z'):.4f}")
    total = sum(res.values())
    for sp, n in res.items():
        print(f"  {sp:9s} {n:8.5f} mol   ({n/total*100:6.2f} mol%)")

    # Boudouard / coking: 2 CO → C(gr) + CO2 (with graphite solid phase)
    print("\n=== Boudouard with graphite: CO only, 600°C, 1 atm ===")
    res = equilibrate_pr(
        inlet_moles={"H2": 0, "CO": 1, "CO2": 0, "H2O": 0, "CH4": 0,
                     "CH3OH": 0, "CH3OCH3": 0, "C(gr)": 0},
        T_celsius=600,
        P_pascal=101325,
        include_graphite=True,
    )
    res.pop("_Z", None)
    for sp, n in res.items():
        print(f"  {sp:9s} {n:8.5f} mol")
