"""Minimal thermo-data fetcher for species missing from Cantera-bundled YAMLs.

Fetches:
  1. NASA7 polynomial coefficients (14 coefficients, low/high T ranges)
     → Burcat's Third Millennium Database (via Mutation++ GitHub mirror).
       Cached locally at ~/.cache/burcat_therm.dat on first call.
  2. Critical parameters (Tc, Pc, omega) for PR/RK EOS
     → `chemicals` library (Caleb Bell, MIT) — aggregates DIPPR, NIST
       WebBook, Yaws 2001, Poling-Prausnitz-Reid, PSRK, etc.

License note:
  Burcat data carries a non-commercial attribution license
  (https://respecth.elte.hu/burcat/READ.ME.txt). This script fetches
  at runtime and caches in the user's home directory — do NOT commit
  the cache to a repo. Proper citation:
    A. Burcat and B. Ruscic, ANL-05/20 TAE-960 (2005),
    "Third Millennium Ideal Gas and Condensed Phase Thermochemical
     Database for Combustion with Updates from Active Thermochemical
     Tables"

Usage:
    from fetch_thermo import fetch_species
    s = fetch_species("Meacetate", cas="79-20-9")
    print(s.cantera_yaml_entry())
"""
from __future__ import annotations

import re
import urllib.request
from dataclasses import dataclass
from pathlib import Path

BURCAT_URL = (
    "https://raw.githubusercontent.com/mutationpp/Mutationpp/"
    "master/data/thermo/burcat_therm.dat"
)
CACHE_PATH = Path.home() / ".cache" / "burcat_therm.dat"


def _fetch_burcat_text() -> str:
    """Download Burcat DB on first call, reuse cached copy thereafter."""
    if CACHE_PATH.exists():
        return CACHE_PATH.read_text(errors="replace")
    CACHE_PATH.parent.mkdir(parents=True, exist_ok=True)
    with urllib.request.urlopen(BURCAT_URL, timeout=60) as r:
        text = r.read().decode("utf-8", errors="replace")
    CACHE_PATH.write_text(text)
    return text


def _parse_composition(line1: str) -> dict[str, int]:
    """Parse element composition from columns 25-44 (4 × 5-char slots)."""
    comp: dict[str, int] = {}
    for slot in range(4):
        s = line1[24 + slot * 5 : 29 + slot * 5]
        el = s[:2].strip()
        n_str = s[2:].strip().rstrip(".")
        if not el or not n_str:
            continue
        n = int(float(n_str))
        if n > 0:
            comp[el] = n
    return comp


def _parse_nasa7(name: str, text: str) -> dict:
    """Extract one species' NASA7 block from Burcat text by exact name match."""
    lines = text.split("\n")
    for i, line in enumerate(lines):
        # CHEMKIN thermo entries have '1' in column 80
        if len(line) < 80 or line[79] != "1":
            continue
        # Burcat's name field is whitespace-delimited tokens in cols 0-18
        tokens = line[:24].split()
        if name not in tokens:
            continue

        comp = _parse_composition(line)
        Tlow = float(line[45:55])
        Thigh = float(line[55:65])
        Tmid = float(line[65:75]) if line[65:75].strip() else 1000.0

        def slice15(s: str, j: int) -> float:
            return float(s[j * 15 : (j + 1) * 15])

        l2, l3, l4 = lines[i + 1], lines[i + 2], lines[i + 3]
        a_high = [slice15(l2, j) for j in range(5)] + [slice15(l3, 0), slice15(l3, 1)]
        a_low = [slice15(l3, j) for j in range(2, 5)] + [slice15(l4, j) for j in range(4)]

        return {
            "burcat_name": name,
            "composition": comp,
            "temperature-ranges": [Tlow, Tmid, Thigh],
            "a_low": a_low,
            "a_high": a_high,
        }
    raise KeyError(f"'{name}' not found in Burcat database (cache: {CACHE_PATH})")


@dataclass
class SpeciesData:
    name: str
    burcat_name: str
    composition: dict[str, int]
    T_ranges: list[float]           # [Tlow, Tmid, Thigh]
    a_low: list[float]
    a_high: list[float]
    Tc: float | None = None         # K
    Pc: float | None = None         # Pa
    omega: float | None = None
    cas: str | None = None

    def cantera_yaml_entry(self) -> str:
        """Render as a Cantera species YAML entry (ready to paste into yaml)."""
        comp = ", ".join(f"{k}: {v}" for k, v in self.composition.items())
        fmt = lambda xs: "[" + ", ".join(f"{x:.9e}" for x in xs) + "]"
        lines = [
            f"- name: {self.name}",
            f"  composition: {{{comp}}}",
            f"  thermo:",
            f"    model: NASA7",
            f"    temperature-ranges: [{self.T_ranges[0]}, "
            f"{self.T_ranges[1]}, {self.T_ranges[2]}]",
            f"    data:",
            f"    - {fmt(self.a_low)}",
            f"    - {fmt(self.a_high)}",
            f"    note: 'Burcat ({self.burcat_name}); see Burcat & Ruscic ANL-05/20'",
        ]
        if self.Tc and self.Pc and self.omega is not None:
            lines += [
                f"  critical-parameters:",
                f"    critical-temperature: {self.Tc}",
                f"    critical-pressure: {self.Pc}",
                f"    acentric-factor: {self.omega}",
                f"  equation-of-state:",
                f"    model: Peng-Robinson",
            ]
        return "\n".join(lines) + "\n"


def fetch_species(
    burcat_name: str,
    cas: str | None = None,
    display_name: str | None = None,
) -> SpeciesData:
    """Fetch NASA7 (Burcat) + critical params (chemicals library).

    Parameters
    ----------
    burcat_name : the name as it appears in burcat_therm.dat (e.g. "Meacetate")
    cas         : CAS number for critical-property lookup. If omitted, Tc/Pc/omega
                  are left as None (useful when only NASA7 is needed).
    display_name: output species name (e.g. "CH3COOCH3"). Defaults to burcat_name.
    """
    nasa = _parse_nasa7(burcat_name, _fetch_burcat_text())

    Tc = Pc = omega = None
    if cas is not None:
        import chemicals
        Tc = chemicals.Tc(cas)
        Pc = chemicals.Pc(cas)
        omega = chemicals.omega(cas)

    return SpeciesData(
        name=display_name or burcat_name,
        burcat_name=nasa["burcat_name"],
        composition=nasa["composition"],
        T_ranges=nasa["temperature-ranges"],
        a_low=nasa["a_low"],
        a_high=nasa["a_high"],
        Tc=Tc, Pc=Pc, omega=omega, cas=cas,
    )


if __name__ == "__main__":
    import argparse

    ap = argparse.ArgumentParser(
        description="Fetch NASA7 (Burcat) + critical parameters (chemicals) "
                    "for a species, printed as a Cantera YAML entry."
    )
    ap.add_argument("burcat_name", help="Name as it appears in Burcat DB (e.g. 'Meacetate', 'H2O')")
    ap.add_argument("cas", nargs="?", default=None,
                    help="CAS number for Tc/Pc/omega lookup. Omit to skip critical params.")
    ap.add_argument("--name", default=None,
                    help="Output display name (default: same as burcat_name)")
    args = ap.parse_args()

    sp = fetch_species(args.burcat_name, cas=args.cas, display_name=args.name)
    print(sp.cantera_yaml_entry())
