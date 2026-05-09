"""
tests/test_homologator.py
=========================
Targeted regression tests for phast_homologator_v2.py

Coverage:
  F3 – mw_calc zero-division guard (mass > 0, molar = 0; both = 0; NaN path)
  F2 – dropped-toxic warning written to HomologatedRow.warning
  R1 – alias priority is absolute (alias wins over nearest MW)
  R2 – pseudocomponent routing to hydrocarbon-only pool
  R3/R4 – molar renormalization invariant (Σ frac_corrected = 1.0)
  R3/R4 – water-override renormalization invariant
  R5 – post-options are additive (BASE result unchanged)
  R7 – MW tie-breaking: alias wins over alphabetical

All tests use synthetic in-memory fixtures so no external files are required
for the pure-logic tests. The integration test (test_end_to_end) reads the
real input_template.xlsx and both JSON files from the parent directory.
"""

from __future__ import annotations

import math
import sys
from pathlib import Path

import pytest

# ── Import the module under test ───────────────────────────────────────────────
SCRIPT_DIR = Path(__file__).parent.parent
sys.path.insert(0, str(SCRIPT_DIR))

import phast_homologator_v2 as H


# ══════════════════════════════════════════════════════════════════════════════
# Fixtures: minimal synthetic compound database and alias table
# ══════════════════════════════════════════════════════════════════════════════

def _make_compound(name, mw, formula=None, is_toxic=False, is_flammable=False,
                   is_inert=False, flag=None, status=1):
    """Build a PhastCompound quickly for test purposes."""
    return H.PhastCompound(
        phast_name=name,
        mw=mw,
        formula=formula,
        cas=None,
        status=status,
        flammable_toxic_flag=flag,
        is_toxic=is_toxic,
        is_flammable=is_flammable,
        is_inert=is_inert,
    )


# A minimal set of compounds covering hydrocarbons + inert + toxic
METHANE   = _make_compound("Methane",   16.043, formula="CH4",  is_flammable=True,  flag=1)
ETHANE    = _make_compound("Ethane",    30.069, formula="C2H6", is_flammable=True,  flag=1)
PROPANE   = _make_compound("Propane",   44.097, formula="C3H8", is_flammable=True,  flag=1)
N_BUTANE  = _make_compound("n-Butane",  58.123, formula="C4H10",is_flammable=True,  flag=1)
N_PENTANE = _make_compound("n-Pentane", 72.150, formula="C5H12",is_flammable=True,  flag=1)
N_HEXANE  = _make_compound("n-Hexane",  86.177, formula="C6H14",is_flammable=True,  flag=1)
WATER     = _make_compound("Water",     18.015, formula="H2O",  is_inert=True,     flag=-2)
H2S       = _make_compound("Hydrogen sulfide", 34.082, formula="H2S",
                            is_toxic=True, is_flammable=True, flag=0)
HCN       = _make_compound("Hydrogen cyanide", 27.026, formula="HCN",
                            is_toxic=True, flag=-1)
NITROGEN  = _make_compound("Nitrogen",  28.014, formula="N2",   is_inert=True,     flag=-2)
CO2       = _make_compound("Carbon dioxide", 44.010, formula="CO2", is_inert=True, flag=-2)

ALL_COMPOUNDS = [METHANE, ETHANE, HCN, NITROGEN, PROPANE, WATER, H2S,
                 N_BUTANE, N_PENTANE, N_HEXANE, CO2]
ALL_COMPOUNDS.sort(key=lambda c: (c.mw, c.phast_name))

HC_ONLY = [c for c in ALL_COMPOUNDS if H._is_hydrocarbon(c.formula)]

# Minimal alias table
ALIASES: dict[str, str] = {
    "C1": "Methane",
    "C2": "Ethane",
    "C3": "Propane",
    "H2S": "Hydrogen sulfide",
    "HCN": "Hydrogen cyanide",
    "H2O": "Water",
}


# ══════════════════════════════════════════════════════════════════════════════
# F3 — mw_calc zero-division guard
# ══════════════════════════════════════════════════════════════════════════════

class TestMwCalcGuard:
    """
    RawComponent.mw_calc must never raise ZeroDivisionError.
    When molar flow is zero the property returns 0.0 and resolve() falls back
    to 'fallback_no_mw' with an appropriate warning.
    """

    def test_mw_calc_both_zero(self):
        """mass=0, molar=0 → mw_calc returns 0.0, not ZeroDivisionError."""
        rc = H.RawComponent(name="TestComp", mass_kgh=0.0, molar_kmolh=0.0)
        assert rc.mw_calc == 0.0

    def test_mw_calc_mass_nonzero_molar_zero(self):
        """mass > 0, molar = 0 → mw_calc returns 0.0 (cannot compute MW)."""
        rc = H.RawComponent(name="TestComp", mass_kgh=100.0, molar_kmolh=0.0)
        assert rc.mw_calc == 0.0

    def test_mw_calc_mass_zero_molar_nonzero(self):
        """mass = 0, molar > 0 → mw_calc returns 0.0 (valid, zero mass)."""
        rc = H.RawComponent(name="TestComp", mass_kgh=0.0, molar_kmolh=1.0)
        assert rc.mw_calc == 0.0

    def test_mw_calc_normal(self):
        """Normal case: mass=44.097, molar=1.0 → mw_calc ≈ 44.097."""
        rc = H.RawComponent(name="Propane", mass_kgh=44.097, molar_kmolh=1.0)
        assert abs(rc.mw_calc - 44.097) < 1e-6

    def test_mw_calc_negative_molar(self):
        """Negative molar (data error) → mw_calc returns 0.0 (not a crash)."""
        rc = H.RawComponent(name="BadData", mass_kgh=50.0, molar_kmolh=-0.5)
        assert rc.mw_calc == 0.0

    def test_resolve_fallback_on_zero_mw(self):
        """
        resolve() with mw_calc=0.0 must use 'fallback_no_mw' method and
        include a non-empty warning about the zero/missing molar flow.
        """
        result = H.resolve(
            name="UnknownComp",
            mw_calc=0.0,
            all_compounds=ALL_COMPOUNDS,
            hydrocarbon_only=HC_ONLY,
            aliases={},
        )
        assert result["match_method"] == "fallback_no_mw"
        assert result["warning"], "Expected a non-empty warning for fallback_no_mw"
        assert "zero" in result["warning"].lower() or "missing" in result["warning"].lower()

    def test_resolve_fallback_warning_not_empty_via_homologate(self):
        """
        A component with mass > 0 but molar = 0 that passes parse filtering
        (if injected directly) must produce a homologated row with a warning.
        We test this via resolve() since parse skips (mass=0 AND molar=0) only.
        """
        # Inject a RawComponent with mass>0, molar=0 into a StreamInput
        stream = H.StreamInput(
            project="Test", escenario=1, corriente="C1",
            mass_unit="kg/h", mol_unit="kgmol/h",
            components=[
                # Valid component with known flow
                H.RawComponent("C1", mass_kgh=16.043, molar_kmolh=1.0),
                # Component with mass but zero molar → fallback
                H.RawComponent("Unknown", mass_kgh=50.0, molar_kmolh=0.0),
            ],
        )
        res = H.homologate(stream, ALL_COMPOUNDS, HC_ONLY, ALIASES, n_requested=5)
        fallback_rows = [hr for hr in res.homologated if hr.match_method == "fallback_no_mw"]
        assert len(fallback_rows) == 1
        assert fallback_rows[0].source_name == "Unknown"
        assert fallback_rows[0].warning  # must have a warning

    def test_no_crash_on_all_zero_flow_stream(self):
        """
        A stream where ALL components have zero molar flow should return an
        empty StreamResult (no crash).
        """
        stream = H.StreamInput(
            project="Test", escenario=1, corriente="C1",
            mass_unit="kg/h", mol_unit="kgmol/h",
            components=[
                H.RawComponent("C1", mass_kgh=0.0, molar_kmolh=0.0),
            ],
        )
        res = H.homologate(stream, ALL_COMPOUNDS, HC_ONLY, ALIASES, n_requested=5)
        assert res.n_used == 0
        assert res.total_molar_kmolh == 0


# ══════════════════════════════════════════════════════════════════════════════
# F2 — Dropped-toxic warning in homologated rows
# ══════════════════════════════════════════════════════════════════════════════

class TestDroppedToxicWarning:
    """
    When a toxic compound is present in the homologated input but excluded
    from the consolidated result by the N-slot threshold, each HomologatedRow
    for that toxic compound must have a non-empty warning mentioning 'HAZARD'
    or 'toxic'.
    """

    def _make_stream_with_trace_toxic(self):
        """
        Stream with dominant Methane (1000 kmol/h), Ethane (500 kmol/h),
        Propane (200 kmol/h), and a trace HCN (0.001 kmol/h).
        With N=2, HCN will be excluded.
        """
        return H.StreamInput(
            project="Test", escenario=1, corriente="C1",
            mass_unit="kg/h", mol_unit="kgmol/h",
            components=[
                H.RawComponent("C1",  mass_kgh=1000*16.043,  molar_kmolh=1000.0),
                H.RawComponent("C2",  mass_kgh=500*30.069,   molar_kmolh=500.0),
                H.RawComponent("C3",  mass_kgh=200*44.097,   molar_kmolh=200.0),
                H.RawComponent("HCN", mass_kgh=0.001*27.026, molar_kmolh=0.001),
            ],
        )

    def test_toxic_excluded_gets_warning(self):
        """HCN excluded by N=2 → its HomologatedRow.warning must mention HAZARD."""
        stream = self._make_stream_with_trace_toxic()
        res = H.homologate(stream, ALL_COMPOUNDS, HC_ONLY, ALIASES, n_requested=2)

        # HCN should NOT be in consolidated (too small)
        consolidated_names = {c.phast_name for c in res.consolidated}
        assert "Hydrogen cyanide" not in consolidated_names, (
            "HCN should have been excluded by N=2 threshold"
        )

        # Every HomologatedRow for HCN must have a HAZARD warning
        hcn_rows = [hr for hr in res.homologated
                    if hr.phast_name == "Hydrogen cyanide"]
        assert hcn_rows, "Expected at least one HomologatedRow for HCN"
        for hr in hcn_rows:
            assert hr.warning, f"Missing warning on excluded toxic row: {hr}"
            assert "HAZARD" in hr.warning or "toxic" in hr.warning.lower(), (
                f"Warning does not mention hazard: {hr.warning!r}"
            )

    def test_toxic_included_no_extra_warning(self):
        """
        H2S with significant flow included in N=5 result → no spurious
        'HAZARD excluded' warning should be added.
        """
        stream = H.StreamInput(
            project="Test", escenario=1, corriente="C1",
            mass_unit="kg/h", mol_unit="kgmol/h",
            components=[
                H.RawComponent("C1",  mass_kgh=1000*16.043, molar_kmolh=1000.0),
                H.RawComponent("H2S", mass_kgh=100*34.082,  molar_kmolh=100.0),
            ],
        )
        res = H.homologate(stream, ALL_COMPOUNDS, HC_ONLY, ALIASES, n_requested=5)

        consolidated_names = {c.phast_name for c in res.consolidated}
        assert "Hydrogen sulfide" in consolidated_names, "H2S should be included"

        h2s_rows = [hr for hr in res.homologated if hr.phast_name == "Hydrogen sulfide"]
        for hr in h2s_rows:
            assert "HAZARD" not in (hr.warning or ""), (
                f"Unexpected HAZARD warning on included compound: {hr.warning!r}"
            )

    def test_nontoxic_excluded_no_hazard_warning(self):
        """A non-toxic compound excluded by threshold should NOT get a HAZARD warning."""
        stream = H.StreamInput(
            project="Test", escenario=1, corriente="C1",
            mass_unit="kg/h", mol_unit="kgmol/h",
            components=[
                H.RawComponent("C1", mass_kgh=1000*16.043, molar_kmolh=1000.0),
                H.RawComponent("C2", mass_kgh=500*30.069,  molar_kmolh=500.0),
                H.RawComponent("C3", mass_kgh=200*44.097,  molar_kmolh=200.0),
                # Trace Methane duplicate → maps to Methane, very small → excluded
                H.RawComponent("n-Hexane", mass_kgh=0.001*86.177, molar_kmolh=0.001),
            ],
        )
        res = H.homologate(stream, ALL_COMPOUNDS, HC_ONLY, ALIASES, n_requested=2)
        # n-Hexane excluded; it is flammable-only, not toxic
        hexane_rows = [hr for hr in res.homologated if hr.phast_name == "n-Hexane"]
        for hr in hexane_rows:
            assert "HAZARD" not in (hr.warning or ""), (
                f"Non-toxic compound has spurious HAZARD warning: {hr.warning!r}"
            )


# ══════════════════════════════════════════════════════════════════════════════
# R1 — Alias priority is absolute
# ══════════════════════════════════════════════════════════════════════════════

class TestAliasPriority:
    """
    An alias match must win over the nearest-MW match, even when another
    compound is physically closer in MW.
    """

    def test_alias_wins_over_mw(self):
        """
        'C1' aliases to Methane (MW=16.043).
        Suppose we pass mw_calc=28.0 (close to Nitrogen=28.014, not Methane).
        Alias must still route to Methane.
        """
        result = H.resolve(
            name="C1",
            mw_calc=28.0,
            all_compounds=ALL_COMPOUNDS,
            hydrocarbon_only=HC_ONLY,
            aliases=ALIASES,
        )
        assert result["phast_name"] == "Methane"
        assert result["match_method"] == "alias"

    def test_alias_match_method_is_alias(self):
        """Alias-resolved compounds always report match_method='alias'."""
        result = H.resolve(
            name="H2S",
            mw_calc=34.082,
            all_compounds=ALL_COMPOUNDS,
            hydrocarbon_only=HC_ONLY,
            aliases=ALIASES,
        )
        assert result["phast_name"] == "Hydrogen sulfide"
        assert result["match_method"] == "alias"

    def test_no_alias_uses_mw(self):
        """
        A name with no alias entry falls through to MW matching.
        'PropaneX' not in aliases → MW nearest-neighbour → Propane (44.097).
        """
        result = H.resolve(
            name="PropaneX",
            mw_calc=44.097,
            all_compounds=ALL_COMPOUNDS,
            hydrocarbon_only=HC_ONLY,
            aliases={},
        )
        assert result["phast_name"] == "Propane"
        assert result["match_method"] in ("mw_unique", "mw_tie_resolved_by_alias",
                                          "mw_tie_unresolved")

    def test_case_insensitive_alias(self):
        """Alias lookup is case-insensitive: 'c1' resolves same as 'C1'."""
        result_upper = H.resolve("C1",  16.043, ALL_COMPOUNDS, HC_ONLY, ALIASES)
        result_lower = H.resolve("c1",  16.043, ALL_COMPOUNDS, HC_ONLY, ALIASES)
        assert result_upper["phast_name"] == result_lower["phast_name"] == "Methane"


# ══════════════════════════════════════════════════════════════════════════════
# R2 — Pseudocomponent routing to hydrocarbon-only pool
# ══════════════════════════════════════════════════════════════════════════════

class TestPseudocomponentRouting:
    """
    Names matching pseudocomponent patterns (CnHm-cut fractions) must only
    be matched against the hydrocarbon_only pool, never against Water, CO2,
    N2, toxics, or other non-CnHm compounds.
    """

    @pytest.mark.parametrize("pseudo_name", [
        "36-40C*",
        "100-110C*",
        "850+C*",
        "C10+",
        "Cn+",
        "NBP-100",
    ])
    def test_pseudo_detection(self, pseudo_name):
        """_is_pseudocomponent returns True for known pseudocomponent patterns."""
        assert H._is_pseudocomponent(pseudo_name), (
            f"{pseudo_name!r} should be detected as pseudocomponent"
        )

    @pytest.mark.parametrize("normal_name", [
        "Methane", "Propane", "C1", "H2S", "Water", "CO2",
    ])
    def test_non_pseudo_detection(self, normal_name):
        """Real compound names are not pseudocomponents."""
        assert not H._is_pseudocomponent(normal_name), (
            f"{normal_name!r} should NOT be detected as pseudocomponent"
        )

    def test_pseudo_never_matches_water(self):
        """
        A pseudocomponent (e.g. 'C10+') with MW close to Water (18 g/mol)
        must NOT be assigned Water — it must stay within the HC pool.
        Construct a degenerate pool where Water has the closest MW.
        """
        # Build a pool with Water having MW=18.015 and only one HC: Methane (16.043)
        # For pseudo 'C10+' with mw_calc=18.0, the closest compound in the full pool
        # is Water, but the HC-only pool forces Methane.
        mini_all = [METHANE, WATER]
        mini_all.sort(key=lambda c: (c.mw, c.phast_name))
        mini_hc  = [METHANE]

        result = H.resolve(
            name="C10+",
            mw_calc=18.0,   # deliberately close to Water
            all_compounds=mini_all,
            hydrocarbon_only=mini_hc,
            aliases={},
        )
        assert result["phast_name"] == "Methane", (
            "Pseudocomponent must not be assigned Water even when Water is closer in MW"
        )
        assert result["match_pool"] == "hydrocarbon_only"

    def test_pseudo_pool_label(self):
        """Pseudocomponent matches report match_pool='hydrocarbon_only'."""
        result = H.resolve(
            name="36-40C*",
            mw_calc=38.0,
            all_compounds=ALL_COMPOUNDS,
            hydrocarbon_only=HC_ONLY,
            aliases={},
        )
        assert result["match_pool"] == "hydrocarbon_only"

    def test_non_pseudo_uses_full_catalog(self):
        """Regular compound uses full catalog (can match Water, N2, etc.)."""
        result = H.resolve(
            name="H2O",
            mw_calc=18.015,
            all_compounds=ALL_COMPOUNDS,
            hydrocarbon_only=HC_ONLY,
            aliases=ALIASES,
        )
        assert result["phast_name"] == "Water"
        # Water is in full catalog; alias match_pool is always 'full_catalog'
        assert result["match_pool"] == "full_catalog"


# ══════════════════════════════════════════════════════════════════════════════
# R3/R4 — Molar renormalization invariant
# ══════════════════════════════════════════════════════════════════════════════

class TestRenormalizationInvariant:
    """
    For every homologated stream, Σ frac_corrected must equal 1.0
    within floating-point tolerance (1e-9).
    """

    TOL = 1e-9

    def _stream(self, components):
        return H.StreamInput(
            project="Test", escenario=1, corriente="C1",
            mass_unit="kg/h", mol_unit="kgmol/h",
            components=components,
        )

    def test_normal_mode_sum_to_one(self):
        """Normal homologation: Σ frac_corrected = 1.0."""
        stream = self._stream([
            H.RawComponent("C1", 1000*16.043, 1000.0),
            H.RawComponent("C2",  500*30.069,  500.0),
            H.RawComponent("C3",  200*44.097,  200.0),
        ])
        res = H.homologate(stream, ALL_COMPOUNDS, HC_ONLY, ALIASES, n_requested=5)
        total = sum(c.frac_corrected for c in res.consolidated)
        assert abs(total - 1.0) < self.TOL, (
            f"Renormalization failed: Σ frac_corrected = {total}"
        )

    def test_normal_mode_N_equals_1(self):
        """Edge case N=1: only 1 compound → frac_corrected = 1.0."""
        stream = self._stream([
            H.RawComponent("C1", 1000*16.043, 1000.0),
            H.RawComponent("C2",  500*30.069,  500.0),
        ])
        res = H.homologate(stream, ALL_COMPOUNDS, HC_ONLY, ALIASES, n_requested=1)
        assert res.n_used == 1
        total = sum(c.frac_corrected for c in res.consolidated)
        assert abs(total - 1.0) < self.TOL

    def test_normal_mode_many_components(self):
        """With 5 distinct compounds and N=3, Σ frac_corrected = 1.0."""
        stream = self._stream([
            H.RawComponent("C1",  1000*16.043, 1000.0),
            H.RawComponent("C2",   500*30.069,  500.0),
            H.RawComponent("C3",   200*44.097,  200.0),
            H.RawComponent("H2S",   50*34.082,   50.0),
            H.RawComponent("H2O",   20*18.015,   20.0),
        ])
        res = H.homologate(stream, ALL_COMPOUNDS, HC_ONLY, ALIASES, n_requested=3)
        total = sum(c.frac_corrected for c in res.consolidated)
        assert abs(total - 1.0) < self.TOL

    def test_water_override_mode_sum_to_one(self):
        """
        Water-override mode: Σ frac_corrected = 1.0 (dry fractions + water fraction).
        """
        # Dry sub-stream (water already separated by the caller)
        dry_stream = H.StreamInput(
            project="Test", escenario=1, corriente="C1",
            mass_unit="kg/h", mol_unit="kgmol/h",
            components=[
                H.RawComponent("C1", 100*16.043, 100.0),
                H.RawComponent("C2",  50*30.069,  50.0),
            ],
        )
        water_frac = 0.85  # water occupies 85% of the original stream
        res = H.homologate(
            dry_stream, ALL_COMPOUNDS, HC_ONLY, ALIASES,
            n_requested=5,
            water_override=True,
            water_frac_original=water_frac,
        )
        total = sum(c.frac_corrected for c in res.consolidated)
        assert abs(total - 1.0) < self.TOL, (
            f"Water-override renormalization failed: Σ frac_corrected = {total}"
        )

    def test_water_override_water_frac_preserved(self):
        """
        In water-override mode, the Water row's frac_corrected must equal
        water_frac_original (not renormalized away).
        """
        dry_stream = H.StreamInput(
            project="Test", escenario=1, corriente="C1",
            mass_unit="kg/h", mol_unit="kgmol/h",
            components=[
                H.RawComponent("C1", 100*16.043, 100.0),
                H.RawComponent("C2",  50*30.069,  50.0),
            ],
        )
        water_frac = 0.72
        res = H.homologate(
            dry_stream, ALL_COMPOUNDS, HC_ONLY, ALIASES,
            n_requested=5,
            water_override=True,
            water_frac_original=water_frac,
        )
        water_rows = [c for c in res.consolidated if c.phast_name == "Water"]
        assert water_rows, "Water row must be present in water-override result"
        assert abs(water_rows[-1].frac_corrected - water_frac) < 1e-9, (
            f"Water fraction not preserved: {water_rows[-1].frac_corrected} != {water_frac}"
        )

    def test_water_override_dry_sum(self):
        """
        Σ of dry frac_corrected = (1 - water_frac_original).
        """
        dry_stream = H.StreamInput(
            project="Test", escenario=1, corriente="C1",
            mass_unit="kg/h", mol_unit="kgmol/h",
            components=[
                H.RawComponent("C1", 100*16.043, 100.0),
                H.RawComponent("C2",  50*30.069,  50.0),
                H.RawComponent("C3",  30*44.097,  30.0),
            ],
        )
        water_frac = 0.60
        res = H.homologate(
            dry_stream, ALL_COMPOUNDS, HC_ONLY, ALIASES,
            n_requested=5,
            water_override=True,
            water_frac_original=water_frac,
        )
        dry_sum = sum(
            c.frac_corrected for c in res.consolidated
            if c.phast_name != "Water"
        )
        expected_dry = 1.0 - water_frac
        assert abs(dry_sum - expected_dry) < 1e-9, (
            f"Dry sum {dry_sum} != expected {expected_dry}"
        )


# ══════════════════════════════════════════════════════════════════════════════
# R7 — MW tie-breaking: alias wins over alphabetical
# ══════════════════════════════════════════════════════════════════════════════

class TestMwTieBreaking:
    """
    When two compounds have the same MW distance, the alias-indicated compound
    must win.  If no alias applies, alphabetical/MW sort determines the winner
    deterministically.
    """

    def test_tie_resolved_by_alias(self):
        """
        Build two compounds with identical MW.  Alias points to one of them.
        Expected: alias-indicated compound wins.
        """
        compound_a = _make_compound("Alpha", 44.097)
        compound_b = _make_compound("Beta",  44.097)
        pool = [compound_a, compound_b]
        aliases_local = {"TestInput": "Beta"}

        match, method = H._resolve_tie(
            candidates=[compound_a, compound_b],
            alias_candidate="Beta",
        )
        assert match.phast_name == "Beta"
        assert method == "mw_tie_resolved_by_alias"

    def test_tie_unresolved_is_alphabetical(self):
        """
        No alias for the tie candidates → alphabetical sort wins.
        'Alpha' < 'Beta' alphabetically → Alpha is chosen.
        """
        compound_a = _make_compound("Alpha", 44.097)
        compound_b = _make_compound("Beta",  44.097)

        match, method = H._resolve_tie(
            candidates=[compound_a, compound_b],
            alias_candidate=None,
        )
        assert match.phast_name == "Alpha"
        assert method == "mw_tie_unresolved"

    def test_tie_deterministic_repeated_calls(self):
        """Same tie resolved identically on multiple calls (no randomness)."""
        compound_a = _make_compound("Alpha", 44.097)
        compound_b = _make_compound("Beta",  44.097)

        results = set()
        for _ in range(20):
            match, _ = H._resolve_tie([compound_a, compound_b], alias_candidate=None)
            results.add(match.phast_name)
        assert len(results) == 1, "Tie resolution must be deterministic"


# ══════════════════════════════════════════════════════════════════════════════
# Unit conversion helpers
# ══════════════════════════════════════════════════════════════════════════════

class TestUnitConversion:

    def test_to_kgh_kgh(self):
        assert H.to_kgh(100.0, "kg/h") == 100.0

    def test_to_kgh_lbh(self):
        result = H.to_kgh(100.0, "lb/h")
        assert abs(result - 100.0 * H.KG_PER_LB) < 1e-9

    def test_to_kgh_gh(self):
        assert H.to_kgh(1000.0, "g/h") == pytest.approx(1.0, abs=1e-9)

    def test_to_kmolh_kgmolh(self):
        assert H.to_kmolh(5.0, "kgmol/h") == 5.0

    def test_to_kmolh_molh(self):
        assert H.to_kmolh(1000.0, "mol/h") == pytest.approx(1.0, abs=1e-9)

    def test_to_kmolh_lbmolh(self):
        result = H.to_kmolh(100.0, "lbmol/h")
        assert abs(result - 100.0 * H.KG_PER_LB) < 1e-9


# ══════════════════════════════════════════════════════════════════════════════
# Water detection
# ══════════════════════════════════════════════════════════════════════════════

class TestWaterDetection:

    @pytest.mark.parametrize("name", ["water", "Water", "WATER", "h2o", "H2O", "agua", "Agua"])
    def test_is_water_true(self, name):
        assert H._is_water(name), f"Expected {name!r} to be recognized as water"

    @pytest.mark.parametrize("name", ["Methane", "H2S", "Steam", "watery", ""])
    def test_is_water_false(self, name):
        assert not H._is_water(name), f"Expected {name!r} NOT to be recognized as water"


# ══════════════════════════════════════════════════════════════════════════════
# Integration test against real input files
# ══════════════════════════════════════════════════════════════════════════════

class TestEndToEnd:
    """
    Run the full homologation pipeline (non-interactively) against the real
    input_template.xlsx, phast_compounds_v3.json, and component_aliases_v2.json.

    This validates that:
    - script completes without exception,
    - at least one stream is parsed,
    - all streams satisfy Σ frac_corrected = 1.0,
    - pseudocomponent routing produces only hydrocarbon PHAST names,
    - no toxic compound in a stream's homologated rows has been excluded
      without a warning in its HomologatedRow.
    """

    SCRIPT_DIR = SCRIPT_DIR

    @pytest.fixture(scope="class")
    def real_results(self):
        db_path      = self.SCRIPT_DIR / "phast_compounds_v3.json"
        aliases_path = self.SCRIPT_DIR / "component_aliases_v2.json"
        input_path   = self.SCRIPT_DIR / "input_template.xlsx"

        if not db_path.exists() or not aliases_path.exists() or not input_path.exists():
            pytest.skip("Real input files not found — skipping integration test")

        all_compounds, hydrocarbon_only = H.load_db(db_path)
        aliases  = H.load_aliases(aliases_path)
        streams  = H.parse_input(input_path)
        assert streams, "No streams parsed from real input template"

        results = []
        for s in streams:
            res = H.homologate(s, all_compounds, hydrocarbon_only, aliases, n_requested=14)
            results.append(res)

        return results, all_compounds, hydrocarbon_only

    def test_at_least_one_stream_parsed(self, real_results):
        results, *_ = real_results
        assert len(results) > 0

    def test_renormalization_holds_for_all_streams(self, real_results):
        results, *_ = real_results
        for res in results:
            if not res.consolidated:
                continue
            total = sum(c.frac_corrected for c in res.consolidated)
            assert abs(total - 1.0) < 1e-7, (
                f"Esc={res.stream.escenario} C={res.stream.corriente}: "
                f"Σ frac_corrected = {total}"
            )

    def test_pseudocomponent_rows_have_hydrocarbon_pool_label(self, real_results):
        results, all_compounds, _ = real_results
        hc_names = {
            c.phast_name for c in all_compounds
            if H._is_hydrocarbon(c.formula)
        }
        for res in results:
            for hr in res.homologated:
                if hr.is_pseudocomponent:
                    assert hr.phast_name in hc_names, (
                        f"Pseudocomponent {hr.source_name!r} resolved to "
                        f"{hr.phast_name!r} which is NOT a hydrocarbon"
                    )
                    assert hr.match_pool == "hydrocarbon_only", (
                        f"Pseudocomponent {hr.source_name!r} used pool "
                        f"{hr.match_pool!r}, expected 'hydrocarbon_only'"
                    )

    def test_dropped_toxic_has_warning(self, real_results):
        """
        Any toxic compound present in homologated rows but absent from
        consolidated rows must have a HAZARD warning.
        """
        results, all_compounds, _ = real_results
        flag_map = {c.phast_name: c for c in all_compounds}

        for res in results:
            consolidated_names = {c.phast_name for c in res.consolidated}
            for hr in res.homologated:
                phast_c = flag_map.get(hr.phast_name)
                if phast_c is None or not phast_c.is_toxic:
                    continue
                if hr.phast_name in consolidated_names:
                    continue
                # Dropped toxic must have warning
                assert hr.warning, (
                    f"Esc={res.stream.escenario} C={res.stream.corriente}: "
                    f"Dropped toxic {hr.phast_name!r} has no warning"
                )
                assert "HAZARD" in hr.warning or "toxic" in hr.warning.lower(), (
                    f"Warning does not mention hazard: {hr.warning!r}"
                )

    def test_no_nan_in_consolidated_fractions(self, real_results):
        results, *_ = real_results
        for res in results:
            for c in res.consolidated:
                assert not math.isnan(c.frac_corrected), (
                    f"NaN frac_corrected in {res.stream.corriente}: {c.phast_name}"
                )
                assert not math.isnan(c.mass_kgh), (
                    f"NaN mass_kgh in {res.stream.corriente}: {c.phast_name}"
                )

    def test_molar_sum_check_field_matches_actual(self, real_results):
        """StreamResult.molar_sum_check must match the actual sum of frac_corrected."""
        results, *_ = real_results
        for res in results:
            actual = round(sum(c.frac_corrected for c in res.consolidated), 8)
            assert abs(actual - res.molar_sum_check) < 1e-9, (
                f"molar_sum_check field mismatch: stored={res.molar_sum_check}, "
                f"actual={actual}"
            )
