1. Architecture Summary
The script is a single-file, 1,659-line monolith structured as a linear pipeline:
CLI (argparse)
  └─ run()
       ├─ load_db()           → PhastCompound[]
       ├─ load_aliases()      → dict (built-in _HYSYS merged with external JSON)
       ├─ parse_input()       → StreamInput[]
       ├─ scan_water_content()→ WaterScanResult[]
       ├─ prompt_water_override() → set[int]
       ├─ homologate()        → StreamResult[]   [one per stream]
       │    ├─ resolve()      [per component]
       │    │    ├─ alias lookup  (priority 1)
       │    │    ├─ pseudocomponent detection
       │    │    └─ MW nearest-neighbour
       │    ├─ consolidate by PHAST name
       │    ├─ rank + threshold + filter
       │    └─ renormalize
       ├─ apply_post_options() → additional StreamResult[] (OPT1/2/3)
       └─ write_output()       → .xlsx
Dataclasses used as typed data containers (not domain objects with behavior): PhastCompound, RawComponent, StreamInput, HomologatedRow, ConsolidatedRow, StreamResult, WaterScanResult.
There are no tests, no modules, no configuration classes, and no separation of I/O from domain logic.
---
2. Critical Domain Rules (inferred from code)
R1 — Alias priority is absolute
resolve() at phast_homologator_v2.py:397–415: if any key in the merged alias table matches name (case-insensitive, stripped), the PHAST compound is assigned immediately. MW mismatch is warned but never overrides the alias. The external component_aliases_v2.json wins over the built-in _HYSYS table because merged.update(ext) is applied last.
R2 — Pseudocomponent routing is hydrocarbon-only
_is_pseudocomponent() at ~line 67–85: names matching four regex patterns (36-40C*, 850+C*, NBP..., C10+) are restricted to the hydrocarbon_only pool (49 CnHm formula compounds). They cannot accidentally match Water, CO2, H2S, or any non-hydrocarbon.
R3 — Water Override mode changes N accounting
When water override is active: n_effective = n_requested - 1. Dry components compete for N−1 slots. Water is appended at its original molar fraction after dry renormalization. Dry fractions are scaled by (1 − x_water).
R4 — Molar renormalization to exactly 1.0
Σ frac_corrected = 1.0 is enforced for every stream. In water-override mode: x_water + Σ_dry × (1 − x_water) = 1.0.
R5 — Post-options are additive, never destructive
apply_post_options() appends additional StreamResult objects tagged OPT1_DRY, OPT2_REM, OPT3_FORCE. The original BASE result is never mutated.
R6 — AUTO_THRESHOLD = 1e-5 is a hard floor
Components with molar_frac < 1e-5 (0.001%) are dropped regardless of N. The threshold actually used is max(1e-5, molar_frac_of_Nth_compound).
R7 — MW tie-breaking favors alias, then alphabetical
_resolve_tie() at ~line 375: when two compounds have identical MW distance, the alias-suggested compound wins. Otherwise alphabetical. This is deterministic and reproducible.
R8 — Hazard tagging is compound-level, not stream-level
is_toxic, is_flammable, is_inert come from the JSON DB (or are derived from flammable_toxic_flag). The post-options logic scans for toxics that were present in the input but dropped by N-slot cutoff — this is the trigger for interactive recovery.
---
3. Likely Failure Modes and Edge Cases
F1 — Excel layout brittleness (parse_input)
parse_input() at ~line 469 detects stream columns by scanning row 1 for cells containing "corriente" (case-insensitive). If the template is reformatted (extra rows, merged cells, different header text), parsing silently yields zero streams with no error. There is no schema validation of the workbook.
F2 — Silent zero-flow exclusion
Components with zero mass AND zero molar flow are silently skipped. If the input contains a non-zero toxic in mass but zero molar (or vice versa), the component is included but mw_calc = mass_kgh / molar_kmolh would raise ZeroDivisionError. There is no guard for this.
F3 — mw_calc ZeroDivisionError
RawComponent.mw_calc is a computed property: mass_kgh / molar_kmolh. If molar_kmolh == 0 but mass_kgh != 0, this raises. The resolve() fallback only triggers on mw_calc <= 0 or NaN — it does not catch ZeroDivisionError from the property itself.
F4 — Stream naming collisions in Excel
Sheet names are Esc{N}_C{corriente}. If the same (escenario, corriente) appears twice (e.g., one BASE + one OPT1), the second sheet appends _OPT1_DRY. But Excel sheet names have a 31-character limit; long corriente names could produce truncated or duplicate sheet names. There is no deduplication guard.
F5 — Water detection is name-based only
WATER_KEYS = frozenset({'water', 'h2o', 'agua'}). A component named "WATER", "H2O", or "agua" triggers the water override pre-scan. But a component aliased to "Water" via an alias entry is not pre-scanned correctly — the scan operates on raw input names, not resolved PHAST names. So "H2O" in the DB alias table resolves to "Water" but the scan checks name.lower() in WATER_KEYS before alias resolution.
Actually, H2O is in WATER_KEYS, so that specific case is covered. But a component named "Vapor_agua" that aliases to Water would be missed.
F6 — apply_post_options uses interactive stdin
apply_post_options() calls input() directly. In non-interactive environments (CI, piped stdin), this would hang or raise EOFError. There is no --non-interactive flag.
F7 — N=18 with water override on a single-component stream
If a stream has 1 component + water, N−1=0 dry slots → effectively zero dry components. The result would be 100% water, which may be valid but is not explicitly guarded or warned.
F8 — Alias target not in DB
load_aliases() does not validate that every alias target exists in phast_compounds_v3.json. If the JSON alias table points to a compound name not in the DB, resolve() finds the alias name but then fails to find it in all_compounds (lookup by name). The fallback would then proceed to MW matching, silently ignoring the alias.
F9 — Pseudocomponent with no calculable MW
If a pseudocomponent has zero mass and zero molar flow (but was not excluded by the zero-flow filter), mw_calc = 0/0 = NaN, and the fallback "fallback_no_mw" returns hydrocarbon_only[0] (the lightest hydrocarbon in the pool). This is correct behavior but not warned in the output.
F10 — _TARGET_REDIRECTS is a hardcoded singleton
"1-Pentene" → "n-Pentane" is the only redirect. If a future alias update inadvertently creates a longer chain or cycle, there is no cycle detection.
---
4. Technical Debt and Maintainability Issues
TD1 — Single 1,659-line file with mixed concerns
I/O (Excel read/write), domain logic (resolve, homologate), interactive prompts, and formatting are completely interleaved. Changes to the Excel format require reading through the homologation logic to find the right section.
TD2 — No tests whatsoever
Zero automated tests. There is no way to confirm that a refactoring or bug fix preserves behavior without manually running the full script and inspecting Excel output.
TD3 — write_output is ~286 lines of imperative openpyxl calls
Style objects are recreated inline (not shared). Column widths are set with magic numbers. Sheet naming logic is embedded inside the writer. Any formatting change requires careful surgery.
TD4 — Interactive prompts interwoven with computation
prompt_water_override() and apply_post_options() call input() directly. This makes the script impossible to test automatically or call from another program.
TD5 — _HYSYS hardcoded inside the module
Built-in alias table is at the top of the file, not in the external JSON. Changes to aliases require editing source code.
TD6 — AUTO_THRESHOLD, PHAST_MAX, etc. are module-level magic constants
Not configurable from CLI (except --n). No way to change MW_WARN_PCT or AUTO_THRESHOLD without editing the source.
TD7 — homologate() signature has 7 parameters, handles two fundamentally different modes (override vs. normal) via boolean flag
The two modes share a complex conditional block at the renormalization step. This is hard to follow.
TD8 — Dataclasses are pure data bags with no invariant enforcement
molar_frac can be negative (no validation). ConsolidatedRow.frac_corrected is set from outside with no bounds check. A future bug could produce fractions > 1 without any assertion.
TD9 — Water pre-scan uses raw names, resolved names are not available at scan time
The scan happens before alias resolution. This is an architectural coupling that means water detection accuracy depends on input naming conventions rather than resolved compound identity (see F5).
TD10 — Excel output path construction uses f"Esc{N}_C{corriente}" without length guard
Sheet name length is not checked against Excel's 31-character limit.
---
5. Validation Strategy
Given the non-negotiable rules in AGENTS.md, the minimum validation suite should cover:
V1 — Happy path end-to-end
Run the script on input_template.xlsx with n=18, no water override, verify:
- Script exits with code 0
- Output .xlsx is created
- Sheet Resumen_Homologacion has ≥1 row per stream
- Each stream sheet has Part A + Part B sections
- Σ frac_corrected = 1.0 for every stream (within 1e-9)
V2 — Water override correctness
Force water override on Stream 5 (C171, which has ~54% water):
- frac_corrected of Water row = original x_water
- Dry sum = 1 − x_water
- N used ≤ n_requested − 1 for dry components
V3 — Alias priority over MW
Test a component whose alias resolves to compound X but whose MW is closest to compound Y. Confirm output shows X, not Y.
V4 — Pseudocomponent hydrocarbon-only routing
Test a pseudocomponent name like "36-40C*" with MW closest to Water (18 g/mol — not realistic but as a unit test). Confirm it maps to the closest hydrocarbon, not Water.
V5 — Renormalization invariant
For every stream in every mode (BASE, OPT1, OPT2, OPT3): |Σ frac_corrected − 1.0| < 1e-9.
V6 — Toxic recovery post-options
Given a stream where a toxic (e.g., HCN) is dropped by N cutoff, verify that OPT3 correctly re-inserts it and renormalizes.
V7 — MW tie determinism
Given two compounds with identical MW, confirm that running the script twice produces the same assignment (no randomness).
V8 — Unit conversion correctness
to_kgh and to_kmolh for all supported units.
V9 — ZeroDivisionError guard (currently missing)
Test a component with mass>0, molar=0 to confirm it does not crash.
---
6. Whether Deeper Modularization Is Advisable Now
Short answer: not yet for the full architecture, but targeted extractions are high value.
What is advisable now
Action	Benefit
Extract resolve() + _find_closest() + _resolve_tie() into a pure function module	Enables unit testing of the core matching logic without Excel I/O
Extract homologate() + _renormalise() into a pure function module	Enables renormalization invariant tests
Extract write_output() into a separate writer module	Isolates Excel formatting from domain logic
Separate prompt_* functions behind an interface	Enables non-interactive / CLI mode and testability
Add pytest tests for V1–V8 above	No production risk, high validation value
What is not advisable without further justification
- Full hexagonal/port-adapter architecture: over-engineered for a ~1,600-line engineering tool used by process safety engineers. The cost in complexity would outweigh the maintainability gain unless the script needs to be embedded in a larger system.
- Rewriting the Excel writer in a template-based approach (Jinja2 + xlsx template): would change observable output formatting, violating the non-negotiable rules.
- Replacing the dataclasses with a full domain model with validation: low ROI given the script's scope.
Recommended phased approach (if modularization proceeds)
1. Phase 1 — Add tests first (no production code changes): tests/test_resolve.py, tests/test_homologate.py, tests/test_renormalize.py
2. Phase 2 — Extract pure domain functions into phast_domain.py (resolve, homologate, renormalize) while keeping phast_homologator_v2.py as the orchestrator
3. Phase 3 — Extract writer into phast_writer.py
4. Phase 4 — Decouple prompts via a thin interface so batch/CI mode is possible
Each phase is independently releasable and verifiable against the test suite added in Phase 1.
---
Summary of highest-risk items to address before any modification:
Priority	Issue	Consequence if unaddressed
Critical	F3 — mw_calc ZeroDivisionError	Script crash on malformed input
Critical	TD2 — No tests	Any change may silently break behavior
High	F2 — Zero-flow toxic silently excluded	Hazardous compound missed without warning
High	F5 — Water detection is name-only	Water override not triggered for aliased water names
Medium	F4 — Sheet name length	Silent Excel truncation / duplicate sheets
Medium	TD6 — Magic constants not configurable	Need source edit for any threshold change
▣  Build · Claude Sonnet 4.6 · 16m 4s
