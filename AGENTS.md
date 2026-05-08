# AGENTS.md

## Repository purpose
This repository benchmarks coding models on a real Python engineering script used to homologate process-stream compositions for PHAST simulations.

## Project goals
- Preserve current script behavior unless a change is explicitly justified and validated.
- Improve robustness, maintainability, and testability.
- Prefer small, verifiable changes over broad rewrites.
- Consider modularization only if it reduces coupling and preserves observable behavior.
- A hexagonal or layered architecture is acceptable only if it clearly improves future extensibility without breaking outputs.

## Non-negotiable rules
- Do not remove existing features.
- Do not weaken pseudocomponent handling.
- Do not break alias priority.
- Do not break hydrocarbon-only routing for pseudocomponents.
- Do not break molar and mass renormalization guarantees.
- Do not arbitrarily change Excel output structure or formatting.
- Do not invent unsupported assumptions about the input files.

## Expected workflow
1. Read the repository files first.
2. Explain current architecture and risks before making major changes.
3. Add or update tests before or alongside meaningful code changes.
4. Validate changes against the real input template and expected output behavior.
5. Summarize modified files, risks, and validation evidence.

## Coding expectations
- Keep functions focused and names explicit.
- Avoid hidden side effects.
- Keep CLI behavior backward compatible unless explicitly requested.
- Prefer deterministic behavior over clever heuristics.
- Document edge cases in tests.

## Validation expectations
At minimum, validate:
- script execution completes successfully,
- expected Excel output is generated,
- normalization constraints still hold,
- pseudocomponent routing still behaves correctly,
- alias handling still has priority,
- no obvious formatting regressions appear in the Excel output.

## Deliverables expected from coding agents
- technical diagnosis,
- implementation plan,
- code changes,
- tests,
- validation summary,
- residual risks.
