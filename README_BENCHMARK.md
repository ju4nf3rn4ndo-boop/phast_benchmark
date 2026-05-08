# README_BENCHMARK

## Objective
Benchmark four coding models on the same real-world engineering task using the same repository, same files, same prompt, and same acceptance criteria.

## Models under test
- Claude
- ChatGPT
- DeepSeek
- Z.AI / GLM-5.1

## Baseline assets
Stored in `00_baseline/`:
- phast_homologator.py
- input template workbook
- PHAST database JSON
- aliases JSON
- any supporting files required by the script

## Benchmark philosophy
This benchmark is not about style only. It measures:
- whether the model finishes,
- whether it preserves core behavior,
- whether it handles real inputs and edge cases,
- whether it produces robust and maintainable changes,
- whether it degrades under long-context work.

## Evaluation criteria
1. Completion: finishes the task without stalling or losing the thread.
2. Correctness: script still works and preserves technical guarantees.
3. Robustness: handles messy or edge-case inputs safely.
4. Output quality: generated Excel remains coherent and professionally usable.
5. Maintainability: improves code structure without destructive rewrites.
6. Regression safety: changes are backed by tests or strong validation.
7. Cost/time: practical efficiency is considered, but only after correctness.

## Repo structure
- `00_baseline/` original baseline files
- `01_runs/` isolated work directories by model
- `02_validation/` validation scripts and comparisons
- `03_prompts/` benchmark prompts
- `04_results/` outputs, reports, logs, scorecards

## Rules for fair comparison
- Same starting files for all models.
- Same main prompt.
- Same acceptance criteria.
- Same validation process.
- No manual code assistance during the run.
- Human review only after the model finishes.

## Current intended future direction
The codebase may evolve toward a more modular architecture, potentially hexagonal or layered, but only if the migration preserves current observable behavior and meaningfully reduces maintenance friction.
