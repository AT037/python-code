# Copilot Instructions for this repository

## Project overview

This repository is a minimal Python learning project rather than a packaged application. Its standalone scripts are:

- [add.py](../add.py), which calculates and prints the sum and product of two values.
- [pocket_money.py](../pocket_money.py), an interactive weekly tracker that totals daily money received and spent by category, checks a savings goal, prints a scaled terminal graph of daily balances, and saves each week to an Excel workbook with `Data` and `Graph` sheets.

There is no application framework, no service layer, no package hierarchy, and no formal test suite or lint configuration in this repo today. Future work should stay lightweight unless the project clearly grows beyond this toy script structure.

## Build, test, and lint commands

The repository does not currently define a build pipeline, test runner, or lint configuration in a `pyproject.toml`, `setup.cfg`, `tox.ini`, `pytest.ini`, or similar file.

Use these commands for the current repo:

```bash
# install the declared Python dependencies
python -m pip install -r requirements.txt

# run the script itself
python add.py

# run the weekly pocket money tracker
python pocket_money.py
```

There are no existing project tests to run as a suite, and there is no configured linter to invoke. If this repository gains tests later, prefer running the smallest relevant target rather than the whole suite.

## High-level architecture

The architecture is intentionally simple:

- `add.py` and `pocket_money.py` are standalone executable scripts.
- Running `pocket_money.py` appends seven daily rows to `pocket_money.xlsx` beside the script and refreshes the line chart on its `Graph` sheet. If the workbook does not yet exist, prior `pocket_money.csv` records are imported. `.gitignore` excludes both files from version control.
- `requirements.txt` declares the Python package dependencies used by the project.
- `README.md` is minimal and describes the project as a beginner-level Python example.

Because the repository has no multi-module structure or runtime service, the main concern is keeping the code straightforward and understandable rather than organizing a complex app architecture.

## Key conventions

- Keep changes minimal and aligned with the repo’s learning-project scope.
- Prefer direct Python scripting patterns over introducing framework scaffolding or package structure unless there is a clear need.
- Maintain compatibility with the current environment and dependencies declared in `requirements.txt`.
- Preserve the simple top-level script style: small constants, direct computations, and printed output instead of a layered application design.
- Avoid adding broad tooling (pytest, ruff, mypy, CI, etc.) unless the project actually grows beyond this basic script.

## Working conventions for future edits

- Prefer extending the relevant existing script rather than creating a new app layout.
- If adding new functionality, keep it small and self-contained so the repository remains easy to understand.
- If dependencies are required, add them to `requirements.txt` rather than introducing a separate environment management workflow.
- When validating changes, use the smallest practical command that checks the behavior you changed (for example, `python pocket_money.py` for tracker changes).
