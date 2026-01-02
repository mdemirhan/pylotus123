# Roadmap

## P0 (Correctness)
- [x] Integrate the recalculation engine and wire manual/auto modes.
- [x] Fix comparison precedence so arithmetic evaluates before comparisons.
- [x] Propagate #NAME? and #N/A errors consistently in the parser.
- [x] Require sorted data for approximate VLOOKUP/HLOOKUP.

## P1 (Maintainability)
- [x] Extract format code normalization into a shared helper.
- [x] Refactor chart range setters to use a single helper.
- [x] Centralize overwrite confirmation dialogs.
- [x] Add a shared view reset helper for grid state.

## P2 (Performance)
- [x] Optimize topological ordering in the recalc engine.

## P3 (Testing)
- [ ] Add shared pytest fixtures in `tests/conftest.py`.
