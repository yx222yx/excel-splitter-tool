# Excel Splitter Performance Optimization Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Reduce large-workbook export time and remove stale blank-row metadata without changing split results.

**Architecture:** Keep the existing `openpyxl` reload-and-filter design. Add focused worksheet metadata compaction, make formula repair linear in actual cells, reuse validated Web plans, then benchmark a bounded process executor before enabling it.

**Tech Stack:** Python 3.11, openpyxl 3.1.5, Flask 3.1.3, pytest 8.4.2.

## Global Constraints

- Preserve the generic per-sheet direct, reference, linked, and full modes.
- Do not shrink intentionally full-column conditional-formatting or validation ranges.
- Keep progress monotonic and preserve per-value error reporting.
- Use `D:\conda_envs\edc\python.exe` for tests and benchmarks.

---

### Task 1: Compact Filtered Worksheet Metadata

**Files:**
- Modify: `src/excel_splitter/engine.py`
- Test: `tests/test_engine.py`

**Interfaces:**
- Consumes: `SheetConfig.header_row` and retained original row numbers.
- Produces: `_compact_filtered_sheet(sheet, config, kept_original_rows) -> None`.

- [ ] Add a failing workbook-level test with custom row dimensions and a stale auto-filter range.
- [ ] Run `D:\conda_envs\edc\python.exe -m pytest tests\test_engine.py -k compact -q` and confirm the output retains metadata below `max_row`.
- [ ] Rebuild retained row dimensions and shrink row-bounded auto-filter/table references.
- [ ] Run the targeted test and all engine tests.

### Task 2: Make Formula Repair Linear

**Files:**
- Modify: `src/excel_splitter/engine.py`
- Test: `tests/test_engine.py`

**Interfaces:**
- Consumes: `kept_rows_per_sheet` produced by filtering.
- Produces: the existing `_fix_formula_references(...)` behavior without per-row `max_column` scans.

- [ ] Add a failing test whose worksheet rejects `max_column` access during formula repair.
- [ ] Run the targeted test and confirm it fails in `_fix_formula_references`.
- [ ] Iterate the cell map once and translate only moved formula cells.
- [ ] Run formula tests and all engine tests.

### Task 3: Reuse Web Split Plans

**Files:**
- Modify: `src/excel_splitter/engine.py`
- Modify: `src/excel_splitter/web/routes.py`
- Test: `tests/test_engine.py`
- Test: `tests/test_web.py`

**Interfaces:**
- Produces: `SplitEngine.execute(job, progress_callback=None, plan=None)`.
- Stores: `_split_plan` and `_split_plan_configs` in the Web job record.

- [ ] Add failing tests proving supplied plans skip rebuilding and mismatched Web configs do not reuse a cached plan.
- [ ] Pass an optional validated plan through the engine and background route.
- [ ] Run targeted engine/Web tests.

### Task 4: Benchmark Bounded Multi-Core Execution

**Files:**
- Modify: `src/excel_splitter/engine.py`
- Test: `tests/test_engine.py`

**Interfaces:**
- Produces: a top-level picklable per-value worker and a conservative worker selector.

- [ ] Add failing tests for worker selection and ordered result/error aggregation.
- [ ] Implement bounded process execution while keeping one-value execution in-process.
- [ ] Benchmark 1, 2, and 3 workers on the real workbook with identical settings.
- [ ] Keep the fastest stable setting only when it materially beats one worker; otherwise retain in-process execution.

### Task 5: Full Verification

**Files:**
- Verify: `src/excel_splitter/engine.py`
- Verify: `src/excel_splitter/web/routes.py`
- Verify: `tests/test_engine.py`
- Verify: `tests/test_web.py`

- [ ] Run `D:\conda_envs\edc\python.exe -m pytest`.
- [ ] Run `D:\conda_envs\edc\python.exe -m compileall -q src`.
- [ ] Export the real workbook to a temporary directory.
- [ ] Verify row-node count, auto-filter range, output count, errors, and readability.
- [ ] Run `git diff --check` and inspect the final diff without staging unrelated files.
