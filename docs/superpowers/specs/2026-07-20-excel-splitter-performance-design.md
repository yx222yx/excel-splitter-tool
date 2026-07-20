# Excel Splitter Performance Optimization Design

## Context

The current row filtering no longer hangs on a 12,182-row worksheet, but a
real ten-value export still spends most of its time repeatedly parsing
workbooks. Output worksheets also retain source row metadata and stale filter
ranges after their cells have been compacted.

## Considered Approaches

1. Optimize the existing `openpyxl` pipeline. Compact row metadata, scan
   formula cells once, and reuse the Web planning result. This is the
   recommended first stage because it preserves the current behavior and has
   a small compatibility surface.
2. Run independent split values in bounded worker processes. This can use
   multiple CPU cores, but each process holds one or two complete workbooks,
   so it must be benchmarked at 1, 2, and 3 workers before choosing a default.
3. Rewrite worksheet XML directly inside the XLSX archive. This offers the
   highest theoretical throughput but risks formulas, styles, merged cells,
   relationships, and extension fidelity. It is deferred.

## Approved Design

After filtering a worksheet, rebuild its row-dimension metadata from the
header rows and retained original data rows. Shrink row-bounded auto-filter and
table ranges to the new last row while leaving full-column conditional
formatting and data validation unchanged.

Formula reference repair will iterate the worksheet cell map once and update
only formula cells whose retained rows moved. It will not recompute
`max_column` for each retained row.

The Web layer will cache the `SplitPlan` produced by `/api/split-values` with
the exact immutable `SheetConfig` tuple and pass it to execution only when the
configuration still matches.

Multi-core execution will use a top-level, picklable worker and a bounded
process count. One selected value remains in-process for detailed progress.
Multiple values may use processes only if the real workbook benchmark shows a
clear improvement without excessive memory pressure. The process count will
be configurable and capped conservatively.

## Correctness And Verification

- Direct and linked filtering keep the same rows and counts.
- Custom row height/style from a retained source row moves to its new row.
- No row-dimension key or auto-filter end row remains below the compacted data.
- Formula references still translate after retained rows move.
- Cached plans are used only for identical sheet configurations.
- Progress is monotonic and reaches 100 percent.
- The full test suite passes.
- A real export is checked for XML row-node count, filter range, file count,
  errors, elapsed time, and workbook readability.
