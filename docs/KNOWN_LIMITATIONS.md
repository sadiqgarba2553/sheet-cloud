# Sheet Cloud - Known Limitations

This document provides a transparent, comprehensive record of architectural, engine, and exchange boundaries in Sheet Cloud.

---

## 1. Macro & Scripting Execution Boundary
- **VBA & Macro Execution**: Sheet Cloud deliberately does not execute Visual Basic for Applications (VBA) macros or XLM legacy macros. Executing arbitrary code from downloaded spreadsheets is a severe security vulnerability.
- **Original Bytes Retention**: When a `.xlsm` or macro-enabled `.xlsx` is imported, the exact original bytes are stored safely in IndexedDB under a separate key. The user can retrieve the unmodified file at any time via the File menu.
- **Office Scripts / JS Automation**: Sandboxed JavaScript automation is not yet connected to the UI.

---

## 2. File Format & Exchange Fidelity
- **Embedded DrawingML in XLSX**: Embedded charts created in Sheet Cloud are saved losslessly in `.scloud` format under `workbook.charts`. When exporting to `.xlsx`, chart data ranges are exported, but native DrawingML vector chart objects are not generated because ExcelJS does not support direct OOXML DrawingML packaging.
- **Pivot Tables in XLSX**: Pivot tables created in Sheet Cloud are rendered as fully populated, styled tabular worksheets with headers and grand totals. They do not export as dynamic Excel OLAP PivotCache records.
- **Cell Comments & Threaded Notes**: Univer open-source core preset does not include threaded cell comments in the current build.

---

## 3. Calculation & Function Corpus
- **Dynamic Arrays & LAMBDA**: The open-source `@univerjs/engine-formula` supports common arithmetic, text, logic, lookup, and statistical formulas (`SUM`, `AVERAGE`, `COUNT`, `IF`, `VLOOKUP`, etc.). Advanced Excel 365 functions (`LAMBDA`, `BYROW`, `SCAN`, `MAP`) are not fully available in this open-source release.
- **Circular References & Iterative Calculation**: Circular references trigger error tokens; iterative calculation loops with max convergence steps are not enabled by default.

---

## 4. Performance & Scalability
- **Populated Cell Limit**: Sheet Cloud enforces an import ceiling of 200,000 populated cells to ensure rapid rendering and avoid memory exhaustion in single-threaded browser JavaScript runtimes.
- **File Size Ceiling**: Native `.scloud` imports are capped at 50 MB, and XLSX imports are capped at 25 MB.

---

## 5. Printing & Pagination
- **Print Table Chunking**: Large print layouts are converted into styled HTML tables with `@page` landscape/portrait rules and repeating table headers. Complex custom header/footer token fields (e.g., page X of Y in margins) require host print driver support.
