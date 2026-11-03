# Implementation architecture

2 November 2026.

`App.tsx` owns the library/editor lifecycle, file commands, local versions, appearance, formula input and chart/print views. `engine.ts` lazy-loads open-source Univer presets and creates one workbook instance at a time. Disposal clears the old engine and its undo state before mounting another workbook. No testing-only global engine API is exposed in production.

`model.ts` defines the versioned native record. Each record contains metadata, the complete supported Univer workbook snapshot and named snapshots. `.scloud` is UTF-8 JSON with format/version validation. Unknown versions, unsafe prototype keys and invalid worksheet dimensions are rejected. Resources supplied by the installed presets remain part of snapshots.

`store.ts` uses IndexedDB transactions for workbook and original-file persistence. Workbook edits debounce for 800 ms. Save status changes only after transaction completion; failures are visible. A mutation sequence prevents an older save from replacing newer editor state. Before navigation, the active workbook is saved. Deletion removes both workbook and retained original in one transaction. Pinned IDs and appearance use localStorage.

`exchange.ts` imports and exports a bounded XLSX subset using ExcelJS: primitive values, formulas and cached values, basic font/fill/number/alignment styles, merges, dimensions, hidden rows/columns and freeze panes. It does not offer full Excel serialization. Imported originals are retained byte-for-byte locally before opening converted data. Native imports with a conflicting local identity open as a separate copy rather than overwrite different local data.

CSV parsing handles quoted fields, escaped quotes and line breaks. Numeric inference preserves leading zeros and long identifiers. Formula-like CSV text is literal input. CSV export uses active-sheet cached values and prefixes unsafe leading formula characters for text.

The application uses a lazy canvas engine rather than a React component per cell. A custom address/formula bar writes through the engine Facade API. The engine supplies recalculation, selections, keyboard editing, formatting, sheet controls and bundled data tools. The current engine chunk is large; large-workbook performance has not been benchmarked.

Charts are standalone previews derived from the selected range's raw numeric values. Print serializes active-sheet displayed values to an escaped HTML table, then opens the print dialog. Neither feature claims embedded Excel chart or print-layout fidelity.

`files.ts` routes browser downloads or native open/save dialogs. Native commands enforce a 50 MB read limit and use unique same-directory temporary files, `sync_all`, and atomic persistence. A baseline byte comparison refuses to overwrite an externally changed native file within the current session. Save As chooses a new target; the local library retains no stale external write path. A file watcher is not implemented.

Tauri capabilities restrict dialog access to the main local window. Production CSP restricts scripts, remote connections and workers. No cloud service, telemetry or proprietary Excel asset is part of the implementation. Native file association launch arguments remain unimplemented.

Tests cover the connected compatibility subset; see TEST_REPORT.md. Desktop build workflow definitions are preparation, not evidence that a platform was tested.
