# Sheet Cloud verification report

2 November 2026, Africa/Lagos. Windows x64 host, Node.js 24.15.0, npm 11.12.1, Rust/Cargo 1.97.1. Browser workflows ran in installed Google Chrome. These results establish a tested development subset, not Excel parity.

| Check | Result |
|---|---|
| `npm test` | 13 tests passed; recorded run 6.72 seconds |
| `npm run test:e2e` | 8 browser workflows passed; final run approximately 1.9 minutes |
| TypeScript / frontend build | Passed with Vite 8.3.5; final packaging build 8.10 seconds |
| `npm audit` | Zero known vulnerabilities, including production and development dependencies |
| `cargo test -j 1` | 2 tests passed: atomic replacement and binary nested-directory save/read |
| Windows desktop runtime | Tauri/WebView2 window opened; home, editable budget grid, calculated totals, dark system appearance and local saved status observed |
| Native save/reopen via dialog | Not completed; not claimed verified |
| Windows EXE/MSI | Unsigned x64 development/debug packages generated; delivered artifacts and hashes are in BUILD_ARTIFACTS.json |
| macOS/Linux/ARM64 | Not run; workflow definitions are preparation only |
| Clean-machine installation / uninstall / signing | Not run |
| Physical printing | Not run |

Total automated tests: 23 passing (13 unit/integration, 8 browser workflows, 2 Rust filesystem tests).

## Browser workflows

1. Create a workbook, enter values, calculate SUM and IF, autosave, return home, reload and reopen formulas/results.
2. Open the budget template, change a source value, observe dependent recalculation, export native data and reopen it.
3. Import a genuine XLSX fixture, calculate its formula, export XLSX, then inspect the generated workbook for values, formula and bold font preservation.
4. Import CSV with formula-like text; verify it remains literal; calculate a numeric sum; observe zero external HTTP(S) requests during the tested local flow.
5. Save a named version, edit, restore, verify original values and pre-restore backup.
6. Type directly into the canvas grid, undo, then create another workbook and verify isolated undo history.
7. Persist dark appearance, select sales data and verify a chart preview with six data bars.
8. Print the active budget sheet, assert calculated values in the print frame and generate a PDF.

The PDF fixture was additionally inspected with pypdf for page size and recoverable budget text/results. See the verification artifact under `output/playwright/budget.pdf`.

## Unit/integration scope

Tests cover unique workbook identities; native formulas/styles/version preservation; rejection of unsupported versions, unsafe properties and invalid dimensions; IndexedDB workbook/original persistence and deletion; CSV quoting/newlines/escaping; leading-zero and long identifiers; literal formula strings; export injection mitigation; XLSX formulas, merged headers, fonts, sheet names and freeze panes; multiple worksheets.

The Rust tests use real temporary files and replace an existing workbook atomically. They do not certify the full native dialog/editor/filesystem interaction.

## REA evidence

REA inspected the installed `EXCEL.EXE`, SHA-256 `1c51767e1136479c38b3e9ce4f0e4d59f8facfa782b9718ad73516539da2815c`. Evidence `ev_05cab8eae7cb06d98e6c615b834f63f7079b75562e0825a3f88d4684d77ded82` records 2,471 resource entries examined and 1,227 resources with complete bounded inventory coverage. This is static resource metadata, not Excel source or runtime calculation evidence. Full metadata and boundaries are in RESEARCH.md and `research/excel-pe-resources.json`.

## Findings addressed

- Fixed empty imported worksheets whose ExcelJS `columns` collection is null.
- Fixed absent worksheet visibility metadata being interpreted as hidden; ordinary XLSX sheets now render.
- Enabled the engine header so its formatting/data ribbon is visible.
- Used raw selected values for chart numeric data rather than formatted representations.
- Isolated workbook undo histories and protected named versions from stale save completion.
- Deleted retained original bytes together with a workbook in one IndexedDB transaction.
- Preserved conflicting native imports as separate local copies.
- Replaced a browser-test readiness race with explicit workspace/engine/canvas readiness checks. Earlier failing runs were not counted as passing.
- Overrode vulnerable transitive minimatch and uuid versions with patched versions, then reran XLSX and browser checks. The package audit is clean; overrides must remain regression tested.

## Remaining boundaries

The lazy spreadsheet engine is approximately 7.1 MB before gzip and optional language assets increase the installed footprint. Performance has not been benchmarked. Some engine-provided data tools are connected but lack dedicated integration tests. Embedded charts, pivots, macros, full OOXML fidelity, secure collaboration, advanced print layout, hostile-file resource exhaustion and complete accessibility remain unfinished.

The build collects license notices for 339 installed production packages. `buffers@0.1.1` has incomplete published license metadata, recorded as a release-review item. Clean dependency advisory results do not replace security or distribution-license review. Refer to KNOWN_LIMITATIONS.md and RELEASE_CHECKLIST.md before public release.
