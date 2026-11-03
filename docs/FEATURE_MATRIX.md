# Excel workflow and implementation status

“Connected” means code/UI integration exists. “Tested subset” is bounded by the named checks; it does not mean Excel parity.

| Expected Excel workflow | Sheet Cloud implementation | Acceptance check | Status |
|---|---|---|---|
| Create, edit and navigate a grid | Univer canvas, address box and formula bar | Direct keyboard entry; A1 navigation; reopen values | Browser tested subset |
| Formulas and dependency recalculation | Open-source Univer formula engine | SUM, IF, template arithmetic and dependency changes | Browser tested subset; no full function corpus |
| Fonts, bold, fill, alignment and number formats | Engine ribbon and styles | Basic style/number-format native and XLSX round-trip | Tested subset; full visual corpus pending |
| Multiple sheets, rename, row/column operations | Engine sheet UI and model | Multiple worksheet XLSX round-trip | Connected; full UI command regression pending |
| Sorting, filters, validation and conditional formatting | Open-source presets | Preset controls load; retain native plugin snapshots | Connected; dedicated behavior tests pending |
| Undo / redo | Engine command history | Undo direct typing; no history leak across workbooks | Browser tested subset |
| Autosave and reopen | IndexedDB and debounced snapshots | Reopen formulas and values after page reload | Browser tested subset |
| Full native save | Versioned `.scloud` JSON | Formula/style/version round-trip; reject invalid versions | Unit and browser tested subset |
| Real XLSX import/export | Local ExcelJS conversion | Open real XLSX; inspect exported values/formulas/font | Tested subset; compatibility warning required |
| CSV | Quote-aware parser and protected values export | Newlines/quotes/IDs; literal formulas; injection mitigation | Unit and browser tested subset |
| Import safety / original preservation | Byte limits, validation and separate original backups | Restore original bytes; delete original with record | Unit tested subset; hostile archives pending |
| Named versions | Local snapshot ledger | Restore and keep pre-restore backup | Browser tested subset |
| Charts | Selected-data bar chart preview | Selected range produces expected bars | Browser tested subset; no embedded chart export |
| Print / PDF | Escaped active-sheet values table | Budget values in print layout and PDF generation | Browser tested subset; no Excel page-layout fidelity |
| Home, templates, pinned/recent files | Original library and editable templates | Budget/sales template calculations | Connected; some library actions need dedicated tests |
| Dark / light / system | Shell and engine theme | Dark appearance survives reload | Browser tested subset |
| Excel executable identity | REA static PE resource metadata | Hash, inventory coverage and limitations recorded | Observed static evidence only |
| Windows desktop | Tauri window, native file dialogs and atomic writes | Rust tests, build and runtime smoke check | See TEST_REPORT.md |
| Mac/Linux/ARM64 | Platform-independent source and CI definition | Build and install on each host | Not tested |
| Macros, pivots, Power Query, embedded Excel charts | No equivalent connected | Excel-specific compatibility corpus | Not implemented |
| Cloud collaboration and permission roles | No service connected | Concurrent editing/auth/reconnect tests | Not implemented |
| Enterprise-scale performance and complete accessibility | No measured claim | Large-file profiling and assistive technology audit | Not validated |
