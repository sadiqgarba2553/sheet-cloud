# References and REA evidence

2 November 2026. Original Sheet Cloud implementation; no Microsoft code or protected visual assets copied.

## Public workflow and component references

- [Microsoft: Basic tasks in Excel](https://support.microsoft.com/en-us/excel/basic-tasks-in-excel): workbook/sheet creation, data entry, formulas, number formats and grid workflows informed acceptance checks.
- [Microsoft: Create a simple formula](https://support.microsoft.com/en-us/excel/create-a-simple-formula): equals-prefixed arithmetic and ordinary function entry informed formula-field behavior.
- [Univer installation](https://docs.univer.ai/guides/sheets/getting-started/installation), [core features](https://docs.univer.ai/guides/sheets/features/core), [Facade sheet APIs](https://docs.univer.ai/guides/sheets/features/core/sheets-api) and [themes](https://docs.univer.ai/guides/sheets/ui/themes): used to integrate the Apache-2.0 grid/editor, snapshots, command events and appearance.
- [ExcelJS repository](https://github.com/exceljs/exceljs) and [license](https://github.com/exceljs/exceljs/blob/master/LICENSE): local XLSX conversion subset, under MIT. Paid Univer import/export services are not used.
- Local installed TypeScript definitions and package source were checked against the integrated versions. They establish API signatures, not passing application behavior.

## Installed Excel: static REA inspection

Target resolved by exact installed path: `C:/Program Files/Microsoft Office/root/Office16/EXCEL.EXE`. Windows file metadata reported product version `16.0.20527.20112`.

- REA operation: `inspect_pe_resources`.
- Evidence ID: `ev_05cab8eae7cb06d98e6c615b834f63f7079b75562e0825a3f88d4684d77ded82`.
- SHA-256: `1c51767e1136479c38b3e9ce4f0e4d59f8facfa782b9718ad73516539da2815c`.
- Coverage: complete bounded resource inventory; 2,471 examined entries and 1,227 resources.
- Complete metadata evidence: `docs/research/excel-pe-resources.json`.
- Observed: PE resource directory/payload locations, resource identities, hashes and static icon-group relationships.
- Unknown: calculation algorithms, live workflows, runtime resource loading, language fallback and UI behavior. Opaque resource payloads were not interpreted as source or behavior. No icons or other proprietary assets were extracted into Sheet Cloud.

This is real static evidence about the installed executable, not decompilation of Excel's spreadsheet engine. Workflow choices come from public documentation and independent design; correctness claims come from Sheet Cloud's tests.

## Implementation decisions

Use established open-source grid/formula functionality instead of a superficial React table. Preserve workbook snapshots locally, retain imported XLSX originals, and warn about lossy conversion. Keep accounts, cloud exchange and collaboration out of the required local path. Ship a clearly labeled development edition and record missing parity instead of claiming a complete Excel replacement.
