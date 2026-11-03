# Known limitations and release gates

This is a free development edition, not an Excel-compatible production release.

- XLSX conversion preserves a basic subset. Macros, legacy XLS, XLSM, XLSB, ODS, pivot tables, Power Query, external workbook links, embedded charts/images, full conditional-formatting/validation conversion, cell comments, advanced borders, named ranges and advanced print metadata are not guaranteed. Warnings are shown and originals are retained locally.
- The open-source formula engine is not Microsoft's engine. SUM, IF and ordinary arithmetic/dependencies have bounded tests. Excel-specific functions, dynamic arrays, circular calculations, locale nuances and complete numerical equivalence need a corpus.
- The engine provides sorting, filtering, validation and conditional-formatting commands. Dedicated integration tests for every command remain outstanding; native snapshots retain installed preset data, but XLSX exchange does not preserve all of it.
- Native snapshots and named versions can grow large. Import limits cover input size/dimensions/populated cells, not every compressed-archive expansion scenario. Resource-limited hostile workbook testing remains required.
- CSV exports a single active sheet and cached values, not formulas or formatting. Text that could execute as a formula is prefixed. This protection intentionally changes those exported text strings.
- The chart is a selected-data preview limited to 30 numeric rows, not an embedded chart object. Negative values use magnitude bars and distinct color/labels, not a full signed-axis charting system.
- Print exports active-sheet displayed values, currently limited to 300 used rows and 30 columns. It does not reproduce Excel print areas, repeated titles, page breaks, complete formatting or print scaling. Physical printers remain untested.
- Local storage is IndexedDB; clearing browser/WebView data deletes the library and retained originals. Export native and original backups. There is no synchronized cloud account, independent crash timeline, SQLite index or file watcher.
- File conflict detection covers an active native save session. Library reopen uses a new save target. Native file associations are configured for `.scloud`, but command-line file opening is not implemented.
- UI keyboard editing and named main inputs are connected. Full screen-reader support, accessibility review and mobile spreadsheet workflows are not validated.
- The spreadsheet engine and optional dictionaries create a large lazy bundle. Startup, memory and very large sheets need profiling. There is no measured million-row performance claim.
- Only the current Windows host and installed Chrome have been exercised. macOS, Linux, ARM64, clean-machine installer/uninstaller behavior and production signing have not been validated.
- License notices and metadata are generated for production dependencies. The `buffers@0.1.1` package has incomplete license metadata in its published package. Complete a transitive distribution-license review before public release; the generated notices are not a certification.
- Security overrides update transitive minimatch and uuid to patched versions. XLSX read/write and the runtime suite must remain passing when these overrides change. A clean npm audit is limited to published advisories, not a complete security assessment.

Next priorities: stronger OOXML compatibility corpus and unsupported-content detection; dedicated sort/filter/validation tests; formula and locale corpus; accessibility and performance; full desktop save/reopen and installer testing; then advanced workbook features and optional authenticated collaboration.
