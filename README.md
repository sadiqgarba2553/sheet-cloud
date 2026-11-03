# Sheet Cloud

A free, local-first spreadsheet editor by BlueCloud AI Technologies. Sheet Cloud is the spreadsheet sibling of Doc Cloud, with original branding and a restrained green interface.

This is a development edition, not full Microsoft Excel compatibility.

## Run

Requires Node.js 22.12+; Node.js 24 was used during development.

```powershell
npm ci
npm run dev
```

Open http://127.0.0.1:1430. Keep the same browser hostname/port to access the same local workbook library.

## Work with your data

- Create a blank workbook or an editable budget, project, sales or inventory template.
- Type directly into the grid, or enter a cell/range in the address box and use the formula field. Press Enter to apply an entry.
- The engine supplies formulas, multiple sheets, formatting, number formats, sorting, filters, validation and conditional formatting controls. Not every supplied command has been independently regression tested.
- Workbooks autosave locally after 800 ms of inactivity. “Saved on this device” refers to the local library, not an exported disk file.
- Export `.scloud` for complete supported workbook snapshots, plugin resources and named versions.
- Open `.xlsx` or `.csv`, and export basic XLSX or active-sheet CSV. XLSX export asks for confirmation of compatibility limits. Imported XLSX originals are retained in a separate local backup store and can be retrieved from File.
- CSV imports preserve leading-zero identifiers and formula-like text. CSV exports prefix potentially executable text; CSV round-trip is not full workbook fidelity.
- Chart preview plots the first and last columns of a selected range, up to 30 numeric rows. This is a preview, not an embedded workbook chart.
- Print produces an active-sheet values table for the browser/native WebView print dialog, up to 300 used rows and 30 columns. Choose Save as PDF where supported.
- Local history keeps named workbook snapshots and a backup before restoration.

Export backups before clearing browser/application data. No account or hosted spreadsheet service is required.

## Desktop

Install Rust and the platform's Tauri prerequisites. Windows requires Microsoft C++ Build Tools and WebView2.

```powershell
$env:CARGO_BUILD_JOBS = '1'
npm run tauri dev
npm run tauri build
```

Unsigned development EXE/MSI packages are delivered in `releases/windows`. Production signing, clean-machine installation and macOS/Linux validation remain release gates. This session reused the adjacent Doc Cloud Cargo build cache through `CARGO_TARGET_DIR`; this is optional, and source/build configuration is independent of Doc Cloud.

## Verify

```powershell
npm test
npm run test:e2e
npm run build
npm audit
cargo test --manifest-path src-tauri/Cargo.toml -j 1
```

Playwright uses installed Chrome. Test downloads, screenshots and PDF output are under `output/playwright`.

## Architecture and evidence

React/TypeScript/Vite provide the application shell. The Apache-2.0 Univer engine supplies the canvas grid and calculation/editor subsystem. ExcelJS provides a local XLSX compatibility subset; no paid Univer exchange or collaboration services are used. IndexedDB stores workbooks and imported originals. Tauri provides native dialogs and atomic Rust file operations.

See `docs/FEATURE_MATRIX.md`, `docs/KNOWN_LIMITATIONS.md`, `docs/RESEARCH.md` and `docs/TEST_REPORT.md`. REA inspected the installed Excel executable's PE resource metadata; this is not a claim to have decompiled or reproduced Excel's calculation engine.

License metadata is recorded in `docs/DEPENDENCY_LICENSES.json`; bundled notices are generated into `public/THIRD_PARTY_NOTICES.txt` during `prebuild`. JSZip's MIT option is selected. An inherited package with incomplete license metadata remains identified for release review.

## Windows alpha releases

Source and downloads belong to `sadiqgarba2553/sheet-cloud`. Build in release mode with `npm run tauri -- build --bundles nsis`, then run `./scripts/package-windows.ps1`. Commit and push source before `npm run release:prepare` and `npm run release:publish`. The publisher verifies repository ownership, source commit, upload sizes and SHA-256 digests. Signed launcher metadata is promoted from the launcher repository only after public downloads pass verification. See RELEASE-NOTES.md for the alpha scope.
