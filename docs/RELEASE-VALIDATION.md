# Windows alpha release validation

3 November 2026, Africa/Lagos.

- Unit/integration: 13 passed. An initial XLSX case exceeded the default five-second timeout during concurrent work; the full rerun passed.
- Production-preview browser workflows: 8 passed, including formulas, native and XLSX/CSV exchange, local persistence, history, chart preview and print.
- TypeScript and production frontend build: passed.
- Rust filesystem tests: 2 passed.
- Optimized Windows x64 executable and NSIS build: passed. The adjacent Doc Cloud Cargo target directory was reused only as a build cache.
- Packaged native executable: home and budget template rendered in an isolated WebView2 profile, with the real grid, ribbon, calculated totals and local saved status.
- Managed executable matches the final native build by SHA-256; ZIP and standalone installer were copied after bundling completed. Exact deliverable hashes are in BUILD_ARTIFACTS.json.
- npm audit: zero reported vulnerabilities.
- App-owned source archive, managed ZIP, standalone installer, release metadata and checksums are prepared for GitHub publication. Launcher 0.7.1 supplies signed-feed download verification.

Unsigned alpha only. Clean-machine install/uninstall, complete native dialog save/reopen, macOS/Linux/ARM64 and full Excel fidelity remain unverified. TEST_REPORT.md records the earlier development baseline. A memory-constrained development-server browser run was abandoned; the complete production-preview run above passed. Upstream notices retain original whitespace and the documented incomplete metadata for the installed (not necessarily bundled) buffers dependency; a complete distribution-license audit is not claimed.
