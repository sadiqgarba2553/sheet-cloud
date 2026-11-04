> Codex review (4 November 2026): Read [CODEX_VERIFICATION.md](CODEX_VERIFICATION.md) for tested behavior, corrected claims and remaining limits. Earlier source-only claims are not proof of feature parity.

# Sheet Cloud - Codex Handoff & Verification Guide

This guide provides Codex with exact verification commands, branch details, and expected behaviors for testing Sheet Cloud.

---

## 1. Environment & Branch Details
- **Repository Directory**: `C:\Users\Garba Ibrahim\Desktop\BlueCloud Apps\apps\office\sheet-cloud`
- **Active Feature Branch**: `feat/office-suite-expansion`
- **Base Clean Commit**: `d42df42c4350904a9c2252e33d6f547aebb05a89`
- **Primary Technology Stack**: React 18, TypeScript, Vite, `@univerjs/presets`, ExcelJS, Tauri (Rust).

---

## 2. Verification Steps for Codex

### Step 1: Run Unit Tests
Execute the unit test suite with Vitest to verify all workbook integrity, exchange, and analytical engine operations:
```powershell
cd "C:\Users\Garba Ibrahim\Desktop\BlueCloud Apps\apps\office\sheet-cloud"
npm test
```
**Expected Outcome**: All tests pass, including:
- Workbook independent IDs and version round-trip.
- Embedded charts persistence in `.scloud`.
- CSV RFC 4180 parsing and formula injection mitigation.
- JSON dataset import and export.
- Pivot table grouping and `SUM` aggregation on regional sales data.
- Goal Seek numerical root-finding convergence.
- Text-to-Columns delimiter splitting.
- Duplicate row removal.
- XLSX multi-worksheet, formula, style, and freeze pane round-trip.

### Step 2: Source Code Typecheck
Run the TypeScript compiler to ensure all interfaces and models are properly typed:
```powershell
npx tsc --noEmit
```
**Expected Outcome**: Clean exit with code 0.

### Step 3: Run Desktop Smoke Test
Launch the development server to verify the UI:
```powershell
npm run dev
```
Navigate to `http://127.0.0.1:1430` in the browser:
1. Verify top toolbar contains `Charts`, `Pivot Table`, `What-If`, `Data Tools`, and `History`.
2. Click `Charts` -> Click `Add Chart` -> Select Column family -> Verify chart renders in the side drawer.
3. Click `Pivot Table` -> Verify source range auto-fills -> Enter row field and value field -> Click `Create Pivot Sheet` -> Verify a new "Pivot 1" worksheet is added with headers and grand totals.
4. Click `What-If` -> Verify Goal Seek inputs -> Click `Solve` -> Verify cell updates to target root.
5. Click `File` -> Verify `.json` option is present alongside `.scloud`, `.xlsx`, and `.csv`.
