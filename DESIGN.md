# Sheet Cloud Design System

Sheet Cloud is the emerald green spreadsheet and analysis member of the BlueCloud desktop suite. Its interface shares Doc Cloud and Slide Cloud's quiet sidebar, practical command bars, local autosave ledger, and native backup workflow.

- **Brand Asset**: Supplied emerald green spreadsheet artwork, engraved in the BlueCloud superellipse squircle silhouette. Never substitute external competitor branding.
- **Typography**: Clean system font stack (Segoe UI on Windows, -apple-system on macOS), 13px base body, 11–12px controls, 9–10px metadata labels, and monospaced cell numbers for clear tabular alignment.
- **Color Palette**:
  - Light mode: Canvas background `#fafbfa`, panel surfaces `#ffffff`, sidebar `#f3f5f4`, line/borders `#e4e9e6`, text `#25302c`, muted text `#838d88`, brand green accent `#15836b` (hover `#106a56`), soft green highlights `#e5f2eb`.
  - Dark mode: Canvas `#1d2220`, panel surfaces `#242a27`, sidebar `#202623`, borders `#353e39`, text `#e7eeea`, muted text `#9faea6`, brand green `#72c7a7`, soft focus `#2b4338`.
- **Layout & Structure**:
  - Left navigation sidebar with quick access to All Workbooks, Recent, Pinned, and Templates.
  - Editor top bar with workbook title, local save indicator, and primary action buttons.
  - Cell address and formula bar (`ƒx`) with calculated display value readout.
  - Univer grid host with full canvas scrolling, row/column headers, and sheet tab strip.
  - Side drawer for embedded vector charts (column, bar, line, area, pie).
- **Spacing & Elevation**: 4/8/12/16/24/32/48px design rhythm. 6–8px control border radii, 1px subtle border lines.
- **Interactions**: Responsive grid keyboard navigation, modal overlays for Pivot Tables, Goal Seek, Data Tools, and Page Setup.
- **Honest Status**: "Saved on this device" reflects local IndexedDB commits; `.scloud` file export produces a full, losslessly serialized JSON workbook package.
