# IssueFlow V2 Alpha.3

This is a test build that runs side-by-side with V1.4.1.

## New in Alpha.3

- Added **Area Code**.
- Added **Room / Space** name.
- Room / Space is shown vertically on the generated issue sheet.
- Created date now includes time.
- Updated date includes time and is shown only after a meaningful change.
- Removed Remark from actions.
- Action structure is now:
  - Action By
  - Action Required
  - Status
- Action rows divide the available action height equally:
  - 3 actions = 3 equal rows.
  - 6 actions = 6 equal rows.
- No unused action rows are generated.
- Reference Images stays manual and has no “manual area” text.
- Cleaner modern layout with fewer borders.
- Section boundaries use line shapes instead of rounded card outlines.
- Removed Daily Painkiller + Codex branding from generated PowerPoint pages.
  Branding remains in the add-in UI.
- Party logo manager added under Settings.
- All parties in the party library get an equal header cell.
- Issue ID and Created/Updated Date-Time sit immediately after the logo area.
- Summary register removes Remark and groups Area/Room with the issue.

## GitHub

Upload these files directly to:

`public/v2/`

- taskpane.html
- taskpane.css
- taskpane.js

Do not create another nested `public/v2` folder.

## Test sequence

1. Add parties in Settings.
2. Upload a logo for each party if available.
3. Set Issue ID.
4. Enter Area Code, e.g. `L1-Z03`.
5. Enter Room / Space, e.g. `Basement-1`.
6. Enter Description.
7. Add 3 actions and refresh the sheet.
   Confirm the action section is divided into 3 equal rows.
8. Add 3 more actions and refresh.
   Confirm it becomes 6 equal rows.
9. Paste reference images manually and refresh again.
   Confirm the manual images remain.
10. Generate Summary and confirm there is no Remark column.
