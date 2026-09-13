# IssueFlow V2 Alpha.5

Stability fix build based on Alpha.4.

## Why this build

Two host-compatibility problems were identified:

1. PowerPoint line shapes were rendering as long diagonal objects on this host.
2. Party logos were stored in the add-in but not appearing on the generated slide.

## Fixes

### 1. No PowerPoint line shapes for layout
All structural separators now use very thin rectangular shapes.
This is more predictable and avoids diagonal connector geometry entirely.

### 2. Stable logo rendering
The build no longer relies on `ShapeCollection.addPicture()`, which Microsoft currently documents as a preview API.
Instead, each logo is rendered inside a normal rectangle using `shape.fill.setImage(base64)`.
This uses the normal PowerPoint shape-fill image workflow.

### Kept unchanged
- Area Code.
- Vertical Room / Space.
- Created date + time always shown.
- Updated date + time only after a meaningful change.
- Manual Reference Images.
- Action By / Action Required / Status only.
- No Remark.
- Equal action row division.
- Equal party header cells.

## GitHub update

Replace only:
- `public/v2/taskpane.html`
- `public/v2/taskpane.css`
- `public/v2/taskpane.js`

## Test

1. Confirm the task pane shows `V2 Alpha.5`.
2. Settings -> upload a party logo.
3. Confirm the small logo preview appears in Settings.
4. Create / Refresh Issue Sheet.
5. Confirm the logo appears in the party header cell.
6. Confirm no long diagonal lines exist.
7. Add 3 actions -> confirm 3 equal rows.
