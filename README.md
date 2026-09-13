# IssueFlow V2 Alpha.6.4

Small stability + layout patch based on Alpha.6.3.

## Fixed

### Navigator
- Removed unsupported `target.select()`.
- Uses `context.presentation.setSelectedSlides([target.id])`.
- Type/paste Issue ID + Go remains unchanged.

### Area / Room strip
- Top 50% = Area Code, vertical.
- Bottom 50% = Room / Space, vertical.
- Thin divider between them.
- Both are centered in the narrow strip.

## Regression checks performed
- JavaScript syntax check passed.
- No `TAG_ACTIONS_JSON`.
- No `window.confirm`.
- Summary helper functions are present.
- Navigator refresh and Go functions are present.
- Unsupported `target.select()` is gone.

## GitHub update
Replace:
- `public/v2/taskpane.html`
- `public/v2/taskpane.js`

CSS is unchanged.
