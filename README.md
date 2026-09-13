# IssueFlow V2 Alpha.6.2

Hotfix for Alpha.6.1.

## Fixed

1. **Summary generation**
   - Restored `deleteExistingSummarySlides()`.
   - Restored `addCleanGeneratedSlide()`.
   - `Generate / Refresh Summary` no longer fails with `deleteExistingSummarySlides is not defined`.

2. **Area Code / Room / Description in Action Register**
   - Reads the correct V2 slide tags.
   - Includes migration fallback from visible managed shapes if an older issue slide is missing the tag.
   - Existing slides can therefore populate the grouped Area / Room and Description fields.

3. **Remove Action**
   - Removed unsupported `window.confirm()`.
   - First click shows `Remove?`.
   - Second click deletes the action.

No visual redesign was made in this hotfix.

## GitHub update
Replace:
- `public/v2/taskpane.html`
- `public/v2/taskpane.js`

CSS is unchanged.
