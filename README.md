# IssueFlow V2 Alpha.6.3

Navigator reliability hotfix.

## Fixed
- Fixed `refreshIssueNavigator is not defined`.
- Added the missing navigator refresh function.
- Added the missing Go-to-Issue function.
- Issue Navigator is now an editable text field with suggestions.
- You can copy/paste an Issue ID from the register.
- Press Enter or click Go.
- Matching is case-insensitive.
- Navigator refresh is included in the common UI refresh flow.

## Verified before packaging
- JavaScript syntax check passed.
- `refreshIssueNavigator()` exists exactly once.
- `goToIssue()` exists exactly once.
- `deleteExistingSummarySlides()` exists exactly once.
- `addCleanGeneratedSlide()` exists exactly once.
- No `TAG_ACTIONS_JSON`.
- No `window.confirm`.
- No old select-only `ui.issueNavigator` references.

## GitHub
Replace:
- public/v2/taskpane.html
- public/v2/taskpane.css
- public/v2/taskpane.js
