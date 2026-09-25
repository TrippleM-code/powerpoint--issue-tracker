# IssueFlow Web P4.8.1 — 26 September 2026

## Changes

- Dashboard cards show net changes since the previous summary refresh: + increase, − decrease, and — unchanged or first baseline. These are count changes, not an audit history.
- Issue sheets include blank SOI, NOV and COI/CVI reference boxes. Type reference numbers directly into these PowerPoint text boxes; their contents survive layout refreshes. Created and Updated dates remain visible.
- Saving/rendering checks the loaded issue revision. Rendering checks again after staging replacement shapes and rolls back staged shapes when a change is detected.
- Updated app icon and Web P4.8.1 labels; hosted manifest version 1.0.0.9, with the existing add-in ID and filename retained.

This is an Office.js web add-in. Windows COM code, installer operations, registry changes and native local backup are not included. Summary refresh does not create backups. Use PowerPoint Save a Copy or your organization's version history for web-version backups.

## Verification

72 automated tests passed, including dashboard comparisons, manual reference preservation, stale edits and a change detected during staged rendering. ESLint and TypeScript/Vite production build passed. These tests use Office mocks; they do not replace a live PowerPoint acceptance test.

## PowerPoint acceptance — use a copy

1. Reopen the web add-in and confirm Web P4.8.1. Existing manifest URLs remain valid; if necessary sideload the hosted issueflow-p4-8.xml manifest again.
2. Render an issue. Fill the three reference boxes on its slide, change an action, and refresh its layout. Confirm references, manual pictures and dates remain correct.
3. Refresh the summary to establish a baseline. Change one action status and refresh again. Confirm the relevant card changes; another refresh without edits should show —.
4. Load the same issue in two panes. Save a change in one, then try saving the stale issue from the other. It should ask you to reload. Reload before entering the second edit.
5. Save, close and reopen the presentation. Verify issue data, reference numbers and generated summaries persist.

## Remaining limitations

Use one editor at a time. Office.js has no atomic compare-and-swap or multiuser lock; a small race remains between the final check and write. Staged changes are not database transactions. Inspect any cleanup failure before retrying. Summaries reflect the issues read when generation started.

Generated summaries are replaced; keep permanent notes on ordinary slides. Issue sheets show the first ten actions plus an overflow count; the register contains all actions. Fixed widescreen geometry and small dates/reference boxes may need manual readability checks for long text. Existing reference boxes keep their placement and formatting, including user edits.

Internet access is required for the Vercel page and Microsoft Office.js. This release does not change document permissions or introduce an application backend, telemetry service or AI service.
