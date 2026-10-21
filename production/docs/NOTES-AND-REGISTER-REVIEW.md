# Web P4.8.3 — slide notes and readable register

Release prepared 21 October 2026 for production. Offline counterpart: desktop 4.8.0.15. The existing manifest ID, SourceLocation and permissions are unchanged. Reopen the pane and check Web P4.8.3; use Create / Refresh Sheet on existing issue slides, and regenerate the register to see the changes.

## Use

- Create/refresh an issue sheet in Standard or ICE format. A bordered text box containing **To type Notes** appears near the bottom of the reference area (Standard) or description area (ICE).
- Click inside the box and replace the placeholder with your own notes. Move, resize or format it using PowerPoint as needed.
- Refreshing the issue, editing its actions, applying settings or switching templates preserves that same box, including its text, formatting and position.
- Notes belong to the slide and are saved with the presentation. They are not issue/action data and are excluded from the dashboard and action register.
- Existing slides get the box on their next issue refresh. If the box is deleted, a fresh placeholder is created on the next refresh. Clearing its text leaves it empty across refreshes.
- The action register uses 11-point headers/body text at the standard 960-point width, with proportional scaling for ICE slides. Ten taller rows fit on each page; all additional actions continue onto subsequent pages.

## Test in PowerPoint before release

Use a copy of a presentation. Test Standard and ICE separately:

1. Generate an issue sheet and check that the bordered notes box is visible.
2. Replace the placeholder with several lines, move/resize the box and change the font.
3. Refresh twice, edit an action and apply settings to all slides. Confirm the same notes remain without duplicate boxes.
4. Save, close and reopen the deck; verify the notes.
5. Generate a register with more than ten actions. Check font readability, page breaks and inclusion of every action. Notes must not appear in the register/dashboard.

## Validation and limits

126 automated web tests, ESLint and the TypeScript/Vite build passed. Regression tests cover notes in both formats, preservation across refresh/settings/template changes, summary exclusion, pagination and rollback after failed staging. These mocks do not confirm the final Office rendering; the PowerPoint checks above remain necessary.

No new network endpoint, permission or manifest change is introduced. Notes use a dedicated shape tag and remain outside managed shape replacement and issue metadata. Existing pictures and manually added text are retained. The initial notes box may overlap an existing picture or annotation; move it once if necessary. Template changes preserve its position rather than relocating it. Very long text can still shrink to fit; resize the notes box or shorten register entries when needed. Standard slide geometry remains 960 x 540 points.
