# IssueFlow P5.1 Regression Checklist

**Rule:** if a test fails, fix only the failing area. Do not refactor unrelated working modules.

## Task pane
- [ ] Add-in opens without error.
- [ ] Issue, Actions, Summary and Settings tabs all work.
- [ ] Repeated tab switching does not freeze the task pane.
- [ ] Navigator works by typed ID, pasted ID, Enter and Go.
- [ ] Unknown ID shows an error without moving slide.

## Issue creation / refresh
- [ ] Blank slide allows a new Issue ID.
- [ ] Existing issue slide keeps its Issue ID.
- [ ] Different Issue ID cannot replace an existing issue slide.
- [ ] Same-ID Area, Room and Description edits refresh correctly.
- [ ] Created timestamp stays unchanged.
- [ ] Updated timestamp changes after meaningful edits.

## Manual content protection
- [ ] Manual image survives refresh.
- [ ] Manual text survives refresh.
- [ ] Manual arrow/shape survives refresh.
- [ ] Only IssueFlow-managed shapes are rebuilt.

## Actions
- [ ] 1, 3 and 6 actions render with equal rows.
- [ ] Action By, Action Required and Status edits work.
- [ ] Remove action uses two-step remove.
- [ ] Open / In Progress / Pending / Closed colors remain correct.
- [ ] Overall status remains correct.

## Settings
- [ ] Party and Status libraries load.
- [ ] Add Party / Status works.
- [ ] Duplicate Party / Status is rejected.
- [ ] Add / Replace / Remove Logo works.
- [ ] Party Up / Down works.
- [ ] Apply Settings to All updates issue slides.
- [ ] Apply Settings to All preserves manual content.

## Summary
- [ ] Summary counts Issues and Actions correctly.
- [ ] Dashboard is slide 1.
- [ ] Register pages follow Dashboard.
- [ ] Issue slides follow generated summary pages.
- [ ] Register contains Issue ID, Area/Room, Description, Action By, Action Required and Status.
- [ ] More than 8 action rows paginate correctly.
- [ ] Regenerate Summary removes old managed summary pages.
- [ ] Regenerate Summary creates no duplicates.
- [ ] Regenerate Summary does not alter issue slides.

## Release gate
- [ ] `npm test` passes.
- [ ] `npm run build` passes.
- [ ] All smoke tests pass.
- [ ] No destructive refresh is observed.
- [ ] No unrelated working UI was changed.
