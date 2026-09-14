# Testing and Acceptance Criteria

## Unit tests

- Issue ID normalization
- duplicate Issue ID rejection
- status validation
- overall issue status calculation
- action add/edit/remove
- Created At immutability
- Updated At only on meaningful changes
- pagination calculation
- equal action-row sizing
- migration functions

## PowerPoint smoke tests

- create issue
- refresh issue
- manual image survives refresh
- manual text survives refresh
- logo appears
- area/room strip renders
- add 3 actions -> 3 equal rows
- add 6 actions -> 6 equal rows
- navigate by pasted ID
- generate summary
- summary groups Issue ID / Area / Room / Description
- status colors match issue sheet
- repeated summary refresh creates no duplicate managed pages

## Production acceptance

A release is not production-ready until:
- no uncaught errors in normal workflow
- no destructive refresh of manual user content
- all tests pass
- accessibility review completed
- manifest validated
- privacy/support pages ready
- known limitations documented
