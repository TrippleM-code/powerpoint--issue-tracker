# V2 Design

## Data model

### Issue
- Issue ID
- Description
- CreatedAt
- UpdatedAt
- Actions[]

### Action
- Action ID (`ACT-001`, `ACT-002`, ...)
- Responsible Party
- Action Required
- Status
- Remark (optional)
- CreatedAt
- UpdatedAt

## Ownership

The add-in owns structured Issue and Action data.

PowerPoint owns manual presentation content, especially the Reference / Images area. User images, arrows, callouts and annotations are not tagged as IssueFlow-managed content and are not removed during Issue Sheet refresh.

## Overall issue status

- All actions Closed -> Issue Closed
- Otherwise Open has highest priority
- Then In Progress
- Then Pending
- Custom statuses fall back to the first action status

## Summary register

Issue ID and Description are issue-level fields and are shown once across the issue's action rows. Party, Action, Remark and Status remain action-level fields.

This prevents duplicated issue information while preserving independent action tracking.
