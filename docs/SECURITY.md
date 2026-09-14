# Security

## Production V1

IssueFlow is a client-side Office Add-in.

- no backend database required
- no credentials stored in code
- no arbitrary file execution
- logos limited to supported image formats
- uploaded logos must be resized/compressed before storage
- manual images remain PowerPoint content
- external links are not automatically opened
- no AI decision-making for contractual, financial, safety, or compliance matters

## Permissions

Request only the PowerPoint document permissions actually required.

## Auditability

Store:
- Created At
- Updated At
- schema version
- action IDs
- issue IDs

Future audit history can be added without changing the V1 issue model.
