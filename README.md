# IssueFlow for PowerPoint — Production V1 Foundation

This package starts the production rebuild of IssueFlow while keeping the latest alpha frozen as a reference.

## Production goal

A Microsoft PowerPoint Office Add-in for structured issue and action management:

- Issue ID
- Area Code
- Room / Space
- Description
- Manual Reference Images area
- Multiple actions
  - Action By
  - Action Required
  - Status
- Party logo library
- Created / Updated timestamps
- Dashboard + grouped Action Register
- Issue navigation by typed/pasted Issue ID

## Production principles

- Keep alpha behavior as reference; do not patch the alpha inside production.
- Separate UI, PowerPoint API, domain logic, layout, storage, and validation.
- Avoid preview Office.js APIs unless there is no stable alternative.
- Keep manual PowerPoint content untouched.
- Use structured metadata as source of truth.
- Add migrations for future schema changes.
- Fail safely: no destructive action without clear intent.
- Keep V1 production backend-free unless a cloud feature creates clear value.

## Recommended stack

- TypeScript
- Vite
- Office.js
- Plain HTML/CSS (no UI framework for V1)
- Vitest for unit tests
- ESLint + Prettier
- Vercel for HTTPS hosting

## Build phases

1. Foundation
2. Data + PowerPoint services
3. Issue sheet renderer
4. Action management
5. Summary / register
6. Validation + migrations
7. Accessibility + error states
8. Marketplace readiness

See `docs/` for the detailed production plan.

## P1.1 deployment hotfix

If P1 failed on Vercel with:

`TS2305: Module '../domain/models' has no exported member 'Party'`

replace the Production P1 files with this P1.1 package. The missing shared domain types are restored.
