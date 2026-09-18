# IssueFlow Stability Policy

## Freeze-and-patch rule
1. Do not modify a working feature unless the current requirement explicitly depends on it.
2. Prefer isolated tests/modules over rewriting shared files.
3. Scope fixes to the failing component.
4. Run regression checks before replacing Stable.
5. Keep one Stable build and one Development line.
6. Preserve manual PowerPoint content during managed refresh.
7. Do not change manifest/deployment paths unless required.

## Protected working areas
- Issue workflow
- Actions workflow
- Party/Status libraries and logos
- Manual Reference Images
- Managed-shape refresh
- Navigator
- Apply Settings to All
- Summary / Action Register
- Current slide geometry/layout

## Release gate
A Development change replaces Stable only after build, unit tests, smoke tests and manual-content protection all pass.
