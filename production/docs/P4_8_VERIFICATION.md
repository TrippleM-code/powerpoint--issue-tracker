# P4.8 local verification results (before publication)

Verified on Windows with Node.js v26.2.0. CI is configured for Node.js 22; that remote CI run has not been triggered by this local delivery.

| Check | Result |
|---|---|
| Root `npm test` | 59 tests passed in 9 files using Vitest 4.1.11 |
| Root `npm run lint` | Passed |
| Root `npm run build` | TypeScript check and Vite production build passed |
| Dependency audit after patched test-runner installation | 0 known vulnerabilities reported |
| Whitespace check | Passed |
| Actual PowerPoint host | Not run; complete the release-note acceptance checklist |
| GitHub push / hosted deployment | Not performed |

The original baseline contained 34 passing tests. This release adds 25 automated cases for workflow/data-integrity boundaries and settings persistence. DOM tests exercise the real panel handlers; PowerPoint service tests use mocked Office objects and failure injection. They do not replace real-host checks of typography, images, host queue behavior or persistence after reopening a deck.

The ZIP contains a generated static build, source, lockfile, manifest and documentation. The packaging step checks ZIP integrity, source-to-archive byte equality and preservation of the relocated legacy route assets.
