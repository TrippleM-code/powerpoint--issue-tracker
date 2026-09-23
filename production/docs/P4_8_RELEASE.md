# P4.8 correction report

Date: 2026-09-23. Base: `dd7e081bcf0a773c370d67bc481a38ff10d494c5`.

## Corrected behavior

1. **Wrong-slide writes:** every issue save/render requires the expected slide ID. The service checks the current selection, resolves a stable slide object and validates existing ownership. Changing selection during later awaits cannot redirect the write. Action removal now follows the same checks and preserves unsaved form edits. Generated summary slides cannot be converted into issue slides by these writes.
2. **Blocked Edit overwriting an action:** edit state is reset before checking party/status availability, and assigned only after successful validation. Saving a stale action ID raises an error.
3. **Failed refresh deleting the previous layout:** new shapes are staged and synced before replacing the previous managed shapes and metadata. Staging failures trigger best-effort removal of temporary shapes and retain the old layout. Summary replacement builds and moves all new pages before removing old summaries.
4. **Library removal breaking other slides:** party and status removals inspect every validated issue in the deck. If any metadata is invalid, the operation stops rather than treating an incomplete scan as safe.
5. **Malformed metadata crashing reports:** validate required fields, dates, actions and unique action IDs at the storage boundary. Unknown schema versions and corrupt JSON produce an error identifying the slide. Unversioned valid v1 records remain readable. Invalid slides are not silently excluded from summaries.
6. **Concurrent panel operations:** mutation buttons and form inputs are disabled while operations run, and handlers guard against re-entry. Stale selection refresh responses are discarded; selection is checked again after operations finish.
7. **Duplicate copied slides:** uniqueness checks include existing issue slides and exclude only the target slide. A duplicated IssueFlow slide is rejected rather than silently accepted; create a new blank slide with a new ID.
8. **Overflow:** show the first ten actions on an issue sheet with an explicit remaining-action count. All actions stay in the Actions tab, stored metadata and paginated Action Register. Text boxes fit text within their bounds; description uses the available column height. No action data is truncated.
9. **Failed settings persistence:** restore the previous in-memory document settings when Office reports a save failure; the UI reloads those settings instead of keeping the rejected change.
10. **Build and maintenance:** one application under `production/`; root delegates to it. Added ESLint configuration, dependency lock, full CI triggers, lint gate, ignore rules and updated version/manifest. Legacy alpha/static paths are preserved separately.

## Verification

- Automated unit, DOM workflow and mocked Office integration suite: 59 tests passed in 9 files; see [verification results](P4_8_VERIFICATION.md).
- TypeScript checking, production build and ESLint run through the root commands.
- Dependency audit run against the locked dependency tree.
- Tests cover selection changes, dynamic control locking, blocked edits, deck-wide party/status use, malformed metadata, settings rollback, replacement failure, manual shape preservation and 18-action register pagination.
- The test-runner dependency was updated to patched Vitest 4.1.11 following [GHSA-82fw-gwwq-j7x9](https://github.com/advisories/GHSA-82fw-gwwq-j7x9). This is development tooling, not a new runtime feature.

## Architecture and data

- UI: `src/ui/issue-panel.ts`, DOM state and operation guards.
- Domain: `src/domain/`, issue/status rules and stored-metadata validation.
- PowerPoint adapter: `src/services/powerpoint-service.ts`, stable slide targeting and orchestration.
- Staging helper: `src/services/staged-render.ts`, temporary shape cleanup on staging failure.
- Settings: `src/services/settings-service.ts`, presentation-scoped party/logo/status library.
- Issue data remains JSON in existing `ISSUEFLOW_*` tags using schema version 1. No migration rewrites are applied automatically.
- Read/write document permission remains necessary; no new external data service or credential is used.

## Limits that still need host verification

- Automated host tests use mocks. No claim is made that this release has passed the checklist below inside real PowerPoint.
- Office.js batches are not database transactions. If the host fails during the final metadata/old-content cleanup, new content may coexist with old content or the final write may be partial. The UI reports this explicitly; inspect the slide, reload and retry. Staging prevents deletion of old content before replacement creation succeeds; it cannot guarantee atomicity across all Office failures.
- Existing layout coordinates assume 960 × 540 points (13⅓ × 7½ inches, 16:9). Other slide sizes are not made responsive in this repair.
- Text fitting can make very long content small. The ten-row issue-sheet limit is a preview limit; use the complete register for all actions. This release does not create issue-sheet continuation pages.
- PowerPointApi 1.4 is required for the pane's layout functionality; summary movement requires 1.8. These are checked before use. [Microsoft text-frame API reference](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.textframe?view=powerpoint-js-1.4).
- Simultaneous editing by multiple add-in instances/users is not transactionally locked. In-pane operations are serialized, but no collaborative conflict-resolution system was added.
- Invalid metadata is reported, not automatically repaired. Preserve the deck and investigate the affected slide before discarding anything.
- Older retained alpha assets are unchanged and are outside this corrective release's application tests.

## PowerPoint acceptance checklist (run on a copy of a deck)

- [ ] Deploy to HTTPS, sideload the P4.8 manifest, and confirm the P4.8 label.
- [ ] Create an issue, add/edit/remove actions, save, close and reopen the presentation; verify fields, dates and actions.
- [ ] Confirm Remove on issue A after selecting B; ensure neither issue is overwritten incorrectly.
- [ ] Click Edit on an action with a missing library value; verify adding another action preserves the original.
- [ ] Attempt removing a party/status used only on another slide; verify it is blocked.
- [ ] Refresh an issue containing manually inserted images, arrows and text; verify each survives.
- [ ] Render 1, 6, 10 and 18 actions and long descriptions. Check fit and the overflow message.
- [ ] Generate and regenerate summaries; all 18 actions appear once, dashboard is first, register pages follow, and issue slides remain intact.
- [ ] Apply Settings to All; verify all target slides update and manual content survives.
- [ ] In a controlled failure test, check the previous layout/summary survives a staging error and temporary content is cleaned up or explicitly reported.
- [ ] Verify a corrupted/unsupported-schema slide produces a named error before summary generation or library removal.

## Release and rollback

The Git-connected deployment target is powerpoint-issue-tracker.vercel.app, using the production root. Check the deployment status for the current commit before loading the manifest. Keep the prior release and a copy of the deck; rollback the hosted static build to restore the previous app. Existing v1 data is retained, so this correction does not require an irreversible data migration. PowerPoint undo or a backup deck is still needed for document edits; changing the hosted app does not undo changes already saved in a presentation.

## Value demonstrated

Regression tests now explicitly enforce the five reproduced bug boundaries, plus duplicate IDs, operation locking, metadata schema checking and overflow preservation. No new backend cost or AI dependency is introduced. Real-user time savings and host compatibility remain to be measured in the acceptance run.
