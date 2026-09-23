## P4.8 corrective release — 2026-09-23

Stable slide-target writes; staged rendering and summaries; blocked-edit reset; deck-wide library guards; stored metadata validation; busy-state protection; overflow notice and text fitting; settings rollback; single production source; lint, lockfile and regression tests. See [P4.8 release notes](P4_8_RELEASE.md).

# Changelog

## Production V1 Foundation
- created production architecture
- froze Alpha.6.6 as reference
- defined production data model
- defined test strategy
- defined security baseline
- defined roadmap

## Foundation v1.1
- added placeholder README files so empty architecture folders appear in ZIP/GitHub

## Production V1 P1
- runnable task-pane shell
- Issue ID, Area Code, Room / Space, Description
- Created / Updated timestamp rules
- presentation-tag persistence
- safe managed-shape refresh
- manual Reference Images preserved
- 50/50 vertical Area/Room strip
- P1 validation tests
- production-development manifest template

## Production V1 P1.1
- fixed TypeScript build failure: restored `Party` and `StatusDefinition` domain exports
- added production `vercel.json` with Vite + `dist` output

## Production V1 P1.2
- fixed PowerPoint `InvalidArgument` during sheet creation
- corrected `ShapeLineFormat.transparency` from `100` to `1.0`
- this allows rendering to continue past the location divider into Area/Room, Reference Images and Actions

## Production V1 P2
- action add/edit/remove workflow
- two-click remove confirmation
- equal-height action rows
- status colors
- overall issue status derived from actions
- issue Updated timestamp changes on action changes
- P2 unit and acceptance tests

## Production V1 P2.1
- restored dedicated Set ID button in the Issue panel
- added issue ID / Created / Updated identity block on the slide
- corrected widescreen layout so all generated content stays within 960x540 slide bounds
- retained P2 actions, safe refresh, vertical Area/Room, and manual Reference Images behavior

## Production V1 P2.2
- removed unnecessary Set ID button; Save / Refresh now owns Issue ID persistence
- added Issue / Actions / Settings tabs
- Action By is now a dropdown controlled by the Settings party library
- Status is now a dropdown controlled by the Settings status library
- party/status libraries are stored with the PowerPoint document using Office document settings
- seeded party library: Architect, C&S, MEP, Main Contractor, ESCS
- protected settings entries that are already used by actions from accidental removal
- retained P2.1 slide-boundary fix and safe refresh behavior

## Production V1 P3
- Party Library supports Add Logo, Replace Logo, and Remove Logo
- P2.2 string party settings migrate automatically to structured Party objects
- logos are resized/compressed before storage
- slide header renders equal-width cells for all configured parties
- logo shown when available; party name shown as fallback
- Issue ID / Created / Updated block remains on the right

## Production V1 P3.1
- fixed Vercel TypeScript build error after Party Library changed from string[] to Party[]
- fixed Party.name comparison while editing an action
- fixed Party membership validation while saving an action
- no functional changes to P3 logo/header behavior

## Production V1 P3.2
- fixed broken party-logo display in PowerPoint
- PowerPoint ShapeFill.setImage requires raw Base64 image data
- IssueFlow had been passing the full browser data URL (`data:image/...;base64,...`)
- logo rendering now strips the data-URL prefix before calling setImage()
- existing stored logos remain compatible; users do not need to upload them again

## Production V1 P4
- added Up / Down party ordering controls
- party order is presentation-level settings and drives every issue-sheet header
- added Apply Settings to All Issue Slides
- global apply preserves manual Reference Images and issue/action data
- added Summary tab with live issue/action/open/closed preview
- added Generate / Refresh Summary
- summary creates one dashboard plus paginated action-register slides
- existing generated summary slides are replaced on refresh

## Production V1 P4.1
- added explicit no-cache headers for `/` and `/index.html`
- retained long immutable caching for hashed Vite assets
- production dev manifest now uses `?v=p4.1` cache-busting query
- prevents PowerPoint task pane from remaining on an older deployed UI after Vercel updates

## Production V1 P4.2
- added a unique task-pane entry URL: `/p4-2.html`
- added a new development add-in ID to bypass Office add-in manifest/webview cache
- added Vite multi-page build so `p4-2.html` is deployed explicitly
- added stronger no-cache headers for all HTML entry points
- P4 functionality is unchanged

## Production V1 P4.3
- fixed strict TypeScript `number | undefined` errors in Action Register column widths
- guarded paginated register page access
- fixed strict Party reorder array-index errors
- source-only strict TypeScript check passes locally
- retains P4 global settings apply, party reorder, dashboard/register and safe refresh
- uses a unique `/p4-3.html` dev entry and fresh add-in ID

## Production V1 P4.4
- tidied Party Library cards for narrow PowerPoint task panes
- Up / Down controls are grouped consistently beside each party
- one Issue ID is now locked to each IssueFlow issue slide
- changing the Issue ID on an existing issue slide is blocked
- same Issue ID can still be edited/refreshed normally
- a new Issue ID requires a new blank slide
- generated dashboard/register slides are moved to the beginning of the presentation
- dashboard is always slide 1; register pages follow
- summary-front placement requires PowerPointApi 1.8 and reports a clear compatibility error otherwise

## Production V1 P4.5
- restored Issue Navigator to the Issue tab
- navigator accepts typed or pasted Issue IDs
- navigator suggestions come from all IssueFlow issue slides
- pressing Enter or clicking Go selects the matching slide
- navigation is case-insensitive
- uses supported `presentation.setSelectedSlides([slideId])`
- P4.4 Issue ID locking, tidied Settings UI and summary-first placement remain unchanged

## Production V1 P4.6
- fixes PowerPoint 404 caused by the temporary `/p4-5.html` task-pane route
- new manifest points to the known-good `/index.html` route
- fresh add-in ID still bypasses the old Office add-in registration cache
- cache-busting query `?v=p4.6` retained
- Navigator and all P4.5 features remain unchanged

## Production V1 P4.7
- fixed task-pane tabs becoming unresponsive
- stopped reading Office document settings during JavaScript module evaluation
- presentation settings now load only after Office.onReady
- tab handlers attach before PowerPoint data reads
- startup failures now appear in the task-pane banner
- Navigator, ID locking, global settings apply, and summary-first behavior retained
