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
