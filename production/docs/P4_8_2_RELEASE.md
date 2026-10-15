# IssueFlow P4.8.2 — Standard and ICE Meeting

Released 15 October 2026 from tested template build 2026.10.15.4.

- Settings now offers Standard (default) and ICE Meeting. Choose a shared title, meeting number and date once per presentation. Area Code and Room / Space still come from each issue.
- Standard header and ICE footer logos retain original proportions.
- Generated geometric boxes/lines stay behind manual images on refresh; generated text stays above. No automatic shape locking is claimed.
- Standard gutters are removed and body fills are transparent. Dashboard text is larger and has no internal margins.
- Issue schema, references, actions, summary counts, conflict checks and manual-content preservation remain. Changing formats may require repositioning manual pictures once.

## Existing manifest
Keep the IT-installed XML unchanged, including ID 3f8b99b8-9bb9-46ca-b9d3-e1ee4b928f4b and https://powerpoint-issue-tracker.vercel.app/index.html?v=p4.8. No new permission or manifest is required. Reopen the pane to load Web P4.8.2. Standard issue rendering now requires API 1.8 for layering; ICE requires API 1.10 to read slide size. The unchanged manifest's 1.4 minimum does not guarantee those newer runtime capabilities; unsupported operations show an update message.

## Validation and acceptance
123 automated web tests, TypeScript/Vite build and ESLint passed. Office mocks cannot verify visual results. In a copy of a deck: select ICE in Settings, apply to all, refresh twice, verify logo proportions/images/reference values, generate summary, then save/reopen. Also check Standard. Existing test add-in settings belong to its separate identity and are not automatically copied into the production add-in.

## Limits
Standard remains 960 x 540 geometry. Ten actions appear on sheets; all actions appear in the register. Generated summaries are replaced. No multiuser lock, transaction guarantee, automatic backup or new data endpoint is introduced. Existing security reports describe the earlier release and are not a new audit of P4.8.2.
