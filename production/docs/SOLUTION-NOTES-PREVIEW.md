# P4.8.8: Separate solution and optional notes

Approved release: Web P4.8.8 and desktop 4.8.0.19. Existing manifest identity, URLs and permissions are unchanged.

## Behavior

- Standard and ICE now use an optional issue-level Solution field, separate from each action's Action Required text.
- Description and Solution start at the top of their areas, leaving the lower areas for manual images and text.
- Action Required text appears in the summary register, not the issue sheet. Action By and Status remain on the issue sheet.
- Independent checkboxes request a Description notes box or Solution notes box on Create / Refresh Sheet. Both default off for old/new issues without saved choices.
- Existing boxes retain their text, formatting, position and size after refresh, bulk settings, and unticking. Unticking never deletes a box. Delete a box directly in PowerPoint; leave its option off to avoid creating a fresh box on the next refresh.
- Notes remain slide-only and excluded from summaries. Existing legacy NOTES boxes are recognized as Description notes without moving or resetting them.
- Closed Action Required text is grey with native strikethrough in the summary. Reopening the action restores normal text on the next summary refresh.
- Native strikethrough requires PowerPointApi 1.8 in the web host. Older summary-capable hosts retain grey text without strikethrough. Existing issue-sheet rendering already requires 1.8; ICE slide-size support requires 1.10.

## Data and safety

Optional `solution`, `descriptionNotes`, and `solutionNotes` fields preserve legacy metadata compatibility. Optional values are type-validated and included in conflict detection. Notes are user-owned shapes, excluded from managed replacement. New notes take part in staging rollback when rendering fails.

No new dependencies, external endpoints, permissions, or uploads. Existing action records are retained; they are not converted into solutions. Empty solutions stay empty until the user enters one.

## Validation

- 133 unit/workflow tests pass, including independent notes options, creation once, untick persistence, manual deletion, regeneration, summary exclusion, reopening actions, and solution conflict checks.
- TypeScript build and ESLint pass.
- A disposable native PowerPoint harness exercises the actual web renderer through the local adapter, saves and reopens both formats, exports slides, and verifies true strikethrough/color and note survival.
- Desktop bridge adds Font2.Strike support and disables initial textbox auto-sizing before text assignment. These changes are included in desktop 4.8.0.19.

## Acceptance checklist

1. Review both preview decks, especially the issue sheet and Action Register.
2. In a test add-in deployment, enter distinct Description, Solution, and Action Required values. Refresh and confirm the correct placement.
3. Start with both notes checkboxes off: no new notes boxes.
4. Tick one, refresh, type notes, then untick and refresh. The same box and text must remain.
5. Repeat for the other notes box. Move/resize boxes and paste an image, then refresh and verify they remain untouched.
6. Delete a notes box with its option off and refresh: it stays absent. Tick and refresh to recreate it.
7. Close an action and refresh the summary, then reopen it and refresh again.
8. Save/reopen the presentation and repeat refresh.

Visual review approved. Actual Office.js task-pane acceptance is to be performed by the user after release using the checklist above.
