# P4.7 Startup / Tabs Fix

The task pane could render while its controls were inert because settings were read from
`Office.context.document.settings` during module import, before Office initialization was guaranteed.

P4.7:
- starts from safe default settings
- reads document settings only inside `initializeIssuePanel()` after `Office.onReady`
- attaches tab click handlers before PowerPoint reads
- shows initialization errors in the task pane
