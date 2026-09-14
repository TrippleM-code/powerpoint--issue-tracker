# P4.2 Office Cache Reset Strategy

PowerPoint continued loading the old task pane even after Vercel deployment and query-string cache busting.

P4.2 therefore uses BOTH:
1. a completely new add-in ID
2. a completely new task-pane URL path

New add-in ID:
`7c73d6fa-1a57-4a6f-9d1f-4b8b64fd0c42`

New task-pane URL:
`https://powerpoint-issue-tracker.vercel.app/p4-2.html`

This avoids relying on PowerPoint honoring query-string cache changes for an existing add-in registration.

## Install
1. Deploy P4.2.
2. Remove/ignore the prior Production Dev manifest from the trusted catalog.
3. Add `issueflow-production-dev-p4-2.xml`.
4. Restart PowerPoint.
5. Add `IssueFlow Production Dev P4.2` from Shared Folder.
