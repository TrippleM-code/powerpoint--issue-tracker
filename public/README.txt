PowerPoint Issue Tracker V1.1 Patch

Replace these files in your Vercel-hosted GitHub repo:
- public/taskpane.html
- public/taskpane.js
- public/taskpane.css

Fixes:
- Changing Issue ID now synchronizes existing tracker box metadata.
- New Repair Current Slide tool.
- Safer validation.
- Safer summary generation/deletion.

After Vercel redeploys:
1. Close and reopen the PowerPoint add-in task pane.
2. Select an old issue slide.
3. Click Repair Current Slide.
4. Validate Presentation.
5. Generate / Refresh Summary.
