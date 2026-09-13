# IssueFlow V2 Alpha

This is a **side-by-side alpha**. It does not require replacing the stable V1.4.1 files.

## What V2 adds

- Modern tabbed task-pane UI: Issue / Actions / Summary / Settings.
- Branding footer: **Brought to you by Daily Painkiller + Codex**.
- `Set ID -> Create / Refresh Sheet` workflow.
- Structured Issue data: Issue ID, Description, Created, Updated.
- Multiple structured actions per issue:
  - Responsible Party
  - Action Required
  - Status
  - Optional Remark
- Managed Party and Status libraries.
- Automatic overall issue status derived from its actions.
- Manual Reference / Images area: the add-in never manages user images or annotations.
- Summary Dashboard.
- Grouped Action Register where the same Issue ID and Description are visually merged across its action rows.
- Generated summary pages are rebuilt as clean pages and moved to the front.
- Existing V1 Issue ID and Description can be reused; V2 will try to migrate a legacy Description box when no V2 description has been saved.

## Safe test installation

Keep V1.4.1 as your working stable version.

Upload this folder to GitHub:

```text
public/
  v2/
    taskpane.html
    taskpane.css
    taskpane.js
```

Do **not** replace the existing root `public/taskpane.*` files yet.

Then add `manifest-v2-alpha.xml` to your trusted add-in catalog. It has a different add-in ID and points to:

```text
https://dailypainkiller-powerpoint-issue-tr.vercel.app/v2/taskpane.html
```

You should then see a second add-in named **IssueFlow V2 Alpha** while your stable V1.4.1 stays available.

## First test

1. Open a test presentation and a normal issue slide.
2. Open **IssueFlow V2 Alpha**.
3. Set `MEP-001`.
4. Enter a Description.
5. Click **Create / Refresh Sheet**.
6. Paste images manually into the Reference / Images area.
7. Add 2-3 actions under the Actions tab.
8. Refresh the Issue Sheet and confirm the manual images remain.
9. Generate the Summary.
10. Confirm the Action Register shows one merged Issue ID / Description with separate Party / Action / Remark / Status rows.

## Important Alpha limitation

Structured data should be edited through the task pane. Manually changing add-in-generated Description or Action text on the slide is presentation-only and can be overwritten on the next refresh. Manual reference images/annotations are intentionally preserved.
