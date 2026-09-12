PowerPoint Issue Tracker V1.3

New:
- Dashboard is kept at the beginning of the presentation.
- Professional KPI cards and status visualization.
- Donut chart is attempted as a locally-generated image; a shape-based status view is used if the host does not support picture insertion.
- Professional paginated issue register using real PowerPoint tables (PowerPointApi 1.8).
- Extra register pages are automatically created when the issue list exceeds one page.
- Old generated summary/register pages are removed on refresh, so duplicates are not accumulated.
- Issue slides move down automatically when the refreshed summary block is placed at the front.
- Visible register no longer uses a slide/location column.
- Issue ID is the user-facing identity.
- Created date is written once and kept.
- Updated date appears only after Description / Status / Remark changes after the first V1.3 baseline refresh.
- Managed footer shows ID, Created date, and Updated date only when one exists.
- Issue Navigator in task pane jumps to an Issue ID reliably.

Hyperlink note:
The current PowerPoint JavaScript API exposes external hyperlink addresses but does not officially expose the native internal slide SubAddress/relationship target used for PowerPoint's "jump to slide" links. V1.3 therefore:
1. tries PowerPoint's documented internal action URI on the visible Issue ID, and
2. provides the Issue Navigator as the guaranteed fallback.
Do not rely on the internal click behavior until verified on your exact PowerPoint build.

Update GitHub/Vercel:
- public/taskpane.html
- public/taskpane.css
- public/taskpane.js

After deployment:
1. Close/reopen the add-in pane.
2. Click Generate / Refresh Summary.
3. Confirm Dashboard becomes Slide 1 and registers follow.
4. Edit an issue, refresh again, and confirm Updated date appears.
5. Test clicking an Issue ID. If native internal linking is not honored by your PowerPoint build, use Issue Navigator > Go.
