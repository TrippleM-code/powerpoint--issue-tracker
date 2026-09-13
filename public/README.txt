V1.3.1 cache-busting package

Replace on GitHub:
public/taskpane.html
public/taskpane.js
public/taskpane.css

The HTML now loads:
taskpane.css?v=1.3.1
taskpane.js?v=1.3.1

The task pane title visibly shows v1.3.1 so you can confirm the new version loaded.

Also replace your local trusted-catalog manifest with:
PowerPoint_Issue_Tracker_manifest_V1_3_1.xml

Then fully close PowerPoint and reopen the add-in.
If PowerPoint still shows the old UI, clear the Office web add-in cache:
%LOCALAPPDATA%\Microsoft\Office\16.0\Wef
(renaming the Wef folder is safer than deleting it).
