PowerPoint Issue Tracker V1.2 patch

Fixes:
- Validation uses one-issue-per-slide logic and no longer rejects tracker boxes due to stale shape Issue IDs.
- Summary generation fixed: SlideCollection.add() returns void, so the code now retrieves the newly added slide after sync.

Replace this file in GitHub:
public/taskpane.js

Then wait for Vercel to redeploy, close/reopen the add-in pane, and test Validate + Generate Summary.
