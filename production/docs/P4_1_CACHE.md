# P4.1 Task Pane Cache Control

## Problem
Vercel deployed the new build successfully, but PowerPoint continued showing an older IssueFlow task-pane version.

## Fix
- `/` and `/index.html` now use `Cache-Control: no-store, no-cache, must-revalidate, proxy-revalidate`.
- `/assets/*` remains long-lived because Vite filenames are content-hashed.
- The development manifest points to:
  `https://powerpoint-issue-tracker.vercel.app/?v=p4.1`

## After deploying P4.1
Replace the old Production Dev manifest with the updated manifest from `manifests/`.
Close and reopen PowerPoint once. Future production-dev UI releases should refresh more reliably.
