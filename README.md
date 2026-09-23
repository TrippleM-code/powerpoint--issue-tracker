# IssueFlow for PowerPoint — P4.8 corrective release

Revised from GitHub commit `dd7e081bcf0a773c370d67bc481a38ff10d494c5` (P4.7).
The app manages issue sheets, responsible parties, actions, dashboard and action register inside a PowerPoint presentation.

**`production/` is the only maintained application source.** Root commands delegate to it. Earlier duplicate root source, tests, docs and manifests have been removed. The Git-connected deployment target is `https://powerpoint-issue-tracker.vercel.app`.

## Set up and verify (Windows PowerShell)

Install Node.js 22.12+ (22 LTS recommended), then run from this folder:

```powershell
npm.cmd run setup
npm.cmd run lint
npm.cmd test
npm.cmd run build
```

Alternatively, inside `production/`, run `npm.cmd ci`, followed by the same lint/test/build commands. `production/package-lock.json` locks the dependency tree for both paths. The root package intentionally has no separate application dependencies.

`npm.cmd run dev` starts the browser development server and must remain running while using it. A normal browser cannot exercise PowerPoint document APIs. To verify the add-in, host the build over HTTPS and sideload its manifest in PowerPoint.

## Deploy the revised version

Use a new extracted folder, not a partial overlay of old source. Keep existing presentations and a copy of the previous project.

Vercel may use either configuration:

| Root Directory | Install command | Build command | Output |
|---|---|---|---|
| Repository root | `npm ci --prefix production` | `npm run build` | `production/dist` |
| `production` | `npm ci` | `npm run build` | `dist` |

Both configurations build the same source. Corresponding `vercel.json` files are included. If Vercel has manual dashboard overrides, align them with this table.

The source ZIP includes the verified `production/dist` output for static HTTPS hosting. Normal repository builds regenerate it; it is ignored by Git.

After the Git-triggered deployment is ready, download `/issueflow-p4-8.xml` from the production host or use `production/manifests/issueflow-production-dev-p4-8.xml`. Sideload it in PowerPoint, reopen the pane and verify the **P4.8** label. For a preview deployment, change `SourceLocation` to that preview's HTTPS `/index.html?v=p4.8` before sideloading.

If replacing a GitHub checkout, commit the removed duplicate directories as deletions along with the new files. Do not commit `node_modules`, `.env`, or `dist`.

## Corrections and acceptance

Read [the release notes](production/docs/P4_8_RELEASE.md) for fixes, test evidence, limitations and the PowerPoint acceptance checklist. The original alpha ZIP stays in `archive/`; legacy `/taskpane.html` and `/v2/taskpane.html` assets are retained under `production/public/` unchanged. They are compatibility artifacts, not the repaired production UI.

Data stays in presentation tags and Office document settings; no backend, new permission scope, paid API, or AI service was introduced.
