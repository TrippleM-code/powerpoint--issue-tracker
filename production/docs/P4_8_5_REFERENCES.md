# Web P4.8.5 / Desktop 4.8.0.16 — register reference numbers

Type SOI, NOV and COI/CVI reference numbers directly into their existing issue-slide boxes, then generate/refresh the summary. The action register reads their current text. Leading zeroes are retained. Empty or whitespace-only fields are omitted; an entirely empty reference column is hidden across every register page. A displayed column has blank cells for issues without a value. COI/CVI remains the existing combined field.

Column widths rebalance automatically. With reference columns, eight taller action rows fit per page; without references the existing ten-row layout remains. All actions are retained. Issue-level references span the issue's action rows and repeat when that issue continues onto another page. The dashboard and slide-only notes are unchanged. Long entries can still shrink to fit.

No new stored issue-schema fields, permissions or endpoints. References are read from tagged shapes before replacing generated summaries; a failed read leaves prior summaries intact. Native snapshots expose text only for the three reference roles. Existing manifests retain their identity and URL.

Acceptance: use an ICE test deck with SOI `001234`, blank NOV and COI/CVI `CVI-001`. Generate a summary and verify only SOI and COI/CVI columns appear. Change a number and regenerate; clear all numbers and regenerate to recover the original column layout. Confirm notes remain absent from summaries and all actions survive pagination.

Validation: 129 web tests, build and lint; 109 desktop tests, build/lint and native safeguards. Real PowerPoint ICE test verified live typed reference numbers, omitted blank NOV, notes preservation and summary exclusion through save/reopen.
