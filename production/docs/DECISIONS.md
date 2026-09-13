# Decisions

## D001 — Manual reference images
Reference Images remains user-controlled PowerPoint content.
IssueFlow only manages the section header and border.

## D002 — Backend-free V1
No cloud backend in Production V1.
Reason: lower cost, simpler deployment, less security risk, easier Marketplace review.

## D003 — Structured metadata is source of truth
Generated shapes are not the canonical data source.

## D004 — Human control
IssueFlow organizes information; it does not make engineering, contractual, safety, or compliance decisions.

## D005 — No preview APIs by default
Use stable Office.js APIs where possible.
Preview APIs require explicit justification and fallback.
