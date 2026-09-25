# Web security and IT deployment guidance — P4.8.1

IssueFlow is a static Office.js PowerPoint add-in hosted on Vercel over HTTPS. It requires internet access. Microsoft Office.js is loaded from Microsoft's CDN. The web version does not install EXE/DLL files, modify the Windows registry or require local administrator rights; organizational add-in deployment remains subject to IT policy.

## Data and permissions

- The manifest requests ReadWriteDocument so the add-in can read/write issue tags, shapes and generated slides in the open presentation. This is a meaningful trust grant to the hosted code.
- Issue JSON is stored in slide tags; settings and resized logos use Office document settings. Reference numbers are ordinary slide text boxes. They travel with the presentation and are not encrypted separately by this app.
- No application backend, database, credential store, AI service or application telemetry endpoint is implemented. Hosting/CDN requests still disclose ordinary request metadata to the hosting providers. Office/OneDrive sharing and storage follow the user's Microsoft configuration.
- Presentation collaborators who can edit the file can change its metadata. Tags are not a security boundary or immutable audit trail.
- User-facing strings use DOM text operations or escaped HTML. Logo inputs are limited to PNG/JPEG and resized before storage. Remote logo URLs are not accepted as settings data.

## Operational controls

IT should approve the publisher, manifest host and requested permission, protect GitHub/Vercel administration with MFA and least privilege, review source changes before production releases, and retain a rollback commit. Hosted code can change without reinstalling the manifest, so repository and deployment access are part of the trust boundary.

Use the organization's approved PowerPoint versions and centrally managed add-in deployment. The manifest keeps its established add-in ID; version 1.0.0.9 identifies this update. Remove the add-in through Office/organizational deployment controls to uninstall it. Removing it does not remove existing presentation content.

Backups are user-managed through PowerPoint Save a Copy or approved version history. Summary refresh does not create a backup. Use one editor at a time: optimistic stale-edit checks reduce accidental overwrites but are not a lock or atomic transaction.

## Assurance limits

Automated tests and source review cannot certify the hosting account, Office installation, tenant policy or future releases. Live PowerPoint acceptance, tenant deployment approval, account controls, provider logging/retention and incident response ownership remain IT responsibilities. Record the reviewed commit and deployment with each security report. Never describe this hosted version as fully offline or guaranteed secure.
