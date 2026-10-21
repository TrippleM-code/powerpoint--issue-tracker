# Web P4.8.4 — notes fill compatibility

21 October 2026. A user reported InvalidArgument on Create / Refresh Sheet in Web P4.8.3. The new notes textbox applied fill transparency without first setting a solid fill, unlike existing transparent layout panels. This is the suspected Office.js compatibility cause; the desktop COM adapter did not reproduce the error.

The notes box now initializes a white solid fill before making it fully transparent. The border, placeholder, notes preservation and summary exclusion are unchanged. The host mock now rejects transparency before solid-fill initialization. Staged-render errors also expose a validated Office property/method location when available, without request bodies or document content. Staged rollback remains in place.

Existing manifests and permissions are unchanged. Reopen the pane and confirm Web P4.8.4; retry Create / Refresh Sheet on an ICE issue. The user's real Office.js host remains the acceptance test; automated mocks and native COM tests do not prove identical Office.js behavior. Installer 4.8.0.15 is unchanged by this web-only hotfix.
