# P3.2 Party Logo Display Fix

## Symptom

PowerPoint showed `The picture can't be displayed` in party header cells.

## Root cause

Browser logo storage uses a data URL:

`data:image/png;base64,iVBOR...`

PowerPoint `ShapeFill.setImage()` expects only the Base64-encoded image data:

`iVBOR...`

## Fix

P3.2 converts stored data URLs to raw Base64 immediately before rendering the logo.

Existing document settings are unchanged, so previously uploaded logos should render after refreshing the issue sheet. Re-uploading should not be necessary.
