# IssueFlow V2 Alpha.6.6

Safe Refresh / Manual Reference Images patch.

## Problem fixed
When users pasted images, text, arrows or annotations into Reference Images,
refreshing the issue sheet recreated a large filled white body shape above
those manual objects, visually hiding them.

## Fix
- Reference Images no longer gets a filled background shape.
- Only four thin border dividers are generated around the manual area.
- Untagged/manual PowerPoint objects are never deleted by refresh.
- The Reference Images header is still managed by IssueFlow.
- Manual pasted images, text, arrows, links and annotations remain visible.

## No other layout changes
Area/Room, Actions, header logos, status styling and summary logic are unchanged.

## GitHub update
Replace:
- `public/v2/taskpane.html`
- `public/v2/taskpane.js`

CSS is unchanged.

## Test
1. Confirm V2 Alpha.6.6.
2. Paste an image into Reference Images.
3. Add a manual text box and arrow on top of the image.
4. Click Create / Refresh Sheet.
5. Confirm the image, text and arrow remain visible.
6. Repeat refresh two or three times.
7. Confirm manual content is still preserved.
