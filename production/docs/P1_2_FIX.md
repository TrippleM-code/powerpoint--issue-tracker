# P1.2 InvalidArgument Fix

## Symptom

PowerPoint displayed `InvalidArgument` while Create / Refresh Sheet partially rendered:

- Issue Description appeared
- Location strip and divider appeared
- Area/Room, Reference Images and Actions did not complete

## Root cause

`addThinRect()` set:

```ts
shape.lineFormat.transparency = 100;
```

PowerPoint JavaScript API requires line transparency to be between `0.0` and `1.0`.

Correct value:

```ts
shape.lineFormat.transparency = 1.0;
```

## Expected result

Create / Refresh Sheet should now continue rendering after the divider.

Manual reference images remain user-owned and are not deleted by refresh.
