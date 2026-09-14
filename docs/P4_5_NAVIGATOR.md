# P4.5 Issue Navigator

Issue Navigator is restored below the Issue metadata card.

Behavior:
- type or paste an Issue ID
- datalist suggests existing Issue IDs
- press Enter or click Go
- matching is case-insensitive
- selected issue slide is opened
- Issue tab fields are reloaded from that slide
- uses `presentation.setSelectedSlides([slideId])`

The previous unsupported `slide.select()` approach is not used.
