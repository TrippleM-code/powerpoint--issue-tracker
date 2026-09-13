# Architecture

```text
Task Pane UI
   |
   v
Application Controllers
   |
   +--> Domain Services
   |     - Issue rules
   |     - Action rules
   |     - Status rules
   |     - Validation
   |
   +--> PowerPoint Service
   |     - selected slide
   |     - tags
   |     - shapes
   |     - navigation
   |     - generated slides
   |
   +--> Layout Renderers
   |     - issue sheet
   |     - dashboard
   |     - action register
   |
   +--> Storage / Migration
         - presentation metadata
         - schema version
         - party library
         - status library
         - logos
```

## Separation rules

UI must not directly manipulate PowerPoint shapes.
Layout renderers must not own business rules.
PowerPoint service must not contain presentation-specific UI logic.
Domain logic must be testable without Office.js.
Storage must version all structured metadata.
