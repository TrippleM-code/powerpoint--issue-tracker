# P2 Acceptance Checklist

## Action editor
- [ ] Add Action requires party, action text and status
- [ ] New action is saved to issue metadata
- [ ] Edit loads existing action into the editor
- [ ] Save Action updates only that action
- [ ] Remove requires two clicks
- [ ] Cancel Edit resets the editor

## Action rendering
- [ ] 1 action uses the full Actions height
- [ ] 3 actions create exactly 3 equal-height rows
- [ ] 6 actions create exactly 6 equal-height rows
- [ ] No unused blank action rows are generated
- [ ] Open is light red
- [ ] In Progress is light blue
- [ ] Pending is light amber
- [ ] Closed is light green
- [ ] Unknown status has neutral styling
- [ ] Overall status is derived from actions

## Timestamps
- [ ] Adding an action updates Issue Updated
- [ ] Editing an action updates Issue Updated
- [ ] Removing an action updates Issue Updated
- [ ] Existing Issue Created timestamp does not change
- [ ] Existing Action Created timestamp does not change when edited

## Regression
- [ ] P1 issue fields still save
- [ ] vertical Area / Room still renders
- [ ] Reference Images manual content survives refresh
- [ ] repeated refresh does not duplicate managed shapes
