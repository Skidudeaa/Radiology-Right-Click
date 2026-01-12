---
status: pending
priority: p3
issue_id: "008"
tags: [code-review, dry, refactoring, ahk]
dependencies: []
---

# GUI Positioning Code Duplicated 15+ Times

## Problem Statement

The same 4-line GUI positioning block appears in every GUI function (15+ occurrences), violating DRY principle.

## Findings

**Pattern repeated throughout:**
```autohotkey
CoordMode, Mouse, Screen
MouseGetPos, mouseX, mouseY
xPos := mouseX + 10
yPos := mouseY + 10
```

**Locations:** Lines 248-252, 318-322, 459-462, 510-514, 584-588, 654-658, 768-772, 796-799, 816-820, 837-841, 862-866, 928, 1063-1066, 2043-2047

**Agent:** pattern-recognition-specialist

## Proposed Solution

Extract to helper function:
```autohotkey
GetMousePosition(ByRef xPos, ByRef yPos, offset := 10) {
    CoordMode, Mouse, Screen
    MouseGetPos, mouseX, mouseY
    xPos := mouseX + offset
    yPos := mouseY + offset
}
```

**Effort:** Small | **Risk:** Low | **LOC Reduction:** ~50 lines

## Acceptance Criteria

- [ ] Helper function created
- [ ] All GUI functions use helper
- [ ] GUIs still appear at mouse position
