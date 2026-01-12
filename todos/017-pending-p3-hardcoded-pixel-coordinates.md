---
status: pending
priority: p3
issue_id: "017"
tags: [code-review, ui, accessibility, ahk]
dependencies: []
---

# Hardcoded Pixel Coordinates (No DPI Awareness)

## Problem Statement

All GUI code uses hardcoded pixel coordinates without DPI awareness. GUIs will appear incorrectly sized on high-DPI displays.

## Findings

**Throughout GUI code:**
```autohotkey
Gui, EllipsoidGui:Add, Text, x10 y10 w280
Gui, EllipsoidGui:Add, Edit, x80 y32 w50 vEllipDim1
Gui, EllipsoidGui:Add, Text, x135 y35, x   T (W):
```

**Agent:** pattern-recognition-specialist

## Proposed Solution

Use DPI scaling or relative positioning:
```autohotkey
; Option 1: DPI scaling factor
dpi := A_ScreenDPI / 96
x := 10 * dpi

; Option 2: AHK v2 built-in DPI awareness
```

**Effort:** Medium-High | **Risk:** Medium (requires testing on various displays)

## Acceptance Criteria

- [ ] GUIs readable on high-DPI displays
- [ ] No regression on standard displays
- [ ] Consider AHK v2 for native DPI support
