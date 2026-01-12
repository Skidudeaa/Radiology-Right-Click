---
status: pending
priority: p3
issue_id: "009"
tags: [code-review, dry, refactoring, ahk]
dependencies: []
---

# Clipboard Save/Restore Pattern Duplicated 8+ Times

## Problem Statement

The clipboard save/restore pattern is repeated 8+ times throughout the codebase.

## Findings

**Pattern repeated:**
```autohotkey
ClipSaved := ClipboardAll
Clipboard := textToInsert
ClipWait, 0.5
Send, ^v
Sleep, 100
Clipboard := ClipSaved
```

**Locations:** Lines 913-923, 936-1000, 1836-1844, 1857-1877, 2041

**Agent:** pattern-recognition-specialist

## Proposed Solution

Extract to helper function:
```autohotkey
PasteText(text, waitMs := 100) {
    ClipSaved := ClipboardAll
    Clipboard := text
    ClipWait, 0.5
    Send, ^v
    Sleep, %waitMs%
    Clipboard := ClipSaved
}
```

**Effort:** Small | **Risk:** Low

## Acceptance Criteria

- [ ] Helper function created
- [ ] All paste operations use helper
- [ ] Clipboard properly restored after each paste
