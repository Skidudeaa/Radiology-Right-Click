---
status: pending
priority: p2
issue_id: "003"
tags: [code-review, performance, ux, ahk]
dependencies: []
---

# InsertAtImpression Has 950ms Fixed Delays

## Problem Statement

The `InsertAtImpression()` function contains 950ms of hard-coded Sleep delays regardless of application responsiveness. This creates noticeable lag on every Fleischner insertion.

**Why it matters:** Users experience nearly 1 second of unnecessary wait time on fast systems where most delays are not needed.

## Findings

**Location:** `RadAssist.ahk` lines 936-1003

```autohotkey
Send, ^a  ; Select all
Sleep, 50
Send, ^c  ; Copy
ClipWait, 0.5
Send, ^{Home}
Sleep, 50
Send, ^f
Sleep, 200      ; Excessive for Find dialog
Send, ^a
Sleep, 50
Send, IMPRESSION:
Sleep, 100
Send, {Enter}
Sleep, 150
Send, {Escape}
Sleep, 50
Send, {Escape}
Sleep, 100
Send, {End}
Sleep, 50
Send, {Enter}{Enter}
Sleep, 50
Send, ^v
Sleep, 100
```

**Total fixed delays:** 950ms

**Agent:** performance-oracle

## Proposed Solutions

### Option A: Reduce Sleep Durations (Recommended)
- **Pros:** Simple change, significant improvement
- **Cons:** May need tuning for slower systems
- **Effort:** Small
- **Risk:** Low

Reduce most 50ms delays to 20ms, reduce 200ms Find dialog delay to 100ms.
Target: ~300ms total (68% reduction).

### Option B: Use SetKeyDelay
- **Pros:** Consistent timing, simpler code
- **Cons:** Global setting may affect other operations
- **Effort:** Small
- **Risk:** Medium

```autohotkey
SetKeyDelay, 20, 20
```

### Option C: Add Configurable Delay Setting
- **Pros:** Users can tune for their system
- **Cons:** Adds complexity
- **Effort:** Medium
- **Risk:** Low

## Recommended Action

(To be filled during triage)

## Technical Details

- **Affected Files:** RadAssist.ahk
- **Lines:** 936-1003
- **Components:** InsertAtImpression()

## Acceptance Criteria

- [ ] Total delays reduced to <=400ms
- [ ] Insertion still works reliably in PowerScribe
- [ ] Insertion still works reliably in Notepad
- [ ] No regression in find/insert flow

## Work Log

| Date | Action | Result/Learning |
|------|--------|-----------------|
| 2026-01-12 | Created from PR review | Initial finding |

## Resources

- PR #1: https://github.com/Skidudeaa/Radiology-Right-Click/pull/1
