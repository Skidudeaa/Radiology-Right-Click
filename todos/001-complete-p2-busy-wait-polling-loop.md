---
status: pending
priority: p2
issue_id: "001"
tags: [code-review, performance, ahk]
dependencies: []
---

# Busy-Wait Polling Loop in Confirmation Dialog

## Problem Statement

The `ShowParseConfirmation()` function uses a busy-wait loop that continuously polls a global variable every 50ms, consuming CPU cycles while waiting for user input.

**Why it matters:** This pattern prevents Windows message processing during the wait, could cause the application to appear unresponsive, and consumes ~2% CPU continuously while waiting.

## Findings

**Location:** `RadAssist.ahk` lines 1091-1095

```autohotkey
g_ConfirmAction := ""
while (g_ConfirmAction = "") {
    Sleep, 50
}
return g_ConfirmAction
```

**Agent:** performance-oracle

## Proposed Solutions

### Option A: Use WinWaitClose (Recommended)
- **Pros:** Event-driven, no CPU polling, native AHK pattern
- **Cons:** Requires restructuring dialog flow
- **Effort:** Medium
- **Risk:** Low

```autohotkey
Gui, ConfirmGui:Show, ...
WinWaitClose, ahk_id %ConfirmGuiHwnd%
return g_ConfirmAction
```

### Option B: Use SetTimer with Callback
- **Pros:** Non-blocking, allows other processing
- **Cons:** More complex callback pattern
- **Effort:** Medium-High
- **Risk:** Medium

## Recommended Action

(To be filled during triage)

## Technical Details

- **Affected Files:** RadAssist.ahk
- **Lines:** 1091-1095
- **Components:** ShowParseConfirmation(), confirmation dialog handlers

## Acceptance Criteria

- [ ] Confirmation dialog does not use busy-wait loop
- [ ] CPU usage remains near 0% while dialog is open
- [ ] Dialog still returns correct action (insert/edit/cancel)
- [ ] No regression in confirmation flow

## Work Log

| Date | Action | Result/Learning |
|------|--------|-----------------|
| 2026-01-12 | Created from PR review | Initial finding |

## Resources

- PR #1: https://github.com/Skidudeaa/Radiology-Right-Click/pull/1
- AutoHotkey WinWaitClose docs
