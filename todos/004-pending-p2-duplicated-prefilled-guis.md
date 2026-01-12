---
status: pending
priority: p2
issue_id: "004"
tags: [code-review, dry, refactoring, ahk]
dependencies: []
---

# Duplicated Prefilled GUI Functions

## Problem Statement

Five "Prefilled" GUI functions are nearly identical to their base counterparts, duplicating ~130 lines of code. Each prefilled version only differs in setting initial values for Edit controls.

**Why it matters:** This duplication creates maintenance burden - any change to a GUI must be made in two places, risking divergence.

## Findings

**Location:** `RadAssist.ahk` lines 767-896

| Base Function | Prefilled Function | Duplication |
|---------------|-------------------|-------------|
| ShowEllipsoidVolumeGui() | ShowEllipsoidVolumeGuiPrefilled() | 90% |
| ShowRVLVGui() | ShowRVLVGuiPrefilled() | 90% |
| ShowNASCETGui() | ShowNASCETGuiPrefilled() | 90% |
| ShowAdrenalWashoutGui() | ShowAdrenalWashoutGuiPrefilled() | 90% |
| ShowFleischnerGui() | ShowFleischnerGuiPrefilled() | 90% |

**Agent:** pattern-recognition-specialist, code-simplicity-reviewer

## Proposed Solutions

### Option A: Add Optional Parameters (Recommended)
- **Pros:** Eliminates duplication, single maintenance point
- **Cons:** Slightly more complex function signature
- **Effort:** Medium
- **Risk:** Low

```autohotkey
ShowEllipsoidVolumeGui(prefill := "") {
    dim1 := prefill.d1 ? Round(prefill.d1, 2) : ""
    dim2 := prefill.d2 ? Round(prefill.d2, 2) : ""
    dim3 := prefill.d3 ? Round(prefill.d3, 2) : ""
    ; ... rest of GUI with dim1, dim2, dim3 as defaults
}
```

### Option B: Factory Function Pattern
- **Pros:** Maximum flexibility
- **Cons:** Over-engineered for this use case
- **Effort:** High
- **Risk:** Medium

## Recommended Action

(To be filled during triage)

## Technical Details

- **Affected Files:** RadAssist.ahk
- **Lines:** 247-268, 767-896 (and 4 more pairs)
- **Components:** All 5 calculator GUI functions
- **LOC Reduction:** ~130 lines

## Acceptance Criteria

- [ ] 5 prefilled GUI functions removed
- [ ] Base GUI functions accept optional prefill parameters
- [ ] Smart parse fallback to GUI still works
- [ ] All calculators function correctly with and without prefill
- [ ] No regression in GUI behavior

## Work Log

| Date | Action | Result/Learning |
|------|--------|-----------------|
| 2026-01-12 | Created from PR review | Initial finding |

## Resources

- PR #1: https://github.com/Skidudeaa/Radiology-Right-Click/pull/1
