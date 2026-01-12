---
status: pending
priority: p3
issue_id: "013"
tags: [code-review, dry, refactoring, ahk]
dependencies: []
---

# NASCET/Stenosis Calculators 90% Identical

## Problem Statement

`ShowStenosisGui()` and `ShowNASCETGui()` are 90% identical, with the same calculation formula and severity thresholds. Only difference is vessel name field and labels.

## Findings

**Location:** `RadAssist.ahk` lines 506-579 (Stenosis), lines 398-505 (NASCET)

Both use: `stenosis := Round((1 - (residual / normal)) * 100, 1)`

Both have same thresholds:
- <50%: Mild
- 50-69%: Moderate
- >=70%: Severe

**Agent:** code-simplicity-reviewer

## Proposed Solution

Combine into single parameterized function:
```autohotkey
ShowStenosisGui(vesselPreset := "") {
    ; If vesselPreset = "NASCET", set labels for ICA
    ; Otherwise, generic stenosis calculator
}
```

**Effort:** Medium | **Risk:** Low | **LOC Reduction:** ~70 lines

## Acceptance Criteria

- [ ] Single stenosis calculator function
- [ ] NASCET preset adds ICA-specific labels
- [ ] Both menu items still work
- [ ] No regression in calculations
