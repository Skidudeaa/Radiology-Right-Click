---
status: pending
priority: p3
issue_id: "016"
tags: [code-review, architecture, ahk]
dependencies: []
---

# Forward Reference: InitPreferencesPath Called Before Defined

## Problem Statement

`InitPreferencesPath()` is called at line 31, but defined at line 2080. While AHK parses the entire file first (so this works), it creates maintenance risk.

## Findings

**Location:** `RadAssist.ahk` lines 31, 2080

```autohotkey
; Line 31
InitPreferencesPath()

; ... 2000 lines later ...

; Line 2080
InitPreferencesPath() {
    ; implementation
}
```

**Agent:** architecture-strategist

## Proposed Solution

Either:
1. Move function definition earlier (near top after globals)
2. Add comment at call site noting forward reference
3. Keep as-is (AHK handles it fine)

**Effort:** Small | **Risk:** Low

## Acceptance Criteria

- [ ] Decision made on placement
- [ ] If moved, function still works correctly
- [ ] Code organization documented
