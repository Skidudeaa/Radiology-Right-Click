---
status: pending
priority: p3
issue_id: "011"
tags: [code-review, dry, refactoring, ahk]
dependencies: []
---

# "No Selection" Check Duplicated 5 Times

## Problem Statement

The same selection validation check is duplicated 5 times with minor message variations.

## Findings

**Location:** `RadAssist.ahk` lines 141-184

```autohotkey
MenuSmartVolume:
    if (g_SelectedText = "") {
        MsgBox, 48, No Selection, Please select text containing dimensions...
        return
    }

MenuSmartRVLV:
    if (g_SelectedText = "") {
        MsgBox, 48, No Selection, Please select text containing RV/LV measurements...
        return
    }
```

**Agent:** pattern-recognition-specialist

## Proposed Solution

Create dispatch function:
```autohotkey
MenuSmartParse(parseType) {
    if (g_SelectedText = "") {
        ShowNoSelectionError(parseType)
        return
    }
    ParseAndInsert%parseType%(g_SelectedText)
}
```

**Effort:** Small | **Risk:** Low | **LOC Reduction:** ~30 lines

## Acceptance Criteria

- [ ] Single dispatch function handles all smart parse menu items
- [ ] Error messages remain contextual
- [ ] Easier to add new parse types
