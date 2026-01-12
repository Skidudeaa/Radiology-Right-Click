---
status: pending
priority: p3
issue_id: "012"
tags: [code-review, error-handling, ahk]
dependencies: []
---

# Silent Failures in IniWriteWithRetry

## Problem Statement

The `IniWriteWithRetry()` function returns false on failure, but callers never check this return value. Users are not notified when settings fail to save.

## Findings

**Location:** `RadAssist.ahk` lines 2024-2034

```autohotkey
IniWriteWithRetry(key, value, maxRetries := 3) {
    Loop, %maxRetries%
    {
        IniWrite, %value%, %PreferencesPath%, Settings, %key%
        if (!ErrorLevel)
            return true
        Sleep, 100
    }
    return false  ; This is never checked!
}
```

**Callers (lines 2005-2013):**
```autohotkey
IniWriteWithRetry("IncludeDatamining", IncludeDatamining)  ; Return ignored
IniWriteWithRetry("ShowCitations", ShowCitations)          ; Return ignored
; ... etc
```

**Agent:** security-sentinel, pattern-recognition-specialist

## Proposed Solution

Check return values and notify user:
```autohotkey
SaveSettings:
    failed := false
    failed := failed || !IniWriteWithRetry("IncludeDatamining", IncludeDatamining)
    ; ... etc
    if (failed)
        MsgBox, 48, Warning, Some settings may not have saved. Check file permissions.
```

**Effort:** Small | **Risk:** Low

## Acceptance Criteria

- [ ] Return values checked in SaveSettings
- [ ] User notified on save failure
- [ ] No silent data loss
