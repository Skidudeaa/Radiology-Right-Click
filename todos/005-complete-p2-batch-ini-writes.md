---
status: pending
priority: p2
issue_id: "005"
tags: [code-review, performance, io, ahk]
dependencies: []
---

# Batch INI Writes Instead of 9 Separate Operations

## Problem Statement

When saving settings, the script performs 9 separate INI write operations with potential 3 retries each. This results in 9-27 file operations with up to 2.6 seconds of delays in worst case.

**Why it matters:** File I/O is slow, especially on OneDrive-synced folders. Batching writes would improve save performance.

## Findings

**Location:** `RadAssist.ahk` lines 2005-2013

```autohotkey
IniWriteWithRetry("IncludeDatamining", IncludeDatamining)
IniWriteWithRetry("ShowCitations", ShowCitations)
IniWriteWithRetry("DataminingPhrase", DataminingPhrase)
IniWriteWithRetry("DefaultSmartParse", DefaultSmartParse)
IniWriteWithRetry("SmartParseConfirmation", SmartParseConfirmation)
IniWriteWithRetry("SmartParseFallbackToGUI", SmartParseFallbackToGUI)
IniWriteWithRetry("DefaultMeasurementUnit", DefaultMeasurementUnit)
IniWriteWithRetry("RVLVOutputFormat", RVLVOutputFormat)
IniWriteWithRetry("FleischnerInsertAfterImpression", FleischnerInsertAfterImpression)
```

**Agent:** performance-oracle

## Proposed Solutions

### Option A: Write to Temp File, Then Rename (Recommended)
- **Pros:** Atomic operation, single write
- **Cons:** More complex implementation
- **Effort:** Medium
- **Risk:** Low

```autohotkey
SaveSettings() {
    tempFile := PreferencesPath . ".tmp"
    FileDelete, %tempFile%
    IniWrite, %IncludeDatamining%, %tempFile%, Settings, IncludeDatamining
    ; ... all other writes to tempFile
    FileMove, %tempFile%, %PreferencesPath%, 1  ; Overwrite
}
```

### Option B: Build Content String, Single FileAppend
- **Pros:** Very fast, single operation
- **Cons:** Loses INI format validation
- **Effort:** Small
- **Risk:** Medium

### Option C: Keep Separate Writes, Remove Retry on Success
- **Pros:** Minimal change
- **Cons:** Still 9 operations
- **Effort:** Small
- **Risk:** Low

Only retry subsequent writes if first one fails.

## Recommended Action

(To be filled during triage)

## Technical Details

- **Affected Files:** RadAssist.ahk
- **Lines:** 2005-2013, 2024-2034
- **Components:** SaveSettings label, IniWriteWithRetry()

## Acceptance Criteria

- [ ] Settings save completes in <500ms on normal systems
- [ ] Settings persist correctly across restarts
- [ ] OneDrive sync conflicts still handled gracefully
- [ ] No data loss on save failure

## Work Log

| Date | Action | Result/Learning |
|------|--------|-----------------|
| 2026-01-12 | Created from PR review | Initial finding |

## Resources

- PR #1: https://github.com/Skidudeaa/Radiology-Right-Click/pull/1
