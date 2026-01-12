---
status: pending
priority: p2
issue_id: "002"
tags: [code-review, security, hipaa, ahk]
dependencies: []
---

# Clipboard PHI Exposure Window

## Problem Statement

Clipboard operations expose Protected Health Information (PHI) for approximately 600ms+ during text insertion and document search operations. Clipboard-monitoring applications (clipboard managers, keyloggers, malware) could capture sensitive patient data.

**Why it matters:** In healthcare environments with strict HIPAA compliance requirements, even brief exposure windows warrant attention.

## Findings

**Locations:**
- `RadAssist.ahk` lines 913-923 (InsertAfterSelection)
- `RadAssist.ahk` lines 940-948 (InsertAtImpression - copies entire document)

```autohotkey
; InsertAfterSelection - ~600ms exposure
ClipSaved := ClipboardAll
Clipboard := textToInsert
ClipWait, 0.5
Send, ^v
Sleep, 100
Clipboard := ClipSaved

; InsertAtImpression - exposes full document
Send, ^a  ; Select all
Send, ^c  ; Copy entire document to clipboard
docText := Clipboard
```

**Agent:** security-sentinel

## Proposed Solutions

### Option A: Minimize Exposure Window (Recommended)
- **Pros:** Simple, reduces exposure time
- **Cons:** Still has brief exposure
- **Effort:** Small
- **Risk:** Low

Reduce Sleep times to minimum needed, clear clipboard immediately after use.

### Option B: Alternative IMPRESSION Search
- **Pros:** Eliminates full document exposure
- **Cons:** Requires different search approach
- **Effort:** Medium
- **Risk:** Medium

Use Windows UI Automation or incremental search instead of ^a^c.

### Option C: Add Exit Cleanup
- **Pros:** Clears clipboard on script exit
- **Cons:** Doesn't address runtime exposure
- **Effort:** Small
- **Risk:** Low

```autohotkey
OnExit("CleanupOnExit")
CleanupOnExit() {
    Clipboard := ""
}
```

## Recommended Action

(To be filled during triage)

## Technical Details

- **Affected Files:** RadAssist.ahk
- **Lines:** 913-923, 940-948
- **Components:** InsertAfterSelection(), InsertAtImpression()

## Acceptance Criteria

- [ ] Clipboard exposure window reduced to minimum necessary
- [ ] Clipboard cleared after InsertAtImpression document search
- [ ] OnExit handler clears clipboard on script termination
- [ ] No regression in insertion functionality

## Work Log

| Date | Action | Result/Learning |
|------|--------|-----------------|
| 2026-01-12 | Created from PR review | Initial finding |

## Resources

- PR #1: https://github.com/Skidudeaa/Radiology-Right-Click/pull/1
- HIPAA Security Rule - Technical Safeguards
