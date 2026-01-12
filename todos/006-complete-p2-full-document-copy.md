---
status: pending
priority: p2
issue_id: "006"
tags: [code-review, performance, security, ahk]
dependencies: []
---

# Full Document Copy to Find IMPRESSION

## Problem Statement

The `InsertAtImpression()` function copies the entire document to clipboard using `^a^c` (Select All, Copy) to search for "IMPRESSION:". This is slow for large documents and exposes full document content.

**Why it matters:** For 50+ page documents, this operation can take significant time and consumes substantial memory. It also exposes the entire document (potentially PHI) to clipboard-monitoring applications.

## Findings

**Location:** `RadAssist.ahk` lines 940-948

```autohotkey
Clipboard := ""
Send, ^a  ; Select all
Sleep, 50
Send, ^c  ; Copy entire document
ClipWait, 0.5
docText := Clipboard

if (!InStr(docText, "IMPRESSION")) {
    ; Fallback behavior
}
```

**Scalability:**
| Document Size | Estimated Time |
|---------------|----------------|
| 1KB | ~1 second |
| 10KB | ~1.2 seconds |
| 100KB | ~2+ seconds |

**Agent:** performance-oracle, security-sentinel

## Proposed Solutions

### Option A: Search from Cursor Position First (Recommended)
- **Pros:** Faster for most cases, less exposure
- **Cons:** May miss IMPRESSION above cursor
- **Effort:** Medium
- **Risk:** Low

Use Find dialog directly without pre-checking, let PowerScribe handle the search.

### Option B: Incremental Search with Ctrl+F
- **Pros:** No clipboard exposure
- **Cons:** Can't detect if IMPRESSION exists before attempting
- **Effort:** Small
- **Risk:** Medium (may fail silently)

Just try the Find operation; if it fails, fall back to cursor insertion.

### Option C: Check First Page Only
- **Pros:** Faster, less exposure
- **Cons:** May miss IMPRESSION on later pages
- **Effort:** Medium
- **Risk:** Medium

## Recommended Action

(To be filled during triage)

## Technical Details

- **Affected Files:** RadAssist.ahk
- **Lines:** 940-948
- **Components:** InsertAtImpression()

## Acceptance Criteria

- [ ] IMPRESSION search does not copy entire document
- [ ] Insertion still works when IMPRESSION exists
- [ ] Fallback still works when IMPRESSION not found
- [ ] Performance improved for large documents

## Work Log

| Date | Action | Result/Learning |
|------|--------|-----------------|
| 2026-01-12 | Created from PR review | Initial finding |

## Resources

- PR #1: https://github.com/Skidudeaa/Radiology-Right-Click/pull/1
