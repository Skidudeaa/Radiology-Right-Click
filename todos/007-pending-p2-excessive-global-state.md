---
status: pending
priority: p2
issue_id: "007"
tags: [code-review, architecture, refactoring, ahk]
dependencies: []
---

# Excessive Global State

## Problem Statement

The script has 14+ global variables scattered throughout, creating tight coupling between functions and making the code harder to test and maintain.

**Why it matters:** Global state makes it difficult to reason about function behavior, creates hidden dependencies, and prevents isolated testing.

## Findings

**Location:** `RadAssist.ahk` lines 27-49, 91

```autohotkey
; Configuration globals (lines 36-49)
global TargetApps
global SectraWindowTitle
global DataminingPhrase
global IncludeDatamining
global ShowCitations
global DefaultSmartParse
global SmartParseConfirmation
global SmartParseFallbackToGUI
global DefaultMeasurementUnit
global RVLVOutputFormat
global UseASCIICharacters
global FleischnerInsertAfterImpression

; Runtime state globals
global g_ConfirmAction      ; Inter-function communication
global g_ParsedResultText   ; Stores parsed result
global g_SelectedText       ; Selected text (line 91)
global PreferencesPath      ; INI file path
```

**Anti-pattern:** `g_ConfirmAction` and `g_ParsedResultText` are used for inter-function communication where parameters should be used instead.

**Agent:** pattern-recognition-specialist, architecture-strategist

## Proposed Solutions

### Option A: Configuration Object (Recommended)
- **Pros:** Organized, namespaced, easier to manage
- **Cons:** Requires refactoring all references
- **Effort:** Medium-High
- **Risk:** Medium

```autohotkey
global Config := {}
Config.Apps := {Targets: [...], Sectra: "Sectra"}
Config.SmartParse := {Default: "Volume", Confirmation: true}
Config.Output := {Citations: true, Datamining: true}
```

### Option B: Separate Runtime from Config
- **Pros:** Clearer separation of concerns
- **Cons:** Still uses globals
- **Effort:** Small
- **Risk:** Low

Group globals into Config (static) and State (runtime) sections with clear comments.

### Option C: Pass Parameters Instead of Globals
- **Pros:** Explicit dependencies, testable
- **Cons:** More verbose function signatures
- **Effort:** High
- **Risk:** Medium

## Recommended Action

(To be filled during triage)

## Technical Details

- **Affected Files:** RadAssist.ahk
- **Lines:** 27-49, 91, and all references
- **Components:** All functions that access global state

## Acceptance Criteria

- [ ] Globals organized into logical groups
- [ ] g_ConfirmAction replaced with function returns or callbacks
- [ ] Configuration accessible via single namespace
- [ ] No regression in functionality

## Work Log

| Date | Action | Result/Learning |
|------|--------|-----------------|
| 2026-01-12 | Created from PR review | Initial finding |

## Resources

- PR #1: https://github.com/Skidudeaa/Radiology-Right-Click/pull/1
