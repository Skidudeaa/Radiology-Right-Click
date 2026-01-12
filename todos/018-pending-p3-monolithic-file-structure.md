---
status: pending
priority: p3
issue_id: "018"
tags: [code-review, architecture, maintainability, ahk]
dependencies: []
---

# Monolithic Single-File Design

## Problem Statement

At 2,160 lines, the script has grown beyond comfortable maintainability for a single file. AutoHotkey supports `#Include` directives for modularization.

## Findings

Current structure in single file:
- Event Handlers (lines 55-133)
- Calculator GUIs (lines 247-759)
- Pre-filled GUI variants (lines 763-896)
- Smart Parsers (lines 898-1823)
- Utility Functions (lines 1829-1918)
- Settings Management (lines 1920-2147)

**Agent:** architecture-strategist

## Proposed Solution

Split into modules:
```
RadAssist.ahk              ; Main entry, hotkeys, menu
Lib/Calculators.ahk        ; GUI calculator functions
Lib/SmartParsers.ahk       ; Text parsing functions
Lib/Preferences.ahk        ; Settings management
Lib/Utilities.ahk          ; Clipboard, insertion helpers
```

**Effort:** Medium-High | **Risk:** Medium

## Acceptance Criteria

- [ ] Logical separation of concerns
- [ ] All includes work correctly
- [ ] No regression in functionality
- [ ] Easier to maintain individual modules
