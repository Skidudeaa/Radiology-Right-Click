---
status: pending
priority: p3
issue_id: "015"
tags: [code-review, validation, ahk]
dependencies: []
---

# No Preferences Validation on Load

## Problem Statement

`LoadPreferences()` trusts INI values without validation. A corrupted INI could set invalid values, causing silent failures.

## Findings

**Location:** `RadAssist.ahk` lines 2110-2135

```autohotkey
IniRead, DefaultSmartParse, %PreferencesPath%, Settings, DefaultSmartParse, Volume
; No validation that DefaultSmartParse is a valid option
```

If `DefaultSmartParse` is set to an invalid value like "Foobar", the MenuQuickSmartParse handler would silently fail to match any parser.

**Agent:** architecture-strategist

## Proposed Solution

Add validation on load:
```autohotkey
validParsers := ["Volume", "RVLV", "NASCET", "Adrenal", "Fleischner"]
if (!HasValue(validParsers, DefaultSmartParse))
    DefaultSmartParse := "Volume"  ; Reset to default
```

**Effort:** Small | **Risk:** Low

## Acceptance Criteria

- [ ] All loaded preferences validated
- [ ] Invalid values reset to defaults
- [ ] Log or notify on reset (optional)
