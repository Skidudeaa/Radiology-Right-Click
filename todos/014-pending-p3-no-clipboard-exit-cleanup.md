---
status: pending
priority: p3
issue_id: "014"
tags: [code-review, security, ahk]
dependencies: ["002"]
---

# No Clipboard Clearing on Script Exit

## Problem Statement

When the script terminates, the last-used clipboard content may remain, potentially containing medical data.

## Findings

No OnExit handler exists to clear clipboard.

**Agent:** security-sentinel

## Proposed Solution

Add exit handler:
```autohotkey
OnExit("CleanupOnExit")

CleanupOnExit() {
    Clipboard := ""
    return 0  ; Allow exit
}
```

**Effort:** Small | **Risk:** Low

## Acceptance Criteria

- [ ] OnExit handler registered
- [ ] Clipboard cleared on script exit
- [ ] Script still exits cleanly
