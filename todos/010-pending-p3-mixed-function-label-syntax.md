---
status: pending
priority: p3
issue_id: "010"
tags: [code-review, consistency, ahk]
dependencies: []
---

# Mixed Function/Label Syntax

## Problem Statement

The code mixes modern function syntax with legacy AHK v1 label-based event handlers, creating inconsistency.

## Findings

```autohotkey
; Function syntax (modern)
ShowEllipsoidVolumeGui() {
    ...
}

; Label syntax (legacy)
EllipsoidGuiClose:
    Gui, EllipsoidGui:Destroy
return

CalcEllipsoid:
    ; calculation logic
return
```

**Agent:** pattern-recognition-specialist

## Proposed Solution

Consider AHK v2 migration or use Func.Bind() for consistency. For now, document the pattern and ensure all new code uses functions.

**Effort:** High (for full migration) | **Risk:** Medium

## Acceptance Criteria

- [ ] Pattern documented in code header
- [ ] New code uses function syntax
- [ ] Consider AHK v2 migration for major version
