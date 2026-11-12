# Radiology Right Click / RadAssist - notes for Claude

Radiologist (PowerScribe 360 now, PowerScribe One later; Sectra PACS) runs these
AutoHotkey tools on a locked-down hospital workstation from OneDrive / a local
folder. Test in Notepad first. Never put patient data in examples.

## The three shipped scripts (each a single self-contained file)
| File | AutoHotkey | Menu |
|---|---|---|
| `RadAssist.ahk` | v2 (portable AutoHotkey64.exe; drag script onto exe) | Shift+Right-click |
| `Radiology Right Click v1.22.ahk` | v1.1.24+ | Right-click |
| `Radiology Right Click v1.22-local.ahk` | v1.1.24+ (no web calls; Quick Call off by default) | Right-click |

**Rule: every feature/fix goes into all three versions.**

## Source modules and embedding
- `QuickCall.ahk` (v2) and `QuickCall_v1.ahk` (v1): floating call panel. Edit these,
  then re-embed between `; >>> QUICKCALL BEGIN` / `; <<< QUICKCALL END` markers in
  RadAssist.ahk (v2 body) and both v1.22 files (v1 body). Strip the standalone block.
- `SmartParse_v1.ahk`: v1 port of RadAssist's parse engine (`PE_*`) + UI (`SP_*`),
  embedded between `; >>> SMARTPARSE BEGIN` / `; <<< SMARTPARSE END` in both v1.22
  files (body from `global SP_Result` to end). Host calls `SP_Init()`,
  `SP_SetTarget(hwnd)`, `SP_AddMenu(...)`, `SP_SetDarkMenus`, `SP_DarkenGui`.
- RadAssist's engine lives inline in RadAssist.ahk (`PE_*` pure functions returning
  `{ok, title, values:[[k,v]...], text, conf, edit, atImpression}`; `SP_Run` shows the
  non-blocking confirm dialog; Insert re-activates the report window then pastes).
- v1 files must stay plain ASCII (AHK v1 reads no-BOM files as ANSI); CRLF endings.
- v1 compatibility: no `switch`, no built-in-only `Min/Max` reliance (module defines
  Min/Max), avoid features newer than v1.1.24. Code after `ExitApp` in v1 hosts has
  no auto-execute: initialize in functions.
- AHK variable names are case-insensitive (`M` and `m` collide) - bit us once.

## Behaviors settled with the user (do not regress)
- RV/LV wording: normal <0.9 / borderline 0.9-0.99 / elevated >=1.0 "compatible with
  RV dilatation, correlate..." - never "severe" or a strain/PE-severity diagnosis.
- Mild/moderate/severe size wording for the PROSTATE only (<30/30-49/50-79/80+ cc).
- Fleischner/TI-RADS/O-RADS blocks insert at end of report (= after IMPRESSION).
- Fleischner GUI defaults to both low- and high-risk recommendations.
- Quick Call: click names to call (NO 1-9 hotkeys - they are W/L presets); resizable;
  collapses to a 175 px rectangle via [+]/[-] or double-click on the header.
- Popups must come to the front over PowerScribe/Sectra and not block the menu.
- Menu must appear fast: quick 0.4 s selection copy before the menu, full copy
  (with retry) after an item is chosen if the quick copy missed.
- Dark mode: real dark menus (uxtheme 135/136), unthemed checkboxes/group titles.
- Bone Age tool off by default. Clinical tables to verify with the user: G&P SD table,
  SRU/O-RADS ovarian rules (written from memory).

## Testing (in the cloud container)
- Wine + Xvfb run real AutoHotkey: v2.0.18, v1.1.37, plus old v1.1.24/26/30 for
  compatibility (download zips from github releases / autohotkey.com/download/1.1/).
- Syntax check: v2 `AutoHotkey64.exe /ErrorStdOut /validate script`,
  v1 `AutoHotkeyU64.exe /ErrorStdOut /iLib NUL script`.
- Engine tests: run the `PE_*` functions on a phrase list and diff v1 vs v2 output
  (must be identical).
- End-to-end: a v2 "driver" script with `SendLevel(1)` opens Wine Notepad, types a
  phrase, selects it, sends (Shift+)RButton, picks the menu item, clicks Insert in
  the "Confirm" window, reads Notepad text.
- Wine quirk: focus changes take ~2 s; don't mistake that for real slowness.

## Repo / workflow
- Feature work on a `claude/...` branch -> PR into `main`; the user merges, then
  downloads Code > Download ZIP (extracts as `Radiology-Right-Click-main`).
- Keep user settings files: `RadAssist_preferences.ini`, `preferences.ini`,
  `QuickCall_PhoneBook.txt`, `QuickCall_Settings.ini`.
