# RadAssist v2.4

A lightweight AutoHotkey tool for radiologists using PowerScribe. Adds right-click calculators and smart text parsing for common radiology calculations.

## What's New in v2.4

### New Calculators
- **ICH Volume (ABC/2)**: Smart parse + GUI for intracerebral hemorrhage volume
- **Follow-up Date Calculator**: Smart parse + GUI for calculating follow-up dates

### Tool Visibility
- Show/hide any of the 8 tools via Settings
- Menu dynamically updates based on your preferences

### Bug Fixes
- RV/LV parser: Added Pattern H for "rvval 5.0 lvval 3.6" PACS format
- NASCET parser: Fixed mixed cm/mm unit handling with automatic conversion
- Stenosis Calculator: Fixed text deletion issue
- Fleischner: Better IMPRESSION detection (handles IMPRESSIONS, CLINICAL IMPRESSION)

### Output Formatting
- All calculators now use consistent sentence-style output with leading space

## What's New in v2.3

- Enhanced Fleischner parser: 5 pattern types, size ranges (8 x 6 mm), prefixes (up to, approximately)
- Enhanced RV/LV parser: 7 pattern types, full word support (Right ventricle), dilated/axial patterns
- Fixed InsertAtImpression timing issues
- Increased text selection timeout for slower applications

## Quick Start

1. **Install AutoHotkey v1.1** from [autohotkey.com](https://www.autohotkey.com/)
2. **Download** `RadAssist.ahk` to any folder
3. **Double-click** to run - look for the green "H" icon in your system tray
4. **Use it**: In PowerScribe or Notepad, highlight text with measurements, then **Shift + Right-click**

## How to Use

### Smart Parsing (Fastest Method)
1. Highlight text containing measurements (e.g., "prostate measures 4.2 x 3.1 x 3.5 cm")
2. **Shift + Right-click** → Choose "Quick Parse" or a specific smart parser
3. Review the parsed result → Click "Insert" to add it to your report

### Manual Calculators
1. **Shift + Right-click** anywhere in PowerScribe/Notepad
2. Choose a calculator:
   - **Ellipsoid Volume** - organs, masses, cysts
   - **RV/LV Ratio** - PE evaluation
   - **NASCET Stenosis** - carotid stenosis
   - **Adrenal Washout** - adenoma vs metastasis
   - **Fleischner 2017** - lung nodule follow-up
   - **General Stenosis** - any vessel stenosis
   - **ICH Volume** - hemorrhage ABC/2 calculation
   - **Date Calculator** - follow-up date planning

### Keyboard Shortcuts
| Key | Action |
|-----|--------|
| `` ` `` (backtick) | Pause/Resume script |
| `Ctrl+Shift+H` | Copy Sectra history |
| `Shift+Right-click` | Open RadAssist menu |

## Settings

Right-click the tray icon → **Settings** to customize:
- Default calculator for Quick Parse
- Show/hide citations
- RV/LV output format (Inline vs Macro)
- Default measurement units (cm/mm)
- Tool visibility (show/hide individual calculators)

## Troubleshooting

- **Script not responding?** Press backtick (`) - it might be paused
- **Menu not appearing?** Make sure you're in PowerScribe or Notepad
- **Settings not saving?** Check if the folder is read-only (OneDrive sync issue)

## Requirements

- Windows 10/11
- AutoHotkey v1.1.x
- PowerScribe 360 or Notepad
