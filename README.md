# RadAssist v2.2

A lightweight AutoHotkey tool for radiologists using PowerScribe. Adds right-click calculators and smart text parsing for common radiology calculations.

## What's New in v2.2

- **Pause/Resume**: Press backtick (`) to pause the script
- **RV/LV Macro Format**: Output matches PowerScribe "Macro right heart" template
- **Fleischner IMPRESSION Insert**: Lung nodule recommendations insert after IMPRESSION: field
- **OneDrive Compatible**: Works from synced folders without file conflicts
- **Better Encoding**: Uses ASCII characters (<=, >=) for PowerScribe compatibility
- **Default cm Units**: Measurements default to cm with auto-detection

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

## Troubleshooting

- **Script not responding?** Press backtick (`) - it might be paused
- **Menu not appearing?** Make sure you're in PowerScribe or Notepad
- **Settings not saving?** Check if the folder is read-only (OneDrive sync issue)

## Requirements

- Windows 10/11
- AutoHotkey v1.1.x
- PowerScribe 360 or Notepad
