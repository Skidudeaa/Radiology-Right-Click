# RadAssist v2.5

A lightweight radiology assistant tool with right-click calculators and smart text parsing. Available as both a Windows desktop app (AutoHotkey) and a modern web interface.

## Web Interface (New!)

Open `index.html` in any browser for a modern, responsive calculator suite featuring:

- **Quick Parse** - Auto-detect and parse measurement text instantly
- **Dark theme** with premium medical aesthetic
- **All 8 calculators** with real-time results and clinical interpretations
- **Copy to clipboard** - One-click result copying
- **Mobile responsive** - Works on tablets and phones

![RadAssist Web UI](https://via.placeholder.com/800x450?text=RadAssist+Web+UI)

### Calculators

| Calculator | Description |
|------------|-------------|
| Ellipsoid Volume | Organ/mass volume from 3D dimensions |
| RV/LV Ratio | Right heart strain evaluation for PE |
| NASCET Stenosis | Carotid stenosis with severity grading |
| Adrenal Washout | Absolute/relative washout for adenoma characterization |
| Fleischner 2017 | Lung nodule follow-up recommendations |
| ICH Volume | ABC/2 hemorrhage volume estimation |
| Vessel Stenosis | General diameter-based stenosis |
| Follow-up Date | Calculate imaging follow-up dates |

## Desktop App (Windows)

### What's New in v2.5

- **AutoHotkey v2 conversion** - Full rewrite to AHK v2 syntax for modern compatibility
- **Web interface** - New browser-based calculator suite

### What's New in v2.4

- **ICH Volume (ABC/2)**: Smart parse + GUI for intracerebral hemorrhage volume
- **Follow-up Date Calculator**: Smart parse + GUI for calculating follow-up dates
- **Tool Visibility**: Show/hide any of the 8 tools via Settings
- **Bug Fixes**: RV/LV parser, NASCET mixed units, Stenosis text deletion, Fleischner IMPRESSION detection

## Quick Start

### Web Interface
1. Open `index.html` in your browser
2. Use Quick Parse or select a specific calculator
3. Copy results to clipboard

### Desktop App (Windows)
1. **Install AutoHotkey v2** from [autohotkey.com](https://www.autohotkey.com/)
2. **Download** `RadAssist.ahk` to any folder
3. **Double-click** to run - look for the green "H" icon in your system tray
4. **Use it**: In PowerScribe or Notepad, highlight text with measurements, then **Shift + Right-click**

## How to Use (Desktop)

### Smart Parsing (Fastest Method)
1. Highlight text containing measurements (e.g., "prostate measures 4.2 x 3.1 x 3.5 cm")
2. **Shift + Right-click** → Choose "Quick Parse" or a specific smart parser
3. Review the parsed result → Click "Insert" to add it to your report

### Keyboard Shortcuts
| Key | Action |
|-----|--------|
| `` ` `` (backtick) | Pause/Resume script |
| `Ctrl+Shift+H` | Copy Sectra history |
| `Shift+Right-click` | Open RadAssist menu |

## Settings (Desktop)

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

### Web Interface
- Any modern browser (Chrome, Firefox, Safari, Edge)

### Desktop App
- Windows 10/11
- AutoHotkey v2.0+
- PowerScribe 360 or Notepad
