; ==========================================
; RadAssist - Radiology Assistant Tool
; Version: 2.5 (AHK v2)
; Description: Lean radiology workflow tool with calculators
;              Triggered by Shift+Right-click in PowerScribe/Notepad
;              Smart text parsing with confirmation dialogs
; ARCHITECTURE: Context-filtered parsing with confidence scoring
; WHY: v2.5 converts to AHK v2 syntax for compatibility
; CHANGES in v2.5:
;   - Converted to AutoHotkey v2 syntax
;   - GUI system rewritten to use v2 Gui class
;   - Menu system converted to v2 Menu class
;   - All commands converted to v2 function syntax
; ==========================================

#Requires AutoHotkey v2.0
#SingleInstance Force
SetWorkingDir(A_ScriptDir)

; -----------------------------------------
; OneDrive Compatibility - Determine preferences path
; WHY: Script may run from OneDrive folder with sync conflicts
; TRADEOFF: Falls back to LOCALAPPDATA if script dir is read-only
; -----------------------------------------
global PreferencesPath := ""

; Initialize preferences path (function defined in Functions section below)
; Call after global declaration to set path based on write permissions
; NOTE: Forward reference - function defined at end of file (AHK parses all functions first)
InitPreferencesPath()

; -----------------------------------------
; Global Configuration
; -----------------------------------------
global TargetApps := ["ahk_exe Nuance.PowerScribe360.exe", "ahk_exe notepad.exe", "ahk_class Notepad"]
global SectraWindowTitle := "Sectra"
global DataminingPhrase := "SSM lung nodule"
global IncludeDatamining := true
global ShowCitations := true
global DefaultSmartParse := "Volume"  ; Options: Volume, RVLV, NASCET, Adrenal, Fleischner
global SmartParseConfirmation := true  ; Always show confirmation dialog before insert
global SmartParseFallbackToGUI := true ; Fall back to GUI when parsing confidence is low
global DefaultMeasurementUnit := "cm"  ; Options: cm, mm - used when no units in input
global RVLVOutputFormat := "Macro"     ; Options: Inline, Macro - output format for RV/LV
global UseASCIICharacters := true      ; Use ASCII <=/>= instead of Unicode
global FleischnerInsertAfterImpression := true  ; Find IMPRESSION: and insert below
global g_ConfirmAction := ""  ; Used by confirmation dialog
global g_ParsedResultText := ""  ; Stores parsed result for insertion

; -----------------------------------------
; Tool Visibility Preferences
; WHY: Allow users to show/hide tools they don't use
; NOTE: These are saved to INI and loaded on startup
; -----------------------------------------
global ShowVolumeTools := true
global ShowRVLVTools := true
global ShowNASCETTools := true
global ShowAdrenalTools := true
global ShowFleischnerTools := true
global ShowStenosisTools := true
global ShowICHTools := true
global ShowDateCalculator := true

; -----------------------------------------
; Quick Call (floating always-on-top call panel, see QuickCall.ahk)
; WHY: Optional module - toggled from Settings, saved to INI
; -----------------------------------------
global QuickCallEnabled := true

; -----------------------------------------
; Exit Cleanup Handler
; WHY: Clear clipboard on exit to prevent PHI exposure
; -----------------------------------------
OnExit(CleanupOnExit)

CleanupOnExit(ExitReason, ExitCode) {
    A_Clipboard := ""
    return 0  ; Allow exit to proceed
}

; -----------------------------------------
; Global Hotkey: Ctrl+Shift+H - Sectra History Copy
; (Can be mapped to Contour ShuttlePRO button)
; -----------------------------------------
^+h:: {
    CopySectraHistory()
}

; -----------------------------------------
; Global Hotkey: Backtick (`) - Pause/Resume Script
; WHY: Quick toggle without opening menu
; -----------------------------------------
`:: {
    Suspend(-1)
    if (A_IsSuspended)
        TrayTip("Script PAUSED - Press `` to resume", "RadAssist", 1)
    else
        TrayTip("Script RESUMED", "RadAssist", 1)
}

; -----------------------------------------
; Shift+Right-Click Menu Handler
; -----------------------------------------
+RButton:: {
    global g_SelectedText, TargetApps, ShowVolumeTools, ShowRVLVTools, ShowNASCETTools
    global ShowAdrenalTools, ShowFleischnerTools, ShowStenosisTools, ShowICHTools
    global ShowDateCalculator, DefaultSmartParse

    ; Check if we're in a target application
    inTargetApp := false
    for app in TargetApps {
        if (WinActive(app)) {
            inTargetApp := true
            break
        }
    }

    if (!inTargetApp) {
        ; Not in target app - send normal shift+right-click
        Send("+{RButton}")
        return
    }

    ; Store any selected text
    ; WHY: Extended timeout (0.5s → 1s) for slower applications like PowerScribe
    ; TRADEOFF: Slightly slower menu appearance vs better reliability
    g_SelectedText := ""
    ClipSaved := ClipboardAll()
    A_Clipboard := ""
    Send("^c")
    if ClipWait(1) {
        g_SelectedText := A_Clipboard
    }
    A_Clipboard := ClipSaved

    ; Build and show menu
    ; WHY: Conditionally build menu based on tool visibility preferences
    ; ARCHITECTURE: Menu items only added if corresponding Show*Tools is true
    RadAssistMenu := Menu()
    SmartParseMenu := Menu()

    smartParseHasItems := false
    if (ShowVolumeTools) {
        SmartParseMenu.Add("Smart Volume (parse dimensions)", MenuSmartVolume)
        smartParseHasItems := true
    }
    if (ShowRVLVTools) {
        SmartParseMenu.Add("Smart RV/LV (parse ratio)", MenuSmartRVLV)
        smartParseHasItems := true
    }
    if (ShowNASCETTools) {
        SmartParseMenu.Add("Smart NASCET (parse stenosis)", MenuSmartNASCET)
        smartParseHasItems := true
    }
    if (ShowAdrenalTools) {
        SmartParseMenu.Add("Smart Adrenal (parse HU values)", MenuSmartAdrenal)
        smartParseHasItems := true
    }
    if (ShowFleischnerTools) {
        if (smartParseHasItems)
            SmartParseMenu.Add()
        SmartParseMenu.Add("Parse Nodules (Fleischner)", MenuSmartFleischner)
        smartParseHasItems := true
    }

    if (smartParseHasItems) {
        RadAssistMenu.Add("Smart Parse", SmartParseMenu)
        RadAssistMenu.Add("Quick Parse (" DefaultSmartParse ")", MenuQuickSmartParse)
        RadAssistMenu.Add()
    }

    ; GUI calculators - conditionally add based on visibility
    guiHasItems := false
    if (ShowVolumeTools) {
        RadAssistMenu.Add("Ellipsoid Volume (GUI)", MenuEllipsoidVolume)
        guiHasItems := true
    }
    if (ShowAdrenalTools) {
        RadAssistMenu.Add("Adrenal Washout (GUI)", MenuAdrenalWashout)
        guiHasItems := true
    }
    if (guiHasItems)
        RadAssistMenu.Add()

    stenosisHasItems := false
    if (ShowNASCETTools) {
        RadAssistMenu.Add("NASCET (Carotid)", MenuNASCET)
        stenosisHasItems := true
    }
    if (ShowStenosisTools) {
        RadAssistMenu.Add("Vessel Stenosis (General)", MenuStenosis)
        stenosisHasItems := true
    }
    if (ShowRVLVTools) {
        RadAssistMenu.Add("RV/LV Ratio (GUI)", MenuRVLV)
        stenosisHasItems := true
    }
    if (stenosisHasItems)
        RadAssistMenu.Add()

    if (ShowFleischnerTools) {
        RadAssistMenu.Add("Fleischner 2017 (GUI)", MenuFleischner)
        RadAssistMenu.Add()
    }

    ; New calculators section
    newCalcHasItems := false
    if (ShowICHTools) {
        RadAssistMenu.Add("ICH Volume (ABC/2)", MenuICHVolume)
        newCalcHasItems := true
    }
    if (ShowDateCalculator) {
        RadAssistMenu.Add("Follow-up Date Calculator", MenuDateCalc)
        newCalcHasItems := true
    }
    if (newCalcHasItems)
        RadAssistMenu.Add()

    RadAssistMenu.Add("Copy Sectra History (Ctrl+Shift+H)", MenuSectraHistory)
    if (QuickCallEnabled)
        RadAssistMenu.Add("Quick Call Panel (Ctrl+Alt+P)", MenuQuickCall)
    RadAssistMenu.Add()
    RadAssistMenu.Add("Settings", MenuSettings)

    RadAssistMenu.Show()
}

; -----------------------------------------
; Utility: Show no-selection error with context
; -----------------------------------------
ShowNoSelectionError(parseType) {
    messages := Map("Volume", "dimensions", "RVLV", "RV/LV measurements", "NASCET", "stenosis values", "Adrenal", "HU values", "Fleischner", "nodule measurements")
    msg := messages.Has(parseType) ? messages[parseType] : "text"
    MsgBox("Please select text containing " msg " first.`n`nHighlight the relevant text and try again.", "No Selection", 48)
}

; -----------------------------------------
; Menu Handlers (converted to functions for v2)
; -----------------------------------------

; Smart Parse handlers (inline text parsing with insertion)
MenuSmartVolume(ItemName, ItemPos, MyMenu) {
    global g_SelectedText
    if (g_SelectedText = "") {
        ShowNoSelectionError("Volume")
        return
    }
    ParseAndInsertVolume(g_SelectedText)
}

MenuSmartRVLV(ItemName, ItemPos, MyMenu) {
    global g_SelectedText
    if (g_SelectedText = "") {
        MsgBox('Please select text containing RV/LV measurements (e.g., "RV 42mm / LV 35mm")', "No Selection", 48)
        return
    }
    ParseAndInsertRVLV(g_SelectedText)
}

MenuSmartNASCET(ItemName, ItemPos, MyMenu) {
    global g_SelectedText
    if (g_SelectedText = "") {
        MsgBox('Please select text containing stenosis measurements (e.g., "distal 5.2mm, stenosis 2.1mm")', "No Selection", 48)
        return
    }
    ParseAndInsertNASCET(g_SelectedText)
}

MenuSmartAdrenal(ItemName, ItemPos, MyMenu) {
    global g_SelectedText
    if (g_SelectedText = "") {
        MsgBox('Please select text containing HU values (e.g., "pre-contrast 10 HU, enhanced 80 HU, delayed 40 HU")', "No Selection", 48)
        return
    }
    ParseAndInsertAdrenalWashout(g_SelectedText)
}

MenuSmartFleischner(ItemName, ItemPos, MyMenu) {
    global g_SelectedText
    if (g_SelectedText = "") {
        MsgBox("Please select the findings section containing nodule descriptions.", "No Selection", 48)
        return
    }
    ParseAndInsertFleischner(g_SelectedText)
}

MenuQuickSmartParse(ItemName, ItemPos, MyMenu) {
    global g_SelectedText, DefaultSmartParse
    if (g_SelectedText = "") {
        MsgBox("Please select text to parse.", "No Selection", 48)
        return
    }
    if (DefaultSmartParse = "Volume")
        ParseAndInsertVolume(g_SelectedText)
    else if (DefaultSmartParse = "RVLV")
        ParseAndInsertRVLV(g_SelectedText)
    else if (DefaultSmartParse = "NASCET")
        ParseAndInsertNASCET(g_SelectedText)
    else if (DefaultSmartParse = "Adrenal")
        ParseAndInsertAdrenalWashout(g_SelectedText)
    else if (DefaultSmartParse = "Fleischner")
        ParseAndInsertFleischner(g_SelectedText)
    else
        ParseAndInsertVolume(g_SelectedText)
}

; GUI-based handlers
MenuEllipsoidVolume(ItemName, ItemPos, MyMenu) {
    ShowEllipsoidVolumeGui()
}

MenuAdrenalWashout(ItemName, ItemPos, MyMenu) {
    ShowAdrenalWashoutGui()
}

MenuNASCET(ItemName, ItemPos, MyMenu) {
    global g_SelectedText
    ; Try text parsing first, fall back to GUI
    if (g_SelectedText != "") {
        result := ParseNASCET(g_SelectedText)
        if (result != "") {
            ShowResult(result)
            return
        }
    }
    ShowNASCETGui()
}

MenuStenosis(ItemName, ItemPos, MyMenu) {
    ShowStenosisGui()
}

MenuRVLV(ItemName, ItemPos, MyMenu) {
    ShowRVLVGui()
}

MenuFleischner(ItemName, ItemPos, MyMenu) {
    ShowFleischnerGui()
}

MenuICHVolume(ItemName, ItemPos, MyMenu) {
    global g_SelectedText
    if (g_SelectedText != "") {
        ParseAndInsertICH(g_SelectedText)
    } else {
        ShowICHVolumeGui()
    }
}

MenuDateCalc(ItemName, ItemPos, MyMenu) {
    global g_SelectedText
    if (g_SelectedText != "") {
        ParseAndInsertDate(g_SelectedText)
    } else {
        ShowDateCalculatorGui()
    }
}

MenuSectraHistory(ItemName, ItemPos, MyMenu) {
    CopySectraHistory()
}

MenuQuickCall(ItemName, ItemPos, MyMenu) {
    QC_ShowPanel()
}

MenuSettings(ItemName, ItemPos, MyMenu) {
    ShowSettings()
}

; =========================================
; CALCULATOR 1: Ellipsoid Volume
; =========================================
; Global GUI object references for v2
global EllipsoidGuiObj := ""

ShowEllipsoidVolumeGui() {
    global EllipsoidGuiObj
    ; Position near mouse
    GetGuiPosition(&xPos, &yPos)

    EllipsoidGuiObj := Gui("+AlwaysOnTop")
    EllipsoidGuiObj.Title := "Ellipsoid Volume"
    EllipsoidGuiObj.Add("Text", "x10 y10 w280", "Ellipsoid Volume Calculator")
    EllipsoidGuiObj.Add("Text", "x10 y35", "AP (L):")
    EllipsoidGuiObj.Add("Edit", "x80 y32 w50 vEllipDim1")
    EllipsoidGuiObj.Add("Text", "x135 y35", "x   T (W):")
    EllipsoidGuiObj.Add("Edit", "x185 y32 w50 vEllipDim2")
    EllipsoidGuiObj.Add("Text", "x10 y60", "CC (H):")
    EllipsoidGuiObj.Add("Edit", "x80 y57 w50 vEllipDim3")
    EllipsoidGuiObj.Add("Text", "x135 y60", "Units:")
    EllipsoidGuiObj.Add("DropDownList", "x185 y57 w50 vEllipUnits Choose1", ["mm", "cm"])
    EllipsoidGuiObj.Add("Button", "x10 y95 w100", "Calculate").OnEvent("Click", CalcEllipsoid)
    EllipsoidGuiObj.Add("Button", "x120 y95 w80", "Cancel").OnEvent("Click", EllipsoidGuiClose)
    EllipsoidGuiObj.OnEvent("Close", EllipsoidGuiClose)
    EllipsoidGuiObj.Show("x" xPos " y" yPos " w250 h135")
}

EllipsoidGuiClose(*) {
    global EllipsoidGuiObj
    if (EllipsoidGuiObj)
        EllipsoidGuiObj.Destroy()
    EllipsoidGuiObj := ""
}

CalcEllipsoid(*) {
    global EllipsoidGuiObj
    saved := EllipsoidGuiObj.Submit(false)

    if (saved.EllipDim1 = "" || saved.EllipDim2 = "" || saved.EllipDim3 = "") {
        MsgBox("Please enter all three dimensions.", "Error", 16)
        return
    }

    d1 := saved.EllipDim1 + 0
    d2 := saved.EllipDim2 + 0
    d3 := saved.EllipDim3 + 0

    if (d1 <= 0 || d2 <= 0 || d3 <= 0) {
        MsgBox("All dimensions must be greater than 0.", "Error", 16)
        return
    }

    ; Convert to cm if input is mm
    if (saved.EllipUnits = "mm") {
        d1 := d1 / 10
        d2 := d2 / 10
        d3 := d3 / 10
    }

    ; Ellipsoid volume formula: (4/3) * pi * (a/2) * (b/2) * (c/2) = 0.5236 * a * b * c
    volume := 0.5236 * d1 * d2 * d3
    volumeRound := Round(volume, 1)

    ; WHY: Sentence style for inline insertion - matches smart parser output
    ; ARCHITECTURE: Leading space for dictation continuity
    ; Format dimensions in display units
    dimStr := saved.EllipDim1 " x " saved.EllipDim2 " x " saved.EllipDim3 " " saved.EllipUnits
    result := " This corresponds to a volume of " volumeRound " cc (" dimStr ")."

    EllipsoidGuiObj.Destroy()
    EllipsoidGuiObj := ""
    ShowResult(result)
}

; =========================================
; CALCULATOR 2: Adrenal Washout
; =========================================
global AdrenalGuiObj := ""

ShowAdrenalWashoutGui() {
    global AdrenalGuiObj
    CoordMode("Mouse", "Screen")
    MouseGetPos(&mouseX, &mouseY)
    xPos := mouseX + 10
    yPos := mouseY + 10

    AdrenalGuiObj := Gui("+AlwaysOnTop")
    AdrenalGuiObj.Title := "Adrenal Washout"
    AdrenalGuiObj.Add("Text", "x10 y10 w280", "Adrenal Washout Calculator")
    AdrenalGuiObj.Add("Text", "x10 y40", "Pre-contrast (HU):")
    AdrenalGuiObj.Add("Edit", "x140 y37 w70 vAdrenalPre")
    AdrenalGuiObj.Add("Text", "x10 y70", "Post-contrast (HU):")
    AdrenalGuiObj.Add("Edit", "x140 y67 w70 vAdrenalPost")
    AdrenalGuiObj.Add("Text", "x10 y100", "Delayed (15 min) (HU):")
    AdrenalGuiObj.Add("Edit", "x140 y97 w70 vAdrenalDelayed")
    AdrenalGuiObj.Add("Button", "x10 y135 w90", "Calculate").OnEvent("Click", CalcAdrenal)
    AdrenalGuiObj.Add("Button", "x110 y135 w80", "Cancel").OnEvent("Click", AdrenalGuiClose)
    AdrenalGuiObj.OnEvent("Close", AdrenalGuiClose)
    AdrenalGuiObj.Show("x" xPos " y" yPos " w230 h175")
}

AdrenalGuiClose(*) {
    global AdrenalGuiObj
    if (AdrenalGuiObj)
        AdrenalGuiObj.Destroy()
    AdrenalGuiObj := ""
}

CalcAdrenal(*) {
    global AdrenalGuiObj, ShowCitations
    saved := AdrenalGuiObj.Submit(false)

    if (saved.AdrenalPre = "" || saved.AdrenalPost = "" || saved.AdrenalDelayed = "") {
        MsgBox("Please enter all three HU values.", "Error", 16)
        return
    }

    pre := saved.AdrenalPre + 0
    post := saved.AdrenalPost + 0
    delayed := saved.AdrenalDelayed + 0

    ; WHY: Sentence style for inline insertion - matches smart parser output
    ; ARCHITECTURE: Leading space for dictation continuity
    result := " Adrenal washout: "

    ; Absolute washout (requires pre-contrast)
    if (post != pre) {
        absWashout := ((post - delayed) / (post - pre)) * 100
        absWashout := Round(absWashout, 1)
        result .= "absolute " absWashout "%"
        if (absWashout >= 60)
            result .= " (likely adenoma)"
        else
            result .= " (indeterminate)"
        result .= ", "
    }

    ; Relative washout (no pre-contrast needed)
    if (post != 0) {
        relWashout := ((post - delayed) / post) * 100
        relWashout := Round(relWashout, 1)
        result .= "relative " relWashout "%"
        if (relWashout >= 40)
            result .= " (likely adenoma)"
        else
            result .= " (indeterminate)"
        result .= "."
    }

    ; Pre-contrast assessment as additional sentence
    if (pre <= 10)
        result .= " Pre-contrast " pre " HU suggests lipid-rich adenoma."

    if (ShowCitations)
        result .= " (Mayo-Smith et al. Radiology 2017)"

    AdrenalGuiObj.Destroy()
    AdrenalGuiObj := ""
    ShowResult(result)
}

; =========================================
; CALCULATOR 3: NASCET Stenosis
; =========================================
global NASCETGuiObj := ""

ParseNASCET(input) {
    ; Try to parse distal and stenosis from text
    input := RegExReplace(input, "`r?\n", " ")

    ; Pattern 1: "distal X mm ... stenosis Y mm"
    if (RegExMatch(input, "i)distal.*?(\d+(?:\.\d+)?)\s*(?:mm|cm).*?stenosis.*?(\d+(?:\.\d+)?)\s*(?:mm|cm)", &m)) {
        return CalculateNASCETResult(m[1], m[2])
    }
    ; Pattern 2: "stenosis X mm ... distal Y mm"
    if (RegExMatch(input, "i)stenosis.*?(\d+(?:\.\d+)?)\s*(?:mm|cm).*?distal.*?(\d+(?:\.\d+)?)\s*(?:mm|cm)", &m)) {
        return CalculateNASCETResult(m[2], m[1])
    }
    ; Pattern 3: Just two numbers - assume larger is distal
    numbers := []
    pos := 1
    while (pos := RegExMatch(input, "(\d+(?:\.\d+)?)\s*(?:mm|cm)?", &m, pos)) {
        numbers.Push(m[1] + 0)
        pos += StrLen(m[0])
    }
    if (numbers.Length >= 2) {
        distal := Max(numbers*)
        stenosis := Min(numbers*)
        return CalculateNASCETResult(distal, stenosis)
    }
    return ""
}

CalculateNASCETResult(distal, stenosis) {
    global ShowCitations
    distal := distal + 0
    stenosis := stenosis + 0

    if (distal <= 0)
        return ""

    nascetVal := ((distal - stenosis) / distal) * 100
    nascetVal := Round(nascetVal, 1)

    ; WHY: Sentence style for inline insertion - matches smart parser output
    ; ARCHITECTURE: Leading space for dictation continuity
    distalRound := Round(distal, 1)
    stenosisRound := Round(stenosis, 1)

    result := " NASCET: " nascetVal "% stenosis (distal " distalRound "mm, stenosis " stenosisRound "mm), "

    if (nascetVal < 50)
        result .= "mild stenosis."
    else if (nascetVal < 70)
        result .= "moderate stenosis, consider intervention if symptomatic."
    else
        result .= "severe stenosis, strong indication for CEA/CAS."

    if (ShowCitations)
        result .= " (NASCET, NEJM 1991)"

    return result
}

ShowNASCETGui() {
    global NASCETGuiObj
    CoordMode("Mouse", "Screen")
    MouseGetPos(&mouseX, &mouseY)
    xPos := mouseX + 10
    yPos := mouseY + 10

    NASCETGuiObj := Gui("+AlwaysOnTop")
    NASCETGuiObj.Title := "NASCET Calculator"
    NASCETGuiObj.Add("Text", "x10 y10 w280", "NASCET Carotid Stenosis Calculator")
    NASCETGuiObj.Add("Text", "x10 y40", "Distal ICA diameter (mm):")
    NASCETGuiObj.Add("Edit", "x160 y37 w60 vNASCETDistal")
    NASCETGuiObj.Add("Text", "x10 y70", "Stenosis diameter (mm):")
    NASCETGuiObj.Add("Edit", "x160 y67 w60 vNASCETStenosis")
    NASCETGuiObj.Add("Button", "x10 y105 w90", "Calculate").OnEvent("Click", CalcNASCETGui)
    NASCETGuiObj.Add("Button", "x110 y105 w80", "Cancel").OnEvent("Click", NASCETGuiClose)
    NASCETGuiObj.OnEvent("Close", NASCETGuiClose)
    NASCETGuiObj.Show("x" xPos " y" yPos " w240 h145")
}

NASCETGuiClose(*) {
    global NASCETGuiObj
    if (NASCETGuiObj)
        NASCETGuiObj.Destroy()
    NASCETGuiObj := ""
}

CalcNASCETGui(*) {
    global NASCETGuiObj
    saved := NASCETGuiObj.Submit(false)

    if (saved.NASCETDistal = "" || saved.NASCETStenosis = "") {
        MsgBox("Please enter both measurements.", "Error", 16)
        return
    }

    distal := saved.NASCETDistal + 0
    stenosis := saved.NASCETStenosis + 0

    if (distal <= 0) {
        MsgBox("Distal ICA must be greater than 0.", "Error", 16)
        return
    }

    if (stenosis >= distal) {
        MsgBox("Stenosis should be less than distal ICA.", "Error", 16)
        return
    }

    result := CalculateNASCETResult(distal, stenosis)
    NASCETGuiObj.Destroy()
    NASCETGuiObj := ""
    ShowResult(result)
}

; -----------------------------------------
; CALCULATOR 3b: General Stenosis Calculator
; NOTE: 90% similar to NASCET calculator above. Consider merging
; into single parameterized function: ShowStenosisGui(preset := "")
; where preset="NASCET" adds ICA-specific labels.
; -----------------------------------------
global StenosisGuiObj := ""

ShowStenosisGui() {
    global StenosisGuiObj
    CoordMode("Mouse", "Screen")
    MouseGetPos(&mouseX, &mouseY)
    xPos := mouseX + 10
    yPos := mouseY + 10

    StenosisGuiObj := Gui("+AlwaysOnTop")
    StenosisGuiObj.Title := "Stenosis Calculator"
    StenosisGuiObj.Add("Text", "x10 y10 w280", "Vessel Stenosis Calculator")
    StenosisGuiObj.Add("Text", "x10 y35", "Vessel (optional):")
    StenosisGuiObj.Add("Edit", "x120 y32 w130 vStenosisVessel")
    StenosisGuiObj.Add("Text", "x10 y65", "Normal diameter (mm):")
    StenosisGuiObj.Add("Edit", "x140 y62 w60 vStenosisNormal")
    StenosisGuiObj.Add("Text", "x10 y95", "Stenosis diameter (mm):")
    StenosisGuiObj.Add("Edit", "x140 y92 w60 vStenosisMeasure")
    StenosisGuiObj.Add("Button", "x10 y130 w90", "Calculate").OnEvent("Click", CalcStenosis)
    StenosisGuiObj.Add("Button", "x110 y130 w80", "Cancel").OnEvent("Click", StenosisGuiClose)
    StenosisGuiObj.OnEvent("Close", StenosisGuiClose)
    StenosisGuiObj.Show("x" xPos " y" yPos " w270 h170")
}

StenosisGuiClose(*) {
    global StenosisGuiObj
    if (StenosisGuiObj)
        StenosisGuiObj.Destroy()
    StenosisGuiObj := ""
}

CalcStenosis(*) {
    global StenosisGuiObj, ShowCitations
    saved := StenosisGuiObj.Submit(false)

    if (saved.StenosisNormal = "" || saved.StenosisMeasure = "") {
        MsgBox("Please enter both diameter measurements.", "Error", 16)
        return
    }

    normal := saved.StenosisNormal + 0
    stenosis := saved.StenosisMeasure + 0

    if (normal <= 0) {
        MsgBox("Normal diameter must be greater than 0.", "Error", 16)
        return
    }

    if (stenosis >= normal) {
        MsgBox("Stenosis should be less than normal diameter.", "Error", 16)
        return
    }

    ; Calculate percent stenosis
    stenosisPercent := ((normal - stenosis) / normal) * 100
    stenosisPercent := Round(stenosisPercent, 1)

    vesselName := saved.StenosisVessel != "" ? saved.StenosisVessel : "Vessel"

    ; WHY: Sentence style for inline insertion - leading space for dictation continuity
    ; ARCHITECTURE: Matches NASCET parser output format for consistency
    normalRound := Round(normal, 1)
    stenosisRound := Round(stenosis, 1)

    ; Build severity interpretation
    if (stenosisPercent < 50)
        severity := "mild stenosis"
    else if (stenosisPercent < 70)
        severity := "moderate stenosis"
    else
        severity := "severe stenosis"

    ; Sentence format: " Stenosis: X% (normal Ymm, stenosis Zmm), severity."
    result := " " vesselName " stenosis: " stenosisPercent "% (normal " normalRound "mm, stenosis " stenosisRound "mm), " severity "."

    StenosisGuiObj.Destroy()
    StenosisGuiObj := ""
    ShowResult(result)
}

; =========================================
; CALCULATOR 4: RV/LV Ratio
; =========================================
global RVLVGuiObj := ""

ShowRVLVGui() {
    global RVLVGuiObj
    CoordMode("Mouse", "Screen")
    MouseGetPos(&mouseX, &mouseY)
    xPos := mouseX + 10
    yPos := mouseY + 10

    RVLVGuiObj := Gui("+AlwaysOnTop")
    RVLVGuiObj.Title := "RV/LV Ratio"
    RVLVGuiObj.Add("Text", "x10 y10 w280", "RV/LV Ratio Calculator (4-chamber axial)")
    RVLVGuiObj.Add("Text", "x10 y40", "RV diameter (mm):")
    RVLVGuiObj.Add("Edit", "x130 y37 w60 vRVDiam")
    RVLVGuiObj.Add("Text", "x10 y70", "LV diameter (mm):")
    RVLVGuiObj.Add("Edit", "x130 y67 w60 vLVDiam")
    RVLVGuiObj.Add("Button", "x10 y105 w90", "Calculate").OnEvent("Click", CalcRVLV)
    RVLVGuiObj.Add("Button", "x110 y105 w80", "Cancel").OnEvent("Click", RVLVGuiClose)
    RVLVGuiObj.OnEvent("Close", RVLVGuiClose)
    RVLVGuiObj.Show("x" xPos " y" yPos " w220 h145")
}

RVLVGuiClose(*) {
    global RVLVGuiObj
    if (RVLVGuiObj)
        RVLVGuiObj.Destroy()
    RVLVGuiObj := ""
}

CalcRVLV(*) {
    global RVLVGuiObj, ShowCitations
    saved := RVLVGuiObj.Submit(false)

    if (saved.RVDiam = "" || saved.LVDiam = "") {
        MsgBox("Please enter both RV and LV diameters.", "Error", 16)
        return
    }

    rv := saved.RVDiam + 0
    lv := saved.LVDiam + 0

    if (lv <= 0) {
        MsgBox("LV diameter must be greater than 0.", "Error", 16)
        return
    }

    ratio := rv / lv
    ratio := Round(ratio, 2)

    ; WHY: Sentence style for inline insertion - matches smart parser output
    ; ARCHITECTURE: Leading space for dictation continuity
    rvRound := Round(rv, 1)
    lvRound := Round(lv, 1)

    ; Build interpretation
    if (ratio >= 1.0)
        interpretation := "significant right heart strain"
    else if (ratio >= 0.9)
        interpretation := "suggestive of right heart strain"
    else
        interpretation := "within normal limits"

    ; Sentence format: " RV/LV ratio: X (RV Ymm, LV Zmm), interpretation."
    result := " RV/LV ratio: " ratio " (RV " rvRound "mm, LV " lvRound "mm), " interpretation "."

    if (ShowCitations)
        result .= " (Meinel et al. Radiology 2015)"

    RVLVGuiObj.Destroy()
    RVLVGuiObj := ""
    ShowResult(result)
}

; =========================================
; CALCULATOR 5: Fleischner 2017
; =========================================
global FleischnerGuiObj := ""

ShowFleischnerGui() {
    global FleischnerGuiObj, IncludeDatamining
    CoordMode("Mouse", "Screen")
    MouseGetPos(&mouseX, &mouseY)
    xPos := mouseX + 10
    yPos := mouseY + 10

    FleischnerGuiObj := Gui("+AlwaysOnTop")
    FleischnerGuiObj.Title := "Fleischner 2017"
    FleischnerGuiObj.Add("Text", "x10 y10 w300", "Fleischner 2017 - Incidental Pulmonary Nodule")
    FleischnerGuiObj.Add("Text", "x10 y40", "Nodule size (mm):")
    FleischnerGuiObj.Add("Edit", "x130 y37 w60 vFleischSize")
    FleischnerGuiObj.Add("Text", "x10 y70", "Nodule type:")
    FleischnerGuiObj.Add("DropDownList", "x130 y67 w120 vFleischType Choose1", ["Solid", "Part-solid", "Ground glass"])
    FleischnerGuiObj.Add("Text", "x10 y100", "Number:")
    FleischnerGuiObj.Add("DropDownList", "x130 y97 w120 vFleischNumber Choose1", ["Single", "Multiple"])
    FleischnerGuiObj.Add("Text", "x10 y130", "Risk:")
    FleischnerGuiObj.Add("DropDownList", "x130 y127 w120 vFleischRisk Choose1", ["Low risk", "High risk"])
    dmChecked := IncludeDatamining ? "Checked" : ""
    FleischnerGuiObj.Add("Checkbox", "x10 y160 w250 vFleischDatamine " dmChecked, "Include datamining phrase (SSM lung nodule)")
    FleischnerGuiObj.Add("Button", "x10 y190 w100", "Get Recommendation").OnEvent("Click", CalcFleischner)
    FleischnerGuiObj.Add("Button", "x120 y190 w80", "Cancel").OnEvent("Click", FleischnerGuiClose)
    FleischnerGuiObj.OnEvent("Close", FleischnerGuiClose)
    FleischnerGuiObj.Show("x" xPos " y" yPos " w280 h230")
}

FleischnerGuiClose(*) {
    global FleischnerGuiObj
    if (FleischnerGuiObj)
        FleischnerGuiObj.Destroy()
    FleischnerGuiObj := ""
}

CalcFleischner(*) {
    global FleischnerGuiObj, DataminingPhrase, ShowCitations
    saved := FleischnerGuiObj.Submit(false)

    if (saved.FleischSize = "") {
        MsgBox("Please enter nodule size.", "Error", 16)
        return
    }

    size := saved.FleischSize + 0
    result := "Fleischner 2017 Recommendation:`n"
    result .= "Size: " size " mm | " saved.FleischType " | " saved.FleischNumber " | " saved.FleischRisk "`n`n"

    ; Determine recommendation based on 2017 Fleischner guidelines
    recommendation := GetFleischnerRecommendation(size, saved.FleischType, saved.FleischNumber, saved.FleischRisk)
    result .= recommendation

    if (saved.FleischDatamine) {
        result .= "`n`n[" DataminingPhrase "]"
    }

    if (ShowCitations)
        result .= "`n`nRef: MacMahon H et al. Radiology 2017;284:228-243"

    FleischnerGuiObj.Destroy()
    FleischnerGuiObj := ""
    ShowResult(result)
}

GetFleischnerRecommendation(size, type, number, risk) {
    ; Solid nodules
    if (type = "Solid") {
        if (number = "Single") {
            if (size < 6) {
                return "No routine follow-up recommended.`n(Optional CT at 12 months if high-risk)"
            } else if (size < 8) {
                if (risk = "Low risk")
                    return "CT at 6-12 months, then consider CT at 18-24 months."
                else
                    return "CT at 6-12 months, then CT at 18-24 months."
            } else {
                return "Consider CT at 3 months, PET/CT, or tissue sampling."
            }
        } else {  ; Multiple
            if (size < 6) {
                return "No routine follow-up recommended.`n(Optional CT at 12 months if high-risk)"
            } else {
                if (risk = "Low risk")
                    return "CT at 3-6 months, then consider CT at 18-24 months."
                else
                    return "CT at 3-6 months, then CT at 18-24 months."
            }
        }
    }
    ; Part-solid nodules
    else if (type = "Part-solid") {
        if (number = "Single") {
            if (size < 6) {
                return "No routine follow-up recommended."
            } else {
                return "CT at 3-6 months. If stable, annual CT for 5 years.`nIf solid component >=6 mm, consider PET/CT or biopsy."
            }
        } else {
            return "CT at 3-6 months. If stable, annual CT for 5 years."
        }
    }
    ; Ground glass nodules
    else if (type = "Ground glass") {
        if (number = "Single") {
            if (size < 6) {
                return "No routine follow-up recommended."
            } else {
                return "CT at 6-12 months, then every 2 years for 5 years."
            }
        } else {
            return "CT at 3-6 months. If stable, consider CT at 2 and 4 years."
        }
    }
    return "Unable to determine recommendation."
}

; =========================================
; PRE-FILLED GUI FUNCTIONS
; WHY: Allow user to edit parsed values in GUI when confidence is low
; =========================================

ShowEllipsoidVolumeGuiPrefilled(d1, d2, d3) {
    global EllipsoidGuiObj
    CoordMode("Mouse", "Screen")
    MouseGetPos(&mouseX, &mouseY)
    xPos := mouseX + 10
    yPos := mouseY + 10

    ; Pre-fill values (convert to string)
    dim1 := Round(d1, 2)
    dim2 := Round(d2, 2)
    dim3 := Round(d3, 2)

    EllipsoidGuiObj := Gui("+AlwaysOnTop")
    EllipsoidGuiObj.Title := "Ellipsoid Volume"
    EllipsoidGuiObj.Add("Text", "x10 y10 w280", "Ellipsoid Volume Calculator (Pre-filled)")
    EllipsoidGuiObj.Add("Text", "x10 y35", "AP (L):")
    EllipsoidGuiObj.Add("Edit", "x80 y32 w50 vEllipDim1", dim1)
    EllipsoidGuiObj.Add("Text", "x135 y35", "x   T (W):")
    EllipsoidGuiObj.Add("Edit", "x185 y32 w50 vEllipDim2", dim2)
    EllipsoidGuiObj.Add("Text", "x10 y60", "CC (H):")
    EllipsoidGuiObj.Add("Edit", "x80 y57 w50 vEllipDim3", dim3)
    EllipsoidGuiObj.Add("Text", "x135 y60", "Units:")
    EllipsoidGuiObj.Add("DropDownList", "x185 y57 w50 vEllipUnits Choose2", ["mm", "cm"])
    EllipsoidGuiObj.Add("Button", "x10 y95 w100", "Calculate").OnEvent("Click", CalcEllipsoid)
    EllipsoidGuiObj.Add("Button", "x120 y95 w80", "Cancel").OnEvent("Click", EllipsoidGuiClose)
    EllipsoidGuiObj.OnEvent("Close", EllipsoidGuiClose)
    EllipsoidGuiObj.Show("x" xPos " y" yPos " w250 h135")
}

ShowRVLVGuiPrefilled(rv, lv) {
    global RVLVGuiObj
    CoordMode("Mouse", "Screen")
    MouseGetPos(&mouseX, &mouseY)
    xPos := mouseX + 10
    yPos := mouseY + 10

    rvVal := Round(rv, 1)
    lvVal := Round(lv, 1)

    RVLVGuiObj := Gui("+AlwaysOnTop")
    RVLVGuiObj.Title := "RV/LV Ratio"
    RVLVGuiObj.Add("Text", "x10 y10 w280", "RV/LV Ratio Calculator (Pre-filled)")
    RVLVGuiObj.Add("Text", "x10 y40", "RV diameter (mm):")
    RVLVGuiObj.Add("Edit", "x130 y37 w60 vRVDiam", rvVal)
    RVLVGuiObj.Add("Text", "x10 y70", "LV diameter (mm):")
    RVLVGuiObj.Add("Edit", "x130 y67 w60 vLVDiam", lvVal)
    RVLVGuiObj.Add("Button", "x10 y105 w90", "Calculate").OnEvent("Click", CalcRVLV)
    RVLVGuiObj.Add("Button", "x110 y105 w80", "Cancel").OnEvent("Click", RVLVGuiClose)
    RVLVGuiObj.OnEvent("Close", RVLVGuiClose)
    RVLVGuiObj.Show("x" xPos " y" yPos " w220 h145")
}

ShowNASCETGuiPrefilled(distal, stenosis) {
    global NASCETGuiObj
    CoordMode("Mouse", "Screen")
    MouseGetPos(&mouseX, &mouseY)
    xPos := mouseX + 10
    yPos := mouseY + 10

    distalVal := Round(distal, 1)
    stenosisVal := Round(stenosis, 1)

    NASCETGuiObj := Gui("+AlwaysOnTop")
    NASCETGuiObj.Title := "NASCET Calculator"
    NASCETGuiObj.Add("Text", "x10 y10 w280", "NASCET Calculator (Pre-filled)")
    NASCETGuiObj.Add("Text", "x10 y40", "Distal ICA diameter (mm):")
    NASCETGuiObj.Add("Edit", "x160 y37 w60 vNASCETDistal", distalVal)
    NASCETGuiObj.Add("Text", "x10 y70", "Stenosis diameter (mm):")
    NASCETGuiObj.Add("Edit", "x160 y67 w60 vNASCETStenosis", stenosisVal)
    NASCETGuiObj.Add("Button", "x10 y105 w90", "Calculate").OnEvent("Click", CalcNASCETGui)
    NASCETGuiObj.Add("Button", "x110 y105 w80", "Cancel").OnEvent("Click", NASCETGuiClose)
    NASCETGuiObj.OnEvent("Close", NASCETGuiClose)
    NASCETGuiObj.Show("x" xPos " y" yPos " w240 h145")
}

ShowAdrenalWashoutGuiPrefilled(pre, post, delayed) {
    global AdrenalGuiObj
    CoordMode("Mouse", "Screen")
    MouseGetPos(&mouseX, &mouseY)
    xPos := mouseX + 10
    yPos := mouseY + 10

    preVal := (pre != "") ? Round(pre, 0) : ""
    postVal := Round(post, 0)
    delayedVal := Round(delayed, 0)

    AdrenalGuiObj := Gui("+AlwaysOnTop")
    AdrenalGuiObj.Title := "Adrenal Washout"
    AdrenalGuiObj.Add("Text", "x10 y10 w280", "Adrenal Washout Calculator (Pre-filled)")
    AdrenalGuiObj.Add("Text", "x10 y40", "Pre-contrast (HU):")
    AdrenalGuiObj.Add("Edit", "x140 y37 w70 vAdrenalPre", preVal)
    AdrenalGuiObj.Add("Text", "x10 y70", "Post-contrast (HU):")
    AdrenalGuiObj.Add("Edit", "x140 y67 w70 vAdrenalPost", postVal)
    AdrenalGuiObj.Add("Text", "x10 y100", "Delayed (15 min) (HU):")
    AdrenalGuiObj.Add("Edit", "x140 y97 w70 vAdrenalDelayed", delayedVal)
    AdrenalGuiObj.Add("Button", "x10 y135 w90", "Calculate").OnEvent("Click", CalcAdrenal)
    AdrenalGuiObj.Add("Button", "x110 y135 w80", "Cancel").OnEvent("Click", AdrenalGuiClose)
    AdrenalGuiObj.OnEvent("Close", AdrenalGuiClose)
    AdrenalGuiObj.Show("x" xPos " y" yPos " w230 h175")
}

ShowFleischnerGuiPrefilled(size, nodeType, number) {
    global FleischnerGuiObj, IncludeDatamining
    CoordMode("Mouse", "Screen")
    MouseGetPos(&mouseX, &mouseY)
    xPos := mouseX + 10
    yPos := mouseY + 10

    sizeVal := Round(size, 0)

    ; Determine dropdown selection for type
    typeSelect := 1
    if (InStr(nodeType, "Part") || InStr(nodeType, "part"))
        typeSelect := 2
    else if (InStr(nodeType, "Ground") || InStr(nodeType, "glass") || InStr(nodeType, "GGN"))
        typeSelect := 3

    ; Determine dropdown selection for number
    numberSelect := (number = "Multiple") ? 2 : 1

    FleischnerGuiObj := Gui("+AlwaysOnTop")
    FleischnerGuiObj.Title := "Fleischner 2017"
    FleischnerGuiObj.Add("Text", "x10 y10 w300", "Fleischner 2017 (Pre-filled)")
    FleischnerGuiObj.Add("Text", "x10 y40", "Nodule size (mm):")
    FleischnerGuiObj.Add("Edit", "x130 y37 w60 vFleischSize", sizeVal)
    FleischnerGuiObj.Add("Text", "x10 y70", "Nodule type:")
    FleischnerGuiObj.Add("DropDownList", "x130 y67 w120 vFleischType Choose" typeSelect, ["Solid", "Part-solid", "Ground glass"])
    FleischnerGuiObj.Add("Text", "x10 y100", "Number:")
    FleischnerGuiObj.Add("DropDownList", "x130 y97 w120 vFleischNumber Choose" numberSelect, ["Single", "Multiple"])
    FleischnerGuiObj.Add("Text", "x10 y130", "Risk:")
    FleischnerGuiObj.Add("DropDownList", "x130 y127 w120 vFleischRisk Choose1", ["Low risk", "High risk"])
    dmChecked := IncludeDatamining ? "Checked" : ""
    FleischnerGuiObj.Add("Checkbox", "x10 y160 w250 vFleischDatamine " dmChecked, "Include datamining phrase")
    FleischnerGuiObj.Add("Button", "x10 y190 w100", "Get Recommendation").OnEvent("Click", CalcFleischner)
    FleischnerGuiObj.Add("Button", "x120 y190 w80", "Cancel").OnEvent("Click", FleischnerGuiClose)
    FleischnerGuiObj.OnEvent("Close", FleischnerGuiClose)
    FleischnerGuiObj.Show("x" xPos " y" yPos " w280 h230")
}

; -----------------------------------------
; Utility: Get mouse position for GUI placement
; WHY: Reduces code duplication across 15+ GUI functions
; -----------------------------------------
GetGuiPosition(&xPos, &yPos, offset := 10) {
    CoordMode("Mouse", "Screen")
    MouseGetPos(&mouseX, &mouseY)
    xPos := mouseX + offset
    yPos := mouseY + offset
}

; =========================================
; SMART PARSE FUNCTIONS
; WHY: Parse dictated text and insert calculations inline without GUI re-entry.
; ARCHITECTURE: Text parsing with regex, calculation, and clipboard-based insertion.
; =========================================

; -----------------------------------------
; Utility: Paste text while preserving clipboard
; WHY: Reduces duplication and ensures clipboard is always restored
; TRADEOFF: Slight overhead for save/restore, but safer
; -----------------------------------------
PasteTextPreserveClipboard(text, waitMs := 100) {
    ClipSaved := ClipboardAll()
    A_Clipboard := text
    ClipWait(0.5)
    Send("^v")
    Sleep(waitMs)
    A_Clipboard := ClipSaved
}

; -----------------------------------------
; Utility: Insert text after current selection
; WHY: Preserves original text and appends calculation result.
; -----------------------------------------
InsertAfterSelection(textToInsert) {
    ; Move cursor to end of selection and insert text
    Send("{Right}")
    Sleep(50)

    ; Use helper to paste text while preserving clipboard
    PasteTextPreserveClipboard(textToInsert)

    ToolTip("Calculation inserted!")
    SetTimer(RemoveToolTip, -1500)
}

; -----------------------------------------
; Utility: Insert text after IMPRESSION: field in document
; WHY: User wants Fleischner recommendations after the impression section
; ARCHITECTURE: Uses Find dialog to locate IMPRESSION:, then inserts below
; TRADEOFF: May not work in all applications; check settings to disable
; NOTE: If IMPRESSION: not found, falls back to inserting after selection
; -----------------------------------------
InsertAtImpression(textToInsert) {
    ; Save current clipboard and cursor position marker
    ClipSaved := ClipboardAll()

    ; Copy document to check for IMPRESSION: existence
    ; NOTE: This briefly exposes full document on clipboard; cleared immediately after check
    ; TRADEOFF: Faster than incremental search, clipboard cleared right after copy
    A_Clipboard := ""
    Send("^a")  ; Select all
    Sleep(20)
    Send("^c")  ; Copy
    ClipWait(0.5)
    docText := A_Clipboard
    A_Clipboard := ""  ; Clear clipboard immediately to minimize PHI exposure

    ; CRITICAL: Deselect document text before any further operations
    ; WHY: Prevents accidental replacement if Find dialog is slow to open
    Send("{Escape}")
    Sleep(30)
    Send("{Right}")
    Sleep(30)

    ; WHY: Match variations of IMPRESSION header used in different templates
    ; TRADEOFF: Multiple checks vs comprehensive coverage
    searchTerm := ""
    if (InStr(docText, "CLINICAL IMPRESSION:"))
        searchTerm := "CLINICAL IMPRESSION:"
    else if (InStr(docText, "IMPRESSIONS:"))
        searchTerm := "IMPRESSIONS:"
    else if (InStr(docText, "IMPRESSION:"))
        searchTerm := "IMPRESSION:"
    else if (InStr(docText, "IMPRESSION"))
        searchTerm := "IMPRESSION"  ; Match without colon as fallback

    if (searchTerm = "") {
        ; IMPRESSION not found - fall back to insert after selection
        A_Clipboard := ClipSaved  ; Restore clipboard first
        InsertAfterSelection(textToInsert)
        ToolTip("IMPRESSION: not found - inserted at cursor")
        SetTimer(RemoveToolTip, -2000)
        return
    }

    ; Go to start of document
    Send("^{Home}")
    Sleep(50)

    ; Open Find dialog (Ctrl+F)
    Send("^f")
    Sleep(500)  ; WHY: PowerScribe Find dialog can be very slow to open

    ; Clear any previous search and type new search text
    ; NOTE: ^a here targets the Find dialog's text field, not the document
    Send("^a")
    Sleep(50)
    Send("{Raw}" searchTerm)
    Sleep(200)

    ; Press Enter to find (or F3/Find Next depending on app)
    Send("{Enter}")
    Sleep(150)

    ; Close Find dialog (Escape) - press twice to ensure closed
    Send("{Escape}")
    Sleep(100)
    Send("{Escape}")
    Sleep(100)

    ; Go to end of line where IMPRESSION: was found
    Send("{End}")
    Sleep(50)

    ; Add blank line (Enter twice for spacing)
    Send("{Enter}{Enter}")
    Sleep(30)

    ; Set clipboard to text and paste
    A_Clipboard := textToInsert
    ClipWait(0.5)
    Send("^v")
    Sleep(100)

    ; Restore clipboard
    A_Clipboard := ClipSaved

    ToolTip("Inserted after IMPRESSION:")
    SetTimer(RemoveToolTip, -1500)
}

; -----------------------------------------
; Utility: Deduplicate array of sizes (removes duplicates within 1mm tolerance)
; WHY: Multiple patterns may match same nodule, avoid double-counting
; -----------------------------------------
DeduplicateSizes(arr) {
    if (arr.Length <= 1)
        return arr

    result := []
    for i, size in arr {
        isDuplicate := false
        for j, existing in result {
            ; Consider sizes within 1mm as duplicates
            if (Abs(size - existing) <= 1) {
                isDuplicate := true
                ; Keep the larger one
                if (size > existing)
                    result[j] := size
                break
            }
        }
        if (!isDuplicate)
            result.Push(size)
    }
    return result
}

; -----------------------------------------
; Utility: Filter out non-measurement numbers
; WHY: Prevents picking up image numbers (23/75), slice numbers, etc.
; TRADEOFF: May occasionally filter legitimate measurements, but safety > convenience
; -----------------------------------------
FilterNonMeasurements(input) {
    filtered := input

    ; Remove image/slice references: (image 23/75), (I23), slice 45, series 3
    filtered := RegExReplace(filtered, "i)\(\s*(?:image|img|i|slice|ser|seq)\s*\d+\s*/\s*\d+\s*\)", "")
    filtered := RegExReplace(filtered, "i)(?:image|img|slice|series|seq|level)\s*[:=#]?\s*\d+(?:\s*/\s*\d+)?", "")

    ; Remove window/level settings: W:400 L:40, WW/WL
    filtered := RegExReplace(filtered, "i)[WL]\s*[:=]?\s*-?\d+", "")
    filtered := RegExReplace(filtered, "i)(?:WW|WL)\s*[:=]?\s*-?\d+", "")

    ; Remove dates: 1/15/2024, 2024-01-15
    filtered := RegExReplace(filtered, "\d{1,2}/\d{1,2}/\d{2,4}", "")
    filtered := RegExReplace(filtered, "\d{4}-\d{1,2}-\d{1,2}", "")

    ; Remove page numbers: page 3 of 10, 3/10
    filtered := RegExReplace(filtered, "i)page\s*\d+\s*(?:of|/)\s*\d+", "")

    ; Remove accession/MRN patterns: ACC# 123456, MRN: 123456
    filtered := RegExReplace(filtered, "i)(?:ACC|MRN|ID)\s*#?\s*:?\s*\d+", "")

    return filtered
}

; -----------------------------------------
; Utility: Show confirmation dialog before inserting parsed result
; WHY: User verification prevents wrong insertions
; RETURNS: "insert", "edit", or "cancel"
; -----------------------------------------
global ConfirmGuiObj := ""

ShowParseConfirmation(parseType, parsedValues, calculatedResult, confidence) {
    global SmartParseFallbackToGUI, g_ConfirmAction, g_ParsedResultText

    ; Store result for later insertion
    g_ParsedResultText := calculatedResult

    ; Build display text
    displayText := "Parsed " . parseType . ":`n`n"
    for key, val in parsedValues {
        displayText .= key . ": " . val . "`n"
    }
    displayText .= "`nResult: " . calculatedResult

    ; Add confidence indicator
    if (confidence < 50) {
        displayText .= "`n`n[!] LOW CONFIDENCE - Values may be incorrect"
    } else if (confidence < 80) {
        displayText .= "`n`n[~] Medium confidence"
    } else {
        displayText .= "`n`n[OK] High confidence"
    }

    ; Build GUI
    global ConfirmGuiObj
    CoordMode("Mouse", "Screen")
    MouseGetPos(&mouseX, &mouseY)
    xPos := mouseX + 10
    yPos := mouseY + 10

    ; Destroy any existing confirm GUI
    if (ConfirmGuiObj != "")
        ConfirmGuiObj.Destroy()

    ConfirmGuiObj := Gui("+AlwaysOnTop")
    ConfirmGuiObj.OnEvent("Close", ConfirmGuiClose)
    ConfirmGuiObj.OnEvent("Escape", ConfirmGuiClose)
    ConfirmGuiObj.Add("Text", "x10 y10 w350 +0x1", "Smart Parse Confirmation")
    ConfirmGuiObj.Add("Edit", "x10 y35 w350 h130 ReadOnly", displayText)

    ; Different buttons based on confidence
    if (confidence < 50 && SmartParseFallbackToGUI) {
        ; Low confidence: default to Cancel, offer GUI edit
        ConfirmGuiObj.Add("Button", "x10 y175 w100", "Insert Anyway").OnEvent("Click", ConfirmInsert)
        ConfirmGuiObj.Add("Button", "x120 y175 w110 Default", "Edit in GUI").OnEvent("Click", ConfirmEdit)
        ConfirmGuiObj.Add("Button", "x240 y175 w100", "Cancel").OnEvent("Click", ConfirmCancel)
    } else {
        ; High/medium confidence: default to Insert
        ConfirmGuiObj.Add("Button", "x10 y175 w100 Default", "Insert").OnEvent("Click", ConfirmInsert)
        ConfirmGuiObj.Add("Button", "x120 y175 w110", "Edit in GUI").OnEvent("Click", ConfirmEdit)
        ConfirmGuiObj.Add("Button", "x240 y175 w100", "Cancel").OnEvent("Click", ConfirmCancel)
    }

    ConfirmGuiObj.Show("x" xPos " y" yPos " w380 h215")
    ConfirmGuiObj.Title := "Confirm Parse"
    hWnd := ConfirmGuiObj.Hwnd

    ; Reset and wait for user action (event-driven via WinWaitClose)
    g_ConfirmAction := ""
    WinWaitClose("ahk_id " hWnd)

    return g_ConfirmAction
}

ConfirmInsert(*) {
    global g_ConfirmAction, ConfirmGuiObj
    g_ConfirmAction := "insert"
    ConfirmGuiObj.Destroy()
}

ConfirmEdit(*) {
    global g_ConfirmAction, ConfirmGuiObj
    g_ConfirmAction := "edit"
    ConfirmGuiObj.Destroy()
}

ConfirmCancel(*) {
    global g_ConfirmAction, ConfirmGuiObj
    g_ConfirmAction := "cancel"
    ConfirmGuiObj.Destroy()
}

ConfirmGuiClose(*) {
    global g_ConfirmAction, ConfirmGuiObj
    g_ConfirmAction := "cancel"
    ConfirmGuiObj.Destroy()
}

; -----------------------------------------
; SMART VOLUME PARSER
; WHY: Parse "Prostate measures 8.0 x 6.0 x 9.0 cm" and insert volume with organ-specific interpretation.
; ARCHITECTURE: Strict patterns with confidence scoring, confirmation dialog, GUI fallback.
; -----------------------------------------
ParseAndInsertVolume(input) {
    global DefaultMeasurementUnit
    ; Step 1: Clean input and filter non-measurements
    input := RegExReplace(input, "`r?\n", " ")
    filtered := FilterNonMeasurements(input)

    ; Step 2: Try patterns from strictest to loosest
    confidence := 0
    organ := ""
    d1 := 0
    d2 := 0
    d3 := 0
    units := ""  ; Will be set by pattern or default

    ; Pattern A (HIGH confidence 95): Organ + "measures" keyword + dimensions + units
    ; Example: "Prostate measures 8.0 x 6.0 x 9.0 cm"
    strictPattern := "i)(prostate|abscess|lesion|cyst|mass|nodule|collection|hematoma|liver|spleen|kidney|bladder|uterus|ovary|thyroid|adrenal)\s+(?:measures?|measuring|sized?|is)\s+(\d+(?:\.\d+)?)\s*[x×]\s*(\d+(?:\.\d+)?)\s*[x×]\s*(\d+(?:\.\d+)?)\s*(cm|mm)"

    if (RegExMatch(filtered, strictPattern, &m)) {
        organ := m[1]
        d1 := m[2]
        d2 := m[3]
        d3 := m[4]
        units := m[5]
        confidence := 95
    }
    ; Pattern B (MEDIUM confidence 70): "measures" keyword + dimensions + units (no organ)
    else if (RegExMatch(filtered, "i)(?:measures?|measuring)\s+(\d+(?:\.\d+)?)\s*[x×]\s*(\d+(?:\.\d+)?)\s*[x×]\s*(\d+(?:\.\d+)?)\s*(cm|mm)", &m)) {
        d1 := m[1]
        d2 := m[2]
        d3 := m[3]
        units := m[4]
        confidence := 70
    }
    ; Pattern C (MEDIUM confidence 60): dimensions + units (no keyword)
    else if (RegExMatch(filtered, "(\d+(?:\.\d+)?)\s*[x×]\s*(\d+(?:\.\d+)?)\s*[x×]\s*(\d+(?:\.\d+)?)\s*(cm|mm)", &m)) {
        d1 := m[1]
        d2 := m[2]
        d3 := m[3]
        units := m[4]
        confidence := 60
    }
    ; Pattern D (LOW confidence 35): just three numbers with x separator (no units)
    ; WHY: Use DefaultMeasurementUnit setting when units not detected
    else if (RegExMatch(filtered, "(\d+(?:\.\d+)?)\s*[x×]\s*(\d+(?:\.\d+)?)\s*[x×]\s*(\d+(?:\.\d+)?)", &m)) {
        d1 := m[1]
        d2 := m[2]
        d3 := m[3]
        units := DefaultMeasurementUnit  ; Use user's default setting
        confidence := 35
    }
    else {
        ; No match found
        MsgBox("Could not find dimensions in text.`n`nExpected: `"measures 8.0 x 6.0 x 9.0 cm`"`n`nTip: Include units (cm or mm) for better detection.", "Parse Error", 48)
        return
    }

    ; Step 3: Convert to numbers and store original values for display
    d1 := d1 + 0
    d2 := d2 + 0
    d3 := d3 + 0
    originalUnits := units
    origD1 := d1
    origD2 := d2
    origD3 := d3

    ; Convert to cm for volume calculation (internal)
    d1_cm := d1
    d2_cm := d2
    d3_cm := d3
    if (units = "mm") {
        d1_cm := d1 / 10
        d2_cm := d2 / 10
        d3_cm := d3 / 10
    }

    ; Step 4: Sanity check - dimensions should be plausible (<50cm)
    if (d1_cm > 50 || d2_cm > 50 || d3_cm > 50) {
        confidence := confidence - 30
    }

    ; Step 5: Calculate ellipsoid volume: (π/6) × L × W × H (using cm values)
    volume := 0.5236 * d1_cm * d2_cm * d3_cm
    volumeRounded := Round(volume, 1)

    ; Step 6: Build result string
    resultText := " This corresponds to a volume of " . volumeRounded . " cc (mL)"

    ; Add organ-specific interpretation for prostate
    organLower := ""
    if (organ != "") {
        organLower := StrLower(organ)
    }
    if (organLower = "prostate") {
        if (volume < 30) {
            resultText .= ", within normal limits."
        } else if (volume < 50) {
            resultText .= ", compatible with an enlarged prostate."
        } else if (volume < 70) {
            resultText .= ", compatible with a moderately enlarged prostate."
        } else {
            resultText .= ", compatible with a massively enlarged prostate."
        }
    } else {
        resultText .= "."
    }

    ; Step 7: Show confirmation dialog (display in original units)
    parsedValues := {}
    parsedValues["Organ"] := organ != "" ? organ : "Not detected"
    parsedValues["Dim 1"] := Round(origD1, 2) . " " . originalUnits
    parsedValues["Dim 2"] := Round(origD2, 2) . " " . originalUnits
    parsedValues["Dim 3"] := Round(origD3, 2) . " " . originalUnits

    action := ShowParseConfirmation("Volume", parsedValues, resultText, confidence)

    if (action = "insert") {
        InsertAfterSelection(resultText)
    } else if (action = "edit") {
        ; Open GUI with pre-filled values (in cm for internal use)
        ShowEllipsoidVolumeGuiPrefilled(d1_cm, d2_cm, d3_cm)
    }
    ; else cancelled - do nothing
}

; -----------------------------------------
; HELPER: Normalize measurement units
; WHY: Standardize various unit formats to mm or cm
; -----------------------------------------
NormalizeUnits(units) {
    if (units = "" || units = "mm" || units = "millimeter" || units = "millimeters")
        return "mm"
    return "cm"
}

; -----------------------------------------
; SMART RV/LV RATIO PARSER
; WHY: Parse RV/LV measurements and insert ratio with PE risk interpretation.
; ARCHITECTURE: Robust multi-pattern matching for various radiology formats.
; TRADEOFF: Macro format uses cm for PowerScribe compatibility; Inline uses mm.
; -----------------------------------------
ParseAndInsertRVLV(input) {
    global ShowCitations, RVLVOutputFormat
    input := RegExReplace(input, "`r?\n", " ")
    filtered := FilterNonMeasurements(input)

    rv := 0
    lv := 0
    rvUnits := "mm"
    lvUnits := "mm"
    confidence := 0

    ; === COMPREHENSIVE RV KEYWORDS ===
    ; Covers: RV, R.V., Right ventricle, Right ventricular, Rt ventricle, Rt. ventricle
    ; Also: "right ventricular diameter", "RV short axis", etc.
    rvKeywords := "(?:RV|R\.V\.|Right\s*(?:ventricle|ventricular|heart)|Rt\.?\s*(?:ventricle|ventricular))"

    ; === COMPREHENSIVE LV KEYWORDS ===
    lvKeywords := "(?:LV|L\.V\.|Left\s*(?:ventricle|ventricular|heart)|Lt\.?\s*(?:ventricle|ventricular))"

    ; === MEASUREMENT QUALIFIERS ===
    ; Things that might appear between keyword and number
    measureQualifiers := "(?:diameter|diam\.?|short[- ]?axis|transverse|dimension|width|size|measurement)?"
    verbQualifiers := "(?:is|are|was|measures?|measuring|=|:)?"

    ; === SIZE PATTERN ===
    ; Handles: "42", "42mm", "42 mm", "4.2cm", "4.2 cm", ".9 cm" (missing leading zero)
    sizeWithUnits := "(\d*\.?\d+)\s*(mm|cm|millimeters?|centimeters?)?"

    ; === PATTERN A: RV ... LV (most common) ===
    ; Examples: "RV 42mm / LV 35mm", "RV: 42, LV: 35", "Right ventricle 4.2 cm, Left ventricle 3.5 cm"
    ; "RV diameter: 42 mm LV diameter: 35 mm", "RV = 42mm; LV = 35mm"
    ; "The RV measures 42mm and the LV measures 35mm"
    patternA := "i)" . rvKeywords . "\s*" . measureQualifiers . "\s*" . verbQualifiers . "\s*" . sizeWithUnits . ".*?" . lvKeywords . "\s*" . measureQualifiers . "\s*" . verbQualifiers . "\s*" . sizeWithUnits

    ; === PATTERN B: LV ... RV (reversed order) ===
    patternB := "i)" . lvKeywords . "\s*" . measureQualifiers . "\s*" . verbQualifiers . "\s*" . sizeWithUnits . ".*?" . rvKeywords . "\s*" . measureQualifiers . "\s*" . verbQualifiers . "\s*" . sizeWithUnits

    ; === PATTERN C: Ratio format ===
    ; Examples: "RV/LV ratio = 42/35", "RV:LV = 42:35", "RV to LV ratio 1.2"
    patternC := "i)(?:RV|Right)\s*[/:]?\s*(?:to\s*)?(?:LV|Left)\s*(?:ratio|diameter)?[:\s=]*(\d+(?:\.\d+)?)\s*[/:]\s*(\d+(?:\.\d+)?)"

    ; === PATTERN D: Table/structured format ===
    ; Examples: "RV    42 mm" followed by "LV    35 mm" (spaces/tabs between)
    patternD := "i)" . rvKeywords . "\s+(\d+(?:\.\d+)?)\s*(mm|cm)?\s+.*?" . lvKeywords . "\s+(\d+(?:\.\d+)?)\s*(mm|cm)?"

    ; === PATTERN E: Sentence with "and" ===
    ; Examples: "RV is 42mm and LV is 35mm", "right ventricle 42 and left ventricle 35"
    patternE := "i)" . rvKeywords . ".{0,20}(\d+(?:\.\d+)?)\s*(mm|cm)?\s*(?:and|,|;)\s*" . lvKeywords . ".{0,20}(\d+(?:\.\d+)?)\s*(mm|cm)?"

    ; === PATTERN F: Dilated/enlarged RV with measurement ===
    ; Examples: "dilated RV measuring 42mm", "enlarged right ventricle 4.2 cm"
    patternF := "i)(?:dilated|enlarged|prominent)\s+" . rvKeywords . ".{0,20}(\d+(?:\.\d+)?)\s*(mm|cm)?.*?" . lvKeywords . ".{0,30}(\d+(?:\.\d+)?)\s*(mm|cm)?"

    ; === PATTERN G: Axial/4-chamber specific ===
    ; Examples: "On axial images, RV 42mm, LV 35mm", "4-chamber view: RV 42 LV 35"
    patternG := "i)(?:axial|4[- ]?chamber|four[- ]?chamber).{0,30}" . rvKeywords . ".{0,15}(\d+(?:\.\d+)?)\s*(mm|cm)?.{0,30}" . lvKeywords . ".{0,15}(\d+(?:\.\d+)?)\s*(mm|cm)?"

    ; === PATTERN H: PACS table format (rvval/lvval) ===
    ; Examples: "rvval 5.0 lvval 3.6", "RVVAL 4.2 LVVAL 3.8"
    ; WHY: Some PACS systems output measurements in this abbreviated format
    ; NOTE: Values assumed to be in cm if no units specified (typical for cardiac CT)
    patternH := "i)rvval\s*(\d+(?:\.\d+)?)\s*(mm|cm)?\s*lvval\s*(\d+(?:\.\d+)?)\s*(mm|cm)?"

    ; Try patterns in order of confidence
    if (RegExMatch(filtered, patternA, &m)) {
        rv := m[1] + 0
        rvUnits := NormalizeUnits(m[2])
        lv := m[3] + 0
        lvUnits := NormalizeUnits(m[4])
        confidence := 90
    }
    else if (RegExMatch(filtered, patternB, &m)) {
        lv := m[1] + 0
        lvUnits := NormalizeUnits(m[2])
        rv := m[3] + 0
        rvUnits := NormalizeUnits(m[4])
        confidence := 85
    }
    else if (RegExMatch(filtered, patternE, &m)) {
        rv := m[1] + 0
        rvUnits := NormalizeUnits(m[2])
        lv := m[3] + 0
        lvUnits := NormalizeUnits(m[4])
        confidence := 85
    }
    else if (RegExMatch(filtered, patternG, &m)) {
        rv := m[1] + 0
        rvUnits := NormalizeUnits(m[2])
        lv := m[3] + 0
        lvUnits := NormalizeUnits(m[4])
        confidence := 85
    }
    else if (RegExMatch(filtered, patternF, &m)) {
        rv := m[1] + 0
        rvUnits := NormalizeUnits(m[2])
        lv := m[3] + 0
        lvUnits := NormalizeUnits(m[4])
        confidence := 80
    }
    else if (RegExMatch(filtered, patternC, &m)) {
        rv := m[1] + 0
        lv := m[2] + 0
        confidence := 80
    }
    else if (RegExMatch(filtered, patternD, &m)) {
        rv := m[1] + 0
        rvUnits := NormalizeUnits(m[2])
        lv := m[3] + 0
        lvUnits := NormalizeUnits(m[4])
        confidence := 75
    }
    else if (RegExMatch(filtered, patternH, &m)) {
        ; PACS "rvval/lvval" format - default to cm if no units (typical cardiac CT)
        rv := m[1] + 0
        rvUnits := (m[2] != "") ? NormalizeUnits(m[2]) : "cm"
        lv := m[3] + 0
        lvUnits := (m[4] != "") ? NormalizeUnits(m[4]) : "cm"
        confidence := 80
    }
    ; No pattern matched
    else {
        MsgBox("Could not find RV/LV measurements.`n`nExpected formats:`n- `"RV 42mm / LV 35mm`"`n- `"RV: 42, LV: 35`"`n- `"Right ventricle 4.2 cm`"`n- `"rvval 5.0 lvval 3.6`"`n`nMust include RV/LV keywords.", "Parse Error", 48)
        return
    }

    ; Convert cm to mm if needed
    if (rvUnits = "cm")
        rv := rv * 10
    if (lvUnits = "cm")
        lv := lv * 10

    ; Validate range (typical RV/LV are 20-60mm)
    if (rv < 10 || rv > 100 || lv < 10 || lv > 100) {
        confidence := confidence - 30
    }

    if (lv <= 0) {
        MsgBox("LV diameter must be greater than 0.", "Error", 48)
        return
    }

    ; Calculate ratio
    ratio := rv / lv
    ratio := Round(ratio, 2)

    ; Build interpretation text
    interpretation := ""
    if (ratio >= 1.0) {
        interpretation := "Significant right heart strain, suggestive of severe PE."
    } else if (ratio >= 0.9) {
        interpretation := "Suggestive of right heart strain."
    } else {
        interpretation := "Within normal limits."
    }

    ; WHY: Both formats now use sentence style for consistency
    ; ARCHITECTURE: Leading space for dictation continuity
    ; TRADEOFF: Macro uses cm (PowerScribe), Inline uses mm (brevity)
    resultText := ""
    interpretLower := StrLower(interpretation)
    if (RVLVOutputFormat = "Macro") {
        ; Macro format: sentence style with cm units for PowerScribe compatibility
        rv_cm := Round(rv / 10, 1)
        lv_cm := Round(lv / 10, 1)
        resultText := " RV/LV ratio: " . ratio . " (RV " . rv_cm . "cm, LV " . lv_cm . "cm), " . interpretLower
    } else {
        ; Inline format: brief sentence with mm units
        resultText := " RV/LV ratio: " . ratio . " (RV " . Round(rv, 1) . "mm, LV " . Round(lv, 1) . "mm), " . interpretLower
    }

    ; Show confirmation dialog
    parsedValues := {}
    parsedValues["RV"] := rv . " mm"
    parsedValues["LV"] := lv . " mm"
    parsedValues["Ratio"] := ratio

    action := ShowParseConfirmation("RV/LV Ratio", parsedValues, resultText, confidence)

    if (action = "insert") {
        InsertAfterSelection(resultText)
    } else if (action = "edit") {
        ShowRVLVGuiPrefilled(rv, lv)
    }
}

; -----------------------------------------
; SMART NASCET PARSER
; WHY: Parse stenosis measurements and insert NASCET percentage inline.
; ARCHITECTURE: Requires "distal" or "stenosis" keywords - no bare number fallback.
; -----------------------------------------
ParseAndInsertNASCET(input) {
    global ShowCitations
    input := RegExReplace(input, "`r?\n", " ")
    filtered := FilterNonMeasurements(input)

    distal := 0
    stenosis := 0
    distalUnits := "mm"
    stenosisUnits := "mm"
    confidence := 0

    ; Pattern A (HIGH confidence 90): "distal X mm ... stenosis Y mm"
    ; WHY: Captures units to handle mixed cm/mm measurements
    if (RegExMatch(filtered, "i)distal\s*(?:ICA|internal\s*carotid)?.*?(\d+(?:\.\d+)?)\s*(mm|cm)?.*?stenosis.*?(\d+(?:\.\d+)?)\s*(mm|cm)?", &m)) {
        distal := m[1] + 0
        distalUnits := (m[2] != "") ? m[2] : "mm"
        stenosis := m[3] + 0
        stenosisUnits := (m[4] != "") ? m[4] : "mm"
        confidence := 90
    }
    ; Pattern B (HIGH confidence 85): "stenosis X mm ... distal Y mm" (reverse order)
    else if (RegExMatch(filtered, "i)stenosis.*?(\d+(?:\.\d+)?)\s*(mm|cm)?.*?distal\s*(?:ICA)?.*?(\d+(?:\.\d+)?)\s*(mm|cm)?", &m)) {
        stenosis := m[1] + 0
        stenosisUnits := (m[2] != "") ? m[2] : "mm"
        distal := m[3] + 0
        distalUnits := (m[4] != "") ? m[4] : "mm"
        confidence := 85
    }
    ; Pattern C (MEDIUM confidence 65): ICA/carotid + narrowing context
    else if (RegExMatch(filtered, "i)(?:ICA|carotid).*?(\d+(?:\.\d+)?)\s*(mm|cm)?.*?(?:narrow|residual).*?(\d+(?:\.\d+)?)\s*(mm|cm)?", &m)) {
        d1 := m[1] + 0
        d1Units := (m[2] != "") ? m[2] : "mm"
        d2 := m[3] + 0
        d2Units := (m[4] != "") ? m[4] : "mm"
        ; Convert to mm before comparison
        d1_mm := (d1Units = "cm") ? d1 * 10 : d1
        d2_mm := (d2Units = "cm") ? d2 * 10 : d2
        ; Larger is distal, smaller is stenosis
        if (d1_mm > d2_mm) {
            distal := d1_mm
            stenosis := d2_mm
        } else {
            distal := d2_mm
            stenosis := d1_mm
        }
        ; Already converted to mm
        distalUnits := "mm"
        stenosisUnits := "mm"
        confidence := 65
    }
    ; NO FALLBACK TO JUST TWO NUMBERS - too risky
    else {
        MsgBox("Could not find stenosis measurements.`n`nExpected formats:`n- `"distal 5.2mm, stenosis 2.1mm`"`n- `"distal ICA 0.5cm, stenosis 3mm`"`n`nMust include `"distal`" or `"stenosis`" keywords.", "Parse Error", 48)
        return
    }

    ; Convert cm to mm for consistent calculation
    if (distalUnits = "cm")
        distal := distal * 10
    if (stenosisUnits = "cm")
        stenosis := stenosis * 10

    ; Validate
    if (distal <= 0) {
        MsgBox("Distal ICA diameter must be greater than 0.", "Error", 48)
        return
    }
    if (stenosis >= distal) {
        MsgBox("Stenosis diameter must be less than distal diameter.`n`nDistal: " distal " mm`nStenosis: " stenosis " mm", "Error", 48)
        return
    }

    ; Calculate NASCET
    nascetVal := ((distal - stenosis) / distal) * 100
    nascetVal := Round(nascetVal, 1)

    ; Validate plausible stenosis (0-99%)
    if (nascetVal < 0 || nascetVal > 99) {
        confidence := confidence - 30
    }

    ; Build result - inline sentence format for continued dictation
    ; Round values to 1 decimal place
    distalRound := Round(distal, 1)
    stenosisRound := Round(stenosis, 1)

    resultText := " NASCET: " . nascetVal . "% stenosis (distal " . distalRound . "mm, stenosis " . stenosisRound . "mm), "

    if (nascetVal < 50) {
        resultText .= "mild stenosis."
    } else if (nascetVal < 70) {
        resultText .= "moderate stenosis, consider intervention if symptomatic."
    } else {
        resultText .= "severe stenosis, strong indication for CEA/CAS."
    }
    ; Note: Citations removed per user preference - inline sentence format

    ; Show confirmation dialog
    parsedValues := {}
    parsedValues["Distal ICA"] := distal . " mm"
    parsedValues["Stenosis"] := stenosis . " mm"
    parsedValues["NASCET"] := nascetVal . "%"

    action := ShowParseConfirmation("NASCET Stenosis", parsedValues, resultText, confidence)

    if (action = "insert") {
        InsertAfterSelection(resultText)
    } else if (action = "edit") {
        ShowNASCETGuiPrefilled(distal, stenosis)
    }
}

; -----------------------------------------
; SMART ADRENAL WASHOUT PARSER
; WHY: Parse HU values and insert washout percentages inline.
; ARCHITECTURE: Requires "HU" labels - no bare number fallback.
; -----------------------------------------
ParseAndInsertAdrenalWashout(input) {
    global ShowCitations
    input := RegExReplace(input, "`r?\n", " ")
    filtered := FilterNonMeasurements(input)

    pre := ""
    post := ""
    delayed := ""
    confidence := 0

    ; Pattern A (HIGH confidence 95): Full pattern with phase labels + HU
    ; Example: "pre-contrast 10 HU, enhanced 80 HU, delayed 40 HU"
    fullPattern := "i)(?:pre-?contrast|unenhanced|baseline|native|non-?con)[:\s]*(-?\d+)\s*HU.*?(?:post-?contrast|enhanced|arterial|portal)[:\s]*(-?\d+)\s*HU.*?(?:delayed|15\s*min|late)[:\s]*(-?\d+)\s*HU"

    if (RegExMatch(filtered, fullPattern, &m)) {
        pre := m[1] + 0
        post := m[2] + 0
        delayed := m[3] + 0
        confidence := 95
    }
    ; Pattern B (MEDIUM confidence 70): Post + delayed with labels
    else if (RegExMatch(filtered, "i)(?:post-?contrast|enhanced)[:\s]*(-?\d+)\s*HU.*?(?:delayed|15\s*min)[:\s]*(-?\d+)\s*HU", &m)) {
        post := m[1] + 0
        delayed := m[2] + 0
        confidence := 70
    }
    ; Pattern C (MEDIUM confidence 55): Three HU values in sequence
    else if (RegExMatch(filtered, "(-?\d+)\s*HU.*?(-?\d+)\s*HU.*?(-?\d+)\s*HU", &m)) {
        pre := m[1] + 0
        post := m[2] + 0
        delayed := m[3] + 0
        confidence := 55
    }
    ; NO raw number fallback - requires HU labels for safety
    else {
        MsgBox("Could not find HU values.`n`nExpected formats:`n- `"pre-contrast 10 HU, enhanced 80 HU, delayed 40 HU`"`n- `"10 HU, 80 HU, 40 HU`"`n`nMust include `"HU`" labels.", "Parse Error", 48)
        return
    }

    if (post = "" || delayed = "") {
        MsgBox("At least post-contrast and delayed HU values are required.", "Error", 48)
        return
    }

    ; Validate HU ranges (typically -20 to 200)
    post := post + 0
    delayed := delayed + 0
    if (pre != "")
        pre := pre + 0

    if (post < -50 || post > 300 || delayed < -50 || delayed > 300) {
        confidence := confidence - 25
    }

    ; Build result - inline sentence format for continued dictation
    resultText := " Adrenal washout: "

    ; Calculate absolute washout if pre is available
    if (pre != "" && post != pre) {
        absWashout := ((post - delayed) / (post - pre)) * 100
        absWashout := Round(absWashout, 1)

        resultText .= "absolute " . absWashout . "%"
        if (absWashout >= 60)
            resultText .= " (likely adenoma)"
        else
            resultText .= " (indeterminate)"
        resultText .= ", "
    }

    ; Calculate relative washout
    if (post != 0) {
        relWashout := ((post - delayed) / post) * 100
        relWashout := Round(relWashout, 1)

        resultText .= "relative " . relWashout . "%"
        if (relWashout >= 40)
            resultText .= " (likely adenoma)"
        else
            resultText .= " (indeterminate)"
        resultText .= "."
    }

    ; Pre-contrast assessment
    if (pre != "" && pre <= 10) {
        resultText .= " Pre-contrast " . pre . " HU suggests lipid-rich adenoma."
    }
    ; Note: Citations removed per user preference - inline sentence format

    ; Show confirmation dialog
    parsedValues := {}
    parsedValues["Pre-contrast"] := (pre != "") ? pre . " HU" : "N/A"
    parsedValues["Post-contrast"] := post . " HU"
    parsedValues["Delayed"] := delayed . " HU"

    action := ShowParseConfirmation("Adrenal Washout", parsedValues, resultText, confidence)

    if (action = "insert") {
        InsertAfterSelection(resultText)
    } else if (action = "edit") {
        ShowAdrenalWashoutGuiPrefilled(pre, post, delayed)
    }
}

; -----------------------------------------
; HELPER: Get larger of two size values
; WHY: Nodule ranges like "8 x 6 mm" should use the larger dimension
; -----------------------------------------
GetLargerSize(size1, size2, units) {
    if (units = "cm" || units = "centimeter" || units = "centimeters")
        size1 := size1 * 10
    if (size2 != "" && size2 > size1)
        return size2
    return size1
}

; -----------------------------------------
; HELPER: Classify nodule by type and store in appropriate array
; WHY: Fleischner guidelines differ by nodule type (solid vs subsolid vs part-solid)
; -----------------------------------------
ClassifyAndStore(size, units, nodeType, &solidArr, &subsolidArr, &partsolidArr) {
    ; Convert to mm
    if (units = "cm" || units = "centimeter" || units = "centimeters")
        size := size * 10

    ; Only accept plausible nodule sizes (1-50mm)
    if (size < 1 || size > 50)
        return false

    ; Classify based on type keywords (case insensitive)
    nodeTypeLower := StrLower(nodeType)
    if (nodeType = "") {
        ; No type specified - default to solid
        solidArr.Push(size)
    } else if (InStr(nodeTypeLower, "part") || InStr(nodeTypeLower, "semi") || nodeTypeLower = "psn") {
        partsolidArr.Push(size)
    } else if (InStr(nodeTypeLower, "ground") || InStr(nodeTypeLower, "gg") || InStr(nodeTypeLower, "subsolid")
            || InStr(nodeTypeLower, "sub-solid") || InStr(nodeTypeLower, "non-solid") || InStr(nodeTypeLower, "nonsolid")
            || InStr(nodeTypeLower, "hazy") || InStr(nodeTypeLower, "ill-defined") || nodeTypeLower = "ssn") {
        subsolidArr.Push(size)
    } else {
        ; "solid", "calcified", "spiculated", etc. = solid
        solidArr.Push(size)
    }
    return true
}

; -----------------------------------------
; SMART FLEISCHNER NODULE PARSER
; WHY: Parse findings text for nodules and generate Fleischner 2017 recommendations.
; ARCHITECTURE: Robust multi-pattern regex for real-world radiology text variations.
; TRADEOFF: More permissive matching may occasionally capture non-nodule measurements,
;           but confirmation dialog prevents incorrect insertions.
; -----------------------------------------
ParseAndInsertFleischner(input) {
    global ShowCitations, DataminingPhrase, IncludeDatamining, FleischnerInsertAfterImpression
    input := RegExReplace(input, "`r?\n", " ")
    filtered := FilterNonMeasurements(input)

    ; Arrays to store found nodules
    solidNodules := []
    subsolidNodules := []
    partsolidNodules := []
    confidence := 0

    ; === COMPREHENSIVE TYPE KEYWORDS ===
    ; WHY: Covers all radiology terminology for nodule morphology
    typePattern := "solid|semi[- ]?solid|part[- ]?solid|ground[- ]?glass|sub[- ]?solid|pure[- ]?GGN|GGN|GGO|GG|SSN|PSN"
    typePattern .= "|calcified|partially calcified|non[- ]?calcified|spiculated|lobulated|irregular|smooth"
    typePattern .= "|cavitary|cystic|necrotic|mixed|indeterminate|non[- ]?solid|hazy|ill[- ]?defined"

    ; === COMPREHENSIVE NODULE KEYWORDS ===
    ; WHY: Matches all ways radiologists describe pulmonary nodules
    nodulePattern := "nodule|nodules|nodular|micronodule|micronodules"
    nodulePattern .= "|opacity|opacities|density|densities"
    nodulePattern .= "|lesion|lesions|mass|masses"
    nodulePattern .= "|focus|foci|finding|findings"
    nodulePattern .= "|abnormality|abnormalities"

    ; === ENHANCED SIZE PATTERNS ===
    ; Handles: "8mm", "8 mm", "8-mm", "0.8 cm", "~8mm", "approximately 8 mm"
    ; Handles: "8 x 6 mm" (captures larger), "8 to 10 mm" (captures larger)
    ; Handles: "up to 8 mm", "at least 6 mm", "measuring 8 mm", "measures 8mm"
    sizePrefix := "(?:~|approximately |approx\.? |about |up to |at least |measuring |measures |sized? )?"
    sizeNum := "(\d+(?:\.\d+)?)"
    sizeRange := "(?:\s*(?:x|-|to)\s*(\d+(?:\.\d+)?))?"  ; Optional second dimension/range
    sizeUnits := "\s*[- ]?(mm|cm|millimeter|centimeter)s?"
    sizePattern := sizePrefix . sizeNum . sizeRange . sizeUnits

    ; === LOCATION KEYWORDS (to filter out non-pulmonary) ===
    ; NOTE: We allow these but don't require them - helps with context
    lungLocations := "lung|pulmonary|lobe|lobar|RUL|RML|RLL|LUL|LLL|lingula|apical|basilar|peripheral|central|subpleural|perifissural"

    ; === PATTERN 1: Size near nodule keyword (within 40 chars) ===
    ; Examples: "8 mm nodule", "8mm pulmonary nodule", "nodule measuring 8 mm"
    pattern1 := "i)" . sizePattern . ".{0,40}(?:" . nodulePattern . ")"
    pattern1b := "i)(?:" . nodulePattern . ").{0,40}" . sizePattern

    ; === PATTERN 2: Type + size + nodule (flexible ordering) ===
    ; Examples: "solid 8mm pulmonary nodule", "8 mm solid nodule", "groundglass nodule 7mm"
    pattern2 := "i)(" . typePattern . ").{0,25}" . sizePattern . ".{0,40}(?:" . nodulePattern . ")"
    pattern2b := "i)(" . typePattern . ").{0,15}(?:" . nodulePattern . ").{0,40}" . sizePattern
    pattern2c := "i)" . sizePattern . ".{0,15}(" . typePattern . ").{0,25}(?:" . nodulePattern . ")"

    ; === PATTERN 3: "largest" or "dominant" nodule ===
    ; Examples: "largest measuring 8 mm", "dominant nodule is 12mm", "the largest is 8 mm"
    pattern3 := "i)(?:largest|dominant|biggest|most prominent|primary).{0,30}" . sizePattern

    ; === PATTERN 4: Sentence structure with "is/are" ===
    ; Examples: "nodule is 8 mm", "nodules are up to 10 mm", "which is 8mm"
    pattern4 := "i)(?:" . nodulePattern . ").{0,20}(?:is|are|was|measures?|measuring).{0,10}" . sizePattern

    ; === PATTERN 5: Parenthetical measurements ===
    ; Examples: "pulmonary nodule (8 mm)", "RUL nodule (0.8 cm)"
    pattern5 := "i)(?:" . nodulePattern . ").{0,15}\(" . sizePattern . "\)"

    ; Search Pattern 1: Size then nodule
    pos := 1
    while (pos := RegExMatch(filtered, pattern1, &m, pos)) {
        size := m[1] + 0
        size2 := m[2] != "" ? m[2] + 0 : 0
        units := m[3]
        finalSize := (size2 > size) ? size2 : size
        ClassifyAndStore(finalSize, units, "", solidNodules, subsolidNodules, partsolidNodules)
        pos += StrLen(m[0])
    }

    ; Search Pattern 1b: Nodule then size
    pos := 1
    while (pos := RegExMatch(filtered, pattern1b, &m, pos)) {
        size := m[1] + 0
        size2 := m[2] != "" ? m[2] + 0 : 0
        units := m[3]
        finalSize := (size2 > size) ? size2 : size
        ClassifyAndStore(finalSize, units, "", solidNodules, subsolidNodules, partsolidNodules)
        pos += StrLen(m[0])
    }

    ; Search Pattern 2: Type + size + nodule
    pos := 1
    while (pos := RegExMatch(filtered, pattern2, &m, pos)) {
        nodeType := m[1]
        size := m[2] + 0
        size2 := m[3] != "" ? m[3] + 0 : 0
        units := m[4]
        finalSize := (size2 > size) ? size2 : size
        ClassifyAndStore(finalSize, units, nodeType, solidNodules, subsolidNodules, partsolidNodules)
        pos += StrLen(m[0])
    }

    ; Search Pattern 2b: Type + nodule + size
    pos := 1
    while (pos := RegExMatch(filtered, pattern2b, &m, pos)) {
        nodeType := m[1]
        size := m[2] + 0
        size2 := m[3] != "" ? m[3] + 0 : 0
        units := m[4]
        finalSize := (size2 > size) ? size2 : size
        ClassifyAndStore(finalSize, units, nodeType, solidNodules, subsolidNodules, partsolidNodules)
        pos += StrLen(m[0])
    }

    ; Search Pattern 2c: Size + type + nodule
    pos := 1
    while (pos := RegExMatch(filtered, pattern2c, &m, pos)) {
        size := m[1] + 0
        size2 := m[2] != "" ? m[2] + 0 : 0
        units := m[3]
        nodeType := m[4]
        finalSize := (size2 > size) ? size2 : size
        ClassifyAndStore(finalSize, units, nodeType, solidNodules, subsolidNodules, partsolidNodules)
        pos += StrLen(m[0])
    }

    ; Search Pattern 3: "largest/dominant" nodule
    pos := 1
    while (pos := RegExMatch(filtered, pattern3, &m, pos)) {
        size := m[1] + 0
        size2 := m[2] != "" ? m[2] + 0 : 0
        units := m[3]
        finalSize := (size2 > size) ? size2 : size
        ClassifyAndStore(finalSize, units, "", solidNodules, subsolidNodules, partsolidNodules)
        pos += StrLen(m[0])
    }

    ; Search Pattern 4: "nodule is/measures X mm"
    pos := 1
    while (pos := RegExMatch(filtered, pattern4, &m, pos)) {
        size := m[1] + 0
        size2 := m[2] != "" ? m[2] + 0 : 0
        units := m[3]
        finalSize := (size2 > size) ? size2 : size
        ClassifyAndStore(finalSize, units, "", solidNodules, subsolidNodules, partsolidNodules)
        pos += StrLen(m[0])
    }

    ; Search Pattern 5: Parenthetical "nodule (8 mm)"
    pos := 1
    while (pos := RegExMatch(filtered, pattern5, &m, pos)) {
        size := m[1] + 0
        size2 := m[2] != "" ? m[2] + 0 : 0
        units := m[3]
        finalSize := (size2 > size) ? size2 : size
        ClassifyAndStore(finalSize, units, "", solidNodules, subsolidNodules, partsolidNodules)
        pos += StrLen(m[0])
    }

    ; Check for "sub 6 mm nodules", "sub-6mm", "punctate nodules", "<6mm nodules", "scattered tiny nodules"
    if (RegExMatch(filtered, "i)(sub[- ]?6|<\s*6|punctate|miliary|tiny|innumerable|scattered small|multiple small|few small).{0,25}(" . nodulePattern . ")")) {
        if (solidNodules.Length = 0)
            solidNodules.Push(5)
    }

    ; Also check for "nodules less than 6 mm" pattern
    if (RegExMatch(filtered, "i)(" . nodulePattern . ").{0,20}(less than|smaller than|under|<)\s*6\s*(mm|cm)?")) {
        if (solidNodules.Length = 0)
            solidNodules.Push(5)
    }

    ; Deduplicate arrays (same nodule might match multiple patterns)
    solidNodules := DeduplicateSizes(solidNodules)
    subsolidNodules := DeduplicateSizes(subsolidNodules)
    partsolidNodules := DeduplicateSizes(partsolidNodules)

    ; If no nodules found, show error
    if (solidNodules.Length = 0 && subsolidNodules.Length = 0 && partsolidNodules.Length = 0) {
        MsgBox("Could not find nodule descriptions in text.`n`nExpected: Text containing `"nodule`" (or similar) with a size measurement.`n`nExamples:`n- `"8 mm pulmonary nodule`"`n- `"solid nodule measuring 8 mm`"`n- `"groundglass opacity, 7mm`"", "Parse Error", 48)
        return
    }

    ; Calculate confidence based on what was found
    totalNodules := solidNodules.Length + subsolidNodules.Length + partsolidNodules.Length
    if (totalNodules > 0)
        confidence := 85
    if (totalNodules > 3)
        confidence := 75  ; Many nodules = more risk of misparse

    ; Find largest nodules in each category
    maxSolid := 0
    maxSubsolid := 0
    maxPartsolid := 0

    for i, size in solidNodules {
        if (size > maxSolid)
            maxSolid := size
    }
    for i, size in subsolidNodules {
        if (size > maxSubsolid)
            maxSubsolid := size
    }
    for i, size in partsolidNodules {
        if (size > maxPartsolid)
            maxPartsolid := size
    }

    ; Determine if multiple nodules
    isMultiple := (totalNodules > 1) ? true : false

    ; Build summary for confirmation
    summaryText := ""
    if (maxSolid > 0)
        summaryText .= "Solid: " . maxSolid . "mm"
    if (maxSubsolid > 0)
        summaryText .= (summaryText != "" ? ", " : "") . "GGN: " . maxSubsolid . "mm"
    if (maxPartsolid > 0)
        summaryText .= (summaryText != "" ? ", " : "") . "Part-solid: " . maxPartsolid . "mm"
    summaryText .= " (" . (isMultiple ? "Multiple" : "Single") . ")"

    ; Build full recommendation block
    resultText := "`n`n___________________________________________________________`n"
    resultText .= "Incidental Lung Nodule Discussion:`n"
    resultText .= "2017 ACR Fleischner Society expert consensus recommendations on incidental pulmonary nodules.`n"
    resultText .= "Target Demographic: 35+ years without known cancer or immunosuppressive disorder.`n`n"

    ; List detected nodules
    resultText .= "Detected nodules:`n"
    noduleNum := 1

    if (maxSolid > 0) {
        resultText .= noduleNum . ". " . maxSolid . " mm solid nodule"
        if (solidNodules.Length > 1)
            resultText .= " (largest of " . solidNodules.Length . ")"
        resultText .= "`n"
        noduleNum++
    }
    if (maxSubsolid > 0) {
        resultText .= noduleNum . ". " . maxSubsolid . " mm ground glass/subsolid nodule"
        if (subsolidNodules.Length > 1)
            resultText .= " (largest of " . subsolidNodules.Length . ")"
        resultText .= "`n"
        noduleNum++
    }
    if (maxPartsolid > 0) {
        resultText .= noduleNum . ". " . maxPartsolid . " mm part-solid nodule"
        if (partsolidNodules.Length > 1)
            resultText .= " (largest of " . partsolidNodules.Length . ")"
        resultText .= "`n"
        noduleNum++
    }

    resultText .= "`n"

    ; Generate recommendations for LOW RISK
    resultText .= "LOW RISK recommendation:`n"
    resultText .= GenerateFleischnerRec(maxSolid, maxSubsolid, maxPartsolid, isMultiple, "Low risk")
    resultText .= "`n`n"

    ; Generate recommendations for HIGH RISK
    resultText .= "HIGH RISK recommendation:`n"
    resultText .= GenerateFleischnerRec(maxSolid, maxSubsolid, maxPartsolid, isMultiple, "High risk")

    ; Add datamining phrase
    if (IncludeDatamining) {
        resultText .= "`n`nData Mining: " . DataminingPhrase . " . ACR and AMA MIPS #364"
    }

    ; Add additional info
    resultText .= "`n`nNote: The need for followup depends on clinical discussion of patient's comorbid conditions, demographics, and willingness to undergo followup imaging and potential intervention."

    if (ShowCitations)
        resultText .= "`nRef: MacMahon H et al. Radiology 2017;284:228-243"

    resultText .= "`n___________________________________________________________"

    ; Show confirmation dialog
    parsedValues := {}
    parsedValues["Nodules found"] := summaryText
    parsedValues["Total count"] := totalNodules

    action := ShowParseConfirmation("Fleischner Nodules", parsedValues, "Full recommendation block will be inserted", confidence)

    if (action = "insert") {
        ; WHY: User wants Fleischner to appear after IMPRESSION: section
        if (FleischnerInsertAfterImpression)
            InsertAtImpression(resultText)
        else
            InsertAfterSelection(resultText)
    } else if (action = "edit") {
        ; Open Fleischner GUI with largest nodule pre-filled
        dominantSize := maxSolid > maxSubsolid ? maxSolid : maxSubsolid
        dominantSize := dominantSize > maxPartsolid ? dominantSize : maxPartsolid
        dominantType := maxSolid >= dominantSize ? "Solid" : (maxSubsolid >= dominantSize ? "Ground glass" : "Part-solid")
        ShowFleischnerGuiPrefilled(dominantSize, dominantType, isMultiple ? "Multiple" : "Single")
    }
}

; -----------------------------------------
; Generate Fleischner recommendation for specific nodule sizes and risk level
; -----------------------------------------
GenerateFleischnerRec(maxSolid, maxSubsolid, maxPartsolid, isMultiple, risk) {
    result := ""
    number := isMultiple ? "Multiple" : "Single"

    ; Solid nodules take priority
    if (maxSolid > 0) {
        if (maxSolid < 6) {
            if (risk = "Low risk")
                result .= "- Solid " . maxSolid . "mm: No routine follow-up required."
            else
                result .= "- Solid " . maxSolid . "mm: Optional CT at 12 months."
        } else if (maxSolid <= 8) {
            if (isMultiple) {
                if (risk = "Low risk")
                    result .= "- Solid " . maxSolid . "mm (multiple): CT at 3-6 months, then consider CT at 18-24 months."
                else
                    result .= "- Solid " . maxSolid . "mm (multiple): CT at 3-6 months, then CT at 18-24 months."
            } else {
                if (risk = "Low risk")
                    result .= "- Solid " . maxSolid . "mm: CT at 6-12 months, then consider CT at 18-24 months."
                else
                    result .= "- Solid " . maxSolid . "mm: CT at 6-12 months, then CT at 18-24 months."
            }
        } else {
            result .= "- Solid " . maxSolid . "mm: Consider CT at 3 months, PET/CT, or tissue sampling."
        }
    }

    ; Subsolid/Ground glass nodules
    if (maxSubsolid > 0) {
        if (result != "")
            result .= "`n"

        if (maxSubsolid < 6) {
            if (isMultiple)
                result .= "- Ground glass " . maxSubsolid . "mm (multiple): CT at 3-6 months. If stable, consider CT at 2 and 4 years."
            else
                result .= "- Ground glass " . maxSubsolid . "mm: No routine follow-up required."
        } else {
            if (isMultiple)
                result .= "- Ground glass " . maxSubsolid . "mm (multiple): CT at 3-6 months, subsequent management based on most suspicious nodule."
            else
                result .= "- Ground glass " . maxSubsolid . "mm: CT at 6-12 months, then every 2 years for 5 years."
        }
    }

    ; Part-solid nodules
    if (maxPartsolid > 0) {
        if (result != "")
            result .= "`n"

        if (maxPartsolid < 6) {
            result .= "- Part-solid " . maxPartsolid . "mm: No routine follow-up required."
        } else {
            if (isMultiple)
                result .= "- Part-solid " . maxPartsolid . "mm (multiple): CT at 3-6 months. If stable, annual CT for 5 years."
            else
                result .= "- Part-solid " . maxPartsolid . "mm: CT at 3-6 months. If stable, annual CT for 5 years. If solid component >=6mm, consider PET/CT or biopsy."
        }
    }

    return result
}

; =========================================
; CALCULATOR 7: ICH Volume (ABC/2)
; WHY: Calculate intracerebral hemorrhage volume using ABC/2 method
; ARCHITECTURE: Smart parse + GUI fallback pattern
; =========================================

; -----------------------------------------
; ICH Smart Parser
; -----------------------------------------
ParseAndInsertICH(input) {
    input := RegExReplace(input, "`r?\n", " ")

    a := 0
    b := 0
    c := 0
    units := "cm"
    confidence := 0

    ; Pattern A (HIGH confidence 90): "X x Y x Z cm" or "X x Y x Z mm"
    ; Example: "hemorrhage measuring 5.0 x 4.0 x 3.5 cm"
    if (RegExMatch(input, "i)(\d+(?:\.\d+)?)\s*x\s*(\d+(?:\.\d+)?)\s*x\s*(\d+(?:\.\d+)?)\s*(cm|mm)?", &m)) {
        a := m[1] + 0
        b := m[2] + 0
        c := m[3] + 0
        units := (m[4] != "") ? m[4] : "cm"
        confidence := 90
    }
    ; Pattern B (MEDIUM confidence 75): Labeled dimensions
    ; Example: "A 5.0, B 4.0, C 3.5" or "length 5, width 4, height 3"
    else if (RegExMatch(input, "i)(?:A|length|L)\s*[:\s=]?\s*(\d+(?:\.\d+)?)\s*(cm|mm)?.*?(?:B|width|W)\s*[:\s=]?\s*(\d+(?:\.\d+)?)\s*(cm|mm)?.*?(?:C|height|H|depth|D)\s*[:\s=]?\s*(\d+(?:\.\d+)?)\s*(cm|mm)?", &m)) {
        a := m[1] + 0
        units := (m[2] != "") ? m[2] : "cm"
        b := m[3] + 0
        c := m[5] + 0
        confidence := 75
    }
    ; Pattern C (LOW confidence 50): Three numbers in sequence
    else {
        numbers := []
        pos := 1
        while (pos := RegExMatch(input, "(\d+(?:\.\d+)?)\s*(cm|mm)?", &m, pos)) {
            numbers.Push({val: m[1] + 0, unit: m[2]})
            pos += StrLen(m[0])
        }
        if (numbers.Length >= 3) {
            a := numbers[1].val
            b := numbers[2].val
            c := numbers[3].val
            units := (numbers[1].unit != "") ? numbers[1].unit : "cm"
            confidence := 50
        }
    }

    if (a = 0 || b = 0 || c = 0) {
        MsgBox("Could not find hemorrhage dimensions.`n`nExpected formats:`n- `"5.0 x 4.0 x 3.5 cm`"`n- `"hemorrhage measuring X x Y x Z`"", "Parse Error", 48)
        return
    }

    ; Convert mm to cm for volume calculation
    if (units = "mm") {
        a := a / 10
        b := b / 10
        c := c / 10
    }

    ; ABC/2 formula
    volume := (a * b * c) / 2
    volume := Round(volume, 1)

    ; Build result text - sentence style with leading space
    resultText := " ICH volume approximately " . volume . " cc (ABC/2: " . Round(a, 1) . " x " . Round(b, 1) . " x " . Round(c, 1) . " cm)."

    ; Show confirmation dialog
    parsedValues := {}
    parsedValues["A"] := Round(a, 1) . " cm"
    parsedValues["B"] := Round(b, 1) . " cm"
    parsedValues["C"] := Round(c, 1) . " cm"
    parsedValues["Volume"] := volume . " cc"

    action := ShowParseConfirmation("ICH Volume (ABC/2)", parsedValues, resultText, confidence)

    if (action = "insert") {
        InsertAfterSelection(resultText)
    } else if (action = "edit") {
        ShowICHVolumeGuiPrefilled(a, b, c)
    }
}

; -----------------------------------------
; ICH Volume GUI
; -----------------------------------------
global ICHGuiObj := ""

ShowICHVolumeGui() {
    global ICHGuiObj
    GetGuiPosition(&xPos, &yPos)

    ICHGuiObj := Gui("+AlwaysOnTop")
    ICHGuiObj.OnEvent("Close", ICHGuiClose)
    ICHGuiObj.OnEvent("Escape", ICHGuiClose)
    ICHGuiObj.Add("Text", "x10 y10 w280", "ICH Volume Calculator (ABC/2)")
    ICHGuiObj.Add("Text", "x10 y40", "A (longest axis, cm):")
    ICHGuiObj.Add("Edit", "x140 y37 w60 vICHDimA")
    ICHGuiObj.Add("Text", "x10 y70", "B (perpendicular, cm):")
    ICHGuiObj.Add("Edit", "x140 y67 w60 vICHDimB")
    ICHGuiObj.Add("Text", "x10 y100", "C (# of slices x thickness, cm):")
    ICHGuiObj.Add("Edit", "x200 y97 w60 vICHDimC")
    ICHGuiObj.Add("Button", "x30 y135 w80", "Calculate").OnEvent("Click", CalcICH)
    ICHGuiObj.Add("Button", "x130 y135 w80", "Cancel").OnEvent("Click", ICHGuiClose)
    ICHGuiObj.Show("x" xPos " y" yPos " w280 h175")
    ICHGuiObj.Title := "ICH Volume (ABC/2)"
}

ShowICHVolumeGuiPrefilled(a, b, c) {
    global ICHGuiObj
    GetGuiPosition(&xPos, &yPos)

    aVal := Round(a, 1)
    bVal := Round(b, 1)
    cVal := Round(c, 1)

    ICHGuiObj := Gui("+AlwaysOnTop")
    ICHGuiObj.OnEvent("Close", ICHGuiClose)
    ICHGuiObj.OnEvent("Escape", ICHGuiClose)
    ICHGuiObj.Add("Text", "x10 y10 w280", "ICH Volume Calculator (Pre-filled)")
    ICHGuiObj.Add("Text", "x10 y40", "A (longest axis, cm):")
    ICHGuiObj.Add("Edit", "x140 y37 w60 vICHDimA", aVal)
    ICHGuiObj.Add("Text", "x10 y70", "B (perpendicular, cm):")
    ICHGuiObj.Add("Edit", "x140 y67 w60 vICHDimB", bVal)
    ICHGuiObj.Add("Text", "x10 y100", "C (# of slices x thickness, cm):")
    ICHGuiObj.Add("Edit", "x200 y97 w60 vICHDimC", cVal)
    ICHGuiObj.Add("Button", "x30 y135 w80", "Calculate").OnEvent("Click", CalcICH)
    ICHGuiObj.Add("Button", "x130 y135 w80", "Cancel").OnEvent("Click", ICHGuiClose)
    ICHGuiObj.Show("x" xPos " y" yPos " w280 h175")
    ICHGuiObj.Title := "ICH Volume (ABC/2)"
}

ICHGuiClose(*) {
    global ICHGuiObj
    ICHGuiObj.Destroy()
}

CalcICH(*) {
    global ICHGuiObj
    saved := ICHGuiObj.Submit(false)

    if (saved.ICHDimA = "" || saved.ICHDimB = "" || saved.ICHDimC = "") {
        MsgBox("Please enter all three dimensions.", "Error", 16)
        return
    }

    a := saved.ICHDimA + 0
    b := saved.ICHDimB + 0
    c := saved.ICHDimC + 0

    if (a <= 0 || b <= 0 || c <= 0) {
        MsgBox("All dimensions must be greater than 0.", "Error", 16)
        return
    }

    ; ABC/2 formula
    volume := (a * b * c) / 2
    volume := Round(volume, 1)

    ; Sentence style output
    result := " ICH volume approximately " . volume . " cc (ABC/2: " . Round(a, 1) . " x " . Round(b, 1) . " x " . Round(c, 1) . " cm)."

    ICHGuiObj.Destroy()
    ShowResult(result)
}

; =========================================
; CALCULATOR 8: Follow-up Date Calculator
; WHY: Calculate follow-up dates from interval specifications
; ARCHITECTURE: Smart parse + GUI fallback pattern
; =========================================

; -----------------------------------------
; Date Smart Parser
; -----------------------------------------
ParseAndInsertDate(input) {
    input := RegExReplace(input, "`r?\n", " ")

    interval := 0
    intervalUnit := ""
    confidence := 0

    ; Pattern A (HIGH confidence 95): "X months", "X weeks", "X days", "X year(s)"
    if (RegExMatch(input, "i)(\d+)\s*(month|months|mo|m)\b", &m)) {
        interval := m[1] + 0
        intervalUnit := "months"
        confidence := 95
    } else if (RegExMatch(input, "i)(\d+)\s*(week|weeks|wk|w)\b", &m)) {
        interval := m[1] + 0
        intervalUnit := "weeks"
        confidence := 95
    } else if (RegExMatch(input, "i)(\d+)\s*(day|days|d)\b", &m)) {
        interval := m[1] + 0
        intervalUnit := "days"
        confidence := 95
    } else if (RegExMatch(input, "i)(\d+)\s*(year|years|yr|y)\b", &m)) {
        interval := m[1] + 0
        intervalUnit := "years"
        confidence := 95
    }
    ; Pattern B (MEDIUM confidence 80): Abbreviated format "3m", "6w", "45d", "1y"
    else if (RegExMatch(input, "i)\b(\d+)(m|mo)\b", &m)) {
        interval := m[1] + 0
        intervalUnit := "months"
        confidence := 80
    } else if (RegExMatch(input, "i)\b(\d+)(w|wk)\b", &m)) {
        interval := m[1] + 0
        intervalUnit := "weeks"
        confidence := 80
    } else if (RegExMatch(input, "i)\b(\d+)(d)\b", &m)) {
        interval := m[1] + 0
        intervalUnit := "days"
        confidence := 80
    } else if (RegExMatch(input, "i)\b(\d+)(y|yr)\b", &m)) {
        interval := m[1] + 0
        intervalUnit := "years"
        confidence := 80
    }

    if (interval = 0 || intervalUnit = "") {
        MsgBox("Could not find follow-up interval.`n`nExpected formats:`n- `"3 months`", `"6 weeks`", `"1 year`"`n- `"3m`", `"6w`", `"45d`", `"1y`"", "Parse Error", 48)
        return
    }

    ; Calculate future date
    futureDate := CalculateFutureDate(interval, intervalUnit)

    ; Build result text - sentence style with leading space
    resultText := " Recommend follow-up in " . interval . " " . intervalUnit . " (approximately " . futureDate . ")."

    ; Show confirmation dialog
    parsedValues := {}
    parsedValues["Interval"] := interval . " " . intervalUnit
    parsedValues["Date"] := futureDate

    action := ShowParseConfirmation("Follow-up Date", parsedValues, resultText, confidence)

    if (action = "insert") {
        InsertAfterSelection(resultText)
    } else if (action = "edit") {
        ShowDateCalculatorGuiPrefilled(interval, intervalUnit)
    }
}

; -----------------------------------------
; Calculate Future Date
; -----------------------------------------
CalculateFutureDate(interval, unit) {
    ; Get current date
    today := FormatTime(, "yyyyMMdd")

    ; Calculate days to add
    daysToAdd := 0
    if (unit = "days")
        daysToAdd := interval
    else if (unit = "weeks")
        daysToAdd := interval * 7
    else if (unit = "months")
        daysToAdd := interval * 30  ; Approximate
    else if (unit = "years")
        daysToAdd := interval * 365  ; Approximate

    ; Add days to current date
    futureDate := DateAdd(today, daysToAdd, "days")

    ; Format the result date
    futureFormatted := FormatTime(futureDate, "MMMM d, yyyy")
    return futureFormatted
}

; -----------------------------------------
; Date Calculator GUI
; -----------------------------------------
global DateGuiObj := ""

ShowDateCalculatorGui() {
    global DateGuiObj
    GetGuiPosition(&xPos, &yPos)

    DateGuiObj := Gui("+AlwaysOnTop")
    DateGuiObj.OnEvent("Close", DateGuiClose)
    DateGuiObj.OnEvent("Escape", DateGuiClose)
    DateGuiObj.Add("Text", "x10 y10 w250", "Follow-up Date Calculator")
    DateGuiObj.Add("Text", "x10 y45", "Interval:")
    DateGuiObj.Add("Edit", "x80 y42 w50 vDateInterval", "3")
    DateGuiObj.Add("DropDownList", "x140 y42 w100 vDateUnit Choose1", ["Months", "Weeks", "Days", "Years"])
    DateGuiObj.Add("Button", "x30 y85 w80", "Calculate").OnEvent("Click", CalcDate)
    DateGuiObj.Add("Button", "x130 y85 w80", "Cancel").OnEvent("Click", DateGuiClose)
    DateGuiObj.Show("x" xPos " y" yPos " w260 h125")
    DateGuiObj.Title := "Follow-up Date"
}

ShowDateCalculatorGuiPrefilled(interval, unit) {
    global DateGuiObj
    GetGuiPosition(&xPos, &yPos)

    ; Determine dropdown selection
    unitSelect := 1
    if (unit = "weeks")
        unitSelect := 2
    else if (unit = "days")
        unitSelect := 3
    else if (unit = "years")
        unitSelect := 4

    DateGuiObj := Gui("+AlwaysOnTop")
    DateGuiObj.OnEvent("Close", DateGuiClose)
    DateGuiObj.OnEvent("Escape", DateGuiClose)
    DateGuiObj.Add("Text", "x10 y10 w250", "Follow-up Date Calculator (Pre-filled)")
    DateGuiObj.Add("Text", "x10 y45", "Interval:")
    DateGuiObj.Add("Edit", "x80 y42 w50 vDateInterval", interval)
    DateGuiObj.Add("DropDownList", "x140 y42 w100 vDateUnit Choose" unitSelect, ["Months", "Weeks", "Days", "Years"])
    DateGuiObj.Add("Button", "x30 y85 w80", "Calculate").OnEvent("Click", CalcDate)
    DateGuiObj.Add("Button", "x130 y85 w80", "Cancel").OnEvent("Click", DateGuiClose)
    DateGuiObj.Show("x" xPos " y" yPos " w260 h125")
    DateGuiObj.Title := "Follow-up Date"
}

DateGuiClose(*) {
    global DateGuiObj
    DateGuiObj.Destroy()
}

CalcDate(*) {
    global DateGuiObj
    saved := DateGuiObj.Submit(false)

    if (saved.DateInterval = "") {
        MsgBox("Please enter an interval.", "Error", 16)
        return
    }

    interval := saved.DateInterval + 0
    if (interval <= 0) {
        MsgBox("Interval must be greater than 0.", "Error", 16)
        return
    }

    ; Normalize unit name
    unit := "months"
    if (saved.DateUnit = "Weeks")
        unit := "weeks"
    else if (saved.DateUnit = "Days")
        unit := "days"
    else if (saved.DateUnit = "Years")
        unit := "years"

    ; Calculate future date
    futureDate := CalculateFutureDate(interval, unit)

    ; Sentence style output
    result := " Recommend follow-up in " . interval . " " . unit . " (approximately " . futureDate . ")."

    DateGuiObj.Destroy()
    ShowResult(result)
}

; =========================================
; TOOL 6: Sectra History Copy
; =========================================
CopySectraHistory() {
    global SectraWindowTitle, g_SelectedText, TargetApps

    ; Save clipboard
    ClipSaved := ClipboardAll()
    A_Clipboard := ""

    ; Check if text is already selected (from our earlier capture)
    if (g_SelectedText != "") {
        A_Clipboard := g_SelectedText
    } else {
        ; Try to get from Sectra
        if (WinExist(SectraWindowTitle)) {
            WinActivate(SectraWindowTitle)
            if (WinWaitActive(SectraWindowTitle, , 2)) {
                Send("^c")
                ClipWait(1)
            }
        }
    }

    if (A_Clipboard = "") {
        MsgBox("No text found. Select history text in Sectra first.", "Info", 48)
        A_Clipboard := ClipSaved
        return
    }

    historyText := A_Clipboard

    ; Find PowerScribe and paste
    foundPS := false
    for index, app in TargetApps {
        if (InStr(app, "PowerScribe") || InStr(app, "Nuance")) {
            if (WinExist(app)) {
                WinActivate(app)
                if (WinWaitActive(app, , 2)) {
                    foundPS := true
                    break
                }
            }
        }
    }

    if (!foundPS) {
        ; Just activate any target app
        WinActivate("ahk_exe notepad.exe")
        WinWaitActive("ahk_exe notepad.exe", , 1)
    }

    Sleep(100)
    Send("^v")

    ; Restore clipboard
    Sleep(100)
    A_Clipboard := ClipSaved

    ToolTip("History pasted!")
    SetTimer(RemoveToolTip, -1500)
    return
}

RemoveToolTip(*) {
    ToolTip()
}

; =========================================
; Settings GUI
; =========================================
global SettingsGuiObj := ""

ShowSettings() {
    global IncludeDatamining, ShowCitations, DataminingPhrase, DefaultSmartParse
    global SmartParseConfirmation, SmartParseFallbackToGUI
    global DefaultMeasurementUnit, RVLVOutputFormat, FleischnerInsertAfterImpression
    global ShowVolumeTools, ShowRVLVTools, ShowNASCETTools
    global ShowAdrenalTools, ShowFleischnerTools, ShowStenosisTools
    global ShowICHTools, ShowDateCalculator, QuickCallEnabled
    global SettingsGuiObj

    CoordMode("Mouse", "Screen")
    MouseGetPos(&mouseX, &mouseY)
    xPos := mouseX + 10
    yPos := mouseY + 10

    dmChecked := IncludeDatamining ? "Checked" : ""
    citChecked := ShowCitations ? "Checked" : ""
    confirmChecked := SmartParseConfirmation ? "Checked" : ""
    fallbackChecked := SmartParseFallbackToGUI ? "Checked" : ""
    fleischnerImprChecked := FleischnerInsertAfterImpression ? "Checked" : ""

    ; Tool visibility checkboxes
    volChecked := ShowVolumeTools ? "Checked" : ""
    rvlvChecked := ShowRVLVTools ? "Checked" : ""
    nascetChecked := ShowNASCETTools ? "Checked" : ""
    adrenalChecked := ShowAdrenalTools ? "Checked" : ""
    fleischnerChecked := ShowFleischnerTools ? "Checked" : ""
    stenosisChecked := ShowStenosisTools ? "Checked" : ""
    ichChecked := ShowICHTools ? "Checked" : ""
    dateChecked := ShowDateCalculator ? "Checked" : ""
    qcChecked := QuickCallEnabled ? "Checked" : ""

    ; Determine which item to select in dropdown (1-based index)
    smartParseSelect := 1
    if (DefaultSmartParse = "RVLV")
        smartParseSelect := 2
    else if (DefaultSmartParse = "NASCET")
        smartParseSelect := 3
    else if (DefaultSmartParse = "Adrenal")
        smartParseSelect := 4
    else if (DefaultSmartParse = "Fleischner")
        smartParseSelect := 5

    ; Unit selection
    unitSelect := (DefaultMeasurementUnit = "mm") ? 2 : 1

    ; RV/LV format selection
    rvlvSelect := (RVLVOutputFormat = "Inline") ? 2 : 1

    SettingsGuiObj := Gui("+AlwaysOnTop")
    SettingsGuiObj.OnEvent("Close", SettingsGuiClose)
    SettingsGuiObj.OnEvent("Escape", SettingsGuiClose)
    SettingsGuiObj.Add("Text", "x10 y10 w280", "RadAssist v2.3 Settings")
    SettingsGuiObj.Add("Text", "x10 y40", "Default Quick Parse:")
    SettingsGuiObj.Add("DropDownList", "x130 y37 w120 vSetDefaultParse Choose" smartParseSelect, ["Volume", "RVLV", "NASCET", "Adrenal", "Fleischner"])

    ; Tool Visibility section (flat checkbox list per user preference)
    SettingsGuiObj.Add("GroupBox", "x10 y65 w270 h125", "Tool Visibility (uncheck to hide)")
    SettingsGuiObj.Add("Checkbox", "x20 y85 w120 vSetShowVolume " volChecked, "Volume")
    SettingsGuiObj.Add("Checkbox", "x150 y85 w120 vSetShowRVLV " rvlvChecked, "RV/LV Ratio")
    SettingsGuiObj.Add("Checkbox", "x20 y108 w120 vSetShowNASCET " nascetChecked, "NASCET")
    SettingsGuiObj.Add("Checkbox", "x150 y108 w120 vSetShowAdrenal " adrenalChecked, "Adrenal")
    SettingsGuiObj.Add("Checkbox", "x20 y131 w120 vSetShowFleischner " fleischnerChecked, "Fleischner")
    SettingsGuiObj.Add("Checkbox", "x150 y131 w120 vSetShowStenosis " stenosisChecked, "Stenosis")
    SettingsGuiObj.Add("Checkbox", "x20 y154 w120 vSetShowICH " ichChecked, "ICH Volume")
    SettingsGuiObj.Add("Checkbox", "x150 y154 w120 vSetShowDate " dateChecked, "Date Calculator")

    SettingsGuiObj.Add("GroupBox", "x10 y195 w270 h105", "Smart Parse Options")
    SettingsGuiObj.Add("Checkbox", "x20 y215 w250 vSetConfirmation " confirmChecked, "Show confirmation dialog before insert")
    SettingsGuiObj.Add("Checkbox", "x20 y240 w250 vSetFallbackGUI " fallbackChecked, "Fall back to GUI when confidence low")
    SettingsGuiObj.Add("Text", "x20 y268", "Default units (no units in text):")
    SettingsGuiObj.Add("DropDownList", "x190 y265 w70 vSetDefaultUnit Choose" unitSelect, ["cm", "mm"])

    SettingsGuiObj.Add("GroupBox", "x10 y305 w270 h80", "RV/LV & Fleischner Output")
    SettingsGuiObj.Add("Text", "x20 y325", "RV/LV format:")
    SettingsGuiObj.Add("DropDownList", "x100 y322 w80 vSetRVLVFormat Choose" rvlvSelect, ["Macro", "Inline"])
    SettingsGuiObj.Add("Text", "x185 y325 w90", "(Macro = cm)")
    SettingsGuiObj.Add("Checkbox", "x20 y350 w250 vSetFleischnerImpression " fleischnerImprChecked, "Insert Fleischner after IMPRESSION:")

    SettingsGuiObj.Add("GroupBox", "x10 y390 w270 h105", "Output Options")
    SettingsGuiObj.Add("Checkbox", "x20 y410 w250 vSetDatamine " dmChecked, "Include datamining phrase by default")
    SettingsGuiObj.Add("Checkbox", "x20 y435 w250 vSetCitations " citChecked, "Show citations in output")
    SettingsGuiObj.Add("Text", "x20 y460", "Datamining phrase:")
    SettingsGuiObj.Add("Edit", "x110 y457 w160 vSetDMPhrase", DataminingPhrase)

    SettingsGuiObj.Add("GroupBox", "x10 y500 w270 h50", "Quick Call")
    SettingsGuiObj.Add("Checkbox", "x20 y520 w250 vSetQuickCall " qcChecked, "Enable Quick Call panel (Ctrl+Alt+P)")

    SettingsGuiObj.Add("Button", "x70 y560 w80", "Save").OnEvent("Click", SaveSettings)
    SettingsGuiObj.Add("Button", "x160 y560 w80", "Cancel").OnEvent("Click", SettingsGuiClose)
    SettingsGuiObj.Show("x" xPos " y" yPos " w295 h600")
    SettingsGuiObj.Title := "Settings"
}

SettingsGuiClose(*) {
    global SettingsGuiObj
    SettingsGuiObj.Destroy()
}

SaveSettings(*) {
    global IncludeDatamining, ShowCitations, DataminingPhrase, DefaultSmartParse
    global SmartParseConfirmation, SmartParseFallbackToGUI
    global DefaultMeasurementUnit, RVLVOutputFormat, FleischnerInsertAfterImpression
    global ShowVolumeTools, ShowRVLVTools, ShowNASCETTools
    global ShowAdrenalTools, ShowFleischnerTools, ShowStenosisTools
    global ShowICHTools, ShowDateCalculator, QuickCallEnabled
    global PreferencesPath
    global SettingsGuiObj

    saved := SettingsGuiObj.Submit()

    ; Convert checkbox values (variable names must match GUI vVar names)
    IncludeDatamining := saved.SetDatamine
    ShowCitations := saved.SetCitations
    SmartParseConfirmation := saved.SetConfirmation
    SmartParseFallbackToGUI := saved.SetFallbackGUI
    FleischnerInsertAfterImpression := saved.SetFleischnerImpression
    DataminingPhrase := saved.SetDMPhrase
    DefaultSmartParse := saved.SetDefaultParse
    DefaultMeasurementUnit := saved.SetDefaultUnit
    RVLVOutputFormat := saved.SetRVLVFormat

    ; Tool visibility settings
    ShowVolumeTools := saved.SetShowVolume
    ShowRVLVTools := saved.SetShowRVLV
    ShowNASCETTools := saved.SetShowNASCET
    ShowAdrenalTools := saved.SetShowAdrenal
    ShowFleischnerTools := saved.SetShowFleischner
    ShowStenosisTools := saved.SetShowStenosis
    ShowICHTools := saved.SetShowICH
    ShowDateCalculator := saved.SetShowDate

    ; Quick Call - apply immediately (open the panel when newly enabled)
    wasQuickCall := QuickCallEnabled
    QuickCallEnabled := saved.SetQuickCall
    if (QuickCallEnabled && !wasQuickCall)
        QC_Enable(1)
    else if (!QuickCallEnabled && wasQuickCall)
        QC_Disable()

    ; Batch all writes together (reduces file operations)
    writeSuccess := true
    writeSuccess := writeSuccess && IniWriteWithRetry("IncludeDatamining", IncludeDatamining)
    writeSuccess := writeSuccess && IniWriteWithRetry("ShowCitations", ShowCitations)
    writeSuccess := writeSuccess && IniWriteWithRetry("DataminingPhrase", DataminingPhrase)
    writeSuccess := writeSuccess && IniWriteWithRetry("DefaultSmartParse", DefaultSmartParse)
    writeSuccess := writeSuccess && IniWriteWithRetry("SmartParseConfirmation", SmartParseConfirmation)
    writeSuccess := writeSuccess && IniWriteWithRetry("SmartParseFallbackToGUI", SmartParseFallbackToGUI)
    writeSuccess := writeSuccess && IniWriteWithRetry("DefaultMeasurementUnit", DefaultMeasurementUnit)
    writeSuccess := writeSuccess && IniWriteWithRetry("RVLVOutputFormat", RVLVOutputFormat)
    writeSuccess := writeSuccess && IniWriteWithRetry("FleischnerInsertAfterImpression", FleischnerInsertAfterImpression)

    ; Tool visibility preferences
    writeSuccess := writeSuccess && IniWriteWithRetry("ShowVolumeTools", ShowVolumeTools)
    writeSuccess := writeSuccess && IniWriteWithRetry("ShowRVLVTools", ShowRVLVTools)
    writeSuccess := writeSuccess && IniWriteWithRetry("ShowNASCETTools", ShowNASCETTools)
    writeSuccess := writeSuccess && IniWriteWithRetry("ShowAdrenalTools", ShowAdrenalTools)
    writeSuccess := writeSuccess && IniWriteWithRetry("ShowFleischnerTools", ShowFleischnerTools)
    writeSuccess := writeSuccess && IniWriteWithRetry("ShowStenosisTools", ShowStenosisTools)
    writeSuccess := writeSuccess && IniWriteWithRetry("ShowICHTools", ShowICHTools)
    writeSuccess := writeSuccess && IniWriteWithRetry("ShowDateCalculator", ShowDateCalculator)
    writeSuccess := writeSuccess && IniWriteWithRetry("QuickCallEnabled", QuickCallEnabled)

    if (!writeSuccess)
        MsgBox("Some settings may not have saved. Check file permissions or OneDrive sync status.", "Warning", 48)

    ToolTip("Settings saved")
    SetTimer(RemoveToolTip, -1500)
}

; -----------------------------------------
; INI Write with retry for OneDrive sync conflicts
; WHY: OneDrive may lock files during sync, retry helps
; -----------------------------------------
IniWriteWithRetry(key, value, maxRetries := 3) {
    global PreferencesPath
    Loop maxRetries
    {
        try {
            IniWrite(value, PreferencesPath, "Settings", key)
            return true
        } catch {
            Sleep(100)  ; Wait 100ms before retry
        }
    }
    return false
}

; =========================================
; Result Display
; =========================================
global ResultGuiObj := ""

ShowResult(text) {
    global ResultGuiObj
    ; Copy to clipboard and show message
    A_Clipboard := text

    ; Create result window
    CoordMode("Mouse", "Screen")
    MouseGetPos(&mouseX, &mouseY)
    xPos := mouseX + 10
    yPos := mouseY + 10

    ResultGuiObj := Gui("+AlwaysOnTop")
    ResultGuiObj.OnEvent("Close", ResultGuiClose)
    ResultGuiObj.OnEvent("Escape", ResultGuiClose)
    ResultGuiObj.Add("Edit", "x10 y10 w380 h200 ReadOnly", text)
    ResultGuiObj.Add("Button", "x10 y220 w120", "Insert into Report").OnEvent("Click", InsertResult)
    ResultGuiObj.Add("Button", "x140 y220 w120", "Copy to Clipboard").OnEvent("Click", CopyResult)
    ResultGuiObj.Add("Button", "x270 y220 w120", "Close").OnEvent("Click", ResultGuiClose)
    ResultGuiObj.Show("x" xPos " y" yPos " w400 h260")
    ResultGuiObj.Title := "Result"
}

ResultGuiClose(*) {
    global ResultGuiObj
    ResultGuiObj.Destroy()
}

CopyResult(*) {
    global ResultGuiObj
    ; Already in clipboard from ShowResult
    ToolTip("Copied to clipboard!")
    SetTimer(RemoveToolTip, -1500)
    ResultGuiObj.Destroy()
}

InsertResult(*) {
    global ResultGuiObj
    ResultGuiObj.Destroy()
    Sleep(100)
    ; WHY: Deselect any selected text first to prevent replacing it
    ; NOTE: Right arrow moves cursor to end of selection without deleting
    Send("{Right}")
    Sleep(50)
    Send("^v")
}

; =========================================
; Initialize Preferences Path (OneDrive Compatibility)
; WHY: Determines writable location for INI file
; TRADEOFF: Falls back to LOCALAPPDATA if script dir is read-only
; =========================================
InitPreferencesPath() {
    global PreferencesPath
    ; Try script directory first
    testPath := A_ScriptDir . "\RadAssist_preferences.ini"
    testFile := A_ScriptDir . "\.radassist_write_test"

    ; Test if we can write to script directory
    canWrite := true
    try {
        FileAppend("test", testFile)
        FileDelete(testFile)
    } catch {
        canWrite := false
    }

    if (canWrite) {
        PreferencesPath := testPath
    } else {
        ; Fall back to LOCALAPPDATA
        fallbackDir := A_AppData . "\..\Local\RadAssist"
        if (!FileExist(fallbackDir)) {
            DirCreate(fallbackDir)
        }
        PreferencesPath := fallbackDir . "\RadAssist_preferences.ini"
    }
}

; =========================================
; Load Preferences on Startup
; =========================================
LoadPreferences() {
    global IncludeDatamining, ShowCitations, DataminingPhrase, DefaultSmartParse
    global SmartParseConfirmation, SmartParseFallbackToGUI
    global DefaultMeasurementUnit, RVLVOutputFormat, FleischnerInsertAfterImpression
    global ShowVolumeTools, ShowRVLVTools, ShowNASCETTools
    global ShowAdrenalTools, ShowFleischnerTools, ShowStenosisTools
    global ShowICHTools, ShowDateCalculator, QuickCallEnabled
    global PreferencesPath

    ; Use PreferencesPath set by InitPreferencesPath() for OneDrive compatibility
    if (FileExist(PreferencesPath)) {
        IncludeDatamining := IniRead(PreferencesPath, "Settings", "IncludeDatamining", "1")
        ShowCitations := IniRead(PreferencesPath, "Settings", "ShowCitations", "1")
        DataminingPhrase := IniRead(PreferencesPath, "Settings", "DataminingPhrase", "SSM lung nodule")
        DefaultSmartParse := IniRead(PreferencesPath, "Settings", "DefaultSmartParse", "Volume")
        SmartParseConfirmation := IniRead(PreferencesPath, "Settings", "SmartParseConfirmation", "1")
        SmartParseFallbackToGUI := IniRead(PreferencesPath, "Settings", "SmartParseFallbackToGUI", "1")
        DefaultMeasurementUnit := IniRead(PreferencesPath, "Settings", "DefaultMeasurementUnit", "cm")
        RVLVOutputFormat := IniRead(PreferencesPath, "Settings", "RVLVOutputFormat", "Macro")
        FleischnerInsertAfterImpression := IniRead(PreferencesPath, "Settings", "FleischnerInsertAfterImpression", "1")

        ; Tool visibility preferences (default to true/1)
        ShowVolumeTools := IniRead(PreferencesPath, "Settings", "ShowVolumeTools", "1")
        ShowRVLVTools := IniRead(PreferencesPath, "Settings", "ShowRVLVTools", "1")
        ShowNASCETTools := IniRead(PreferencesPath, "Settings", "ShowNASCETTools", "1")
        ShowAdrenalTools := IniRead(PreferencesPath, "Settings", "ShowAdrenalTools", "1")
        ShowFleischnerTools := IniRead(PreferencesPath, "Settings", "ShowFleischnerTools", "1")
        ShowStenosisTools := IniRead(PreferencesPath, "Settings", "ShowStenosisTools", "1")
        ShowICHTools := IniRead(PreferencesPath, "Settings", "ShowICHTools", "1")
        ShowDateCalculator := IniRead(PreferencesPath, "Settings", "ShowDateCalculator", "1")
        QuickCallEnabled := IniRead(PreferencesPath, "Settings", "QuickCallEnabled", "1")

        IncludeDatamining := (IncludeDatamining = "1")
        ShowCitations := (ShowCitations = "1")
        SmartParseConfirmation := (SmartParseConfirmation = "1")
        SmartParseFallbackToGUI := (SmartParseFallbackToGUI = "1")
        FleischnerInsertAfterImpression := (FleischnerInsertAfterImpression = "1")

        ; Convert visibility settings to boolean
        ShowVolumeTools := (ShowVolumeTools = "1")
        ShowRVLVTools := (ShowRVLVTools = "1")
        ShowNASCETTools := (ShowNASCETTools = "1")
        ShowAdrenalTools := (ShowAdrenalTools = "1")
        ShowFleischnerTools := (ShowFleischnerTools = "1")
        ShowStenosisTools := (ShowStenosisTools = "1")
        ShowICHTools := (ShowICHTools = "1")
        ShowDateCalculator := (ShowDateCalculator = "1")
        QuickCallEnabled := (QuickCallEnabled = "1")
    }

    ; Validate loaded preferences
    validParsers := "Volume,RVLV,NASCET,Adrenal,Fleischner"
    if (!InStr(validParsers, DefaultSmartParse))
        DefaultSmartParse := "Volume"

    validUnits := "cm,mm"
    if (!InStr(validUnits, DefaultMeasurementUnit))
        DefaultMeasurementUnit := "cm"

    validFormats := "Inline,Macro"
    if (!InStr(validFormats, RVLVOutputFormat))
        RVLVOutputFormat := "Macro"
}
LoadPreferences()

; =========================================
; Tray Menu
; =========================================
A_TrayMenu.Delete()
A_TrayMenu.Add("RadAssist v2.4", TrayAbout)
A_TrayMenu.Add()
A_TrayMenu.Add("Settings", MenuSettings)
A_TrayMenu.Add("Reload", TrayReload)
A_TrayMenu.Add("Exit", TrayExit)
A_IconTip := "RadAssist v2.4 - Smart Radiology Tools"

; =========================================
; Quick Call module (floating always-on-top call panel)
; NOTE: Embedded here (same code as QuickCall.ahk) so RadAssist is a single file.
;       Must come after the tray menu and before QC_Enable(). Toggle in Settings > Quick Call.
; =========================================
; ================= SETTINGS =================
global QC_PopupHotkey    := "^!p"   ; Ctrl+Alt+P  show / minimize panel
global QC_DialSelHotkey  := "^!d"   ; Ctrl+Alt+D  dial highlighted number
global QC_LogHotkey      := "^!l"   ; Ctrl+Alt+L  paste critical-result statement
global QC_OpenAtStart    := true    ; open the panel when the script starts
global QC_DialMode       := "url"   ; "url" = voice.google.com link | "tel" = tel: link
global QC_GVAccount      := 0       ; Google account index if signed into several
global QC_RadName        := "Dr. Harwood"
global QC_IncludeReadBack := true

global QC_BookFile := A_ScriptDir "\QuickCall_PhoneBook.txt"
global QC_IniFile  := A_ScriptDir "\QuickCall_Settings.ini"

; Used only to create QuickCall_PhoneBook.txt the first time
global QC_DefaultBook := "
(
; QuickCall phone book. One contact per line: Group|Label|Number
; Group Main = big buttons at the top. Any other group = a tech column.
; Edit here or with the Edit #s button on the panel.
Main|HC ER|636-947-5111
Main|LSL ER|636-625-5300
Main|HC ICU|
Main|LSL ICU|
HC|ED Charge RN|636-947-7503
HC|Stroke Unit|636-443-7550
HC|House Sup|636-443-7596
HC|CT ED|636-947-5285
HC|CT Street|636-947-5410
HC|MRI|636-947-5606
HC|US|636-947-5441
HC|Rad / XR (ED)|636-443-2855
HC|Nuc Med|636-947-5494
HC|Vascular|636-947-5438
HC|IR RN|
HC|Front Desk (opt 1)|636-947-5444
LSL|Stroke Unit|636-755-3617
LSL|House Sup|636-625-7889
LSL|CT|636-755-3635
LSL|MRI|636-625-5407
LSL|Rad / XR|636-625-7801
LSL|US|636-755-3930
LSL|Nuc Med|636-755-3112
LSL|Vascular|636-625-5275
LSL|IR RN|
LSL|Front Desk (opt 2)|636-625-5303
WZ|ER|636-327-1100
WZ|ER MD|636-327-1102
WZ|CT/Rad|636-327-1140
WZ|CT|636-327-1088
WZ|Rad|636-327-1145
)"

; ================= STATE =================
global QC_Entries := [], QC_Gui := "", QC_BtnMap := Map(), QC_Editor := ""
global QC_Expanded := false, QC_LastCall := "", QC_DialogOpen := 0
global QC_X := "", QC_Y := "", QC_LastExtWin := 0
global QC_Pid := DllCall("GetCurrentProcessId")
global QC_Initialized := false, QC_Enabled := false
global QC_TickMs := 500     ; how often the panel checks it is still on top

; ================= ENABLE / DISABLE =================
QC_Init() {
    global QC_Initialized, QC_X, QC_Y, QC_Expanded
    if QC_Initialized
        return
    QC_Initialized := true
    QC_LoadBook()
    try QC_X := IniRead(QC_IniFile, "Panel", "X", "")
    try QC_Y := IniRead(QC_IniFile, "Panel", "Y", "")
    try QC_Expanded := (IniRead(QC_IniFile, "Panel", "Expanded", "0") = "1")

    Hotkey(QC_PopupHotkey, QC_TogglePanel, "Off")
    Hotkey(QC_DialSelHotkey, QC_DialSelection, "Off")
    Hotkey(QC_LogHotkey, QC_LogCall, "Off")
    HotIf(QC_PanelActive)
    Loop 9
        Hotkey(String(A_Index), QC_MainKey, "Off")
    HotIf()
    OnExit(QC_OnExit)
}

; showPanel: -1 = follow QC_OpenAtStart, 0 = don't open, 1 = open now
QC_Enable(showPanel := -1) {
    global QC_Enabled
    QC_Init()
    if !QC_Enabled {
        QC_Enabled := true
        QC_SetHotkeys("On")
        QC_SetTray(true)
        SetTimer(QC_Tick, QC_TickMs)
    }
    if (showPanel = 1 || (showPanel = -1 && QC_OpenAtStart))
        QC_ShowPanel()
}

QC_Disable() {
    global QC_Enabled, QC_Gui
    if !QC_Enabled
        return
    QC_Enabled := false
    SetTimer(QC_Tick, 0)
    QC_SetHotkeys("Off")
    QC_SetTray(false)
    if IsObject(QC_Editor)
        try QC_Editor.Destroy()
    QC_SetEditor("")
    if IsObject(QC_Gui) {
        QC_SavePos()
        try QC_Gui.Destroy()
        QC_Gui := ""
    }
}

QC_SetHotkeys(state) {
    Hotkey(QC_PopupHotkey, state)
    Hotkey(QC_DialSelHotkey, state)
    Hotkey(QC_LogHotkey, state)
    HotIf(QC_PanelActive)
    Loop 9
        Hotkey(String(A_Index), state)
    HotIf()
}

QC_SetTray(add) {
    for name in ["Quick Call panel", "Edit phone numbers"] {
        try A_TrayMenu.Delete(name)
    }
    if !add
        return
    try {
        ; RadAssist tray: put the items just above Reload
        A_TrayMenu.Insert("Reload", "Quick Call panel", QC_ShowPanel)
        A_TrayMenu.Insert("Reload", "Edit phone numbers", QC_ShowEditor)
    } catch {
        A_TrayMenu.Add("Quick Call panel", QC_ShowPanel)
        A_TrayMenu.Add("Edit phone numbers", QC_ShowEditor)
    }
}

; ================= RIGHT-CLICK MENU HOOK =================
QC_AddPhoneMenu(parentMenu) {
    pm := Menu()
    pm.Add("Show call panel", QC_ShowPanel)
    pm.Add()
    for e in QC_Entries {
        if (e.sec = "Main")
            pm.Add("Call " e.label (e.num = "" ? "  (add #)" : ""), QC_DialEntry.Bind(e))
    }
    pm.Add()
    pm.Add("Dial highlighted number", QC_DialSelection)
    pm.Add("Log last call to report", QC_LogCall)
    pm.Add("Edit phone numbers...", QC_ShowEditor)
    parentMenu.Add("Phone", pm)
}

; ================= PHONE BOOK FILE =================
QC_LoadBook() {
    global QC_Entries
    if !FileExist(QC_BookFile)
        FileAppend(QC_DefaultBook "`n", QC_BookFile, "UTF-8")
    QC_Entries := []
    for line in StrSplit(FileRead(QC_BookFile, "UTF-8"), "`n", "`r") {
        line := Trim(line)
        if (line = "" || SubStr(line, 1, 1) = ";")
            continue
        p := StrSplit(line, "|")
        if (p.Length < 2)
            continue
        QC_Entries.Push({sec: Trim(p[1]), label: Trim(p[2]), num: p.Length >= 3 ? Trim(p[3]) : ""})
    }
}

QC_SaveBook() {
    txt := "; QuickCall phone book. One contact per line: Group|Label|Number`n"
    txt .= "; Group Main = big buttons at the top. Any other group = a tech column.`n"
    for e in QC_Entries
        txt .= e.sec "|" e.label "|" e.num "`n"
    try FileDelete(QC_BookFile)
    FileAppend(txt, QC_BookFile, "UTF-8")
}

QC_SetEntries(arr) {
    global QC_Entries
    QC_Entries := arr
    QC_SaveBook()
}

QC_Groups() {
    main := [], secs := [], grp := Map()
    for e in QC_Entries {
        if (e.sec = "Main") {
            main.Push(e)
            continue
        }
        if !grp.Has(e.sec) {
            grp[e.sec] := []
            secs.Push(e.sec)
        }
        grp[e.sec].Push(e)
    }
    return {main: main, secs: secs, grp: grp}
}

QC_Full(e) => (e.sec = "Main" ? e.label : e.sec " " e.label)

QC_IsValidNum(v) {
    d := RegExReplace(v, "\D")
    return StrLen(d) = 10 || (StrLen(d) = 11 && SubStr(d, 1, 1) = "1")
}

; ================= PANEL =================
QC_BuildPanel() {
    global QC_Gui, QC_BtnMap
    if IsObject(QC_Gui) {
        QC_SavePos()
        try QC_Gui.Destroy()
        QC_Gui := ""
    }
    QC_BtnMap := Map()
    grps := QC_Groups()

    ; +ToolWindow = floating panel with no taskbar button; it is hidden (not
    ; minimized) by Esc / X and comes back with Ctrl+Alt+P, the menu, or tray.
    g := Gui("+AlwaysOnTop +ToolWindow -MaximizeBox -MinimizeBox", "Quick Call")
    g.BackColor := "202020"
    g.MarginX := 10, g.MarginY := 10
    g.SetFont("s11", "Segoe UI")
    g.OnEvent("Close", QC_OnPanelClose)
    g.OnEvent("Escape", (*) => (QC_SavePos(), QC_Gui.Hide()))
    g.OnEvent("ContextMenu", QC_PanelContext)

    BtnW := 150, BtnH := 42, Gap := 8
    for i, e in grps.main {
        if (i = 1)
            opt := "xm ym"
        else if (Mod(i, 2) = 1)
            opt := "xm y+" Gap
        else
            opt := "x+" Gap " yp"
        b := g.AddButton(opt " w" BtnW " h" BtnH, (i <= 9 ? i "   " : "") e.label (e.num = "" ? " (add #)" : ""))
        b.OnEvent("Click", QC_DialEntry.Bind(e))
        QC_BtnMap[b.Hwnd] := e
    }

    fullW := BtnW * 2 + Gap
    g.SetFont("s10")
    tb := g.AddButton((grps.main.Length ? "xm y+12" : "xm ym") " w" fullW " h30", QC_Expanded ? "−   Techs" : "+   Techs")
    tb.OnEvent("Click", QC_ToggleTechs)
    tb.GetPos(, &ty, , &th)
    bottom := ty + th

    if QC_Expanded {
        ColW := 150, RowH := 32
        for k, sec in grps.secs {
            cx := 10 + (k - 1) * (ColW + Gap)
            cy := ty + th + 10
            g.SetFont("s10 bold c9FC5E8")
            g.AddText("x" cx " y" cy " w" ColW, sec)
            g.SetFont("s9 norm cDefault")
            cy += 22
            for e in grps.grp[sec] {
                b := g.AddButton("x" cx " y" cy " w" ColW " h28", e.label (e.num = "" ? " (add #)" : ""))
                b.OnEvent("Click", QC_DialEntry.Bind(e))
                QC_BtnMap[b.Hwnd] := e
                cy += RowH
            }
            if (cy > bottom)
                bottom := cy
        }
    }

    g.SetFont("s9 norm cDefault")
    w3 := Floor((fullW - 2 * Gap) / 3)
    b1 := g.AddButton("x10 y" (bottom + 10) " w" w3 " h28", "Dial #...")
    b1.OnEvent("Click", QC_ManualDial)
    b2 := g.AddButton("x+" Gap " yp w" w3 " h28", "Log call")
    b2.OnEvent("Click", QC_LogCall)
    b3 := g.AddButton("x+" Gap " yp w" w3 " h28", "Edit #s")
    b3.OnEvent("Click", QC_ShowEditor)

    QC_Gui := g
    g.Show(QC_ValidPos(QC_X, QC_Y) ? "x" QC_X " y" QC_Y : "")
    QC_RaiseTop(g.Hwnd)
}

QC_ShowPanel(*) {
    if !IsObject(QC_Gui) {
        QC_BuildPanel()
        return
    }
    h := QC_Gui.Hwnd
    if (WinGetMinMax(h) = -1)
        QC_Gui.Restore()
    else
        QC_Gui.Show()
    QC_RaiseTop(h)
}

QC_TogglePanel(*) {
    if !IsObject(QC_Gui) {
        QC_BuildPanel()
        return
    }
    h := QC_Gui.Hwnd
    if (!DllCall("IsWindowVisible", "Ptr", h) || WinGetMinMax(h) = -1) {
        QC_ShowPanel()
    } else {
        QC_SavePos()
        QC_Gui.Hide()
    }
}

QC_ToggleTechs(*) {
    global QC_Expanded
    QC_Expanded := !QC_Expanded
    try IniWrite(QC_Expanded ? 1 : 0, QC_IniFile, "Panel", "Expanded")
    QC_BuildPanel()
}

QC_OnPanelClose(*) {
    QC_SavePos()          ; window hides; reopen with Ctrl+Alt+P, menu, or tray
}

QC_SavePos() {
    global QC_X, QC_Y
    if !IsObject(QC_Gui)
        return
    try {
        h := QC_Gui.Hwnd
        if !DllCall("IsWindowVisible", "Ptr", h) || WinGetMinMax(h) = -1
            return
        WinGetPos(&x, &y, , , h)
        QC_X := x, QC_Y := y
        IniWrite(x, QC_IniFile, "Panel", "X")
        IniWrite(y, QC_IniFile, "Panel", "Y")
    }
}

QC_ValidPos(x, y) {
    if (x = "" || y = "")
        return false
    Loop MonitorGetCount() {
        MonitorGetWorkArea(A_Index, &l, &t, &r, &b)
        if (x >= l - 50 && x < r - 50 && y >= t - 10 && y < b - 50)
            return true
    }
    return false
}

QC_OnExit(*) {
    QC_SavePos()
}

QC_PanelActive(*) {
    try return IsObject(QC_Gui) && WinActive("ahk_id " QC_Gui.Hwnd)
    return false
}

QC_MainKey(hk) {
    grps := QC_Groups()
    i := Integer(hk)
    if (i <= grps.main.Length)
        QC_DialEntry(grps.main[i])
}

; Keeps the panel above Sectra/PowerScribe and remembers the last non-browser window
QC_Tick() {
    global QC_LastExtWin
    try {
        a := WinExist("A")
        if (a && WinGetPID(a) != QC_Pid) {
            if !RegExMatch(WinGetProcessName(a), "i)^(chrome|msedge|firefox|brave)\.exe$")
                QC_LastExtWin := a
        }
    }
    if (!IsObject(QC_Gui) || QC_DialogOpen > 0)
        return
    try {
        h := QC_Gui.Hwnd
        if !DllCall("IsWindowVisible", "Ptr", h) || WinGetMinMax(h) = -1
            return
        if (WinGetPID("A") = QC_Pid)   ; one of our own windows or menus is active
            return
        ; Re-assert topmost every tick (cheap, no flicker)
        DllCall("SetWindowPos", "Ptr", h, "Ptr", -1, "Int", 0, "Int", 0, "Int", 0, "Int", 0, "UInt", 0x13)
        ; Sectra / PowerScribe windows that are themselves topmost can still sit
        ; above us; if anything covers the panel, bump it to the top of the band.
        if QC_IsObscured(h)
            QC_RaiseTop(h)
    }
}

; Move the panel to the very top of the always-on-top band without activating it.
; Toggling NOTOPMOST -> TOPMOST is what lifts it above other topmost windows.
QC_RaiseTop(h) {
    ; 0x13 = SWP_NOSIZE | SWP_NOMOVE | SWP_NOACTIVATE
    try {
        DllCall("SetWindowPos", "Ptr", h, "Ptr", -2, "Int", 0, "Int", 0, "Int", 0, "Int", 0, "UInt", 0x13)
        DllCall("SetWindowPos", "Ptr", h, "Ptr", -1, "Int", 0, "Int", 0, "Int", 0, "Int", 0, "UInt", 0x13)
    }
}

; True if another application's window is drawn over part of the panel.
; Shell UI (taskbar, Start, Alt+Tab, menus, tooltips) and our own windows are ignored.
QC_IsObscured(h) {
    static shellCls := "i)^(Shell_TrayWnd|Shell_SecondaryTrayWnd|#32768|tooltips_class32|NotifyIconOverflowWindow"
        . "|TopLevelWindowForOverflowXamlIsland|Windows\.UI\.Core\.CoreWindow|ForegroundStaging"
        . "|MultitaskingViewFrame|XamlExplorerHostIslandWindow|TaskListThumbnailWnd|Progman|WorkerW)$"
    try WinGetPos(&x, &y, &w, &hh, h)
    catch
        return false
    pts := [[x + w // 2, y + hh // 2], [x + 12, y + 12], [x + w - 12, y + 12]
          , [x + 12, y + hh - 12], [x + w - 12, y + hh - 12]]
    for p in pts {
        hit := DllCall("WindowFromPoint", "Int64", (p[2] << 32) | (p[1] & 0xFFFFFFFF), "Ptr")
        if !hit
            continue
        root := DllCall("GetAncestor", "Ptr", hit, "UInt", 2, "Ptr")   ; GA_ROOT
        if (!root || root = h)
            continue
        try {
            if (WinGetPID("ahk_id " root) = QC_Pid)
                continue
            if RegExMatch(WinGetClass("ahk_id " root), shellCls)
                continue
        } catch {
            continue
        }
        return true
    }
    return false
}

QC_OwnIfVisible() {
    if !IsObject(QC_Gui)
        return
    try {
        h := QC_Gui.Hwnd
        if DllCall("IsWindowVisible", "Ptr", h) && WinGetMinMax(h) != -1
            QC_Gui.Opt("+OwnDialogs")
    }
}

; Right-click a button on the panel to edit it
QC_PanelContext(GuiObj, Ctrl, Item, IsRightClick, X, Y) {
    if (!IsObject(Ctrl) || !QC_BtnMap.Has(Ctrl.Hwnd))
        return
    e := QC_BtnMap[Ctrl.Hwnd]
    m := Menu()
    m.Add("Edit number for " QC_Full(e) "...", (*) => QC_EditField(e, "num"))
    m.Add("Rename...", (*) => QC_EditField(e, "label"))
    m.Add()
    m.Add("Open phone book editor...", QC_ShowEditor)
    m.Show()
}

QC_EditField(e, field) {
    global QC_DialogOpen
    QC_DialogOpen++
    QC_OwnIfVisible()
    prompt := (field = "num") ? "Phone number for " QC_Full(e) ":" : "New label for " QC_Full(e) ":"
    ib := InputBox(prompt, "Quick Call", "w320 h120", e.%field%)
    QC_DialogOpen--
    if (ib.Result != "OK")
        return
    val := Trim(ib.Value)
    if (field = "num" && val != "" && !QC_IsValidNum(val)) {
        MsgBox("Enter a 10-digit number (e.g. 636-947-5111) or leave it blank.", "Quick Call", "Icon!")
        return
    }
    if (field = "label" && val = "")
        return
    e.%field% := val
    QC_SaveBook()
    QC_BuildPanel()
}

; ================= PHONE BOOK EDITOR =================
QC_SetEditor(g) {
    global QC_Editor, QC_DialogOpen
    if IsObject(g)
        QC_DialogOpen++
    else if IsObject(QC_Editor)
        QC_DialogOpen--
    QC_Editor := g
}

QC_ShowEditor(*) {
    if IsObject(QC_Editor) {
        try {
            QC_Editor.Show()
            return
        }
    }
    work := []
    groups := []
    for e in QC_Entries {
        work.Push({sec: e.sec, label: e.label, num: e.num})
        found := false
        for gname in groups {
            if (gname = e.sec)
                found := true
        }
        if !found
            groups.Push(e.sec)
    }

    eg := Gui("+AlwaysOnTop -MaximizeBox", "Quick Call - Edit Phone Book")
    eg.SetFont("s10", "Segoe UI")
    eg.AddText("xm", "Group Main = big buttons at top. Other groups = tech columns. Type a new group name to add a column.")
    lv := eg.AddListView("xm w560 r16 -Multi Grid", ["Group", "Label", "Number"])
    eg.AddText("xm y+10", "Group")
    cbSec := eg.AddComboBox("x+5 yp-3 w100", groups)
    eg.AddText("x+12 yp+3", "Label")
    edLab := eg.AddEdit("x+5 yp-3 w170")
    eg.AddText("x+12 yp+3", "Number")
    edNum := eg.AddEdit("x+5 yp-3 w130")
    bAdd := eg.AddButton("xm y+12 w90", "Add new")
    bUpd := eg.AddButton("x+6 yp w90", "Update")
    bDel := eg.AddButton("x+6 yp w90", "Delete")
    bUp  := eg.AddButton("x+6 yp w50", "▲")
    bDn  := eg.AddButton("x+6 yp w50", "▼")
    bSave := eg.AddButton("xm y+16 w130", "Save && Close")
    bCan  := eg.AddButton("x+6 yp w90", "Cancel")
    state := {closed: false}

    Refresh(sel := 0) {
        lv.Delete()
        for e in work
            lv.Add(, e.sec, e.label, e.num)
        lv.ModifyCol(1, 90), lv.ModifyCol(2, 260), lv.ModifyCol(3, 180)
        if (sel > 0 && sel <= work.Length)
            lv.Modify(sel, "Select Focus Vis")
    }
    Fill(row) {
        e := work[row]
        cbSec.Text := e.sec
        edLab.Value := e.label
        edNum.Value := e.num
    }
    ReadFields() {
        s := Trim(cbSec.Text), l := Trim(edLab.Value), n := Trim(edNum.Value)
        if (s = "" || l = "") {
            MsgBox("Group and Label are required.", "Quick Call", "Icon! Owner" eg.Hwnd)
            return ""
        }
        if (n != "" && !QC_IsValidNum(n)) {
            MsgBox("Enter a 10-digit number (e.g. 636-947-5111) or leave it blank.", "Quick Call", "Icon! Owner" eg.Hwnd)
            return ""
        }
        return {sec: s, label: l, num: n}
    }
    DoAdd(*) {
        f := ReadFields()
        if !IsObject(f)
            return
        r := lv.GetNext(0)
        if (r = 0) {
            work.Push(f)
            r := work.Length
        } else {
            work.InsertAt(r + 1, f)
            r := r + 1
        }
        Refresh(r)
    }
    DoUpdate(*) {
        r := lv.GetNext(0)
        if (r = 0) {
            MsgBox("Select a row to update.", "Quick Call", "Owner" eg.Hwnd)
            return
        }
        f := ReadFields()
        if !IsObject(f)
            return
        work[r] := f
        Refresh(r)
    }
    DoDelete(*) {
        r := lv.GetNext(0)
        if (r = 0)
            return
        work.RemoveAt(r)
        Refresh(r > work.Length ? work.Length : r)
    }
    DoMove(dir, *) {
        r := lv.GetNext(0)
        t := r + dir
        if (r = 0 || t < 1 || t > work.Length)
            return
        tmp := work[r]
        work[r] := work[t]
        work[t] := tmp
        Refresh(t)
    }
    DoClose(*) {
        if state.closed
            return
        state.closed := true
        QC_SetEditor("")
        try eg.Destroy()
    }
    DoSave(*) {
        QC_SetEntries(work)
        DoClose()
        QC_BuildPanel()
    }

    lv.OnEvent("ItemSelect", (ctrl, item, selected) => (selected ? Fill(item) : 0))
    bAdd.OnEvent("Click", DoAdd)
    bUpd.OnEvent("Click", DoUpdate)
    bDel.OnEvent("Click", DoDelete)
    bUp.OnEvent("Click", DoMove.Bind(-1))
    bDn.OnEvent("Click", DoMove.Bind(1))
    bSave.OnEvent("Click", DoSave)
    bCan.OnEvent("Click", DoClose)
    eg.OnEvent("Close", DoClose)
    eg.OnEvent("Escape", DoClose)

    Refresh()
    QC_SetEditor(eg)
    eg.Show()
}

; ================= DIALING =================
QC_DialEntry(e, *) {
    if (e.num = "") {
        QC_EditField(e, "num")
        return
    }
    QC_Dial(e.num, QC_Full(e))
}

QC_Dial(num, label := "") {
    global QC_LastCall
    d := RegExReplace(num, "\D")
    if (StrLen(d) = 11 && SubStr(d, 1, 1) = "1")
        d := SubStr(d, 2)
    if (StrLen(d) != 10) {
        MsgBox("Can't dial '" num "'. Need a full 10-digit number.", "Quick Call", "Icon! 262144")
        return
    }
    if (QC_DialMode = "tel")
        Run("tel:+1" d)
    else
        Run("https://voice.google.com/u/" QC_GVAccount "/calls?a=nc,%2B1" d)
    QC_LastCall := {label: label != "" ? label : QC_FormatPhone(d), time: A_Now}
    QC_Tip("Calling " QC_LastCall.label "  " QC_FormatPhone(d), 3000)
}

QC_DialText(txt) {
    if !RegExMatch(txt, "(\+?1[\s.-]?)?\(?\d{3}\)?[\s.-]?\d{3}[\s.-]?\d{4}|\b8[057]-?\d{4}\b", &m) {
        QC_Tip("No phone number found")
        return
    }
    d := RegExReplace(m[0], "\D")
    if (StrLen(d) = 6) {
        pre := SubStr(d, 1, 2), ext := SubStr(d, 3)
        if (pre = "85")
            d := "636947" ext
        else if (pre = "80")
            d := "636327" ext
        else {
            MsgBox("87- extensions are split between 636-625 and 636-755. Dial the full number.", "Quick Call", "Icon! 262144")
            return
        }
    }
    QC_Dial(d)
}

QC_DialSelection(*) {
    Sleep(100)
    saved := ClipboardAll()
    A_Clipboard := ""
    Send("^c")
    if !ClipWait(0.5) {
        A_Clipboard := saved
        QC_Tip("Highlight a phone number first")
        return
    }
    txt := A_Clipboard
    A_Clipboard := saved
    QC_DialText(txt)
}

QC_ManualDial(*) {
    global QC_DialogOpen
    QC_DialogOpen++
    QC_OwnIfVisible()
    ib := InputBox("Number to dial (10-digit, or 85-xxxx / 80-xxxx):", "Quick Call", "w320 h120")
    QC_DialogOpen--
    if (ib.Result = "OK" && Trim(ib.Value) != "")
        QC_DialText(ib.Value)
}

QC_FormatPhone(d) => SubStr(d, 1, 3) "-" SubStr(d, 4, 3) "-" SubStr(d, 7)

QC_Tip(msg, ms := 2000) {
    ToolTip(msg, , , 2)
    SetTimer(() => ToolTip(, , , 2), -ms)
}

; ================= CRITICAL RESULT STATEMENT =================
QC_LogCall(*) {
    global QC_DialogOpen
    if !IsObject(QC_LastCall) {
        QC_Tip("No call placed yet this session")
        return
    }
    target := 0
    try {
        a := WinExist("A")
        if (a && WinGetPID(a) != QC_Pid && !RegExMatch(WinGetProcessName(a), "i)^(chrome|msedge|firefox|brave)\.exe$"))
            target := a
    }
    if !target
        target := QC_LastExtWin

    QC_DialogOpen++
    QC_OwnIfVisible()
    ib := InputBox("Spoke with (name / title):", "Critical result: " QC_LastCall.label, "w340 h120")
    QC_DialogOpen--
    if (ib.Result != "OK" || Trim(ib.Value) = "")
        return

    t := FormatTime(QC_LastCall.time, "h:mm tt")
    dt := FormatTime(QC_LastCall.time, "M/d/yyyy")
    s := "These findings were discussed by telephone with " Trim(ib.Value) " (" QC_LastCall.label ") by " QC_RadName " at " t " on " dt "."
    if QC_IncludeReadBack
        s .= " Read-back was performed."

    A_Clipboard := s
    ClipWait(1)
    if (target && WinExist("ahk_id " target)) {
        WinActivate("ahk_id " target)
        if WinWaitActive("ahk_id " target, , 1) {
            Send("^v")
            return
        }
    }
    QC_Tip("Statement copied. Paste it into the report.", 3000)
}

if (QuickCallEnabled)
    QC_Enable()

TrayAbout(*) {
    MsgBox("RadAssist v2.4`n`nShift+Right-click in PowerScribe or Notepad`nto access radiology calculators.`n`nv2.2 Features:`n- Pause/Resume with backtick key (``)`n- RV/LV macro format for PowerScribe`n- Fleischner inserts after IMPRESSION:`n- OneDrive compatible (auto-fallback path)`n- ASCII chars for encoding compatibility`n`nSmart Parsers:`n- Smart Volume: Organ detection + prostate interpretation`n- Smart RV/LV: PE risk (Macro or Inline format)`n- Smart NASCET: Stenosis severity grading`n- Smart Adrenal: Washout percentages`n- Parse Nodules: Fleischner 2017 recommendations`n`nUtilities:`n- Report Header Template`n- Sectra History Copy (Ctrl+Shift+H)", "RadAssist", 64)
}

TrayReload(*) {
    Reload()
}

TrayExit(*) {
    ExitApp()
}
