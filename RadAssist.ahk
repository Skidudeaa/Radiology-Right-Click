; ==========================================
; RadAssist - Radiology Assistant Tool
; Version: 2.2
; Description: Lean radiology workflow tool with calculators
;              Triggered by Shift+Right-click in PowerScribe/Notepad
;              Smart text parsing with confirmation dialogs
; ARCHITECTURE: Context-filtered parsing with confidence scoring
; WHY: v2.1 added smart parse; v2.2 adds OneDrive support, RV/LV macro format
; CHANGES in v2.2:
;   - Pause script with backtick key (`)
;   - Fixed character encoding (<=/>= for ASCII compatibility)
;   - cm default with auto-detection for measurements
;   - RV/LV macro format for PowerScribe "Macro right heart"
;   - Fleischner insert after IMPRESSION: field
;   - OneDrive compatibility (fallback path, retry logic)
; ==========================================

#NoEnv
#SingleInstance, Force
SetWorkingDir, %A_ScriptDir%

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
; Exit Cleanup Handler
; WHY: Clear clipboard on exit to prevent PHI exposure
; -----------------------------------------
OnExit("CleanupOnExit")

CleanupOnExit() {
    Clipboard := ""
    return 0  ; Allow exit to proceed
}

; -----------------------------------------
; Global Hotkey: Ctrl+Shift+H - Sectra History Copy
; (Can be mapped to Contour ShuttlePRO button)
; -----------------------------------------
^+h::
    CopySectraHistory()
return

; -----------------------------------------
; Global Hotkey: Backtick (`) - Pause/Resume Script
; WHY: Quick toggle without opening menu
; -----------------------------------------
`::
    Suspend, Toggle
    if (A_IsSuspended)
        TrayTip, RadAssist, Script PAUSED - Press ` to resume, 1
    else
        TrayTip, RadAssist, Script RESUMED, 1
return

; -----------------------------------------
; Shift+Right-Click Menu Handler
; -----------------------------------------
+RButton::
    ; Check if we're in a target application
    inTargetApp := false
    for index, app in TargetApps {
        if (WinActive(app)) {
            inTargetApp := true
            break
        }
    }

    if (!inTargetApp) {
        ; Not in target app - send normal shift+right-click
        Send, +{RButton}
        return
    }

    ; Store any selected text
    ; WHY: Extended timeout (0.5s → 1s) for slower applications like PowerScribe
    ; TRADEOFF: Slightly slower menu appearance vs better reliability
    global g_SelectedText := ""
    ClipSaved := ClipboardAll
    Clipboard := ""
    Send, ^c
    ClipWait, 1  ; Increased from 0.3 for reliability
    if (!ErrorLevel)
        g_SelectedText := Clipboard
    Clipboard := ClipSaved

    ; Build and show menu
    ; WHY: Conditionally build menu based on tool visibility preferences
    ; ARCHITECTURE: Menu items only added if corresponding Show*Tools is true
    Menu, RadAssistMenu, Add
    Menu, RadAssistMenu, DeleteAll

    ; Smart Parse submenu (text parsing with inline insertion)
    Menu, SmartParseMenu, Add
    Menu, SmartParseMenu, DeleteAll
    smartParseHasItems := false
    if (ShowVolumeTools) {
        Menu, SmartParseMenu, Add, Smart Volume (parse dimensions), MenuSmartVolume
        smartParseHasItems := true
    }
    if (ShowRVLVTools) {
        Menu, SmartParseMenu, Add, Smart RV/LV (parse ratio), MenuSmartRVLV
        smartParseHasItems := true
    }
    if (ShowNASCETTools) {
        Menu, SmartParseMenu, Add, Smart NASCET (parse stenosis), MenuSmartNASCET
        smartParseHasItems := true
    }
    if (ShowAdrenalTools) {
        Menu, SmartParseMenu, Add, Smart Adrenal (parse HU values), MenuSmartAdrenal
        smartParseHasItems := true
    }
    if (ShowFleischnerTools) {
        if (smartParseHasItems)
            Menu, SmartParseMenu, Add
        Menu, SmartParseMenu, Add, Parse Nodules (Fleischner), MenuSmartFleischner
        smartParseHasItems := true
    }

    if (smartParseHasItems) {
        Menu, RadAssistMenu, Add, Smart Parse, :SmartParseMenu
        Menu, RadAssistMenu, Add, Quick Parse (%DefaultSmartParse%), MenuQuickSmartParse
        Menu, RadAssistMenu, Add
    }

    ; GUI calculators - conditionally add based on visibility
    guiHasItems := false
    if (ShowVolumeTools) {
        Menu, RadAssistMenu, Add, Ellipsoid Volume (GUI), MenuEllipsoidVolume
        guiHasItems := true
    }
    if (ShowAdrenalTools) {
        Menu, RadAssistMenu, Add, Adrenal Washout (GUI), MenuAdrenalWashout
        guiHasItems := true
    }
    if (guiHasItems)
        Menu, RadAssistMenu, Add

    stenosisHasItems := false
    if (ShowNASCETTools) {
        Menu, RadAssistMenu, Add, NASCET (Carotid), MenuNASCET
        stenosisHasItems := true
    }
    if (ShowStenosisTools) {
        Menu, RadAssistMenu, Add, Vessel Stenosis (General), MenuStenosis
        stenosisHasItems := true
    }
    if (ShowRVLVTools) {
        Menu, RadAssistMenu, Add, RV/LV Ratio (GUI), MenuRVLV
        stenosisHasItems := true
    }
    if (stenosisHasItems)
        Menu, RadAssistMenu, Add

    if (ShowFleischnerTools) {
        Menu, RadAssistMenu, Add, Fleischner 2017 (GUI), MenuFleischner
        Menu, RadAssistMenu, Add
    }

    ; New calculators section
    newCalcHasItems := false
    if (ShowICHTools) {
        Menu, RadAssistMenu, Add, ICH Volume (ABC/2), MenuICHVolume
        newCalcHasItems := true
    }
    if (ShowDateCalculator) {
        Menu, RadAssistMenu, Add, Follow-up Date Calculator, MenuDateCalc
        newCalcHasItems := true
    }
    if (newCalcHasItems)
        Menu, RadAssistMenu, Add

    Menu, RadAssistMenu, Add, Copy Sectra History (Ctrl+Shift+H), MenuSectraHistory
    Menu, RadAssistMenu, Add
    Menu, RadAssistMenu, Add, Settings, MenuSettings

    Menu, RadAssistMenu, Show
return

; -----------------------------------------
; Utility: Show no-selection error with context
; -----------------------------------------
ShowNoSelectionError(parseType) {
    messages := {Volume: "dimensions", RVLV: "RV/LV measurements", NASCET: "stenosis values", Adrenal: "HU values", Fleischner: "nodule measurements"}
    msg := messages[parseType] ? messages[parseType] : "text"
    MsgBox, 48, No Selection, Please select text containing %msg% first.`n`nHighlight the relevant text and try again.
}

; -----------------------------------------
; Menu Handlers
; -----------------------------------------

; Smart Parse handlers (inline text parsing with insertion)
MenuSmartVolume:
    if (g_SelectedText = "") {
        ShowNoSelectionError("Volume")
        return
    }
    ParseAndInsertVolume(g_SelectedText)
return

MenuSmartRVLV:
    if (g_SelectedText = "") {
        MsgBox, 48, No Selection, Please select text containing RV/LV measurements (e.g., "RV 42mm / LV 35mm")
        return
    }
    ParseAndInsertRVLV(g_SelectedText)
return

MenuSmartNASCET:
    if (g_SelectedText = "") {
        MsgBox, 48, No Selection, Please select text containing stenosis measurements (e.g., "distal 5.2mm, stenosis 2.1mm")
        return
    }
    ParseAndInsertNASCET(g_SelectedText)
return

MenuSmartAdrenal:
    if (g_SelectedText = "") {
        MsgBox, 48, No Selection, Please select text containing HU values (e.g., "pre-contrast 10 HU, enhanced 80 HU, delayed 40 HU")
        return
    }
    ParseAndInsertAdrenalWashout(g_SelectedText)
return

MenuSmartFleischner:
    if (g_SelectedText = "") {
        MsgBox, 48, No Selection, Please select the findings section containing nodule descriptions.
        return
    }
    ParseAndInsertFleischner(g_SelectedText)
return

MenuQuickSmartParse:
    if (g_SelectedText = "") {
        MsgBox, 48, No Selection, Please select text to parse.
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
return

; GUI-based handlers
MenuEllipsoidVolume:
    ShowEllipsoidVolumeGui()
return

MenuAdrenalWashout:
    ShowAdrenalWashoutGui()
return

MenuNASCET:
    ; Try text parsing first, fall back to GUI
    if (g_SelectedText != "") {
        result := ParseNASCET(g_SelectedText)
        if (result != "") {
            ShowResult(result)
            return
        }
    }
    ShowNASCETGui()
return

MenuStenosis:
    ShowStenosisGui()
return

MenuRVLV:
    ShowRVLVGui()
return

MenuFleischner:
    ShowFleischnerGui()
return

MenuICHVolume:
    if (g_SelectedText != "") {
        ParseAndInsertICH(g_SelectedText)
    } else {
        ShowICHVolumeGui()
    }
return

MenuDateCalc:
    if (g_SelectedText != "") {
        ParseAndInsertDate(g_SelectedText)
    } else {
        ShowDateCalculatorGui()
    }
return

MenuSectraHistory:
    CopySectraHistory()
return

MenuSettings:
    ShowSettings()
return

; =========================================
; CALCULATOR 1: Ellipsoid Volume
; =========================================
ShowEllipsoidVolumeGui() {
    ; Position near mouse
    GetGuiPosition(xPos, yPos)

    Gui, EllipsoidGui:New, +AlwaysOnTop
    Gui, EllipsoidGui:Add, Text, x10 y10 w280, Ellipsoid Volume Calculator
    Gui, EllipsoidGui:Add, Text, x10 y35, AP (L):
    Gui, EllipsoidGui:Add, Edit, x80 y32 w50 vEllipDim1
    Gui, EllipsoidGui:Add, Text, x135 y35, x   T (W):
    Gui, EllipsoidGui:Add, Edit, x185 y32 w50 vEllipDim2
    Gui, EllipsoidGui:Add, Text, x10 y60, CC (H):
    Gui, EllipsoidGui:Add, Edit, x80 y57 w50 vEllipDim3
    Gui, EllipsoidGui:Add, Text, x135 y60, Units:
    Gui, EllipsoidGui:Add, DropDownList, x185 y57 w50 vEllipUnits Choose1, mm|cm
    Gui, EllipsoidGui:Add, Button, x10 y95 w100 gCalcEllipsoid, Calculate
    Gui, EllipsoidGui:Add, Button, x120 y95 w80 gEllipsoidGuiClose, Cancel
    Gui, EllipsoidGui:Show, x%xPos% y%yPos% w250 h135, Ellipsoid Volume
    return
}

EllipsoidGuiClose:
    Gui, EllipsoidGui:Destroy
return

CalcEllipsoid:
    Gui, EllipsoidGui:Submit, NoHide

    if (EllipDim1 = "" || EllipDim2 = "" || EllipDim3 = "") {
        MsgBox, 16, Error, Please enter all three dimensions.
        return
    }

    d1 := EllipDim1 + 0
    d2 := EllipDim2 + 0
    d3 := EllipDim3 + 0

    if (d1 <= 0 || d2 <= 0 || d3 <= 0) {
        MsgBox, 16, Error, All dimensions must be greater than 0.
        return
    }

    ; Convert to cm if input is mm
    if (EllipUnits = "mm") {
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
    dimStr := EllipDim1 . " x " . EllipDim2 . " x " . EllipDim3 . " " . EllipUnits
    result := " This corresponds to a volume of " . volumeRound . " cc (" . dimStr . ")."

    Gui, EllipsoidGui:Destroy
    ShowResult(result)
return

; =========================================
; CALCULATOR 2: Adrenal Washout
; =========================================
ShowAdrenalWashoutGui() {
    CoordMode, Mouse, Screen
    MouseGetPos, mouseX, mouseY
    xPos := mouseX + 10
    yPos := mouseY + 10

    Gui, AdrenalGui:New, +AlwaysOnTop
    Gui, AdrenalGui:Add, Text, x10 y10 w280, Adrenal Washout Calculator
    Gui, AdrenalGui:Add, Text, x10 y40, Pre-contrast (HU):
    Gui, AdrenalGui:Add, Edit, x140 y37 w70 vAdrenalPre
    Gui, AdrenalGui:Add, Text, x10 y70, Post-contrast (HU):
    Gui, AdrenalGui:Add, Edit, x140 y67 w70 vAdrenalPost
    Gui, AdrenalGui:Add, Text, x10 y100, Delayed (15 min) (HU):
    Gui, AdrenalGui:Add, Edit, x140 y97 w70 vAdrenalDelayed
    Gui, AdrenalGui:Add, Button, x10 y135 w90 gCalcAdrenal, Calculate
    Gui, AdrenalGui:Add, Button, x110 y135 w80 gAdrenalGuiClose, Cancel
    Gui, AdrenalGui:Show, x%xPos% y%yPos% w230 h175, Adrenal Washout
    return
}

AdrenalGuiClose:
    Gui, AdrenalGui:Destroy
return

CalcAdrenal:
    Gui, AdrenalGui:Submit, NoHide

    if (AdrenalPre = "" || AdrenalPost = "" || AdrenalDelayed = "") {
        MsgBox, 16, Error, Please enter all three HU values.
        return
    }

    pre := AdrenalPre + 0
    post := AdrenalPost + 0
    delayed := AdrenalDelayed + 0

    ; WHY: Sentence style for inline insertion - matches smart parser output
    ; ARCHITECTURE: Leading space for dictation continuity
    result := " Adrenal washout: "

    ; Absolute washout (requires pre-contrast)
    if (post != pre) {
        absWashout := ((post - delayed) / (post - pre)) * 100
        absWashout := Round(absWashout, 1)
        result .= "absolute " . absWashout . "%"
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
        result .= "relative " . relWashout . "%"
        if (relWashout >= 40)
            result .= " (likely adenoma)"
        else
            result .= " (indeterminate)"
        result .= "."
    }

    ; Pre-contrast assessment as additional sentence
    if (pre <= 10)
        result .= " Pre-contrast " . pre . " HU suggests lipid-rich adenoma."

    if (ShowCitations)
        result .= " (Mayo-Smith et al. Radiology 2017)"

    Gui, AdrenalGui:Destroy
    ShowResult(result)
return

; =========================================
; CALCULATOR 3: NASCET Stenosis
; =========================================
ParseNASCET(input) {
    ; Try to parse distal and stenosis from text
    input := RegExReplace(input, "`r?\n", " ")

    ; Pattern 1: "distal X mm ... stenosis Y mm"
    if (RegExMatch(input, "i)distal.*?(\d+(?:\.\d+)?)\s*(?:mm|cm).*?stenosis.*?(\d+(?:\.\d+)?)\s*(?:mm|cm)", m)) {
        return CalculateNASCETResult(m1, m2)
    }
    ; Pattern 2: "stenosis X mm ... distal Y mm"
    if (RegExMatch(input, "i)stenosis.*?(\d+(?:\.\d+)?)\s*(?:mm|cm).*?distal.*?(\d+(?:\.\d+)?)\s*(?:mm|cm)", m)) {
        return CalculateNASCETResult(m2, m1)
    }
    ; Pattern 3: Just two numbers - assume larger is distal
    numbers := []
    pos := 1
    while (pos := RegExMatch(input, "(\d+(?:\.\d+)?)\s*(?:mm|cm)?", m, pos)) {
        numbers.Push(m1 + 0)
        pos += StrLen(m)
    }
    if (numbers.Length() >= 2) {
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

    result := " NASCET: " . nascetVal . "% stenosis (distal " . distalRound . "mm, stenosis " . stenosisRound . "mm), "

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
    CoordMode, Mouse, Screen
    MouseGetPos, mouseX, mouseY
    xPos := mouseX + 10
    yPos := mouseY + 10

    Gui, NASCETGui:New, +AlwaysOnTop
    Gui, NASCETGui:Add, Text, x10 y10 w280, NASCET Carotid Stenosis Calculator
    Gui, NASCETGui:Add, Text, x10 y40, Distal ICA diameter (mm):
    Gui, NASCETGui:Add, Edit, x160 y37 w60 vNASCETDistal
    Gui, NASCETGui:Add, Text, x10 y70, Stenosis diameter (mm):
    Gui, NASCETGui:Add, Edit, x160 y67 w60 vNASCETStenosis
    Gui, NASCETGui:Add, Button, x10 y105 w90 gCalcNASCETGui, Calculate
    Gui, NASCETGui:Add, Button, x110 y105 w80 gNASCETGuiClose, Cancel
    Gui, NASCETGui:Show, x%xPos% y%yPos% w240 h145, NASCET Calculator
    return
}

NASCETGuiClose:
    Gui, NASCETGui:Destroy
return

CalcNASCETGui:
    Gui, NASCETGui:Submit, NoHide

    if (NASCETDistal = "" || NASCETStenosis = "") {
        MsgBox, 16, Error, Please enter both measurements.
        return
    }

    distal := NASCETDistal + 0
    stenosis := NASCETStenosis + 0

    if (distal <= 0) {
        MsgBox, 16, Error, Distal ICA must be greater than 0.
        return
    }

    if (stenosis >= distal) {
        MsgBox, 16, Error, Stenosis should be less than distal ICA.
        return
    }

    result := CalculateNASCETResult(distal, stenosis)
    Gui, NASCETGui:Destroy
    ShowResult(result)
return

; -----------------------------------------
; CALCULATOR 3b: General Stenosis Calculator
; NOTE: 90% similar to NASCET calculator above. Consider merging
; into single parameterized function: ShowStenosisGui(preset := "")
; where preset="NASCET" adds ICA-specific labels.
; -----------------------------------------
ShowStenosisGui() {
    CoordMode, Mouse, Screen
    MouseGetPos, mouseX, mouseY
    xPos := mouseX + 10
    yPos := mouseY + 10

    Gui, StenosisGui:New, +AlwaysOnTop
    Gui, StenosisGui:Add, Text, x10 y10 w280, Vessel Stenosis Calculator
    Gui, StenosisGui:Add, Text, x10 y35, Vessel (optional):
    Gui, StenosisGui:Add, Edit, x120 y32 w130 vStenosisVessel,
    Gui, StenosisGui:Add, Text, x10 y65, Normal diameter (mm):
    Gui, StenosisGui:Add, Edit, x140 y62 w60 vStenosisNormal
    Gui, StenosisGui:Add, Text, x10 y95, Stenosis diameter (mm):
    Gui, StenosisGui:Add, Edit, x140 y92 w60 vStenosisMeasure
    Gui, StenosisGui:Add, Button, x10 y130 w90 gCalcStenosis, Calculate
    Gui, StenosisGui:Add, Button, x110 y130 w80 gStenosisGuiClose, Cancel
    Gui, StenosisGui:Show, x%xPos% y%yPos% w270 h170, Stenosis Calculator
    return
}

StenosisGuiClose:
    Gui, StenosisGui:Destroy
return

CalcStenosis:
    Gui, StenosisGui:Submit, NoHide
    global ShowCitations

    if (StenosisNormal = "" || StenosisMeasure = "") {
        MsgBox, 16, Error, Please enter both diameter measurements.
        return
    }

    normal := StenosisNormal + 0
    stenosis := StenosisMeasure + 0

    if (normal <= 0) {
        MsgBox, 16, Error, Normal diameter must be greater than 0.
        return
    }

    if (stenosis >= normal) {
        MsgBox, 16, Error, Stenosis should be less than normal diameter.
        return
    }

    ; Calculate percent stenosis
    stenosisPercent := ((normal - stenosis) / normal) * 100
    stenosisPercent := Round(stenosisPercent, 1)

    vesselName := StenosisVessel != "" ? StenosisVessel : "Vessel"

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
    result := " " . vesselName . " stenosis: " . stenosisPercent . "% (normal " . normalRound . "mm, stenosis " . stenosisRound . "mm), " . severity . "."

    Gui, StenosisGui:Destroy
    ShowResult(result)
return

; =========================================
; CALCULATOR 4: RV/LV Ratio
; =========================================
ShowRVLVGui() {
    CoordMode, Mouse, Screen
    MouseGetPos, mouseX, mouseY
    xPos := mouseX + 10
    yPos := mouseY + 10

    Gui, RVLVGui:New, +AlwaysOnTop
    Gui, RVLVGui:Add, Text, x10 y10 w280, RV/LV Ratio Calculator (4-chamber axial)
    Gui, RVLVGui:Add, Text, x10 y40, RV diameter (mm):
    Gui, RVLVGui:Add, Edit, x130 y37 w60 vRVDiam
    Gui, RVLVGui:Add, Text, x10 y70, LV diameter (mm):
    Gui, RVLVGui:Add, Edit, x130 y67 w60 vLVDiam
    Gui, RVLVGui:Add, Button, x10 y105 w90 gCalcRVLV, Calculate
    Gui, RVLVGui:Add, Button, x110 y105 w80 gRVLVGuiClose, Cancel
    Gui, RVLVGui:Show, x%xPos% y%yPos% w220 h145, RV/LV Ratio
    return
}

RVLVGuiClose:
    Gui, RVLVGui:Destroy
return

CalcRVLV:
    Gui, RVLVGui:Submit, NoHide

    if (RVDiam = "" || LVDiam = "") {
        MsgBox, 16, Error, Please enter both RV and LV diameters.
        return
    }

    rv := RVDiam + 0
    lv := LVDiam + 0

    if (lv <= 0) {
        MsgBox, 16, Error, LV diameter must be greater than 0.
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
    result := " RV/LV ratio: " . ratio . " (RV " . rvRound . "mm, LV " . lvRound . "mm), " . interpretation . "."

    if (ShowCitations)
        result .= " (Meinel et al. Radiology 2015)"

    Gui, RVLVGui:Destroy
    ShowResult(result)
return

; =========================================
; CALCULATOR 5: Fleischner 2017
; =========================================
ShowFleischnerGui() {
    global IncludeDatamining
    CoordMode, Mouse, Screen
    MouseGetPos, mouseX, mouseY
    xPos := mouseX + 10
    yPos := mouseY + 10

    Gui, FleischnerGui:New, +AlwaysOnTop
    Gui, FleischnerGui:Add, Text, x10 y10 w300, Fleischner 2017 - Incidental Pulmonary Nodule
    Gui, FleischnerGui:Add, Text, x10 y40, Nodule size (mm):
    Gui, FleischnerGui:Add, Edit, x130 y37 w60 vFleischSize
    Gui, FleischnerGui:Add, Text, x10 y70, Nodule type:
    Gui, FleischnerGui:Add, DropDownList, x130 y67 w120 vFleischType Choose1, Solid|Part-solid|Ground glass
    Gui, FleischnerGui:Add, Text, x10 y100, Number:
    Gui, FleischnerGui:Add, DropDownList, x130 y97 w120 vFleischNumber Choose1, Single|Multiple
    Gui, FleischnerGui:Add, Text, x10 y130, Risk:
    Gui, FleischnerGui:Add, DropDownList, x130 y127 w120 vFleischRisk Choose1, Low risk|High risk
    dmChecked := IncludeDatamining ? "Checked" : ""
    Gui, FleischnerGui:Add, Checkbox, x10 y160 w250 vFleischDatamine %dmChecked%, Include datamining phrase (SSM lung nodule)
    Gui, FleischnerGui:Add, Button, x10 y190 w100 gCalcFleischner, Get Recommendation
    Gui, FleischnerGui:Add, Button, x120 y190 w80 gFleischnerGuiClose, Cancel
    Gui, FleischnerGui:Show, x%xPos% y%yPos% w280 h230, Fleischner 2017
    return
}

FleischnerGuiClose:
    Gui, FleischnerGui:Destroy
return

CalcFleischner:
    Gui, FleischnerGui:Submit, NoHide

    if (FleischSize = "") {
        MsgBox, 16, Error, Please enter nodule size.
        return
    }

    size := FleischSize + 0
    result := "Fleischner 2017 Recommendation:`n"
    result .= "Size: " . size . " mm | " . FleischType . " | " . FleischNumber . " | " . FleischRisk . "`n`n"

    ; Determine recommendation based on 2017 Fleischner guidelines
    recommendation := GetFleischnerRecommendation(size, FleischType, FleischNumber, FleischRisk)
    result .= recommendation

    if (FleischDatamine) {
        result .= "`n`n[" . DataminingPhrase . "]"
    }

    if (ShowCitations)
        result .= "`n`nRef: MacMahon H et al. Radiology 2017;284:228-243"

    Gui, FleischnerGui:Destroy
    ShowResult(result)
return

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
    ; Position near mouse
    CoordMode, Mouse, Screen
    MouseGetPos, mouseX, mouseY
    xPos := mouseX + 10
    yPos := mouseY + 10

    ; Pre-fill values (convert to string)
    dim1 := Round(d1, 2)
    dim2 := Round(d2, 2)
    dim3 := Round(d3, 2)

    Gui, EllipsoidGui:New, +AlwaysOnTop
    Gui, EllipsoidGui:Add, Text, x10 y10 w280, Ellipsoid Volume Calculator (Pre-filled)
    Gui, EllipsoidGui:Add, Text, x10 y35, AP (L):
    Gui, EllipsoidGui:Add, Edit, x80 y32 w50 vEllipDim1, %dim1%
    Gui, EllipsoidGui:Add, Text, x135 y35, x   T (W):
    Gui, EllipsoidGui:Add, Edit, x185 y32 w50 vEllipDim2, %dim2%
    Gui, EllipsoidGui:Add, Text, x10 y60, CC (H):
    Gui, EllipsoidGui:Add, Edit, x80 y57 w50 vEllipDim3, %dim3%
    Gui, EllipsoidGui:Add, Text, x135 y60, Units:
    Gui, EllipsoidGui:Add, DropDownList, x185 y57 w50 vEllipUnits Choose2, mm|cm
    Gui, EllipsoidGui:Add, Button, x10 y95 w100 gCalcEllipsoid, Calculate
    Gui, EllipsoidGui:Add, Button, x120 y95 w80 gEllipsoidGuiClose, Cancel
    Gui, EllipsoidGui:Show, x%xPos% y%yPos% w250 h135, Ellipsoid Volume
    return
}

ShowRVLVGuiPrefilled(rv, lv) {
    CoordMode, Mouse, Screen
    MouseGetPos, mouseX, mouseY
    xPos := mouseX + 10
    yPos := mouseY + 10

    rvVal := Round(rv, 1)
    lvVal := Round(lv, 1)

    Gui, RVLVGui:New, +AlwaysOnTop
    Gui, RVLVGui:Add, Text, x10 y10 w280, RV/LV Ratio Calculator (Pre-filled)
    Gui, RVLVGui:Add, Text, x10 y40, RV diameter (mm):
    Gui, RVLVGui:Add, Edit, x130 y37 w60 vRVDiam, %rvVal%
    Gui, RVLVGui:Add, Text, x10 y70, LV diameter (mm):
    Gui, RVLVGui:Add, Edit, x130 y67 w60 vLVDiam, %lvVal%
    Gui, RVLVGui:Add, Button, x10 y105 w90 gCalcRVLV, Calculate
    Gui, RVLVGui:Add, Button, x110 y105 w80 gRVLVGuiClose, Cancel
    Gui, RVLVGui:Show, x%xPos% y%yPos% w220 h145, RV/LV Ratio
    return
}

ShowNASCETGuiPrefilled(distal, stenosis) {
    CoordMode, Mouse, Screen
    MouseGetPos, mouseX, mouseY
    xPos := mouseX + 10
    yPos := mouseY + 10

    distalVal := Round(distal, 1)
    stenosisVal := Round(stenosis, 1)

    Gui, NASCETGui:New, +AlwaysOnTop
    Gui, NASCETGui:Add, Text, x10 y10 w280, NASCET Calculator (Pre-filled)
    Gui, NASCETGui:Add, Text, x10 y40, Distal ICA diameter (mm):
    Gui, NASCETGui:Add, Edit, x160 y37 w60 vNASCETDistal, %distalVal%
    Gui, NASCETGui:Add, Text, x10 y70, Stenosis diameter (mm):
    Gui, NASCETGui:Add, Edit, x160 y67 w60 vNASCETStenosis, %stenosisVal%
    Gui, NASCETGui:Add, Button, x10 y105 w90 gCalcNASCETGui, Calculate
    Gui, NASCETGui:Add, Button, x110 y105 w80 gNASCETGuiClose, Cancel
    Gui, NASCETGui:Show, x%xPos% y%yPos% w240 h145, NASCET Calculator
    return
}

ShowAdrenalWashoutGuiPrefilled(pre, post, delayed) {
    CoordMode, Mouse, Screen
    MouseGetPos, mouseX, mouseY
    xPos := mouseX + 10
    yPos := mouseY + 10

    preVal := (pre != "") ? Round(pre, 0) : ""
    postVal := Round(post, 0)
    delayedVal := Round(delayed, 0)

    Gui, AdrenalGui:New, +AlwaysOnTop
    Gui, AdrenalGui:Add, Text, x10 y10 w280, Adrenal Washout Calculator (Pre-filled)
    Gui, AdrenalGui:Add, Text, x10 y40, Pre-contrast (HU):
    Gui, AdrenalGui:Add, Edit, x140 y37 w70 vAdrenalPre, %preVal%
    Gui, AdrenalGui:Add, Text, x10 y70, Post-contrast (HU):
    Gui, AdrenalGui:Add, Edit, x140 y67 w70 vAdrenalPost, %postVal%
    Gui, AdrenalGui:Add, Text, x10 y100, Delayed (15 min) (HU):
    Gui, AdrenalGui:Add, Edit, x140 y97 w70 vAdrenalDelayed, %delayedVal%
    Gui, AdrenalGui:Add, Button, x10 y135 w90 gCalcAdrenal, Calculate
    Gui, AdrenalGui:Add, Button, x110 y135 w80 gAdrenalGuiClose, Cancel
    Gui, AdrenalGui:Show, x%xPos% y%yPos% w230 h175, Adrenal Washout
    return
}

ShowFleischnerGuiPrefilled(size, nodeType, number) {
    global IncludeDatamining
    CoordMode, Mouse, Screen
    MouseGetPos, mouseX, mouseY
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

    Gui, FleischnerGui:New, +AlwaysOnTop
    Gui, FleischnerGui:Add, Text, x10 y10 w300, Fleischner 2017 (Pre-filled)
    Gui, FleischnerGui:Add, Text, x10 y40, Nodule size (mm):
    Gui, FleischnerGui:Add, Edit, x130 y37 w60 vFleischSize, %sizeVal%
    Gui, FleischnerGui:Add, Text, x10 y70, Nodule type:
    Gui, FleischnerGui:Add, DropDownList, x130 y67 w120 vFleischType Choose%typeSelect%, Solid|Part-solid|Ground glass
    Gui, FleischnerGui:Add, Text, x10 y100, Number:
    Gui, FleischnerGui:Add, DropDownList, x130 y97 w120 vFleischNumber Choose%numberSelect%, Single|Multiple
    Gui, FleischnerGui:Add, Text, x10 y130, Risk:
    Gui, FleischnerGui:Add, DropDownList, x130 y127 w120 vFleischRisk Choose1, Low risk|High risk
    dmChecked := IncludeDatamining ? "Checked" : ""
    Gui, FleischnerGui:Add, Checkbox, x10 y160 w250 vFleischDatamine %dmChecked%, Include datamining phrase
    Gui, FleischnerGui:Add, Button, x10 y190 w100 gCalcFleischner, Get Recommendation
    Gui, FleischnerGui:Add, Button, x120 y190 w80 gFleischnerGuiClose, Cancel
    Gui, FleischnerGui:Show, x%xPos% y%yPos% w280 h230, Fleischner 2017
    return
}

; -----------------------------------------
; Utility: Get mouse position for GUI placement
; WHY: Reduces code duplication across 15+ GUI functions
; -----------------------------------------
GetGuiPosition(ByRef xPos, ByRef yPos, offset := 10) {
    CoordMode, Mouse, Screen
    MouseGetPos, mouseX, mouseY
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
    ClipSaved := ClipboardAll
    Clipboard := text
    ClipWait, 0.5
    Send, ^v
    Sleep, %waitMs%
    Clipboard := ClipSaved
}

; -----------------------------------------
; Utility: Insert text after current selection
; WHY: Preserves original text and appends calculation result.
; -----------------------------------------
InsertAfterSelection(textToInsert) {
    ; Move cursor to end of selection and insert text
    Send, {Right}
    Sleep, 50

    ; Use helper to paste text while preserving clipboard
    PasteTextPreserveClipboard(textToInsert)

    ToolTip, Calculation inserted!
    SetTimer, RemoveToolTip, -1500
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
    ClipSaved := ClipboardAll

    ; Copy document to check for IMPRESSION: existence
    ; NOTE: This briefly exposes full document on clipboard; cleared immediately after check
    ; TRADEOFF: Faster than incremental search, clipboard cleared right after copy
    Clipboard := ""
    Send, ^a  ; Select all
    Sleep, 20
    Send, ^c  ; Copy
    ClipWait, 0.5
    docText := Clipboard
    Clipboard := ""  ; Clear clipboard immediately to minimize PHI exposure

    ; CRITICAL: Deselect document text before any further operations
    ; WHY: Prevents accidental replacement if Find dialog is slow to open
    Send, {Escape}
    Sleep, 30
    Send, {Right}
    Sleep, 30

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
        Clipboard := ClipSaved  ; Restore clipboard first
        InsertAfterSelection(textToInsert)
        ToolTip, IMPRESSION: not found - inserted at cursor
        SetTimer, RemoveToolTip, -2000
        return
    }

    ; Go to start of document
    Send, ^{Home}
    Sleep, 50

    ; Open Find dialog (Ctrl+F)
    Send, ^f
    Sleep, 500  ; WHY: PowerScribe Find dialog can be very slow to open

    ; Clear any previous search and type new search text
    ; NOTE: ^a here targets the Find dialog's text field, not the document
    Send, ^a
    Sleep, 50
    Send, {Raw}%searchTerm%
    Sleep, 200

    ; Press Enter to find (or F3/Find Next depending on app)
    Send, {Enter}
    Sleep, 150

    ; Close Find dialog (Escape) - press twice to ensure closed
    Send, {Escape}
    Sleep, 100
    Send, {Escape}
    Sleep, 100

    ; Go to end of line where IMPRESSION: was found
    Send, {End}
    Sleep, 50

    ; Add blank line (Enter twice for spacing)
    Send, {Enter}{Enter}
    Sleep, 30

    ; Set clipboard to text and paste
    Clipboard := textToInsert
    ClipWait, 0.5
    Send, ^v
    Sleep, 100

    ; Restore clipboard
    Clipboard := ClipSaved

    ToolTip, Inserted after IMPRESSION:
    SetTimer, RemoveToolTip, -1500
}

; -----------------------------------------
; Utility: Deduplicate array of sizes (removes duplicates within 1mm tolerance)
; WHY: Multiple patterns may match same nodule, avoid double-counting
; -----------------------------------------
DeduplicateSizes(arr) {
    if (arr.Length() <= 1)
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
    CoordMode, Mouse, Screen
    MouseGetPos, mouseX, mouseY
    xPos := mouseX + 10
    yPos := mouseY + 10

    ; Destroy any existing confirm GUI
    Gui, ConfirmGui:Destroy

    Gui, ConfirmGui:New, +AlwaysOnTop
    Gui, ConfirmGui:Add, Text, x10 y10 w350 +0x1, Smart Parse Confirmation
    Gui, ConfirmGui:Add, Edit, x10 y35 w350 h130 ReadOnly, %displayText%

    ; Different buttons based on confidence
    if (confidence < 50 && SmartParseFallbackToGUI) {
        ; Low confidence: default to Cancel, offer GUI edit
        Gui, ConfirmGui:Add, Button, x10 y175 w100 gConfirmInsert, Insert Anyway
        Gui, ConfirmGui:Add, Button, x120 y175 w110 gConfirmEdit Default, Edit in GUI
        Gui, ConfirmGui:Add, Button, x240 y175 w100 gConfirmCancel, Cancel
    } else {
        ; High/medium confidence: default to Insert
        Gui, ConfirmGui:Add, Button, x10 y175 w100 gConfirmInsert Default, Insert
        Gui, ConfirmGui:Add, Button, x120 y175 w110 gConfirmEdit, Edit in GUI
        Gui, ConfirmGui:Add, Button, x240 y175 w100 gConfirmCancel, Cancel
    }

    Gui, ConfirmGui:Show, x%xPos% y%yPos% w380 h215, Confirm Parse
    WinGet, hWnd, ID, A

    ; Reset and wait for user action (event-driven via WinWaitClose)
    g_ConfirmAction := ""
    WinWaitClose, ahk_id %hWnd%

    return g_ConfirmAction
}

ConfirmInsert:
    global g_ConfirmAction
    g_ConfirmAction := "insert"
    Gui, ConfirmGui:Destroy
return

ConfirmEdit:
    global g_ConfirmAction
    g_ConfirmAction := "edit"
    Gui, ConfirmGui:Destroy
return

ConfirmCancel:
    global g_ConfirmAction
    g_ConfirmAction := "cancel"
    Gui, ConfirmGui:Destroy
return

ConfirmGuiClose:
    global g_ConfirmAction
    g_ConfirmAction := "cancel"
    Gui, ConfirmGui:Destroy
return

ConfirmGuiEscape:
    global g_ConfirmAction
    g_ConfirmAction := "cancel"
    Gui, ConfirmGui:Destroy
return

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

    if (RegExMatch(filtered, strictPattern, m)) {
        organ := m1
        d1 := m2
        d2 := m3
        d3 := m4
        units := m5
        confidence := 95
    }
    ; Pattern B (MEDIUM confidence 70): "measures" keyword + dimensions + units (no organ)
    else if (RegExMatch(filtered, "i)(?:measures?|measuring)\s+(\d+(?:\.\d+)?)\s*[x×]\s*(\d+(?:\.\d+)?)\s*[x×]\s*(\d+(?:\.\d+)?)\s*(cm|mm)", m)) {
        d1 := m1
        d2 := m2
        d3 := m3
        units := m4
        confidence := 70
    }
    ; Pattern C (MEDIUM confidence 60): dimensions + units (no keyword)
    else if (RegExMatch(filtered, "(\d+(?:\.\d+)?)\s*[x×]\s*(\d+(?:\.\d+)?)\s*[x×]\s*(\d+(?:\.\d+)?)\s*(cm|mm)", m)) {
        d1 := m1
        d2 := m2
        d3 := m3
        units := m4
        confidence := 60
    }
    ; Pattern D (LOW confidence 35): just three numbers with x separator (no units)
    ; WHY: Use DefaultMeasurementUnit setting when units not detected
    else if (RegExMatch(filtered, "(\d+(?:\.\d+)?)\s*[x×]\s*(\d+(?:\.\d+)?)\s*[x×]\s*(\d+(?:\.\d+)?)", m)) {
        d1 := m1
        d2 := m2
        d3 := m3
        units := DefaultMeasurementUnit  ; Use user's default setting
        confidence := 35
    }
    else {
        ; No match found
        MsgBox, 48, Parse Error, Could not find dimensions in text.`n`nExpected: "measures 8.0 x 6.0 x 9.0 cm"`n`nTip: Include units (cm or mm) for better detection.
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
        StringLower, organLower, organ
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
    if (RegExMatch(filtered, patternA, m)) {
        rv := m1 + 0
        rvUnits := NormalizeUnits(m2)
        lv := m3 + 0
        lvUnits := NormalizeUnits(m4)
        confidence := 90
    }
    else if (RegExMatch(filtered, patternB, m)) {
        lv := m1 + 0
        lvUnits := NormalizeUnits(m2)
        rv := m3 + 0
        rvUnits := NormalizeUnits(m4)
        confidence := 85
    }
    else if (RegExMatch(filtered, patternE, m)) {
        rv := m1 + 0
        rvUnits := NormalizeUnits(m2)
        lv := m3 + 0
        lvUnits := NormalizeUnits(m4)
        confidence := 85
    }
    else if (RegExMatch(filtered, patternG, m)) {
        rv := m1 + 0
        rvUnits := NormalizeUnits(m2)
        lv := m3 + 0
        lvUnits := NormalizeUnits(m4)
        confidence := 85
    }
    else if (RegExMatch(filtered, patternF, m)) {
        rv := m1 + 0
        rvUnits := NormalizeUnits(m2)
        lv := m3 + 0
        lvUnits := NormalizeUnits(m4)
        confidence := 80
    }
    else if (RegExMatch(filtered, patternC, m)) {
        rv := m1 + 0
        lv := m2 + 0
        confidence := 80
    }
    else if (RegExMatch(filtered, patternD, m)) {
        rv := m1 + 0
        rvUnits := NormalizeUnits(m2)
        lv := m3 + 0
        lvUnits := NormalizeUnits(m4)
        confidence := 75
    }
    else if (RegExMatch(filtered, patternH, m)) {
        ; PACS "rvval/lvval" format - default to cm if no units (typical cardiac CT)
        rv := m1 + 0
        rvUnits := (m2 != "") ? NormalizeUnits(m2) : "cm"
        lv := m3 + 0
        lvUnits := (m4 != "") ? NormalizeUnits(m4) : "cm"
        confidence := 80
    }
    ; No pattern matched
    else {
        MsgBox, 48, Parse Error, Could not find RV/LV measurements.`n`nExpected formats:`n- "RV 42mm / LV 35mm"`n- "RV: 42, LV: 35"`n- "Right ventricle 4.2 cm"`n- "rvval 5.0 lvval 3.6"`n`nMust include RV/LV keywords.
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
        MsgBox, 48, Error, LV diameter must be greater than 0.
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
    StringLower, interpretLower, interpretation
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
    if (RegExMatch(filtered, "i)distal\s*(?:ICA|internal\s*carotid)?.*?(\d+(?:\.\d+)?)\s*(mm|cm)?.*?stenosis.*?(\d+(?:\.\d+)?)\s*(mm|cm)?", m)) {
        distal := m1 + 0
        distalUnits := (m2 != "") ? m2 : "mm"
        stenosis := m3 + 0
        stenosisUnits := (m4 != "") ? m4 : "mm"
        confidence := 90
    }
    ; Pattern B (HIGH confidence 85): "stenosis X mm ... distal Y mm" (reverse order)
    else if (RegExMatch(filtered, "i)stenosis.*?(\d+(?:\.\d+)?)\s*(mm|cm)?.*?distal\s*(?:ICA)?.*?(\d+(?:\.\d+)?)\s*(mm|cm)?", m)) {
        stenosis := m1 + 0
        stenosisUnits := (m2 != "") ? m2 : "mm"
        distal := m3 + 0
        distalUnits := (m4 != "") ? m4 : "mm"
        confidence := 85
    }
    ; Pattern C (MEDIUM confidence 65): ICA/carotid + narrowing context
    else if (RegExMatch(filtered, "i)(?:ICA|carotid).*?(\d+(?:\.\d+)?)\s*(mm|cm)?.*?(?:narrow|residual).*?(\d+(?:\.\d+)?)\s*(mm|cm)?", m)) {
        d1 := m1 + 0
        d1Units := (m2 != "") ? m2 : "mm"
        d2 := m3 + 0
        d2Units := (m4 != "") ? m4 : "mm"
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
        MsgBox, 48, Parse Error, Could not find stenosis measurements.`n`nExpected formats:`n- "distal 5.2mm, stenosis 2.1mm"`n- "distal ICA 0.5cm, stenosis 3mm"`n`nMust include "distal" or "stenosis" keywords.
        return
    }

    ; Convert cm to mm for consistent calculation
    if (distalUnits = "cm")
        distal := distal * 10
    if (stenosisUnits = "cm")
        stenosis := stenosis * 10

    ; Validate
    if (distal <= 0) {
        MsgBox, 48, Error, Distal ICA diameter must be greater than 0.
        return
    }
    if (stenosis >= distal) {
        MsgBox, 48, Error, Stenosis diameter must be less than distal diameter.`n`nDistal: %distal% mm`nStenosis: %stenosis% mm
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

    if (RegExMatch(filtered, fullPattern, m)) {
        pre := m1 + 0
        post := m2 + 0
        delayed := m3 + 0
        confidence := 95
    }
    ; Pattern B (MEDIUM confidence 70): Post + delayed with labels
    else if (RegExMatch(filtered, "i)(?:post-?contrast|enhanced)[:\s]*(-?\d+)\s*HU.*?(?:delayed|15\s*min)[:\s]*(-?\d+)\s*HU", m)) {
        post := m1 + 0
        delayed := m2 + 0
        confidence := 70
    }
    ; Pattern C (MEDIUM confidence 55): Three HU values in sequence
    else if (RegExMatch(filtered, "(-?\d+)\s*HU.*?(-?\d+)\s*HU.*?(-?\d+)\s*HU", m)) {
        pre := m1 + 0
        post := m2 + 0
        delayed := m3 + 0
        confidence := 55
    }
    ; NO raw number fallback - requires HU labels for safety
    else {
        MsgBox, 48, Parse Error, Could not find HU values.`n`nExpected formats:`n- "pre-contrast 10 HU, enhanced 80 HU, delayed 40 HU"`n- "10 HU, 80 HU, 40 HU"`n`nMust include "HU" labels.
        return
    }

    if (post = "" || delayed = "") {
        MsgBox, 48, Error, At least post-contrast and delayed HU values are required.
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
ClassifyAndStore(size, units, nodeType, ByRef solidArr, ByRef subsolidArr, ByRef partsolidArr) {
    ; Convert to mm
    if (units = "cm" || units = "centimeter" || units = "centimeters")
        size := size * 10

    ; Only accept plausible nodule sizes (1-50mm)
    if (size < 1 || size > 50)
        return false

    ; Classify based on type keywords (case insensitive)
    StringLower, nodeTypeLower, nodeType
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
    while (pos := RegExMatch(filtered, pattern1, m, pos)) {
        size := m1 + 0
        size2 := m2 != "" ? m2 + 0 : 0
        units := m3
        finalSize := (size2 > size) ? size2 : size
        ClassifyAndStore(finalSize, units, "", solidNodules, subsolidNodules, partsolidNodules)
        pos += StrLen(m)
    }

    ; Search Pattern 1b: Nodule then size
    pos := 1
    while (pos := RegExMatch(filtered, pattern1b, m, pos)) {
        size := m1 + 0
        size2 := m2 != "" ? m2 + 0 : 0
        units := m3
        finalSize := (size2 > size) ? size2 : size
        ClassifyAndStore(finalSize, units, "", solidNodules, subsolidNodules, partsolidNodules)
        pos += StrLen(m)
    }

    ; Search Pattern 2: Type + size + nodule
    pos := 1
    while (pos := RegExMatch(filtered, pattern2, m, pos)) {
        nodeType := m1
        size := m2 + 0
        size2 := m3 != "" ? m3 + 0 : 0
        units := m4
        finalSize := (size2 > size) ? size2 : size
        ClassifyAndStore(finalSize, units, nodeType, solidNodules, subsolidNodules, partsolidNodules)
        pos += StrLen(m)
    }

    ; Search Pattern 2b: Type + nodule + size
    pos := 1
    while (pos := RegExMatch(filtered, pattern2b, m, pos)) {
        nodeType := m1
        size := m2 + 0
        size2 := m3 != "" ? m3 + 0 : 0
        units := m4
        finalSize := (size2 > size) ? size2 : size
        ClassifyAndStore(finalSize, units, nodeType, solidNodules, subsolidNodules, partsolidNodules)
        pos += StrLen(m)
    }

    ; Search Pattern 2c: Size + type + nodule
    pos := 1
    while (pos := RegExMatch(filtered, pattern2c, m, pos)) {
        size := m1 + 0
        size2 := m2 != "" ? m2 + 0 : 0
        units := m3
        nodeType := m4
        finalSize := (size2 > size) ? size2 : size
        ClassifyAndStore(finalSize, units, nodeType, solidNodules, subsolidNodules, partsolidNodules)
        pos += StrLen(m)
    }

    ; Search Pattern 3: "largest/dominant" nodule
    pos := 1
    while (pos := RegExMatch(filtered, pattern3, m, pos)) {
        size := m1 + 0
        size2 := m2 != "" ? m2 + 0 : 0
        units := m3
        finalSize := (size2 > size) ? size2 : size
        ClassifyAndStore(finalSize, units, "", solidNodules, subsolidNodules, partsolidNodules)
        pos += StrLen(m)
    }

    ; Search Pattern 4: "nodule is/measures X mm"
    pos := 1
    while (pos := RegExMatch(filtered, pattern4, m, pos)) {
        size := m1 + 0
        size2 := m2 != "" ? m2 + 0 : 0
        units := m3
        finalSize := (size2 > size) ? size2 : size
        ClassifyAndStore(finalSize, units, "", solidNodules, subsolidNodules, partsolidNodules)
        pos += StrLen(m)
    }

    ; Search Pattern 5: Parenthetical "nodule (8 mm)"
    pos := 1
    while (pos := RegExMatch(filtered, pattern5, m, pos)) {
        size := m1 + 0
        size2 := m2 != "" ? m2 + 0 : 0
        units := m3
        finalSize := (size2 > size) ? size2 : size
        ClassifyAndStore(finalSize, units, "", solidNodules, subsolidNodules, partsolidNodules)
        pos += StrLen(m)
    }

    ; Check for "sub 6 mm nodules", "sub-6mm", "punctate nodules", "<6mm nodules", "scattered tiny nodules"
    if (RegExMatch(filtered, "i)(sub[- ]?6|<\s*6|punctate|miliary|tiny|innumerable|scattered small|multiple small|few small).{0,25}(" . nodulePattern . ")")) {
        if (solidNodules.Length() = 0)
            solidNodules.Push(5)
    }

    ; Also check for "nodules less than 6 mm" pattern
    if (RegExMatch(filtered, "i)(" . nodulePattern . ").{0,20}(less than|smaller than|under|<)\s*6\s*(mm|cm)?")) {
        if (solidNodules.Length() = 0)
            solidNodules.Push(5)
    }

    ; Deduplicate arrays (same nodule might match multiple patterns)
    solidNodules := DeduplicateSizes(solidNodules)
    subsolidNodules := DeduplicateSizes(subsolidNodules)
    partsolidNodules := DeduplicateSizes(partsolidNodules)

    ; If no nodules found, show error
    if (solidNodules.Length() = 0 && subsolidNodules.Length() = 0 && partsolidNodules.Length() = 0) {
        MsgBox, 48, Parse Error, Could not find nodule descriptions in text.`n`nExpected: Text containing "nodule" (or similar) with a size measurement.`n`nExamples:`n- "8 mm pulmonary nodule"`n- "solid nodule measuring 8 mm"`n- "groundglass opacity, 7mm"
        return
    }

    ; Calculate confidence based on what was found
    totalNodules := solidNodules.Length() + subsolidNodules.Length() + partsolidNodules.Length()
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
        if (solidNodules.Length() > 1)
            resultText .= " (largest of " . solidNodules.Length() . ")"
        resultText .= "`n"
        noduleNum++
    }
    if (maxSubsolid > 0) {
        resultText .= noduleNum . ". " . maxSubsolid . " mm ground glass/subsolid nodule"
        if (subsolidNodules.Length() > 1)
            resultText .= " (largest of " . subsolidNodules.Length() . ")"
        resultText .= "`n"
        noduleNum++
    }
    if (maxPartsolid > 0) {
        resultText .= noduleNum . ". " . maxPartsolid . " mm part-solid nodule"
        if (partsolidNodules.Length() > 1)
            resultText .= " (largest of " . partsolidNodules.Length() . ")"
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
    if (RegExMatch(input, "i)(\d+(?:\.\d+)?)\s*x\s*(\d+(?:\.\d+)?)\s*x\s*(\d+(?:\.\d+)?)\s*(cm|mm)?", m)) {
        a := m1 + 0
        b := m2 + 0
        c := m3 + 0
        units := (m4 != "") ? m4 : "cm"
        confidence := 90
    }
    ; Pattern B (MEDIUM confidence 75): Labeled dimensions
    ; Example: "A 5.0, B 4.0, C 3.5" or "length 5, width 4, height 3"
    else if (RegExMatch(input, "i)(?:A|length|L)\s*[:\s=]?\s*(\d+(?:\.\d+)?)\s*(cm|mm)?.*?(?:B|width|W)\s*[:\s=]?\s*(\d+(?:\.\d+)?)\s*(cm|mm)?.*?(?:C|height|H|depth|D)\s*[:\s=]?\s*(\d+(?:\.\d+)?)\s*(cm|mm)?", m)) {
        a := m1 + 0
        units := (m2 != "") ? m2 : "cm"
        b := m3 + 0
        c := m5 + 0
        confidence := 75
    }
    ; Pattern C (LOW confidence 50): Three numbers in sequence
    else {
        numbers := []
        pos := 1
        while (pos := RegExMatch(input, "(\d+(?:\.\d+)?)\s*(cm|mm)?", m, pos)) {
            numbers.Push({val: m1 + 0, unit: m2})
            pos += StrLen(m)
        }
        if (numbers.Length() >= 3) {
            a := numbers[1].val
            b := numbers[2].val
            c := numbers[3].val
            units := (numbers[1].unit != "") ? numbers[1].unit : "cm"
            confidence := 50
        }
    }

    if (a = 0 || b = 0 || c = 0) {
        MsgBox, 48, Parse Error, Could not find hemorrhage dimensions.`n`nExpected formats:`n- "5.0 x 4.0 x 3.5 cm"`n- "hemorrhage measuring X x Y x Z"
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
ShowICHVolumeGui() {
    GetGuiPosition(xPos, yPos)

    Gui, ICHGui:New, +AlwaysOnTop
    Gui, ICHGui:Add, Text, x10 y10 w280, ICH Volume Calculator (ABC/2)
    Gui, ICHGui:Add, Text, x10 y40, A (longest axis, cm):
    Gui, ICHGui:Add, Edit, x140 y37 w60 vICHDimA
    Gui, ICHGui:Add, Text, x10 y70, B (perpendicular, cm):
    Gui, ICHGui:Add, Edit, x140 y67 w60 vICHDimB
    Gui, ICHGui:Add, Text, x10 y100, C (# of slices x thickness, cm):
    Gui, ICHGui:Add, Edit, x200 y97 w60 vICHDimC
    Gui, ICHGui:Add, Button, x30 y135 w80 gCalcICH, Calculate
    Gui, ICHGui:Add, Button, x130 y135 w80 gICHGuiClose, Cancel
    Gui, ICHGui:Show, x%xPos% y%yPos% w280 h175, ICH Volume (ABC/2)
    return
}

ShowICHVolumeGuiPrefilled(a, b, c) {
    GetGuiPosition(xPos, yPos)

    aVal := Round(a, 1)
    bVal := Round(b, 1)
    cVal := Round(c, 1)

    Gui, ICHGui:New, +AlwaysOnTop
    Gui, ICHGui:Add, Text, x10 y10 w280, ICH Volume Calculator (Pre-filled)
    Gui, ICHGui:Add, Text, x10 y40, A (longest axis, cm):
    Gui, ICHGui:Add, Edit, x140 y37 w60 vICHDimA, %aVal%
    Gui, ICHGui:Add, Text, x10 y70, B (perpendicular, cm):
    Gui, ICHGui:Add, Edit, x140 y67 w60 vICHDimB, %bVal%
    Gui, ICHGui:Add, Text, x10 y100, C (# of slices x thickness, cm):
    Gui, ICHGui:Add, Edit, x200 y97 w60 vICHDimC, %cVal%
    Gui, ICHGui:Add, Button, x30 y135 w80 gCalcICH, Calculate
    Gui, ICHGui:Add, Button, x130 y135 w80 gICHGuiClose, Cancel
    Gui, ICHGui:Show, x%xPos% y%yPos% w280 h175, ICH Volume (ABC/2)
    return
}

ICHGuiClose:
    Gui, ICHGui:Destroy
return

CalcICH:
    Gui, ICHGui:Submit, NoHide

    if (ICHDimA = "" || ICHDimB = "" || ICHDimC = "") {
        MsgBox, 16, Error, Please enter all three dimensions.
        return
    }

    a := ICHDimA + 0
    b := ICHDimB + 0
    c := ICHDimC + 0

    if (a <= 0 || b <= 0 || c <= 0) {
        MsgBox, 16, Error, All dimensions must be greater than 0.
        return
    }

    ; ABC/2 formula
    volume := (a * b * c) / 2
    volume := Round(volume, 1)

    ; Sentence style output
    result := " ICH volume approximately " . volume . " cc (ABC/2: " . Round(a, 1) . " x " . Round(b, 1) . " x " . Round(c, 1) . " cm)."

    Gui, ICHGui:Destroy
    ShowResult(result)
return

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
    if (RegExMatch(input, "i)(\d+)\s*(month|months|mo|m)\b", m)) {
        interval := m1 + 0
        intervalUnit := "months"
        confidence := 95
    } else if (RegExMatch(input, "i)(\d+)\s*(week|weeks|wk|w)\b", m)) {
        interval := m1 + 0
        intervalUnit := "weeks"
        confidence := 95
    } else if (RegExMatch(input, "i)(\d+)\s*(day|days|d)\b", m)) {
        interval := m1 + 0
        intervalUnit := "days"
        confidence := 95
    } else if (RegExMatch(input, "i)(\d+)\s*(year|years|yr|y)\b", m)) {
        interval := m1 + 0
        intervalUnit := "years"
        confidence := 95
    }
    ; Pattern B (MEDIUM confidence 80): Abbreviated format "3m", "6w", "45d", "1y"
    else if (RegExMatch(input, "i)\b(\d+)(m|mo)\b", m)) {
        interval := m1 + 0
        intervalUnit := "months"
        confidence := 80
    } else if (RegExMatch(input, "i)\b(\d+)(w|wk)\b", m)) {
        interval := m1 + 0
        intervalUnit := "weeks"
        confidence := 80
    } else if (RegExMatch(input, "i)\b(\d+)(d)\b", m)) {
        interval := m1 + 0
        intervalUnit := "days"
        confidence := 80
    } else if (RegExMatch(input, "i)\b(\d+)(y|yr)\b", m)) {
        interval := m1 + 0
        intervalUnit := "years"
        confidence := 80
    }

    if (interval = 0 || intervalUnit = "") {
        MsgBox, 48, Parse Error, Could not find follow-up interval.`n`nExpected formats:`n- "3 months", "6 weeks", "1 year"`n- "3m", "6w", "45d", "1y"
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
    FormatTime, today, , yyyyMMdd
    today := today + 0

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
    futureDate := today
    futureDate += daysToAdd, days

    ; Format the result date
    FormatTime, futureFormatted, %futureDate%, MMMM d, yyyy
    return futureFormatted
}

; -----------------------------------------
; Date Calculator GUI
; -----------------------------------------
ShowDateCalculatorGui() {
    GetGuiPosition(xPos, yPos)

    Gui, DateGui:New, +AlwaysOnTop
    Gui, DateGui:Add, Text, x10 y10 w250, Follow-up Date Calculator
    Gui, DateGui:Add, Text, x10 y45, Interval:
    Gui, DateGui:Add, Edit, x80 y42 w50 vDateInterval, 3
    Gui, DateGui:Add, DropDownList, x140 y42 w100 vDateUnit Choose1, Months|Weeks|Days|Years
    Gui, DateGui:Add, Button, x30 y85 w80 gCalcDate, Calculate
    Gui, DateGui:Add, Button, x130 y85 w80 gDateGuiClose, Cancel
    Gui, DateGui:Show, x%xPos% y%yPos% w260 h125, Follow-up Date
    return
}

ShowDateCalculatorGuiPrefilled(interval, unit) {
    GetGuiPosition(xPos, yPos)

    ; Determine dropdown selection
    unitSelect := 1
    if (unit = "weeks")
        unitSelect := 2
    else if (unit = "days")
        unitSelect := 3
    else if (unit = "years")
        unitSelect := 4

    Gui, DateGui:New, +AlwaysOnTop
    Gui, DateGui:Add, Text, x10 y10 w250, Follow-up Date Calculator (Pre-filled)
    Gui, DateGui:Add, Text, x10 y45, Interval:
    Gui, DateGui:Add, Edit, x80 y42 w50 vDateInterval, %interval%
    Gui, DateGui:Add, DropDownList, x140 y42 w100 vDateUnit Choose%unitSelect%, Months|Weeks|Days|Years
    Gui, DateGui:Add, Button, x30 y85 w80 gCalcDate, Calculate
    Gui, DateGui:Add, Button, x130 y85 w80 gDateGuiClose, Cancel
    Gui, DateGui:Show, x%xPos% y%yPos% w260 h125, Follow-up Date
    return
}

DateGuiClose:
    Gui, DateGui:Destroy
return

CalcDate:
    Gui, DateGui:Submit, NoHide

    if (DateInterval = "") {
        MsgBox, 16, Error, Please enter an interval.
        return
    }

    interval := DateInterval + 0
    if (interval <= 0) {
        MsgBox, 16, Error, Interval must be greater than 0.
        return
    }

    ; Normalize unit name
    unit := "months"
    if (DateUnit = "Weeks")
        unit := "weeks"
    else if (DateUnit = "Days")
        unit := "days"
    else if (DateUnit = "Years")
        unit := "years"

    ; Calculate future date
    futureDate := CalculateFutureDate(interval, unit)

    ; Sentence style output
    result := " Recommend follow-up in " . interval . " " . unit . " (approximately " . futureDate . ")."

    Gui, DateGui:Destroy
    ShowResult(result)
return

; =========================================
; TOOL 6: Sectra History Copy
; =========================================
CopySectraHistory() {
    global SectraWindowTitle

    ; Save clipboard
    ClipSaved := ClipboardAll
    Clipboard := ""

    ; Check if text is already selected (from our earlier capture)
    if (g_SelectedText != "") {
        Clipboard := g_SelectedText
    } else {
        ; Try to get from Sectra
        if (WinExist(SectraWindowTitle)) {
            WinActivate, %SectraWindowTitle%
            WinWaitActive, %SectraWindowTitle%, , 2
            if (!ErrorLevel) {
                Send, ^c
                ClipWait, 1
            }
        }
    }

    if (Clipboard = "") {
        MsgBox, 48, Info, No text found. Select history text in Sectra first.
        Clipboard := ClipSaved
        return
    }

    historyText := Clipboard

    ; Find PowerScribe and paste
    foundPS := false
    for index, app in TargetApps {
        if (InStr(app, "PowerScribe") || InStr(app, "Nuance")) {
            if (WinExist(app)) {
                WinActivate, %app%
                WinWaitActive, %app%, , 2
                if (!ErrorLevel) {
                    foundPS := true
                    break
                }
            }
        }
    }

    if (!foundPS) {
        ; Just activate any target app
        WinActivate, ahk_exe notepad.exe
        WinWaitActive, ahk_exe notepad.exe, , 1
    }

    Sleep, 100
    Send, ^v

    ; Restore clipboard
    Sleep, 100
    Clipboard := ClipSaved

    ToolTip, History pasted!
    SetTimer, RemoveToolTip, -1500
    return
}

RemoveToolTip:
    ToolTip
return

; =========================================
; Settings GUI
; =========================================
ShowSettings() {
    global IncludeDatamining, ShowCitations, DataminingPhrase, DefaultSmartParse
    global SmartParseConfirmation, SmartParseFallbackToGUI
    global DefaultMeasurementUnit, RVLVOutputFormat, FleischnerInsertAfterImpression
    global ShowVolumeTools, ShowRVLVTools, ShowNASCETTools
    global ShowAdrenalTools, ShowFleischnerTools, ShowStenosisTools
    global ShowICHTools, ShowDateCalculator

    CoordMode, Mouse, Screen
    MouseGetPos, mouseX, mouseY
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

    ; Determine which item to select in dropdown
    smartParseOptions := "Volume|RVLV|NASCET|Adrenal|Fleischner"
    if (DefaultSmartParse = "RVLV")
        smartParseOptions := "Volume|RVLV||NASCET|Adrenal|Fleischner"
    else if (DefaultSmartParse = "NASCET")
        smartParseOptions := "Volume|RVLV|NASCET||Adrenal|Fleischner"
    else if (DefaultSmartParse = "Adrenal")
        smartParseOptions := "Volume|RVLV|NASCET|Adrenal||Fleischner"
    else if (DefaultSmartParse = "Fleischner")
        smartParseOptions := "Volume|RVLV|NASCET|Adrenal|Fleischner||"
    else
        smartParseOptions := "Volume||RVLV|NASCET|Adrenal|Fleischner"

    ; Unit options
    unitOptions := (DefaultMeasurementUnit = "mm") ? "cm|mm||" : "cm||mm"

    ; RV/LV format options
    rvlvOptions := (RVLVOutputFormat = "Inline") ? "Macro|Inline||" : "Macro||Inline"

    Gui, SettingsGui:New, +AlwaysOnTop
    Gui, SettingsGui:Add, Text, x10 y10 w280, RadAssist v2.3 Settings
    Gui, SettingsGui:Add, Text, x10 y40, Default Quick Parse:
    Gui, SettingsGui:Add, DropDownList, x130 y37 w120 vSetDefaultParse, %smartParseOptions%

    ; Tool Visibility section (flat checkbox list per user preference)
    Gui, SettingsGui:Add, GroupBox, x10 y65 w270 h125, Tool Visibility (uncheck to hide)
    Gui, SettingsGui:Add, Checkbox, x20 y85 w120 vSetShowVolume %volChecked%, Volume
    Gui, SettingsGui:Add, Checkbox, x150 y85 w120 vSetShowRVLV %rvlvChecked%, RV/LV Ratio
    Gui, SettingsGui:Add, Checkbox, x20 y108 w120 vSetShowNASCET %nascetChecked%, NASCET
    Gui, SettingsGui:Add, Checkbox, x150 y108 w120 vSetShowAdrenal %adrenalChecked%, Adrenal
    Gui, SettingsGui:Add, Checkbox, x20 y131 w120 vSetShowFleischner %fleischnerChecked%, Fleischner
    Gui, SettingsGui:Add, Checkbox, x150 y131 w120 vSetShowStenosis %stenosisChecked%, Stenosis
    Gui, SettingsGui:Add, Checkbox, x20 y154 w120 vSetShowICH %ichChecked%, ICH Volume
    Gui, SettingsGui:Add, Checkbox, x150 y154 w120 vSetShowDate %dateChecked%, Date Calculator

    Gui, SettingsGui:Add, GroupBox, x10 y195 w270 h105, Smart Parse Options
    Gui, SettingsGui:Add, Checkbox, x20 y215 w250 vSetConfirmation %confirmChecked%, Show confirmation dialog before insert
    Gui, SettingsGui:Add, Checkbox, x20 y240 w250 vSetFallbackGUI %fallbackChecked%, Fall back to GUI when confidence low
    Gui, SettingsGui:Add, Text, x20 y268, Default units (no units in text):
    Gui, SettingsGui:Add, DropDownList, x190 y265 w70 vSetDefaultUnit, %unitOptions%

    Gui, SettingsGui:Add, GroupBox, x10 y305 w270 h80, RV/LV & Fleischner Output
    Gui, SettingsGui:Add, Text, x20 y325, RV/LV format:
    Gui, SettingsGui:Add, DropDownList, x100 y322 w80 vSetRVLVFormat, %rvlvOptions%
    Gui, SettingsGui:Add, Text, x185 y325 w90, (Macro = cm)
    Gui, SettingsGui:Add, Checkbox, x20 y350 w250 vSetFleischnerImpression %fleischnerImprChecked%, Insert Fleischner after IMPRESSION:

    Gui, SettingsGui:Add, GroupBox, x10 y390 w270 h105, Output Options
    Gui, SettingsGui:Add, Checkbox, x20 y410 w250 vSetDatamine %dmChecked%, Include datamining phrase by default
    Gui, SettingsGui:Add, Checkbox, x20 y435 w250 vSetCitations %citChecked%, Show citations in output
    Gui, SettingsGui:Add, Text, x20 y460, Datamining phrase:
    Gui, SettingsGui:Add, Edit, x110 y457 w160 vSetDMPhrase, %DataminingPhrase%

    Gui, SettingsGui:Add, Button, x70 y505 w80 gSaveSettings, Save
    Gui, SettingsGui:Add, Button, x160 y505 w80 gSettingsGuiClose, Cancel
    Gui, SettingsGui:Show, x%xPos% y%yPos% w295 h545, Settings
    return
}

SettingsGuiClose:
    Gui, SettingsGui:Destroy
return

SaveSettings:
    global IncludeDatamining, ShowCitations, DataminingPhrase, DefaultSmartParse
    global SmartParseConfirmation, SmartParseFallbackToGUI
    global DefaultMeasurementUnit, RVLVOutputFormat, FleischnerInsertAfterImpression
    global ShowVolumeTools, ShowRVLVTools, ShowNASCETTools
    global ShowAdrenalTools, ShowFleischnerTools, ShowStenosisTools
    global ShowICHTools, ShowDateCalculator
    global PreferencesPath

    Gui, SettingsGui:Submit

    ; Convert checkbox values (variable names must match GUI vVar names)
    IncludeDatamining := SetDatamine
    ShowCitations := SetCitations
    SmartParseConfirmation := SetConfirmation
    SmartParseFallbackToGUI := SetFallbackGUI
    FleischnerInsertAfterImpression := SetFleischnerImpression
    DataminingPhrase := SetDMPhrase
    DefaultSmartParse := SetDefaultParse
    DefaultMeasurementUnit := SetDefaultUnit
    RVLVOutputFormat := SetRVLVFormat

    ; Tool visibility settings
    ShowVolumeTools := SetShowVolume
    ShowRVLVTools := SetShowRVLV
    ShowNASCETTools := SetShowNASCET
    ShowAdrenalTools := SetShowAdrenal
    ShowFleischnerTools := SetShowFleischner
    ShowStenosisTools := SetShowStenosis
    ShowICHTools := SetShowICH
    ShowDateCalculator := SetShowDate

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

    if (!writeSuccess)
        MsgBox, 48, Warning, Some settings may not have saved. Check file permissions or OneDrive sync status.

    Gui, SettingsGui:Destroy
    ToolTip, Settings saved
    SetTimer, RemoveToolTip, -1500
return

; -----------------------------------------
; INI Write with retry for OneDrive sync conflicts
; WHY: OneDrive may lock files during sync, retry helps
; -----------------------------------------
IniWriteWithRetry(key, value, maxRetries := 3) {
    global PreferencesPath
    Loop, %maxRetries%
    {
        IniWrite, %value%, %PreferencesPath%, Settings, %key%
        if (!ErrorLevel)
            return true
        Sleep, 100  ; Wait 100ms before retry
    }
    return false
}

; =========================================
; Result Display
; =========================================
ShowResult(text) {
    ; Copy to clipboard and show message
    Clipboard := text

    ; Create result window
    CoordMode, Mouse, Screen
    MouseGetPos, mouseX, mouseY
    xPos := mouseX + 10
    yPos := mouseY + 10

    Gui, ResultGui:New, +AlwaysOnTop
    Gui, ResultGui:Add, Edit, x10 y10 w380 h200 ReadOnly, %text%
    Gui, ResultGui:Add, Button, x10 y220 w120 gInsertResult, Insert into Report
    Gui, ResultGui:Add, Button, x140 y220 w120 gCopyResult, Copy to Clipboard
    Gui, ResultGui:Add, Button, x270 y220 w120 gResultGuiClose, Close
    Gui, ResultGui:Show, x%xPos% y%yPos% w400 h260, Result
    return
}

ResultGuiClose:
    Gui, ResultGui:Destroy
return

CopyResult:
    ; Already in clipboard from ShowResult
    ToolTip, Copied to clipboard!
    SetTimer, RemoveToolTip, -1500
    Gui, ResultGui:Destroy
return

InsertResult:
    Gui, ResultGui:Destroy
    Sleep, 100
    ; WHY: Deselect any selected text first to prevent replacing it
    ; NOTE: Right arrow moves cursor to end of selection without deleting
    Send, {Right}
    Sleep, 50
    Send, ^v
return

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
    FileAppend, test, %testFile%
    if (ErrorLevel) {
        canWrite := false
    } else {
        FileDelete, %testFile%
    }

    if (canWrite) {
        PreferencesPath := testPath
    } else {
        ; Fall back to LOCALAPPDATA
        fallbackDir := A_AppData . "\..\Local\RadAssist"
        if (!FileExist(fallbackDir)) {
            FileCreateDir, %fallbackDir%
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
    global ShowICHTools, ShowDateCalculator
    global PreferencesPath

    ; Use PreferencesPath set by InitPreferencesPath() for OneDrive compatibility
    if (FileExist(PreferencesPath)) {
        IniRead, IncludeDatamining, %PreferencesPath%, Settings, IncludeDatamining, 1
        IniRead, ShowCitations, %PreferencesPath%, Settings, ShowCitations, 1
        IniRead, DataminingPhrase, %PreferencesPath%, Settings, DataminingPhrase, SSM lung nodule
        IniRead, DefaultSmartParse, %PreferencesPath%, Settings, DefaultSmartParse, Volume
        IniRead, SmartParseConfirmation, %PreferencesPath%, Settings, SmartParseConfirmation, 1
        IniRead, SmartParseFallbackToGUI, %PreferencesPath%, Settings, SmartParseFallbackToGUI, 1
        IniRead, DefaultMeasurementUnit, %PreferencesPath%, Settings, DefaultMeasurementUnit, cm
        IniRead, RVLVOutputFormat, %PreferencesPath%, Settings, RVLVOutputFormat, Macro
        IniRead, FleischnerInsertAfterImpression, %PreferencesPath%, Settings, FleischnerInsertAfterImpression, 1

        ; Tool visibility preferences (default to true/1)
        IniRead, ShowVolumeTools, %PreferencesPath%, Settings, ShowVolumeTools, 1
        IniRead, ShowRVLVTools, %PreferencesPath%, Settings, ShowRVLVTools, 1
        IniRead, ShowNASCETTools, %PreferencesPath%, Settings, ShowNASCETTools, 1
        IniRead, ShowAdrenalTools, %PreferencesPath%, Settings, ShowAdrenalTools, 1
        IniRead, ShowFleischnerTools, %PreferencesPath%, Settings, ShowFleischnerTools, 1
        IniRead, ShowStenosisTools, %PreferencesPath%, Settings, ShowStenosisTools, 1
        IniRead, ShowICHTools, %PreferencesPath%, Settings, ShowICHTools, 1
        IniRead, ShowDateCalculator, %PreferencesPath%, Settings, ShowDateCalculator, 1

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
Menu, Tray, Tip, RadAssist v2.4 - Smart Radiology Tools
Menu, Tray, NoStandard
Menu, Tray, Add, RadAssist v2.4, TrayAbout
Menu, Tray, Add
Menu, Tray, Add, Settings, MenuSettings
Menu, Tray, Add, Reload, TrayReload
Menu, Tray, Add, Exit, TrayExit
return

TrayAbout:
    MsgBox, 64, RadAssist, RadAssist v2.4`n`nShift+Right-click in PowerScribe or Notepad`nto access radiology calculators.`n`nv2.2 Features:`n- Pause/Resume with backtick key (`)`n- RV/LV macro format for PowerScribe`n- Fleischner inserts after IMPRESSION:`n- OneDrive compatible (auto-fallback path)`n- ASCII chars for encoding compatibility`n`nSmart Parsers:`n- Smart Volume: Organ detection + prostate interpretation`n- Smart RV/LV: PE risk (Macro or Inline format)`n- Smart NASCET: Stenosis severity grading`n- Smart Adrenal: Washout percentages`n- Parse Nodules: Fleischner 2017 recommendations`n`nUtilities:`n- Report Header Template`n- Sectra History Copy (Ctrl+Shift+H)
return

TrayReload:
    Reload
return

TrayExit:
    ExitApp
return
