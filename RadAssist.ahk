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
    global g_SelectedText := ""
    ClipSaved := ClipboardAll
    Clipboard := ""
    Send, ^c
    ClipWait, 0.3
    if (!ErrorLevel)
        g_SelectedText := Clipboard
    Clipboard := ClipSaved

    ; Build and show menu
    Menu, RadAssistMenu, Add
    Menu, RadAssistMenu, DeleteAll

    ; Smart Parse submenu (text parsing with inline insertion)
    Menu, SmartParseMenu, Add
    Menu, SmartParseMenu, DeleteAll
    Menu, SmartParseMenu, Add, Smart Volume (parse dimensions), MenuSmartVolume
    Menu, SmartParseMenu, Add, Smart RV/LV (parse ratio), MenuSmartRVLV
    Menu, SmartParseMenu, Add, Smart NASCET (parse stenosis), MenuSmartNASCET
    Menu, SmartParseMenu, Add, Smart Adrenal (parse HU values), MenuSmartAdrenal
    Menu, SmartParseMenu, Add
    Menu, SmartParseMenu, Add, Parse Nodules (Fleischner), MenuSmartFleischner

    Menu, RadAssistMenu, Add, Smart Parse, :SmartParseMenu
    Menu, RadAssistMenu, Add, Quick Parse (%DefaultSmartParse%), MenuQuickSmartParse
    Menu, RadAssistMenu, Add

    Menu, RadAssistMenu, Add, Ellipsoid Volume (GUI), MenuEllipsoidVolume
    Menu, RadAssistMenu, Add, Adrenal Washout (GUI), MenuAdrenalWashout
    Menu, RadAssistMenu, Add
    Menu, RadAssistMenu, Add, NASCET (Carotid), MenuNASCET
    Menu, RadAssistMenu, Add, Vessel Stenosis (General), MenuStenosis
    Menu, RadAssistMenu, Add, RV/LV Ratio (GUI), MenuRVLV
    Menu, RadAssistMenu, Add
    Menu, RadAssistMenu, Add, Fleischner 2017 (GUI), MenuFleischner
    Menu, RadAssistMenu, Add
    Menu, RadAssistMenu, Add, Insert Report Header, MenuInsertHeader
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

MenuInsertHeader:
    InsertReportHeader()
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

    result := "Ellipsoid Volume Calculation:`n"
    result .= "Dimensions: " . EllipDim1 . " x " . EllipDim2 . " x " . EllipDim3 . " " . EllipUnits . "`n"
    result .= "Volume: " . Round(volume, 2) . " mL (cc)`n"

    ; Also show in mm³ if input was cm
    if (EllipUnits = "cm") {
        volumeMm3 := volume * 1000
        result .= "Volume: " . Round(volumeMm3, 0) . " mm³"
    }

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

    result := "Adrenal Washout Calculation:`n"
    result .= "Pre-contrast: " . pre . " HU`n"
    result .= "Post-contrast: " . post . " HU`n"
    result .= "Delayed (15 min): " . delayed . " HU`n`n"

    ; Absolute washout (requires pre-contrast)
    if (post != pre) {
        absWashout := ((post - delayed) / (post - pre)) * 100
        absWashout := Round(absWashout, 1)
        result .= "Absolute Washout: " . absWashout . "%`n"
        if (absWashout >= 60)
            result .= "  -> >=60%: Likely adenoma`n"
        else
            result .= "  -> <60%: Indeterminate/suspicious`n"
    }

    ; Relative washout (no pre-contrast needed)
    if (post != 0) {
        relWashout := ((post - delayed) / post) * 100
        relWashout := Round(relWashout, 1)
        result .= "`nRelative Washout: " . relWashout . "%`n"
        if (relWashout >= 40)
            result .= "  -> >=40%: Likely adenoma`n"
        else
            result .= "  -> <40%: Indeterminate/suspicious`n"
    }

    ; Pre-contrast density assessment
    result .= "`nPre-contrast Assessment:`n"
    if (pre <= 10)
        result .= "  -> <=10 HU: Lipid-rich adenoma (no washout needed)`n"
    else
        result .= "  -> >10 HU: Lipid-poor, washout analysis required`n"

    if (ShowCitations)
        result .= "`nRef: Mayo-Smith WW et al. Radiology 2017"

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

    result := "NASCET Calculation:`n"
    result .= "Distal ICA: " . distal . " mm`n"
    result .= "Stenosis: " . stenosis . " mm`n"
    result .= "NASCET: " . nascetVal . "%`n`n"

    if (nascetVal < 50) {
        result .= "Interpretation: Mild stenosis (<50%)`n"
        result .= "- Medical management typically recommended."
    } else if (nascetVal < 70) {
        result .= "Interpretation: Moderate stenosis (50-69%)`n"
        result .= "- Consider intervention in symptomatic patients."
    } else {
        result .= "Interpretation: Severe stenosis (>=70%)`n"
        result .= "- Strong indication for CEA/CAS in symptomatic patients."
    }

    if (ShowCitations)
        result .= "`n`nCitation: NASCET (N Engl J Med 1991;325:445-53)"

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

    result := "Stenosis Calculation:`n"
    result .= vesselName . "`n"
    result .= "Normal diameter: " . normal . " mm`n"
    result .= "Stenosis diameter: " . stenosis . " mm`n"
    result .= "Stenosis: " . stenosisPercent . "%`n`n"

    ; Same severity cutoffs as NASCET
    if (stenosisPercent < 50) {
        result .= "Interpretation: Mild stenosis (<50%)"
    } else if (stenosisPercent < 70) {
        result .= "Interpretation: Moderate stenosis (50-69%)"
    } else {
        result .= "Interpretation: Severe stenosis (>=70%)"
    }

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

    result := "RV/LV Ratio Calculation:`n"
    result .= "RV diameter: " . rv . " mm`n"
    result .= "LV diameter: " . lv . " mm`n"
    result .= "RV/LV Ratio: " . ratio . "`n`n"

    if (ratio >= 1.0) {
        result .= "Interpretation: Significant right heart strain`n"
        result .= "- Suggestive of severe PE or chronic pulmonary hypertension`n"
        result .= "- Consider ICU admission and advanced therapies"
    } else if (ratio >= 0.9) {
        result .= "Interpretation: Suggestive of right heart strain`n"
        result .= "- Ratio 0.9-0.99 associated with increased adverse events in acute PE`n"
        result .= "- Close monitoring recommended"
    } else {
        result .= "Interpretation: Normal`n"
        result .= "- RV size normal relative to LV`n"
        result .= "- Low risk for RV dysfunction in setting of PE"
    }

    if (ShowCitations)
        result .= "`n`nRef: Meinel et al. Radiology 2015;275:583-591"

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

    ; Check if IMPRESSION: exists in document
    if (!InStr(docText, "IMPRESSION")) {
        ; IMPRESSION not found - fall back to insert after selection
        Clipboard := ClipSaved  ; Restore clipboard first
        Send, {Right}  ; Deselect and move cursor
        Sleep, 20
        InsertAfterSelection(textToInsert)
        ToolTip, IMPRESSION: not found - inserted at cursor
        SetTimer, RemoveToolTip, -2000
        return
    }

    ; Deselect and go to start
    Send, ^{Home}
    Sleep, 20

    ; Open Find dialog (Ctrl+F)
    Send, ^f
    Sleep, 100

    ; Clear any previous search and type new search text
    Send, ^a
    Sleep, 20
    Send, IMPRESSION:
    Sleep, 100

    ; Press Enter to find (or F3/Find Next depending on app)
    Send, {Enter}
    Sleep, 50

    ; Close Find dialog (Escape) - press twice to ensure closed
    Send, {Escape}
    Sleep, 20
    Send, {Escape}
    Sleep, 20

    ; Go to end of line
    Send, {End}
    Sleep, 20

    ; Add blank line (Enter twice for spacing)
    Send, {Enter}{Enter}
    Sleep, 20

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
; SMART RV/LV RATIO PARSER
; WHY: Parse "RV 42mm / LV 35mm" and insert ratio with PE risk interpretation.
; ARCHITECTURE: Requires explicit RV/LV keywords - no fallback to bare numbers (too risky).
; TRADEOFF: Macro format uses cm for PowerScribe compatibility; Inline uses mm.
; -----------------------------------------
ParseAndInsertRVLV(input) {
    global ShowCitations, RVLVOutputFormat
    input := RegExReplace(input, "`r?\n", " ")
    filtered := FilterNonMeasurements(input)

    rv := 0
    lv := 0
    confidence := 0

    ; Pattern A (HIGH confidence 90): "RV X mm ... LV Y mm" or "RV: Xmm, LV: Ymm"
    ; Example: "RV 42mm / LV 35mm", "RV: 42, LV: 35"
    if (RegExMatch(filtered, "i)RV\s*(?:diameter|diam)?[:\s=]*(\d+(?:\.\d+)?)\s*(?:mm|cm)?.*?LV\s*(?:diameter|diam)?[:\s=]*(\d+(?:\.\d+)?)\s*(?:mm|cm)?", m)) {
        rv := m1 + 0
        lv := m2 + 0
        confidence := 90
    }
    ; Pattern B (HIGH confidence 85): "RV/LV ratio = X / Y"
    else if (RegExMatch(filtered, "i)RV\s*/\s*LV\s*(?:ratio)?\s*[=:]\s*(\d+(?:\.\d+)?)\s*/\s*(\d+(?:\.\d+)?)", m)) {
        rv := m1 + 0
        lv := m2 + 0
        confidence := 85
    }
    ; NO FALLBACK TO JUST TWO NUMBERS - too risky (could be image numbers, dates, etc.)
    else {
        MsgBox, 48, Parse Error, Could not find RV/LV measurements.`n`nExpected formats:`n- "RV 42mm / LV 35mm"`n- "RV: 42, LV: 35"`n`nMust include "RV" and "LV" keywords.
        return
    }

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

    ; Build result based on output format setting
    resultText := ""
    if (RVLVOutputFormat = "Macro") {
        ; Macro format for PowerScribe "Macro right heart" - convert mm to cm
        rv_cm := Round(rv / 10, 1)
        lv_cm := Round(lv / 10, 1)
        resultText := "`n`nEVALUATION FOR RIGHT HEART STRAIN:`n"
        resultText .= "Standard axial measurements demonstrate:`n"
        resultText .= "Right Ventricle: " . rv_cm . " cm`n"
        resultText .= "Left Ventricle: " . lv_cm . " cm`n"
        resultText .= "RV/LV Ratio: " . ratio . "`n"
        resultText .= "Impression: " . interpretation
    } else {
        ; Inline format for continued dictation (original behavior)
        resultText := " RV/LV ratio: " . ratio . " (RV " . Round(rv, 1) . "mm / LV " . Round(lv, 1) . "mm), "
        StringLower, interpretLower, interpretation
        resultText .= interpretLower
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
    confidence := 0

    ; Pattern A (HIGH confidence 90): "distal X mm ... stenosis Y mm"
    if (RegExMatch(filtered, "i)distal\s*(?:ICA|internal\s*carotid)?.*?(\d+(?:\.\d+)?)\s*(?:mm|cm).*?stenosis.*?(\d+(?:\.\d+)?)\s*(?:mm|cm)?", m)) {
        distal := m1 + 0
        stenosis := m2 + 0
        confidence := 90
    }
    ; Pattern B (HIGH confidence 85): "stenosis X mm ... distal Y mm" (reverse order)
    else if (RegExMatch(filtered, "i)stenosis.*?(\d+(?:\.\d+)?)\s*(?:mm|cm).*?distal\s*(?:ICA)?.*?(\d+(?:\.\d+)?)\s*(?:mm|cm)?", m)) {
        stenosis := m1 + 0
        distal := m2 + 0
        confidence := 85
    }
    ; Pattern C (MEDIUM confidence 65): ICA/carotid + narrowing context
    else if (RegExMatch(filtered, "i)(?:ICA|carotid).*?(\d+(?:\.\d+)?)\s*(?:mm|cm).*?(?:narrow|residual).*?(\d+(?:\.\d+)?)\s*(?:mm|cm)?", m)) {
        d1 := m1 + 0
        d2 := m2 + 0
        ; Larger is distal, smaller is stenosis
        if (d1 > d2) {
            distal := d1
            stenosis := d2
        } else {
            distal := d2
            stenosis := d1
        }
        confidence := 65
    }
    ; NO FALLBACK TO JUST TWO NUMBERS - too risky
    else {
        MsgBox, 48, Parse Error, Could not find stenosis measurements.`n`nExpected formats:`n- "distal 5.2mm, stenosis 2.1mm"`n- "distal ICA 5.2mm, stenosis 2.1mm"`n`nMust include "distal" or "stenosis" keywords.
        return
    }

    ; Validate
    if (distal <= 0) {
        MsgBox, 48, Error, Distal ICA diameter must be greater than 0.
        return
    }
    if (stenosis >= distal) {
        MsgBox, 48, Error, Stenosis diameter must be less than distal diameter.
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
; SMART FLEISCHNER NODULE PARSER
; WHY: Parse findings text for nodules and generate Fleischner 2017 recommendations.
; ARCHITECTURE: Multi-pattern regex, requires "nodule" keyword, confirmation before insert.
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

    ; Pattern 1 (HIGH confidence): Size + type + "nodule" keyword
    ; Example: "8 mm solid nodule"
    pattern1 := "i)(\d+(?:\.\d+)?)\s*(mm|cm)\s*(solid|part-?solid|ground[- ]?glass|subsolid|GGN|GGO|SSN)?\s*(?:nodule|opacity|lesion|nodular)"

    ; Pattern 2 (HIGH confidence): Type + "nodule" + size
    ; Example: "solid nodule measuring 8 mm"
    pattern2 := "i)(solid|part-?solid|ground[- ]?glass|subsolid|GGN|GGO|SSN)\s*(?:nodule|opacity|lesion|nodular).*?(\d+(?:\.\d+)?)\s*(mm|cm)"

    ; Search with Pattern 1
    pos := 1
    while (pos := RegExMatch(filtered, pattern1, m, pos)) {
        size := m1 + 0
        units := m2
        nodeType := m3

        ; Convert to mm if cm
        if (units = "cm")
            size := size * 10

        ; Only accept plausible nodule sizes (1-50mm)
        if (size >= 1 && size <= 50) {
            ; Classify nodule type
            if (nodeType = "" || (InStr(nodeType, "solid") && !InStr(nodeType, "part") && !InStr(nodeType, "sub"))) {
                solidNodules.Push(size)
            } else if (InStr(nodeType, "part")) {
                partsolidNodules.Push(size)
            } else {
                subsolidNodules.Push(size)
            }
        }

        pos += StrLen(m)
    }

    ; Search with Pattern 2
    pos := 1
    while (pos := RegExMatch(filtered, pattern2, m, pos)) {
        nodeType := m1
        size := m2 + 0
        units := m3

        if (units = "cm")
            size := size * 10

        ; Only accept plausible sizes
        if (size >= 1 && size <= 50) {
            if (InStr(nodeType, "solid") && !InStr(nodeType, "part") && !InStr(nodeType, "sub")) {
                solidNodules.Push(size)
            } else if (InStr(nodeType, "part")) {
                partsolidNodules.Push(size)
            } else {
                subsolidNodules.Push(size)
            }
        }

        pos += StrLen(m)
    }

    ; Check for "sub 6 mm nodules" or "punctate nodules" (multiple small)
    if (InStr(filtered, "sub 6") || InStr(filtered, "sub-6") || InStr(filtered, "punctate")) {
        if (solidNodules.Length() = 0)
            solidNodules.Push(5)
    }

    ; If no nodules found, show error
    if (solidNodules.Length() = 0 && subsolidNodules.Length() = 0 && partsolidNodules.Length() = 0) {
        MsgBox, 48, Parse Error, Could not find nodule descriptions in text.`n`nExpected patterns:`n- "8 mm solid nodule"`n- "ground glass nodule measuring 6 mm"`n`nMust include "nodule" keyword with size.
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

; -----------------------------------------
; REPORT HEADER TEMPLATE
; WHY: Insert standardized report header with placeholder for Sectra paste.
; -----------------------------------------
InsertReportHeader() {
    header := "PROCEDURE:  `n"
    header .= "COMPARISON: No relevant available comparison.`n"
    header .= "CLINICAL INFORMATION (none relevant/not provided if blank):`n"
    header .= "Indication:    `n"
    header .= "Additional History:    `n`n"

    ; Save clipboard and paste header
    ClipSaved := ClipboardAll
    Clipboard := header
    ClipWait, 0.5

    Send, ^v
    Sleep, 100

    Clipboard := ClipSaved

    ToolTip, Header inserted!
    SetTimer, RemoveToolTip, -1500
}

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

    CoordMode, Mouse, Screen
    MouseGetPos, mouseX, mouseY
    xPos := mouseX + 10
    yPos := mouseY + 10

    dmChecked := IncludeDatamining ? "Checked" : ""
    citChecked := ShowCitations ? "Checked" : ""
    confirmChecked := SmartParseConfirmation ? "Checked" : ""
    fallbackChecked := SmartParseFallbackToGUI ? "Checked" : ""
    fleischnerImprChecked := FleischnerInsertAfterImpression ? "Checked" : ""

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
    Gui, SettingsGui:Add, Text, x10 y10 w280, RadAssist v2.2 Settings
    Gui, SettingsGui:Add, Text, x10 y40, Default Quick Parse:
    Gui, SettingsGui:Add, DropDownList, x130 y37 w120 vSetDefaultParse, %smartParseOptions%
    Gui, SettingsGui:Add, GroupBox, x10 y65 w270 h105, Smart Parse Options
    Gui, SettingsGui:Add, Checkbox, x20 y85 w250 vSetConfirmation %confirmChecked%, Show confirmation dialog before insert
    Gui, SettingsGui:Add, Checkbox, x20 y110 w250 vSetFallbackGUI %fallbackChecked%, Fall back to GUI when confidence low
    Gui, SettingsGui:Add, Text, x20 y138, Default units (no units in text):
    Gui, SettingsGui:Add, DropDownList, x190 y135 w70 vSetDefaultUnit, %unitOptions%
    Gui, SettingsGui:Add, GroupBox, x10 y175 w270 h80, RV/LV & Fleischner Output
    Gui, SettingsGui:Add, Text, x20 y195, RV/LV format:
    Gui, SettingsGui:Add, DropDownList, x100 y192 w80 vSetRVLVFormat, %rvlvOptions%
    Gui, SettingsGui:Add, Text, x185 y195 w90, (Macro = cm)
    Gui, SettingsGui:Add, Checkbox, x20 y220 w250 vSetFleischnerImpression %fleischnerImprChecked%, Insert Fleischner after IMPRESSION:
    Gui, SettingsGui:Add, GroupBox, x10 y260 w270 h105, Output Options
    Gui, SettingsGui:Add, Checkbox, x20 y280 w250 vSetDatamine %dmChecked%, Include datamining phrase by default
    Gui, SettingsGui:Add, Checkbox, x20 y305 w250 vSetCitations %citChecked%, Show citations in output
    Gui, SettingsGui:Add, Text, x20 y330, Datamining phrase:
    Gui, SettingsGui:Add, Edit, x110 y327 w160 vSetDMPhrase, %DataminingPhrase%
    Gui, SettingsGui:Add, Button, x70 y375 w80 gSaveSettings, Save
    Gui, SettingsGui:Add, Button, x160 y375 w80 gSettingsGuiClose, Cancel
    Gui, SettingsGui:Show, x%xPos% y%yPos% w295 h415, Settings
    return
}

SettingsGuiClose:
    Gui, SettingsGui:Destroy
return

SaveSettings:
    global IncludeDatamining, ShowCitations, DataminingPhrase, DefaultSmartParse
    global SmartParseConfirmation, SmartParseFallbackToGUI
    global DefaultMeasurementUnit, RVLVOutputFormat, FleischnerInsertAfterImpression
    global PreferencesPath

    Gui, SettingsGui:Submit

    ; Convert checkbox values
    IncludeDatamining := SetDatamining
    ShowCitations := SetCitations
    SmartParseConfirmation := SetSmartParseConfirm
    SmartParseFallbackToGUI := SetFallbackToGUI
    FleischnerInsertAfterImpression := SetFleischnerImpression

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

        IncludeDatamining := (IncludeDatamining = "1")
        ShowCitations := (ShowCitations = "1")
        SmartParseConfirmation := (SmartParseConfirmation = "1")
        SmartParseFallbackToGUI := (SmartParseFallbackToGUI = "1")
        FleischnerInsertAfterImpression := (FleischnerInsertAfterImpression = "1")
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
Menu, Tray, Tip, RadAssist v2.2 - Smart Radiology Tools
Menu, Tray, NoStandard
Menu, Tray, Add, RadAssist v2.2, TrayAbout
Menu, Tray, Add
Menu, Tray, Add, Settings, MenuSettings
Menu, Tray, Add, Reload, TrayReload
Menu, Tray, Add, Exit, TrayExit
return

TrayAbout:
    MsgBox, 64, RadAssist, RadAssist v2.2`n`nShift+Right-click in PowerScribe or Notepad`nto access radiology calculators.`n`nv2.2 Features:`n- Pause/Resume with backtick key (`)`n- RV/LV macro format for PowerScribe`n- Fleischner inserts after IMPRESSION:`n- OneDrive compatible (auto-fallback path)`n- ASCII chars for encoding compatibility`n`nSmart Parsers:`n- Smart Volume: Organ detection + prostate interpretation`n- Smart RV/LV: PE risk (Macro or Inline format)`n- Smart NASCET: Stenosis severity grading`n- Smart Adrenal: Washout percentages`n- Parse Nodules: Fleischner 2017 recommendations`n`nUtilities:`n- Report Header Template`n- Sectra History Copy (Ctrl+Shift+H)
return

TrayReload:
    Reload
return

TrayExit:
    ExitApp
return
