; ==========================================
; RadAssist - Radiology Assistant Tool
; Version: 2.0
; Description: Lean radiology workflow tool with calculators
;              Triggered by Shift+Right-click in PowerScribe/Notepad
;              Now with intelligent text parsing for inline calculations
; ==========================================

#NoEnv
#SingleInstance, Force
SetWorkingDir, %A_ScriptDir%

; -----------------------------------------
; Global Configuration
; -----------------------------------------
global TargetApps := ["ahk_exe Nuance.PowerScribe360.exe", "ahk_exe notepad.exe", "ahk_class Notepad"]
global SectraWindowTitle := "Sectra"
global DataminingPhrase := "SSM lung nodule"
global IncludeDatamining := true
global ShowCitations := true

; -----------------------------------------
; Global Hotkey: Ctrl+Shift+H - Sectra History Copy
; (Can be mapped to Contour ShuttlePRO button)
; -----------------------------------------
^+h::
    CopySectraHistory()
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
; Menu Handlers
; -----------------------------------------

; Smart Parse handlers (inline text parsing with insertion)
MenuSmartVolume:
    if (g_SelectedText = "") {
        MsgBox, 48, No Selection, Please select text containing dimensions (e.g., "Prostate measures 8.0 x 6.0 x 9.0 cm")
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
    CoordMode, Mouse, Screen
    MouseGetPos, mouseX, mouseY
    xPos := mouseX + 10
    yPos := mouseY + 10

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
            result .= "  -> ≥60%: Likely adenoma`n"
        else
            result .= "  -> <60%: Indeterminate/suspicious`n"
    }

    ; Relative washout (no pre-contrast needed)
    if (post != 0) {
        relWashout := ((post - delayed) / post) * 100
        relWashout := Round(relWashout, 1)
        result .= "`nRelative Washout: " . relWashout . "%`n"
        if (relWashout >= 40)
            result .= "  -> ≥40%: Likely adenoma`n"
        else
            result .= "  -> <40%: Indeterminate/suspicious`n"
    }

    ; Pre-contrast density assessment
    result .= "`nPre-contrast Assessment:`n"
    if (pre <= 10)
        result .= "  -> ≤10 HU: Lipid-rich adenoma (no washout needed)`n"
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
        result .= "Interpretation: Severe stenosis (≥70%)`n"
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

; =========================================
; CALCULATOR 3B: General Vessel Stenosis
; (Same cutoffs as NASCET, for any vessel)
; =========================================
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
        result .= "Interpretation: Severe stenosis (≥70%)"
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
                return "CT at 3-6 months. If stable, annual CT for 5 years.`nIf solid component ≥6 mm, consider PET/CT or biopsy."
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
; SMART PARSE FUNCTIONS
; WHY: Parse dictated text and insert calculations inline without GUI re-entry.
; ARCHITECTURE: Text parsing with regex, calculation, and clipboard-based insertion.
; =========================================

; -----------------------------------------
; Utility: Insert text after current selection
; WHY: Preserves original text and appends calculation result.
; -----------------------------------------
InsertAfterSelection(textToInsert) {
    ; Move cursor to end of selection and insert text
    Send, {Right}
    Sleep, 50

    ; Save current clipboard
    ClipSaved := ClipboardAll
    Clipboard := textToInsert
    ClipWait, 0.5

    ; Paste the text
    Send, ^v
    Sleep, 100

    ; Restore clipboard
    Clipboard := ClipSaved

    ToolTip, Calculation inserted!
    SetTimer, RemoveToolTip, -1500
}

; -----------------------------------------
; SMART VOLUME PARSER
; WHY: Parse "Prostate measures 8.0 x 6.0 x 9.0 cm" and insert volume with organ-specific interpretation.
; TRADEOFF: Complex regex for flexibility vs potential false positives.
; -----------------------------------------
ParseAndInsertVolume(input) {
    ; Clean input
    input := RegExReplace(input, "`r?\n", " ")

    ; Pattern: Detect organ keyword and dimensions
    ; Matches: "Prostate measures 8.0 x 6.0 x 9.0 cm" or just "8.0 x 6.0 x 9.0 cm"
    organPattern := "i)(prostate|abscess|lesion|cyst|mass|nodule|collection|hematoma|liver|spleen|kidney|bladder|uterus|ovary|thyroid|adrenal)?"
    dimPattern := "(\d+(?:\.\d+)?)\s*[x×,]\s*(\d+(?:\.\d+)?)\s*[x×,]\s*(\d+(?:\.\d+)?)\s*(cm|mm)?"

    fullPattern := organPattern . ".*?" . dimPattern

    if (!RegExMatch(input, fullPattern, m)) {
        ; Try simpler pattern: just dimensions
        if (!RegExMatch(input, dimPattern, m)) {
            MsgBox, 48, Parse Error, Could not parse dimensions from text.`n`nExpected format: "8.0 x 6.0 x 9.0 cm" or "Prostate measures 8.0 x 6.0 x 9.0 cm"
            return
        }
        organ := ""
        d1 := m1
        d2 := m2
        d3 := m3
        units := m4
    } else {
        organ := m1
        d1 := m2
        d2 := m3
        d3 := m4
        units := m5
    }

    ; Convert to numbers
    d1 := d1 + 0
    d2 := d2 + 0
    d3 := d3 + 0

    ; Default to cm if no unit specified, handle mm conversion
    if (units = "" || units = "cm") {
        ; Already in cm or assume cm
    } else if (units = "mm") {
        d1 := d1 / 10
        d2 := d2 / 10
        d3 := d3 / 10
    }

    ; Calculate ellipsoid volume: (π/6) × L × W × H
    volume := 0.5236 * d1 * d2 * d3
    volumeRounded := Round(volume, 1)

    ; Build result string
    result := " This corresponds to a volume of " . volumeRounded . " cc (mL)"

    ; Add organ-specific interpretation if prostate detected
    if (organ != "" && (organ = "prostate" || organ = "Prostate")) {
        if (volume < 30) {
            result .= ", within normal limits."
        } else if (volume < 50) {
            result .= ", compatible with an enlarged prostate."
        } else if (volume < 70) {
            result .= ", compatible with a moderately enlarged prostate."
        } else {
            result .= ", compatible with a massively enlarged prostate."
        }
    } else if (organ != "") {
        ; Generic organ volume
        result .= "."
    } else {
        result .= "."
    }

    ; Insert after selection
    InsertAfterSelection(result)
}

; -----------------------------------------
; SMART RV/LV RATIO PARSER
; WHY: Parse "RV 42mm / LV 35mm" and insert ratio with PE risk interpretation.
; -----------------------------------------
ParseAndInsertRVLV(input) {
    global ShowCitations
    input := RegExReplace(input, "`r?\n", " ")

    rv := 0
    lv := 0

    ; Pattern 1: "RV/LV ratio = X / Y" or "RV/LV = X/Y"
    if (RegExMatch(input, "i)RV\s*[/:]?\s*LV\s*(?:ratio)?\s*[=:]\s*(\d+(?:\.\d+)?)\s*(?:cm|mm)?\s*/\s*(\d+(?:\.\d+)?)\s*(?:cm|mm)?", m)) {
        rv := m1 + 0
        lv := m2 + 0
    }
    ; Pattern 2: "RV X mm ... LV Y mm" or "RV: Xmm, LV: Ymm"
    else if (RegExMatch(input, "i)RV\s*(?:diameter|diam)?[:\s]*(\d+(?:\.\d+)?)\s*(?:mm|cm)?.*?LV\s*(?:diameter|diam)?[:\s]*(\d+(?:\.\d+)?)\s*(?:mm|cm)?", m)) {
        rv := m1 + 0
        lv := m2 + 0
    }
    ; Pattern 3: Just two numbers separated by / or comma (assume RV first, LV second)
    else if (RegExMatch(input, "(\d+(?:\.\d+)?)\s*(?:cm|mm)?\s*[/,]\s*(\d+(?:\.\d+)?)\s*(?:cm|mm)?", m)) {
        rv := m1 + 0
        lv := m2 + 0
    }
    else {
        MsgBox, 48, Parse Error, Could not parse RV/LV measurements.`n`nExpected formats:`n- "RV 42mm / LV 35mm"`n- "RV/LV = 42/35"`n- "RV: 42, LV: 35"
        return
    }

    if (lv <= 0) {
        MsgBox, 48, Error, LV diameter must be greater than 0.
        return
    }

    ; Calculate ratio
    ratio := rv / lv
    ratio := Round(ratio, 2)

    ; Build result
    result := "`nRV/LV Ratio: " . ratio . " (RV " . rv . "mm / LV " . lv . "mm). "

    ; Interpretation
    if (ratio >= 1.0) {
        result .= "Interpretation: Significant right heart strain, suggestive of severe PE."
    } else if (ratio >= 0.9) {
        result .= "Interpretation: Suggestive of right heart strain (ratio 0.9-0.99)."
    } else {
        result .= "Interpretation: Normal RV/LV ratio, low risk for RV dysfunction."
    }

    if (ShowCitations)
        result .= " [Ref: Meinel et al. Radiology 2015;275:583-591]"

    InsertAfterSelection(result)
}

; -----------------------------------------
; SMART NASCET PARSER
; WHY: Parse stenosis measurements and insert NASCET percentage inline.
; -----------------------------------------
ParseAndInsertNASCET(input) {
    global ShowCitations
    input := RegExReplace(input, "`r?\n", " ")

    distal := 0
    stenosis := 0

    ; Pattern 1: "distal X mm ... stenosis Y mm"
    if (RegExMatch(input, "i)distal.*?(\d+(?:\.\d+)?)\s*(?:mm|cm).*?stenosis.*?(\d+(?:\.\d+)?)\s*(?:mm|cm)", m)) {
        distal := m1 + 0
        stenosis := m2 + 0
    }
    ; Pattern 2: "stenosis X mm ... distal Y mm"
    else if (RegExMatch(input, "i)stenosis.*?(\d+(?:\.\d+)?)\s*(?:mm|cm).*?distal.*?(\d+(?:\.\d+)?)\s*(?:mm|cm)", m)) {
        distal := m2 + 0
        stenosis := m1 + 0
    }
    ; Pattern 3: Just two numbers - larger is distal
    else {
        numbers := []
        pos := 1
        while (pos := RegExMatch(input, "(\d+(?:\.\d+)?)\s*(?:mm|cm)?", m, pos)) {
            numbers.Push(m1 + 0)
            pos += StrLen(m)
        }
        if (numbers.Length() >= 2) {
            distal := Max(numbers*)
            stenosis := Min(numbers*)
        } else {
            MsgBox, 48, Parse Error, Could not parse stenosis measurements.`n`nExpected formats:`n- "distal 5.2mm, stenosis 2.1mm"`n- "5.2 / 2.1 mm"
            return
        }
    }

    if (distal <= 0) {
        MsgBox, 48, Error, Distal ICA diameter must be greater than 0.
        return
    }

    ; Calculate NASCET
    nascetVal := ((distal - stenosis) / distal) * 100
    nascetVal := Round(nascetVal, 1)

    ; Build result
    result := "`nNASCET: " . nascetVal . "% stenosis (distal " . distal . "mm, stenosis " . stenosis . "mm). "

    ; Interpretation
    if (nascetVal < 50) {
        result .= "Mild stenosis (<50%)."
    } else if (nascetVal < 70) {
        result .= "Moderate stenosis (50-69%), consider intervention in symptomatic patients."
    } else {
        result .= "Severe stenosis (≥70%), strong indication for CEA/CAS."
    }

    if (ShowCitations)
        result .= " [NASCET: N Engl J Med 1991;325:445-53]"

    InsertAfterSelection(result)
}

; -----------------------------------------
; SMART ADRENAL WASHOUT PARSER
; WHY: Parse HU values and insert washout percentages inline.
; NOTE: Removed timing from output per user request.
; -----------------------------------------
ParseAndInsertAdrenalWashout(input) {
    global ShowCitations
    input := RegExReplace(input, "`r?\n", " ")

    pre := ""
    post := ""
    delayed := ""

    ; Pattern 1: Full pattern with pre, post, delayed
    fullPattern := "i)(?:pre-?contrast|unenhanced|baseline|native|non-?con)[:\s]*(-?\d+)\s*(?:HU)?.*?(?:post-?contrast|enhanced|arterial|portal)[:\s]*(-?\d+)\s*(?:HU)?.*?(?:delayed|15\s*min|late)[:\s]*(-?\d+)\s*(?:HU)?"

    if (RegExMatch(input, fullPattern, m)) {
        pre := m1 + 0
        post := m2 + 0
        delayed := m3 + 0
    }
    ; Pattern 2: Just post and delayed (for relative washout only)
    else {
        fallbackPattern := "i)(?:post-?contrast|enhanced|arterial|portal)[:\s]*(-?\d+)\s*(?:HU)?.*?(?:delayed|15\s*min|late)[:\s]*(-?\d+)\s*(?:HU)?"
        if (RegExMatch(input, fallbackPattern, m)) {
            post := m1 + 0
            delayed := m2 + 0
        } else {
            ; Try to find any three numbers
            numbers := []
            pos := 1
            while (pos := RegExMatch(input, "(-?\d+)\s*(?:HU)?", m, pos)) {
                numbers.Push(m1 + 0)
                pos += StrLen(m)
            }
            if (numbers.Length() >= 3) {
                pre := numbers[1]
                post := numbers[2]
                delayed := numbers[3]
            } else if (numbers.Length() >= 2) {
                post := numbers[1]
                delayed := numbers[2]
            } else {
                MsgBox, 48, Parse Error, Could not parse HU values.`n`nExpected formats:`n- "pre-contrast 10 HU, enhanced 80 HU, delayed 40 HU"`n- "10, 80, 40 HU"
                return
            }
        }
    }

    if (post = "" || delayed = "") {
        MsgBox, 48, Error, At least post-contrast and delayed HU values are required.
        return
    }

    ; Build result
    result := "`nAdrenal washout analysis: "

    ; Calculate absolute washout if pre is available
    if (pre != "" && post != pre) {
        absWashout := ((post - delayed) / (post - pre)) * 100
        absWashout := Round(absWashout, 1)

        result .= "Absolute washout: " . absWashout . "%"
        if (absWashout >= 60)
            result .= " (≥60%: likely adenoma)"
        else
            result .= " (<60%: indeterminate)"
        result .= ". "
    }

    ; Calculate relative washout
    if (post != 0) {
        relWashout := ((post - delayed) / post) * 100
        relWashout := Round(relWashout, 1)

        result .= "Relative washout: " . relWashout . "%"
        if (relWashout >= 40)
            result .= " (≥40%: likely adenoma)"
        else
            result .= " (<40%: indeterminate)"
        result .= ". "
    }

    ; Pre-contrast assessment
    if (pre != "" && pre <= 10) {
        result .= "Pre-contrast " . pre . " HU (≤10 HU: lipid-rich adenoma)."
    }

    if (ShowCitations)
        result .= " [Ref: Mayo-Smith WW et al. Radiology 2017]"

    InsertAfterSelection(result)
}

; -----------------------------------------
; SMART FLEISCHNER NODULE PARSER
; WHY: Parse findings text for nodules and generate Fleischner 2017 recommendations.
; ARCHITECTURE: Multi-pattern regex to find all nodule mentions, classify, and apply algorithm.
; -----------------------------------------
ParseAndInsertFleischner(input) {
    global ShowCitations, DataminingPhrase, IncludeDatamining
    input := RegExReplace(input, "`r?\n", " ")

    ; Arrays to store found nodules
    solidNodules := []
    subsolidNodules := []
    partsolidNodules := []

    ; Pattern 1: Size before type - "8 mm solid nodule"
    pattern1 := "i)(\d+(?:\.\d+)?)\s*(mm|cm)\s*(solid|part-?solid|ground[- ]?glass|subsolid|GGN|GGO|SSN)?\s*(?:nodule|opacity|lesion|nodular)"

    ; Pattern 2: Type before size - "solid nodule measuring 8 mm"
    pattern2 := "i)(solid|part-?solid|ground[- ]?glass|subsolid|GGN|GGO|SSN)\s*(?:nodule|opacity|lesion|nodular).*?(\d+(?:\.\d+)?)\s*(mm|cm)"

    ; Pattern 3: Generic "X mm nodule" (assume solid if not specified)
    pattern3 := "i)(\d+(?:\.\d+)?)\s*(mm|cm)\s*(?:pulmonary\s+)?(?:nodule|opacity)"

    ; Search with Pattern 1
    pos := 1
    while (pos := RegExMatch(input, pattern1, m, pos)) {
        size := m1 + 0
        units := m2
        nodeType := m3

        ; Convert to mm if cm
        if (units = "cm")
            size := size * 10

        ; Classify nodule type
        if (nodeType = "" || nodeType = "solid" || nodeType = "Solid") {
            solidNodules.Push(size)
        } else if (InStr(nodeType, "part") || InStr(nodeType, "Part")) {
            partsolidNodules.Push(size)
        } else {
            subsolidNodules.Push(size)
        }

        pos += StrLen(m)
    }

    ; Search with Pattern 2
    pos := 1
    while (pos := RegExMatch(input, pattern2, m, pos)) {
        nodeType := m1
        size := m2 + 0
        units := m3

        if (units = "cm")
            size := size * 10

        if (InStr(nodeType, "solid") && !InStr(nodeType, "part") && !InStr(nodeType, "sub")) {
            solidNodules.Push(size)
        } else if (InStr(nodeType, "part")) {
            partsolidNodules.Push(size)
        } else {
            subsolidNodules.Push(size)
        }

        pos += StrLen(m)
    }

    ; Check for "sub 6 mm nodules" or "punctate nodules" (multiple small)
    if (InStr(input, "sub 6") || InStr(input, "sub-6") || InStr(input, "punctate") || InStr(input, "scattered") || InStr(input, "multiple")) {
        ; Mark as having multiple small nodules
        if (solidNodules.Length() = 0)
            solidNodules.Push(5)  ; Placeholder for <6mm
    }

    ; If no nodules found, show error
    if (solidNodules.Length() = 0 && subsolidNodules.Length() = 0 && partsolidNodules.Length() = 0) {
        MsgBox, 48, Parse Error, Could not find nodule descriptions in text.`n`nExpected patterns:`n- "8 mm solid nodule"`n- "ground glass nodule measuring 6 mm"`n- "6 mm subsolid nodule"
        return
    }

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
    totalNodules := solidNodules.Length() + subsolidNodules.Length() + partsolidNodules.Length()
    isMultiple := (totalNodules > 1) ? true : false

    ; Build recommendation block
    result := "`n`n___________________________________________________________`n"
    result .= "Incidental Lung Nodule Discussion:`n"
    result .= "2017 ACR Fleischner Society expert consensus recommendations on incidental pulmonary nodules.`n"
    result .= "Target Demographic: 35+ years without known cancer or immunosuppressive disorder.`n`n"

    ; List detected nodules
    result .= "Detected nodules:`n"
    noduleNum := 1

    if (maxSolid > 0) {
        result .= noduleNum . ". " . maxSolid . " mm solid nodule"
        if (solidNodules.Length() > 1)
            result .= " (largest of " . solidNodules.Length() . ")"
        result .= "`n"
        noduleNum++
    }
    if (maxSubsolid > 0) {
        result .= noduleNum . ". " . maxSubsolid . " mm ground glass/subsolid nodule"
        if (subsolidNodules.Length() > 1)
            result .= " (largest of " . subsolidNodules.Length() . ")"
        result .= "`n"
        noduleNum++
    }
    if (maxPartsolid > 0) {
        result .= noduleNum . ". " . maxPartsolid . " mm part-solid nodule"
        if (partsolidNodules.Length() > 1)
            result .= " (largest of " . partsolidNodules.Length() . ")"
        result .= "`n"
        noduleNum++
    }

    result .= "`n"

    ; Generate recommendations for LOW RISK
    result .= "LOW RISK recommendation:`n"
    result .= GenerateFleischnerRec(maxSolid, maxSubsolid, maxPartsolid, isMultiple, "Low risk")
    result .= "`n`n"

    ; Generate recommendations for HIGH RISK
    result .= "HIGH RISK recommendation:`n"
    result .= GenerateFleischnerRec(maxSolid, maxSubsolid, maxPartsolid, isMultiple, "High risk")

    ; Add datamining phrase
    if (IncludeDatamining) {
        result .= "`n`nData Mining: " . DataminingPhrase . " . ACR and AMA MIPS #364"
    }

    ; Add additional info
    result .= "`n`nNote: The need for followup depends on clinical discussion of patient's comorbid conditions, demographics, and willingness to undergo followup imaging and potential intervention."

    if (ShowCitations)
        result .= "`nRef: MacMahon H et al. Radiology 2017;284:228-243"

    result .= "`n___________________________________________________________"

    InsertAfterSelection(result)
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
                result .= "- Part-solid " . maxPartsolid . "mm: CT at 3-6 months. If stable, annual CT for 5 years. If solid component ≥6mm, consider PET/CT or biopsy."
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
    global IncludeDatamining, ShowCitations, DataminingPhrase

    CoordMode, Mouse, Screen
    MouseGetPos, mouseX, mouseY
    xPos := mouseX + 10
    yPos := mouseY + 10

    dmChecked := IncludeDatamining ? "Checked" : ""
    citChecked := ShowCitations ? "Checked" : ""

    Gui, SettingsGui:New, +AlwaysOnTop
    Gui, SettingsGui:Add, Text, x10 y10 w250, RadAssist Settings
    Gui, SettingsGui:Add, Checkbox, x10 y40 w250 vSetDatamine %dmChecked%, Include datamining phrase by default
    Gui, SettingsGui:Add, Checkbox, x10 y65 w250 vSetCitations %citChecked%, Show citations in output
    Gui, SettingsGui:Add, Text, x10 y95, Datamining phrase:
    Gui, SettingsGui:Add, Edit, x10 y115 w200 vSetDMPhrase, %DataminingPhrase%
    Gui, SettingsGui:Add, Button, x10 y150 w80 gSaveSettings, Save
    Gui, SettingsGui:Add, Button, x100 y150 w80 gSettingsGuiClose, Cancel
    Gui, SettingsGui:Show, x%xPos% y%yPos% w230 h190, Settings
    return
}

SettingsGuiClose:
    Gui, SettingsGui:Destroy
return

SaveSettings:
    Gui, SettingsGui:Submit, NoHide
    global IncludeDatamining, ShowCitations, DataminingPhrase

    IncludeDatamining := SetDatamine
    ShowCitations := SetCitations
    DataminingPhrase := SetDMPhrase

    ; Save to INI file
    IniWrite, %IncludeDatamining%, %A_ScriptDir%\RadAssist_preferences.ini, Settings, IncludeDatamining
    IniWrite, %ShowCitations%, %A_ScriptDir%\RadAssist_preferences.ini, Settings, ShowCitations
    IniWrite, %DataminingPhrase%, %A_ScriptDir%\RadAssist_preferences.ini, Settings, DataminingPhrase

    Gui, SettingsGui:Destroy
    ToolTip, Settings saved!
    SetTimer, RemoveToolTip, -1500
return

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
; Load Preferences on Startup
; =========================================
LoadPreferences() {
    global IncludeDatamining, ShowCitations, DataminingPhrase

    prefsFile := A_ScriptDir . "\RadAssist_preferences.ini"
    if (FileExist(prefsFile)) {
        IniRead, IncludeDatamining, %prefsFile%, Settings, IncludeDatamining, 1
        IniRead, ShowCitations, %prefsFile%, Settings, ShowCitations, 1
        IniRead, DataminingPhrase, %prefsFile%, Settings, DataminingPhrase, SSM lung nodule

        IncludeDatamining := (IncludeDatamining = "1")
        ShowCitations := (ShowCitations = "1")
    }
}
LoadPreferences()

; =========================================
; Tray Menu
; =========================================
Menu, Tray, Tip, RadAssist v2.0 - Smart Radiology Tools
Menu, Tray, NoStandard
Menu, Tray, Add, RadAssist v2.0, TrayAbout
Menu, Tray, Add
Menu, Tray, Add, Settings, MenuSettings
Menu, Tray, Add, Reload, TrayReload
Menu, Tray, Add, Exit, TrayExit
return

TrayAbout:
    MsgBox, 64, RadAssist, RadAssist v2.0`n`nShift+Right-click in PowerScribe or Notepad`nto access radiology calculators.`n`nSmart Parse Features (NEW):`n- Smart Volume: Parse dimensions with organ detection`n- Smart RV/LV: Parse ratio with PE interpretation`n- Smart NASCET: Parse stenosis measurements`n- Smart Adrenal: Parse HU washout values`n- Parse Nodules: Fleischner 2017 from findings text`n`nGUI Calculators:`n- Ellipsoid Volume`n- Adrenal Washout`n- NASCET Stenosis`n- RV/LV Ratio`n- Fleischner 2017`n`nUtilities:`n- Report Header Template`n- Sectra History Copy (Ctrl+Shift+H)
return

TrayReload:
    Reload
return

TrayExit:
    ExitApp
return
