; ==========================================
; RadAssist - Radiology Assistant Tool
; Version: 1.0
; Description: Lean radiology workflow tool with calculators
;              Triggered by Shift+Right-click in PowerScribe/Notepad
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

    Menu, RadAssistMenu, Add, Ellipsoid Volume, MenuEllipsoidVolume
    Menu, RadAssistMenu, Add, Adrenal Washout, MenuAdrenalWashout
    Menu, RadAssistMenu, Add, NASCET Stenosis, MenuNASCET
    Menu, RadAssistMenu, Add, RV/LV Ratio, MenuRVLV
    Menu, RadAssistMenu, Add, Fleischner 2017, MenuFleischner
    Menu, RadAssistMenu, Add
    Menu, RadAssistMenu, Add, Copy Sectra History, MenuSectraHistory
    Menu, RadAssistMenu, Add
    Menu, RadAssistMenu, Add, Settings, MenuSettings

    Menu, RadAssistMenu, Show
return

; -----------------------------------------
; Menu Handlers
; -----------------------------------------
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
    Gui, EllipsoidGui:Add, Text, x10 y35, Dimension 1:
    Gui, EllipsoidGui:Add, Edit, x100 y32 w60 vEllipDim1
    Gui, EllipsoidGui:Add, Text, x10 y60, Dimension 2:
    Gui, EllipsoidGui:Add, Edit, x100 y57 w60 vEllipDim2
    Gui, EllipsoidGui:Add, Text, x10 y85, Dimension 3:
    Gui, EllipsoidGui:Add, Edit, x100 y82 w60 vEllipDim3
    Gui, EllipsoidGui:Add, Text, x10 y115, Units:
    Gui, EllipsoidGui:Add, DropDownList, x100 y112 w60 vEllipUnits Choose1, mm|cm
    Gui, EllipsoidGui:Add, Button, x10 y145 w80 gCalcEllipsoid, Calculate
    Gui, EllipsoidGui:Add, Button, x100 y145 w80 gEllipsoidGuiClose, Cancel
    Gui, EllipsoidGui:Show, x%xPos% y%yPos% w200 h180, Ellipsoid Volume
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
Menu, Tray, Tip, RadAssist - Radiology Tools
Menu, Tray, NoStandard
Menu, Tray, Add, RadAssist v1.0, TrayAbout
Menu, Tray, Add
Menu, Tray, Add, Settings, MenuSettings
Menu, Tray, Add, Reload, TrayReload
Menu, Tray, Add, Exit, TrayExit
return

TrayAbout:
    MsgBox, 64, RadAssist, RadAssist v1.0`n`nShift+Right-click in PowerScribe or Notepad`nto access radiology calculators.`n`nCalculators:`n- Ellipsoid Volume`n- Adrenal Washout`n- NASCET Stenosis`n- RV/LV Ratio`n- Fleischner 2017
return

TrayReload:
    Reload
return

TrayExit:
    ExitApp
return
