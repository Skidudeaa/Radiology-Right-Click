; ==========================================
; Radiologist's Helper Script (No OCR)
; Version: 1.20
; Description: This AutoHotkey script provides various calculation tools and utilities
;              for radiologists, including volume calculations, date estimations,
;              and other analyses, **without** OCR dependencies.
; ==========================================

#NoEnv
#SingleInstance, Force
SetWorkingDir, %A_ScriptDir%

; -----------------------------------------
; Global User-Modifiable Toggles (read from Preferences)
; -----------------------------------------
global DisplayUnits := true
global DisplayAllValues := true
global ShowEllipsoidVolume := true
global ShowBulletVolume := true
global ShowPSADensity := true
global ShowPregnancyDates := true
global ShowMenstrualPhase := true
global ShowAdrenalWashout := true
global ShowThymusChemicalShift := true
global ShowHepaticSteatosis := true
global ShowMRILiverIron := true
global ShowStatistics := true
global ShowNumberRange := true
global PauseDuration := 180000
global DarkMode := false
global ShowCalciumScorePercentile := true
global ShowCitations := true
global ShowArterialAge := true
global ShowContrastPremedication := true
global ShowFleischnerCriteria := true
global ShowNASCETCalculator := true
global ShowSpellCheck := true

; -------------- NEW/CHANGED CODE --------------
; Menu-sorting-related Globals
global MenuSortingMethod := "frequency"         ; can be "none", "frequency", or "custom"
global DefaultCustomMenuOrder := "CalculateEllipsoidVolume,CalculateBulletVolume,CalculatePSADensity,CompareNoduleSizes,SortNoduleSizes,CalculateStatistics,Range,CalculateCalciumScorePercentile,CalculatePregnancyDates,CalculateMenstrualPhase,CalculateAdrenalWashout,CalculateThymusChemicalShift,CalculateHepaticSteatosis,CalculateIronContent,CalculateContrastPremedication,CalculateFleischnerCriteria,CalculateNASCET"
global CustomMenuOrder := ""  ; comma-separated function-list from .ini
global g_FunctionFrequency := {}  ; stores usage frequency per function

; Some global variables used throughout
global g_SelectedText := ""
global TargetApps := ["ahk_class Notepad", "ahk_exe notepad.exe", "ahk_class PowerScribe", "ahk_exe PowerScribe.exe", "ahk_class PowerScribe360", "ahk_exe Nuance.PowerScribe360.exe", "ahk_class PowerScribe | Reporting"]
global ResultText
global InvisibleControl
global originalMouseX, originalMouseY

; Some global reference variables
global g_References := {}  ; Will store {name: {type: "url|file|mapped", path: "path", uses: 0}}
global g_MaxReferences := 15
global vRefPath
global vRefType
global vRefList
global vRefName
global vRefPat
global RefList

; Some global AI variables
global g_AIPrompts := {}  ; Will store {name: {prompt: "prompt text", uses: 0}}
global g_MaxAIPrompts := 15
global g_PreferredAI := "claude"  ; can be "claude" or "chatgpt"
global vPromptName
global vPromptText
global PromptList
global g_OriginalPromptName := ""
; Default AI prompts
; Initialize the global object.
global DEFAULT_PROMPTS := Object()

; Build the "Edit Report" prompt object.
DEFAULT_PROMPTS["Edit Report"] := Object()
DEFAULT_PROMPTS["Edit Report"].prompt := "Edit this radiology report for clarity while preserving its original meaning. Do not explain your reasoning. Provide two outputs:`n`nVERSION WITH EDITS, in a standard output method:`nShow all changes using:`n- ~~strikethrough~~ for deletions`n- **bold** for additions`n`nFINAL VERSION, in an artifact or code based output window:`n{clean version in a copy-friendly format}`n`nREQUIREMENTS:`n1. Fix all:`n   - Grammar`n   - Punctuation`n   - Sentence structure`n   - Sentence fragments`n2. Maintain report structure:`n   - Keep narrative format if narrative`n   - Keep structured format if structured`n   - Preserve all original sections`n3. Verify presence and accuracy of:`n   - History`n   - Technique (do not modify)`n   - Comparison`n   - Findings`n   - Impression`n4. Flag potential errors in:`n   - Technique descriptions`n   - Comparison dates/studies`n   - Anatomical sidedness`n   - Contextual consistency`n`nALERTS:`n- Flag if technique section appears incorrect relative to the findings and impression (mismatch for billing)`n- Flag if comparison section contains errors`n- Flag any sidedness discrepancies"
DEFAULT_PROMPTS["Edit Report"].uses := 0

; Build the "Generate Impression" prompt object.
DEFAULT_PROMPTS["Generate Impression"] := Object()
DEFAULT_PROMPTS["Generate Impression"].prompt := "Generate a concise radiology report impression following these guidelines:`n`nSTRUCTURE`n- Address emergent findings and clinical questions first`n- Use complete sentences with minimal necessary words`n- Include only clinically significant diagnoses`n- Format for easy copy/paste`n`nREQUIRED STANDARDS`n- Follow established reporting systems where applicable:`n  - BI-RADS for breast imaging`n  - TI-RADS for thyroid nodules`n  - LI-RADS for liver lesions`n  - PI-RADS for prostate`n  - SRU consensus guidelines`n  - Fleischner criteria for pulmonary nodules`n  - Other ACR or society based guidelines.`n`nFOLLOW-UP RECOMMENDATIONS`n- Only recommend follow-up for:`n  - Established guideline-based scenarios`n  - Diagnostic uncertainty`n- Default intervals (unless guidelines specify otherwise):`n  - Concerning findings: 3 months`n  - Equivocal findings: 6 months`n  - Likely benign findings: 12 months`n Do not provide a ""no follow up recommended"" or ""correlate clinically"" recommendation.`n`nSIGNIFICANT INCIDENTALS`nInclude if present:`n- Pulmonary nodules`n- Indeterminate renal lesions`n- Vascular anomalies/aneurysms`n- Coronary artery calcification`n- Other findings requiring follow-up or intervention`n`nEXCLUDE`n- Benign or clinically insignificant findings`n- Chronic unchanged conditions`n- Anatomic variants without clinical impact`n- Normal post-surgical changes`n- Only state something is stable if there is a comparison referenced`n`nExample format:`n1. [Emergent findings/Clinical question answer]`n2. [Major diagnoses in order of clinical significance]`n3. [Relevant incidental findings requiring action]`n4. [Evidence-based follow-up recommendations]"
DEFAULT_PROMPTS["Generate Impression"].uses := 0


; For Fleischner
global g_Nodules := []
global g_FleischnerNodules := []
global g_ShowFleischnerCitation := false
global g_ShowFleischnerExclusions := false
global recommendations := {}

; -------------------------------------------------------------------------
; Load user preferences from an .ini file
; -------------------------------------------------------------------------
LoadPreferencesFromFile() {
    global DisplayUnits, DisplayAllValues, ShowEllipsoidVolume, ShowBulletVolume
    global ShowPSADensity, ShowPregnancyDates, ShowMenstrualPhase, ShowAdrenalWashout
    global ShowThymusChemicalShift, ShowHepaticSteatosis, ShowMRILiverIron
    global ShowStatistics, ShowNumberRange, PauseDuration, DarkMode
    global ShowCalciumScorePercentile, ShowCitations, ShowArterialAge
    global ShowContrastPremedication, ShowFleischnerCriteria, ShowNASCETCalculator
    global MenuSortingMethod, CustomMenuOrder
    global g_FunctionFrequency
	
	if (!CheckPreferencesIntegrity()) {
        MsgBox, Warning: Preferences file was corrupted and has been restored from backup.
    }

    preferencesFile := A_ScriptDir . "\preferences.ini"
    if (FileExist(preferencesFile)) {
        ; --- [Display] section ---
        IniRead, DisplayUnits, %preferencesFile%, Display, DisplayUnits, 1
        IniRead, DisplayAllValues, %preferencesFile%, Display, DisplayAllValues, 1
        IniRead, DarkMode, %preferencesFile%, Display, DarkMode, 0
        IniRead, ShowCitations, %preferencesFile%, Display, ShowCitations, 1
        IniRead, ShowArterialAge, %preferencesFile%, Display, ShowArterialAge, 1

        ; --- [Calculations] section ---
        IniRead, ShowEllipsoidVolume, %preferencesFile%, Calculations, ShowEllipsoidVolume, 1
        IniRead, ShowBulletVolume, %preferencesFile%, Calculations, ShowBulletVolume, 1
        IniRead, ShowPSADensity, %preferencesFile%, Calculations, ShowPSADensity, 1
        IniRead, ShowPregnancyDates, %preferencesFile%, Calculations, ShowPregnancyDates, 1
        IniRead, ShowMenstrualPhase, %preferencesFile%, Calculations, ShowMenstrualPhase, 1
        IniRead, ShowAdrenalWashout, %preferencesFile%, Calculations, ShowAdrenalWashout, 1
        IniRead, ShowThymusChemicalShift, %preferencesFile%, Calculations, ShowThymusChemicalShift, 1
        IniRead, ShowHepaticSteatosis, %preferencesFile%, Calculations, ShowHepaticSteatosis, 1
        IniRead, ShowMRILiverIron, %preferencesFile%, Calculations, ShowMRILiverIron, 1
        IniRead, ShowStatistics, %preferencesFile%, Calculations, ShowStatistics, 1
        IniRead, ShowNumberRange, %preferencesFile%, Calculations, ShowNumberRange, 1
        IniRead, ShowCalciumScorePercentile, %preferencesFile%, Calculations, ShowCalciumScorePercentile, 1
        IniRead, ShowContrastPremedication, %preferencesFile%, Calculations, ShowContrastPremedication, 1
        IniRead, ShowFleischnerCriteria, %preferencesFile%, Calculations, ShowFleischnerCriteria, 1
        IniRead, ShowNASCETCalculator, %preferencesFile%, Calculations, ShowNASCETCalculator, 1

        ; --- [Script] section ---
        IniRead, PauseDuration, %preferencesFile%, Script, PauseDuration, 180000

        ; --- [Menu] or [Sorting] section ---
        IniRead, MenuSortingMethod, %preferencesFile%, Menu, SortingMethod, none
        IniRead, CustomMenuOrder, %preferencesFile%, Menu, CustomMenuOrder,
		
		; --- [References] section ---
		IniRead, referencesList, %preferencesFile%, References, SavedReferences, %A_Space%
		if (referencesList != "") {
			references := StrSplit(referencesList, "|")
			for _, ref in references {
				refData := StrSplit(ref, ":::")
				if (refData.Length() >= 4)
					g_References[refData[1]] := {type: refData[2], path: refData[3], uses: refData[4]}
			}
		}
		
		; --- [AI Prompt] section ---
		IniRead, g_PreferredAI, %preferencesFile%, AIAssistant, PreferredAI, claude
		IniRead, promptsList, %preferencesFile%, AIAssistant, SavedPrompts, %A_Space%
		if (promptsList = "") {
			; Load default prompts
			for name, data in DEFAULT_PROMPTS {
				g_AIPrompts[name] := {prompt: data.prompt, uses: 0}
			}
		}
		if (promptsList != "") {
			prompts := StrSplit(promptsList, "|")
			for _, prompt in prompts {
				promptData := StrSplit(prompt, ":::")
				if (promptData.Length() >= 3) {
					decodedPrompt := StrReplace(promptData[2], "\n", "`n")
					decodedPrompt := StrReplace(decodedPrompt, "\p", "|")
					g_AIPrompts[promptData[1]] := {prompt: decodedPrompt, uses: promptData[3]}
				}
			}
		}
		
		; [Spell Check]
		IniRead, ShowSpellCheck, %preferencesFile%, Calculations, ShowSpellCheck, 1
		

        ; Convert string "1" to boolean true, etc.
        DisplayUnits := (DisplayUnits = "1")
        DisplayAllValues := (DisplayAllValues = "1")
        ShowEllipsoidVolume := (ShowEllipsoidVolume = "1")
        ShowBulletVolume := (ShowBulletVolume = "1")
        ShowPSADensity := (ShowPSADensity = "1")
        ShowPregnancyDates := (ShowPregnancyDates = "1")
        ShowMenstrualPhase := (ShowMenstrualPhase = "1")
        ShowAdrenalWashout := (ShowAdrenalWashout = "1")
        ShowThymusChemicalShift := (ShowThymusChemicalShift = "1")
        ShowHepaticSteatosis := (ShowHepaticSteatosis = "1")
        ShowMRILiverIron := (ShowMRILiverIron = "1")
        ShowStatistics := (ShowStatistics = "1")
        ShowNumberRange := (ShowNumberRange = "1")
        DarkMode := (DarkMode = "1")
        ShowCalciumScorePercentile := (ShowCalciumScorePercentile = "1")
        ShowCitations := (ShowCitations = "1")
        ShowArterialAge := (ShowArterialAge = "1")
        ShowContrastPremedication := (ShowContrastPremedication = "1")
        ShowFleischnerCriteria := (ShowFleischnerCriteria = "1")
        ShowNASCETCalculator := (ShowNASCETCalculator = "1")
        PauseDuration += 0
		ShowSpellCheck := (ShowSpellCheck = "1")

        ; Load frequency data
        LoadMenuFrequencies()
    } else {
        ; If no .ini file, do nothing special. Will be created on saving prefs.
    }
}
LoadPreferencesFromFile()
InitializeRecommendations()  ; For Fleischner

; -------------------------------------------------------------------------
; Load stored frequency data from .ini
; -------------------------------------------------------------------------
LoadMenuFrequencies() {
    global g_FunctionFrequency
    preferencesFile := A_ScriptDir . "\preferences.ini"
    section := "Frequency"
    freqKeys := []

    IniRead, entireSection, %preferencesFile%, %section%
    if (entireSection != "ERROR") {
        Loop, Parse, entireSection, `n, `r
        {
            line := A_LoopField
            if (RegExMatch(line, "^(.*?)=(.*)$", m)) {
                cmd := m1
                val := m2+0
                g_FunctionFrequency[cmd] := val
            }
        }
    }
}

; -------------------------------------------------------------------------
; Save frequency to .ini after increment
; -------------------------------------------------------------------------
IncrementFunctionFrequency(cmd) {
    global g_FunctionFrequency
    preferencesFile := A_ScriptDir . "\preferences.ini"

    if !(g_FunctionFrequency.HasKey(cmd)) {
        g_FunctionFrequency[cmd] := 0
    }
    g_FunctionFrequency[cmd]++

    IniWrite, % g_FunctionFrequency[cmd], %preferencesFile%, Frequency, %cmd%
}

; ------------------------------------------
; HOTKEYS, MENUS, AND SCRIPT LOGIC
; ------------------------------------------
#If IsTargetApp()

; Minor key hook: reset timer if user typed ". " quickly
$Space::
    if (A_PriorKey == "." and A_TimeSincePriorHotkey < 500) {
        lastPeriodTime := 0
    }
    SendInput {Space}
return

; -------------------------------------------------------------------------
; Right-click context menu activation (only in specified apps)
; -------------------------------------------------------------------------
CoordMode, Mouse, Screen
RButton::
{
    ; Store original mouse position
    CoordMode, Mouse, Screen  ; Ensure we're using screen coordinates
    MouseGetPos, originalX, originalY, windowUnderCursor
    
    if (windowUnderCursor != WinActive("A")) {
        WinActivate, ahk_id %windowUnderCursor%
        Sleep, 50
        ; Restore mouse position after activation
        MouseMove, %originalX%, %originalY%, 0
    }

    ; Grab selected text
    g_SelectedText := GetSelectedText()

    ; Build + show menu
    CreateCustomMenu()
    Menu, CustomMenu, Show
    Menu, CustomMenu, DeleteAll
    return
}
#If

; -------------------------------------------------------------------------
; Validate whether the current window is a target app
; -------------------------------------------------------------------------
IsTargetApp() {
    MouseGetPos, , , windowUnderCursor
    WinGetClass, windowClass, ahk_id %windowUnderCursor%
    WinGet, windowExe, ProcessName, ahk_id %windowUnderCursor%

    for index, app in TargetApps {
        if (windowClass == StrReplace(app, "ahk_class ", "")
         || windowExe   == StrReplace(app, "ahk_exe ",   "")) {
            return true
        }
    }
    return false
}

; -------------------------------------------------------------------------
; Build the context (right-click) menu
; -------------------------------------------------------------------------
CreateCustomMenu() {
    global DarkMode

    Menu, CustomMenu, Add
    Menu, CustomMenu, DeleteAll

    ; Set menu colors based on Dark Mode
    if (DarkMode) {
        Menu, CustomMenu, Color, 0xA9A9A9
    } else {
        Menu, CustomMenu, Color, Default
    }

    ; Standard editing items (always at the top)
    Menu, CustomMenu, Add, Cut, MenuCut
    Menu, CustomMenu, Add, Copy, MenuCopy
    Menu, CustomMenu, Add, Paste, MenuPaste
    Menu, CustomMenu, Add, Delete, MenuDelete
	Menu, CustomMenu, Add
	
	; Add spell check suggestion if enabled and text is selected
	if (ShowSpellCheck = 1 && g_SelectedText != "") {
		Menu, CustomMenu, Add, Check Spelling, CheckSpelling
		Menu, CustomMenu, Add  ; Separator
	}
	
	
    ; Build + sort menu items
    menuItems := BuildMenuItemsArray()
    SortMenuItems(menuItems)

    ; Add sorted items
    For _, item in menuItems {
        if (item.show) {
            Menu, CustomMenu, Add, % item.title, % item.command
        }
    }
	
	; Add References submenu
    Menu, CustomMenu, Add  ; Separator
    CreateReferencesMenu()
    Menu, CustomMenu, Add, References, :ReferencesMenu
	
	Menu, CustomMenu, Add
	CreateAIAssistantMenu()
	Menu, CustomMenu, Add, AI Assistant, :AIAssistantMenu

    ; Final items (always at bottom)
    Menu, CustomMenu, Add
    Menu, CustomMenu, Add, Pause Script, PauseScript
    Menu, CustomMenu, Add, Preferences, ShowPreferences
}

; -------------------------------------------------------------------------
; Return a list of all possible menu items with relevant data
; -------------------------------------------------------------------------
BuildMenuItemsArray() {
    global ShowCalciumScorePercentile, ShowEllipsoidVolume, ShowBulletVolume, ShowPSADensity
    global ShowPregnancyDates, ShowMenstrualPhase, ShowAdrenalWashout, ShowThymusChemicalShift, ShowHepaticSteatosis
    global ShowMRILiverIron, ShowStatistics, ShowNumberRange, ShowContrastPremedication
    global ShowFleischnerCriteria, ShowNASCETCalculator, g_FunctionFrequency

    items := []
    ; Each entry => {title, command, freq, show, customKey}
    ; customKey is used to see how it matches in the CustomMenuOrder
    ; 'command' must match our label name exactly for the "IncrementFunctionFrequency"

    items.Push({title: "Compare Measurement Sizes", command: "CompareNoduleSizes", freq: g_FunctionFrequency["CompareNoduleSizes"]+0, show: true,  customKey: "CompareNoduleSizes"})
    items.Push({title: "Sort Measurement Sizes", command: "SortNoduleSizes", freq: g_FunctionFrequency["SortNoduleSizes"]+0, show: true,  customKey: "SortNoduleSizes"})

    items.Push({title: "Calculate Calcium Score Percentile", command: "CalculateCalciumScorePercentile", freq: g_FunctionFrequency["CalculateCalciumScorePercentile"]+0, show: ShowCalciumScorePercentile, customKey: "CalculateCalciumScorePercentile"})
    items.Push({title: "Calculate Ellipsoid Volume", command: "CalculateEllipsoidVolume", freq: g_FunctionFrequency["CalculateEllipsoidVolume"]+0, show: ShowEllipsoidVolume, customKey: "CalculateEllipsoidVolume"})
    items.Push({title: "Calculate Bullet Volume", command: "CalculateBulletVolume", freq: g_FunctionFrequency["CalculateBulletVolume"]+0, show: ShowBulletVolume, customKey: "CalculateBulletVolume"})
    items.Push({title: "Calculate PSA Density", command: "CalculatePSADensity", freq: g_FunctionFrequency["CalculatePSADensity"]+0, show: ShowPSADensity, customKey: "CalculatePSADensity"})
    items.Push({title: "Calculate Pregnancy Dates", command: "CalculatePregnancyDates", freq: g_FunctionFrequency["CalculatePregnancyDates"]+0, show: ShowPregnancyDates, customKey: "CalculatePregnancyDates"})
    items.Push({title: "Calculate Menstrual Phase", command: "CalculateMenstrualPhase", freq: g_FunctionFrequency["CalculateMenstrualPhase"]+0, show: ShowMenstrualPhase, customKey: "CalculateMenstrualPhase"})
    items.Push({title: "Calculate Adrenal Washout", command: "CalculateAdrenalWashout", freq: g_FunctionFrequency["CalculateAdrenalWashout"]+0, show: ShowAdrenalWashout, customKey: "CalculateAdrenalWashout"})
    items.Push({title: "Calculate Thymus Chemical Shift", command: "CalculateThymusChemicalShift", freq: g_FunctionFrequency["CalculateThymusChemicalShift"]+0, show: ShowThymusChemicalShift, customKey: "CalculateThymusChemicalShift"})
    items.Push({title: "Calculate Hepatic Steatosis", command: "CalculateHepaticSteatosis", freq: g_FunctionFrequency["CalculateHepaticSteatosis"]+0, show: ShowHepaticSteatosis, customKey: "CalculateHepaticSteatosis"})
    items.Push({title: "MRI Liver Iron Content", command: "CalculateIronContent", freq: g_FunctionFrequency["CalculateIronContent"]+0, show: ShowMRILiverIron, customKey: "CalculateIronContent"})
    items.Push({title: "Calculate Statistics", command: "Statistics", freq: g_FunctionFrequency["Statistics"]+0, show: ShowStatistics, customKey: "Statistics"})
    items.Push({title: "Calculate Number Range", command: "Range", freq: g_FunctionFrequency["Range"]+0, show: ShowNumberRange, customKey: "Range"})
    items.Push({title: "Calculate Contrast Premedication", command: "CalculateContrastPremedication", freq: g_FunctionFrequency["CalculateContrastPremedication"]+0, show: ShowContrastPremedication, customKey: "CalculateContrastPremedication"})
    items.Push({title: "Calculate Fleischner Criteria", command: "CalculateFleischnerCriteria", freq: g_FunctionFrequency["CalculateFleischnerCriteria"]+0, show: ShowFleischnerCriteria, customKey: "CalculateFleischnerCriteria"})
    items.Push({title: "Calculate NASCET", command: "CalculateNASCET", freq: g_FunctionFrequency["CalculateNASCET"]+0, show: ShowNASCETCalculator, customKey: "CalculateNASCET"})
	; items.Push({title: "Check Spelling", command: "CheckSpelling", freq: g_FunctionFrequency["CheckSpelling"]+0, show: ShowSpellCheck, customKey: "CheckSpelling"})

    return items
}

; -------------------------------------------------------------------------
; Sort the menu items (by freq or custom)
; -------------------------------------------------------------------------
FrequencySort(a, b) {
    return b.freq - a.freq
}

CustomSort(a, b) {
    global CustomMenuOrder
    static mapOrder := {}

    ; Clear and rebuild mapOrder
    mapOrder := {}
    customList := StrSplit(CustomMenuOrder, ",")
    For i, cmd in customList {
        cmd := Trim(cmd)
        mapOrder[cmd] := i
    }

    aPos := mapOrder.HasKey(a.customKey) ? mapOrder[a.customKey] : 999999
    bPos := mapOrder.HasKey(b.customKey) ? mapOrder[b.customKey] : 999999
    return aPos - bPos
}

SortMenuItems(ByRef items) {
    global MenuSortingMethod, CustomMenuOrder

    if (MenuSortingMethod = "none" || MenuSortingMethod = "") {
        return  ; do nothing
    } else if (MenuSortingMethod = "frequency") {
        tempArray := []
        for index, item in items {
            tempArray.Insert(item)
        }

        tempSorted := []
        while (tempArray.Length() > 0) {
            highestFreq := -1
            highestIndex := 0
            for index, item in tempArray {
                if (item.freq > highestFreq) {
                    highestFreq := item.freq
                    highestIndex := index
                }
            }
            tempSorted.Insert(tempArray[highestIndex])
            tempArray.RemoveAt(highestIndex)
        }
        items := tempSorted
    } else if (MenuSortingMethod = "custom") {
        if (CustomMenuOrder = "") {
            return  ; no custom order specified
        }

        orderMap := {}
        customList := StrSplit(CustomMenuOrder, ",")
        for index, cmd in customList {
            cmd := Trim(cmd)
            orderMap[cmd] := index
        }

        tempArray := []
        for index, item in items {
            tempArray.Insert(item)
        }

        tempSorted := []
        while (tempArray.Length() > 0) {
            lowestOrder := 999999
            lowestIndex := 0
            for index, item in tempArray {
                itemOrder := orderMap.HasKey(item.customKey) ? orderMap[item.customKey] : 999999
                if (itemOrder < lowestOrder) {
                    lowestOrder := itemOrder
                    lowestIndex := index
                }
            }
            tempSorted.Insert(tempArray[lowestIndex])
            tempArray.RemoveAt(lowestIndex)
        }
        items := tempSorted
    }
}

; -------------------------------------------------------------------------
; Basic editing items
; -------------------------------------------------------------------------
MenuCut:
    Send, ^x
return

MenuCopy:
    Send, ^c
return

MenuPaste:
    Send, ^v
return

MenuDelete:
    Send, {Delete}
return

; -------------------------------------------------------------------------
; GetSelectedText: Captures text from the clipboard after Ctrl+C
; -------------------------------------------------------------------------
GetSelectedText() {
    OldClipboard := ClipboardAll
    Clipboard := ""
    Send, ^c
    ClipWait, 0.1
    if (ErrorLevel) {
        Clipboard := OldClipboard
        return ""
    }
    SelectedText := Clipboard
    Clipboard := OldClipboard
    return SelectedText
}

; -------------------------------------------------------------------------
; ShowResult( ResultString )
; Displays the text in a small, always-on-top GUI near the mouse pointer.
; -------------------------------------------------------------------------
ShowResult(Result) {
    global DarkMode, originalMouseX, originalMouseY

    CoordMode, Mouse, Screen
    MouseGetPos, originalMouseX, originalMouseY

    ; Which monitor are we on?
    SysGet, monitorCount, MonitorCount
    Loop, %monitorCount%
    {
        SysGet, monArea, Monitor, %A_Index%
        if (originalMouseX >= monAreaLeft && originalMouseX <= monAreaRight
         && originalMouseY >= monAreaTop  && originalMouseY <= monAreaBottom) {
            activeMonitor := A_Index
            break
        }
    }

    ; Get monitor area
    SysGet, workArea, MonitorWorkArea, %activeMonitor%
    monitorWidth := workAreaRight - workAreaLeft
    monitorHeight := workAreaBottom - workAreaTop

    ; We'll measure text first
    maxWidth := monitorWidth * 0.5
    maxHeight := monitorHeight * 0.5

    Gui, TempMeasure:New, +AlwaysOnTop
    Gui, TempMeasure:Font, s10, Segoe UI
    Gui, TempMeasure:Add, Text, w%maxWidth% wrap, %Result%
    GuiControlGet, TextSize, TempMeasure:Pos, Static1
    Gui, TempMeasure:Destroy

    requiredWidth := TextSizeW + 40
    requiredHeight := TextSizeH + 80
    guiWidth := (requiredWidth > maxWidth) ? maxWidth : requiredWidth
    guiHeight := (requiredHeight > maxHeight) ? maxHeight : requiredHeight
    guiWidth := (guiWidth < 300) ? 300 : guiWidth
    guiHeight := (guiHeight < 200) ? 200 : guiHeight

    xPos := originalMouseX + 10
    yPos := originalMouseY + 10
    if (xPos + guiWidth > workAreaRight)
        xPos := workAreaRight - guiWidth
    if (yPos + guiHeight > workAreaBottom)
        yPos := workAreaBottom - guiHeight

    Gui, ResultBox:New, +AlwaysOnTop -SysMenu +Owner
    Gui, ResultBox:Margin, 10, 10

    if (DarkMode) {
        Gui, ResultBox:Color, 0x2C2C2C, 0x2C2C2C
        textColor := "cE0E0E0"
        buttonOptions := "Background333333 c999999"
    } else {
        Gui, ResultBox:Color, 0xF0F0F0, 0xF0F0F0
        textColor := "c000000"
        buttonOptions := "Background777777 cFFFFFF"
    }

    Gui, ResultBox:Font, s10 %textColor%, Segoe UI
    editHeight := guiHeight - 50
    Gui, ResultBox:Add, Edit, vResultText ReadOnly -E0x200 +E0x20000 Wrap VScroll w%guiWidth% h%editHeight%, %Result%

    Gui, ResultBox:Font, s9 bold, Segoe UI
    Gui, ResultBox:Add, Button, gCloseResultBox w90 x10 y+10 %buttonOptions%, Close

    Gui, ResultBox:Add, Text, Hidden vInvisibleControl
    Gui, ResultBox:Show, x%xPos% y%yPos% w%guiWidth% h%guiHeight%, Result
    GuiControl, Focus, InvisibleControl

    ; Move mouse over close button to make it easy for user
    GuiControlGet, ClosePos, ResultBox:Pos, Close
    MouseMove, % xPos + ClosePosX + (ClosePosW / 2), % yPos + ClosePosY + (ClosePosH / 2), 0

    ; Copy result to clipboard
    Clipboard := Result
    return
}

CloseResultBox:
    Gui, ResultBox:Destroy
return

; -------------------------------------------------------------------------
; Calculation Functions (menu-driven)
; Each function increments usage frequency
; -------------------------------------------------------------------------
CalculateCalciumScorePercentile:
    IncrementFunctionFrequency("CalculateCalciumScorePercentile")
    Result := CalculateCalciumScorePercentile(g_SelectedText)
    ShowResult(Result)
return

CalculateEllipsoidVolume:
    IncrementFunctionFrequency("CalculateEllipsoidVolume")
    Result := CalculateEllipsoidVolume(g_SelectedText)
    ShowResult(Result)
return

CalculateBulletVolume:
    IncrementFunctionFrequency("CalculateBulletVolume")
    Result := CalculateBulletVolume(g_SelectedText)
    ShowResult(Result)
return

CalculatePSADensity:
    IncrementFunctionFrequency("CalculatePSADensity")
    Result := CalculatePSADensity(g_SelectedText)
    ShowResult(Result)
return

CalculatePregnancyDates:
    IncrementFunctionFrequency("CalculatePregnancyDates")
    Result := CalculatePregnancyDates(g_SelectedText)
    ShowResult(Result)
return

CalculateMenstrualPhase:
    IncrementFunctionFrequency("CalculateMenstrualPhase")
    Result := CalculateMenstrualPhase(g_SelectedText)
    ShowResult(Result)
return

CompareNoduleSizes:
    IncrementFunctionFrequency("CompareNoduleSizes")
    Result := CompareNoduleSizes(g_SelectedText)
    ShowResult(Result)
return

SortNoduleSizes:
    IncrementFunctionFrequency("SortNoduleSizes")
    ProcessedText := ProcessAllNoduleSizes(g_SelectedText)
    if (ProcessedText != g_SelectedText) {
        leadingSpace := (SubStr(g_SelectedText, 1, 1) == " ") ? " " : ""
        trailingSpace := (SubStr(g_SelectedText, 0) == " ") ? " " : ""
        Clipboard := leadingSpace . Trim(ProcessedText) . trailingSpace
        Send, ^v
    }
return

CalculateAdrenalWashout:
    IncrementFunctionFrequency("CalculateAdrenalWashout")
    Result := CalculateAdrenalWashout(g_SelectedText)
    ShowResult(Result)
return

CalculateThymusChemicalShift:
    IncrementFunctionFrequency("CalculateThymusChemicalShift")
    Result := CalculateThymusChemicalShift(g_SelectedText)
    ShowResult(Result)
return

CalculateHepaticSteatosis:
    IncrementFunctionFrequency("CalculateHepaticSteatosis")
    Result := CalculateHepaticSteatosis(g_SelectedText)
    ShowResult(Result)
return

CalculateIronContent:
    IncrementFunctionFrequency("CalculateIronContent")
    Result := EstimateIronContent(g_SelectedText)
    ShowResult(Result)
return

Statistics:
    IncrementFunctionFrequency("Statistics")
    Result := CalculateStatistics(g_SelectedText)
    ShowResult(Result)
return

Range:
    IncrementFunctionFrequency("Range")
    Result := CalculateRange(g_SelectedText)
    ShowResult(Result)
return

CalculateFleischnerCriteria:
    IncrementFunctionFrequency("CalculateFleischnerCriteria")
    Result := ProcessNodules(g_SelectedText)
    ShowResult(Result)
return

CalculateNASCET:
    IncrementFunctionFrequency("CalculateNASCET")
    Result := CalculateNASCET(g_SelectedText)
    ShowResult(Result)
return

CalculateContrastPremedication:
    IncrementFunctionFrequency("CalculateContrastPremedication")
    Result := ""  ; We'll trigger a small UI for premed
    CalculateContrastPremedication()
return

CheckSpelling:
    IncrementFunctionFrequency("CheckSpelling")
    selectedWord := Trim(g_SelectedText)
    if (selectedWord = "") {
        return
    }
    
    correctedWord := GoogleAutoCorrect(SanitizePHI(selectedWord))
    
    if (correctedWord != "" && correctedWord != selectedWord) {
        ; Create right-click menu with suggestions
        Menu, SpellSuggestions, Add
        Menu, SpellSuggestions, DeleteAll
        Menu, SpellSuggestions, Add, %correctedWord%, ReplaceWithSuggestion
        Menu, SpellSuggestions, Show
        Menu, SpellSuggestions, DeleteAll
    } else {
        ShowResult("No spelling suggestions found.")
    }
return

ReplaceWithSuggestion:
    Clipboard := A_ThisMenuItem
    Send, ^v
return

; -------------------------------------------------------------------------
; RestoreDefaults: Resets certain fields in the Preferences GUI
; -------------------------------------------------------------------------
RestoreDefaults:
    global DefaultCustomMenuOrder
    GuiControl,, MenuSortingMethodChoice, |none|frequency||custom
    GuiControl,, CustomMenuOrderBox, % DefaultCustomMenuOrder
return

; -------------------------------------------------------------------------
; Script pause/resume
; -------------------------------------------------------------------------
PauseScript() {
    global PauseDuration
    if (PauseDuration < 3600000) {
        pauseMinutes := Floor(PauseDuration / 60000)
        pauseDisplay := pauseMinutes . " minute" . (pauseMinutes != 1 ? "s" : "")
    } else {
        pauseHours := Floor(PauseDuration / 3600000)
        pauseDisplay := pauseHours . " hour" . (pauseHours != 1 ? "s" : "")
    }
    Suspend, On
    SetTimer, ResumeScript, %PauseDuration%
    MsgBox, 0, Script Paused, Script paused for %pauseDisplay%. Click OK to resume immediately.
    Suspend, Off
    SetTimer, ResumeScript, Off
}

ResumeScript:
    Suspend, Off
    SetTimer, ResumeScript, Off
    MsgBox, 0, Script Resumed, The script has been automatically resumed.
return

; -------------------------------------------------------------------------
; Preferences GUI
; -------------------------------------------------------------------------
ShowPreferences() {
    global DisplayUnits, DisplayAllValues, ShowEllipsoidVolume, ShowBulletVolume, ShowPSADensity
    global ShowPregnancyDates, ShowMenstrualPhase, PauseDuration, ShowAdrenalWashout
    global ShowThymusChemicalShift, ShowHepaticSteatosis, ShowMRILiverIron, ShowStatistics
    global ShowNumberRange, DarkMode, ShowCalciumScorePercentile, ShowCitations, ShowArterialAge
    global ShowContrastPremedication, ShowFleischnerCriteria, ShowNASCETCalculator
    global MenuSortingMethod, CustomMenuOrder

    ; Position near mouse
    CoordMode, Mouse, Screen
    MouseGetPos, px, py

    SysGet, monitorCount, MonitorCount
    activeMonitor := 1
    Loop, %monitorCount% {
        SysGet, monArea, Monitor, %A_Index%
        if (px >= monAreaLeft && px <= monAreaRight && py >= monAreaTop && py <= monAreaBottom) {
            activeMonitor := A_Index
            break
        }
    }
    SysGet, workArea, MonitorWorkArea, %activeMonitor%
    monitorWidth := workAreaRight - workAreaLeft
    monitorHeight := workAreaBottom - workAreaTop

    ; Increased height to accommodate new elements
    guiW := 300
    guiH := 860
    xPos := px + 10
    yPos := py + 10
    if (xPos + guiW > workAreaRight)
        xPos := workAreaRight - guiW
    if (yPos + guiH > workAreaBottom)
        yPos := workAreaBottom - guiH

    if (DarkMode) {
        bgColor := "0x2C2C2C"
        textColor := "cE0E0E0"
        buttonOptions := "Background333333 c999999"
    } else {
        bgColor := "0xF0F0F0"
        textColor := "c000000"
        buttonOptions := "Background777777 cFFFFFF"
    }

    if (PauseDuration = 180000)
        currentPauseDuration := "3 minutes"
    else if (PauseDuration = 600000)
        currentPauseDuration := "10 minutes"
    else if (PauseDuration = 1800000)
        currentPauseDuration := "30 minutes"
    else if (PauseDuration = 3600000)
        currentPauseDuration := "1 hour"
    else if (PauseDuration = 36000000)
        currentPauseDuration := "10 hours"
    else
        currentPauseDuration := PauseDuration . " ms"

    Gui, Preferences:New, +AlwaysOnTop
    Gui, Preferences:Color, %bgColor%, %bgColor%
    Gui, Preferences:Font, s10 %textColor%, Segoe UI

    ; Basic settings section
    Gui, Add, Text, x10 y10 w200, Select functions to display:
    Gui, Add, Checkbox, x10 y30 w200 vDarkMode Checked%DarkMode%, Dark Mode

    ; Calculation functions (vertical list)
    y := 60
    Gui, Add, Checkbox, x10 y%y% w200 vShowEllipsoidVolume Checked%ShowEllipsoidVolume%, Ellipsoid Volume
    y += 30
    Gui, Add, Checkbox, x10 y%y% w200 vShowBulletVolume Checked%ShowBulletVolume%, Bullet Volume
    y += 30
    Gui, Add, Checkbox, x10 y%y% w200 vShowPSADensity Checked%ShowPSADensity%, PSA Density
    y += 30
    Gui, Add, Checkbox, x10 y%y% w200 vShowPregnancyDates Checked%ShowPregnancyDates%, Pregnancy Dates
    y += 30
    Gui, Add, Checkbox, x10 y%y% w200 vShowMenstrualPhase Checked%ShowMenstrualPhase%, Menstrual Phase
    y += 30
    Gui, Add, Checkbox, x10 y%y% w200 vShowAdrenalWashout Checked%ShowAdrenalWashout%, Adrenal Washout
    y += 30
    Gui, Add, Checkbox, x10 y%y% w200 vShowThymusChemicalShift Checked%ShowThymusChemicalShift%, Thymus Chemical Shift
    y += 30
    Gui, Add, Checkbox, x10 y%y% w200 vShowHepaticSteatosis Checked%ShowHepaticSteatosis%, Hepatic Steatosis
    y += 30
    Gui, Add, Checkbox, x10 y%y% w200 vShowMRILiverIron Checked%ShowMRILiverIron%, MRI Liver Iron Content
    y += 30
    Gui, Add, Checkbox, x10 y%y% w200 vShowStatistics Checked%ShowStatistics%, Calculate Statistics
    y += 30
    Gui, Add, Checkbox, x10 y%y% w200 vShowNumberRange Checked%ShowNumberRange%, Calculate Number Range
    y += 30
    Gui, Add, Checkbox, x10 y%y% w200 vShowCalciumScorePercentile Checked%ShowCalciumScorePercentile%, Calcium Score Percentile
    y += 30
    Gui, Add, Checkbox, x10 y%y% w200 vShowCitations Checked%ShowCitations%, Show Citations in Output
    y += 30
    Gui, Add, Checkbox, x10 y%y% w200 vShowArterialAge Checked%ShowArterialAge%, Show Arterial Age
    y += 30
    Gui, Add, Checkbox, x10 y%y% w200 vShowContrastPremedication Checked%ShowContrastPremedication%, Contrast Premedication
    y += 30
    Gui, Add, Checkbox, x10 y%y% w200 vShowFleischnerCriteria Checked%ShowFleischnerCriteria%, Fleischner Criteria
    y += 30
    Gui, Add, Checkbox, x10 y%y% w200 vShowNASCETCalculator Checked%ShowNASCETCalculator%, NASCET Calculator
	y += 30
	Gui, Add, Checkbox, x10 y%y% w200 vShowSpellCheck Checked%ShowSpellCheck%, Spell Check

    ; Pause duration
    y += 40
    Gui, Add, Text, x10 y%y% w250, Pause Length (current: %currentPauseDuration%):
    y += 30
    Gui, Add, DropDownList, x10 y%y% w200 vPauseDurationChoice, 3 minutes|10 minutes|30 minutes|1 hour|10 hours
    if (PauseDuration = 180000)
        GuiControl, Choose, PauseDurationChoice, 1
    else if (PauseDuration = 600000)
        GuiControl, Choose, PauseDurationChoice, 2
    else if (PauseDuration = 1800000)
        GuiControl, Choose, PauseDurationChoice, 3
    else if (PauseDuration = 3600000)
        GuiControl, Choose, PauseDurationChoice, 4
    else if (PauseDuration = 36000000)
        GuiControl, Choose, PauseDurationChoice, 5

    ; Menu sorting
    y += 40
    Gui, Add, Text, x10 y%y% w250, Menu Sorting Method:
    y += 20
    Gui, Add, DropDownList, x10 y%y% w200 vMenuSortingMethodChoice, none|frequency|custom
    if (MenuSortingMethod = "none")
        GuiControl, Choose, MenuSortingMethodChoice, 1
    else if (MenuSortingMethod = "frequency")
        GuiControl, Choose, MenuSortingMethodChoice, 2
    else if (MenuSortingMethod = "custom")
        GuiControl, Choose, MenuSortingMethodChoice, 3

    y += 30
    Gui, Add, Text, x10 y%y% w280, Custom Menu Order (comma-separated function names):
    y += 20
    Gui, Add, Edit, x10 y%y% w280 h50 vCustomMenuOrderBox, % CustomMenuOrder

    ; Buttons at the bottom
    y += 60
    Gui, Add, Button, x10 y%y% w135 gRestoreDefaults %buttonOptions%, Restore Defaults
    Gui, Add, Button, x155 y%y% w135 gSavePreferences %buttonOptions%, Save

    Gui, Show, x%xPos% y%yPos% w%guiW% h%guiH%, Preferences
}

; -------------------------------------------------------------------------
; SavePreferences: stores user selection into .ini
; -------------------------------------------------------------------------
SavePreferences:
    Gui, Submit, NoHide
    global DisplayUnits, DisplayAllValues, DarkMode
    global ShowEllipsoidVolume, ShowBulletVolume, ShowPSADensity, ShowPregnancyDates
    global ShowMenstrualPhase, ShowAdrenalWashout, ShowThymusChemicalShift
    global ShowHepaticSteatosis, ShowMRILiverIron, ShowStatistics, ShowNumberRange
    global PauseDuration, ShowCitations, ShowArterialAge
    global ShowCalciumScorePercentile, ShowContrastPremedication
    global ShowFleischnerCriteria, ShowNASCETCalculator
    global MenuSortingMethod, CustomMenuOrder
    global MenuSortingMethodChoice, CustomMenuOrderBox

    if (PauseDurationChoice = "3 minutes")
        PauseDuration := 180000
    else if (PauseDurationChoice = "10 minutes")
        PauseDuration := 600000
    else if (PauseDurationChoice = "30 minutes")
        PauseDuration := 1800000
    else if (PauseDurationChoice = "1 hour")
        PauseDuration := 3600000
    else if (PauseDurationChoice = "10 hours")
        PauseDuration := 36000000

    MenuSortingMethod := MenuSortingMethodChoice
    CustomMenuOrder := CustomMenuOrderBox

    SavePreferencesToFile()
    Gui, Destroy
return

PreferencesGuiClose:
PreferencesGuiEscape:
    Gui, Destroy
return

; -------------------------------------------------------------------------
; SavePreferencesToFile: Writes final user preferences to .ini
; -------------------------------------------------------------------------
SavePreferencesToFile() {
    ; Create a backup of current preferences before making changes
    preferencesFile := A_ScriptDir . "\preferences.ini"
    backupFile := A_ScriptDir . "\preferences.backup.ini"
    if (FileExist(preferencesFile))
        FileCopy, %preferencesFile%, %backupFile%, 1

    ; 1. First verify we have valid data before writing
    if (!g_References)
        g_References := {}
    if (!g_AIPrompts) {
        g_AIPrompts := {}
        ; Load default prompts if empty
        for name, data in DEFAULT_PROMPTS {
            g_AIPrompts[name] := {prompt: data.prompt, uses: 0}
        }
    }

    ; 2. Build the references string with validation
    referencesList := ""
    for name, data in g_References {
        if (name && data.type && data.path) {  ; Validate required fields
            if (referencesList != "")
                referencesList .= "|"
            referencesList .= name . ":::" . data.type . ":::" . data.path . ":::" . (data.uses ? data.uses : 0)
        }
    }

    ; 3. Build the prompts string with validation
    promptsList := ""
    for name, data in g_AIPrompts {
        if (name && data.prompt) {  ; Validate required fields
            if (promptsList != "")
                promptsList .= "|"
            encodedPrompt := StrReplace(data.prompt, "`n", "\n")
            encodedPrompt := StrReplace(encodedPrompt, "|", "\p")
            promptsList .= name . ":::" . encodedPrompt . ":::" . (data.uses ? data.uses : 0)
        }
    }

    ; 4. Write to a temporary file first
    tempFile := A_ScriptDir . "\preferences.tmp"
    
    ; Write display settings
    IniWrite, %DisplayUnits%, %tempFile%, Display, DisplayUnits
    IniWrite, %DisplayAllValues%, %tempFile%, Display, DisplayAllValues
    IniWrite, %DarkMode%, %tempFile%, Display, DarkMode
    IniWrite, %ShowCitations%, %tempFile%, Display, ShowCitations
    IniWrite, %ShowArterialAge%, %tempFile%, Display, ShowArterialAge

    ; Write calculation settings
    IniWrite, %ShowEllipsoidVolume%, %tempFile%, Calculations, ShowEllipsoidVolume
    IniWrite, %ShowBulletVolume%, %tempFile%, Calculations, ShowBulletVolume
    IniWrite, %ShowPSADensity%, %tempFile%, Calculations, ShowPSADensity
    IniWrite, %ShowPregnancyDates%, %tempFile%, Calculations, ShowPregnancyDates
    IniWrite, %ShowMenstrualPhase%, %tempFile%, Calculations, ShowMenstrualPhase
    IniWrite, %ShowAdrenalWashout%, %tempFile%, Calculations, ShowAdrenalWashout
    IniWrite, %ShowThymusChemicalShift%, %tempFile%, Calculations, ShowThymusChemicalShift
    IniWrite, %ShowHepaticSteatosis%, %tempFile%, Calculations, ShowHepaticSteatosis
    IniWrite, %ShowMRILiverIron%, %tempFile%, Calculations, ShowMRILiverIron
    IniWrite, %ShowStatistics%, %tempFile%, Calculations, ShowStatistics
    IniWrite, %ShowNumberRange%, %tempFile%, Calculations, ShowNumberRange
    IniWrite, %ShowCalciumScorePercentile%, %tempFile%, Calculations, ShowCalciumScorePercentile
    IniWrite, %ShowContrastPremedication%, %tempFile%, Calculations, ShowContrastPremedication
    IniWrite, %ShowFleischnerCriteria%, %tempFile%, Calculations, ShowFleischnerCriteria
    IniWrite, %ShowNASCETCalculator%, %tempFile%, Calculations, ShowNASCETCalculator
    IniWrite, %ShowSpellCheck%, %tempFile%, Calculations, ShowSpellCheck

    ; Write script settings
    IniWrite, %PauseDuration%, %tempFile%, Script, PauseDuration

    ; Write menu settings
    IniWrite, %MenuSortingMethod%, %tempFile%, Menu, SortingMethod
    IniWrite, %CustomMenuOrder%, %tempFile%, Menu, CustomMenuOrder

    ; Write references with validation
    if (referencesList != "")
        IniWrite, %referencesList%, %tempFile%, References, SavedReferences

    ; Write AI settings with validation
    if (g_PreferredAI)
        IniWrite, %g_PreferredAI%, %tempFile%, AIAssistant, PreferredAI
    if (promptsList != "")
        IniWrite, %promptsList%, %tempFile%, AIAssistant, SavedPrompts

    ; 5. Verify the temporary file was written successfully
    if (!FileExist(tempFile)) {
        ; If temp file creation failed, restore from backup
        if (FileExist(backupFile))
            FileCopy, %backupFile%, %preferencesFile%, 1
        return
    }

    ; 6. Replace the actual preferences file with the temporary file
    FileCopy, %tempFile%, %preferencesFile%, 1
    FileDelete, %tempFile%
    
    ; 7. Keep latest backup
    FileCopy, %preferencesFile%, %backupFile%, 1
}

; Add this function to check preferences integrity
CheckPreferencesIntegrity() {
    preferencesFile := A_ScriptDir . "\preferences.ini"
    backupFile := A_ScriptDir . "\preferences.backup.ini"
    
    if (!FileExist(preferencesFile)) {
        if (FileExist(backupFile)) {
            FileCopy, %backupFile%, %preferencesFile%, 1
            return true
        }
        return false
    }
    
    ; Read critical sections to verify integrity
    IniRead, references, %preferencesFile%, References, SavedReferences, %A_Space%
    IniRead, prompts, %preferencesFile%, AIAssistant, SavedPrompts, %A_Space%
    
    if (references = "" && prompts = "") {
        if (FileExist(backupFile)) {
            FileCopy, %backupFile%, %preferencesFile%, 1
            return true
        }
    }
    
    return true
}


; ------------------------------------------
; "Actual" Calculation Code
; ------------------------------------------

; -------------------------------------------------------------
; 1) CalculateEllipsoidVolume
; -------------------------------------------------------------
CalculateEllipsoidVolume(input) {
    RegExNeedle := "\s*(\d+(?:\.\d+)?)\s*[x,]\s*(\d+(?:\.\d+)?)\s*[x,]\s*(\d+(?:\.\d+)?)\s*"
    if (RegExMatch(input, RegExNeedle, match)) {
        dimensions := [match1, match2, match3]
        dimensions := SortDimensions(dimensions)

        ; mm vs cm check
        isMillimeters := ((InStr(dimensions[1], ".") = 0)
                       && (InStr(dimensions[2], ".") = 0)
                       && (InStr(dimensions[3], ".") = 0))
        if (isMillimeters) {
            dimensions[1] := dimensions[1] / 10
            dimensions[2] := dimensions[2] / 10
            dimensions[3] := dimensions[3] / 10
        }

        volume := (1/6) * 3.141592653589793 * dimensions[1] * dimensions[2] * dimensions[3]
        volumeRounded := (volume < 1) ? Round(volume, 3) : Round(volume, 1)
        result := input . " (" . volumeRounded . (DisplayUnits ? " cc" : "") . ")"
        return result
    } else {
        return "Invalid input format for ellipsoid volume.`nExample: 3 x 2 x 1 cm"
    }
}

; -------------------------------------------------------------
; 2) CalculateBulletVolume
; -------------------------------------------------------------
CalculateBulletVolume(input) {
    RegExNeedle := "\s*(\d+(?:\.\d+)?)\s*[x,]\s*(\d+(?:\.\d+)?)\s*[x,]\s*(\d+(?:\.\d+)?)\s*"
    if (RegExMatch(input, RegExNeedle, match)) {
        dimensions := [match1, match2, match3]
        dimensions := SortDimensions(dimensions)

        isMillimeters := ((InStr(dimensions[1], ".") = 0)
                       && (InStr(dimensions[2], ".") = 0)
                       && (InStr(dimensions[3], ".") = 0))
        if (isMillimeters) {
            dimensions[1] := dimensions[1] / 10
            dimensions[2] := dimensions[2] / 10
            dimensions[3] := dimensions[3] / 10
        }

        volume := dimensions[1] * dimensions[2] * dimensions[3] * (5 * 3.141592653589793 / 24)
        volumeRounded := (volume < 1) ? Round(volume, 3) : Round(volume, 1)
        return input . " (" . volumeRounded . (DisplayUnits ? " cc" : "") . ")"
    } else {
        return "Invalid input format for bullet volume.`nExample: 3 x 2 x 1 cm"
    }
}

; -------------------------------------------------------------
; 3) CalculatePSADensity
; -------------------------------------------------------------
CalculatePSADensity(input) {
    volNotGiven := 1
    volumeMethod := "User Supplied"

    PSARegEx := "i)PSA\s*(?:level|value)?:?\s*(\d+(?:\.\d+)?)(?:\s*(?:ng\/ml|ng/mL|ng\/cc|ng/cc)?)"
    VolumeRegEx := "i)(?:(?:volume:?\s*(\d+(?:\.\d+)?)(?:\s*(?:cc|cm3|mL|ml)))|(?:Prostate )?Size:?.*?\((\d+(?:\.\d+?)\s*cc\)|(\d+(?:\.\d+)?)(?:\s*x\s*\d+(?:\.\d+)?\s*x\s*\d+(?:\.\d+)?\s*cm\s*\((\d+(?:\.\d+?)\s*cc\)))"

    if (RegExMatch(input, PSARegEx, PSAMatch)) {
        PSALevel := PSAMatch1
    } else {
        return "Invalid format for PSA density.`nExample:`nPSA: 5.6 ng/mL`nSize: 3.5 x 5.4 x 2.5 cm"
    }

    if (RegExMatch(input, VolumeRegEx, VolumeMatch)) {
        if (VolumeMatch1 != "")
            ProstateVolume := VolumeMatch1
        else if (VolumeMatch2 != "")
            ProstateVolume := VolumeMatch2
        else if (VolumeMatch4 != "")
            ProstateVolume := VolumeMatch4
    } else {
        volNotGiven := 0
        bulletResult := CalculateBulletVolume(input)
        if (!InStr(bulletResult, "Invalid input")) {
            ProstateVolume := RegExReplace(bulletResult, "s).*?(\d+(?:\.\d+)?)(?:\s*cc)?\).*", "$1")
            volumeMethod := "Bullet Volume"

            if (ProstateVolume >= 55) {
                ellipsoidResult := CalculateEllipsoidVolume(input)
                volumeMethod := "Ellipsoid Volume"
                if (!InStr(ellipsoidResult, "Invalid input")) {
                    ProstateVolume := RegExReplace(ellipsoidResult, "s).*?(\d+(?:\.\d+)?)(?:\s*cc)?\).*", "$1")
                }
            }
        } else {
            return "Prostate volume not found.`nExample:`nPSA: 5.6 ng/mL`nSize: 3.5 x 5.4 x 2.5 cm"
        }
    }

    PSADensity := PSALevel / ProstateVolume
    PSADensity := Round(PSADensity, 3)

    if (volNotGiven = 0) {
        result := input . "`nProstate volume: " . ProstateVolume . " cc (" . volumeMethod . ")`n"
        result .= "PSA Density: " . PSADensity . (DisplayUnits ? " ng/mL/cc" : "")
    } else {
        result := input . "`nPSA Density: " . PSADensity . (DisplayUnits ? " ng/mL/cc" : "")
    }
    return result
}

; -------------------------------------------------------------
; 4) CalculatePregnancyDates
; -------------------------------------------------------------
CalculatePregnancyDates(input) {
    LMPRegEx := "i)(?:LMP|Last\s*Menstrual\s*Period).*?(\d{1,2}[-/\.]\d{1,2}[-/\.]\d{2,4})"
    GARegEx := "i)(\d+)(?:\s*(?:weeks?|w))?\s*(?:and|&|,|-)?\s*(\d+)?(?:\s*(?:days?|d))?(?:.*?(?:as of|on)\s+(today|\d{1,2}[-/\.]\d{1,2}[-/\.]\d{2,4}))?"

    if (RegExMatch(input, LMPRegEx, LMPMatch)) {
        LMPDate := ParseDate(LMPMatch1)
        if (LMPDate = "Invalid Date") {
            return "Invalid LMP date. Please use MM/DD/YYYY or DD/MM/YYYY."
        }
        return CalculateDatesFromLMP(LMPDate)
    } else if (RegExMatch(input, GARegEx, GAMatch)) {
        WeeksGA := GAMatch1+0
        DaysGA := (GAMatch2 != "") ? GAMatch2+0 : 0
        ReferenceDate := (GAMatch3 != "") ? (GAMatch3 = "today" ? A_Now : ParseDate(GAMatch3)) : A_Now
        return CalculateDatesFromGA(WeeksGA, DaysGA, ReferenceDate)
    } else {
        return "Invalid format for pregnancy date calculation.`nExample:`nLMP: 01/15/2023 or GA: 12 weeks and 3 days as of today"
    }
}

; -------------------------------------------------------------
; 5) CalculateMenstrualPhase
; -------------------------------------------------------------
CalculateMenstrualPhase(input) {
    LMPRegEx := "i)(?:LMP|Last\s*Menstrual\s*Period)\s*:?\s*(\d{1,2}[-/\.]\d{1,2}[-/\.]\d{2,4})"
    if (RegExMatch(input, LMPRegEx, LMPMatch)) {
        LMPDate := ParseDate(LMPMatch1)
        if (LMPDate = "Invalid Date") {
            return "Invalid LMP date. Use MM/DD/YYYY or DD/MM/YYYY."
        }
        return DetermineMenstrualPhase(LMPDate)
    } else {
        return "Invalid format for menstrual phase calculation.`nExample: LMP: 05/01/2023"
    }
}

; -------------------------------------------------------------
; 6) Compare / sort nodule sizes
; -------------------------------------------------------------
CompareNoduleSizes(input) {
    static RegExNeedle := "i)(?:(\d{1,2}/\d{1,2}/\d{2,4})[:.]?\s*)?(\d+(?:\.\d+)?(?:\s*(?:x|\*)\s*\d+(?:\.\d+)?){0,2})\s*(cm|mm)?.*?(?:previous(?:ly)?|prior|before|old|initial).*?(?:(\d{1,2}/\d{1,2}/\d{2,4})[:.]?\s*)?(\d+(?:\.\d+)?(?:\s*(?:x|\*)\s*\d+(?:\.\d+)?){0,2})\s*(cm|mm)?(?:\s*(?:on|dated?)\s*(\d{1,2}/\d{1,2}/\d{2,4}))?"
    
    if (!RegExMatch(input, RegExNeedle, match)) {
        RegExNeedle := "i)(?:previous(?:ly)?|prior|before|old|initial).*?(?:(\d{1,2}/\d{1,2}/\d{2,4})[:.]?\s*)?(\d+(?:\.\d+)?(?:\s*(?:x|\*)\s*\d+(?:\.\d+)?){0,2})\s*(cm|mm)?(?:\s*(?:on|dated?)\s*(\d{1,2}/\d{1,2}/\d{2,4}))?.*?(?:now|current(?:ly)?|present|new|recent(?:ly)?|follow[- ]?up).*?(?:(\d{1,2}/\d{1,2}/\d{2,4})[:.]?\s*)?(\d+(?:\.\d+)?(?:\s*(?:x|\*)\s*\d+(?:\.\d+)?){0,2})\s*(cm|mm)?"
        if (!RegExMatch(input, RegExNeedle, match)) {
            return "Invalid input format. Please provide both current and previous measurements."
        }
        current := match6 . " " . match7, previous := match2 . " " . match3
        currentDate := match5, previousDate := match1 ? match1 : match4
    } else {
        current := match2 . " " . match3, previous := match5 . " " . match6
        currentDate := match1, previousDate := match4 ? match4 : match7
    }
    
    return CompareMeasurements(previous, current, previousDate, currentDate, input)
}

ProcessAllNoduleSizes(input) {
    RegExNeedleComma3 := "\s*(\d+(?:\.\d+)?)\s*,\s*(\d+(?:\.\d+)?)\s*,\s*(\d+(?:\.\d+)?)\s*"
    RegExNeedleX3 := "\s*(\d+(?:\.\d+)?)\s*x\s*(\d+(?:\.\d+)?)\s*x\s*(\d+(?:\.\d+)?)\s*"
    RegExNeedleComma2 := "\s*(\d+(?:\.\d+)?)\s*,\s*(\d+(?:\.\d+)?)\s*"
    RegExNeedleX2 := "\s*(\d+(?:\.\d+)?)\s*x\s*(\d+(?:\.\d+)?)\s*"

    input := ProcessPattern(input, RegExNeedleComma3, 3)
    input := ProcessPattern(input, RegExNeedleX3, 3)
    input := ProcessPattern(input, RegExNeedleComma2, 2)
    input := ProcessPattern(input, RegExNeedleX2, 2)

    return input
}

ProcessPattern(input, RegExNeedle, dimensions) {
    pos := 1
    while (pos := RegExMatch(input, RegExNeedle, match, pos)) {
        if (dimensions == 3)
            processed := ProcessNoduleSizes(match1, match2, match3)
        else
            processed := ProcessNoduleSizes(match1, match2)

        if (processed != match) {
            input := SubStr(input, 1, pos-1) . processed . SubStr(input, pos+StrLen(match))
        }
        pos += StrLen(processed)
    }
    return input
}

ProcessNoduleSizes(a, b, c := "") {
    aNum := a+0, bNum := b+0
    cNum := (c != "") ? c+0 : ""

    if (c != "") {
        if (aNum < bNum) {
            temp := a, a := b, b := temp
            tempNum := aNum, aNum := bNum, bNum := tempNum
        }
        if (bNum < cNum) {
            temp := b, b := c, c := temp
            tempNum := bNum, bNum := cNum, cNum := tempNum
        }
        if (aNum < bNum) {
            temp := a, a := b, b := temp
            tempNum := aNum, aNum := bNum, bNum := tempNum
        }
        return " " . Trim(a) . " x " . Trim(b) . " x " . Trim(c) . " "
    } else {
        if (aNum < bNum) {
            temp := a, a := b, b := temp
        }
        return " " . Trim(a) . " x " . Trim(b) . " "
    }
}

; -------------------------------------------------------------
; 7) Basic Statistics
; -------------------------------------------------------------
CalculateStatistics(input) {
    numbers := ExtractNumbers(input)
    count := numbers.Length()
    if (count = 0) {
        return "No numbers found in text."
    }

    result := "Statistics:`n"
    result .= "Count: " . count . "`n"
    result .= "Sum: " . Round(CalculateSum(numbers), 1) . "`n"
    result .= "Mean: " . Round(CalculateMean(numbers), 1) . "`n"
    result .= "Median: " . Round(CalculateMedian(numbers), 1) . "`n"
    result .= "Min: " . Round(Min(numbers*), 1) . "`n"
    result .= "Max: " . Round(Max(numbers*), 1) . "`n"

    if (count >= 9) {
        Q1 := Round(CalculateQuartile(numbers, 0.25), 1)
        Q3 := Round(CalculateQuartile(numbers, 0.75), 1)
        IQR := Q3 - Q1
        median := Round(CalculateMedian(numbers), 1)
        result .= "Q1: " . Q1 . "`n"
        result .= "Q3: " . Q3 . "`n"
        result .= "IQR: " . Round(IQR, 1) . "`n"
        result .= "IQR/Median: " . Round(IQR / median, 2) . "`n"
        result .= "Standard Deviation: " . Round(CalculateStandardDeviation(numbers), 1) . "`n"
    }
    return result
}

ExtractNumbers(input) {
    numbers := []
    input := RegExReplace(input, "i)(?:^|\n|\s)(?:slice|observation|sample|number|#|no\.?)\s*\d+:?\s*", "`n")
    RegExNeedle := "(-?\d+(?:\.\d+)?)(?:\s*(?:cm|mm)?)"
    pos := 1
    while (pos := RegExMatch(input, RegExNeedle, match, pos)) {
        numbers.Push(match1+0)
        pos += StrLen(match)
    }
    return numbers
}

CalculateSum(arr) {
    total := 0
    for index, value in arr {
        total += value
    }
    return total
}

CalculateMean(numbers) {
    sum := 0
    for i, num in numbers {
        sum += num
    }
    return (sum / numbers.Length())
}

CalculateMedian(numbers) {
    sortedNumbers := SortArray(numbers)
    count := sortedNumbers.Length()

    if (count = 0)
        return 0
    else if (Mod(count, 2) = 0) {
        middle1 := sortedNumbers[count//2]
        middle2 := sortedNumbers[(count//2)+1]
        return (middle1 + middle2) / 2
    } else {
        return sortedNumbers[Floor(count/2)+1]
    }
}

SortArray(numbers) {
    sortedNumbers := []
    for index, value in numbers {
        sortedNumbers.Push(value)
    }
    sortedNumbers.Sort()
    return sortedNumbers
}

CalculateQuartile(numbers, percentile) {
    sortedNumbers := SortArray(numbers)
    count := sortedNumbers.Length()
    position := (count - 1) * percentile + 1
    lower := Floor(position)
    upper := Ceil(position)
    if (lower = upper) {
        return sortedNumbers[lower]
    } else {
        return sortedNumbers[lower] + (position - lower)*(sortedNumbers[upper] - sortedNumbers[lower])
    }
}

CalculateStandardDeviation(numbers) {
    mean := CalculateMean(numbers)
    sumSquaredDiff := 0
    for i, num in numbers {
        diff := num - mean
        sumSquaredDiff += diff*diff
    }
    variance := sumSquaredDiff / (numbers.Length()-1)
    return Sqrt(variance)
}

; -------------------------------------------------------------
; 8) Number Range
; -------------------------------------------------------------
CalculateRange(input) {
    numbers := []
    unit := ""
    RegExNeedle := "(-?\d+(?:\.\d+)?)(?:\s*((?:cm/s|mm/s|m/s|km/h|mph|cm|mm|Hz|T|mg|m|ml|mL|cc|s|min|hr|days?|weeks?|months?|years?|g|ng|ng/ml|ng/mL|mmol/L|µmol/L|°F|°C)(?:/(?:day|week|month|year))?))?"
    pos := 1
    while (pos := RegExMatch(input, RegExNeedle, match, pos)) {
        numbers.Push(match1+0)
        if (match2 != "" && unit = "")
            unit := match2
        pos += StrLen(match)
    }
    if (numbers.Length() = 0) {
        return "No numbers found."
    }
    minValue := Min(numbers*)
    maxValue := Max(numbers*)
    result := Round(minValue,1) . " - " . Round(maxValue,1)
    if (unit != "")
        result .= " " . unit
    return result
}

; -------------------------------------------------------------
; 9) Adrenal Washout
; -------------------------------------------------------------
CalculateAdrenalWashout(input) {
	global ShowCitations
	
    RegExNeedle := "i)(?:(?:unenhanced|non-?enhanced|intrinsic|pre-?contrast|baseline|native)(?:\s+CT)?(?:\s+density|\s+HU|\s+hounsfield\s+units?)?[\s:]*(-?\d+(?:\.\d+)?)\s*(?:HU|hounsfield\s+units?)?).*?(?:(?:enhanced|post-?contrast|arterial|portal\s+venous?|60-?75\s*(?:second|sec)|1-?2\s*min(?:ute)?)(?:\s+CT)?(?:\s+density|\s+HU|\s+hounsfield\s+units?)?[\s:]*(-?\d+(?:\.\d+)?)\s*(?:HU|hounsfield\s+units?)?).*?(?:(?:delayed|late|15\s*min(?:ute)?|10-?15\s*min(?:ute)?|post-?contrast)(?:\s+CT)?(?:\s+density|\s+HU|\s+hounsfield\s+units?)?[\s:]*(-?\d+(?:\.\d+)?)\s*(?:HU|hounsfield\s+units?)?)"

    if (RegExMatch(input, RegExNeedle, match)) {
        unenhanced := match1
        enhanced := match2
        delayed := match3

        absoluteWashout := ((enhanced - delayed) / (enhanced - unenhanced)) * 100
        relativeWashout := ((enhanced - delayed) / enhanced) * 100

        result := input . "`n`n"
        result .= "Absolute Washout: " . Round(absoluteWashout, 1) . "% ... (Ref adenomas: >=60%)" . "`n"
        result .= "Relative Washout: " . Round(relativeWashout, 1) . "% ... (Ref adenomas: >=40%)" . "`n`n"
		
		if(ShowCitations=1){
				result .= "`nMayo-Smith WW, Song JH, Boland GL, Francis IR, Israel GM, Mazzaglia PJ, Berland LL, Pandharipande PV. Management of Incidental Adrenal Masses: A White Paper of the ACR Incidental Findings Committee. J Am Coll Radiol. 2017 Aug;14(8):1038-1044.`n`n"
		}
		
        result .= InterpretAdrenalWashout(absoluteWashout, relativeWashout, unenhanced, enhanced, delayed)
        return result
    } else {
        RegExNeedle := "i)(?:(?:enhanced|post-?contrast|arterial|portal\s+venous?|60-?75\s*(?:second|sec)|1-?2\s*min(?:ute)?)(?:\s+CT)?(?:\s+density|\s+HU|\s+hounsfield\s+units?)?[\s:]*(-?\d+(?:\.\d+)?)\s*(?:HU|hounsfield\nits?)?).*?(?:(?:delayed|late|15\s*min(?:ute)?|10-?15\s*min(?:ute)?|post-?contrast)(?:\s+CT)?(?:\s+density|\s+HU|\s+hounsfield\s+units?)?[\s:]*(-?\d+(?:\.\d+)?)\s*(?:HU|hounsfield\s+units?)?)"

        if (RegExMatch(input, RegExNeedle, match)) {
            enhanced := match1
            delayed := match2

            relativeWashout := ((enhanced - delayed) / enhanced) * 100

            result := input . "`n`n"
            result .= "Relative Washout: " . Round(relativeWashout, 1) . "%`n`n"
            result .= InterpretAdrenalWashout(0, relativeWashout, unenhanced, enhanced, delayed)
			
			if(ShowCitations=1){
				result .= "`n`n`nMayo-Smith WW, Song JH, Boland GL, Francis IR, Israel GM, Mazzaglia PJ, Berland LL, Pandharipande PV. Management of Incidental Adrenal Masses: A White Paper of the ACR Incidental Findings Committee. J Am Coll Radiol. 2017 Aug;14(8):1038-1044.`n`n"
			}
            return result
        } else {
            return "Invalid input format for adrenal washout calculation.`nSample syntax: Unenhanced: 10 HU, Enhanced: 80 HU, Delayed: 40 HU"
        }
    }
}

InterpretAdrenalWashout(absoluteWashout, relativeWashout, unenhanced, enhanced, delayed) {
    result := ""

    ; Check if unenhanced is null or empty
    if (unenhanced = "" or unenhanced = "NULL") {
        unenhanced := "N/A"
        isUnenhancedAvailable := false
    } else {
        isUnenhancedAvailable := true
        unenhanced += 0  ; Convert to number
    }

    ; Check for enhancing mass
    if (isUnenhancedAvailable) {
        enhancedChange := CheckEnhancement(enhanced, unenhanced)
        delayedChange := CheckEnhancement(delayed, unenhanced)
		
         ; Add interpretation based on unenhanced HU
        if (unenhanced <= 10) {
            result .= "Unenhanced HU less than 10 is typically suggestive of a benign adrenal adenoma. "
        } else if (unenhanced > 43) {
            result .= "Unenhanced HU >43 in a noncalcified, nonhemorrhagic lesion is suspicious for malignancy, regardless of washout characteristics. "
        }
		
        if (Abs(enhancedChange) >= 10 or Abs(delayedChange) >= 10) {
            if (enhancedChange < 0 or delayedChange < 0) {
                result .= "The adrenal mass demonstrates unexpected de-enhancement (decrease in HU in enhanced or delayed phase). This is an atypical finding. "
                result .= "Caution: Standard washout calculations may not be applicable in this case. "
            } else {
                result .= "The adrenal mass demonstrates enhancement. "
                
                ; Interpret washout only if there's positive enhancement
                if (absoluteWashout >= 60 or relativeWashout >= 40) {
                    result .= "Adrenal washout characteristics are suggestive of a benign adrenal adenoma. "
                } else {
                    result .= "Washout characteristics are indeterminate. "
                }
            }
        } else {
            result .= "The adrenal mass does not demonstrate significant enhancement (<10 HU change in both enhanced and delayed phases compared to unenhanced). This may represent a cyst, hemorrhage, or other non-enhancing lesion. Further characterization with additional imaging may be necessary. "
        }

    } else {
        result .= "Unenhanced HU value is not available. "
        
        if (absoluteWashout >= 60 or relativeWashout >= 40) {
            result .= "Based on the provided washout values alone, characteristics are suggestive of a benign adrenal adenoma. However, this interpretation is limited without the unenhanced HU value. "
        } else {
			result .= "Washout characteristics are indeterminate. "
		}
    }

    return Trim(result)
}
; Function to check enhancement
CheckEnhancement(value, baseline) {
    if (value = "" or value = "NULL" or baseline = "N/A")
        return 0
    return (value - baseline)
}

; -------------------------------------------------------------
; 10) Thymus Chemical Shift
; -------------------------------------------------------------
CalculateThymusChemicalShift(input) {
    RegExNeedle := "i)thymus.*?((?:in[- ]?phase|IP|T1IP)).*?(\d+).*?((?:out[- ]?of[- ]?phase|OP|OOP|T1OP)).*?(\d+)(?:.*?paraspinous.*?((?:in[- ]?phase|IP|T1IP)).*?(\d+).*?((?:out[- ]?of[- ]?phase|OP|OOP|T1OP)).*?(\d+))?"

    if (RegExMatch(input, RegExNeedle, match)) {
        thymusInPhase := match2
        thymusOutOfPhase := match4
        paraspinousInPhase := match6
        paraspinousOutOfPhase := match8

        signalIntensityIndex := ((thymusInPhase-thymusOutOfPhase)/thymusInPhase)*100

        result := input
        if (paraspinousInPhase && paraspinousOutOfPhase) {
            outOfPhaseSignalRatio := (thymusOutOfPhase) / paraspinousOutOfPhase
            inPhaseSignalRatio := (thymusInPhase) / paraspinousInPhase
            chemicalShiftRatio := outOfPhaseSignalRatio / inPhaseSignalRatio

            if (DisplayAllValues) {
                result .= "`n`nChemical Shift Ratio: " . Round(chemicalShiftRatio, 3) . " (hyperplasia < 0.849)`n"    
                result .= "Thymus Signal Intensity Index (SII): " . round(signalIntensityIndex, 2) . "% (hyperplasia > 8.92)`n"
            }
            result .= "`n" . InterpretThymusChemicalShift(chemicalShiftRatio, signalIntensityIndex)
        } else {
            if (DisplayAllValues) {
                result .= "`nThymus Signal Intensity Index (SII): " . round(signalIntensityIndex, 2) . "%`n"
            }
            result .= "`n" . InterpretThymusChemicalShift("", signalIntensityIndex)
        }
        return result
    } else {
        return "Invalid input format for thymus chemical shift calculation.`nSample syntax: Thymus IP: 100, OP: 80, Paraspinous IP: 90, OP: 85`nOr for thymus only: Thymus IP: 100, OP: 80"
    }
}

InterpretThymusChemicalShift(chemicalShiftRatio := "", signalIntensityIndex := "") {
    global ShowCitations
    result := ""

    if (chemicalShiftRatio != "" && signalIntensityIndex != "") {
        ; Both chemical shift ratio and signal intensity index are provided
        if (chemicalShiftRatio > 0.849 && signalIntensityIndex > 8.92) {
            result := "Interpretation: Chemical Shift Ratio is greater than 0.849 and Signal Intensity Index is greater than 8.92%. Calculations are in conflict and therefore indeterminate, though probably consistent with thymic hyperplasia with single dual echo technique.`n`n"
        } else if (chemicalShiftRatio <= 0.849 && signalIntensityIndex > 8.92) {
            result := "Interpretation: Chemical Shift Ratio is less than or equal to 0.849 and Signal Intensity Index is greater than 8.92%. Findings are consistent with thymic hyperplasia with single dual echo technique. `n`n"
        } else if (chemicalShiftRatio > 0.849 && signalIntensityIndex <= 8.92) {
            result := "Interpretation: Chemical Shift Ratio is greater than 0.849 and Signal Intensity Index is less than or equal to 8.92%. Findings are not consistent with with typical thymic hyperplasia with single dual echo technique.`n`n"
        } else {  ; chemicalShiftRatio <= 0.849 && signalIntensityIndex <= 8.92
            result := "Interpretation: Chemical Shift Ratio is less than or equal to 0.849 and Signal Intensity Index is less than or equal to 8.92%. Calculations are in conflict and therefore indeterminate, possibly thymic hyperplasia with single dual echo technique.`n`n"
        }
    } else if (chemicalShiftRatio != "" && signalIntensityIndex == "") {
        ; Only chemical shift ratio is provided
        result := "Interpretation: Chemical Shift Ratio is " . (chemicalShiftRatio > 0.849 ? "greater than" : "less than or equal to") . " 0.849 with single dual echo technique.`n`n"
    } else if (chemicalShiftRatio == "" && signalIntensityIndex != "") {
        ; Only signal intensity index is provided
        if (signalIntensityIndex > 8.92) {
            result := "Interpretation: Signal Intensity Index is greater than 8.92%. This suggests thymic hyperplasia with single dual echo technique.`n`n"
        } else {
            result := "Interpretation: Signal Intensity Index is less than or equal to 8.92%. Findings are not consistent with typical thymic hyperplasia with single dual echo technique.`n`n"
        }
    } else {
        ; Neither chemical shift ratio nor signal intensity index is provided
        result := "Error: Both Chemical Shift Ratio and Signal Intensity Index are missing. At least one value is required for interpretation.`n`n"
    }

    if (ShowCitations = 1) {
        result .= "Citation: Priola AM, Priola SM, Ciccone G, Evangelista A, Cataldi A, Gned D, Pazè F, Ducco L, Moretti F, Brundu M, Veltri A. Differentiation of rebound and lymphoid thymic hyperplasia from anterior mediastinal tumors with dual-echo chemical-shift MR imaging in adulthood: reliability of the chemical-shift ratio and signal intensity index. Radiology. 2015 Jan;274(1):238-49. doi: 10.1148/radiol.14132665. Epub 2014 Aug 7. PMID: 25105246.`n"
    }

    return result
}

; -------------------------------------------------------------
; 11) Hepatic Steatosis
; -------------------------------------------------------------
CalculateHepaticSteatosis(inputText) {
    global ShowCitations
    RegExNeedleLocal := "i)liver.*?((?:in[- ]?phase|IP|T1IP)).*?(\d+).*?((?:out[- ]?of[- ]?phase|OP|OOP|T1OP)).*?(\d+)"
    RegExNeedleSpleenLocal := "i)spleen.*?((?:in[- ]?phase|IP|T1IP)).*?(\d+).*?((?:out[- ]?of[- ]?phase|OP|OOP|T1OP)).*?(\d+)"
    
    if (RegExMatch(inputText, RegExNeedleLocal, matchLocal)) {
        liverInPhaseLocal := matchLocal2
        liverOutOfPhaseLocal := matchLocal4
        
        ; Calculate Fat fraction
        fatFractionLocal := 100 * (liverInPhaseLocal - liverOutOfPhaseLocal) / (2 * liverInPhaseLocal)
        
        resultLocal := inputText . " (Fat Fraction: " . Round(fatFractionLocal, 1) . "%)"
        resultLocal .= "`n`nFat Fraction " . InterpretHepaticSteatosis(fatFractionLocal) . "`n"
        
        ; Check if spleen values are provided
        if (RegExMatch(inputText, RegExNeedleSpleenLocal, matchSpleenLocal)) {
            spleenInPhaseLocal := matchSpleenLocal2
            spleenOutOfPhaseLocal := matchSpleenLocal4
            
            ; Calculate Fat percentage
            fatPercentageLocal := 100 * ((liverInPhaseLocal / spleenInPhaseLocal) - (liverOutOfPhaseLocal / spleenOutOfPhaseLocal)) / (2 * (liverInPhaseLocal / spleenInPhaseLocal))
            
            resultLocal := StrReplace(resultLocal, ")", ", Fat Percentage: " . Round(fatPercentageLocal, 1) . "%)")
            resultLocal .= "Fat Percentage " . InterpretHepaticSteatosis(fatPercentageLocal) . "`n"
        }
        
        if (ShowCitations = 1) {
            resultLocal .= "`n`nSirlin CB. Invited Commentary on Image-based quantification of hepatic fat: methods and clinical applications. Radiographics 2009; 29:1277-80`n"
        }
        
        return resultLocal
    } else {
        return "Invalid input format for hepatic steatosis calculation.`nSample syntax: Liver IP: 100, OP: 80, Spleen IP: 90, OP: 88"
    }
}

InterpretHepaticSteatosis(hepaticFatFraction) {
    if (hepaticFatFraction < 5) {
        return "Interpretation: No significant hepatic steatosis."
    } else if (hepaticFatFraction < 15) {
        return "Interpretation: Mild hepatic steatosis."
    } else if (hepaticFatFraction < 30) {
        return "Interpretation: Moderate hepatic steatosis."
    } else {
        return "Interpretation: Severe hepatic steatosis."
    }
}
; -------------------------------------------------------------
; 12) MRI Liver Iron Content
; -------------------------------------------------------------
EstimateIronContent(input) {
	global ShowCitations

    RegExMatch(input, "i)(?:\b|^)(1[.,]5|3[.,]0|1\.5|3\.0|2\.89)(?:\s*-?\s*)?T(?:esla)?", fieldStrength)
    if (!fieldStrength1) {
        return "Error: Magnetic field strength (1.5T, 2.89T or 3.0T) not found in the input."
    }
    fieldStrength1 := StrReplace(fieldStrength1, ",", ".")

    R2StarPattern := "i)R2\*?\s*(?:value|reading|measurement)?(?:[\s:=]+of)?\s*[:=]?\s*(\d+(?:[.,]\d+)?)\s*(?:Hz|hertz|s(?:ec(?:ond)?)?Ã¢ÂÂ»Ã‚Â¹|1/s)"
    RegExMatch(input, R2StarPattern, R2Star)
    if (!R2Star1) {
        return "Error: R2* value not found in the input.`nSample syntax: 1.5T, R2*: 50 Hz"
    }
    R2StarValue := StrReplace(R2Star1, ",", ".")
    R2StarValue += 0

    if (fieldStrength1 == "1.5") {
		ironContent :=  0.02603 * R2StarValue - 0.16
    } else if (fieldStrength1 == "2.89") {
	    ironContent := 0.01400 * R2StarValue - 0.03
	} else if (fieldStrength1 == "3.0") {
        ironContent := 0.01349 * R2StarValue - 0.03
    } else { 
	    ironContent := 0
	}

    result := input . "`nEstimated Iron Content: " . Round(ironContent, 2) . " mg Fe/g dry liver`n"
	
    if (DisplayAllValues) {
        result .= "`nMagnetic Field Strength: " . fieldStrength1 . "T`n"
        result .= "R2* Value: " . R2StarValue . " Hz`n"
    }
	
	if (ShowCitations=1){
		result .= "`nGuglielmo FF, Barr RG, Yokoo T, Ferraioli G, Lee JT, Dillman JR, Horowitz JM, Jhaveri KS, Miller FH, Modi RY, Mojtahed A, Ohliger MA, Pirasteh A, Reeder SB, Shanbhogue K, Silva AC, Smith EN, Surabhi VR, Taouli B, Welle CL, Yeh BM, Venkatesh SK. Liver Fibrosis, Fat, and Iron Evaluation with MRI and Fibrosis and Fat Evaluation with US: A Practical Guide for Radiologists. Radiographics. 2023 Jun;43(6):e220181. doi: 10.1148/rg.220181. PMID: 37227944.`n"
	}
    return result
}

; -------------------------------------------------------------
; 13) Calcium Score Percentile (no OCR)
; -------------------------------------------------------------
CalculateCalciumScorePercentile(input) {
    ; Extract age, sex, race, and calcium score from input
    RegExMatch(input, "i)Age:\s*(\d+)", age)
    RegExMatch(input, "i)Sex:\s*(Male|Female)", sex)
    RegExMatch(input, "i)Race:\s*(White|Black|Hispanic|Chinese|[A-Za-z]+)", race)
    RegExMatch(input, "i)(?:your\s+)?(?:coronary\s+artery\s+)?calcium\s+score(?:\s+is)?:?\s*(\d+(?:\.\d+)?)\s*(?:\(?\s*Agatston\s*\)?)?|(?:total\s+)?calcium\s+score:?\s*(\d+(?:\.\d+)?)", score)
	
	if (!age1) {
        return input . "`n`nError: Age not found or invalid. Please provide age in the format 'Age: 55'."
    }
    if (!sex1) {
        return input . "`n`nError: Sex not found or invalid. Please specify either Male or Female."
    }
    if (!score1 && score1 !=0 ) {
        return input . "`n`nError: Calcium score not found or invalid. Please ensure the score is provided in the format 'YOUR CORONARY ARTERY CALCIUM SCORE: 0.0 (Agatston)'."
    }

    age := age1
    sex := sex1
    race := race1 ? race1 : "Unspecified"
    score := score1

    result := ""
    if (age >= 45 && age <= 84 && (race = "White" || race = "Black" || race = "Hispanic" || race = "Chinese")) {
        result := CalculateMESAScore(age, race, sex, score)
    } else if (age >= 30) {
        result := CalculateHoffScore(age, sex, score, race)
    } else {
        result := "Error: The calcium score calculators are only valid for ages 30 and above."
    }

    return input . "`n`nCONTEXT:`n" . result
}

CalculateMESAScore(age, race, sex, score) {
	global ShowCitations
	global ShowArterialAge

    url := "https://www.mesa-nhlbi.org/Calcium/input.aspx"
    
    ; First, we need to get the initial page to retrieve some hidden values
    initialResponse := DownloadToString(url)
    if (InStr(initialResponse, "Error:")) {
        return "Initial page load failed: " . initialResponse
    }
    
    ; Extract necessary hidden values
    RegExMatch(initialResponse, "id=""__VIEWSTATE"" value=""([^""]+)""", viewState)
    RegExMatch(initialResponse, "id=""__VIEWSTATEGENERATOR"" value=""([^""]+)""", viewStateGenerator)
    RegExMatch(initialResponse, "id=""__EVENTVALIDATION"" value=""([^""]+)""", eventValidation)
    
    ; Prepare the post data
    postData := "__VIEWSTATE=" . UrlEncode(viewState1)
              . "&__VIEWSTATEGENERATOR=" . viewStateGenerator1
              . "&__EVENTVALIDATION=" . UrlEncode(eventValidation1)
              . "&Age=" . age
              . "&gender=" . GetSexValue(sex)
              . "&Race=" . GetRaceValue(race)
              . "&Score=" . score
              . "&Calculate=Calculate"

    ; Send the calculation request
    response := DownloadToString(url, postData)
    if (InStr(response, "Error:")) {
        return "Calculation request failed: " . response
    }
    
    ; Extract the results
    RegExMatch(response, "id=""Label10""[^>]*>([^<]+)</span>", probability)
    RegExMatch(response, "id=""scoreLabel""[^>]*>([^<]+)</span>", observedScore)
    RegExMatch(response, "id=""percLabel""[^>]*>([^<]+)</span>", percentile)
    
    if (probability1 || observedScore1 || percentile1) {
        result := "Probability of non-zero calcium score: " . probability1 . "`n`n"
        
        result .= "Plaque Burden: " . DeterminePlaqueBurden(score) . "`n`n"
        
        percentileNum := percentile1 + 0
        comparison := DetermineComparison(percentileNum)
        
        result .= "Comparison to people of the same age, race and sex: " . comparison . "`n`n"
        result .= "Observed calcium score of " . score . " Agatston is at percentile " . percentile1 . " for age, race and sex`n`n"
        
        if(ShowArterialAge=1){
			result .= "Arterial Age: " . CalculateCoronaryAge(score) . " years" . "`n`n"
        }
		if(ShowCitations=1) {
			result .= "Citation 1: McClelland RL, et al. Distribution of coronary artery calcium by race, gender, and age: results from the Multi-Ethnic Study of Atherosclerosis (MESA). Circulation. 2006 Jan 3;113(1):30-7.`n`n"
			result .= "Citation 2: McClelland RL, Nasir K, Budoff M, Blumenthal RS, Kronmal RA. Arterial age as a function of coronary artery calcium (from the Multi-Ethnic Study of Atherosclerosis [MESA]). Am J Cardiol. 2009 Jan 1;103(1):59-63."
        }
		
		return result
		
	} else {
		; Save the response for debugging
		FileAppend, %response%, %A_Desktop%\mesa_debug_response.html
		return "Error: Unable to retrieve result from MESA calculator. The response has been saved to your desktop as 'mesa_debug_response.html' for further investigation."
	}
}

CalculateHoffScore(age, sex, score, race) {
	global ShowCitations
	global ShowArterialAge

    ageGroup := GetAgeGroup(age)
    percentile := GetHoffPercentile(ageGroup, sex, score)
    
    result := "Plaque Burden: " . DeterminePlaqueBurden(score) . "`n`n"
    
    comparison := DetermineComparison(percentile)
    result .= "Comparison to people of the same age and sex: " . comparison . "`n`n"
    
    if(ShowArterialAge=1){
		result .= "Arterial Age: " . CalculateCoronaryAge(score) . " years" . "`n`n"
    }
    if(ShowCitations=1){
		result .= "Citation 1: Hoff JA, et al. Age and gender distributions of coronary artery calcium detected by electron beam tomography in 35,246 adults. Am J Cardiol. 2001 Jun 15;87(12):1335-9.`n`n"
		result .= "Citation 2: McClelland RL, Nasir K, Budoff M, Blumenthal RS, Kronmal RA. Arterial age as a function of coronary artery calcium (from the Multi-Ethnic Study of Atherosclerosis [MESA]). Am J Cardiol. 2009 Jan 1;103(1):59-63."
    }
	
    return result
}

GetAgeGroup(age) {
    if (age < 40)
        return "<40"
    else if (age < 45)
        return "40-44"
    else if (age < 50)
        return "45-49"
    else if (age < 55)
        return "50-54"
    else if (age < 60)
        return "55-59"
    else if (age < 65)
        return "60-64"
    else if (age < 70)
        return "65-69"
    else if (age <= 74)
        return "70-74"
    else
        return ">74"
}

GetHoffPercentile(ageGroup, sex, score) {
    percentiles := {"<40": {}, "40-44": {}, "45-49": {}, "50-54": {}, "55-59": {}, "60-64": {}, "65-69": {}, "70-74": {}, ">74": {}}
    
    ; Male percentiles
    percentiles["<40"]["Male"] := [0, 1, 3, 14]
    percentiles["40-44"]["Male"] := [0, 1, 9, 59]
    percentiles["45-49"]["Male"] := [0, 3, 36, 154]
    percentiles["50-54"]["Male"] := [1, 15, 103, 332]
    percentiles["55-59"]["Male"] := [4, 48, 215, 554]
    percentiles["60-64"]["Male"] := [13, 113, 410, 994]
    percentiles["65-69"]["Male"] := [32, 180, 566, 1299]
    percentiles["70-74"]["Male"] := [64, 310, 892, 1774]
    percentiles[">74"]["Male"] := [166, 473, 1071, 1982]
    
    ; Female percentiles
    percentiles["<40"]["Female"] := [0, 0, 1, 3]
    percentiles["40-44"]["Female"] := [0, 0, 1, 4]
    percentiles["45-49"]["Female"] := [0, 0, 2, 22]
    percentiles["50-54"]["Female"] := [0, 0, 5, 55]
    percentiles["55-59"]["Female"] := [0, 1, 23, 121]
    percentiles["60-64"]["Female"] := [0, 3, 57, 193]
    percentiles["65-69"]["Female"] := [1, 24, 145, 410]
    percentiles["70-74"]["Female"] := [3, 52, 210, 631]
    percentiles[">74"]["Female"] := [9, 75, 241, 709]
    
    agePercentiles := percentiles[ageGroup][sex]
    
    if (score <= agePercentiles[1])
        return 25
    else if (score <= agePercentiles[2])
        return 50
    else if (score <= agePercentiles[3])
        return 75
    else if (score <= agePercentiles[4])
        return 90
    else
        return 99
}

DeterminePlaqueBurden(score) {
    if (score = 0)
        return "None. Risk of coronary artery disease is very low, generally less than 5 percent."
    else if (score > 0 && score <= 10)
        return "Minimal identifiable plaque. Risk of coronary artery disease is very unlikely, less than 10 percent."
    else if (score > 10 && score <= 100)
        return "At least mild atherosclerotic plaque. Mild or minimal coronary narrowings likely."
    else if (score > 100 && score <= 400)
        return "At least moderate atherosclerotic plaque. Mild coronary artery disease highly likely, significant narrowings possible."
    else
        return "Extensive atherosclerotic plaque. High likelihood of at least one significant coronary narrowing."
}

DetermineComparison(percentile) {
    if (percentile<=25)
        return "Low (≤25%)"
    else if (percentile<=50)
        return "Average (25-50%)"
    else if (percentile<=75)
        return "Average (50-75%)"
    else if (percentile<=90)
        return "High (75-90%)"
    else
        return "Very high (>90%)"
}

CalculateCoronaryAge(score) {
    logScore := Ln(score+1)
    effectiveAge := Round(39.1 + 7.25*logScore, 0)
    return effectiveAge
}

; -------------------------------------------------------------
; 14) CalculateNASCET
; -------------------------------------------------------------
CalculateNASCET(input) {
    global ShowCitations
    raw := input
    input := RegExReplace(input, "`r?\n", " ")

    RegExNeedle := "i)(?:distal.*?(\d+(?:\.\d+)?)(?:mm|cm)).*?(?:stenosis.*?(\d+(?:\.\d+)?)(?:mm|cm))"
    if (!RegExMatch(input, RegExNeedle, match)) {
        RegExNeedle2 := "i)(?:stenosis.*?(\d+(?:\.\d+)?)(?:mm|cm)).*?(?:distal.*?(\d+(?:\.\d+)?)(?:mm|cm))"
        if (!RegExMatch(input, RegExNeedle2, match2)) {
            numbers := []
            patternAny := "(\d+(?:\.\d+)?)(?:mm|cm)?"
            pos := 1
            while (pos := RegExMatch(input, patternAny, m, pos)) {
                numbers.Push(m1+0)
                pos += StrLen(m)
            }
            if (numbers.Length()<2)
                return "Could not find two diameters for NASCET. Example: 'Distal ICA = 6 mm, Stenosis = 2 mm'"
            distal := Max(numbers*)
            stenosis := Min(numbers*)
        } else {
            stenosis := match21+0
            distal := match22+0
        }
    } else {
        distal := match1+0
        stenosis := match2+0
    }

    nascetVal := (distal - stenosis)/distal*100
    nascetVal := Round(nascetVal,1)

    result := raw . "`n`nNASCET Calculation:`nDistal: " . distal . " mm`nStenosis: " . stenosis . " mm`nNASCET: " . nascetVal . "%"

    if (nascetVal<50)
        result .= "`nMild (<50%)"
    else if (nascetVal<70)
        result .= "`nModerate (50-69%)"
    else
        result .= "`nSevere (≥70%)"

    if (ShowCitations=1) {
        result .= "`nCitation: NASCET (N Engl J Med 1991;325:445-53)."
    }
    return result
}

; ------------------------------------------
; Utility Subfunctions
; ------------------------------------------
; -------------------------------------------------------------
; GoogleAutoCorrect: Checks spelling via Google search
; -------------------------------------------------------------
GoogleAutoCorrect(Word) {
    ; Sanitize and encode the word before sending to Google
    Word := Trim(Word)
    
    ; Use the Google search "Did you mean" feature
    try {
        OutputVar := URLDownloadToVarWithHeader("https://www.google.com/search?q=" Word)
        
        pos := InStr(OutputVar, "spell=1")
        if (pos) {
            cut := SubStr(OutputVar, 1, pos-1)
            equalSign := InStr(cut, "=",, 0)
            
            if (equalSign) {
                equalSignString := SubStr(cut, equalSign+1)
                
                lengthOfOptional := 0
                if (SubStr(equalSignString, -4) = "&amp;") {
                    lengthOfOptional := 5
                }
                
                autocorrected := SubStr(equalSignString, 1, StrLen(equalSignString) - lengthOfOptional)
                return StrReplace(autocorrected, "+" , " ")
            }
        }
    } catch {}
    
    ; If no suggestion found, return empty string to indicate no correction
    return ""
}

URLDownloadToVarWithHeader(url){
    hObject:=ComObjCreate("WinHttp.WinHttpRequest.5.1")
    hObject.Open("GET",url)
    hObject.SetRequestHeader("User-Agent", "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; WOW64; Trident/5.0)")

    hObject.Send()
    return hObject.ResponseText
}

; ==

SortDimensions(dimensions) {
    if (dimensions[1] < dimensions[2]) {
        temp := dimensions[1], dimensions[1] := dimensions[2], dimensions[2] := temp
    }
    if (dimensions[2] < dimensions[3]) {
        temp := dimensions[2], dimensions[2] := dimensions[3], dimensions[3] := temp
    }
    if (dimensions[1] < dimensions[2]) {
        temp := dimensions[1], dimensions[1] := dimensions[2], dimensions[2] := temp
    }
    return dimensions
}

ParseDate(dateStr) {
    dateStr := StrReplace(dateStr, ".", "/")
    dateStr := StrReplace(dateStr, "-", "/")
    if (RegExMatch(dateStr, "(\d{1,2})/(\d{1,2})/(\d{2,4})", match)) {
        month := match1
        day := match2
        year := match3
        if (StrLen(year) == 2)
            year := "20" . year
        if (month > 12) {
            temp := month
            month := day
            day := temp
        }
        if (day > 31 || month > 12)
            return "Invalid Date"
        month := SubStr("0" . month, -1)
        day := SubStr("0" . day, -1)
        return year . month . day
    }
    return "Invalid Date"
}

CalculateDatesFromLMP(LMPDate) {
    FormatTime, LMPFormatted, %LMPDate%, MM/dd/yyyy
    EDDDate := DateCalc(LMPDate, 280)
    FormatTime, EDDFormatted, %EDDDate%, MM/dd/yyyy
    GA := DateCalc(A_Now, 0)
    GA -= LMPDate, days
    GAWeeks := Floor(GA / 7)
    GADays := Mod(GA, 7)
    return % "LMP: " . LMPFormatted . "`n"
        . "Estimated Delivery Date: " . EDDFormatted . "`n"
        . "Current Gestational Age: " . GAWeeks . " weeks " . GADays . " days"
}

CalculateDatesFromGA(WeeksGA, DaysGA, ReferenceDate) {
    FetusAgeDays := (WeeksGA * 7) + DaysGA
    LMPDate := DateCalc(ReferenceDate, -FetusAgeDays)
    FormatTime, LMPFormatted, %LMPDate%, MM/dd/yyyy
    EDDDate := DateCalc(LMPDate, 280)
    FormatTime, EDDFormatted, %EDDDate%, MM/dd/yyyy
    CurrentGA := DateCalc(A_Now, 0)
    CurrentGA -= LMPDate, days
    CurrentGAWeeks := Floor(CurrentGA / 7)
    CurrentGADays := Mod(CurrentGA, 7)
    FormatTime, ReferenceDateFormatted, %ReferenceDate%, MM/dd/yyyy
    return % "LMP: " . LMPFormatted . "`n"
        . "Estimated Delivery Date: " . EDDFormatted . "`n"
        . "Gestational Age as of " . ReferenceDateFormatted . ": " . WeeksGA . " weeks " . DaysGA . " days`n"
        . "Current Gestational Age: " . CurrentGAWeeks . " weeks " . CurrentGADays . " days"
}

DetermineMenstrualPhase(LMPDate) {
    DaysSinceLMP := A_Now
    DaysSinceLMP -= LMPDate, days
    CycleDay := Mod(DaysSinceLMP, 28) + 1
    FormatTime, LMPFormatted, %LMPDate%, MM/dd/yyyy
    Result := "LMP: " . LMPFormatted . "`n"
    Result .= "Current Cycle Day: " . CycleDay . "/28`n`n"
    if (CycleDay >= 1 && CycleDay <= 5) {
        Result .= "Menstrual Phase`nExpected endometrial stripe thickness: 1-4 mm"
    } else if (CycleDay >= 6 && CycleDay <= 13) {
        Result .= "Early Proliferative Phase`nExpected endometrial stripe thickness: 5-7 mm"
    } else if (CycleDay == 14) {
        Result .= "Ovulation`nExpected endometrial appearance: Trilaminar, approximately 11 mm"
    } else if (CycleDay >= 15 && CycleDay <= 28) {
        Result .= "Secretory Phase`nExpected endometrial stripe thickness: 7-16 mm"
    } else {
        Result .= "Error: Invalid cycle day calculated"
    }
    return Result
}

DateCalc(date, days) {
    date += %days%, days
    FormatTime, out, %date%, yyyyMMdd
    return out
}

; -------------------------------------------------------------
; CompareMeasurements: single function to compare old vs new
; -------------------------------------------------------------
CompareMeasurements(previous, current, previousDate, currentDate, input) {
    ; Process measurements
    prev := ProcessMeasurement(previous)
    curr := ProcessMeasurement(current)
    
    ; Check for dimension mismatch
    if (prev.dimensions.MaxIndex() != curr.dimensions.MaxIndex()) {
        return "Error: Mismatch in number of dimensions between previous and current measurements.`nPrevious: " . previous . "`nCurrent: " . current
    }

    ; Initialize result string
    result := input . "`n`n"
    result .= "Previous Date: " . (previousDate ? previousDate : "Not provided") . "`n"
    result .= "Current Date: " . (currentDate ? currentDate : "Assumed as today") . "`n`n"

    ; Process dimensions
    prevLongestDim := 0
    currLongestDim := 0
    
    Loop, % prev.dimensions.MaxIndex()
    {
        prevDim := prev.dimensions[A_Index]
        currDim := curr.dimensions[A_Index]
        
        ; Convert to cm if necessary
        prevDimCm := (prev.unit == "mm") ? prevDim / 10 : prevDim
        currDimCm := (curr.unit == "mm") ? currDim / 10 : currDim
        
        ; Track longest dimension
        prevLongestDim := (prevDimCm > prevLongestDim) ? prevDimCm : prevLongestDim
        currLongestDim := (currDimCm > currLongestDim) ? currDimCm : currLongestDim
        
        ; Calculate change
        change := (currDimCm / prevDimCm - 1) * 100
        
        ; Append to result
        result .= "Dimension " . A_Index . ": " 
                . Round(prevDim, 2) . " " . prev.unit 
                . " -> " 
                . Round(currDim, 2) . " " . curr.unit 
                . " (" . (change >= 0 ? "+" : "") . Round(change, 1) . "%)`n"
    }
    
    ; Calculate longest dimension change
    longestDimChange := (currLongestDim / prevLongestDim - 1) * 100
    result .= "`nLongest dimension change: " . (longestDimChange >= 0 ? "+" : "") . Round(longestDimChange, 1) . "%`n"
    
    ; Calculate volumes
    prevVolumeNum := CalculateVolume(prev.dimensions, prev.unit)
    currVolumeNum := CalculateVolume(curr.dimensions, curr.unit)
    
    ; Process volume calculations if valid
    if (prevVolumeNum != "Invalid input" and currVolumeNum != "Invalid input") {
        volumeChange := (currVolumeNum / prevVolumeNum - 1) * 100
        result .= "Volume change: " . (volumeChange >= 0 ? "+" : "") . Round(volumeChange, 1) . "%`n"
        result .= "Previous volume: " . FormatVolume(prevVolumeNum) . "`n"
        result .= "Current volume: " . FormatVolume(currVolumeNum) . "`n"
        
        ; Process dates and calculate time-based metrics
        if (previousDate) {
            parsedPreviousDate := ParseDate(previousDate)
            if (parsedPreviousDate == "Invalid Date") {
                return result . "`nError: Invalid previous date format."
            }
            
            if (!currentDate) {
                FormatTime, currentDate, , MM/dd/yyyy
            }
            parsedCurrentDate := ParseDate(currentDate)
            if (parsedCurrentDate == "Invalid Date") {
                return result . "`nError: Invalid current date format."
            }
            
            ; Calculate time difference
            timeDiff := DateDiff(parsedPreviousDate, parsedCurrentDate) / 365.25
            result .= "Time difference: " . Round(timeDiff, 2) . " years`n"
            
            ; Calculate growth metrics if time difference is positive
            if (timeDiff > 0) {
                ; Calculate doubling time in days
                doublingTime := CalculateDoublingTime(prevVolumeNum, currVolumeNum, timeDiff)
                doublingTimeDays := doublingTime * 365.25
                result .= "Doubling time: " . (doublingTime != "N/A" ? Round(doublingTimeDays, 0) . " days" : doublingTime) . "`n"
                
                ; Calculate exponential growth rate in % per year
                growthRate := CalculateExponentialGrowth(prevVolumeNum, currVolumeNum, timeDiff)
                result .= "Exponential Growth Rate: " . Round(growthRate * 100, 2) . "% per year"
            } else {
                result .= "Note: Doubling time and Growth Rate not calculated due to invalid time difference."
            }
        } else {
            result .= "Note: Doubling time and Growth Rate not calculated due to missing previous date."
        }
    } else {
        result .= "Error: Unable to calculate volume for one or both measurements."
    }
    
    return result
}

ProcessMeasurement(input) {
    dimensions := []
    RegExMatch(input, "i)(\d+(?:\.\d+)?)(?:\s*(?:x|\*)\s*(\d+(?:\.\d+)?))?(?:\s*(?:x|\*)\s*(\d+(?:\.\d+)?))?(?=\s*(cm|mm)?)", match)
    dimensions.Push(match1 + 0)  ; Convert to number to preserve all decimal places
    if (match2 != "")
        dimensions.Push(match2 + 0)
    if (match3 != "")
        dimensions.Push(match3 + 0)
    unit := (match4 != "") ? match4 : ((InStr(match1, ".") > 0 || InStr(match2, ".") > 0 || InStr(match3, ".") > 0) ? "cm" : "mm")
    return {dimensions: dimensions, unit: unit}
}

CalculateVolume(dimensions, unit) {
    static PI := 3.14159265358979
    
    if (dimensions.MaxIndex() == 1)
        volume := (4/3) * PI * (dimensions[1] / 2) ** 3  ; Sphere
    else if (dimensions.MaxIndex() == 2)
        volume := (4/3) * PI * (dimensions[1] / 2) * (dimensions[2] / 2) * ((dimensions[1] + dimensions[2]) / 4)  ; Ellipsoid with 3rd dim as average
    else if (dimensions.MaxIndex() == 3)
        volume := CalculateEllipsoidVolumeNumeric(dimensions[1] . " x " . dimensions[2] . " x " . dimensions[3], unit)
    else
        return "Invalid input"
    
    return (unit == "mm") ? volume / 1000 : volume  ; Convert to cm³ if necessary
}

FormatVolume(volume) {
    return (volume < 1) ? Round(volume * 1000, 1) . " cu-mm" : Round(volume, 1) . " cc"
}

CalculateDoublingTime(initialVolume, finalVolume, time) {
    growthRate := (finalVolume / initialVolume) ** (1 / time) - 1
    return (growthRate > 0) ? (Ln(2) / Ln(1 + growthRate)) : "N/A"
}

CalculateExponentialGrowth(initialVolume, finalVolume, time) {
    growthRate := Ln(finalVolume/initialVolume) / time
    return growthRate  ; Return as a decimal, will be converted to percentage in the main function
}

DateDiff(date1, date2) {
    EnvSub, date2, %date1%, Days
    return date2
}

CalculateEllipsoidVolumeNumeric(input, unit) {
    RegExNeedle := "\s*(\d+(?:\.\d+)?)\s*[x,]\s*(\d+(?:\.\d+)?)\s*[x,]\s*(\d+(?:\.\d+)?)\s*"
    if (RegExMatch(input, RegExNeedle, match)) {
        dimensions := [match1 + 0, match2 + 0, match3 + 0]  ; Convert to numbers
        dimensions := SortDimensions(dimensions)

        volume := (1/6) * 3.14159265358979323846 * (dimensions[1]) * (dimensions[2]) * (dimensions[3])
        return volume
    } else {
        return "Invalid input format"
    }
}
JoinDimensions(dimensions) {
    return Round(dimensions[1], 1) . (dimensions.Length() > 1 ? " x " . Round(dimensions[2], 1) : "") . (dimensions.Length() > 2 ? " x " . Round(dimensions[3], 1) : "")
}
; -------------------------------------------------------------------------
; 15) Contrast Premedication
; -------------------------------------------------------------------------
CalculateContrastPremedication() {
    defaultDateTime := GetDefaultDateTime()
    FormatTime, defaultDate, %defaultDateTime%, yyyyMMdd
    FormatTime, defaultTime, %defaultDateTime%, HH:mm

    ; Get current mouse position (relative to the entire desktop)
    CoordMode, Mouse, Screen
    MouseGetPos, mouseX, mouseY

    ; Determine which monitor the mouse is on
    SysGet, monitorCount, MonitorCount
    Loop, %monitorCount%
    {
        SysGet, monArea, Monitor, %A_Index%
        if (mouseX >= monAreaLeft && mouseX <= monAreaRight && mouseY >= monAreaTop && mouseY <= monAreaBottom)
        {
            activeMonitor := A_Index
            break
        }
    }

    ; Get dimensions of the active monitor
    SysGet, workArea, MonitorWorkArea, %activeMonitor%
    monitorWidth := workAreaRight - workAreaLeft
    monitorHeight := workAreaBottom - workAreaTop

    ; Calculate GUI dimensions and position
    guiWidth := 250
    guiHeight := 160
    xPos := mouseX + 10
    yPos := mouseY + 10

    ; Ensure the GUI doesn't go off-screen
    if (xPos + guiWidth > workAreaRight)
        xPos := workAreaRight - guiWidth
    if (yPos + guiHeight > workAreaBottom)
        yPos := workAreaBottom - guiHeight

    ; Create GUI with AlwaysOnTop option and position it near the mouse
    Gui, ContrastPremed:New, +AlwaysOnTop
    Gui, ContrastPremed:Add, Text, x10 y10, Scan Date:
    Gui, ContrastPremed:Add, DateTime, x10 y30 w120 vScanDate Choose%defaultDate%
    Gui, ContrastPremed:Add, Text, x140 y10, Scan Time:
    Gui, ContrastPremed:Add, DropDownList, x140 y30 w100 vScanTime, % CreateTimeList(defaultTime)
    Gui, ContrastPremed:Add, Radio, x10 y60 w200 vPremedProtocol Checked, Prednisone (13-7-1 hour)
    Gui, ContrastPremed:Add, Radio, x10 y80 w200, Methylprednisolone (12-2 hour)
    Gui, ContrastPremed:Add, Checkbox, x10 y100 w200 vIncludeDiphenhydramine Checked, Include Diphenhydramine
    Gui, ContrastPremed:Add, Button, x10 y130 w100 gCalculatePremedTiming, Calculate
    Gui, ContrastPremed:Add, Button, x120 y130 w100 gShowPremedDosages, Show Dosages
    Gui, ContrastPremed:Show, x%xPos% y%yPos% w%guiWidth% h%guiHeight%, Contrast Premedication

    return
}

GetDefaultDateTime() {
    defaultDateTime := A_Now
    EnvAdd, defaultDateTime, 13, Hours
    
    ; Extract hours and minutes
    FormatTime, hours, %defaultDateTime%, HH
    FormatTime, minutes, %defaultDateTime%, mm
    
    ; Round up to nearest 5 minutes
    minutes := Ceil(minutes / 5) * 5
    if (minutes = 60) {
        minutes := 0
        hours := hours + 1
    }
    
    ; Handle day change if hours exceed 23
    if (hours >= 24) {
        hours := Mod(hours, 24)
        EnvAdd, defaultDateTime, 1, Days
    }
    
    ; Format the time back into the datetime
    formattedTime := Format("{:02d}{:02d}00", hours, minutes)
    FormatTime, formattedDate, %defaultDateTime%, yyyyMMdd
    return formattedDate . formattedTime
}

CreateTimeList(defaultTime) {
    timeList := ""
    Loop, 24 {
        hour := A_Index - 1
        Loop, 12 {
            minute := (A_Index - 1) * 5
            time := Format("{:02d}:{:02d}", hour, minute)
            timeList .= time . "|"
        }
    }
    ; Remove the trailing pipe
    timeList := RTrim(timeList, "|")
    
    ; Set the default time
    if (defaultTime) {
        timeList := StrReplace(timeList, defaultTime, defaultTime . "||")
    }
    
    return timeList
}

GetScanDateTime(scanDate, scanTime) {
    if (scanDate = "" || scanTime = "") {
        return GetDefaultDateTime()
    }
    return CombineDateTime(scanDate, scanTime)
}

CalculatePremedTiming:
    Gui, ContrastPremed:Submit, NoHide
    scanDateTime := GetScanDateTime(ScanDate, ScanTime)
    if (!scanDateTime) {
        MsgBox, 0, Error, Please enter a valid date and time.
        return
    }
    CalculateAndShowPremedResult(scanDateTime, PremedProtocol, IncludeDiphenhydramine)
return

CombineDateTime(scanDate, scanTime) {
    FormatTime, formattedDate, %scanDate%, yyyyMMdd
    formattedTime := StrReplace(scanTime, ":")
    fullDateTime := formattedDate . formattedTime . "00"
    
    if (!IsValidDateTime(fullDateTime)) {
        return false
    }
    
    return fullDateTime
}

CalculateAndShowPremedResult(scanDateTime, premedProtocol, includeDiphenhydramine) {
    FormatTime, scanDateTimeFormatted, %scanDateTime%, MM/dd/yyyy hh:mm tt
    
    result := "Scan time: " . scanDateTimeFormatted . " Contrast Premedication Schedule:`n`n"
    
    if (premedProtocol = 1) {
        ; Prednisone-based protocol
        result .= FormatPremedTime(scanDateTime, -13, "13", premedProtocol, includeDiphenhydramine)
        result .= FormatPremedTime(scanDateTime, -7, "7", premedProtocol, includeDiphenhydramine)
        result .= FormatPremedTime(scanDateTime, -1, "1", premedProtocol, includeDiphenhydramine)
    } else {
        ; Methylprednisolone-based protocol
        result .= FormatPremedTime(scanDateTime, -12, "12", premedProtocol, includeDiphenhydramine)
        result .= FormatPremedTime(scanDateTime, -2, "2", premedProtocol, includeDiphenhydramine)
    }
    
    result .= "`nNote: Premedication regimens less than 4-5 hours in duration (oral or IV) have not been shown to be effective.`n"
    result .= "If a patient is unable to take oral medication, 200 mg hydrocortisone IV may be substituted for each dose of oral prednisone in the 13-7-1 premedication regimen.`n"
    
    if (ShowCitations) {
        result .= "`nCitation: ACR Manual on Contrast Media. 2023 American College of Radiology. https://www.acr.org/Clinical-Resources/Contrast-Manual`n"
    }
    
    ShowResult(result)
}

FormatPremedTime(scanTime, hoursOffset, label, protocol, includeDiphenhydramine) {
    premedTime := DateAdd(scanTime, hoursOffset, "hours")
    FormatTime, premedTimeFormatted, %premedTime%, MM/dd/yyyy hh:mm tt
    
    medication := GetMedicationInfo(label, protocol, includeDiphenhydramine)
    
    return label . " hours before (" . premedTimeFormatted . "):`n" . medication . "`n`n"
}

GetMedicationInfo(label, protocol, includeDiphenhydramine) {
    if (protocol = 1) {  ; Prednisone-based protocol
        medication := "- Prednisone 50 mg PO"
        if (label = "1" && includeDiphenhydramine) {
            medication .= "`n- Diphenhydramine 50 mg IV, IM, or PO"
        }
    } else {  ; Methylprednisolone-based protocol
        medication := "- Methylprednisolone 32 mg PO"
        if (label = "2" && includeDiphenhydramine) {
            medication .= "`n- Diphenhydramine 50 mg IV, IM, or PO"
        }
    }
    return medication
}

IsValidDateTime(dateTimeStr) {
    try {
        FormatTime, test, %dateTimeStr%, yyyy-MM-dd HH:mm:ss
        return true
    } catch {
        return false
    }
}

ShowPremedDosages:
    Gui, ContrastPremed:Submit, NoHide
    ShowPremedDosages(PremedProtocol, IncludeDiphenhydramine)
return

ShowPremedDosages(premedProtocol, includeDiphenhydramine) {
    dosages := "Contrast Premedication Dosages:`n`n"
    
    if (premedProtocol = 1) {
        dosages .= "Prednisone-based Protocol (13-7-1 hour):`n`n"
        dosages .= "13 hours before:`n- Prednisone 50 mg PO`n`n"
        dosages .= "7 hours before:`n- Prednisone 50 mg PO`n`n"
        dosages .= "1 hour before:`n- Prednisone 50 mg PO`n"
        if (includeDiphenhydramine) {
            dosages .= "- Diphenhydramine 50 mg IV, IM, or PO`n"
        }
    } else {
        dosages .= "Methylprednisolone-based Protocol (12-2 hour):`n`n"
        dosages .= "12 hours before:`n- Methylprednisolone 32 mg PO`n`n"
        dosages .= "2 hours before:`n- Methylprednisolone 32 mg PO`n"
        if (includeDiphenhydramine) {
            dosages .= "- Diphenhydramine 50 mg IV, IM, or PO`n"
        }
    }
    
    dosages .= "`nNotes:`n"
    dosages .= "- Premedication regimens less than 4-5 hours in duration (oral or IV) have not been shown to be effective.`n"
    dosages .= "- If a patient is unable to take oral medication, 200 mg hydrocortisone IV may be substituted for each dose of oral prednisone in the 13-7-1 premedication regimen.`n"
    dosages .= "- Diphenhydramine is considered optional. If a patient is allergic to diphenhydramine, an alternate anti-histamine without cross-reactivity may be considered, or the anti-histamine may be omitted.`n"
    dosages .= "- These dosages are based on the ACR Manual on Contrast Media. Please consult with a healthcare professional for patient-specific recommendations.`n"
    
    if (ShowCitations) {
        dosages .= "`nCitation: ACR Manual on Contrast Media. 2023 American College of Radiology. https://www.acr.org/Clinical-Resources/Contrast-Manual`n"
    }
    
    ShowResult(dosages)
}

DateAdd(datetime, value, unit) {
    EnvAdd, datetime, %value%, %unit%
    return datetime
}

; -------------------------------------------------------------------------
; 16) Fleischner Code
; -------------------------------------------------------------------------
class Nodule {
    __New(nString) {
        this.Description := nString
        this.Composition := ""
        this.Calcified := false
        this.mString := ""
        this.Units := ""
        this.Measurements := []
        this.HighRisk := false
        this.Location := ""
        this.Multiplicity := "single"
        this.Perifissural := false
        this.Morphology := ""
        this.GlobalHighRisk := false

        if (!this.ContainsNoduleReference(nString))
            throw Exception("No nodule reference", -1)
        
        this.ParseNoduleProperties(nString)
        this.ExtractMeasurements(nString)

        ; Handle cases where only the largest measurement is given for multiple nodules
        if (InStr(nString, "up to") and InStr(nString, "multiple")) {
            this.Multiplicity := "multiple"
        }
    }

    ContainsNoduleReference(nString) {
        noduleTerms := ["nodule", "nodules", "mass", "masses", "opacity", "opacities", "lesion", "lesions", "micronodule", "micronodules"]
        for _, term in noduleTerms {
            if (InStr(nString, term))
                return true
        }
        return false
    }

    ParseNoduleProperties(nString) {
        words := StrSplit(nString, A_Space)
        for index, word in words {
            if (this.IsMultiplicityWord(word))
                this.Multiplicity := "multiple"
            
            if (this.FuzzyMatch(word, "solid")) {
                if (this.Composition = "")
                    this.Composition := "solid"
                else if (this.Composition = "ground glass")
                    this.Composition := "part solid"
            }
            
            if (this.FuzzyMatch(word, "ground") and index < words.Length() and this.FuzzyMatch(words[index+1], "glass")) {
                if (this.Composition = "")
                    this.Composition := "ground glass"
                else if (this.Composition = "solid")
                    this.Composition := "part solid"
            }
            
            if (this.FuzzyMatch(word, "groundglass")) {
                if (this.Composition = "")
                    this.Composition := "ground glass"
                else if (this.Composition = "solid")
                    this.Composition := "part solid"
            }
            
            if ((this.FuzzyMatch(word, "part") and index < words.Length() and this.FuzzyMatch(words[index+1], "solid")) 
                or this.FuzzyMatch(word, "part-solid") or this.FuzzyMatch(word, "partsolid")) {
                this.Composition := "part solid"
            }
            
            if (this.FuzzyMatch(word, "calcified") or this.FuzzyMatch(word, "calcification") or this.FuzzyMatch(word, "calcifications"))
                this.Calcified := true
            
            if (this.FuzzyMatch(word, "noncalcified") or this.FuzzyMatch(word, "non-calcified"))
                this.Calcified := false

            if (this.FuzzyMatch(word, "emphysema") or this.FuzzyMatch(word, "fibrosis"))
                this.GlobalHighRisk := true

            lobeKeywords := ["upper", "middle", "lower", "lingula", "apical", "basal"]
            for _, keyword in lobeKeywords {
                if (this.FuzzyMatch(word, keyword)) {
                    this.Location .= " " . word
                }
            }
            if (this.FuzzyMatch(word, "right") or this.FuzzyMatch(word, "left")) {
                this.Location .= " " . word
            }

            if (this.FuzzyMatch(word, "perifissural") or this.FuzzyMatch(word, "fissure"))
                this.Perifissural := true

            morphologyKeywords := ["spiculated", "lobulated", "irregular", "smooth"]
            for _, keyword in morphologyKeywords {
                if (this.FuzzyMatch(word, keyword)) {
                    this.Morphology := keyword
                    break
                }
            }
        }
        this.Location := Trim(this.Location)

        if (InStr(nString, "multiple") or InStr(nString, "several") or InStr(nString, "numerous"))
            this.Multiplicity := "multiple"
    }

    IsMultiplicityWord(word) {
        multiplicityWords := ["nodules", "multiple", "several", "few", "numerous", "masses", "opacities", "lesions", "micronodules"]
        for _, w in multiplicityWords {
            if (this.FuzzyMatch(word, w))
                return true
        }
        return false
    }

    FuzzyMatch(word1, word2, threshold := 2) {
        return (LevenshteinDistance(word1, word2) <= threshold)
    }

	ExtractMeasurements(nString) {
	needle := "i)(?:up to|approximately|about|~)?\s*((?:\d*\.)?\d+)\s*(?:x\s*((?:\d*\.)?\d+))?\s*(?:x\s*((?:\d*\.)?\d+))?\s*([cm]m)"
		if (RegExMatch(nString, needle, match)) {
			this.mString := match
			this.Units := match4
			Loop, 3 {
				if (match%A_Index% != "")
					this.Measurements.Push(match%A_Index%)
			}
		}
		else {
			needle := "i)(?:up to|approximately|about|~)?\s*((?:\d*\.)?\d+)\s*([cm]m)"
			if (RegExMatch(nString, needle, match)) {
				this.mString := match
				this.Units := match2
				this.Measurements.Push(match1)
			}
			else if (InStr(nString, "micronodule") or InStr(nString, "micronodules")) {
				; Default size for micronodules when no measurement is given
				this.mString := "5 mm"
					this.Units := "mm"
						this.Measurements.Push(5)
				}
				else {
					; If no specific measurement found, look for any number followed by units
					needle := "i)((?:\d*\.)?\d+)\s*([cm]m)"
					if (RegExMatch(nString, needle, match)) {
						this.mString := match
						this.Units := match2
						this.Measurements.Push(match1)
					}
					else {
					throw Exception("No measurements found", -2)
					}
				}
			}
		; Store original measurements
		this.OriginalMeasurements := this.Measurements.Clone()
		this.OriginalUnits := this.Units
	}

    Size() {
        if (this.Measurements.Length() = 0)
            return 0
        
        totalSize := 0
        for _, measurement in this.Measurements {
            totalSize += measurement
        }
        averageSize := totalSize / this.Measurements.Length()

        ; Convert to mm if necessary
        if (this.Units = "cm")
            return averageSize * 10  ; Convert cm to mm
        return averageSize  ; Already in mm
    }
	
    UpdateMString() {
        if (this.OriginalMeasurements.Length() = 0)
            return

        sizes := []
        for _, measurement in this.OriginalMeasurements {
            sizes.Push(Format("{:.1f}", measurement))
        }
        
        this.mString := Join(sizes, " x ") . " " . this.OriginalUnits
    }

    Category() {
        s := this.Size()  ; s is now always in mm
        
        if (this.Composition == "solid" or this.Composition == "") {
            if (this.Multiplicity == "multiple") {
                if (s < 6)
                    return "multiple_solid_small"
                else if (s >= 6 and s <= 8)
                    return "multiple_solid_medium"
                else
                    return "multiple_solid_large"
            } else {
                if (s < 6)
                    return "single_solid_small"
                else if (s >= 6 and s <= 8)
                    return "single_solid_medium"
                else
                    return "single_solid_large"
            }
        } else if (this.Composition == "ground glass") {
            if (this.Multiplicity == "multiple") {
                if (s < 6)
                    return "multiple_subsolid_small"
                else
                    return "multiple_subsolid_large"
            } else {
                if (s < 6)
                    return "single_gg_small"
                else
                    return "single_gg_large"
            }
        } else if (this.Composition == "part solid") {
            if (this.Multiplicity == "multiple") {
                if (s < 6)
                    return "multiple_subsolid_small"
                else
                    return "multiple_subsolid_large"
            } else {
                if (s < 6)
                    return "single_ps_small"
                else
                    return "single_ps_large"
            }
        }
    }

    Recommendation() {
        category := this.Category()
        risk := this.HighRisk ? "high" : "low"
        
        if (recommendations.HasKey(category)) {
            if (IsObject(recommendations[category]))
                return recommendations[category][risk]
            else
                return recommendations[category]
        }
        return "Unable to determine recommendation."
    }
}

InitializeRecommendations() {
    ObjRawSet(recommendations, "single_solid_small", {low: "No routine follow-up.", high: "Optional CT at 12 months."})
    ObjRawSet(recommendations, "single_solid_medium", {low: "CT at 6-12 months, then consider CT at 18-24 months.", high: "CT at 6-12 months, then CT at 18-24 months."})
    ObjRawSet(recommendations, "single_solid_large", {low: "Consider CT at 3 months, PET/CT, or tissue sampling.", high: "Consider CT at 3 months, PET/CT, or tissue sampling."})
    ObjRawSet(recommendations, "multiple_solid_small", {low: "No routine follow-up.", high: "Optional CT at 12 months."})
    ObjRawSet(recommendations, "multiple_solid_medium", {low: "CT at 3-6 months, then consider CT at 18-24 months.", high: "CT at 3-6 months, then at 18-24 months."})
    ObjRawSet(recommendations, "multiple_solid_large", {low: "CT at 3-6 months, then consider CT at 18-24 months.", high: "CT at 3-6 months, then at 18-24 months."})
    ObjRawSet(recommendations, "single_gg_small", "No routine follow-up.")
    ObjRawSet(recommendations, "single_gg_large", "CT at 6-12 months to confirm persistence, then every 2 years until 5 years.")
    ObjRawSet(recommendations, "single_ps_small", "No routine follow-up.")
    ObjRawSet(recommendations, "single_ps_large", "CT at 3-6 months to confirm persistence. If unchanged and solid component remains <6mm, annual CT should be performed for 5 years.")
    ObjRawSet(recommendations, "multiple_subsolid_small", "CT at 3-6 months. If stable, consider CT at 2 and 4 years.")
	ObjRawSet(recommendations, "multiple_subsolid_large", "CT at 3-6 months. Subsequent management based on the most suspicious nodule(s).")
}

Join(arr, sep) {
    result := ""
    for index, element in arr {
        if (index > 1)
            result .= sep
        result .= element
    }
    return result
}

PreprocessTextNodules(text) {
    sentences := []
    lines := StrSplit(text, "`n", "`r")
    for _, line in lines {
        ; Remove bullet points
        line := RegExReplace(line, "^\s*[\*\-•]\s*", "")
        
        ; Split by periods, but be careful with measurements
        parts := StrSplit(line, ".")
        for index, part in parts {
            part := Trim(part)
            ; Check if this part is a continuation of a measurement
            if (index > 1 && RegExMatch(parts[index-1], "\d+$") && RegExMatch(part, "^\d+"))
                sentences[sentences.Length()] .= "." . part
            else if (part != "")
                sentences.Push(part)
        }
    }
    return sentences
}

ProcessNodules(text) {
    result := ""
    sentences := PreprocessTextNodules(text)
    nodules := []
    globalHighRisk := CheckHighRiskConditions(text)
    isMultiple := false

    for _, sentence in sentences {
        multipleNodulesProbability := CalculateMultipleNodulesProbability(sentence)
        isMultiple := (multipleNodulesProbability >= 0.45) or isMultiple  ; Adjusted threshold and accumulate

        noduleDescriptions := SplitNoduleDescriptions(sentence)
        for _, description in noduleDescriptions {
            try {
                nodule := new Nodule(description)
                nodule.GlobalHighRisk := globalHighRisk
                nodule.Multiplicity := isMultiple ? "multiple" : "single"
                nodules.Push(nodule)
            } catch e {
                ; Skip invalid nodule descriptions
            }
        }
    }

    if (nodules.Length() = 0) {
        ; Check for general mention of micronodules or multiple nodules
        if (InStr(text, "micronodule") or InStr(text, "micronodules") or (InStr(text, "multiple") and InStr(text, "nodule"))) {
            nodule := new Nodule("Multiple nodules measuring up to 5 mm")
            nodule.GlobalHighRisk := globalHighRisk
            nodule.Multiplicity := "multiple"
            nodules.Push(nodule)
        } else {
            result := "Error: No valid nodules found."
            return result
        }
    }

    ; Find the most significant nodule
    mostSignificantNodule := nodules[1]
    for _, nodule in nodules {
        nodule.UpdateMString()  ; Update mString for each nodule
        if (nodule.Size() > mostSignificantNodule.Size() or (nodule.Size() = mostSignificantNodule.Size() and IsMoreSignificant(nodule, mostSignificantNodule))) {
            mostSignificantNodule := nodule
        }
    }

    ; Ensure multiplicity is set correctly
    if (nodules.Length() > 1 or isMultiple or InStr(text, "multiple") or InStr(text, "micronodules")) {
        mostSignificantNodule.Multiplicity := "multiple"
    }

    ; Generate result string
    result := text . "`n`n"  ; Preserve input text at the top
    result .= "Fleischner 2017 assessment:`n"
    if (mostSignificantNodule.Multiplicity == "multiple") {
        result .= "Multiple pulmonary nodules are present. The most significant nodule forming the basis of follow-up:`n"
    } else {
        result .= "A solitary pulmonary nodule is described forming the basis of follow-up:`n"
    }
    esult .= "- Location: " . (mostSignificantNodule.Location ? mostSignificantNodule.Location : "Not specified") . "`n"
    result .= "- Extracted (or inferred) Size: " . mostSignificantNodule.mString . "`n"
    result .= "- Fleischner Size (mm): " . Format("{:.1f} mm", mostSignificantNodule.Size()) . "`n"
    result .= "- Composition (or inferred): " . (mostSignificantNodule.Composition ? mostSignificantNodule.Composition : "solid") . "`n"
    if (mostSignificantNodule.Calcified)
        result .= "- Calcification: Present`n"
    if (mostSignificantNodule.Morphology)
        result .= "- Morphology: " . mostSignificantNodule.Morphology . "`n"
		
    currentDate := A_Now

    result .= "`nFLEISCHNER SOCIETY RECOMMENDATION:`n"
    if ((globalHighRisk || mostSignificantNodule.Morphology == "spiculated") && !mostSignificantNodule.Calcified) {
        mostSignificantNodule.HighRisk := true
        recommendation := mostSignificantNodule.Recommendation()
        result .= AddFollowUpDates(recommendation, currentDate)
    } else {
        if (!mostSignificantNodule.Calcified && mostSignificantNodule.Composition == "solid" || mostSignificantNodule.Composition == "") {
            lowRiskRec := mostSignificantNodule.Recommendation()
            result .= "For low-risk patients: " . AddFollowUpDates(lowRiskRec, currentDate) . "`n"
            mostSignificantNodule.HighRisk := true
            highRiskRec := mostSignificantNodule.Recommendation()
            result .= "`nFor high-risk patients: " . AddFollowUpDates(highRiskRec, currentDate) . "`n"
		} else if (!mostSignificantNodule.Calcified && (mostSignificantNodule.Composition == "part solid" || mostSignificantNodule.Composition == "ground glass")) {
			recommendation := mostSignificantNodule.Recommendation()
			result .= AddFollowUpDates(recommendation, currentDate)
		} else {
            result .= "Incidental calcified nodules do not typically require routine follow up.`n"
        }
    }

    if (globalHighRisk || mostSignificantNodule.Morphology == "spiculated")
        result .= "`n`nNote: Patient has risk factors or morphologic characteristics that may increase the risk of lung cancer."

    if (ShowCitations)
        result .= "`n`nCitation: MacMahon H, Naidich DP, Goo JM, et al. Guidelines for Management of Incidental Pulmonary Nodules Detected on CT Images: From the Fleischner Society 2017. Radiology. 2017;284(1):228-243. doi:10.1148/radiol.2017161659"
   
    return result
}

IsMoreSignificant(nodule1, nodule2) {
    ; Order of significance: solid > part solid > ground glass
    compositionOrder := {"solid": 3, "part solid": 2, "ground glass": 1}
    return compositionOrder[nodule1.Composition] > compositionOrder[nodule2.Composition]
}

SplitNoduleDescriptions(text) {
    descriptions := []
    ; Modified regex to capture nodule descriptions more accurately, including cases where size appears before location and "nodule"
    needle := "i)(?:(?:\d+(?:\.\d+)?\s*(?:x\s*\d+(?:\.\d+)?)*\s*(?:mm|cm))\s*(?:[a-z\s]+\s+)?(?:pulmonary\s+)?(?:nodule|mass|opacity|lesion)s?|(?:(?:multiple\s+)?(?:nodule|mass|opacity|lesion)s?(?:\s+(?:up to|approximately|about|~)?)?\s*(?:\d+(?:\.\d+)?\s*(?:x\s*\d+(?:\.\d+)?)*\s*(?:mm|cm))?|(?:(?:up to|approximately|about|~)?\s*\d+(?:\.\d+)?\s*(?:x\s*\d+(?:\.\d+)?)*\s*(?:mm|cm))?\s*(?:(?:solid|ground glass|part[- ]solid|calcified|noncalcified|spiculated|lobulated|irregular|smooth)?\s*(?:pulmonary\s+)?(?:nodule|mass|opacity|lesion)s?)))(?:\s*\([^)]+\))?"
    
    ; Split the text into lines
    lines := StrSplit(text, "`n", "`r")
    for _, line in lines {
        pos := 1
        while (pos := RegExMatch(line, needle, match, pos)) {
            descriptions.Push(Trim(match))
            pos += StrLen(match)
        }
    }
    
    ; If no specific nodule descriptions found, check for general mentions of nodules
    if (descriptions.Length() = 0) {
        if (RegExMatch(text, "i)(?:multiple\s+)?nodules?.*(?:measure|up to).*(\d+(?:\.\d+)?)\s*(?:x\s*\d+(?:\.\d+)?)*\s*(mm|cm)", match)) {
            descriptions.Push(match)
        } else if (InStr(text, "micronodule") or InStr(text, "micronodules")) {
            descriptions.Push("Multiple micronodules")
        } else if (InStr(text, "multiple") and InStr(text, "nodule")) {
            descriptions.Push("Multiple nodules")
        }
    }
    
    ; If still no descriptions found, return the entire text
    if (descriptions.Length() = 0) {
        descriptions.Push(text)
    }
    
    return descriptions
}

AddFollowUpDates(recommendation, currentDate) {
    followUpPeriods := []
    
    if (InStr(recommendation, "3 months"))
        followUpPeriods.Push({min: 90, max: 90, text: "3 months"})
	if (InStr(recommendation, "3-6 months"))
        followUpPeriods.Push({min: 90, max: 183, text: "3-6 months"})
    if (InStr(recommendation, "6-12 months"))
        followUpPeriods.Push({min: 180, max: 365, text: "6-12 months"})
    if (InStr(recommendation, "18-24 months"))
        followUpPeriods.Push({min: 540, max: 730, text: "18-24 months"})
	if (InStr(recommendation, "2 years until 5 years"))
        followUpPeriods.Push({min: 1095, max: 1825, text: "2 years until 5 years"})
    
    if (followUpPeriods.Length() > 0) {
        recommendation .= "`nFollow-up dates:"
        for _, period in followUpPeriods {
            minDate := DateCalc(currentDate, period.min)
            maxDate := DateCalc(currentDate, period.max)
            FormatTime, formattedMinDate, %minDate%, MMMM yyyy
            FormatTime, formattedMaxDate, %maxDate%, MMMM yyyy
			FormatTime, formattedCurrentDate, %currentDate%, MMMM yyyy
            if (formattedMinDate != formattedMaxDate)
                recommendation .= " " . period.text . " is " . formattedMinDate . " to " . formattedMaxDate . " from " . formattedCurrentDate . ". "
            else
                recommendation .= " " . period.text . " is " . formattedMinDate . " from " . formattedCurrentDate . ". "
        }
    }
    
    return recommendation
}

CheckHighRiskConditions(text) {
    StringLower text, text ; Convert to lowercase
    highRiskTerms := ["emphysema", "fibrosis", "fibrotic", "emphysematous"]
    
    for _, term in highRiskTerms {
        words := StrSplit(text, A_Space)
        for _, word in words {
            if (FuzzyMatchGlobal(word, term))
                return true
        }
    }
    return false
}

FuzzyMatchGlobal(word1, word2, threshold := 2) {
    return (LevenshteinDistance(word1, word2) <= threshold)
}

CalculateMultipleNodulesProbability(text) {
    probability := 0
    
    ; Check for explicit mentions of multiple nodules
    if (InStr(text, "nodules") or InStr(text, "multiple") or InStr(text, "several") or InStr(text, "few"))
        probability += 0.7
    
    ; Check for scattered micronodules
    if (InStr(text, "scattered") and InStr(text, "micronodules"))
        probability += 0.6
    
    ; Check for multiple measurements
    measurementCount := 0
    pos := 1
    while (pos := RegExMatch(text, "i)(\d+(?:\.\d+)?\s*(mm|cm))", match, pos + StrLen(match)))
        measurementCount++
    
    if (measurementCount > 1)
        probability += 0.5  ; Increased weight for multiple measurements
    
    ; Check for multiple locations
    lobeKeywords := ["upper", "middle", "lower", "lingula", "apical", "basal", "right", "left"]
    uniqueLocations := {}
    for _, keyword in lobeKeywords {
        if (InStr(text, keyword))
            uniqueLocations[keyword] := true
    }
    
    if (uniqueLocations.Count() > 1)
        probability += 0.3  ; Increased weight for multiple locations
    
    ; Count number of times "nodule" (singular) is mentioned
    noduleCount := 0
    pos := 1
    while (pos := InStr(text, "nodule", false, pos))
    {
        noduleCount++
        pos += 6  ; length of "nodule"
    }
    
    if (noduleCount > 1)
        probability += 0.4  ; Add probability if "nodule" is mentioned multiple times
    
    return (probability > 1) ? 1 : probability
}
; ------------------------------------------
; Go to MESA website and get informatoin
; ------------------------------------------
DownloadToString(url, postData := "") {
    static INTERNET_FLAG_RELOAD := 0x80000000
    static INTERNET_FLAG_SECURE := 0x00800000
    static SECURITY_FLAG_IGNORE_UNKNOWN_CA := 0x00000100
    
    hModule := DllCall("LoadLibrary", "Str", "wininet.dll", "Ptr")
    if (!hModule)
        return "Error: Failed to load wininet.dll. Error code: " . A_LastError
    
    hInternet := DllCall("wininet\InternetOpenA", "Str", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/93.0.4577.63 Safari/537.36", "UInt", 1, "Ptr", 0, "Ptr", 0, "UInt", 0, "Ptr")
    if (!hInternet) {
        DllCall("FreeLibrary", "Ptr", hModule)
        return "Error: InternetOpen failed. Error code: " . A_LastError
    }
    
    hConnect := DllCall("wininet\InternetConnectA", "Ptr", hInternet, "Str", "www.mesa-nhlbi.org", "UShort", 443, "Ptr", 0, "Ptr", 0, "UInt", 3, "UInt", 0, "Ptr", 0, "Ptr")
    if (!hConnect) {
        DllCall("wininet\InternetCloseHandle", "Ptr", hInternet)
        DllCall("FreeLibrary", "Ptr", hModule)
        return "Error: InternetConnect failed. Error code: " . A_LastError
    }
    
    flags := INTERNET_FLAG_RELOAD | INTERNET_FLAG_SECURE | SECURITY_FLAG_IGNORE_UNKNOWN_CA
    hRequest := DllCall("wininet\HttpOpenRequestA", "Ptr", hConnect, "Str", (postData ? "POST" : "GET"), "Str", "/Calcium/input.aspx", "Str", "HTTP/1.1", "Ptr", 0, "Ptr", 0, "UInt", flags, "Ptr", 0, "Ptr")
    if (!hRequest) {
        DllCall("wininet\InternetCloseHandle", "Ptr", hConnect)
        DllCall("wininet\InternetCloseHandle", "Ptr", hInternet)
        DllCall("FreeLibrary", "Ptr", hModule)
        return "Error: HttpOpenRequest failed. Error code: " . A_LastError
    }
    
    headers := "Content-Type: application/x-www-form-urlencoded`r`n"
             . "User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/93.0.4577.63 Safari/537.36`r`n"
             . "Accept: text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9`r`n"
             . "Accept-Language: en-US,en;q=0.9`r`n"
    DllCall("wininet\HttpAddRequestHeadersA", "Ptr", hRequest, "Str", headers, "UInt", -1, "UInt", 0x10000000, "Int")
    
    VarSetCapacity(buffer, 8192, 0)
    if (postData) {
        VarSetCapacity(postDataBuffer, StrLen(postData), 0)
        StrPut(postData, &postDataBuffer, "UTF-8")
        result := DllCall("wininet\HttpSendRequestA", "Ptr", hRequest, "Ptr", 0, "UInt", 0, "Ptr", &postDataBuffer, "UInt", StrLen(postData), "Int")
    } else {
        result := DllCall("wininet\HttpSendRequestA", "Ptr", hRequest, "Ptr", 0, "UInt", 0, "Ptr", 0, "UInt", 0, "Int")
    }
    
    if (!result) {
        errorCode := A_LastError
        DllCall("wininet\InternetCloseHandle", "Ptr", hRequest)
        DllCall("wininet\InternetCloseHandle", "Ptr", hConnect)
        DllCall("wininet\InternetCloseHandle", "Ptr", hInternet)
        DllCall("FreeLibrary", "Ptr", hModule)
        return "Error: HttpSendRequest failed. Error code: " . errorCode
    }
    
    VarSetCapacity(responseText, 1024*1024)  ; Allocate 1MB for the response
    bytesRead := 0
    totalBytesRead := 0
    
    Loop {
        result := DllCall("wininet\InternetReadFile", "Ptr", hRequest, "Ptr", &buffer, "UInt", 8192, "Ptr", &bytesRead, "Int")
        bytesRead := NumGet(bytesRead, 0, "UInt")
        if (bytesRead == 0)
            break
        DllCall("RtlMoveMemory", "Ptr", &responseText + totalBytesRead, "Ptr", &buffer, "Ptr", bytesRead)
        totalBytesRead += bytesRead
    }
    
    responseText := StrGet(&responseText, totalBytesRead, "UTF-8")
    
    DllCall("wininet\InternetCloseHandle", "Ptr", hRequest)
    DllCall("wininet\InternetCloseHandle", "Ptr", hConnect)
    DllCall("wininet\InternetCloseHandle", "Ptr", hInternet)
    DllCall("FreeLibrary", "Ptr", hModule)
    
    return responseText
}


GetRaceValue(race) {
    switch race {
        case "White": return 3
        case "Black": return 0
        case "Hispanic": return 2
        case "Chinese": return 1
    }
}

GetSexValue(sex) {
    switch sex {
        case "Male": return 1
        case "Female": return 0
    }
}

GetLastErrorMessage(errorCode) {
    VarSetCapacity(msg, 1024)
    DllCall("FormatMessage"
        , "UInt", 0x1000      ; FORMAT_MESSAGE_FROM_SYSTEM
        , "Ptr", 0
        , "UInt", errorCode
        , "UInt", 0           ; Default language
        , "Str", msg
        , "UInt", 1024
        , "Ptr", 0)
    return msg
}

SendHttpRequest(url, postData := "") {
    try {
        whr := ComObject("WinHttp.WinHttpRequest.5.1")
        whr.Open(postData ? "POST" : "GET", url, true)
        whr.SetRequestHeader("Content-Type", "application/x-www-form-urlencoded")
        whr.SetRequestHeader("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/93.0.4577.63 Safari/537.36")
        whr.Send(postData)
        whr.WaitForResponse()
        return whr.ResponseText
    }
    catch e {
        return "Error: " . e.message
    }
}

UrlEncode(str) {
    oldFormat := A_FormatInteger
    SetFormat, Integer, Hex
    VarSetCapacity(var, StrPut(str, "UTF-8"), 0)
    StrPut(str, &var, "UTF-8")
    StringLower, str, str
    Loop
    {
        code := NumGet(var, A_Index - 1, "UChar")
        If (!code)
            break
        If (code >= 0x30 && code <= 0x39 ; 0-9
            || code >= 0x41 && code <= 0x5A ; A-Z
            || code >= 0x61 && code <= 0x7A) ; a-z
            result .= Chr(code)
        Else
            result .= "%" . SubStr(code + 0x100, -1)
    }
    SetFormat, Integer, %oldFormat%[]
    return result
}
; ------------------------------------------
; Levenshtein Distance for Fleischner
; ------------------------------------------
LevenshteinDistance(s, t) {
    m := StrLen(s)
    n := StrLen(t)
    d := []

    Loop, % m + 1
    {
        d[A_Index] := []
        d[A_Index, 1] := A_Index - 1
    }

    Loop, % n + 1
        d[1, A_Index] := A_Index - 1

    Loop, % m
    {
        i := A_Index
        Loop, % n
        {
            j := A_Index
            cost := (SubStr(s, i, 1) = SubStr(t, j, 1)) ? 0 : 1
            d[i+1, j+1] := Min(d[i, j+1] + 1, d[i+1, j] + 1, d[i, j] + cost)
        }
    }

    return d[m+1, n+1]
}

; -------------------------------------------------------------
; References Management
; -------------------------------------------------------------
CreateReferencesMenu() {
    Menu, ReferencesMenu, Add
    Menu, ReferencesMenu, DeleteAll
    
    ; Add management options at top
    Menu, ReferencesMenu, Add, Add Reference..., AddReference
    Menu, ReferencesMenu, Add, Remove References..., RemoveReferences
    Menu, ReferencesMenu, Add  ; Separator
    
    ; Create sorted array of references
    refs := []
    for name, data in g_References {
        refs.Push({name: name, uses: data.uses})
    }
    
    ; Sort references by usage count (descending)
    SortReferencesByUse(refs)
    
    ; Add sorted references to menu
    count := 0
    for _, ref in refs {
        if (count >= g_MaxReferences)
            break
        Menu, ReferencesMenu, Add, % ref.name, OpenReference
        count++
    }
}

; Helper function to sort references by usage count
SortReferencesByUse(ByRef refs) {
    ; Create temporary arrays for sorting
    tempArray := []
    sortedArray := []
    
    ; Copy all items to temp array
    for index, ref in refs {
        tempArray.Push(ref)
    }
    
    ; Sort by uses (descending)
    while (tempArray.Length() > 0) {
        highestUses := -1
        highestIndex := 0
        for index, ref in tempArray {
            if (ref.uses > highestUses) {
                highestUses := ref.uses
                highestIndex := index
            }
        }
        sortedArray.Push(tempArray[highestIndex])
        tempArray.RemoveAt(highestIndex)
    }
    
    ; Clear and repopulate refs with sorted items
    refs := []
    for _, ref in sortedArray {
        refs.Push(ref)
    }
    
    return refs
}

AddReference() {
    global RefName, RefPath
    
    ; Get current mouse position
    CoordMode, Mouse, Screen
    MouseGetPos, mouseX, mouseY
    
    ; Determine which monitor the mouse is on
    SysGet, monitorCount, MonitorCount
    Loop, %monitorCount%
    {
        SysGet, monArea, Monitor, %A_Index%
        if (mouseX >= monAreaLeft && mouseX <= monAreaRight && mouseY >= monAreaTop && mouseY <= monAreaBottom)
        {
            activeMonitor := A_Index
            break
        }
    }
    
    ; Get dimensions of the active monitor
    SysGet, workArea, MonitorWorkArea, %activeMonitor%
    monitorWidth := workAreaRight - workAreaLeft
    monitorHeight := workAreaBottom - workAreaTop
    
    ; Calculate GUI dimensions and position
    guiWidth := 300
    guiHeight := 150
    xPos := mouseX + 10
    yPos := mouseY + 10
    
    ; Ensure the GUI doesn't go off-screen
    if (xPos + guiWidth > workAreaRight)
        xPos := workAreaRight - guiWidth
    if (yPos + guiHeight > workAreaBottom)
        yPos := workAreaBottom - guiHeight
    
    Gui, AddRef:New, +AlwaysOnTop
    Gui, AddRef:Add, Text,, Enter reference details:
    Gui, AddRef:Add, Edit, vRefName w280, Reference Name
    Gui, AddRef:Add, Edit, vRefPath w280, Location (URL or File Path)
    Gui, AddRef:Add, Button, gBrowseReference w90, Browse...
    Gui, AddRef:Add, Button, gSaveReference w90, Save
    Gui, AddRef:Add, Button, x+10 w90 gAddRefGuiClose, Cancel
    Gui, AddRef:Show, x%xPos% y%yPos% w%guiWidth% h%guiHeight%, Add Reference
}

BrowseReference() {
    FileSelectFile, filePath, 3, , Select File, Documents (*.pdf; *.doc; *.docx; *.xls; *.xlsx, *.ppt, *.pptx)
    if (filePath) {
        GuiControl, AddRef:, RefPath, %filePath%
        SplitPath, filePath, fileName
        GuiControl, AddRef:, RefName, %fileName%
    }
}

SaveReference() {
    global RefName, RefPath, g_References
    Gui, AddRef:Submit, NoHide
    
    RefName := Trim(RefName)  ; Trim any whitespace
    RefPath := Trim(RefPath)  ; Trim any whitespace
    
    if (RefName = "") {
        MsgBox, Please enter a reference name.
        return
    }
    
    if (RefPath = "") {
        MsgBox, Please enter a URL or file path.
        return
    }

    ; Automatically determine if it's a file or URL
    if (FileExist(RefPath)) {
        ; It's a file
        SplitPath, RefPath, fileName, , fileExt
        if !IsValidFileType(fileExt) {
            MsgBox, Invalid file type. Only PDF, PPT, PPTX, XLS, XLSX, DOC, and DOCX files are allowed.
            return
        }
        
        referencesDir := A_ScriptDir . "\References"
        if (!FileExist(referencesDir))
            FileCreateDir, %referencesDir%
        
        destPath := referencesDir . "\" . fileName
        if (FileExist(destPath)) {
            MsgBox, 4, File Exists, File already exists. Would you like to create a new copy?
            IfMsgBox Yes
            {
                destPath := GetUniqueFilePath(referencesDir, fileName)
                FileCopy, %RefPath%, %destPath%
            }
        } else {
            FileCopy, %RefPath%, %destPath%
        }
        
        g_References[RefName] := {type: "file", path: destPath, uses: 0}
    } else if (IsValidURL(RefPath)) {
        ; It's a URL
        g_References[RefName] := {type: "url", path: RefPath, uses: 0}
    } else {
        MsgBox, The location must be either a valid file path or URL.
        return
    }
    
    SavePreferencesToFile()
    CreateReferencesMenu()  ; Refresh the menu after adding new reference
    Gui, AddRef:Destroy
}

MapReference() {
    ; Get current mouse position
    CoordMode, Mouse, Screen
    MouseGetPos, mouseX, mouseY
    
    ; Determine which monitor the mouse is on
    SysGet, monitorCount, MonitorCount
    Loop, %monitorCount%
    {
        SysGet, monArea, Monitor, %A_Index%
        if (mouseX >= monAreaLeft && mouseX <= monAreaRight && mouseY >= monAreaTop && mouseY <= monAreaBottom)
        {
            activeMonitor := A_Index
            break
        }
    }
    
    ; Get dimensions of the active monitor
    SysGet, workArea, MonitorWorkArea, %activeMonitor%
    monitorWidth := workAreaRight - workAreaLeft
    monitorHeight := workAreaBottom - workAreaTop
    
    ; Calculate GUI dimensions and position
    guiWidth := 300
    guiHeight := 180
    xPos := mouseX + 10
    yPos := mouseY + 10
    
    ; Ensure the GUI doesn't go off-screen
    if (xPos + guiWidth > workAreaRight)
        xPos := workAreaRight - guiWidth
    if (yPos + guiHeight > workAreaBottom)
        yPos := workAreaBottom - guiHeight
    
    Gui, MapRef:New, +AlwaysOnTop
    Gui, MapRef:Add, Text,, Select file to map:
    Gui, MapRef:Add, Edit, vRefName w280, Reference Name
    Gui, MapRef:Add, Edit, vRefPath w280, File Path
    Gui, MapRef:Add, Button, gBrowseMapReference w90, Browse...
    Gui, MapRef:Add, Button, gSaveMapReference w90, Save
    Gui, MapRef:Add, Button, x+10 w90 gMapRefGuiClose, Cancel
    Gui, MapRef:Show, x%xPos% y%yPos% w%guiWidth% h%guiHeight%, Map Reference
}

SaveMapReference() {
    Gui, MapRef:Submit, NoHide
    
    if (RefName = "") {
        MsgBox, Please enter a reference name.
        return
    }
    
    if (!FileExist(RefPath)) {
        MsgBox, File does not exist.
        return
    }
    
    SplitPath, RefPath, , , fileExt
    if !IsValidFileType(fileExt) {
        MsgBox, Invalid file type. Only PDF, XLS, XLSX, DOC, and DOCX files are allowed.
        return
    }
    
    g_References[RefName] := {type: "mapped", path: RefPath}
    SavePreferencesToFile()
    Gui, MapRef:Destroy
}

RemoveReferences() {
    global RefType, RefName, RefPath
    
    ; Get current mouse position
    CoordMode, Mouse, Screen
    MouseGetPos, mouseX, mouseY
    
    ; Determine which monitor the mouse is on
    SysGet, monitorCount, MonitorCount
    Loop, %monitorCount%
    {
        SysGet, monArea, Monitor, %A_Index%
        if (mouseX >= monAreaLeft && mouseX <= monAreaRight && mouseY >= monAreaTop && mouseY <= monAreaBottom)
        {
            activeMonitor := A_Index
            break
        }
    }
    
    ; Get dimensions of the active monitor
    SysGet, workArea, MonitorWorkArea, %activeMonitor%
    monitorWidth := workAreaRight - workAreaLeft
    monitorHeight := workAreaBottom - workAreaTop
    
    ; Calculate GUI dimensions and position
    guiWidth := 400
    guiHeight := 300
    xPos := mouseX + 10
    yPos := mouseY + 10
    
    ; Ensure the GUI doesn't go off-screen
    if (xPos + guiWidth > workAreaRight)
        xPos := workAreaRight - guiWidth
    if (yPos + guiHeight > workAreaBottom)
        yPos := workAreaBottom - guiHeight
    
    Gui, RemoveRef:New, +AlwaysOnTop
    Gui, RemoveRef:Add, ListView, vRefList w380 h230 gRefListHandler -Multi, Reference|Type|Path
    
    ; In v1.1, objects are handled with key-value pairs
    for name, data in g_References {
        LV_Add("", name, data.type, data.path)
    }
    
    Gui, RemoveRef:Add, Button, w120 gRemoveSelectedReferences, Remove Selected
    Gui, RemoveRef:Add, Button, x+10 w120 gRemoveRefGuiClose, Close
    Gui, RemoveRef:Show, x%xPos% y%yPos% w%guiWidth% h%guiHeight%, Remove References
}

RemoveSelectedReferences() {
    ; Get the position of the Remove References GUI
    Gui, RemoveRef:+LastFound
    WinGetPos, guiX, guiY, guiW, guiH
    
    selectedRows := []
    row := 0
    
    ; Collect all selected rows first
    while (row := LV_GetNext(row)) {
        selectedRows.Insert(row)
    }
    
    removedCount := 0
    
    ; Process selected rows in reverse order to maintain correct indices
    Loop % selectedRows.MaxIndex() {
        row := selectedRows[selectedRows.MaxIndex() - A_Index + 1]
        LV_GetText(name, row, 1)
        
        ; Calculate position for the MsgBox relative to the GUI
        msgBoxX := guiX + (guiW / 2) - 150
        msgBoxY := guiY + (guiH / 2) - 50
        
        ; Ensure message box stays on screen
        SysGet, workArea, MonitorWorkArea
        if (msgBoxX < workAreaLeft)
            msgBoxX := workAreaLeft
        if (msgBoxX + 300 > workAreaRight)
            msgBoxX := workAreaRight - 300
        if (msgBoxY < workAreaTop)
            msgBoxY := workAreaTop
        if (msgBoxY + 100 > workAreaBottom)
            msgBoxY := workAreaBottom - 100
        
        ; Position and set AlwaysOnTop for the message box
        SetTimer, MoveMsgBoxTop, 50
        MsgBox, 4, Confirm Removal, Are you sure you want to remove "%name%"?
        SetTimer, MoveMsgBoxTop, Off
        
        IfMsgBox Yes
        {
            g_References.Delete(name)
            removedCount++
        }
    }
    
    if (removedCount > 0) {
        SavePreferencesToFile()
        Gui, RemoveRef:Destroy
        RemoveReferences()  ; Refresh the list
    }
}

MoveMsgBoxTop:
    WinGetPos, guiX, guiY, guiW, guiH, Remove References
    if WinExist("Confirm Removal") {
        msgBoxX := guiX + (guiW / 2) - 150
        msgBoxY := guiY + (guiH / 2) - 50
        WinMove, Confirm Removal,, msgBoxX, msgBoxY
        WinSet, AlwaysOnTop, On, Confirm Removal
        ; Also activate the window to ensure focus
        WinActivate, Confirm Removal
    }
return

OpenReference(ItemName) {
    if (g_References.HasKey(ItemName)) {
        thisRef := g_References[ItemName]
        refType := thisRef["type"]
        refPath := thisRef["path"]
        
        if (refType = "url") {
            ; Ensure URL has proper protocol
            if (!RegExMatch(refPath, "^https?://")) {
                refPath := "https://" . refPath
            }
            Run, %refPath%  ; Direct browser launch
            thisRef["uses"] += 1
            SavePreferencesToFile()
        } else {  ; file or mapped
            if (FileExist(refPath)) {
                Run, %refPath%
                thisRef["uses"] += 1
                SavePreferencesToFile()
            } else {
                MsgBox, File not found: %refPath%
            }
        }
    }
}

IsValidURL(url) {
    return RegExMatch(url, "i)^(?:(?:https?|ftp)://)?(?:\w+\.)?\w+\.\w+(?:/|$)")
}

IsValidFileType(ext) {
    static validTypes := {pdf: 1, xls: 1, xlsx: 1, doc: 1, docx: 1}
    StringLower, ext, ext
    return validTypes.HasKey(ext)
}

GetUniqueFilePath(dir, fileName) {
    SplitPath, fileName, , , ext, nameNoExt
    newPath := dir . "\" . fileName
    counter := 1
    
    while (FileExist(newPath)) {
        newPath := dir . "\" . nameNoExt . " (" . counter . ")." . ext
        counter++
    }
    
    return newPath
}

AddRefGuiClose:
    Gui, Destroy
return
MapRefGuiClose:
RemoveRefGuiClose:
    Gui, Destroy
return

RefListHandler:
return

; ------------------------------------------
; AI Prompt
; ------------------------------------------

CreateAIAssistantMenu() {
    Menu, AIAssistantMenu, Add
    Menu, AIAssistantMenu, DeleteAll
    
    ; Add management options at top
    Menu, AIAssistantMenu, Add, Add Prompt..., AddAIPrompt
    Menu, AIAssistantMenu, Add, Manage Prompts..., ManageAIPrompts
    Menu, AIAssistantMenu, Add, AI Settings..., AISettings
    Menu, AIAssistantMenu, Add  ; Separator
    
    ; Create sorted array of prompts
    prompts := []
    for name, data in g_AIPrompts {
        prompts.Push({name: name, uses: data.uses})
    }
    
    ; Sort prompts by usage count (descending)
    SortPromptsByUse(prompts)
    
    ; Add sorted prompts to menu
    count := 0
    for _, prompt in prompts {
        if (count >= g_MaxAIPrompts)
            break
        Menu, AIAssistantMenu, Add, % prompt.name, ExecuteAIPrompt
        count++
    }
}

SortPromptsByUse(ByRef prompts) {
    tempArray := []
    sortedArray := []
    
    for index, prompt in prompts {
        tempArray.Push(prompt)
    }
    
    while (tempArray.Length() > 0) {
        highestUses := -1
        highestIndex := 0
        for index, prompt in tempArray {
            if (prompt.uses > highestUses) {
                highestUses := prompt.uses
                highestIndex := index
            }
        }
        sortedArray.Push(tempArray[highestIndex])
        tempArray.RemoveAt(highestIndex)
    }
    
    prompts := []
    for _, prompt in sortedArray {
        prompts.Push(prompt)
    }
    
    return prompts
}

AddAIPrompt() {
    CoordMode, Mouse, Screen
    MouseGetPos, mouseX, mouseY
    
    ; Get active monitor
    SysGet, monitorCount, MonitorCount
    Loop, %monitorCount% {
        SysGet, monArea, Monitor, %A_Index%
        if (mouseX >= monAreaLeft && mouseX <= monAreaRight && mouseY >= monAreaTop && mouseY <= monAreaBottom) {
            activeMonitor := A_Index
            break
        }
    }
    
    ; Get monitor work area
    SysGet, workArea, MonitorWorkArea, %activeMonitor%
    monitorWidth := workAreaRight - workAreaLeft
    monitorHeight := workAreaBottom - workAreaTop
    
    ; Calculate GUI position
    guiWidth := 400
    guiHeight := 200
    xPos := mouseX + 10
    yPos := mouseY + 10
    
    ; Ensure GUI stays on screen
    if (xPos + guiWidth > workAreaRight)
        xPos := workAreaRight - guiWidth
    if (yPos + guiHeight > workAreaBottom)
        yPos := workAreaBottom - guiHeight
    
    Gui, AddPrompt:New, +AlwaysOnTop
    Gui, AddPrompt:Add, Text,, Prompt Name:
    Gui, AddPrompt:Add, Edit, vvPromptName w380
    Gui, AddPrompt:Add, Text,, Prompt Text:
    Gui, AddPrompt:Add, Edit, vvPromptText w380 h80
    Gui, AddPrompt:Add, Button, gSaveAIPrompt w90, Save
    Gui, AddPrompt:Add, Button, x+10 w90 gAddPromptGuiClose, Cancel
    Gui, AddPrompt:Show, x%xPos% y%yPos% w%guiWidth% h%guiHeight%, Add AI Prompt
}

SaveAIPrompt() {
    global vPromptName, vPromptText, g_AIPrompts
    Gui, AddPrompt:Submit, NoHide
    
    if (vPromptName = "" || vPromptText = "") {
        MsgBox, Please fill in both fields.
        return
    }
    
    g_AIPrompts[vPromptName] := {prompt: vPromptText, uses: 0}
    SavePreferencesToFile()
    CreateAIAssistantMenu()
    Gui, AddPrompt:Destroy
}

ManageAIPrompts() {
    CoordMode, Mouse, Screen
    MouseGetPos, mouseX, mouseY
    
    ; Get active monitor
    SysGet, monitorCount, MonitorCount
    Loop, %monitorCount% {
        SysGet, monArea, Monitor, %A_Index%
        if (mouseX >= monAreaLeft && mouseX <= monAreaRight && mouseY >= monAreaTop && mouseY <= monAreaBottom) {
            activeMonitor := A_Index
            break
        }
    }
    
    ; Get monitor work area
    SysGet, workArea, MonitorWorkArea, %activeMonitor%
    
    ; Calculate GUI position
    guiWidth := 500
    guiHeight := 400
    xPos := mouseX + 10
    yPos := mouseY + 10
    
    ; Ensure GUI stays on screen
    if (xPos + guiWidth > workAreaRight)
        xPos := workAreaRight - guiWidth
    if (yPos + guiHeight > workAreaBottom)
        yPos := workAreaBottom - guiHeight
    
    Gui, ManagePrompts:New, +AlwaysOnTop
    Gui, ManagePrompts:Add, ListView, vPromptList w480 h330 gPromptListHandler -Multi, Prompt Name|Prompt Text|Uses
    
    for name, data in g_AIPrompts {
        LV_Add("", name, data.prompt, data.uses)
    }
    
    ; Add three buttons with equal spacing
    Gui, ManagePrompts:Add, Button, w120 gEditSelectedPrompt, Edit
    Gui, ManagePrompts:Add, Button, x+10 w120 gRemoveSelectedPrompts, Remove Selected
    Gui, ManagePrompts:Add, Button, x+10 w120 gManagePromptsGuiClose, Close
    Gui, ManagePrompts:Show, x%xPos% y%yPos% w%guiWidth% h%guiHeight%, Manage AI Prompts
}

RemoveSelectedPrompts() {
    selectedRows := []
    row := 0
    
    while (row := LV_GetNext(row)) {
        selectedRows.Insert(row)
    }
    
    removedCount := 0
    
    Loop % selectedRows.MaxIndex() {
        row := selectedRows[selectedRows.MaxIndex() - A_Index + 1]
        LV_GetText(name, row, 1)
        
        MsgBox, 4, Confirm Removal, Are you sure you want to remove "%name%"?
        IfMsgBox Yes
        {
            g_AIPrompts.Delete(name)
            removedCount++
        }
    }
    
    if (removedCount > 0) {
        SavePreferencesToFile()
        Gui, ManagePrompts:Destroy
        ManageAIPrompts()
    }
}

AISettings() {
    CoordMode, Mouse, Screen
    MouseGetPos, mouseX, mouseY
    
    ; Get active monitor
    SysGet, monitorCount, MonitorCount
    Loop, %monitorCount% {
        SysGet, monArea, Monitor, %A_Index%
        if (mouseX >= monAreaLeft && mouseX <= monAreaRight && mouseY >= monAreaTop && mouseY <= monAreaBottom) {
            activeMonitor := A_Index
            break
        }
    }
    
    ; Get monitor work area
    SysGet, workArea, MonitorWorkArea, %activeMonitor%
    
    ; Calculate GUI position
    guiWidth := 400
    guiHeight := 200
    xPos := mouseX + 10
    yPos := mouseY + 10
    
    ; Ensure GUI stays on screen
    if (xPos + guiWidth > workAreaRight)
        xPos := workAreaRight - guiWidth
    if (yPos + guiHeight > workAreaBottom)
        yPos := workAreaBottom - guiHeight
    
    ; Create GUI with better spacing and organization
    ;Gui, AISettings:New, +AlwaysOnTop
    
    ; Add a header
    ; Gui, AISettings:Font, s12 bold
    Gui, AISettings:Add, Text, x20 y20 w360 h30, AI Assistant Settings
    
    ; Reset font for rest of controls
    ;Gui, AISettings:Font, s10 normal
    
    ; Add section for AI selection
    ; , AISettings:Add, GroupBox, x20 y60 w360 h70, AI Service Selection
    Gui, AISettings:Add, Text, x35 y85, Select your preferred AI service:
    Gui, AISettings:Add, DropDownList, x35 y105 w330 vg_PreferredAI, claude||chatgpt
    GuiControl, Choose, g_PreferredAI, %g_PreferredAI%
    
    ; Add buttons with proper spacing and grouping
    Gui, AISettings:Add, Button, x20 y150 w110 h30 gRestoreAIDefaults, Restore Prompts
    Gui, AISettings:Add, Button, x240 y150 w70 h30 gSaveAISettings, Save
    Gui, AISettings:Add, Button, x310 y150 w70 h30 gAISettingsGuiClose, Cancel
    
    ; Show the GUI
    Gui, AISettings:Show, x%xPos% y%yPos% w%guiWidth% h%guiHeight%, AI Assistant Settings
}

SaveAISettings() {
    Gui, AISettings:Submit, NoHide
    SavePreferencesToFile()
    Gui, AISettings:Destroy
}

RestoreAIDefaults() {
    global g_AIPrompts, DEFAULT_PROMPTS
    
    MsgBox, 4, Confirm Restore, Warning: This will delete all your current prompts and restore the default prompts. Continue?
    IfMsgBox No
        return
        
    ; Clear existing prompts and copy defaults
    g_AIPrompts := {}
    for name, data in DEFAULT_PROMPTS {
        g_AIPrompts[name] := {prompt: data.prompt, uses: 0}
    }
    
    SavePreferencesToFile()
    CreateAIAssistantMenu()
    
    MsgBox, Default prompts have been restored.
}

ExecuteAIPrompt(ItemName) {
    if (!g_AIPrompts.HasKey(ItemName))
        return
    
    
    selectedText := g_SelectedText
    if (!selectedText) {
        MsgBox, No text selected.
        return
    }
    
    promptData := g_AIPrompts[ItemName]
    promptText := promptData.prompt
	
	; First update the clipboard with the new content
    Clipboard := promptText . "`n`n" . SanitizePHI(selectedText)
	; ShowResult(promptText . "`n`n" . SanitizePHI(selectedText)) ; debug
    ClipWait, 2  ; Wait for the clipboard to be ready

	MsgBox, 4, Warning, IMPORTANT: Do not send protected health information (PHI) to external AI services. Continue?
    IfMsgBox No
        return
    
    if (g_PreferredAI = "claude") {
        Run, https://claude.ai
        Sleep, 4000  ; Wait for Claude to load
        Send, ^v  ; Paste into input box
        Send, {Enter}  ; Send the message
    } else {
        Run, https://chat.openai.com
        Sleep, 4000  ; Wait for ChatGPT to load
        Send, ^v  ; Paste into input box
        Send, {Enter}  ; Send the message
    }
       
    ; Increment usage counter
    promptData.uses += 1
    SavePreferencesToFile()
}

; Add these new functions to your script
EditSelectedPrompt() {
    row := LV_GetNext(0)
    if (!row) {
        MsgBox, Please select a prompt to edit.
        return
    }

    ; Get the current prompt data
    LV_GetText(originalName, row, 1)
    LV_GetText(originalPrompt, row, 2)
    LV_GetText(usageCount, row, 3)

    ; Store original name for later comparison
    g_OriginalPromptName := originalName

    ; Get the current GUI position
    Gui, ManagePrompts:+LastFound
    WinGetPos, currentX, currentY, currentW, currentH

    ; Calculate position for edit window
    editWidth := 400
    editHeight := 200
    editX := currentX + (currentW - editWidth) / 2
    editY := currentY + (currentH - editHeight) / 2

    ; Create edit GUI
    Gui, EditPrompt:New, +AlwaysOnTop
    Gui, EditPrompt:Add, Text,, Prompt Name:
    Gui, EditPrompt:Add, Edit, vvPromptName w380, %originalName%
    Gui, EditPrompt:Add, Text,, Prompt Text:
    Gui, EditPrompt:Add, Edit, vvPromptText w380 h80, %originalPrompt%
    Gui, EditPrompt:Add, Button, gSaveEditedPrompt w90, Save
    Gui, EditPrompt:Add, Button, x+10 w90 gEditPromptGuiClose, Cancel
    Gui, EditPrompt:Show, x%editX% y%editY% w%editWidth% h%editHeight%, Edit AI Prompt
}

SaveEditedPrompt() {
    global vPromptName, vPromptText, g_AIPrompts, g_OriginalPromptName
    Gui, EditPrompt:Submit, NoHide

    if (vPromptName = "" || vPromptText = "") {
        MsgBox, Please fill in both fields.
        return
    }

    ; Get the usage count from the original prompt
    originalUsage := g_AIPrompts[g_OriginalPromptName].uses

    ; If name changed, delete the old prompt
    if (vPromptName != g_OriginalPromptName) {
        g_AIPrompts.Delete(g_OriginalPromptName)
    }

    ; Save the edited prompt with original usage count
    g_AIPrompts[vPromptName] := {prompt: vPromptText, uses: originalUsage}
    
    SavePreferencesToFile()
    Gui, EditPrompt:Destroy
    
    ; Refresh the manage prompts window
    Gui, ManagePrompts:Destroy
    ManageAIPrompts()
}

SanitizePHI(text) {
    ; 1. Names with labels 
	text := RegExReplace(text, "i)(?:name|patient):\s*[a-zA-Z][a-z]+(?:\s+[a-zA-Z]\.?)?\s+[a-zA-Z][a-z\-']+\b", "Name: [NAME]")
	text := RegExReplace(text, "i)(?<![\w'])\b(?:mr|dr|mrs|ms|miss|rev|hon|prof)(?:\.|\b)\s+[a-zA-Z][a-z\-']+(?:\s+[a-zA-Z][a-z\-']+)?\b", "[NAME]")

	; 2. Standalone names at start of lines/sections
	text := RegExReplace(text, "mi)^(?!(?:MRN|ACC|INDICATIONS|TECHNIQUE|COMPARISON|FINDINGS|IMPRESSION|HISTORY|EXAM|CLINICAL)\b)[a-zA-Z][a-z]+(?:\s+[a-zA-Z]\.?)?\s+[a-zA-Z][a-z]+(?:\s*$|\R)", "[NAME]")
	text := RegExReplace(text, "mi)^(?!(?:MRN|ACC|CT|MRI)\b)[a-zA-Z][a-z]+,\s+[a-zA-Z][a-z]+(?:\s*$|\R)", "[NAME]")

	; 3. Names before medical report headers
	text := RegExReplace(text, "mi)^(?!(?:MRN|ACC|CT|MRI)\b)[a-zA-Z][a-z]+(?:\s+[a-zA-Z]\.?)?\s+[a-zA-Z][a-z\-']+(?=\R+(?:MRI|CT|CLINICAL|TECHNIQUE|FINDINGS|IMPRESSION|HISTORY|EXAM)\b)", "[NAME]")

	; 4. Relative names
	text := RegExReplace(text, "i)(?:Father|Mother|Brother|Sister|Son|Daughter|Wife|Husband|Spouse)'s?\s+(?:name|history):\s*[a-zA-Z][a-z]+\s+[a-zA-Z][a-z\-']+\b", "[RELATIVE]")
    
    ; 2. Geographic Subdivisions
    text := RegExReplace(text, "\b(?:AL|AK|AZ|AR|CA|CO|CT|DE|FL|GA|HI|ID|IL|IN|IA|KS|KY|LA|ME|MD|MA|MI|MN|MS|MO|MT|NE|NV|NH|NJ|NM|NY|NC|ND|OH|OK|OR|PA|RI|SC|SD|TN|TX|UT|VT|VA|WA|WV|WI|WY)\b(?=\s+\d{5})", "[STATE]")  ; State abbreviations
    text := RegExReplace(text, "\b\d{5}(?:-\d{4})?\b", "[ZIP]")  ; ZIP codes
    text := RegExReplace(text, "\b[A-Z][a-z\.]+(?: County| Parish| Borough)\b", "[COUNTY]")  ; Counties
    text := RegExReplace(text, "\b\d{1,5}\s+[A-Za-z0-9\s\-\.]+(?:Street|St|Avenue|Ave|Road|Rd|Boulevard|Blvd|Lane|Ln|Drive|Dr|Way|Court|Ct|Circle|Cir|Trail|Trl|Highway|Hwy|Route|Rte|Plaza|Pl|Square|Sq|Terrace|Ter|Parkway|Pkwy|Commons|Point|Pt)\b(?:,?\s+(?:Apt|Suite|Ste|Room|Rm|Unit)\s+[-A-Z0-9]+)?", "[ADDRESS]")  ; Street addresses with units
    text := RegExReplace(text, "\b(?:North|South|East|West|N|S|E|W)\s+[A-Z][a-z\.]+(?: Street| St| Avenue| Ave| Road| Rd)\b", "[ADDRESS]")  ; Directional streets

    ; 3. Dates
    ; text := RegExReplace(text, "(?<!\d\s)(?:\d{1,2}[-/]\d{1,2}[-/]\d{2,4}|\d{4}[-/]\d{1,2}[-/]\d{1,2})", "[DATE]")  ; All date formats
    ; text := RegExReplace(text, "(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)\s+\d{1,2}(?:st|nd|rd|th)?,?\s+\d{4}", "[DATE]")  ; Written dates
    ; text := RegExReplace(text, "i)DOB:?\s*.*?(?=\s|$)", "DOB: [DOB]")  ; Date of birth variations
    
    ; 4. Telephone Numbers
    text := RegExReplace(text, "(?:Tel|Telephone|Phone|Mobile|Cell|Home|Work|Fax)(?:\s*#|\s*:|\s+)\s*(?:\+\d{1,2}\s*)?(?:\(?\d{3}\)?[-\.\s]?)?\d{3}[-\.\s]?\d{4}", "[PHONE]")  ; Phone numbers with labels
    text := RegExReplace(text, "\b(?:\+\d{1,2}\s*)?(?:\(?\d{3}\)?[-\.\s]?)?\d{3}[-\.\s]?\d{4}\b", "[PHONE]")  ; Generic phone numbers
    text := RegExReplace(text, "Extension\s*:?\s*\d+", "[EXT]")  ; Extensions
    
    ; 5. Fax Numbers
    ; text := RegExReplace(text, "(?:Fax|Facsimile)(?:\s*#|\s*:|\s+)\s*(?:\+\d{1,2}\s*)?(?:\(?\d{3}\)?[-\.\s]?)?\d{3}[-\.\s]?\d{4}", "[FAX]")  ; Fax numbers
    
    ; 6. Email Addresses
    text := RegExReplace(text, "\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b", "[EMAIL]")  ; Email addresses
    
    ; 7. Social Security Numbers
    text := RegExReplace(text, "(?:SSN|SIN|Social Security)(?:\s*#|\s*:|\s+)\s*\d{3}[-\s]?\d{2}[-\s]?\d{4}", "[SSN]")  ; SSN with labels
    text := RegExReplace(text, "\b\d{3}[-\s]?\d{2}[-\s]?\d{4}\b(?!\s*(?:mm|cm|tesla))", "[SSN]")  ; Generic SSN pattern (avoiding measurements)
    
    ; 8. Medical Record Numbers and Accession Numbers
    text := RegExReplace(text, "i)(?:MRN|Medical Record Number|Medical Record|Record Number|Chart Number|Patient ID)(?:\s*#|\s*:|\s*Number:|\s+)\s*[A-Za-z0-9\-]+", "[MRN]")  ; Medical record numbers
    text := RegExReplace(text, "i)(?:ACC(?:ESSION)?|Accession Number)(?:\s*#|\s*:|\s*Number:|\s+)\s*[A-Za-z0-9\-]+", "[ACC]")  ; Accession numbers with letters
    text := RegExReplace(text, "i)(?:ACC|Accession):\s*\d+[A-Za-z0-9\-]*", "[ACC]")  ; Additional accession pattern
    
    ; 9. Health Plan Numbers
    text := RegExReplace(text, "i)(?:Health Plan|Insurance|Policy|Group|Member|Coverage)(?:\s*#|\s*:|\s*ID:?|\s+)\s*[A-Za-z0-9\-\._/\\]+", "[INSURANCE]")  ; Insurance IDs
    text := RegExReplace(text, "i)Insurance:?\s*[A-Za-z0-9\-\._/\\]+", "[INSURANCE]")  ; Additional insurance pattern
    
    ; 10. Account Numbers
    text := RegExReplace(text, "i)(?:Account|Acct|Order|Reference|Visit|Encounter)(?:\s*#|\s*:|\s*ID:?|\s+)\s*[A-Za-z0-9\-\._/\\]+", "[ACCOUNT]")  ; Account numbers
    text := RegExReplace(text, "i)(?:Account|Acct):?\s*[A-Za-z0-9\-\._/\\]+", "[ACCOUNT]")  ; Additional account pattern
    
    ; 11. Certificate/License Numbers
    text := RegExReplace(text, "i)(?:Certificate|License|Registration|Certification|Reg|ID)(?:\s*#|\s*:|\s*ID:?|\s+)\s*[A-Za-z0-9\-\._/\\]+", "[LICENSE]")  ; License numbers
    text := RegExReplace(text, "i)(?:License|Cert):?\s*[A-Za-z0-9\-\._/\\]+", "[LICENSE]")  ; Additional license pattern
    
    ; 12. Vehicle Identifiers
    ; text := RegExReplace(text, "i)\b[A-Za-z0-9]{17}\b", "[VIN]")  ; VINs (includes letters)
    ; text := RegExReplace(text, "i)(?:Vehicle|Car|Auto|License|Tag)(?:\s*#|\s*:|\s*ID:?|\s+)\s*[A-Za-z0-9\-\._/\\]+", "[VEHICLE]")  ; Removed "Plate" from this group
	; text := RegExReplace(text, "i)(?:Plate\s*#|Plate\s*:|Plate\s*ID:?|\bPlate\s+)[A-Za-z0-9\-\._/\\]+", "[VEHICLE]")  ; Separate pattern for "Plate" with identifiers
    ; text := RegExReplace(text, "i)(?:VIN|Vehicle):?\s*[A-Za-z0-9\-\._/\\]+", "[VEHICLE]")  ; Additional vehicle pattern
    
    ; 13. Device Identifiers
    text := RegExReplace(text, "i)(?:Device|Serial|Model|Equipment|Unit|Product)(?:\s*#|\s*:|\s*ID:?|\s+)\s*[A-Za-z0-9\-\._/\\]+", "[DEVICE]")  ; Device IDs
    text := RegExReplace(text, "i)(?:Serial|Model):?\s*[A-Za-z0-9\-\._/\\]+", "[DEVICE]")  ; Additional device pattern
    
    ; 14. Web URLs
    ; text := RegExReplace(text, "https?://[^\s<>""]+", "[URL]")  ; URLs
    
    ; 15. IP Addresses
    text := RegExReplace(text, "\b(?:\d{1,3}\.){3}\d{1,3}\b", "[IP]")  ; IPv4 addresses
    text := RegExReplace(text, "\b(?:[A-F0-9]{1,4}:){7}[A-F0-9]{1,4}\b", "[IP]")  ; IPv6 addresses
    
    ; 16. Biometric Identifiers
    ; text := RegExReplace(text, "(?:Fingerprint|Retinal|Iris|Face|Voice)(?:\s*Scan|\s*Image|\s*Print|\s*ID)(?:\s*#|\s*:|\s+)\s*[A-Z0-9\-]+", "[BIOMETRIC]")  ; Biometric data
    
    ; 17. Photographic Images
    ; text := RegExReplace(text, "(?:Photo|Image|Picture|Portrait)(?:\s*#|\s*:|\s+)\s*[A-Z0-9\-]+", "[PHOTO]")  ; Photo references
    
    ; 18. Any other unique identifying number, characteristic, or code
    text := RegExReplace(text, "\b\d{6,}\b(?!\s*(?:mm|cm|tesla))", "[ID]")  ; Long number sequences (avoiding measurements)
    
    ; Age Information (while preserving clinical duration references)
    text := RegExReplace(text, "i)Age:\s*\d+", "Age: [AGE]")  ; Age with label
    text := RegExReplace(text, "\b(?:age[d]?\s*:?\s*|)\d{1,3}\s*(?:years?|yrs?|y\.?o\.?|months?|mos?)\s*(?:old|of?\s*age)\b", "[AGE]")  ; Age variations
    
    ; Provider Names & Facility Information (if not already part of public record)
    text := RegExReplace(text, "(?:Dr\.|Doctor|Provider|Physician)\s+[A-Z][a-z\-']+(?:\s+[A-Z][a-z\-']+)?", "[PROVIDER]")  ; Provider names
    
    return text
}

; Add these labels to your existing GUI labels section
EditPromptGuiClose:
    Gui, EditPrompt:Destroy
return
RestoreAIDefaults:
    RestoreAIDefaults()
return
AddPromptGuiClose:
ManagePromptsGuiClose:
AISettingsGuiClose:
    Gui, Destroy
return

PromptListHandler:
return

; ------------------------------------------
; End of Script
; ------------------------------------------
ExitApp

