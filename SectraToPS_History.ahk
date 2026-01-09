; ==========================================
; Sectra to PowerScribe History Copier
; Version: 0.1
; Description: Copies clinical history from Sectra side panel
;              and pastes into PowerScribe History field
; ==========================================

#NoEnv
#SingleInstance, Force
SetWorkingDir, %A_ScriptDir%

; -----------------------------------------
; Configuration - ADJUST THESE
; -----------------------------------------
; Window title fragments (partial match)
global SectraWindowTitle := "Sectra"           ; Adjust to match your Sectra window title
global PowerScribeWindowTitle := "PowerScribe" ; Adjust to match your PowerScribe window

; The field label in PowerScribe to target
; Common names: "History", "Clinical History", "Clinical history", "HISTORY"
global PSHistoryFieldLabel := "History"

; -----------------------------------------
; HOTKEY: Ctrl+Shift+H - Copy Sectra History to PowerScribe
; -----------------------------------------
^+h::
    CopySectraHistoryToPS()
return

; -----------------------------------------
; Main Function
; -----------------------------------------
CopySectraHistoryToPS() {
    global SectraWindowTitle, PowerScribeWindowTitle, PSHistoryFieldLabel

    ; Store current clipboard
    ClipSaved := ClipboardAll
    Clipboard := ""

    ; Step 1: Make sure we're in Sectra and copy selected text
    if (!WinActive(SectraWindowTitle)) {
        ; Try to activate Sectra
        WinActivate, %SectraWindowTitle%
        WinWaitActive, %SectraWindowTitle%, , 2
        if (ErrorLevel) {
            MsgBox, 16, Error, Could not find Sectra window.`nLooking for: %SectraWindowTitle%
            Clipboard := ClipSaved
            return
        }
    }

    ; Copy current selection (user should have text selected)
    Send, ^c
    ClipWait, 2
    if (ErrorLevel) {
        MsgBox, 16, Error, No text selected in Sectra or copy failed.`nSelect the history text first, then press the hotkey.
        Clipboard := ClipSaved
        return
    }

    HistoryText := Clipboard

    ; Step 2: Switch to PowerScribe
    WinActivate, %PowerScribeWindowTitle%
    WinWaitActive, %PowerScribeWindowTitle%, , 2
    if (ErrorLevel) {
        MsgBox, 16, Error, Could not find PowerScribe window.`nLooking for: %PowerScribeWindowTitle%
        Clipboard := ClipSaved
        return
    }

    ; Step 3: Find the History field and paste
    ; Method A: Try clicking on the field label (if it's clickable)
    ; Method B: Use Tab navigation
    ; Method C: Use PowerScribe's field navigation

    ; For now, just paste - user positions cursor first
    ; TODO: Add smart field detection

    Sleep, 100
    Send, ^v

    ; Restore clipboard
    Sleep, 100
    Clipboard := ClipSaved

    ; Optional: Show tooltip confirmation
    ToolTip, History pasted!
    SetTimer, RemoveToolTip, -1500
    return
}

RemoveToolTip:
    ToolTip
return

; -----------------------------------------
; HOTKEY: Ctrl+Shift+D - Debug: Show window info
; Use this to find the exact window titles
; -----------------------------------------
^+d::
    WinGetActiveTitle, activeTitle
    WinGetClass, activeClass
    WinGet, activePID, PID, A
    WinGet, activeExe, ProcessName, A

    MsgBox, 0, Window Info,
    (
    Active Window Debug Info:

    Title: %activeTitle%
    Class: %activeClass%
    Process: %activeExe%
    PID: %activePID%

    Use these values to configure the script.
    )
return

; -----------------------------------------
; HOTKEY: Ctrl+Shift+S - Select All in current control and copy
; Useful if history panel needs select-all first
; -----------------------------------------
^+s::
    Send, ^a
    Sleep, 50
    Send, ^c
    ClipWait, 2
    if (!ErrorLevel) {
        ToolTip, Text copied! (%Clipboard% chars)
        SetTimer, RemoveToolTip, -1500
    }
return

; -----------------------------------------
; Tray menu
; -----------------------------------------
Menu, Tray, Tip, Sectra to PowerScribe History
Menu, Tray, Add, Debug Window Info, ShowDebugInfo
Menu, Tray, Add, Reload Script, ReloadScript
Menu, Tray, Add, Exit, ExitScript
return

ShowDebugInfo:
    GoSub, ^+d
return

ReloadScript:
    Reload
return

ExitScript:
    ExitApp
return
