#NoEnv ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input ; Recommended for new scripts due to its superior speed and reliability.

#SingleInstance force

#include getActive.ahk

SetCapsLockState, AlwaysOff
SetNumLockState, AlwaysOn
SetScrollLockState, AlwaysOff

GetActiveExplorerPath()
{
    explorerHwnd := WinActive("ahk_class CabinetWClass")
    if (explorerHwnd)
    {
        for window in ComObjCreate("Shell.Application").Windows
        {
            if (window.hwnd==explorerHwnd)
            {
                return window.Document.Folder.Self.Path
            }
        }
    }
    Return "~"
}

GetSelectedText() {
    tmp = %ClipboardAll% ; save clipboard
    Clipboard := "" ; clear clipboard
    Send, ^c ; simulate Ctrl+C (=selection in clipboard)
    ClipWait, 0.2, 0
    selection = %Clipboard% ; save the content of the clipboard
    Clipboard = %tmp% ; restore old content of the clipboard
    return selection
}

#IfWinActive AHK_CLASS CabinetWClass
    #If, GetKeyState("CapsLock", "P")

    f:: ; Create new Text file
        Send {PgUp} ; Force select the first file
        Send ^{Space} ; Clear the selection
        Send {AppsKey} ; Menu key
        Send w ; New
        Send t ; Text Document
    return

    +f:: ; Create new Folder
        Send {PgUp} ; Force select the first file
        Send ^{Space} ; Clear the selection
        Send {AppsKey} ; Menu key
        Send w ; New
        Send f ; Folder
    return

    c:: ; Open folder in code
        path := Explorer_GetPath()
        sel := Explorer_GetSelected()
        if (sel) {
            Run powershell -noProfile -command "code '%sel%'"
        }
        else if(path) {
            Run powershell -noProfile -Command "code '%path%'"
        }
        else {
            Run powershell -noProfile -Command "code"
        }
    return
    s:: ; Open folder in code
        path := Explorer_GetPath()
        sel := Explorer_GetSelected()
        if (sel) {
            Run powershell -noProfile -command "subl '%sel%'"
        }
        else if(path) {
            Run powershell -noProfile -Command "subl '%path%'"
        }
        else {
            Run powershell -noProfile -Command "subl"
        }
    return
    #If
#IfWinActive

#If, GetKeyState("CapsLock", "P")
    t:: ; Open folder in Terminal
    path := Explorer_GetPath()
    if (path == "ERROR")
    {
        path := "~"
    }
    selectedText := GetSelectedText()
    Run, pwsh.exe -NoExit -workingdirectory %path%
    Sleep, 1000
    Send, %selectedText%
return
g::
    selectedText := GetSelectedText()
    url := "https://www.google.com/search?q="
    is_it_an_url := SubStr(selectedText, 1 , 8)
    if (is_it_an_url = "https://") { ; if it starts with "https://" go to, rather than search in google search
        run, %selectedText%
    }
    else { ;search using google search
        joined_url = %url%%selectedText%
        run, %joined_url%
    }
return
y::
    selectedText := GetSelectedText()
    if (selectedText = "") {
        run, "https://www.youtube.com/"
    }else {
        run, "https://www.youtube.com/results?search_query=%selectedText%"
    }
return
Left::
    Send {CtrlDown}+{Tab}{CtrlUp}
return
Right::
    Send {CtrlDown}{Tab}{CtrlUp}
return
#If

; replace windows + e with open this pc
#e::
    Run, explorer.exe shell:MyComputerFolder