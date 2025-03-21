; This script is for snipping a part of the screen and pasting it into a Word file.
; The hotkey is Ctrl+Q.
^Q::
{
    IF NOT ProcessExist("Snipaste.exe")
    {
        MsgBox "Snipaste is not open."
    }
    else
    {
        Sleep 100
        Send "{F1}"
        Sleep 100
        Send "{Enter}"
        Sleep 100
        IF WinExist("ahk_class OpusApp")
        {
            WinActivate
            WinWaitActive("ahk_class OpusApp")
            Sleep 100
            Send "^v"
            Sleep 100
            if WinExist("ahk_class Qt5QWindowOwnDCIcon")
            {
                WinActivate
            }
            else
            {
                 MsgBox "cMTViewer is not open."
            }
        }
        else
        {
            MsgBox "Word file is not open."
        }
        
    }
}
return