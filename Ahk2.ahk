#Requires AutoHotkey v2.0

; This script is for snipping a part of the screen and pasting it into a Word file.
; The hotkey is Ctrl+Q.
^Q::
{
/*
    IF WinExist("ahk_class Qt5QWindowToolSaveBits")
    {
        WinActivate
    }
    else
    {
        MsgBox "Snipaste is not open."
    }
*/
    ;Sleep 100
    Send "{F1}"
    Sleep 500
    Send "{Enter}"
    Sleep 500
    IF WinExist("ahk_class OpusApp")
    {
        WinActivate
    }
    else
    {
        MsgBox "Word file is not open."
    }
    Sleep 500
    Send "^v"
}
return