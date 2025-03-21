
#Requires Autohotkey v2
#SingleInstance force
;AutoGUI creator: Alguimist autohotkey.com/boards/viewtopic.php?f=64&t=89901
;AHKv2converter creator: github.com/mmikeww/AHK-v2-script-converter
;EasyAutoGUI-AHKv2 github.com/samfisherirl/Easy-Auto-GUI-for-AHK-v2

^1::
{
	if A_LineFile = A_ScriptFullPath && !A_IsCompiled
	{
		myGui := Constructor()
		myGui.Show("w230 h126")
	}

	Constructor()
	{	
		myGui := Gui()
		myGui.Title := "Window"

		myGui.Add("Text", "x16 y8 w40 h20 +0x200", "Start：")
		Start := myGui.Add("Edit", "x56 y8 w80 h20 +Number")

		myGui.Add("Text", "x16 y35 w40 h20 +0x200", "End：")
		End := myGui.Add("Edit", "x56 y35 w80 h20 +Number")

		ButtonOK := myGui.Add("Button", "x40 y80 w80 h23", "&OK")

		;Start.OnEvent("Change", OnEventHandler)

		;End.OnEvent("Change", OnEventHandler)

		ButtonOK.OnEvent("Click", OnEventHandler)

		myGui.OnEvent('Close', (*) => ExitApp())

		OnEventHandler(*)
		{
			myGui.Hide()
			IF WinExist("ahk_exe SysmacStudio.exe")
			{
				WinActivate
				WinWaitActive
				loop End.Value-Start.Value+1
				{
					iNum := Start.Value+A_Index-1
					Send "{Enter}"
					Sleep 500
					Send iNum
					Sleep 500
					Send "{Enter}"  
					Sleep 300
					continue
				}
			}
			else
			{
				MsgBox "SysmacStudio is not open."
			}
		}
		
		return myGui
	}
}


