#Requires Autohotkey v2
;#SingleInstance force

{
	myGui := Constructor()
	myGui.Show("w180 h100")
}

Constructor()
{
	myGui := Gui()
	myGui.Title := "W"

	myGui.Add("Text", "x16 y8 w40 h20 +0x200", "Start:")
	Start := myGui.Add("Edit", "x56 y8 w80 h20 +Number")

	myGui.Add("Text", "x16 y35 w40 h20 +0x200", "End:")
	End := myGui.Add("Edit", "x56 y35 w80 h20 +Number")

	ButtonOK := myGui.Add("Button", "x40 y70 w80 h23", "&OK")

	;Start.OnEvent("Change", OnEventHandler)

	;End.OnEvent("Change", OnEventHandler)

	ButtonOK.OnEvent("Click", OnEventHandler)

	;myGui.OnEvent('Close', (*) => ExitApp())

	OnEventHandler(*)
	{
		IF WinExist("ahk_exe SysmacStudio.exe")
		{
			WinActivate
			WinWaitActive
			Send "{Tab}"
			loop End.Value-Start.Value+1
			{
				iNum := Start.Value+A_Index-1
				Send "{Enter}"
				Sleep 300
				Send iNum
				Sleep 300
				Send "{Enter}"
				Sleep 100
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
