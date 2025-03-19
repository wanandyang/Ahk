#Requires AutoHotkey v2.0

MyGui:=Gui()

MyGui.Add("Text", "w200", "Hello, World!")
MyGui.Add("Button", "w100", "OK")
MyGui.Add("Button", "w100", "Cancel")

MyGui.Show()
