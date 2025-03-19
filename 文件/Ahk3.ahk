#Requires AutoHotkey v2.0
;AutoGUI creator: Alguimist autohotkey.com/boards/viewtopic.php?f=64&t=89901
;AHKv2converter creator: github.com/mmikeww/AHK-v2-script-converter
;EasyAutoGUI-AHKv2 github.com/samfisherirl/Easy-Auto-GUI-for-AHK-v2

if A_LineFile = A_ScriptFullPath && !A_IsCompiled
{
	myGui := Constructor()

    ; 显示 GUI 窗口，宽度为 620，高度为 420
    myGui.Show("w620 h420")
}

Constructor()
{	
	myGui := Gui()
    ; 添加一个编辑框控件，位置为 (32, 32)，宽度为 120，高度为 21，并且只允许输入数字
    Edit1 := myGui.Add("Edit", "x32 y32 w120 h21 +Number")

    ; 为 Edit1 控件添加事件处理程序，当内容改变时触发 OnEventHandler 函数
    Edit1.OnEvent("Change", OnEventHandler)

    ; 为 GUI 窗口添加关闭事件处理程序，当窗口关闭时触发 ExitApp 函数
    myGui.OnEvent('Close', (*) => ExitApp())
	myGui.Title := "自动写节点"
	
	OnEventHandler(*)
	{
        ; 显示工具提示，内容为示例操作和当前 GUI 元素的值
        ToolTip("点击！这是一个示例操作。`n"
        . "当前 GUI 元素的值包括：`n"  
        . "Edit1 => " Edit1.Value "`n", 77, 277)


        ; 设置一个定时器，在 3 秒后隐藏工具提示
        SetTimer(() => ToolTip(), -3000)
	}
	
	return myGui
}