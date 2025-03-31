;#Requires AutoHotkey v2.0+
;#SingleInstance force
; 检查是否存在运行中的 Excel 应用程序
if WinExist("ahk_exe Excel.exe")
{
    ; 获取活动的 Excel 应用程序对象
    objExcel := ComObjActive("Excel.Application")
    
    ; 设置 Excel 应用程序为可见
    objExcel.Visible := true

    ; 获取活动的工作簿对象
    objWorkbook := objExcel.ActiveWorkbook
}
else
{
    ; 创建新的 Excel 应用程序对象
    objExcel := ComObject("Excel.Application")

    ; 设置 Excel 应用程序为可见
    objExcel.Visible := true

    ; 添加一个新的工作簿
    objWorkbook := objExcel.Workbooks.Add()
}

; 获取第一个工作表
objSheet := objWorkbook.Sheets(1)
; 设置工作表名称
if objSheet.Name != "示例工作表"
{
    objSheet.Name := "示例工作表"
}

; 如果工作簿中的工作表数量少于 2，则添加第二个工作表
if objWorkbook.Sheets.Count < 2
{
   objSheet2 := objWorkbook.Sheets.Add()
   objSheet2.Name := "第二个工作表"
}
else
{
    ; 添加第二个工作表
   objSheet2 := objWorkbook.Sheets(2)
   if objSheet2.Name != "第二个工作表"
   {
      objSheet2.Name := "第二个工作表"
   }
}

; 设置单元格值
objSheet.Cells(1, 1).Value := "Hello, Excel!"
objSheet.Cells(1, 2).Value := "This is a test!"
objSheet.Cells(2, 1).Value := "第2行，第1列"
objSheet.Cells(2, 2).Value := "第2行，第2列"
objSheet.Cells(3, 1).Value := "第3行，第1列"
objSheet.Cells(3, 2).Value := "第3行，第2列"

; 设置单元格背景颜色
objSheet.Cells(1, 1).Interior.Color := 0xFF0000 ; 
; 设置单元格字体颜色
objSheet.Cells(1, 2).Font.Color := 0xFFF000 ; 
; 设置单元格字体加粗
objSheet.Cells(2, 1).Font.Bold := true
; 设置单元格字体大小
objSheet.Cells(2, 2).Font.Size := 14
; 设置单元格对齐方式
objSheet.Cells(3, 1).HorizontalAlignment := -4108 ; xlCenter
objSheet.Cells(3, 1).VerticalAlignment := -4108 ; xlCenter
; 设置单元格字体名称
objSheet.Cells(3, 1).Font.Name := "Arial" ; 设置字体为 Arial
; 设置单元格字体斜体
objSheet.Cells(3, 1).Font.Italic := true ; 设置字体为斜体
; 设置单元格字体下划线
objSheet.Cells(3, 1).Font.Underline := true ; 设置字体为下划线
; 设置单元格字体删除线
objSheet.Cells(3, 1).Font.Strikethrough := true ; 设置字体为删除线

; 设置单元格字体大小
objSheet.Cells(3, 1).Font.Size := 12 ; 设置字体大小为 12

; 设置单元格字体颜色

objSheet.Cells(3, 1).Font.Color := 0x0000FF ; 设置字体颜色为蓝色

; 设置单元格背景颜色
objSheet.Cells(3, 1).Interior.Color := 0xFFFF00 ; 设置背景颜色为黄色

; 设置单元格背景颜色透明度
objSheet.Cells(3, 1).Interior.Pattern := 1 ; 
; 设置背景颜色透明度为绿色
objSheet.Cells(3, 1).Interior.PatternColor := 0x00FF00 

; 设置单元格边框颜色
objSheet.Cells(3, 1).Borders.Color := 0x00FF00 ; 设置边框颜色为绿色

; 设置单元格边框线条粗细
objSheet.Cells(3, 1).Borders.Weight := 2 ; 设置边框线条粗细为 2

; 设置单元格边框线条样式
objSheet.Cells(3, 1).Borders.LineStyle := 1 ; 设置边框线条样式为实线




; 设置中文字符字体为微软雅黑
objSheet.Cells.Font.Name := "微软雅黑" 
; 设置英文字符字体为 Times New Roman
objSheet.Cells.Font.Name := "Times New Roman" 





; 设置单元格边框
objSheet.Cells(3, 1).Borders.LineStyle := 1 ; xlContinuous
objSheet.Cells(3, 1).Borders.Weight := 2 ; xlThin
objSheet.Cells(3, 1).Borders.Color := 0x000000 ; 设置边框颜色为黑色
; 设置单元格格式为日期
objSheet.Cells(4, 1).NumberFormat := "yyyy-mm-dd" ; 设置日期格式
objSheet.Cells(4, 1).Value := "2023-10-01" ; 设置日期值为 2023-10-01
; 保存工作簿
;objWorkbook.SaveAs("D:\Example.xlsx")

; 关闭 Excel
;objExcel.Quit()