Attribute VB_Name = "A019"
Function 单个工程图Bom(保存, 运参, 桌面, 判断, 同名, 显参)

Call 模型属性读取(判断, 保存)

保存.List(0) = ""

保存.List(1) = 运参.List(1)

保存.List(2) = 运参.List(2) & "." & 运参.List(3)

Call DWG和PDF批量(保存, 运参, 桌面)

For i = 0 To 13

同名.List(i) = 判断.List(i)

Next

If (运参.List(3) = "SLDDRW") Then

Call 生成Bom(运参, 保存, 桌面)
  
Call 写入表格(同名, 保存)

显参.List(0) = "正确"
显参.List(1) = "PDF/DWG以及BOM生成完成"

Else

显参.List(0) = "错误"
显参.List(1) = "PDF/DWG以及BOM生成错误，请选择工程图"

End If

End Function

