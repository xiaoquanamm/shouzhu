Attribute VB_Name = "A017"
Public Function 生成Bom(运参, 保存, 桌面)

Dim xlApp As Excel.Application
Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Set xlApp = CreateObject("Excel.Application") '创建EXCEL应用类
xlApp.Visible = True
Set xlBook = xlApp.Workbooks.Open("C:\sw2016\TE模板SW2016 版\模板\bom\加工件 BOM表.XLS") '打开标准件
Set xlSheet = xlBook.Worksheets("加工件")  '操作一个SHEET
xlSheet.Cells(6, 5) = " 设备名称：" & 运参.List(7)
xlSheet.Cells(6, 1) = "申请日期：" & 运参.List(6) & "    项目负责人：" & "刘茂松"
    
xlSheet.Cells(5, 1) = "加工件  申请BOM（表单编号：" & 运参.List(4) & "）"
    
Set xlBook = xlApp.Workbooks("加工件 BOM表.XLS")  '打开标准件

ww1 = Right(运参.List(6), Len(运参.List(6)) - InStrRev(运参.List(6), "/"))

ww2 = Left(运参.List(6), InStrRev(运参.List(6), "/") - 1)

ww3 = Right(ww2, Len(ww2) - InStrRev(ww2, "/"))

ww4 = Left(运参.List(6), 4)

If (ww1 < 10) Then

ww1 = "0" & ww1

End If

If (ww3 < 10) Then

ww3 = "0" & ww3

End If

aa = ww4 & ww3 & ww1
 
If (桌面 = 1) Then

BOM生成文件位置 = 运参.List(5) & "\" & aa & " " & 运参.List(7) & ".XLS"

Else

BOM生成文件位置 = 运参.List(1) & aa & " " & 运参.List(7) & ".XLS"

End If

xlBook.SaveCopyAs BOM生成文件位置

Set xlBook = xlApp.Workbooks("加工件 BOM表.XLS")

 保存.List(2) = BOM生成文件位置

xlBook.Close False
Set xlBook = Nothing
xlApp.Quit
Set xlApp = Nothing

End Function

Public Function 写入表格(同名, 保存)

Dim xlApp As Excel.Application
Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet

Set xlApp = CreateObject("Excel.Application") '创建EXCEL应用类
xlApp.Visible = True
Set xlBook = xlApp.Workbooks.Open(保存.List(2)) '打开标准件
 Set xlSheet = xlBook.Worksheets("加工件")  '操作一个SHEET
       
For i3 = 0 To 同名.ListCount / 13
       
xlSheet.Cells(i3 + 8, 1) = i3 + 1
 
xlSheet.Cells(i3 + 8, 2) = 同名.List(i3 * 13 + 3)

xlSheet.Cells(i3 + 8, 3) = 同名.List(i3 * 13 + 2)

xlSheet.Cells(i3 + 8, 4) = 同名.List(i3 * 13 + 12)

xlSheet.Cells(i3 + 8, 5) = 同名.List(i3 * 13 + 4)

xlSheet.Cells(i3 + 8, 6) = 同名.List(i3 * 13 + 5)

xlSheet.Cells(i3 + 8, 7) = 同名.List(i3 * 13 + 12)

xlSheet.Cells(i3 + 8, 9) = 同名.List(i3 * 13 + 6)

xlSheet.Cells(i3 + 8, 11) = 同名.List(i3 * 13 + 1)

xlSheet.Cells(i3 + 8, 16) = 同名.List(i3 * 13 + 7)

Next

  xlBook.Close (True)

Set xlBook = Nothing
xlApp.Quit
Set xlApp = Nothing

End Function


Public Function 工程图查找(运参, 保存, 判断, 同名, File1, 桌面, 显参)

File1.Path = Left(运参.List(1), Len(运参.List(1)) - 1)

File1.Pattern = "*.SLDDRW" '匹配 txt 文件

For i2 = 0 To File1.ListCount - 1

保存.List(0) = i2

保存.List(1) = File1.Path

保存.List(2) = File1.List(i2)

Dim boolstatus As Boolean
Dim longstatus As Long, longwarnings As Long
Set swapp = GetObject(, "SldWorks.Application")
Set part = swapp.ActiveDoc
Set swModel2 = swapp.ActiveDoc
Set swModel = swapp.ActiveDoc
Set part = swapp.OpenDoc6(保存.List(1) & "\" & 保存.List(2), 3, 0, "", longstatus, longwarnings)

Call 模型属性读取(判断, 保存)

mm = i2 * 13

For i1 = 0 To 13

同名.List(i1 + mm) = 判断.List(i1)

Next

Call DWG和PDF批量(保存, 运参, 桌面)

Call 关闭文件(运参)

Next

显参.List(0) = "正确"
显参.List(1) = "PDF/DWG以及BOM生成完成，共计" & i2 & "个工程图"

End Function

Public Function 模型属性读取(判断, 保存)

Dim vCustInfoName2_temp As String
Dim vCustInfoName_temp As String
Dim a() As String
Dim part As Object
Dim swapp As Object
Dim m As Integer

Dim boolstatus As Boolean
Dim longstatus As Long, longwarnings As Long
Set swapp = GetObject(, "SldWorks.Application")
Set part = swapp.ActiveDoc
Set swModel2 = swapp.ActiveDoc
Set swModel = swapp.ActiveDoc
Set part = swapp.OpenDoc6(保存.List(1) & "\" & 保存.List(2), 3, 0, "", longstatus, longwarnings)


vCustInfoNameArr = swModel2.GetCustomInfoNames2(vConfigName)

vCustInfoNameArr2 = swModel2.GetCustomInfoNames

If Not IsEmpty(vCustInfoNameArr2) Then '取得自定义属性表的属性数据

For Each vCustInfoName2 In vCustInfoNameArr2

vCustInfoName2_temp = CStr(vCustInfoName2)

vCustInfoName_temp_value2 = swModel2.CustomInfo(vCustInfoName2)

ReDim Preserve a(1, m)

a(0, m) = Trim(vCustInfoName2_temp)

a(1, m) = Trim(vCustInfoName_temp_value2)

m = m + 1

'属性名称.Text = vCustInfoName2_temp

'属性值.Text = vCustInfoName_temp_value2

判断.List(m - 1) = vCustInfoName_temp_value2

ReDim Preserve a(1, m)
Next
End If

End Function

Public Function DWG和PDF批量(保存, 运参, 桌面)

Dim swapp As Object
Dim part As Object
Dim Filename As String
Set swapp = GetObject(, "SldWorks.Application")
Set part = swapp.ActiveDoc
Set swModel = swapp.ActiveDoc
swapp.Visible = True

If (桌面 = 1) Then

BOM生成文件位置 = 运参.List(5) & "\"
Else

BOM生成文件位置 = 运参.List(1)

End If

mm1 = Left(保存.List(2), InStrRev(保存.List(2), ".") - 1)

mm2 = UCase(Right(保存.List(2), Len(保存.List(2)) - InStrRev(保存.List(2), ".")))

Filename1 = BOM生成文件位置 & mm1

If (mm2 = "SLDDRW") Then

part.SaveAs2 Filename1 & ".DWG", 0, True, True
part.SaveAs2 Filename1 & ".PDF", 0, True, True

End If

End Function


