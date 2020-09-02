Attribute VB_Name = "A021"
Public Function 颜色修改(显参, 运参, colure1)

Dim part As Object
Dim swapp As Object
Dim boolstatus As Boolean
Dim longstatus As Long
Set swapp = GetObject(, "SldWorks.Application")
Set part = swapp.ActiveDoc
Set SelMgr = part.SelectionManager
Set swModel = swapp.ActiveDoc
swapp.Visible = True

strFilePathName = swModel.GetPathName '文件所在目录+文件名+扩展名  没有点选
strFilePath = Left(strFilePathName, InStrRev(strFilePathName, "\") - 1) & "\" '文件所在目录
strFileName = Mid(strFilePathName, InStrRev(strFilePathName, "\") + 1)
strFileName = Left(strFileName, InStrRev(strFileName, ".") - 1) '文件名
strFileType = UCase(Mid(strFilePathName, InStrRev(strFilePathName, ".") + 1))   '文件扩展名，大写

运参.List(0) = strFilePath
运参.List(1) = strFileName
运参.List(2) = strFileType

If (运参.List(2) = "SLDDRW") Then

显参.List(0) = "错误"

显参.List(1) = "颜色修改失败，请选择零件"

Exit Function

End If

If (运参.List(2) = "SLDASM") Then

part.OpenCompFile

Set part = swapp.ActiveDoc

swapp.Visible = True

运参.List(3) = part.GetTitle

运参.List(4) = "打开"

End If

If (运参.List(3) <> 运参.List(1)) Then

Dim strd As String
strd = colure1.Text
Select Case strd

Case Is = "1"
 h1 = 255
 h2 = 128
 h3 = 128
Case Is = "2"
 h1 = 192
 h2 = 192
 h3 = 192
 Case Is = "3"
 h1 = 128
 h2 = 255
 h3 = 128
 Case Is = "4"
 h1 = 255
 h2 = 255
 h3 = 0
 Case Is = "5"
 h1 = 128
 h2 = 128
 h3 = 255
 Case Is = "6"
 h1 = 255
 h2 = 128
 h3 = 255
 Case Is = "7"
 h1 = 0
 h2 = 255
 h3 = 255
 Case Is = "8"
 h1 = 0
 h2 = 192
 h3 = 0
Case Is = "9"
 h1 = 250
 h2 = 0
 h3 = 0
End Select

r = Val(h1)
g = Val(h2)
b = Val(h3)
Set swapp = GetObject(, "SldWorks.Application")
Set swComp = swapp.ActiveDoc
vMatProp = swComp.MaterialPropertyValues
vMatProp(0) = r / 255 '红
vMatProp(1) = g / 255 '绿
vMatProp(2) = b / 255 '蓝
swComp.MaterialPropertyValues = vMatProp
swComp.EditRebuild
swComp.GraphicsRedraw2
If (colure1.Text < 9) Then
colure1.Text = colure1.Text + 1
Else
colure1.Text = 1
End If

显参.List(0) = "正确"

显参.List(1) = "颜色修改完成"

Call 保存文件

If (运参.List(4) = "打开") Then

Call 关闭文件(运参)

End If

Else

显参.List(0) = "错误"

显参.List(1) = "颜色修改失败，请选择零件"

End If

End Function

