Attribute VB_Name = "A011"
Public Function 新建工程图(同名, 运参, 显参, 工程图模板位置)

Dim swapp As Object
Dim part As Object
Dim longstatus As Long, longwarnings As Long
Set swapp = GetObject(, "SldWorks.Application")
Set part = swapp.ActiveDoc
Set swModel = swapp.ActiveDoc
swapp.Visible = True

strFilePathName = swModel.GetPathName '文件所在目录+文件名+扩展名  没有点选
strFilePath = Left(strFilePathName, InStrRev(strFilePathName, "\") - 1) & "\" '文件所在目录
strFileName = Mid(strFilePathName, InStrRev(strFilePathName, "\") + 1)
strFileName = Left(strFileName, InStrRev(strFileName, ".") - 1) '文件名
strFileType = UCase(Mid(strFilePathName, InStrRev(strFilePathName, ".") + 1))   '文件扩展名，大写
WW1 = Left(strFilePathName, Len(strFilePathName) - 6) & "SLDDRW"

运参.List(9) = 工程图模板位置.Text

If (运参.List(9) = "") Then

显参.List(0) = "错误"

显参.List(1) = "未找到工程图模板位置"

显参.List(3) = 1

Exit Function

Else

显参.List(3) = 0

If (Dir(WW1) = "") Then

ww3 = Right(运参.List(9), Len(运参.List(9)) - InStrRev(运参.List(9), "\"))

Set part = swapp.OpenDoc6(运参.List(9), 3, 0, "", longstatus, longwarnings)

boolstatus = part.Create3rdAngleViews(strFilePathName)

longstatus = part.SaveAs3(WW1, 0, 2)

swapp.CloseDoc ww3 & " - V"

运参.List(10) = "新建"

End If

Set part = swapp.OpenDoc6(WW1, 3, 0, "", longstatus, longwarnings)

End If

End Function

