Attribute VB_Name = "A006"
Public Function 读取基本信息(显参, 运参, 保存, 保存DWG文档位置, 当前日期, 机台名字)

On Error Resume Next
Dim part As Object
Dim swapp As Object

Set swapp = GetObject(, "SldWorks.Application")

Set swModel = swapp.ActiveDoc

Set swSelMgr = swModel.SelectionManager '激活选择管理器

ks = swSelMgr.GetSelectedObjectCount2(0) '获取被点选零件的数目

If (ks = 0) Then

strFilePathName = swModel.GetPathName '文件所在目录+文件名+扩展名  没有点选

运参.List(0) = "0"

Else

Set swComp = swSelMgr.GetSelectedObjectsComponent3(1, 0) '获取被点选的零件

strFilePathName = swComp.GetPathName '文件所在目录+文件名+扩展名  点选后的

If (strFilePathName = "") Then

strFilePathName = swModel.GetPathName '文件所在目录+文件名+扩展名  点选后没得二次名字

End If

运参.List(0) = "1"

End If

strFilePath = Left(strFilePathName, InStrRev(strFilePathName, "\") - 1) & "\" '文件所在目录
strFileName = Mid(strFilePathName, InStrRev(strFilePathName, "\") + 1)
strFileName = Left(strFileName, InStrRev(strFileName, ".") - 1) '文件名
strFileType = UCase(Mid(strFilePathName, InStrRev(strFilePathName, ".") + 1))   '文件扩展名，大写

mm = Left(strFileName, InStrRev(strFileName, "-") - 1)

运参.List(1) = strFilePath

运参.List(2) = strFileName

运参.List(3) = strFileType

运参.List(4) = mm

运参.List(5) = 保存DWG文档位置

运参.List(6) = 当前日期

运参.List(7) = 机台名字

保存.List(7) = 运参.List(1)

保存.List(8) = 运参.List(3)

显参.List(0) = "正确"

显参.List(1) = "读取完成" & "   " & 运参.List(2) & " ." & 运参.List(3)

显参.List(2) = 运参.List(2)

保存.List(4) = 运参.List(1)

End Function

