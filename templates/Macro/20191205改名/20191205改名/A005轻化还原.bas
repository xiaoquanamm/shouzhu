Attribute VB_Name = "A005"
Public Function 轻化还原(显参)

On Error Resume Next
Dim swModel As Object
Set swapp = GetObject(, "SldWorks.Application")
Set swModel = swapp.ActiveDoc

strFilePathName = swModel.GetPathName
strFilePath = Left(strFilePathName, InStrRev(strFilePathName, "\") - 1) & "\" '文件所在目录

strFileName = Mid(strFilePathName, InStrRev(strFilePathName, "\") + 1)

strFileName = Left(strFileName, InStrRev(strFileName, ".") - 1) '文件名

strFileType = UCase(Mid(strFilePathName, InStrRev(strFilePathName, ".") + 1))   '文件扩展名，大写

If (StrComp(strFileType, "SLDASM") = 0) Then

value_temp = swModel.ResolveAllLightWeightComponents(False) '轻化取消到还原状态

显参.List(0) = "正确"
显参.List(1) = strFileName & "  " & "轻化还原完成"
显参.List(2) = strFileName

Else

显参.List(0) = "错误"
显参.List(1) = strFileName & "  " & "轻化还原失败"
显参.List(2) = ""

End If

Set swpart = Nothing
Set swModel = Nothing
Set swapp = Nothing

End Function
