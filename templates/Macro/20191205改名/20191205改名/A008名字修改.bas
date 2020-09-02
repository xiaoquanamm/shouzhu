Attribute VB_Name = "A008"
Function 改名(运参, 显参, 同名, 保存, 判断, 项目编号, 改名选择) '

Dim part As Object
Dim swapp As Object
Set swapp = GetObject(, "SldWorks.Application")
Set part = swapp.ActiveDoc
Set swModel = swapp.ActiveDoc

'手动修改零件

If (显参.List(1) = "未选择需要操作的零件名字") Then
     Exit Function
End If

If (保存.List(0) = "主装配体") Then

  Call 手动修改名字(项目编号, 保存, 显参)

Else

'自动修改零件

If (改名选择.Text = "装配1") = True Then

Call 装配体1改名(同名, 显参, 判断, 运参, 保存)

End If

If (改名选择.Text = "装配2") = True Then

Call 装配体2改名(同名, 显参, 判断, 运参, 保存)

End If

If (改名选择.Text = "加工") = True Then

Call 零件改名字(同名, 显参, 判断, 保存)

End If

If (改名选择.Text = "标件") = True Then

Call 标件名字(项目编号, 保存, 显参, 读取)

End If

If (改名选择.Text = "附件") = True Then

Call 附件改名字(同名, 显参, 保存, 判断)

End If

End If

End Function

Function 手动修改名字(项目编号, 保存, 显参)

Dim part As Object
Dim swapp As Object
Set swapp = GetObject(, "SldWorks.Application")
Set part = swapp.ActiveDoc
Set swModel = swapp.ActiveDoc

   L1 = 项目编号.Text

    mm5 = 保存.List(7) & L1 & "." & 保存.List(8)

       If Dir(mm5) <> "" Then '判断文件夹里是否存在文件

         ee4 = ee4 + 1

         显参.List(0) = "错误"

         显参.List(1) = "改名失败，有同名文件"

         Exit Function
            
        End If

      longstatus = part.Extension.RenameDocument(L1)

      显参.List(0) = "正确"

      显参.List(1) = "手工修改名字完成" & "  " & L1

      显参.List(2) = L1

End Function

Public Function 零件改名字(同名, 显参, 判断, 保存)

Dim part As Object
Dim swapp As Object
Set swapp = GetObject(, "SldWorks.Application")
Set part = swapp.ActiveDoc
Set swModel = swapp.ActiveDoc

wq1 = 判断.ListCount

'单个零件修改
ee4 = 0

If (wq1 = 1) Then

For qq = 0 To 100

ee2 = ee4 + 1

If (ee2 < 10) Then

ee5 = "0" & ee2

Else

ee5 = ee4 + 1

End If

L1 = 保存.List(2) & "-" & 保存.List(3) & ee5

mm5 = 保存.List(4) & L1 & "." & 保存.List(5)

If Dir(mm5) <> "" Then '判断文件夹里是否存在文件

ee4 = ee4 + 1

显参.List(0) = "错误"

显参.List(1) = "改名失败，有同名文件"

Else

Exit For

End If

Next

longstatus = part.Extension.RenameDocument(L1)

显参.List(0) = "正确"

显参.List(1) = "零件修改完成" & "  " & L1

Else

'连续修改零件
For ee1 = 0 To 判断.ListCount - 1

qee1 = 同名.List(ee1)

qee2 = 保存.List(1)

boolstatus = part.Extension.SelectByID2(qee1 & "@" & qee2, "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)

ee2 = ee1 + 1

If (ee2 < 10) Then

ee3 = "0" & ee2

Else

ee3 = ee1 + 1

End If

L1 = 保存.List(2) & "-" & 保存.List(3) & ee3

mm5 = 保存.List(4) & L1 & "." & 保存.List(5)

If Dir(mm5) <> "" Then '判断文件夹里是否存在文件

显参.List(0) = "错误"

显参.List(1) = "改名失败，有同名文件"
                 
Else

longstatus = part.Extension.RenameDocument(L1)

显参.List(0) = "正确"

显参.List(1) = "零件修改完成" & "  " & ee2 & "  个零件"

End If

Next

End If


End Function

Public Function 装配体1改名(同名, 显参, 判断, 运参, 保存)

Dim part As Object
Dim swapp As Object
Set swapp = GetObject(, "SldWorks.Application")
Set part = swapp.ActiveDoc
Set swModel = swapp.ActiveDoc

For ee1 = 0 To 判断.ListCount - 1
mm1 = 运参.List(ee1)
mm1 = UCase(mm1)
If (mm1 = "SLDPRT") Then
显参.List(0) = "错误"
显参.List(1) = "选择中包含零件,不能修改为子装配体1名字"
Exit Function
End If
Next

wq1 = 判断.ListCount

'单个零件修改
ee4 = 0

If (wq1 = 1) Then

For qq = 0 To 100

ee2 = ee4 + 1

If (ee2 < 10) Then

ee5 = "0" & ee2

Else

ee5 = ee4 + 1

End If

L1 = 保存.List(2) & "-" & ee5 & "000"

mm5 = 保存.List(4) & L1 & "." & 保存.List(5)

If Dir(mm5) <> "" Then '判断文件夹里是否存在文件

ee4 = ee4 + 1

显参.List(0) = "错误"

显参.List(1) = "改名失败，有同名文件"

Else

Exit For

End If

Next

longstatus = part.Extension.RenameDocument(L1)

显参.List(0) = "正确"

显参.List(1) = "装配体1改名完成" & "  " & L1

Else

'连续修改零件
For ee1 = 0 To 判断.ListCount - 1

qee1 = 同名.List(ee1)

qee2 = 保存.List(1)

boolstatus = part.Extension.SelectByID2(qee1 & "@" & qee2, "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)

ee2 = ee1 + 1

If (ee2 < 10) Then

ee3 = "0" & ee2

Else

ee3 = ee1 + 1

End If

L1 = 保存.List(2) & "-" & ee3 & "000"

mm5 = 保存.List(4) & L1 & "." & 保存.List(5)

If Dir(mm5) <> "" Then '判断文件夹里是否存在文件

显参.List(0) = "错误"

显参.List(1) = "改名失败，有同名文件"
                 
Else

longstatus = part.Extension.RenameDocument(L1)

显参.List(0) = "正确"

显参.List(1) = "装配体1改名完成" & "  " & ee2 & "  个零件"

End If

Next

End If

End Function


Public Function 装配体2改名(同名, 显参, 判断, 运参, 保存)

Dim part As Object
Dim swapp As Object
Set swapp = GetObject(, "SldWorks.Application")
Set part = swapp.ActiveDoc
Set swModel = swapp.ActiveDoc


MM22 = 判断.ListCount
If (MM22 > 9) Then
显参.List(0) = "错误"
显参.List(1) = "装配体2点选数量不能多余9个"
Exit Function
End If

For ee1 = 0 To 判断.ListCount - 1
mm1 = 运参.List(ee1)
mm1 = UCase(mm1)
If (mm1 = "SLDPRT") Then
显参.List(0) = "错误"
显参.List(1) = "选择中包含零件,不能修改为子装配体2名字"
Exit Function
End If
Next

wq1 = 判断.ListCount

'单个零件修改
ee4 = 0

If (wq1 = 1) Then

For qq = 0 To 10

ee2 = ee4 + 1

L1 = 保存.List(2) & "-" & 保存.List(6) & ee2 & "00"

mm5 = 保存.List(4) & L1 & "." & 保存.List(5)

If Dir(mm5) <> "" Then '判断文件夹里是否存在文件

ee4 = ee4 + 1

显参.List(0) = "错误"

显参.List(1) = "改名失败，有同名文件"

Else

Exit For

End If

Next

longstatus = part.Extension.RenameDocument(L1)

显参.List(0) = "正确"

显参.List(1) = "装配体2改名完成" & "  " & L1

Else

'连续修改零件
For ee1 = 0 To 判断.ListCount - 1

qee1 = 同名.List(ee1)

qee2 = 保存.List(1)

boolstatus = part.Extension.SelectByID2(qee1 & "@" & qee2, "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)

ee2 = ee1 + 1

L1 = 保存.List(2) & "-" & 保存.List(6) & ee2 & "00"

mm5 = 保存.List(4) & L1 & "." & 保存.List(5)

If Dir(mm5) <> "" Then '判断文件夹里是否存在文件

显参.List(0) = "错误"

显参.List(1) = "改名失败，有同名文件"
                 
Else

longstatus = part.Extension.RenameDocument(L1)

显参.List(0) = "正确"

显参.List(1) = "装配体2改名完成" & "  " & ee2 & "  个零件"

End If

Next

End If

End Function

Public Function 附件改名字(同名, 显参, 保存, 判断)

Dim part As Object
Dim swapp As Object
Dim aa001 As String
Set swapp = GetObject(, "SldWorks.Application")
Set part = swapp.ActiveDoc
Set swModel = swapp.ActiveDoc

For ee1 = 0 To 判断.ListCount - 1

mm = Now
mm = Replace(mm, "/", "")
mm = Replace(mm, ":", "")
mm = Replace(mm, " ", "")
mm = Replace(mm, " ", "") & ee1

qee1 = 同名.List(ee1)

qee2 = 保存.List(1)

boolstatus = part.Extension.SelectByID2(qee1 & "@" & qee2, "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)

L1 = 保存.List(1) & "-" & "附件" & mm

longstatus = part.Extension.RenameDocument(L1)

显参.List(0) = "正确"

显参.List(1) = "附件改名完成" & "  " & ee1 + 1 & "  个零件"

Next
End Function



Public Function 标件名字(项目编号, 保存, 显参, 读取)


Dim part As Object
Dim swapp As Object
Set swapp = GetObject(, "SldWorks.Application")
Set part = swapp.ActiveDoc
Set swModel = swapp.ActiveDoc

   L1 = 项目编号.Text
   
    mm5 = 保存.List(4) & L1 & "." & 保存.List(5)

       If Dir(mm5) <> "" Then '判断文件夹里是否存在文件

         显参.List(0) = "错误"

         显参.List(1) = "改名失败，有同名文件"

         Exit Function
            
        End If

      longstatus = part.Extension.RenameDocument(L1)

      显参.List(0) = "正确"

      显参.List(1) = "主装配体手工修改名字完成" & "  " & L1

      显参.List(2) = L1




End Function






