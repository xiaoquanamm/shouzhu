Attribute VB_Name = "A007"
Dim 主装配名字, 项目编号, 装配体编号1, 装配体编号2, strFilePath, strFileType As String
Dim d As String
Dim a As String
Dim swModel As SldWorks.ModelDoc2
Dim swSelMgr As SldWorks.SelectionMgr
Dim swComp As SldWorks.Component2
Dim arr As String
Dim boolstatus As Boolean
Dim longstatus As Long
Dim myMate As Object
Dim i As Long
Dim ks As Long
Dim Str(100) As String
Dim Strm(100) As String

Function 点选数量统计(运参, 同名, 判断, 显参, 保存)

On Error Resume Next
Set swapp = GetObject(, "SldWorks.Application")
Set swModel = swapp.ActiveDoc
Set swSelMgr = swModel.SelectionManager '激活选择管理器
 
主装配名字 = swapp.ActiveDoc.GetTitle() '提取装配体名称

项目编号 = Left(主装配名字, InStrRev(主装配名字, "-") - 1) '项目编号

装配体编号 = Left(Right(主装配名字, Len(主装配名字) - InStrRev(主装配名字, "-")), 3)

装配体编号1 = Left(Right(主装配名字, Len(主装配名字) - InStrRev(主装配名字, "-")), 2)

ks = swSelMgr.GetSelectedObjectCount2(0) '获取被点选零件的数目

If (ks = 0) Then

显参.List(0) = "错误"

显参.List(1) = "未选择需要操作的零件名字"

Exit Function

Else

显参.List(0) = "错误"

显参.List(1) = "程序运行中"

End If

For i = 1 To ks '获取全部被点选零件的名称，循环开始
Set swComp = swSelMgr.GetSelectedObjectsComponent3(i, 0) '获取被点选的零件

arr = swComp.Name2 '提取被点选零件名称

strFilePathName = swComp.GetPathName '文件所在目录+文件名+扩展名
strFilePath = Left(strFilePathName, InStrRev(strFilePathName, "\") - 1) & "\" '文件所在目录
strFileName = Mid(strFilePathName, InStrRev(strFilePathName, "\") + 1)
strFileName = Left(strFileName, InStrRev(strFileName, ".") - 1) '文件名
strFileType = UCase(Mid(strFilePathName, InStrRev(strFilePathName, ".") + 1))   '文件扩展名，大写

If (strFilePathName = "") Then

显参.List(0) = "错误"

显参.List(1) = "选择零件为主装配体"

保存.List(0) = "主装配体"

Exit Function

End If

Str(i - 1) = arr

Strm(i - 1) = strFileType

运参.List(i - 1) = arr

同名.List(i - 1) = Left(arr, InStrRev(arr, "-") - 1)

Next

Call 判断位置相同(判断, 同名)

Call 改名字零件放入(判断, 运参, 同名)

保存.List(0) = ""

保存.List(1) = 主装配名字

保存.List(2) = 项目编号

保存.List(3) = 装配体编号

保存.List(4) = strFilePath

保存.List(5) = strFileType

保存.List(6) = 装配体编号1

运参.Clear

For i = 1 To ks

运参.List(i - 1) = Strm(i - 1)

Next

End Function

Function 判断位置相同(判断, 同名) '分开点选中的相同零件名字

判断.List(0) = "否"

For mm1 = 1 To ks - 1

abc = 同名.List(mm1)

    For mm2 = 0 To mm1
      
      If (mm1 = mm2) Then
      
        ' 判断.List(mm1) = "自身"
      Else
        If abc = 同名.List(mm2) Then
 
          判断.List(mm1) = "是"
             
          GoTo 1
             
        Else
            
          判断.List(mm1) = "否"

        End If
             
      End If
        
     Next
1:

Next

End Function

Function 改名字零件放入(判断, 运参, 同名) '统计需要修改零件的名字

同名.Clear

For WW1 = 0 To 判断.ListCount - 1

qwe1 = 判断.List(WW1)

  If (qwe1 = "否") Then

     同名.List(ww2) = 运参.List(WW1)
     
     ww2 = ww2 + 1

  End If

Next

End Function
