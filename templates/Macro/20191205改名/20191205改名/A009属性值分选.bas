Attribute VB_Name = "A009"
Function 属性值分选(材料, 处理方式, 零件名字, 输入1, 输入2, 同名, 机台名字, 项目编号, 设计, 审核, 批准, 当前日期, 加工件数量)

Dim part As Object
Dim swapp As Object
Dim longstatus As Long, longwarnings As Long
Set swapp = GetObject(, "SldWorks.Application")
Set swModel = swapp.ActiveDoc
strFilePathName = swModel.GetPathName '文件所在目录+文件名+扩展名  没有点选
strFilePath = Left(strFilePathName, InStrRev(strFilePathName, "\") - 1) & "\" '文件所在目录
strFileName = Mid(strFilePathName, InStrRev(strFilePathName, "\") + 1)
strFileName = Left(strFileName, InStrRev(strFileName, ".") - 1) '文件名
strFileType = UCase(Mid(strFilePathName, InStrRev(strFilePathName, ".") + 1))   '文件扩展名，大写
qq1 = Left(strFileName, InStrRev(strFileName, "-") - 1) '项目编号
qq2 = Right(strFileName, Len(strFileName) - Len(qq1) - 1)
qq3 = Left(qq2, 3) & "00" '模组编号

If (输入2 = "1") Then

材料 = 材料
处理方式 = 处理方式
零件名字 = 零件名字
备注 = 零件名字

Else

WW1 = 输入1 & 材料
Dim strlms As String
strlms = WW1
Select Case strlms
Case Is = "0S45C"
处理方式.Text = "NI"
Case Is = "0SKD11"
处理方式.Text = "HRC58~62°NI"
Case Is = "0A6061"
处理方式.Text = "喷砂，本色阳极"
Case Is = "0SUS304"
处理方式.Text = "--"
Case Is = "0SPCC"
处理方式.Text = "橘纹RAL7035"
Case Is = "0POM"
处理方式.Text = "黑色"
Case Is = "0PC"
处理方式.Text = "透明"
Case Is = "0SUJ2"
处理方式.Text = "HRC58~62°NI"
Case Is = "0ASP23"
处理方式.Text = "HRC64~66"
Case Is = "0优力胶"
处理方式.Text = "--"
Case Is = "0亚克力"
处理方式.Text = "透明"
Case Is = "0--"
处理方式.Text = "--"
'第二种选择
Case Is = "1S45C"
处理方式.Text = "Hard Cr"
Case Is = "1SKD11"
处理方式.Text = "HRC58~62°Hard Cr"
Case Is = "1A6061"
处理方式.Text = "硬质阳极"
Case Is = "1SUS304"
处理方式.Text = "--"
Case Is = "1SPCC"
处理方式.Text = "红色烤漆"
Case Is = "1POM"
处理方式.Text = "黑色"
Case Is = "1PC"
处理方式.Text = "透明"
Case Is = "1SUJ2"
处理方式.Text = "HRC58~62°Hard Cr"
Case Is = "1ASP23"
处理方式.Text = "HRC64~66"
Case Is = "1优力胶"
处理方式.Text = "--"
Case Is = "1亚克力"
处理方式.Text = "透明"
Case Is = "1--"
处理方式.Text = "--"
End Select

If (材料 = "SPCC" Or 材料 = "SUS304") Then
  备注 = "钣金件"
  零件名字 = "安装件"
Else
   If (零件名字 = "圆柱") Then
       备注 = "圆件"
   Else
       备注 = "加工件"
   End If
End If

End If


同名.List(0) = 机台名字
同名.List(1) = qq1
同名.List(2) = strFileName
同名.List(3) = 零件名字
同名.List(4) = 材料
同名.List(5) = 处理方式
同名.List(6) = qq3
同名.List(7) = 备注
同名.List(8) = 设计
同名.List(9) = 审核
同名.List(10) = 批准
同名.List(11) = 当前日期
同名.List(12) = 加工件数量
同名.List(13) = strFileType
同名.List(14) = 输入1
同名.List(15) = 输入2
同名.List(16) = strFilePath

End Function


