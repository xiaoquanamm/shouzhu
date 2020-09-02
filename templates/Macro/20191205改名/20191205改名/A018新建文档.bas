Attribute VB_Name = "A018"
Function 新建文档(保存DWG文档位置, 显参, 桌面, 运参)

On Error Resume Next

If (桌面 = 1) Then

cc1 = 保存DWG文档位置.Text

cc2 = Right(cc1, Len(cc1) - InStrRev(cc1, "\"))

If (cc2 = "批量工程图") Then

cc3 = 保存DWG文档位置.Text

Else

cc3 = 保存DWG文档位置.Text & "\" & "批量工程图"

End If

MkDir (cc3) '建立项目文件夹

End If

运参.List(5) = cc3
End Function
