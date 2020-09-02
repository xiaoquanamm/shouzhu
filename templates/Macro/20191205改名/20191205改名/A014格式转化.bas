Attribute VB_Name = "A014"
Public Function DWG和PDF保存(保存DWG文档位置, 显示内容, 显参, 桌面, 运参)

Dim swapp As Object
Dim part As Object
Dim Filename As String
Set swapp = GetObject(, "SldWorks.Application")
Set part = swapp.ActiveDoc
Set swModel = swapp.ActiveDoc
swapp.Visible = True

If (桌面.Value = 1) Then

保存文档全名 = 保存DWG文档位置.Text & "\" & 运参.List(2)

Else

保存文档全名 = 运参.List(1) & 运参.List(2)
        
End If

Filename1 = 保存文档全名

If (运参.List(3) = "SLDDRW") Then

part.SaveAs2 Filename1 & ".DWG", 0, True, True
part.SaveAs2 Filename1 & ".PDF", 0, True, True

显参.List(0) = "正确"
显参.List(1) = "PDF/DWG保存完成" & "  " & 运参.List(2) & ".DWG"

Else

part.SaveAs2 Filename1 & ".Step", 0, True, True

显参.List(0) = "正确"
显参.List(1) = "Step保存完成" & "  " & 运参.List(2) & ".Step"

End If

End Function

Public Function Part保存(保存DWG文档位置, 显示内容, 显参, 桌面, 运参)

Dim swapp As Object
Dim part As Object
Dim Filename As String
Set swapp = GetObject(, "SldWorks.Application")
Set part = swapp.ActiveDoc
Set swModel = swapp.ActiveDoc
swapp.Visible = True

If (桌面.Value = 1) Then

保存文档全名 = 保存DWG文档位置.Text & "\" & 运参.List(2)

Else

保存文档全名 = 运参.List(1) & 运参.List(2)
        
End If

Filename1 = 保存文档全名

If (运参.List(3) = "SLDDRW") Then

显参.List(0) = "错误"
显参.List(1) = "Part保存错误，请选择零件或装配体"

Else

part.SaveAs2 Filename1 & ".SLDPRT", 0, True, True

显参.List(0) = "正确"

显参.List(1) = "Part输出完成" & "  " & 运参.List(2) & ".SLDPRT"

End If

End Function

