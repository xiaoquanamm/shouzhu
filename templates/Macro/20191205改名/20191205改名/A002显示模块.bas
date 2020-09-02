Attribute VB_Name = "A002"
Public Function 显示模块(显参, 显示内容, 项目编号)

显示内容.Text = 显参.List(1)

项目编号.Text = 显参.List(2)

If (显参.List(0) = "正确") Then

显示内容.BackColor = vbGreen

显示内容.ForeColor = vbBlue

Else

显示内容.BackColor = vbYellow

显示内容.ForeColor = vbRed
End If


End Function


