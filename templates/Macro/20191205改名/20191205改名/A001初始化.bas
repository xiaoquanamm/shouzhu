Attribute VB_Name = "A001"
Function 初始化(运参, 显参, 同名, 判断, 显示内容, 保存)

显示内容.Text = ""

显示内容.Text = "程序运行中"

显示内容.BackColor = vbWhite

显示内容.ForeColor = vbRed

For i = 0 To 100

运参.List(i) = ""

显参.List(i) = ""

同名.List(i) = ""

Next

同名.Clear

保存.Clear

判断.Clear

End Function



