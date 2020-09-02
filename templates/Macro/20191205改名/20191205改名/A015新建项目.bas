Attribute VB_Name = "A015"
Public Function 项目文件建立(保存DWG文档位置, 机台名字, 显参)

On Error Resume Next

CC1 = 保存DWG文档位置.Text & "\"

rr1 = CC1 & 机台名字.Text

MkDir (rr1) '建立项目文件夹

rr2 = CC1 & 机台名字.Text & "\" & "3D"
MkDir (rr2)

rr3 = CC1 & 机台名字.Text & "\" & "3D" & "\" & "3D"
MkDir (rr3)

rr4 = CC1 & 机台名字.Text & "\" & "3D" & "\" & "备件"
MkDir (rr4)

rr5 = CC1 & 机台名字.Text & "\" & "3D" & "\" & "标准件"
MkDir (rr5)

rr6 = CC1 & 机台名字.Text & "\" & "3D" & "\" & "加工件"
MkDir (rr6)

rr7 = CC1 & 机台名字.Text & "\" & "3D" & "\" & "追加发包"
MkDir (rr7)

rr7 = CC1 & 机台名字.Text & "\" & "3D" & "\" & "返修文档"
MkDir (rr7)

rr8 = CC1 & 机台名字.Text & "\" & "报价以及方案"
MkDir (rr8)

rr9 = CC1 & 机台名字.Text & "\" & "参考图面"
MkDir (rr9)

rr10 = CC1 & 机台名字.Text & "\" & "产品图面"
MkDir (rr10)

rr11 = CC1 & 机台名字.Text & "\" & "移交客户资料"
MkDir (rr11)

rr12 = CC1 & 机台名字.Text & "\" & "客户产品资料"
MkDir (rr12)

rr13 = CC1 & 机台名字.Text & "\" & "审核会议记录"
MkDir (rr13)

rr14 = CC1 & 机台名字.Text & "\" & "调机记录表"
MkDir (rr14)

rr15 = CC1 & 机台名字.Text & "\" & "项目计划"
MkDir (rr15)

rr16 = CC1 & 机台名字.Text & "\" & "振动盘资料"
MkDir (rr16)

rr17 = CC1 & 机台名字.Text & "\" & "CCD资料"
MkDir (rr17)



FileCopy "C:\sw2016\TE模板SW2016 版\模板\项目管理\190way包装设备方案及布局.ppt", rr8 & "\" & 机台名字.Text & "设备方案及布局.ppt"

FileCopy "C:\sw2016\TE模板SW2016 版\模板\项目管理\HVA630  手动治具振动盘来料方向说明.ppt", rr16 & "\" & 机台名字.Text & "振动盘来料方向说明.ppt"

FileCopy "C:\sw2016\TE模板SW2016 版\模板\项目管理\设计审查会议记录.xls", rr13 & "\" & 机台名字.Text & "设计审查会议记录.xls"

FileCopy "C:\sw2016\TE模板SW2016 版\模板\项目管理\调机记录表.xls", rr14 & "\" & 机台名字.Text & "调机记录表.xls"

FileCopy "C:\sw2016\TE模板SW2016 版\模板\项目管理\项目计划表.xls", rr15 & "\" & 机台名字.Text & "项目计划表.xls"

显参.List(0) = "正确"

显参.List(1) = "新项目建立完成" & "  " & 机台名字.Text
 
End Function
