Attribute VB_Name = "A010"
Public Function 清除属性()

Dim swModel As Object
Dim swapp As Object
Set swapp = GetObject(, "SldWorks.Application")
Set swModel = swapp.ActiveDoc
'清除原有属性值
 Set swModel = swapp.ActiveDoc
 Set cpm = swModel.Extension.CustomPropertyManager("")
 vCustInfoNameArr2 = swModel.GetCustomInfoNames
   If Not IsEmpty(vCustInfoNameArr2) Then
      For Each vCustInfoName2 In vCustInfoNameArr2
          bRet = swModel.DeleteCustomInfo(vCustInfoName2)
       Next
   End If
  
End Function

Function 配置属性(同名, 显参)

Set swapp = GetObject(, "SldWorks.Application")
Set swModel = swapp.ActiveDoc
'swCustomInfoText = "30"

    swModel.DeleteCustomInfo2 "", "项目名字"
    swModel.DeleteCustomInfo2 "", "项目编号"
    swModel.DeleteCustomInfo2 "", "图号或型号"
    swModel.DeleteCustomInfo2 "", "品名"
    swModel.DeleteCustomInfo2 "", "设计"
    swModel.DeleteCustomInfo2 "", "审核"
    swModel.DeleteCustomInfo2 "", "批准"
    swModel.DeleteCustomInfo2 "", "REV"
    swModel.DeleteCustomInfo2 "", "材料"
    swModel.DeleteCustomInfo2 "", "品牌或技术要求"
    swModel.DeleteCustomInfo2 "", "模组编号"
    swModel.DeleteCustomInfo2 "", "Date"
    swModel.DeleteCustomInfo2 "", "备注"
    swModel.DeleteCustomInfo2 "", "加工件数量"
    '赋值到属性
    swModel.AddCustomInfo3 "", "项目名字", swCustomInfoText, 同名.List(0)
    swModel.AddCustomInfo3 "", "项目编号", swCustomInfoText, 同名.List(1)
    swModel.AddCustomInfo3 "", "图号或型号", swCustomInfoText, 同名.List(2)
    swModel.AddCustomInfo3 "", "品名", swCustomInfoText, 同名.List(3)
    swModel.AddCustomInfo3 "", "材料", swCustomInfoText, 同名.List(4)
    swModel.AddCustomInfo3 "", "品牌或技术要求", swCustomInfoText, 同名.List(5)
    swModel.AddCustomInfo3 "", "模组编号", swCustomInfoText, 同名.List(6)
    swModel.AddCustomInfo3 "", "备注", swCustomInfoText, 同名.List(7)
    swModel.AddCustomInfo3 "", "设计", swCustomInfoText, 同名.List(8)
    swModel.AddCustomInfo3 "", "审核", swCustomInfoText, 同名.List(9)
    swModel.AddCustomInfo3 "", "批准", swCustomInfoText, 同名.List(10)
    swModel.AddCustomInfo3 "", "Date", swCustomInfoText, 同名.List(11)
    swModel.AddCustomInfo3 "", "加工件数量", swCustomInfoText, 同名.List(12)
  oldDate = swModel.CustomInfo2("", "Date1")
  
  
显参.List(0) = "正确"

显参.List(1) = 同名.List(2) & "  数量" & 同名.List(12) & "件  " & 同名.List(4) & "   " & 同名.List(5)
       
显参.List(2) = 同名.List(2)

End Function


