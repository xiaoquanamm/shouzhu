Attribute VB_Name = "A016"
Dim i As Integer
Dim PartsCollect()     As String '遍历清单（阵列）
Dim InCollectCount     As Double '遍历清单长度
Dim CustomInfoQTY      As String
Dim S1, S2                  As Integer
Dim Path_Name          As String
Dim Code_Name_C        As String
Dim Code_              As String
Dim Name_              As String
Dim swapp              As SldWorks.SldWorks
Dim swModelDoc         As SldWorks.ModelDoc2
Dim swConfig           As SldWorks.Configuration
Dim CustPropMgr        As SldWorks.CustomPropertyManager
Public swModel2 As SldWorks.ModelDoc2
Public PARTNAME_Value_temp As String
Public MATERIAL_Value2_temp As String
'============================================
'批量附件属性值写入
'============================================


Public Function 批量附件属性值写入(显参)

Dim swModel As SldWorks.ModelDoc2

Dim swModelDocExt As SldWorks.ModelDocExtension

Dim swSelMgr As SldWorks.SelectionMgr

Dim swBOMAnnotation As SldWorks.BomTableAnnotation

Dim swBOMFeature As SldWorks.BomFeature

Dim swBomTable As Variant

Dim boolstatus As Boolean

Dim BomType As Long

Dim Configuration As String

Dim TemplateName As String

Dim i, j, n, k As Integer

Dim swBOM_name As String

Dim component As Component2

Dim value_temp As Integer

Dim time_start As String

Dim txt_path As String

Set swapp = GetObject(, "SldWorks.Application")

Set swModel = swapp.ActiveDoc

txt_path = swModel.GetPathName() & " .csv"

Open txt_path For Output Shared As #400

Configuration = swModel.GetActiveConfiguration().Name

If swModel.GetType = 1 Then

Set swModel2 = swapp.ActiveDoc

Call Custominfo_change(Configuration)

ElseIf swModel.GetType = 2 Then

value_temp = swModel.ResolveAllLightWeightComponents(False) '轻化取消到还原状态

Dim Components As Variant

Dim SingleComponent As Variant

Dim swComponent As SldWorks.Component2

Components = swModel.GetComponents(False) '获取整个装配体的组成部件（零件或者装配体）

For Each SingleComponent In Components '遍历

Set swComponent = SingleComponent

If Not swComponent Is Nothing Then

If swComponent.GetModelDoc() Is Nothing Then '判断子件对象模型是否存在；轻化状态下获取不到，为空

Debug.Print "没有通过"

Else

Dim X As Integer

Do '此循环实现处理当前模型和子件属性

If Not swComponent Is Nothing And X < 99 Then '一个很原始的方法强制使用当前的模型

Set swModel2 = swModel

X = 100

Else

Set swModel2 = swComponent.GetModelDoc() '取得子件对象模型

X = 101

End If

Call Custominfo_change(swModel2.GetActiveConfiguration().Name)

Loop Until X = 101

End If

Else

Debug.Print " 不能获取到子件"

End If

Next

Else

MsgBox "不是零件或者装配体模型"

End If

swModel.Save '保存文件

Close #400

显参.List(0) = "正确"
显参.List(1) = "批量清除属性完成"

End Function


Private Function Custominfo_change(ByVal vConfigName As String) '处理模型的属性

Dim vConfigNameArr As Variant

Dim vCustInfoNameArr As Variant

Dim vCustInfoName As Variant

Dim vCustInfoName2 As Variant

Dim vCustInfoNameArr2 As Variant

Dim vCustInfoName2_temp As String

Dim vCustInfoName_temp As String

Dim a() As String

Dim b() As String

Dim m, n As Integer

vCustInfoNameArr = swModel2.GetCustomInfoNames2(vConfigName)

vCustInfoNameArr2 = swModel2.GetCustomInfoNames

m = 0

If Not IsEmpty(vCustInfoNameArr2) Then '取得自定义属性表的属性数据

For Each vCustInfoName2 In vCustInfoNameArr2

vCustInfoName2_temp = CStr(vCustInfoName2)

If vCustInfoName2_temp = "" Then '处理属性表中的空白数据行

m = m - 1

ReDim Preserve a(1, m)

Exit For

End If

vCustInfoName_temp_value2 = swModel2.CustomInfo(vCustInfoName2)

ReDim Preserve a(1, m)

a(0, m) = Trim(vCustInfoName2_temp)

a(1, m) = Trim(vCustInfoName_temp_value2)

m = m + 1

ReDim Preserve a(1, m)

Next

End If

n = 0

If Not IsEmpty(vCustInfoNameArr) Then '取得配置特定属性表的属性数据

For Each vCustInfoName In vCustInfoNameArr

vCustInfoName_temp = CStr(vCustInfoName)

If vCustInfoName_temp = "" Then '处理属性表中的空白数据行

n = n - 1
ReDim Preserve b(1, n)

Exit For

End If

vCustInfoName_temp_value = swModel2.CustomInfo2(vConfigName, vCustInfoName)

ReDim Preserve b(1, n)

b(0, n) = Trim(vCustInfoName_temp)

b(1, n) = Trim(vCustInfoName_temp_value)

n = n + 1

ReDim Preserve b(1, n)
Next

End If

Dim s, t As Integer

If m > 0 Then '当数组a中有数据时

For s = 0 To UBound(a, 2) '循环取出a中存储的每一条数据

If a(0, s) <> "" And a(1, s) <> "" Then '当数据有效时

ElseIf a(0, s) = "" Then '当数据无效时（此情况只会是取得最后一条数据时）

Exit For

End If

Next s

If n > 0 Then

For t = 0 To UBound(b, 2) '循环取出b中存储的每一条数据

If b(0, t) <> "" And b(1, t) <> "" Then '当数据有效时

ElseIf b(0, t) = "" Then '当数据无效时（此情况只会是取得最后一条数据时）

Exit For

End If

Next t

End If

End If



Dim DRAWNO_value As String

Dim DRAWNO_value1, DRAWNO_value2, DRAWNO_value13 As String

Dim DRAWNO_value_N As Integer


DRAWNO_value1 = swModel2.GetPathName

DRAWNO_value2 = StrReverse(DRAWNO_value1)

DRAWNO_value_N = InStr(1, DRAWNO_value2, "\")

DRAWNO_value3 = Mid(DRAWNO_value2, 8, DRAWNO_value_N - 8)

DRAWNO_value = StrReverse(DRAWNO_value3)


If Not IsEmpty(vCustInfoNameArr2) Then

For Each vCustInfoName2 In vCustInfoNameArr2 '删除自定义属性

bRet = swModel2.DeleteCustomInfo(vCustInfoName2)


Next
End If

'swModel2.AddCustomInfo3 "", "备注", swCustomInfoText, "附件" '写入图样代号DRAWNO

End Function


