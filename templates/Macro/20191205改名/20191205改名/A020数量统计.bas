Attribute VB_Name = "A020"
Dim TopDocPathOnly     As String
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

Public Function 数量统计(显参)

On Error Resume Next
Set swapp = GetObject(, "SldWorks.Application") 'SW对象
Set TopDoc = swapp.ActiveDoc '总装对象
If TopDoc.GetType <> 2 Then
显参.List(1) = "零件统计失败"
显参.List(0) = "错误"
Exit Function
End If

TopDocPathSplit = Split(TopDoc.GetPathName, "\") '分割
TopDocName = TopDocPathSplit(UBound(TopDocPathSplit)) '总装文件名称
TopDocName = Left(TopDocName, Len(TopDocName) - 7) '总装文件名称（排除.SLDASM）
TopDocPathOnly = Mid(TopDoc.GetPathName, 1, InStrRev(TopDoc.GetPathName, "\", -1)) '总装的完整目录
TopConfString = TopDoc.GetActiveConfiguration.Name '總裝配置名稱
CustomInfoQTY = "零件个数" '可按个人喜好修改预设值
InCollectCount = 1 '遍历清单长度基数
ReDim PartsCollect(InCollectCount) '定义阵列项数
SubAsm TopDoc, TopConfString '遍历
Set swApo = GetObject(, "SldWorks.Application") '重建模型
Set part = swApo.ActiveDoc
 part.EditRebuild
 Set swModel = swApo.ActiveDoc '保存当前文件
swModel.Save
 Set swApo = GetObject(, "SldWorks.Application")
 
显参.List(1) = "零件统计完成"
显参.List(0) = "正确"
   
Set swpart = Nothing
Set swModel = Nothing
Set swapp = Nothing

End Function

 Function SubAsm(AsmDoc, ConfString) '统计装配体各零件数量
 On Error Resume Next
 
Set Configuration = AsmDoc.GetConfigurationByName(ConfString)
 Set RootComponent = Configuration.GetRootComponent
 Components = RootComponent.GetChildren
 For Each Child In Components
 Set ChildModel = Child.GetModelDoc
Dim swModel As Object
Set swapp = GetObject(, "SldWorks.Application")
Set swModel = swapp.ActiveDoc
If Not (ChildModel Is Nothing) Then '排除抑制及轻化
ChildConfString = Child.ReferencedConfiguration '零件配置名称
ChildType = ChildModel.GetType
 ChildPathSplit = Split(Child.GetPathName, "\") '分割
ChildName = ChildPathSplit(UBound(ChildPathSplit)) '零件文件名称
ChildPathOnly = Mid(Child.GetPathName, 1, InStrRev(Child.GetPathName, "\", -1)) '零件的完整目录
If ChildPathOnly = Replace(ChildPathOnly, TopDocPathOnly, "") Then SamePath = False Else SamePath = True '零件是否在总装目录或往下目录
If SamePath And (Not Child.ExcludeFromBOM) And (Not Child.IsEnvelope) Then '跳過：不在总装目錄或其往下目錄 或 不包括在材料明細表中 或 是个封套
If (Not Child.ExcludeFromBOM) And (Not Child.IsEnvelope) Then '跳过：不包括在材料明細表中 及 封套
UNIT_OF_MEASURE = ChildModel.CustomInfo2(ChildConfString, UNIT_OF_MEASURE_Name) '备用量
If (UNIT_OF_MEASURE = 0) Or (UNIT_OF_MEASURE = "") Then UNIT_OF_MEASURE = 1 '备用量除错
inCollect = False '重置判断变量
For Each PartinCollect In PartsCollect '判断是否已在遍历清单內
If ChildConfString & "@" & ChildName = PartinCollect Then inCollect = True
 Next
 If inCollect Then '已在遍历清单內
ht_Qty = ChildModel.CustomInfo2("", CustomInfoQTY) + 1 * UNIT_OF_MEASURE
 ChildModel.DeleteCustomInfo2 "", CustomInfoQTY
 ChildModel.AddCustomInfo3 "", CustomInfoQTY, 30, ht_Qty
 Else '不在遍历清单內（首次处理）
ChildModel.DeleteCustomInfo2 "", CustomInfoQTY
 ChildModel.AddCustomInfo3 "", CustomInfoQTY, 30, UNIT_OF_MEASURE
InCollectCount = InCollectCount + 1 '遍历清单长度基数+1
 ReDim Preserve PartsCollect(InCollectCount) '重新定义阵列项数（保留內含数据）
PartsCollect(InCollectCount - 1) = ChildConfString & "@" & ChildName '加入到遍历清单中
End If
 If ChildType = 2 Then
 SubAsm ChildModel, ChildConfString '如果是装配则向下遍历
End If
 End If
 End If
 End If
Next
 End Function
