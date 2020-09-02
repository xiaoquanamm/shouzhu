Attribute VB_Name = "PropertyEditorGlobal1"
'---------------------------------------------------------------------03/08/2007
' PropertyEditorGlobal.swp          Written by Leonard Kikstra,
'                                   Copyright 2003-2007, Leonard Kikstra
'                                   Downloaded from Lenny's SolidWorks Resources
'                                        at http://www.lennyworks.com/solidworks
'-------------------------------------------------------------------------------
' Notes:            Source file must be in same directory as macro file.
'                   Source file must have macro name with '.ini' extension.
'------------------------------------------------------------------------------
' Edit/Add/Delete/Rename custom file properties and configuration
' specific properties in one easy to use interface.
' Works on all SolidWorks documents.
' See 'Form_PropertyEditor' for version history
' -----------------------------------------------------------------------------
' Loader module
' -----------------------------------------------------------------------------
Global swApp                    As SldWorks.SldWorks
Global ModelDoc2                As SldWorks.ModelDoc2
Global StepEdit                 As Boolean
Global StepAdded                As Boolean
Global Source                   As String
Global NewDate                  As String
Global ForcePropertyAdd         As Boolean
Global SourceExists             As Boolean
Global ListPosition             As Integer
Global CstmPropType(3)          As Integer                          ' v2.0c
Global PropAdd                  As Boolean
Global WorkDirectory            As String
Global Const Version = "2.00"           ' See individual modules for revisions

Public Enum swDocumentTypes_e
        swDocNONE = 0                   '  Used to be TYPE_NONE
        swDocPART = 1                   '  Used to be TYPE_PART
        swDocASSEMBLY = 2               '  Used to be TYPE_ASSEMBLY
        swDocDRAWING = 3                '  Used to be TYPE_DRAWING
        swDocSDM = 4                    '  Solid data manager.
End Enum

Public Enum swCustomInfoType_e
        swCustomInfoUnknown = 0
        swCustomInfoText = 30           '  VT_LPSTR
        swCustomInfoDate = 64           '  VT_FILETIME
        swCustomInfoNumber = 3          '  VT_I4
        swCustomInfoYesOrNo = 11        '  VT_BOOL
End Enum

Public Enum swOpenDocOptions_e
        swOpenDocOptions_Silent = &H1               '  Open silently or not
        swOpenDocOptions_ReadOnly = &H2             '  Open read only or not
        swOpenDocOptions_ViewOnly = &H4             '  Open view only or not
        swOpenDocOptions_RapidDraft = &H8           '  Convert to RapidDraft
        swOpenDocOptions_LoadModel = &H10           '  Load detached models
        swOpenDocOptions_AutoMissingConfig = &H20   '  Handle missing configs
End Enum

Public Enum swSaveAsOptions_e
        swSaveAsOptions_Silent = &H1                '  Save silently or not
        swSaveAsOptions_Copy = &H2                  '  Save as a copy or not
        swSaveAsOptions_SaveReferenced = &H4        '  Save ref documents
        swSaveAsOptions_AvoidRebuildOnSave = &H8    '  Avoid rebuild on Save
        swSaveAsOptions_UpdateInactiveViews = &H10  '  Update inactive sheets,
        swSaveAsOptions_OverrideSaveEmodel = &H20   '  Override system setting
        swSaveAsOptions_SaveEmodelData = &H40       '  If OverrideSaveEmodel
End Enum

Function CloseAll()
  Set OpenDoc = swApp.ActiveDoc()
  While Not OpenDoc Is Nothing
    swApp.QuitDoc (OpenDoc.GetTitle)
    Set OpenDoc = swApp.ActiveDoc()
  Wend
End Function

Sub ForceProperties()                                               ' v1.37
  ForcePropertyAdd = True                   ' Default Value         ' routine
  Open Source For Input As #1               ' Open source file.
  Do While Not EOF(1)                       ' Loop until end of file.
    Input #1, Reader                        ' Read data line.
    If Reader = "[OPTIONS]" Then            ' Looking for section.
      Do While Not EOF(1)                   ' Loop until end of file.
        Input #1, PropName                  ' Read data line.
        If PropName <> "" Then              ' Is name valid?
          If UCase(Left$(PropName, 16)) = UCase("ForcePropertyAdd") Then
            If UCase(Right$(PropName, 4)) = "TRUE" Then
              ForcePropertyAdd = True
            ElseIf UCase(Right$(PropName, 5)) = "FALSE" Then
              ForcePropertyAdd = False
            End If
            GoTo EndForceRead               ' No more data for section
          End If
        Else
          GoTo EndForceRead                 ' No more data for section
        End If
      Loop
    End If
  Loop
EndForceRead:
  Close #1    ' Close file.
End Sub

Sub ReadProperties()
  Const swDocPART = 1                       ' Consistent with swconst.bas
  Const swDocASSEMBLY = 2
  Const swDocDRAWING = 3
  FileTyp = ModelDoc2.GetType               ' Get doc type
  Open Source For Input As #1               ' Open source file.
  If FileTyp = swDocPART Or FileTyp = swDocASSEMBLY Then
    ' Do the following for all parts and assemblies
    Do While Not EOF(1)                     ' Loop until end of file.
      Input #1, Reader                      ' Read data line.
      If Reader = "[MODEL-CUSTOM]" Then     ' Looking for section.
        Do While Not EOF(1)                 ' Loop until end of file.
          Input #1, PropName                ' Read data line.
          If PropName <> "" Then            ' Is name valid?
            Input #1, PropType              ' Read data line.
            AddProp "", PropName, PropType
          Else
            GoTo EndRead1                   ' No more data for section
          End If
        Loop
EndRead1:
      ElseIf Reader = "[MODEL-CONFIGURATION]" Then ' Looking for section.
        Do While Not EOF(1)                 ' Loop until end of file.
          Input #1, PropName                ' Read data line.
          If PropName <> "" Then            ' Is name valid?
            Input #1, PropType              ' Read data line.
            ConfNames = ModelDoc2.GetConfigurationNames
            For i = 0 To UBound(ConfNames)
              AddProp ConfNames(i), PropName, PropType
            Next i
          Else
            GoTo EndRead2                   ' No more data for section
          End If
        Loop
EndRead2:
      End If
    Loop
  ElseIf FileTyp = swDocDRAWING Then
    ' Do the following for all drawings
    Do While Not EOF(1)                     ' Loop until end of file.
      Input #1, Reader                      ' Read data line.
      If Reader = "[DRAWING-CUSTOM]" Then   ' Looking for section.
        Do While Not EOF(1)                 ' Loop until end of file.
          Input #1, PropName                ' Read data line.
          If PropName <> "" Then            ' Is name valid?
            Input #1, PropType              ' Read data line.
            AddProp "", PropName, PropType
          Else
            GoTo EndRead3                   ' No more data for section
          End If
        Loop
EndRead3:
      End If
    Loop
  End If
  Close #1    ' Close file.
End Sub

Sub AddProp(Conf, Nam, Typ)
  If Typ = CstmPropType(0) Then
    DefaultValue = ""
  ElseIf Typ = CstmPropType(1) Then
    DefaultValue = Format(Now, "mm/dd/yyyy")
  ElseIf Typ = CstmPropType(2) Then
    DefaultValue = 0
  ElseIf Typ = CstmPropType(3) Then
    DefaultValue = "No"
  End If
  ModelDoc2.AddCustomInfo3 Conf, Nam, Typ, DefaultValue
End Sub

Sub Main()
  CstmPropType(0) = 30  ' TEXT      Values match SolidWorks swconst.bas file
  CstmPropType(1) = 64  ' DATE
  CstmPropType(2) = 3   ' NUMBER
  CstmPropType(3) = 11  ' YES/NO
  Set swApp = CreateObject("SldWorks.Application")
  swApp.Visible = True
  swApp.UserControl = True
  Set ModelDoc2 = swApp.ActiveDoc      ' Grab currently active document
  Source = swApp.GetCurrentMacroPathName             ' Get macro path+name
  Source = Left$(Source, Len(Source) - 3) + "ini"    ' Set source file name
  Set FileSys = CreateObject("Scripting.FileSystemObject")
  SourceExists = FileSys.FileExists(Source)
  If ModelDoc2 Is Nothing Then         ' Check to see if a document
    Form_GlobalPropEditor.Show                                      ' V2.0a
  Else
    Response = MsgBox("A document is open in SolidWorks." & Chr(13) _
               & "Should I close documents and proceed?", _
                vbYesNo + vbCritical + vbDefaultButton2, "Selections?")
    If Response = vbYes Then    ' User chose Yes.
      CloseAll                  ' Close all files
      Main
    Else    ' User chose No.
      Exit Sub
    End If
  End If            ' document loaded
End Sub

