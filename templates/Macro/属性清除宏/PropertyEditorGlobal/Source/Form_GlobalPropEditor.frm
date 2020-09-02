VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_GlobalPropEditor 
   Caption         =   "PropertyEditorGlobal: RULE BUILDER"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9555
   OleObjectBlob   =   "Form_GlobalPropEditor.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Form_GlobalPropEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------03/08/2007
' PropertyEditorGlobal.swp          Written by Leonard Kikstra,
'                                   Copyright 2003-2007, Leonard Kikstra
'                                   Downloaded from Lenny's SolidWorks Resources
'                                        at http://www.lennyworks.com/solidworks
'-------------------------------------------------------------------------------
' Rules Based Global Property Editor Module
' -----------------------------------------------------------------------------
' Edit/Add/Delete/Rename custom file properties and configuration specific
' properties in one easy to use interface. Works on all SolidWorks documents
' -----------------------------------------------------------------------------
' Columns in 'ListBoxContent' are used as follows:
'   0   Verbose text displayed to user
'   1   Action Code ------------------------------> 1 = Add     3 = Delete
'                                                   2 = Revise  4 = Rename
'   2   Property Name
'   3   Property Type ----------------------------> 1 = Test    3 = Number
'                                                   2 = Date    4 = Yes/No
'   4   Property Value      (New property value)
'       New Property Name   (Renaming property)
'   5   Document Type -----------> 0 = Part    1 = Assembly    2 = Drawing
'   6   Property style ----------> 0 = Custom Property    1 = Config Specific
'   7   Overwrite existing values
' ------------------------------------------------------------------------------
' Version   2.0a  * Added "Global Properties" capabilities
'           2.0b  * Added "Global Properties" rules import capabilities
'                 * Enhanced "Global Properties" capabilities
'           2.0c  * Added "Global Properties" rules export capabilities
' ------------------------------------------------------------------------------
Dim TempNam As String
Dim TempVal As String
Dim swConfigMgr             As SldWorks.ConfigurationManager
Dim vConfName               As Variant
Dim vConfParam              As Variant
Dim vConfValue              As Variant
Dim vPropName               As Variant
Dim PropNames               As Variant
Dim nNumProp                As Integer
Dim RestrictInput           As Boolean
Dim Title                   As String
Dim ProcessParts            As Boolean
Dim ProcessAssemblies       As Boolean
Dim ProcessDrawings         As Boolean
Dim FileError               As Long
Dim FileWarn                As Long
Dim OpenRO                  As Boolean
Dim FileRO                  As Boolean
Const swSumInfoTitle = 0
Const swSumInfoSubject = 1
Const swSumInfoAuthor = 2
Const swSumInfoKeywords = 3
Const swSumInfoComment = 4

'------------------------------------------------------------------------------
' This function will count the occurance of "stringToFind" inside "theString".
'------------------------------------------------------------------------------
Function ParseStringCount(theString As String, stringToFind As String)
  Dim CountPosition As Long
  CountPosition = 1
  CountNum = 0
  While CountPosition > 0
    CountPosition = InStr(1, theString, stringToFind)
    If CountPosition <> 0 Then    ' Which characters do we keep
      theString = Mid(theString, CountPosition + Len(stringToFind))
      CountNum = CountNum + 1
    End If
  Wend
  ParseStringCount = LTrim(CountNum)
End Function

'------------------------------------------------------------------------------
' This function will locate the occurance of "stringToFind" inside "theString".
' From that position, it will return the remaining characters
' based on the 'FirstLast' variable passed into this routine.
'   FirstLast = 0       Get first occurrance and return text before
'               1       Get last occurrance and return text after
'------------------------------------------------------------------------------
Function reverseParseString(theString As String, stringToFind As String, _
                            FirstLast As Boolean)
    Dim slashPosition As Long
    slashPosition = 1
    While slashPosition > 0
      slashPosition = InStr(1, theString, stringToFind)
      If slashPosition <> 0 Then    ' Which characters do we keep
        If FirstLast = True Then    '   0 = Return text before
                                    '   1 = Return text after
          theString = Mid(theString, slashPosition + Len(stringToFind))
        Else
          theString = Mid(theString, 1, slashPosition - 1)
        End If
         
      End If
    Wend
    reverseParseString = theString
End Function

'------------------------------------------------------------------------------
' This function will locate the occurance of "stringToFind" inside "theString".
'   Occur  = 0       Get first occurrance
'            1       Get last occurrance
' From that position, it will return text based on the 'Retrieve' variable.
'   StrRet = 0       Get text before
'            1       Get text after
'            2       If no occurrance found, return blank
'------------------------------------------------------------------------------
Function NewParseString(theString As String, stringToFind As String, _
                            Occur As Boolean, StrRet As Integer)
  Location = 0
  TempLoc = 0
  PosCount = 1
  While PosCount < Len(theString)
    TempLoc = InStr(PosCount, theString, stringToFind)
    If TempLoc <> 0 Then
      If Occur = 0 Then
        If Location = 0 Then
          Location = TempLoc
        End If
      Else
        Location = TempLoc
      End If
      PosCount = TempLoc + 1
    Else
      PosCount = Len(theString)
    End If
  Wend
  If Location <> 0 Then
    If StrRet = 0 Then
      theString = Mid(theString, 1, Location - 1)
    Else
      theString = Mid(theString, Location + Len(stringToFind))
    End If
  ElseIf StrRet = 2 Then
    theString = ""
  End If
  NewParseString = theString
End Function

' ---------------------------------------------------------------------------------
' Read attributes for file
' ---------------------------------------------------------------------------------
Sub FileAttrib(filespec, Bit, SetClear)
  Set FileSys = CreateObject("Scripting.FileSystemObject")
  Set f = FileSys.GetFile(filespec)
  If SetClear = 1 Then
    f.Attributes = f.Attributes + Bit
  Else
    f.Attributes = f.Attributes - Bit
  End If
  ' 0  Normal       4  System
  ' 1  ReadOnly     32 Archive
  ' 2  Hidden
End Sub

' ---------------------------------------------------------------------------------
' Change attributes for file
' ---------------------------------------------------------------------------------
Sub FileAttribRead(filespec, Bit)
  Set FileSys = CreateObject("Scripting.FileSystemObject")
  Set f = FileSys.GetFile(filespec)
  If f.Attributes And Bit Then
    FileRO = True
  Else
    FileRO = False
  End If
  ' 0  Normal       4  System
  ' 1  ReadOnly     32 Archive
  ' 2  Hidden
End Sub

Private Sub CheckBoxReadOnly_Click()
  If CheckBoxReadOnly.Value = True Then
    MessageText = "Enabling this option will allow SolidWorks to modify files " _
                & "that are marked as ReadOnly (Normally not modifiable) " _
                & Chr$(13) & Chr$(13) _
                & "This procedure may not be a permissable within your department " _
                & "and/or organization." _
                & Chr$(13) & Chr$(13) _
                & "Continue?"
    Title = "Warning: Change Read/Only files?"
    Answer = MsgBox(MessageText, vbYesNo, Title)
    If Answer <> vbYes Then
      CheckBoxReadOnly.Value = False
    End If
  End If
End Sub

Private Sub ComboBoxPresets_Change()
  ListBoxContent.Clear
  Open Source For Input As #1             ' Open source file.
  Do While Not EOF(1)                     ' Loop until end of file.
    Input #1, Reader                      ' Read data line.
    If Reader = "[" & ComboBoxPresets.Value & "]" Then ' section.
      Do While Not EOF(1)                 ' Loop until end of file.
        Input #1, PropName                ' Read data line.
        If PropName <> "" Then            ' Is name valid?
          ListBoxContent.AddItem PropName
          For Cnt = 1 To 7
            Input #1, Data
              ListBoxContent.List(ListBoxContent.ListCount - 1, Cnt) = Data
          Next Cnt
        Else
          GoTo EndOptRead                 ' No more data for section
        End If
      Loop
    End If
  Loop
EndOptRead:
  Close #1    ' Close file.
  VerifyListBoxContent
End Sub

Private Sub CommandProcess_Click()
  ProcessParts = False
  ProcessAssemblies = False
  ProcessDrawings = False
  For Row = 0 To ListBoxContent.ListCount - 1       ' What do we process?
    If ListBoxContent.List(Row, 5) = 0 Then ProcessParts = True
    If ListBoxContent.List(Row, 5) = 1 Then ProcessAssemblies = True
    If ListBoxContent.List(Row, 5) = 2 Then ProcessDrawings = True
  Next Row
  If ProcessParts = True Then ProcessDocuments "*.sldprt", swDocPART
  If ProcessAssemblies = True Then ProcessDocuments "*.sldasm", swDocASSEMBLY
  If ProcessDrawings = True Then ProcessDocuments "*.slddrw", swDocDRAWING
End Sub

Private Sub ProcessDocuments(Extension As String, MyFileType As Integer)
  FileName = Dir(WorkDirectory + Extension)        ' gets file list
  Do While FileName <> ""
    OpenRO = False
    FileRO = False
    Set Document = swApp.OpenDoc6(WorkDirectory & FileName, _
                                  MyFileType, swOpenDocOptions_Silent, _
                                  "Default", FileError, FileWarn)
    If Not Document Is Nothing Then
      If Document.IsOpenedReadOnly() Then OpenRO = True
        If (OpenRO = False) Or (CheckBoxReadOnly.Value = True) Then
        If OpenRO = True Then             ' Was file loaded ReadOnly?
          FileAttribRead WorkDirectory + FileName, 1
                                          ' If file is loaded ReadOnly
          If FileRO = True Then           ' If file on disk is ReadOnly,
            FileAttrib WorkDirectory + FileName, 1, 0
                                          ' Change attrib of file on disk.
            swApp.CloseDoc FileName       ' Close file in SolidWorks
                                          ' Reload file into SolidWorks
            Set Document = swApp.OpenDoc6(WorkDirectory & FileName, _
                              MyFileType, swOpenDocOptions_Silent, _
                              "Default", FileError, FileWarn)
            ProcessChanges Document, MyFileType ' Process file per rules
            Document.Save2 swSaveAsOptions_Silent ' Save the document
            Set Document = Nothing        ' Clear Document object
            swApp.CloseDoc FileName       ' Close the file
            FileAttrib WorkDirectory + FileName, 1, 1
                                          ' Change attrib of file on disk.
          Else        ' FileRO            ' File loaded as ReadOnly,
                                          ' file on disk is not ReadOnly.
                                          ' Skip because file is in use.
            Set Document = Nothing        ' Clear Document object
            swApp.QuitDoc FileName        ' Close the file
          End If      ' FileRO
        Else          ' OpenRO            ' File on disk has Read-Write access
          ProcessChanges Document, MyFileType ' Process file per rules
          Document.Save2 swSaveAsOptions_Silent     ' Save the document
          Set Document = Nothing          ' Clear Document object
          swApp.QuitDoc FileName          ' Close the file
        End If        ' Nothing Open
      Else              ' PropCheck "REVISION" Skip file because property exists
        Set Document = Nothing            ' Clear Document object
        swApp.QuitDoc FileName            ' Close the file
      End If          ' PropCheck "REVISION"
    End If            ' Selected
    PropertyEditorGlobal1.CloseAll
    FileName = Dir
  Loop
End Sub

Private Sub ProcessChanges(ThisDoc, MyFileType As Integer)
  For Row = 0 To ListBoxContent.ListCount - 1            ' What do we process?
    If ListBoxContent.List(Row, 5) + 1 = MyFileType Then ' Process this step
      ' Note: Value in ListBoxContent.List(ROW, 5) is offset one less than
      '       swconst.bas values to match position in list boxes on prop forms
      '       This must be adjusted to accommodate the API requirements.
      PropName = ListBoxContent.List(Row, 2)
      PropType = CstmPropType(ListBoxContent.List(Row, 3))
      PropValu = ListBoxContent.List(Row, 4)
      PropCatg = ListBoxContent.List(Row, 6)
      Overwrite = ListBoxContent.List(Row, 7)
      TempValu = ""
      Select Case ListBoxContent.List(Row, 1)
          Case 1 ' Add property to current document
                 If PropCatg = 0 Then
                   ' Add property to ensure it exists
                   ThisDoc.AddCustomInfo3 "", PropName, _
                            PropType, PropValu
                   Else
                     ConfNames = ThisDoc.GetConfigurationNames
                     For i = 0 To UBound(ConfNames)
                       ' Add property to ensure it exists
                       ThisDoc.AddCustomInfo3 ConfNames(i), PropName, _
                             PropType, PropValu
                      Next i
                  End If
                  Select Case Overwrite
                      Case 0    ' Add only.
                                ' No additional procedures
                      Case 1    ' Add or overwrite only if blank
                             If PropCatg = 0 Then
                               ' Get existing property value
                               TempValu = ThisDoc.CustomInfo2("", PropName)
                               ' Is it blank?
                               If TempValu = "" Or TempValu = " " Then
                                 ' Write new value if existing value is blank
                                 ThisDoc.CustomInfo2("", PropName) = PropValu
                               End If
                             Else
                               For i = 0 To UBound(ConfNames)
                                 ' Get existing property value
                                 TempValu = ThisDoc.CustomInfo2(ConfNames(i), _
                                        PropName)
                                 ' Is it blank?
                                 If TempValu = "" Or TempValu = " " Then
                                   ' Write new value if existing value is blank
                                   ThisDoc.CustomInfo2(ConfNames(i), PropName) _
                                        = PropValu
                                 End If
                               Next i
                             End If
                      Case 2    ' Add or overwrite all values
                             If PropCatg = 0 Then
                               ThisDoc.CustomInfo2("", PropName) = PropValu
                             Else
                               For i = 0 To UBound(ConfNames)
                                 ThisDoc.CustomInfo2(ConfNames(i), PropName) _
                                        = PropValu
                               Next i
                             End If
                  End Select
          Case 2 ' Revise property value
                 If PropCatg = 0 Then
                   ' Custom File Property
                   ThisDoc.CustomInfo2("", PropName) = PropValu
                 Else
                   ' Configuration Specific Property
                   ConfNames = ThisDoc.GetConfigurationNames
                   For i = 0 To UBound(ConfNames)
                     ThisDoc.CustomInfo2(ConfNames(i), PropName) = PropValu
                   Next i
                 End If
          Case 3 ' Delete property
                 If PropCatg = 0 Then
                   ThisDoc.DeleteCustomInfo2 "", PropName
                 Else
                   ConfNames = ThisDoc.GetConfigurationNames
                   For i = 0 To UBound(ConfNames)
                     ThisDoc.DeleteCustomInfo2 ConfNames(i), PropName
                   Next i
                 End If
          Case 4 ' Rename property
                 If PropCatg = 0 Then
                   ' Custom File Property
                   TempValu = ThisDoc.CustomInfo2("", PropName)
                   TempType = ThisDoc.GetCustomInfoType3("", PropName)
                   ThisDoc.DeleteCustomInfo2 "", PropName
                   ThisDoc.AddCustomInfo3 "", PropValu, TempType, TempValu
                 Else
                   ConfNames = ThisDoc.GetConfigurationNames
                   For i = 0 To UBound(ConfNames)
                     TempValu = ThisDoc.CustomInfo2(ConfNames(i), PropName)
                     TempType = ThisDoc.GetCustomInfoType3(ConfNames(i), PropName)
                     ThisDoc.DeleteCustomInfo2 ConfNames(i), PropName
                   Next i
                 End If
          Case Else
      End Select
    End If                                              ' Step Processed
  Next Row
End Sub

Private Sub CommandStepCopy_Click()
  ListBoxContent.AddItem ""                         ' Add step
  For Column = 0 To 7
    ListBoxContent.List(ListBoxContent.ListCount - 1, Column) = _
            ListBoxContent.List(ListBoxContent.ListIndex, Column)
  Next Column
  ListBoxContent.ListIndex = ListBoxContent.ListCount - 1
  CommandStepEdit_Click                             ' Edit new step
  If StepAdded = False Then                         ' Delete step if cancelled
    ListBoxContent.RemoveItem (ListBoxContent.ListCount - 1)
  End If
End Sub

Private Sub CommandStepEdit_Click()
  ListPosition = ListBoxContent.ListIndex
  StepEdit = True
  Select Case ListBoxContent.List(ListBoxContent.ListCount - 1, 1)
      Case 1
             Form_GlobalPropAdd.Show
      Case 2
             Form_GlobalPropRevise.Show
      Case 3
             Form_GlobalPropDelete.Show
      Case 4
             Form_GlobalPropRename.Show
      Case Else
  End Select
End Sub

Private Sub CommandStepRemove_Click()
  ListBoxContent.RemoveItem (ListBoxContent.ListIndex)
  VerifyListBoxContent
End Sub

Private Sub CommandMoveUp_Click()
  For Column = 0 To 7
    TempValue = ListBoxContent.List(ListBoxContent.ListIndex - 1, Column)
    ListBoxContent.List(ListBoxContent.ListIndex - 1, Column) = _
            ListBoxContent.List(ListBoxContent.ListIndex, Column)
    ListBoxContent.List(ListBoxContent.ListIndex, Column) = TempValue
  Next Column
  ListBoxContent.ListIndex = ListBoxContent.ListIndex - 1
End Sub

Private Sub CommandMoveDown_Click()
  For Column = 0 To 7
    TempValue = ListBoxContent.List(ListBoxContent.ListIndex + 1, Column)
    ListBoxContent.List(ListBoxContent.ListIndex + 1, Column) = _
            ListBoxContent.List(ListBoxContent.ListIndex, Column)
    ListBoxContent.List(ListBoxContent.ListIndex, Column) = TempValue
  Next Column
  ListBoxContent.ListIndex = ListBoxContent.ListIndex + 1
End Sub

Private Sub ListBoxContent_Click()
  VerifyListBoxContent
End Sub

Private Sub CommandAdd_Click()
  ' Routine to create "ADD PROPERTY" step
  StepEdit = False
  Form_GlobalPropAdd.Show
End Sub

Private Sub CommandDelete_Click()
  ' Routine to create "DELETE PROPERTY" step
  StepEdit = False
  Form_GlobalPropDelete.Show
End Sub

Private Sub CommandRename_Click()
  ' Routine to create "RENAME PROPERTY" step
  StepEdit = False
  Form_GlobalPropRename.Show
End Sub

Private Sub CommandRevise_Click()
  ' Routine to create "REVISE PROPERTY VALUE" step
  StepEdit = False
  Form_GlobalPropRevise.Show
End Sub

Private Sub ControlLock(Control As Object, Mode As Boolean)
  Control.Locked = Mode
  If Control.Locked = True Then
    Control.BackColor = vbButtonFace
  Else
    Control.BackColor = vbWindowBackground
  End If
End Sub

Private Sub CommandClose_Click()
  End
End Sub

Private Sub VerifyListBoxContent()
  FrameModifyRule.Enabled = True                    ' Turn everything on
  CommandStepEdit.Enabled = True
  CommandStepCopy.Enabled = True
  CommandStepRemove.Enabled = True
  CommandMoveUp.Enabled = True
  CommandMoveDown.Enabled = True
  CommandExportRules.Enabled = True
  CommandProcess.Enabled = True
  If ListBoxContent.ListCount < 1 Then              ' Turn off not needed
    FrameModifyRule.Enabled = False
    CommandStepEdit.Enabled = False
    CommandStepRemove.Enabled = False
    CommandStepCopy.Enabled = False
    CommandMoveUp.Enabled = False
    CommandMoveDown.Enabled = False
    CommandExportRules.Enabled = False
    CommandProcess.Enabled = False
  ElseIf ListBoxContent.ListCount = 1 Then          ' Turn off not needed
    CommandMoveUp.Enabled = False
    CommandMoveDown.Enabled = False
  End If
  If ListBoxContent.ListIndex = 0 Then
    CommandMoveUp.Enabled = False
  ElseIf ListBoxContent.ListIndex = ListBoxContent.ListCount - 1 Then
    CommandMoveDown.Enabled = False
  End If
End Sub

Sub ReadOptions()
  ComboBoxPresets.Clear
  ComboBoxPresets.AddItem "<- None - >"
  Open Source For Input As #1             ' Open source file.
  Do While Not EOF(1)                     ' Loop until end of file.
    Input #1, Reader                      ' Read data line.
    If Reader = "[OPTIONS]" Then          ' Looking for section.
      Do While Not EOF(1)                 ' Loop until end of file.
        Input #1, PropName                ' Read data line.
        If PropName <> "" Then            ' Is name valid?
          If UCase(Left$(PropName, 20)) = UCase("ForceUpperCaseValues") Then
            If UCase(Right$(PropName, 4)) = "TRUE" Then
              CheckBoxUpperCase.Value = True
            End If
          ElseIf UCase(Left$(PropName, 20)) = UCase("AllowUpperCaseChange") Then
            If UCase(Right$(PropName, 4)) = "TRUE" Then
              CheckBoxUpperCase.Enabled = True
            Else
              CheckBoxUpperCase.Enabled = False
            End If
          End If
        Else
          GoTo EndOptRead                 ' No more data for section
        End If
      Loop
EndOptRead:
    ElseIf Reader = "[GLOBAL RULES]" Then ' Looking for section.
      Do While Not EOF(1)                 ' Loop until end of file.
        Input #1, RuleName                ' Read data line.
        If RuleName <> "" Then            ' Is name valid?
          ComboBoxPresets.AddItem RuleName
        Else
          GoTo EndRuleRead                ' No more data for section
        End If
      Loop
    End If
EndRuleRead:
  Loop
  Close #1    ' Close file.
  If ComboBoxPresets.ListCount > 0 Then
    FramePresetRules.Enabled = True
    ComboBoxPresets.Enabled = True
    ComboBoxPresets.BackColor = vbWindowBackground
  End If
End Sub

Private Sub CommandExportRules_Click()                              ' v2.0c
  Form_GlobalExportRules.Show                                       ' routine
End Sub                                                             ' added

Private Sub CommandAbout_Click()
  Form_About.Show
End Sub

Private Sub LSRBanner_Click()
  Form_About.Show
End Sub

Private Sub UserForm_Initialize()
  ProgramVersion.Caption = "Version: v" & Version
  WorkDirectory = swApp.GetCurrentWorkingDirectory
  FramePresetRules.Enabled = False
  ComboBoxPresets.Enabled = False
  ComboBoxPresets.BackColor = vbButtonFace
  If SourceExists = True Then
    ReadOptions
    ComboBoxPresets.ListIndex = 0
  End If
  PropAdd = False
  VerifyListBoxContent
  ' Remove comment indicator from line below to see hidden content in list box.
  ' ListBoxContent.ColumnWidths = _
    "60 pt;20 pt;60 pt;20 pt;60 pt;20 pt;20 pt;20 pt;20 pt;20 pt"
End Sub

