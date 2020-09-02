VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_GlobalPropRename 
   Caption         =   "PropertyEditorGlobal: RENAME PROP"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3990
   OleObjectBlob   =   "Form_GlobalPropRename.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Form_GlobalPropRename"
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
' This module is only used for adding new properties.
' -----------------------------------------------------------------------------
' Version:  2.0a  * Added "Global Properties" capabilities
' -----------------------------------------------------------------------------
Dim ListBoxContent          As Object

Private Sub ComboBoxNewName_Change()
  CheckSelections
End Sub

Private Sub CommandAdd_Click()
  Dim SelectTyp         As Boolean
  Dim SelectDoc         As Boolean
  Dim SelectCat         As Boolean
  Dim DisplayText       As String
  SelectTyp = False
  SelectDoc = False
  SelectCat = False
  
  For X = 0 To ListBoxDocType.ListCount - 1
    If ListBoxDocType.Selected(X) Then SelectDoc = True
  Next X
  For X = 0 To ListBoxCategory.ListCount - 1
    If ListBoxCategory.Selected(X) Then SelectCat = True
  Next X
  
  If (SelectDoc = True) And (SelectCat = True) Then
    DisplayText = "RENAME: "
    DisplayText = DisplayText _
                & "(" & ComboBoxOldName.Value & ") " & ListBoxCategory.Value _
                & " to (" & ComboBoxNewName.Value & ")" _
                & " in all " & ListBoxDocType.Value
    If StepEdit = False Then
      ListBoxContent.AddItem ""
      ListBoxContent.ListIndex = ListBoxContent.ListCount - 1
      ListPosition = ListBoxContent.ListIndex
    End If
    ListBoxContent.List(ListPosition, 0) = DisplayText
    ListBoxContent.List(ListPosition, 1) = 4
    ListBoxContent.List(ListPosition, 2) = ComboBoxOldName.Value
    ListBoxContent.List(ListPosition, 3) = 0        ' Not used
    ListBoxContent.List(ListPosition, 4) = ComboBoxNewName.Value
    ListBoxContent.List(ListPosition, 5) = ListBoxDocType.ListIndex
    ListBoxContent.List(ListPosition, 6) = ListBoxCategory.ListIndex
    ListBoxContent.List(ListPosition, 7) = True
    StepAdded = True
    Me.Hide
  Else
    MsgBox "Incomplete.  Check your selections above."
  End If
End Sub

Private Sub CommandCancel_Click()
  StepAdded = False
  Me.Hide
End Sub

Private Sub ComboBoxOldName_Change()
  If ComboBoxOldName.Value = "" Then
    CommandAdd.Enabled = False
  Else
    CommandAdd.Enabled = True
    If ComboBoxOldName.ListIndex > -1 Then
      ListBoxType.ListIndex = -1
      For X = 0 To ListBoxType.ListCount - 1
        If ComboBoxOldName.List(ComboBoxOldName.ListIndex, 1) = ListBoxType.List(X, 1) Then
          ListBoxType.ListIndex = X
        End If
      Next
    End If
  End If
  CheckSelections
End Sub

Private Sub GetList()
' StatPart "Reading property list.", 0
  Dim PropName As String, PropType As Integer
  ComboBoxOldName.Clear
  Open Source For Input As #1               ' Open source file.
  Do While Not EOF(1)                       ' Loop until end of file.
    Input #1, Reader                        ' Read data line.
    If Reader = "[SPECIAL PROPERTIES]" Then ' Looking for section.
      Do While Not EOF(1)                   ' Loop until end of file.
        Input #1, PropName                  ' Read data line.
        If PropName <> "" Then              ' Is name valid?
          Input #1, PropType                ' Read data line.
          ComboBoxOldName.AddItem PropName     ' Add to combo box
          ComboBoxOldName.List(ComboBoxOldName.ListCount - 1, 1) = PropType
        Else
          GoTo EndRead0                     ' No more data for section
        End If
      Loop
EndRead0:
    End If
  Loop
  Close #1    ' Close file.
End Sub

Private Sub ListBoxCategory_Click()
  CheckSelections
End Sub

Private Sub ListBoxDocType_Click()
  If ListBoxDocType.ListIndex = 2 Then
    ListBoxCategory.ListIndex = 0
    ListBoxCategory.Enabled = False
  Else
    ListBoxCategory.Enabled = True
  End If
  CheckSelections
End Sub

Private Sub CheckSelections()
  SelectDoc = False
  SelectCat = False
  For X = 0 To ListBoxDocType.ListCount - 1         ' Is document type selected?
    If ListBoxDocType.Selected(X) Then SelectDoc = True
  Next X
  For X = 0 To ListBoxCategory.ListCount - 1        ' Is prop category selected?
    If ListBoxCategory.Selected(X) Then SelectCat = True
  Next X
  If (SelectDoc = True) And (SelectCat = True) _
        And (ComboBoxOldName.Value <> "") And (ComboBoxNewName.Value <> "") Then
    ' All selections are valid?
    CommandAdd.Enabled = True
  Else
    CommandAdd.Enabled = False
  End If
End Sub

Private Sub UserForm_Activate()
  Set ListBoxContent = Form_GlobalPropEditor.ListBoxContent
  If SourceExists = True Then
    GetList
  End If
  ListBoxDocType.Clear
  ListBoxDocType.AddItem "Parts"
  ListBoxDocType.AddItem "Assemblies"
  ListBoxDocType.AddItem "Drawings"
  ListBoxCategory.Clear
  ListBoxCategory.AddItem "Custom File Property"
  ListBoxCategory.AddItem "Config Specific Property"
  If StepEdit = True Then
    ComboBoxOldName.Value = _
                ListBoxContent.List(ListPosition, 2)
    ComboBoxNewName.Value = _
                ListBoxContent.List(ListPosition, 4)
    ListBoxDocType.ListIndex = _
                ListBoxContent.List(ListPosition, 5)
    ListBoxCategory.ListIndex = _
                ListBoxContent.List(ListPosition, 6)
    CommandAdd.Caption = "Apply Change"
  Else
    ComboBoxOldName.Value = ""
    ComboBoxNewName.Value = ""
  End If
  ComboBoxOldName.SetFocus
  ComboBoxOldName_Change
End Sub
