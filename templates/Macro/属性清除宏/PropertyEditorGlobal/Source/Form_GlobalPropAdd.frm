VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_GlobalPropAdd 
   Caption         =   "PropertyEditorGlobal: ADD PROP"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3990
   OleObjectBlob   =   "Form_GlobalPropAdd.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Form_GlobalPropAdd"
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
Dim Initialize              As Boolean

Private Sub CommandAdd_Click()
  Dim SelectTyp         As Boolean
  Dim SelectDoc         As Boolean
  Dim SelectCat         As Boolean
  Dim DisplayText       As String
  Select Case ListBoxOverwrite.ListIndex
         Case 0
                DisplayText = "ADD: "
         Case 1
                DisplayText = "ADD/OVERWRITE(Blank): "
         Case 2
                DisplayText = "ADD/OVERWRITE(All): "
  End Select
  DisplayText = DisplayText _
              & "(" & ComboBoxName.Value & ") " & ListBoxCategory.Value _
              & " with value of (" & TextBoxValue.Value & ")" _
              & " to all " & ListBoxDocType.Value
  If StepEdit = False Then                        ' If were not editing a
    ListBoxContent.AddItem ""                     ' step, then add a step
    ListBoxContent.ListIndex = ListBoxContent.ListCount - 1
    ListPosition = ListBoxContent.ListIndex       ' Reset ListIndex
  End If
  ' Text displayed to user
  ListBoxContent.List(ListPosition, 0) = DisplayText
  ' Action code (for this macro)
  ListBoxContent.List(ListPosition, 1) = 1
  ' Property name
  ListBoxContent.List(ListPosition, 2) = ComboBoxName.Value
  ' Property Type
  ListBoxContent.List(ListPosition, 3) = ListBoxType.ListIndex
  ' Property value
  Select Case ListBoxType.ListIndex
         Case 0 ' TEXT
                ListBoxContent.List(ListPosition, 4) = TextBoxValue.Value
         Case 1 ' DATE
                ListBoxContent.List(ListPosition, 4) = TextBoxValue.Value
         Case 2 ' NUMBER
                ListBoxContent.List(ListPosition, 4) = TextBoxValue.Value
         Case 3 ' YES/NO
                If OptionYes.Value = True Then
                  ListBoxContent.List(ListPosition, 4) = "Yes"
                Else
                  ListBoxContent.List(ListPosition, 4) = "No"
                End If
  End Select
  ' Document type
  ListBoxContent.List(ListPosition, 5) = ListBoxDocType.ListIndex
  ' Property Category
  ListBoxContent.List(ListPosition, 6) = ListBoxCategory.ListIndex
  ' Overwrite flag
  ListBoxContent.List(ListPosition, 7) = ListBoxOverwrite.ListIndex
  StepAdded = True
  Me.Hide
End Sub

Private Sub CommandCancel_Click()
  StepAdded = False
  Me.Hide
End Sub

Private Sub ComboBoxName_Change()
  If Initialize = False Then
    If ComboBoxName.Value = "" Then
      CommandAdd.Enabled = False
    Else
      CommandAdd.Enabled = True
        ListBoxType.Enabled = False
        ListBoxType.BackColor = vbButtonFace
      If ComboBoxName.ListIndex > -1 Then
        ListBoxType.ListIndex = -1
        For X = 0 To ListBoxType.ListCount - 1
          If ComboBoxName.List(ComboBoxName.ListIndex, 1) = ListBoxType.List(X, 1) Then
            ListBoxType.ListIndex = X
          End If
        Next
      Else
        ListBoxType.Enabled = True
        ListBoxType.BackColor = vbWindowBackground
        ListBoxType.ListIndex = 1
        ListBoxType.ListIndex = 0
      End If
    End If
    CheckSelections
  End If
End Sub

Private Sub GetList()
' StatPart "Reading property list.", 0
  Dim PropName As String, PropType As Integer
  ComboBoxName.Clear
  Open Source For Input As #1               ' Open source file.
  Do While Not EOF(1)                       ' Loop until end of file.
    Input #1, Reader                        ' Read data line.
    If Reader = "[SPECIAL PROPERTIES]" Then ' Looking for section.
      Do While Not EOF(1)                   ' Loop until end of file.
        Input #1, PropName                  ' Read data line.
        If PropName <> "" Then              ' Is name valid?
          Input #1, PropType                ' Read data line.
          ComboBoxName.AddItem PropName     ' Add to combo box
          ComboBoxName.List(ComboBoxName.ListCount - 1, 1) = PropType
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

Private Sub ComboDateSelect_Change()
  Select Case ComboDateSelect.ListIndex
         Case 0 ' Today
           TextBoxValue = Format(Now, "mm/dd/yyyy")
         Case 1 ' Other
           Form_DocDate.Show
           If NewDate <> "00/00/0000" Then TextBoxValue = NewDate
  End Select
End Sub

Private Sub ListBoxOverwrite_Click()
  CheckSelections
End Sub

Private Sub ListBoxType_Click()
  TextBoxValue.Enabled = True
  TextBoxValue.Visible = False
  TextBoxValue.BackColor = vbWindowBackground
  TextBoxValue.Width = 128
  ComboDateSelect.Visible = False
  OptionYes.Visible = False
  OptionNo.Visible = False
  Select Case ListBoxType.ListIndex
         Case 0 ' TEXT
                TextBoxValue.Visible = True
         Case 1 ' DATE
                ComboDateSelect.Visible = True
                TextBoxValue.Enabled = False
                TextBoxValue.Visible = True
                TextBoxValue.BackColor = vbButtonFace
                TextBoxValue.Width = 64
         Case 2 ' NUMBER
                TextBoxValue.Visible = True
                TextBoxValue.Value = Val(TextBoxValue.Value)
         Case 3 ' YESNO
                OptionYes.Visible = True
                OptionNo.Visible = True
  End Select
  CheckSelections
End Sub

Private Sub TextBoxValue_Change()
  If Form_GlobalPropEditor.CheckBoxUpperCase.Value = True Then
    TextBoxValue.Value = UCase(TextBoxValue.Value)
  End If
End Sub

Private Sub TextBoxValue_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
  If ListBoxType.ListIndex = 2 Then
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 46 Then KeyAscii = 0
  End If
  If KeyAscii = 13 Then
    TextBoxValue.Value = TextBoxValue.Value + Chr$(13)
  End If
End Sub

Private Sub CheckSelections()
  SelectTyp = False
  SelectDoc = False
  SelectCat = False
  SelectOWT = False
  For X = 0 To ListBoxType.ListCount - 1            ' Is property type selected?
    If ListBoxType.Selected(X) Then SelectTyp = True
  Next X
  For X = 0 To ListBoxDocType.ListCount - 1         ' Is document type selected?
    If ListBoxDocType.Selected(X) Then SelectDoc = True
  Next X
  For X = 0 To ListBoxCategory.ListCount - 1        ' Is prop category selected?
    If ListBoxCategory.Selected(X) Then SelectCat = True
  Next X
  For X = 0 To ListBoxOverwrite.ListCount - 1       ' Is overwrite opt selected?
    If ListBoxOverwrite.Selected(X) Then SelectOWT = True
  Next X
  If (SelectTyp = True) And (SelectDoc = True) And (SelectCat = True) _
        And (SelectOWT = True) And (ComboBoxName.Value <> "") Then
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
  ListBoxType.Clear
  ListBoxType.AddItem "Text"
  ListBoxType.List(ListBoxType.ListCount - 1, 1) = 30
  ListBoxType.AddItem "Date"
  ListBoxType.List(ListBoxType.ListCount - 1, 1) = 64
  ListBoxType.AddItem "Number"
  ListBoxType.List(ListBoxType.ListCount - 1, 1) = 3
  ListBoxType.AddItem "Yes/No"
  ListBoxType.List(ListBoxType.ListCount - 1, 1) = 11
  ComboDateSelect.Clear
  ComboDateSelect.AddItem "Today"
  ComboDateSelect.AddItem "Other"
  ListBoxDocType.Clear
  ListBoxDocType.AddItem "Parts"
  ListBoxDocType.AddItem "Assemblies"
  ListBoxDocType.AddItem "Drawings"
  ListBoxCategory.Clear
  ListBoxCategory.AddItem "Custom File Property"
  ListBoxCategory.AddItem "Config Specific Property"
  ListBoxOverwrite.Clear
  ListBoxOverwrite.AddItem "Add only."
  ListBoxOverwrite.AddItem "Add and Overwrite blank."
  ListBoxOverwrite.AddItem "Add and Overwrite all."
  If StepEdit = True Then
    Initialize = True
    ListPosition = ListBoxContent.ListIndex       ' Reset ListIndex
    ' Property name
    ComboBoxName.Value = _
                ListBoxContent.List(ListPosition, 2)
    ' Property type
    ListBoxType.ListIndex = _
                ListBoxContent.List(ListPosition, 3)
    ListBoxType_Click
    ' Property value
    Select Case ListBoxType.ListIndex
           Case 0 ' TEXT
                  TextBoxValue.Value = _
                      ListBoxContent.List(ListPosition, 4)
           Case 1 ' DATE
                  TextBoxValue.Value = _
                      ListBoxContent.List(ListPosition, 4)
           Case 2 ' NUMBER
                  TextBoxValue.Value = _
                      ListBoxContent.List(ListPosition, 4)
           Case 3 ' YESNO
                  If ListBoxContent.List(ListPosition, 4) = True Then
                    OptionYes.Value = True
                    OptionNo.Value = False
                  Else
                    OptionYes.Value = False
                    OptionNo.Value = True
                  End If
    End Select
    ' Document type
    ListBoxDocType.ListIndex = _
                ListBoxContent.List(ListPosition, 5)
    ' Property category
    ListBoxCategory.ListIndex = _
                ListBoxContent.List(ListPosition, 6)
    ' Overwrite flag
    ListBoxOverwrite.ListIndex = _
                ListBoxContent.List(ListPosition, 7)
    CommandAdd.Caption = "Apply Change"
  Else
    ListBoxType.ListIndex = 0
    ListBoxType_Click
    ComboBoxName.Value = ""
    TextBoxValue.Value = ""
    CommandAdd.Caption = "Add Step"
  End If
  ComboBoxName.SetFocus
  ComboBoxName_Change
  Initialize = False
End Sub
