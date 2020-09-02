VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_GlobalExportRules 
   Caption         =   "PropertyEditorGlobal: RULE EXPORT"
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4785
   OleObjectBlob   =   "Form_GlobalExportRules.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Form_GlobalExportRules"
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
' This module is only used for exporting global property rules.
' -----------------------------------------------------------------------------
' Version:  2.0c  * Added "Global Properties" rules export capabilities
' -----------------------------------------------------------------------------

Private Sub CommandCancel_Click()
  Me.Hide
End Sub

Private Sub CommandExport_Click()
  TextBoxRulesName.Value = ""
  Dim PrintText As String
  Set ListBoxContent = Form_GlobalPropEditor.ListBoxContent
  FileName = WorkDirectory + "GlobalPropEdRules.TXT"
  Open FileName For Append Access Write Lock Write As #1
  Print #1, "' The following lines must be added to"
  Print #1, "' " & Source
  Print #1, "' in the appropriate section to be used in the macro."
  Print #1, ""
  Print #1, "' Add the 1 line below under section [GLOBAL RULES]"
  Print #1, Chr(34) & TextBoxRulesName.Value & Chr(34)
  Print #1, ""
  Print #1, "' Add all lines below at end of file for new section"
  Print #1, "[" & TextBoxRulesName.Value & "]"
  For Row = 0 To ListBoxContent.ListCount - 1
    PrintText = Chr(34) & ListBoxContent.List(Row, 0) & Chr(34) & ","
    For Col = 1 To 6
      PrintText = PrintText & ListBoxContent.List(Row, Col) & ","
    Next Col
    PrintText = PrintText & ListBoxContent.List(Row, 7)
    Print #1, PrintText
  Next Row
  Print #1, vbNewLine
  Close #1
  MsgBox "Rules exported to: " & Chr(13) & Chr(13) & FileName _
         & Chr(13) & Chr(13) & "To use this data, as a preset rule, " _
         & "in this macro, you must manually copy the data from the " _
         & "file above, into the file:" _
         & Chr(13) & Chr(13) & Source
  Me.Hide
End Sub
