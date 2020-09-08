Imports System.IO
Imports SolidWorks.Interop.swdocumentmgr
'Imports System.Management.Automation
'Imports Microsoft.VisualBasic.CompilerServices

Public Class BatchPropForm
    Private LogPath As String = Application.StartupPath & "\Log.txt"
    Private swDocMgr As SwDMApplication
    Private DontUpdate As Boolean


    Private dtDgv1 As New DataTable()
    Private dtDgv2 As New DataTable()

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.Text = "Batch Custom Properties V0.1"

        ' ------------------handle to catch application crashes
        ' Get the your application's application domain.
        'Dim currentDomain As AppDomain = AppDomain.CurrentDomain
        '' Define a handler for unhandled exceptions.
        'AddHandler currentDomain.UnhandledException, AddressOf MYExHandler
        '' Define a handler for unhandled exceptions for threads behind forms.
        'AddHandler Application.ThreadException, AddressOf MYThreadHandler

        If File.Exists(LogPath) Then File.Delete(LogPath)

        'Load last search folder
        TbFolder.Text = My.Settings.SearchFolder

        'Load search filter
        TbFilter.Text = My.Settings.SearchFilter
        If TbFilter.Text = String.Empty Then TbFilter.Text = "*.sldprt,*.sldasm"

        'set check boxes
        CbLoadColumns.Checked = My.Settings.LoadColumns
        CbIncludeSubFolders.Checked = My.Settings.IncludeSubFolder
        CbLoadCustomProp.Checked = My.Settings.LoadCustomProp
        CbLoadAllConfig.Checked = My.Settings.LoadAllConfig
        CbLoadActiveConfig.Checked = My.Settings.LoadActiveConfig

        'load license key
        LoadLicenseKey()

        Log("Form loaded")
    End Sub



    Private Sub ResetDGVs()

        dtDgv1.Clear()
        dtDgv1.Columns.Clear()
        DGV1.DataSource = dtDgv1

        dtDgv1.Columns.Add("FilePath", GetType(String))
        dtDgv1.Columns.Add("FileName", GetType(String))

        dtDgv1.Rows.Add("Type")
        dtDgv1.Rows.Add("Config")
        dtDgv1.Rows.Add("Rules")

        'Freeze first 2 columns
        DGV1.Columns(1).Frozen = True

        'prevent first 2 columns' cells editing
        DGV1.Columns(0).ReadOnly = True
        DGV1.Columns(1).ReadOnly = True

        'prevent row's cells editing
        DGV1.Rows(0).ReadOnly = True
        DGV1.Rows(1).ReadOnly = True


        If CbLoadColumns.Checked Then LoadColumns()

        '====================
        dtDgv2.Clear()
        dtDgv2.Columns.Clear()

        dtDgv2 = dtDgv1.Clone()
        DGV2.DataSource = dtDgv2

        'set first 2 columns columns width
        DGV2.Columns(0).Width = 100
        DGV2.Columns(1).Width = 100

        'Freeze first 2 columns
        DGV2.Columns(1).Frozen = True

        'prevent first 2 columns' cells editing
        DGV2.Columns(0).ReadOnly = True
        DGV2.Columns(1).ReadOnly = True

        DGV2.ScrollBars = ScrollBars.Vertical

    End Sub

    Private Sub BtSearch_Click(sender As Object, e As EventArgs) Handles BtSearch.Click
        DontUpdate = True
        ResetDGVs()


        'AddFirstRows()

        If Not Directory.Exists(TbFolder.Text) Then
            MessageBox.Show("Can not find folder " & TbFolder.Text)
            Exit Sub
        End If

        'Save search folder
        My.Settings.SearchFolder = TbFolder.Text

        'Save search filter
        My.Settings.SearchFilter = TbFilter.Text

        'set files search filter based on radio button choice
        Dim myFilters As List(Of String) = TbFilter.Text.Split(",").ToList

        'search files
        Dim FilePaths = SearchForFiles(TbFolder.Text, myFilters)

        Log("add to list")
        For Each FilePath As String In FilePaths
            'bypass Solidworks temp files
            If Not FilePath.Contains("~$") Then AddFilesToDt(FilePath)
        Next

        ReformatDGVs()

        BtSearch.Text = "Refresh"

        DontUpdate = False
    End Sub


    Private Sub AddFilesToDt(ByVal FilePath As String)

        Log("   Add file: " & FilePath)

        Dim swDoc As SwDMDocument = OpenDocument(FilePath)
        If swDoc Is Nothing Then Exit Sub

        'add a new row with file path and name
        dtDgv2.Rows.Add(FilePath, Path.GetFileNameWithoutExtension(FilePath))

        'load general custom properties
        If CbLoadCustomProp.Checked Then AddCustomProp(swDoc)

        If CbLoadAllConfig.Checked OrElse CbLoadActiveConfig.Checked Then
            Dim swConfig As SwDMConfiguration
            Dim swConfigMgr As SwDMConfigurationMgr = swDoc.ConfigurationManager
            If CbLoadAllConfig.Checked Then
                'load all configurations custom properties
                Dim ConfigNames As String() = swConfigMgr.GetConfigurationNames
                For Each ConfigName As String In ConfigNames
                    swConfig = swConfigMgr.GetConfigurationByName(ConfigName)
                    AddCustomProp(swConfig)
                Next
            Else
                'load active configuration custom properties
                Dim ConfigName As String = swConfigMgr.GetActiveConfigurationName
                swConfig = swConfigMgr.GetConfigurationByName(ConfigName)
                AddCustomProp(swConfig)
            End If
        End If

        swDoc.CloseDoc()
    End Sub

    Private Sub AddCustomProp(swObj As Object)
        Dim swDoc As SwDMDocument = Nothing
        Dim swConfig As SwDMConfiguration = Nothing
        Dim CustPropNameArr As New Object
        Dim ConfigName As String = String.Empty

        If TypeOf swObj Is SwDMDocument Then
            swDoc = swObj
            CustPropNameArr = swDoc.GetCustomPropertyNames
        ElseIf TypeOf swObj Is SwDMConfiguration Then
            swConfig = swObj
            CustPropNameArr = swConfig.GetCustomPropertyNames
            ConfigName = swConfig.Name
        Else
            Exit Sub
        End If

        If CustPropNameArr Is Nothing Then Exit Sub

        Dim CustPropType As Long
        Dim CustPropValue As String
        Dim CustPropHeader As String

        For Each CustPropName As String In CustPropNameArr
            If swDoc IsNot Nothing Then
                CustPropValue = swDoc.GetCustomProperty(CustPropName, CustPropType)
            Else
                CustPropValue = swConfig.GetCustomProperty(CustPropName, CustPropType)
            End If
            CustPropHeader = String.Join(";", CustPropName, PropTypeToString(CustPropType), ConfigName)
            '    Dim ColumnHeaders = DGV1.Columns.Cast(Of DataGridViewColumn)().Select(Function(column) column.HeaderText.ToLower)
            If Not dtDgv1.Columns.Contains(CustPropHeader) Then '
                dtDgv1.Columns.Add(CustPropHeader, GetType(String))
                dtDgv1.Rows(0).Item(CustPropHeader) = PropTypeToString(CustPropType)
                dtDgv1.Rows(1).Item(CustPropHeader) = ConfigName

                dtDgv2.Columns.Add(CustPropHeader, GetType(String))
            End If

            dtDgv2.Rows(dtDgv2.Rows.Count - 1).Item(CustPropHeader) = CustPropValue
        Next

    End Sub

    Private Sub BtSaveColumns_Click(sender As Object, e As EventArgs) Handles BtSaveColumns.Click
        If My.Settings.ColumnsNames IsNot Nothing Then
            My.Settings.ColumnsNames.Clear()
            My.Settings.ColumnsTypes.Clear()
            My.Settings.ColumnsConfigs.Clear()
            My.Settings.ColumnsRules.Clear()
        End If
        My.Settings.ColumnsNames = New Specialized.StringCollection
        My.Settings.ColumnsTypes = New Specialized.StringCollection
        My.Settings.ColumnsConfigs = New Specialized.StringCollection
        My.Settings.ColumnsRules = New Specialized.StringCollection

        Dim ColumnsInDisplayedOrder = DGV1.Columns.Cast(Of DataGridViewColumn)().OrderBy(Function(column) column.DisplayIndex)
        For Each col As DataGridViewColumn In ColumnsInDisplayedOrder
            Dim index As Integer = col.Index
            If col.Index < 2 Then Continue For
            My.Settings.ColumnsNames.Add(col.HeaderText)
            My.Settings.ColumnsTypes.Add(dtDgv1.Rows(0).Item(index).ToString)
            My.Settings.ColumnsConfigs.Add(dtDgv1.Rows(1).Item(index).ToString)
            My.Settings.ColumnsRules.Add(dtDgv1.Rows(2).Item(index).ToString)
        Next

        My.Settings.Save()
    End Sub

    Private Sub LoadColumns()
        If My.Settings.ColumnsNames Is Nothing Then Exit Sub

        'For Each colName As String In My.Settings.ColumnsNames
        For i = 0 To My.Settings.ColumnsNames.Count - 1
            Dim CustPropHeader As String = My.Settings.ColumnsNames(i)
            If dtDgv1.Columns.Contains(CustPropHeader) Then Continue For

            dtDgv1.Columns.Add(CustPropHeader, GetType(String))
            dtDgv1.Rows(0).Item(CustPropHeader) = My.Settings.ColumnsTypes(i)
            dtDgv1.Rows(1).Item(CustPropHeader) = My.Settings.ColumnsConfigs(i)
            dtDgv1.Rows(2).Item(CustPropHeader) = My.Settings.ColumnsRules(i)
        Next
    End Sub


    'Private Sub MYExHandler(ByVal sender As Object, ByVal e As UnhandledExceptionEventArgs)
    '    ' catch application error
    '    'MessageBox.Show("An error has occured." & vbcr & "See Log for more informations")
    '    Log("ExHandler error: " & e.ExceptionObject.StackTrace)
    '    Me.Close()
    'End Sub

    'Private Sub MYThreadHandler(ByVal sender As Object, ByVal e As Threading.ThreadExceptionEventArgs)
    '    ' catch application error
    '    'MessageBox.Show("An error has occured." & vbCr & "See Log for more informations")
    '    Log("ThreadHandler error: " & e.Exception.StackTrace)
    '    Me.Close()
    'End Sub

    Public Sub Log(ByVal logMessage As String)
        Dim swt As StreamWriter
        If Not File.Exists(LogPath) Then
            swt = File.CreateText(LogPath)
        Else
            swt = File.AppendText(LogPath)
        End If
        swt.WriteLine(logMessage)
        swt.Flush()
        swt.Close()
    End Sub

    Private Sub LoadLicenseKey()
        'Load License Key
        Dim LicensePath As String = Application.StartupPath & "\LicenseKey.txt"
        If Not File.Exists(LicensePath) Then
            Dim swt As StreamWriter
            swt = File.CreateText(LicensePath)
            swt.WriteLine("myCompany:swdocmgr_general-00000-11111-22222-33333-44444-55555-.....")
            swt.Flush()
            swt.Close()
        End If
        Dim LicenseText As List(Of String) = File.ReadAllLines(LicensePath).ToList
        If LicenseText.Count = 0 OrElse LicenseText(0).Length < 80 Then
            MessageBox.Show("Enter Solidworks Document Manager License in file:" & vbCr & LicensePath, "Error")
            Process.Start(LicensePath)
            Exit Sub
        End If
        Dim LicenseKey As String = LicenseText(0)
        'initialize swdocumentmgr
        Dim swClassFact As SwDMClassFactory = CreateObject("SwDocumentMgr.SwDMClassFactory")
        swDocMgr = swClassFact.GetApplication(LicenseKey)
    End Sub

    'search for files in a given folder and its sub folders, with a given files filter
    Private Function SearchForFiles(ByVal RootFolder As String, ByVal FileFilters As List(Of String)) As List(Of String)
        Dim ReturnedData As New List(Of String)
        Dim myFiles As New List(Of String)
        Dim FolderStack As New Stack(Of String)
        FolderStack.Push(RootFolder)
        Do While FolderStack.Count > 0
            Dim ThisFolder As String = FolderStack.Pop
            Try
                For Each SubFolder As String In Directory.GetDirectories(ThisFolder)
                    FolderStack.Push(SubFolder)
                Next
                For Each FileFilter As String In FileFilters
                    ReturnedData.AddRange(Directory.GetFiles(ThisFolder, FileFilter))
                Next
            Catch ex As Exception
            End Try
        Loop
        Return ReturnedData
    End Function

    Private Function OpenDocument(ByVal FilePath As String) As SwDMDocument

        Dim swDocType As Long
        If Path.GetExtension(FilePath).Contains("SLDPRT") Then
            swDocType = SwDmDocumentType.swDmDocumentPart
        ElseIf Path.GetExtension(FilePath).Contains("SLDASM") Then
            swDocType = SwDmDocumentType.swDmDocumentAssembly
        Else
            Return Nothing
        End If

        Dim swDoc As SwDMDocument
        Try
            Dim RetVal As Long
            ' open solidworks files with document manager
            swDoc = swDocMgr.GetDocument(FilePath, swDocType, False, RetVal)

            Select Case RetVal
                Case SwDmDocumentOpenError.swDmDocumentOpenErrorNone
                    'Log("File " & FilePath & " opened successfully")
                Case SwDmDocumentOpenError.swDmDocumentOpenErrorFail
                    Log("File " & FilePath & " failed to open; reasons could be related to permissions or the file is in use by some other application or the file does not exist")

                Case SwDmDocumentOpenError.swDmDocumentOpenErrorNonSW
                    Log("Non-SOLIDWORKS file " & FilePath)

                Case SwDmDocumentOpenError.swDmDocumentOpenErrorFileNotFound
                    Log("File " & FilePath & " not found")

                Case SwDmDocumentOpenError.swDmDocumentOpenErrorFileReadOnly
                    Log("File " & FilePath & " is read only")

                Case SwDmDocumentOpenError.swDmDocumentOpenErrorNoLicense
                    Log("No SOLIDWORKS Document Manager API license")

                Case SwDmDocumentOpenError.swDmDocumentOpenErrorFutureVersion
                    Log("File " & FilePath & " was created in a version of SOLIDWORKS more recent than the version of Document manager attempting to open the file" & " - Also: The program MUST be compile for x64 CPU only")

            End Select
        Catch ex As Exception
            ' Display error message for unlisted errors
            Log("Document Manager API Error. Probable causes:" & vbCr &
                " - invalid license number" & vbCr &
                " - project not compiled for x64 CPU")
            Return Nothing
        End Try
        Return swDoc
    End Function

    Private Function PropTypeToString(ByVal LongType As Long) As String
        Select Case LongType
            Case SwDmCustomInfoType.swDmCustomInfoText
                Return "Text"
            Case SwDmCustomInfoType.swDmCustomInfoNumber
                Return "Number"
            Case SwDmCustomInfoType.swDmCustomInfoDate
                Return "Date"
            Case SwDmCustomInfoType.swDmCustomInfoYesOrNo
                Return "Yes/No"
            Case Else
                Return "Other"
        End Select
    End Function
    Private Function PropTypeToLong(ByVal StringType As String) As Long
        Select Case StringType
            Case "Text"
                Return SwDmCustomInfoType.swDmCustomInfoText
            Case "Number"
                Return SwDmCustomInfoType.swDmCustomInfoNumber
            Case "Date"
                Return SwDmCustomInfoType.swDmCustomInfoDate
            Case "Yes/No"
                Return SwDmCustomInfoType.swDmCustomInfoYesOrNo
            Case Else
                Return SwDmCustomInfoType.swDmCustomInfoUnknown
        End Select
    End Function


    Private Sub BtBrowse_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles BtBrowse.Click
        If FolderBrowserDialog1.ShowDialog() = DialogResult.OK Then
            TbFolder.Text = FolderBrowserDialog1.SelectedPath
        End If
    End Sub

    Private Sub CbLoadColumns_CheckedChanged(sender As Object, e As EventArgs) Handles CbLoadColumns.CheckedChanged
        My.Settings.LoadColumns = CbLoadColumns.Checked
    End Sub
    Private Sub CbIncludeSubFolders_CheckedChanged(sender As Object, e As EventArgs) Handles CbIncludeSubFolders.CheckedChanged
        My.Settings.IncludeSubFolder = CbIncludeSubFolders.Checked
    End Sub
    Private Sub CbLoadCustomProp_CheckedChanged(sender As Object, e As EventArgs) Handles CbLoadCustomProp.CheckedChanged
        My.Settings.LoadCustomProp = CbLoadCustomProp.Checked
    End Sub
    Private Sub CbLoadAllConfig_CheckedChanged(sender As Object, e As EventArgs) Handles CbLoadAllConfig.CheckedChanged
        My.Settings.LoadAllConfig = CbLoadAllConfig.Checked
        If CbLoadActiveConfig.Checked And CbLoadAllConfig.Checked Then CbLoadActiveConfig.Checked = False
    End Sub
    Private Sub CbLoadActiveConfig_CheckedChanged(sender As Object, e As EventArgs) Handles CbLoadActiveConfig.CheckedChanged
        My.Settings.LoadActiveConfig = CbLoadActiveConfig.Checked
        If CbLoadActiveConfig.Checked And CbLoadAllConfig.Checked Then CbLoadAllConfig.Checked = False
    End Sub

    Private Sub DGV1_CellPainting(ByVal sender As Object, ByVal e As DataGridViewCellPaintingEventArgs) Handles DGV1.CellPainting
        'Draw custom cell borders.
        e.Paint(e.CellBounds, DataGridViewPaintParts.All And Not DataGridViewPaintParts.ContentBackground)
        If e.ColumnIndex = 1 Then
            ControlPaint.DrawBorder(e.Graphics, e.CellBounds,
                                    DGV1.GridColor, 0, ButtonBorderStyle.None,
                                    DGV1.GridColor, 0, ButtonBorderStyle.None,
                                    DGV1.GridColor, 3, ButtonBorderStyle.Outset,
                                    DGV1.GridColor, 0, ButtonBorderStyle.None)
        End If
        e.Handled = True
    End Sub

    Private Sub DGV2_CellPainting(ByVal sender As Object, ByVal e As DataGridViewCellPaintingEventArgs) Handles DGV2.CellPainting
        'Draw custom cell borders.
        e.Paint(e.CellBounds, DataGridViewPaintParts.All And Not DataGridViewPaintParts.ContentBackground)
        If e.ColumnIndex = 1 Then
            ControlPaint.DrawBorder(e.Graphics, e.CellBounds,
                                    DGV2.GridColor, 0, ButtonBorderStyle.None,
                                    DGV2.GridColor, 0, ButtonBorderStyle.None,
                                    DGV2.GridColor, 3, ButtonBorderStyle.Outset,
                                    DGV2.GridColor, 0, ButtonBorderStyle.None)
        End If
        e.Handled = True
    End Sub


    Private Sub DGV2_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs) Handles DGV2.KeyDown
        If e.KeyCode = Keys.V And Keys.ControlKey Then PasteData(sender)
    End Sub

    'function to paste the clipboard into a given datagridview
    Public Sub PasteData(ByRef dgv As DataGridView)
        Dim ColNum, ColNumTemp, LineNum As Integer
        Dim myRows As List(Of String)
        Dim myRowItems As List(Of String)
        myRows = Clipboard.GetText.Trim.Split(vbCr).ToList

        'Get Paste location
        LineNum = dgv.CurrentCellAddress.Y()
        ColNum = dgv.CurrentCellAddress.X()

        ' Increase the number of rows as necessary (Disabled for this application)
        'If myRows.Count >= (dgv.Rows.Count - LineNum) Then dgv.Rows.Add(myRows.Count - dgv.Rows.Count + LineNum) ' + 1 to add empty row

        'Paste
        For Each myRow As String In myRows
            If Not String.IsNullOrEmpty(myRow) Then
                If LineNum > dgv.Rows.Count - 1 Then Exit Sub
                ColNumTemp = ColNum
                myRowItems = myRow.Split(vbTab).ToList
                For Each myRowItem As String In myRowItems
                    If ColNumTemp < 2 Then Continue For
                    If ColNumTemp > dgv.ColumnCount - 1 Then Exit For
                    dgv.Item(ColNumTemp, LineNum).Value = myRowItem.TrimStart
                    ColNumTemp = ColNumTemp + 1
                Next
                LineNum = LineNum + 1
            End If
        Next

    End Sub


    Private Sub DGV2_MouseUp(ByVal sender As Object, ByVal e As MouseEventArgs) Handles DGV2.MouseUp
        'Prevent first 2 columns from moving
        If dtDgv2.Rows.Count = 0 Then Exit Sub
        DGV2.Columns("FilePath").DisplayIndex = 0
        DGV2.Columns("FileName").DisplayIndex = 1
    End Sub
    Private Sub DGV1_Scroll(ByVal sender As Object, ByVal e As ScrollEventArgs) Handles DGV1.Scroll
        'lock scrolling of both DGV
        DGV2.HorizontalScrollingOffset = DGV1.HorizontalScrollingOffset
    End Sub
    Private Sub DGV2_ColumnWidthChanged(ByVal sender As Object, ByVal e As DataGridViewColumnEventArgs) Handles DGV2.ColumnWidthChanged
        'same column width resizing on both DGV
        DGV1.Columns(e.Column.Index).Width = e.Column.Width
    End Sub
    Private Sub DGV2_ColumnDisplayIndexChanged(ByVal sender As Object, ByVal e As DataGridViewColumnEventArgs) Handles DGV2.ColumnDisplayIndexChanged
        'same column order on both DGV
        If DGV1.Columns.Count <> DGV2.Columns.Count Then Exit Sub
        DGV1.Columns(e.Column.Index).DisplayIndex = e.Column.DisplayIndex
    End Sub

    Private Sub ReformatDGVs()
        For Each col As DataGridViewColumn In DGV2.Columns
            Dim index As Integer = col.Index

            'disable columns sort
            DGV1.Columns(index).SortMode = DataGridViewColumnSortMode.NotSortable

            'set width
            AutoSizeColumn(index)

            'rename column with custom prop name only
            Dim colTxt As String = col.HeaderText
            If Not colTxt.Contains(";") Then Continue For
            col.HeaderText = colTxt.Substring(0, colTxt.IndexOf(";"))
        Next
    End Sub

    Private Sub AutoSizeColumn(index As Integer)
        'set width
        If DGV1.Columns.Count <> DGV2.Columns.Count Then Exit Sub
        DGV1.Columns(index).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCellsExceptHeader
        DGV2.Columns(index).AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        Dim ColWidth As Integer = Math.Min(Math.Max(DGV1.Columns(index).Width, DGV2.Columns(index).Width), 120)
        DGV1.Columns(index).AutoSizeMode = DataGridViewAutoSizeColumnMode.NotSet
        DGV2.Columns(index).AutoSizeMode = DataGridViewAutoSizeColumnMode.NotSet
        '  DGV1.Columns(index).Width = ColWidth
        DGV2.Columns(index).Width = ColWidth
    End Sub

    Private Sub DGV1_CellValueChanged(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) Handles DGV1.CellValueChanged
        'autosize column in both DGV
        AutoSizeColumn(e.ColumnIndex)
    End Sub

    Private Sub DGV2_CellValueChanged(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) Handles DGV2.CellValueChanged
        'autosize column in both DGV
        AutoSizeColumn(e.ColumnIndex)

        If DontUpdate Then Exit Sub
        If e.ColumnIndex < 2 Then Exit Sub
        If e.RowIndex = -1 Then Exit Sub
        UpdateCustomProp(e.RowIndex, e.ColumnIndex)

    End Sub

    Private Sub UpdateCustomProp(ByVal RowIndex As Integer, ByVal ColumnIndex As Integer)

        Dim FilePath As String = dtDgv2.Rows(RowIndex).Item(0).ToString
        Dim CustPropName As String = dtDgv2.Columns(ColumnIndex).ColumnName
        CustPropName = CustPropName.Substring(0, CustPropName.IndexOf(";"))
        Dim CustPropType As String = dtDgv1.Rows(0).Item(ColumnIndex).ToString
        Dim ConfigName As String = dtDgv1.Rows(1).Item(ColumnIndex).ToString
        Dim CustPropValue As String = dtDgv2.Rows(RowIndex).Item(ColumnIndex).ToString

        Select Case CustPropType
            Case "Number"
                Dim Num As Integer
                If Not Integer.TryParse(CustPropValue, Num) Then
                    DGV2.Rows(RowIndex).Cells(ColumnIndex).Value = ""
                    Exit Sub
                End If

            Case "Date"
                Dim result As Date
                If Not DateTime.TryParse(CustPropValue, result) Then
                    DGV2.Rows(RowIndex).Cells(ColumnIndex).Value = ""
                    Exit Sub
                End If

            Case "Yes/No"
                If CustPropValue = "True" OrElse CustPropValue = "Y" OrElse CustPropValue = "T" OrElse CustPropValue = "1" Then
                    CustPropValue = "Yes"
                    DGV2.Rows(RowIndex).Cells(ColumnIndex).Value = "Yes"
                End If
                If CustPropValue = "False" OrElse CustPropValue = "N" OrElse CustPropValue = "F" OrElse CustPropValue = "0" Then
                    CustPropValue = "No"
                    DGV2.Rows(RowIndex).Cells(ColumnIndex).Value = "No"
                End If

        End Select


        Log("   updating value: " & CustPropName & ", of type: " & CustPropType & " , to value: " & CustPropValue &
            ", in file: " & FilePath & ", config: " & ConfigName)


        Dim swDoc As SwDMDocument = OpenDocument(FilePath)
        If swDoc Is Nothing Then Exit Sub

        If ConfigName = String.Empty Then
            swDoc.DeleteCustomProperty(CustPropName)
            If CustPropValue <> String.Empty Then
                swDoc.AddCustomProperty(CustPropName, PropTypeToLong(CustPropType), CustPropValue)
            End If
        Else
            Dim swConfig As SwDMConfiguration
            Dim swConfigMgr As SwDMConfigurationMgr = swDoc.ConfigurationManager

            Dim ConfigNames As String() = swConfigMgr.GetConfigurationNames
            If Not ConfigNames.Contains(ConfigName) Then
                DGV2.Rows(RowIndex).Cells(ColumnIndex).Value = ""
                swDoc.CloseDoc()
                Exit Sub
            End If

            If CbLoadActiveConfig.Checked AndAlso ConfigName <> swConfigMgr.GetActiveConfigurationName Then
                ' do not update this config because even if it exists in this document,
                ' user only requested the active configuration
                DGV2.Rows(RowIndex).Cells(ColumnIndex).Value = ""
                swDoc.CloseDoc()
                Exit Sub
            End If

            swConfig = swConfigMgr.GetConfigurationByName(ConfigName)
            swConfig.DeleteCustomProperty(CustPropName)
            If CustPropValue <> String.Empty Then
                swConfig.AddCustomProperty(CustPropName, PropTypeToLong(CustPropType), CustPropValue)
            End If
        End If

        swDoc.Save()
        swDoc.CloseDoc()
    End Sub

    Private Sub DGV1_CellMouseClick(ByVal sender As Object, ByVal e As DataGridViewCellMouseEventArgs) Handles DGV1.CellMouseClick
        If e.Button <> MouseButtons.Right Then Exit Sub
        If e.RowIndex = 2 AndAlso e.ColumnIndex > 1 Then
            Dim CustPropType As String = dtDgv1.Rows(0).Item(ColumnClickIndex).ToString

            'populate menu
            ContextMenuStrip2.Items.Clear()
            Dim WildCards As New List(Of String)

            If CustPropType = "Yes/No" Then
                WildCards.Add("Yes")
                WildCards.Add("No")
            End If

            If CustPropType = "Date" Then
                WildCards.Add("<Today>")
            End If

            If CustPropType = "Number" Then
                WildCards.Add("<#(1)>")
            End If

            For Each col As DataColumn In dtDgv2.Columns
                Dim index As Integer = dtDgv2.Columns.IndexOf(col)
                If index < 2 AndAlso CustPropType <> "Text" Then Continue For
                WildCards.Add("<" & col.ColumnName & ">")
            Next

            If CustPropType = "Text" Then
                WildCards.Add("""SW-Material""")
                WildCards.Add("""SW-Mass""")
                WildCards.Add("<Today>")
                WildCards.Add("<UserName>")
                WildCards.Add("<####(1)>")
            End If

            WildCards.Add("<Folder(-2)>")

            For Each WildCard As String In WildCards
                Dim SubMenuItem = New ToolStripMenuItem
                SubMenuItem.Text = WildCard
                ContextMenuStrip2.Items.Add(SubMenuItem)
            Next

            ' show menu
            ContextMenuStrip2.Show(Cursor.Position)
        End If
    End Sub

    Private Sub DGV2_CellMouseClick(ByVal sender As Object, ByVal e As DataGridViewCellMouseEventArgs) Handles DGV2.CellMouseClick
        If e.Button <> MouseButtons.Right Then Exit Sub

        If e.RowIndex = -1 AndAlso e.ColumnIndex > 1 Then
            ' show menu
            ContextMenuStrip1.Show(Cursor.Position)

            ' populate submenu with config
            Dim DocConfigs As New List(Of String)
            Dim ConfigName As String
            For Each col As DataColumn In dtDgv1.Columns
                If dtDgv1.Columns.IndexOf(col) < 2 Then Continue For
                ConfigName = dtDgv1.Rows(1).Item(col).ToString
                If ConfigName <> String.Empty AndAlso Not DocConfigs.Contains(ConfigName) Then DocConfigs.Add(ConfigName)
            Next
            If DocConfigs.Count > 0 Then DocConfigs.Insert(0, "-General-")

            If DocConfigs.Count > 0 Then
                For Each MenuInsertType As ToolStripMenuItem In MenuInsertColumn.DropDownItems
                    MenuInsertType.DropDownItems.Clear()
                    For Each Config As String In DocConfigs
                        Dim SubMenuItem = New ToolStripMenuItem
                        SubMenuItem.Text = Config
                        MenuInsertType.DropDownItems.Add(SubMenuItem)
                    Next
                Next
            End If
        End If
    End Sub

    Private ColumnClickIndex As Integer
    Private Sub DGV1_CellMouseEnter(ByVal sender As Object, ByVal location As DataGridViewCellEventArgs) Handles DGV1.CellMouseEnter
        ColumnClickIndex = location.ColumnIndex
    End Sub
    Private Sub DGV2_CellMouseEnter(ByVal sender As Object, ByVal location As DataGridViewCellEventArgs) Handles DGV2.CellMouseEnter
        ColumnClickIndex = location.ColumnIndex
    End Sub

    Private Sub ApplyRuleToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ApplyRuleToolStripMenuItem.Click
        Dim StoredIndex As Integer = ColumnClickIndex
        Dim ColRule As String = dtDgv1.Rows(2).Item(StoredIndex).ToString
        If ColRule = String.Empty Then Exit Sub
        For Each row As DataRow In dtDgv2.Rows
            Dim CellValue As String = ColRule
            'Wildcards
            ' <Folder(2)>
            CellValue = FolderWildCard(CellValue, row.Item(0).ToString)
            ' <FilePath> <FileName> <myCustomProp;Type;Config>
            CellValue = ColumnWildCard(CellValue, row)
            ' <#(1)> <##(33)>
            CellValue = CounterWildCard(CellValue, dtDgv2.Rows.IndexOf(row))
            ' <Today>
            CellValue = CellValue.Replace("<Today>", Date.Today)
            ' <UserName>
            CellValue = CellValue.Replace("<UserName>", Environment.UserName)

            row.Item(StoredIndex) = CellValue
            UpdateCustomProp(dtDgv2.Rows.IndexOf(row), StoredIndex)
        Next
    End Sub

    Private Function FolderWildCard(ByVal CellValue As String, FilePath As String) As String
        Dim Pos As Integer = CellValue.IndexOf("<Folder(")
        If Pos = -1 Then Return CellValue
        Dim Pos2 As Integer = CellValue.IndexOf(")>", Pos + 8)
        If Pos2 = -1 Then Return CellValue
        Dim IndexStr As String = CellValue.Substring(Pos + 8, Pos2 - Pos - 8)

        ' remove the wildcard string
        CellValue = CellValue.Substring(0, Pos) & CellValue.Substring(Pos2 + 2)

        ' recursive
        CellValue = FolderWildCard(CellValue, FilePath)

        Dim Index As Integer
        If Int32.TryParse(IndexStr, Index) = 0 Then Return CellValue

        Dim SubFolders As String() = FilePath.Split("\")
        If Math.Abs(Index) > SubFolders.Count Then Return CellValue

        Dim SubFolder As String
        If Index > 0 Then
            SubFolder = SubFolders(Index - 1)
        Else
            SubFolder = SubFolders(SubFolders.Count + Index)
        End If
        Return CellValue.Insert(Pos, SubFolder)

    End Function

    Private Function ColumnWildCard(ByVal CellValue As String, ByVal row As DataRow) As String
        For Each col As DataColumn In dtDgv2.Columns
            Dim index As Integer = dtDgv2.Columns.IndexOf(col)
            '  If index < 2 Then Continue For
            CellValue = CellValue.Replace("<" & col.ColumnName & ">", row.Item(index).ToString)
        Next
        Return CellValue
    End Function

    Private Function CounterWildCard(ByVal CellValue As String, ByVal Index As Integer) As String
        Dim Pos As Integer = CellValue.IndexOf("<#")
        If Pos = -1 Then Return CellValue
        Dim Pos2 As Integer = CellValue.IndexOf("#(", Pos)
        If Pos2 = -1 Then Return CellValue
        Dim Pos3 As Integer = CellValue.IndexOf(")>", Pos)
        If Pos3 = -1 Then Return CellValue
        Dim IndexStr As String = CellValue.Substring(Pos2 + 2, Pos3 - Pos2 - 2)

        ' remove the wildcard string
        CellValue = CellValue.Substring(0, Pos) & CellValue.Substring(Pos3 + 2)

        ' recursive
        CellValue = CounterWildCard(CellValue, Index)

        If Int32.TryParse(IndexStr + Index, Index) = 0 Then Return CellValue
        Return CellValue.Insert(Pos, Index.ToString(StrDup(Pos2 - Pos, "0")))
    End Function

    Private Sub MenuDeleteColumn_Click(sender As Object, e As EventArgs) Handles MenuDeleteColumn.Click
        ' remove column
        'app crash if columns are frozen 
        DGV1.Columns(1).Frozen = False
        DGV2.Columns(1).Frozen = False

        dtDgv1.Columns.RemoveAt(ColumnClickIndex)
        dtDgv2.Columns.RemoveAt(ColumnClickIndex)

        'app crash if columns are frozen 
        DGV1.Columns(1).Frozen = True
        DGV2.Columns(1).Frozen = True
    End Sub

    Private Sub MenuDeleteColAndCustProp_Click(sender As Object, e As EventArgs) Handles MenuDeleteColAndCustProp.Click
        Dim StoredIndex As Integer = ColumnClickIndex
        If MessageBox.Show("Are you sure you want to delete this Custom Property?", "Confirmation", MessageBoxButtons.YesNo) <> DialogResult.Yes Then Exit Sub

        'find custom prop to delete and part configuration where it is situated
        Dim CustPropName As String = DGV2.Columns(StoredIndex).HeaderText
        Dim ConfigName As String = dtDgv1.Rows(1).Item(StoredIndex).ToString

        'delete custom prop in every documents
        For Each row As DataRow In dtDgv2.Rows
            Dim swDoc As SwDMDocument = OpenDocument(row.Item(0).ToString)
            If swDoc Is Nothing Then Continue For

            If ConfigName = String.Empty Then
                swDoc.DeleteCustomProperty(CustPropName)
            Else
                Dim swConfig As SwDMConfiguration
                Dim swConfigMgr As SwDMConfigurationMgr = swDoc.ConfigurationManager
                swConfig = swConfigMgr.GetConfigurationByName(ConfigName)
                swConfig.DeleteCustomProperty(CustPropName)
            End If

            swDoc.Save()
            swDoc.CloseDoc()
        Next

        ' remove column
        dtDgv1.Columns.RemoveAt(StoredIndex)
        dtDgv2.Columns.RemoveAt(StoredIndex)
    End Sub


    Private Sub MenuInsertColumn_DropDownItemClicked(ByVal sender As Object, ByVal e As ToolStripItemClickedEventArgs) Handles MenuInsertColumn.DropDownItemClicked, MenuInsertTypeText.DropDownItemClicked, MenuInsertTypeNumber.DropDownItemClicked, MenuInsertTypeDate.DropDownItemClicked, MenuInsertTypeYesNo.DropDownItemClicked
        Dim StoredIndex As Integer = ColumnClickIndex
        Dim ColName As String = InputBox("Enter a Custom Property Name", "Custom Property Name")
        If ColName = "" Then Exit Sub

        Dim CustPropType As String
        Dim ConfigName As String = String.Empty
        If sender.Text = "Insert Column" Then
            CustPropType = e.ClickedItem.Text
        Else
            CustPropType = sender.Text
            If e.ClickedItem.Text <> "-General-" Then ConfigName = e.ClickedItem.Text
        End If

        'check if the name is already attributed to a column
        For Each col As DataColumn In dtDgv2.Columns
            Dim index As Integer = dtDgv2.Columns.IndexOf(col)
            '  If index < 2 Then Continue For
            If String.Compare(DGV2.Columns(index).HeaderText, ColName, True) = 0 AndAlso String.Compare(dtDgv1.Rows(1).Item(index).ToString, ConfigName) = 0 Then
                MessageBox.Show("Column Name is already taken")
                Exit Sub
            End If
        Next

        Dim CustPropHeader As String = String.Join(";", ColName, CustPropType, ConfigName)

        dtDgv1.Columns.Add(CustPropHeader, GetType(String))
        dtDgv1.Rows(0).Item(CustPropHeader) = CustPropType
        dtDgv1.Rows(1).Item(CustPropHeader) = ConfigName

        dtDgv2.Columns.Add(CustPropHeader, GetType(String))
        'move column where user clicked
        DGV2.Columns(dtDgv2.Columns.Count - 1).DisplayIndex = DGV2.Columns(StoredIndex).DisplayIndex
        ReformatDGVs()

    End Sub

    Private Sub ContextMenuStrip2_ItemClicked(ByVal sender As Object, ByVal e As ToolStripItemClickedEventArgs) Handles ContextMenuStrip2.ItemClicked
        dtDgv1.Rows(2).Item(ColumnClickIndex) = dtDgv1.Rows(2).Item(ColumnClickIndex).ToString & e.ClickedItem.Text
        ReformatDGVs()
    End Sub


End Class