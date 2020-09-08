Imports System.IO
Imports SolidWorks.Interop.swdocumentmgr
Imports SolidWorks.Interop.sldworks
Imports SolidWorks.Interop.swconst


Public Class BatchPropForm
    Private LogPath As String = Application.StartupPath & "\Log.txt"
    Private swDocMgr As SwDMApplication
    Private DontUpdate As Boolean
    Private DocConfigs As New List(Of String)
    Private swApp As SldWorks
    Private SlowMode As Boolean = True
    Private Title As String = "Batch Custom Properties V0.32"

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        ' ------------------handle to catch application crashes
        ' Get your application's application domain.
        Dim currentDomain As AppDomain = AppDomain.CurrentDomain
        ' Define a handler for unhandled exceptions.
        '  AddHandler currentDomain.UnhandledException, AddressOf MYExHandler
        ' Define a handler for unhandled exceptions for threads behind forms.
        '  AddHandler Application.ThreadException, AddressOf MYThreadHandler

        If File.Exists(LogPath) Then File.Delete(LogPath)

        DontUpdate = False

        'check if setting file is corrupted
        Try
            'Load last search folder
            TbFolder.Text = My.Settings.SearchFolder

        Catch ex As Exception
            'if config file is corrupted : reset config file
            Dim ConfigPath As String = Application.LocalUserAppDataPath
            If ConfigPath.Contains("BatchCustom") Then
                ConfigPath = ConfigPath.Substring(0, ConfigPath.IndexOf("BatchCustom") - 1)
                For Each SubFolder As String In Directory.GetDirectories(ConfigPath)
                    If SubFolder.Contains("BatchCustom") Then
                        For Each foundFile As String In My.Computer.FileSystem.GetFiles(SubFolder, FileIO.SearchOption.SearchAllSubDirectories, "*.config")
                            File.Delete(foundFile)
                        Next
                    End If
                Next
            End If
            If MessageBox.Show("Error in user settings file" & vbCr & "Restart?", "Restart?", MessageBoxButtons.YesNo) = DialogResult.Yes Then
                Application.Restart()
            End If
            Me.Close()
            Exit Sub
        End Try


        'Load search filter
        TbFilter.Text = My.Settings.SearchFilter

        If TbFilter.Text = String.Empty Then TbFilter.Text = "*.sldprt,*.sldasm"
        'desactivate context menu in tbfilter textbox
        TbFilter.ShortcutsEnabled = False

        'set check boxes
        ToolStripLoadColumns.Checked = My.Settings.LoadColumns
        ToolStripOnlyShowSavedCol.Checked = My.Settings.OnlyShowSavedCol
        ToolStripDeleteProp.Checked = My.Settings.DeletePropertiesNotInSaved
        ToolStripCreateMissingProp.Checked = My.Settings.CreateMissingProperties
        CbIncludeSubFolders.Checked = My.Settings.IncludeSubFolder
        ToolStripLoadCustomProp.Checked = My.Settings.LoadCustomProp
        ToolStripLoadAllConfig.Checked = My.Settings.LoadAllConfig
        ToolStripLoadActiveConfig.Checked = My.Settings.LoadActiveConfig
        ToolStripLoadCutList.Checked = My.Settings.LoadCutList

        'disable ToolStripOnlyShowSavedCol if ToolStripLoadColumns is not selected
        ToolStripOnlyShowSavedCol.Enabled = ToolStripLoadColumns.Checked
        ToolStripDeleteProp.Enabled = ToolStripLoadColumns.Checked
        ToolStripCreateMissingProp.Enabled = ToolStripLoadColumns.Checked

        'load license key
        LoadLicenseKey()
        SetTitle()

        log("Form loaded")
    End Sub

    Private Sub ResetDGVs()

        DGV2.Columns.Clear()

        Dim col As New DataGridViewTextBoxColumn
        col.HeaderText = "FilePath"
        col.Name = "FilePath"
        DGV2.Columns.Add(col)

        col = New DataGridViewTextBoxColumn
        col.HeaderText = "FileName"
        col.Name = "FileName"
        DGV2.Columns.Add(col)

        'set first 2 columns columns width
        DGV2.Columns(0).Width = 100
        DGV2.Columns(1).Width = 100

        'Freeze first 2 columns
        DGV2.Columns(1).Frozen = True

        'prevent first 2 columns' cells editing
        DGV2.Columns(0).ReadOnly = True
        DGV2.Columns(1).ReadOnly = True

        DGV2.ScrollBars = ScrollBars.Vertical


        '====================

        DGV1.Columns.Clear()

        col = New DataGridViewTextBoxColumn
        col.HeaderText = "FilePath"
        col.Name = "FilePath"
        DGV1.Columns.Add(col)

        col = New DataGridViewTextBoxColumn
        col.HeaderText = "FileName"
        col.Name = "FileName"
        DGV1.Columns.Add(col)

        DGV1.Rows.Add("Type")
        DGV1.Rows.Add("Config")
        DGV1.Rows.Add("Rules")

        'Freeze first 2 columns
        DGV1.Columns(1).Frozen = True

        'prevent first 2 columns' cells editing
        DGV1.Columns(0).ReadOnly = True
        DGV1.Columns(1).ReadOnly = True

        'prevent row's cells editing
        DGV1.Rows(0).ReadOnly = True
        DGV1.Rows(1).ReadOnly = True

        If ToolStripLoadColumns.Checked Then LoadColumns()

    End Sub

    Private Sub BtSearch_Click(sender As Object, e As EventArgs) Handles BtSearch.Click
        log("BtSearch_Click")
        DontUpdate = True

        If Not Directory.Exists(TbFolder.Text) AndAlso Not File.Exists(TbFolder.Text) Then
            MessageBox.Show("Can not resolve search path " & TbFolder.Text)
            DontUpdate = False
            Exit Sub
        End If


        Try
            'Save search folder
            My.Settings.SearchFolder = TbFolder.Text
            'Save search filter
            My.Settings.SearchFilter = TbFilter.Text

            My.Settings.Save()

        Catch ex As Exception
            log("Error saving settings")
        End Try


        Dim FilePaths As New List(Of String)
        If Directory.Exists(TbFolder.Text) Then
            'set files search filter based on radio button choice
            Dim myFilters As List(Of String) = TbFilter.Text.Split(",").ToList

            'search files
            FilePaths = SearchForFiles(TbFolder.Text, myFilters)
        Else
            If SlowMode Then
                Dim swModel As ModelDoc2 = OpenDocument2(TbFolder.Text)
                SearchForComponentFiles2(swModel, FilePaths)
                swApp.CommandInProgress = False
                swApp.Visible = True
            Else
                Dim swDoc As SwDMDocument21 = OpenDocument(TbFolder.Text)
                SearchForComponentFiles(swDoc, FilePaths)
            End If
        End If

        ResetDGVs()

        DocConfigs.Clear()
        DocConfigs.Add("@General")

        ProgressBar1.Value = 0
        ProgressBar1.Maximum = FilePaths.Count
        For Each FilePath As String In FilePaths
            'bypass Solidworks temp files
            If FilePath.Contains("~$") Then Continue For

            If SlowMode Then
                AddFilesToDt2(FilePath)
            Else
                AddFilesToDt(FilePath)
            End If
            ProgressBar1.PerformStep()
        Next

        If ToolStripLoadColumns.Checked AndAlso My.Settings.ColumnsNames IsNot Nothing AndAlso DGV2.Columns.Count > 2 Then
            For Each col As DataGridViewColumn In DGV2.Columns
                Dim index As Integer = DGV2.Columns.IndexOf(col)
                If index < 2 Then Continue For
                Dim CustPropHeader As String = col.Name
                If My.Settings.ColumnsNames.Contains(CustPropHeader) Then
                    ' create empty properties
                    If ToolStripCreateMissingProp.Checked Then
                        For Each row In DGV2.Rows
                            If row.Cells(index).Value = "" Then
                                If SlowMode Then
                                    UpdateCustomProp2(DGV2.Rows.IndexOf(row), index)
                                Else
                                    UpdateCustomProp(DGV2.Rows.IndexOf(row), index)
                                End If
                            End If
                        Next
                    End If
                Else
                    'delete unsaved properties
                    If ToolStripDeleteProp.Checked Then
                        For Each row In DGV2.Rows
                            row.Cells(index).Value = "<DEL>"
                            If SlowMode Then
                                UpdateCustomProp2(DGV2.Rows.IndexOf(row), index)
                            Else
                                UpdateCustomProp(DGV2.Rows.IndexOf(row), index)
                            End If
                        Next
                    End If
                End If
            Next
        End If


        ReformatDGVs()

        ProgressBar1.Value = 0
        BtSearch.Text = "Refresh"
        DontUpdate = False
    End Sub

    Private Sub AddFilesToDt(ByVal FilePath As String)
        log("   Add file: " & FilePath)

        Dim swDoc As SwDMDocument = OpenDocument(FilePath)
        If swDoc Is Nothing Then Exit Sub

        'add a new row with file path and name
        DGV2.Rows.Add(FilePath, Path.GetFileNameWithoutExtension(FilePath))

        'load general custom properties
        If ToolStripLoadCustomProp.Checked Then ReadCustomPropGeneral(swDoc)

        If Path.GetExtension(FilePath).ToUpper <> ".SLDDRW" AndAlso (ToolStripLoadAllConfig.Checked OrElse ToolStripLoadActiveConfig.Checked) Then
            Dim swConfig As SwDMConfiguration
            Dim swConfigMgr As SwDMConfigurationMgr = swDoc.ConfigurationManager
            If ToolStripLoadAllConfig.Checked Then
                'load all configurations custom properties
                Dim ConfigNames As String() = swConfigMgr.GetConfigurationNames
                For Each ConfigName As String In ConfigNames
                    swConfig = swConfigMgr.GetConfigurationByName(ConfigName)
                    ReadCustomPropConfig(swConfig)
                Next
            Else
                'load active configuration custom properties
                Dim ConfigName As String = swConfigMgr.GetActiveConfigurationName
                swConfig = swConfigMgr.GetConfigurationByName(ConfigName)
                ReadCustomPropConfig(swConfig)
            End If
        End If

        If Path.GetExtension(FilePath).ToUpper = ".SLDPRT" AndAlso ToolStripLoadCutList.Checked Then
            Dim swDocument13 As SwDMDocument13
            swDocument13 = swDoc
            Dim vCutListItems As Object
            vCutListItems = swDocument13.GetCutListItems2

            If vCutListItems IsNot Nothing Then
                Dim Cutlist As SwDMCutListItem2
                Dim CustPropNameArr As String()

                For Each vCutListItem In vCutListItems
                    Cutlist = vCutListItem
                    If Cutlist.Quantity = 0 Then Continue For

                    CustPropNameArr = Cutlist.GetCustomPropertyNames
                    If CustPropNameArr Is Nothing Then Continue For

                    DGV2.Rows.Add(FilePath, Path.GetFileNameWithoutExtension(FilePath) & "\" & Cutlist.Name)
                    ReadCustomPropCutList(Cutlist)
                Next
            End If
        End If

        swDoc.CloseDoc()
    End Sub

    Private Sub AddFilesToDt2(ByVal FilePath As String)
        log("   Add file2: " & FilePath)

        Dim swModel As ModelDoc2 = OpenDocument2(FilePath)
        If swModel Is Nothing Then Exit Sub

        'add a new row with file path and name
        DGV2.Rows.Add(FilePath, Path.GetFileNameWithoutExtension(FilePath))

        'load general custom properties
        If ToolStripLoadCustomProp.Checked Then ReadCustomPropConfig2(swModel, String.Empty)

        If Path.GetExtension(FilePath).ToUpper <> ".SLDDRW" AndAlso (ToolStripLoadAllConfig.Checked OrElse ToolStripLoadActiveConfig.Checked) Then
            Dim swConfig As SolidWorks.Interop.sldworks.Configuration
            If ToolStripLoadAllConfig.Checked Then
                'load all configurations custom properties
                Dim ConfigNames As String() = swModel.GetConfigurationNames
                For Each ConfigName As String In ConfigNames
                    ReadCustomPropConfig2(swModel, ConfigName)
                Next
            Else
                'load active configuration custom properties
                swConfig = swModel.GetActiveConfiguration
                ReadCustomPropConfig2(swModel, swConfig.Name)
            End If
        End If

        If Path.GetExtension(FilePath).ToUpper = ".SLDPRT" AndAlso ToolStripLoadCutList.Checked Then

            Dim swFeature As Feature = swModel.FirstFeature
            While Not swFeature Is Nothing
                Dim swSubFeature As Feature = swFeature.GetFirstSubFeature
                While swSubFeature IsNot Nothing
                    If swSubFeature.GetTypeName2 = "CutListFolder" Then
                        Dim swCutFolder As BodyFolder = swSubFeature.GetSpecificFeature2
                        If swCutFolder IsNot Nothing AndAlso swCutFolder.GetBodyCount > 0 Then
                            DGV2.Rows.Add(FilePath, Path.GetFileNameWithoutExtension(FilePath) & "\" & swSubFeature.Name)
                            ReadCustomPropCutList2(swSubFeature)
                        End If
                    End If
                    swSubFeature = swSubFeature.GetNextSubFeature
                End While
                swFeature = swFeature.GetNextFeature
            End While
        End If

        swApp.CloseDoc(swModel.GetTitle)
    End Sub



    Private Sub ReadCustomPropGeneral(ByVal swDoc As SwDMDocument5)
        'log("      ReadCustomPropGeneral")
        Dim CustPropNameArr As String() = swDoc.GetCustomPropertyNames
        If CustPropNameArr Is Nothing Then Exit Sub

        Dim CustPropType As Long
        Dim CustPropValue As String = String.Empty
        Dim CustPropEval As String = String.Empty

        For Each CustPropName As String In CustPropNameArr
            'CustPropValue = swDoc.GetCustomProperty(CustPropName, CustPropType)
            CustPropEval = swDoc.GetCustomPropertyValues(CustPropName, CustPropType, CustPropValue)
            AddCustomPropToDGV(CustPropName, PropTypeToString(CustPropType), "", CustPropValue, CustPropEval)
        Next
    End Sub


    Private Sub ReadCustomPropConfig(swConfig As SwDMConfiguration5)
        'log("      ReadCustomPropConfig")
        'save config name in list to populate menu
        Dim ConfigName As String = swConfig.Name
        If ConfigName <> String.Empty AndAlso Not DocConfigs.Contains(ConfigName) Then DocConfigs.Add(ConfigName)

        Dim CustPropNameArr As String() = swConfig.GetCustomPropertyNames
        If CustPropNameArr Is Nothing Then Exit Sub

        Dim CustPropType As Long
        Dim CustPropValue As String = String.Empty
        Dim CustPropEval As String = String.Empty
        For Each CustPropName As String In CustPropNameArr
            'CustPropValue = swConfig.GetCustomProperty(CustPropName, CustPropType)
            CustPropEval = swConfig.GetCustomPropertyValues(CustPropName, CustPropType, CustPropValue)
            AddCustomPropToDGV(CustPropName, PropTypeToString(CustPropType), swConfig.Name, CustPropValue, CustPropEval)
        Next
    End Sub


    Private Sub ReadCustomPropConfig2(ByVal swModel As ModelDoc2, ByVal ConfigName As String)
        'log("      ReadCustomPropConfig2")
        'save config name in list to populate menu
        If ConfigName <> String.Empty AndAlso Not DocConfigs.Contains(ConfigName) Then DocConfigs.Add(ConfigName)

        Dim cpm As CustomPropertyManager
        cpm = swModel.Extension.CustomPropertyManager(ConfigName)
        Dim CustPropNameArr As String() = cpm.GetNames
        If CustPropNameArr Is Nothing Then Exit Sub

        Dim CustPropType As Long
        Dim CustPropValue As String = String.Empty
        Dim CustPropEval As String = String.Empty
        For Each CustPropName As String In CustPropNameArr
            cpm.Get2(CustPropName, CustPropValue, CustPropEval)
            CustPropType = cpm.GetType2(CustPropName)
            AddCustomPropToDGV(CustPropName, PropTypeToString(CustPropType), ConfigName, CustPropValue, CustPropEval)
        Next
    End Sub


    Private Sub ReadCustomPropCutList(swCutListItem As SwDMCutListItem2)
        'log("      ReadCustomPropCutList")
        Dim ConfigName As String = "@CutList"
        DocConfigs.Add(ConfigName)

        Dim CustPropNameArr As String() = swCutListItem.GetCustomPropertyNames
        If CustPropNameArr Is Nothing Then Exit Sub
        Dim CustPropEval As String
        Dim CustPropType As Long
        Dim CustPropValue As String = String.Empty

        For Each CustPropName As String In CustPropNameArr
            CustPropEval = swCutListItem.GetCustomPropertyValue2(CustPropName, CustPropType, CustPropValue)
            AddCustomPropToDGV(CustPropName, PropTypeToString(CustPropType), ConfigName, CustPropValue, CustPropEval)
        Next
    End Sub

    Sub ReadCustomPropCutList2(swDocFeat As Feature)
        Dim ConfigName As String = "@CutList"
        DocConfigs.Add(ConfigName)

        Dim cpm As CustomPropertyManager
        Dim CustPropNames As String()
        Dim CustPropValue As String = String.Empty
        Dim CustPropEval As String = String.Empty
        Dim CustPropType As Long
        cpm = swDocFeat.CustomPropertyManager
        If cpm IsNot Nothing Then
            CustPropNames = cpm.GetNames
            If Not IsNothing(CustPropNames) Then
                For Each CustPropName As String In CustPropNames
                    cpm.Get2(CustPropName, CustPropValue, CustPropEval)
                    CustPropType = cpm.GetType2(CustPropName)
                    AddCustomPropToDGV(CustPropName, PropTypeToString(CustPropType), ConfigName, CustPropValue, CustPropEval)
                Next
            End If
        End If
    End Sub

    Private Sub AddCustomPropToDGV(ByVal CustPropName As String, ByVal CustPropType As String, ByVal ConfigName As String, ByVal CustPropValue As String, ByVal CustPropEval As String)
        'log("        AddCustomPropToDGV")
        Dim CustPropHeader As String = String.Join(";", CustPropName, CustPropType, ConfigName)
        '    Dim ColumnHeaders = DGV1.Columns.Cast(Of DataGridViewColumn)().Select(Function(column) column.HeaderText.ToLower)

        If Not DGV2.Columns.Contains(CustPropHeader) Then
            'log("add column " & CustPropHeader)


            Dim col As New DataGridViewTextBoxColumn
            col.HeaderText = CustPropName
            col.Name = CustPropHeader
            DGV2.Columns.Add(col)

            col = New DataGridViewTextBoxColumn
            col.HeaderText = CustPropName
            col.Name = CustPropHeader
            DGV1.Columns.Add(col)


            DGV1.Rows(0).Cells(CustPropHeader).Value = CustPropType
            DGV1.Rows(1).Cells(CustPropHeader).Value = ConfigName
        End If

        If CustPropValue <> "" Then
            CustPropValue = CustPropValue & "$$$" & CustPropEval
        Else
            CustPropValue = CustPropEval
        End If

        Dim row As DataGridViewRow
        row = DGV2.Rows(DGV2.Rows.Count - 1)
        Try
            row.Cells(CustPropHeader).Value = CustPropValue
        Catch ex As Exception
            log(" => exception during Dgv2 row writing")
        End Try

    End Sub
    Private Sub DGV2_CellFormating(ByVal sender As Object, ByVal e As DataGridViewCellFormattingEventArgs) Handles DGV2.CellFormatting

        'change cell text and tooltip for cut list item
        'before: cell text = CustomPropValue$$$CustomPropEval
        'after:  cell text = CustomPropValue and cell tooltip = CustomPropEval


        Dim CellText As String = e.Value
        If CellText Is Nothing Then Exit Sub
        Dim Pos As Integer = CellText.IndexOf("$$$")

        If Pos <> -1 Then
            e.Value = CellText.Substring(0, Pos)
            Dim cell As DataGridViewCell = DGV2.Rows(e.RowIndex).Cells(e.ColumnIndex)
            cell.ToolTipText = CellText.Substring(Pos + 3)
        End If

    End Sub

    Private Sub ToolStripSaveColumns_Click(sender As Object, e As EventArgs) Handles ToolStripSaveColumns.Click
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
            My.Settings.ColumnsNames.Add(col.Name)
            My.Settings.ColumnsTypes.Add(DGV1.Rows(0).Cells(index).Value)
            My.Settings.ColumnsConfigs.Add(DGV1.Rows(1).Cells(index).Value)
            My.Settings.ColumnsRules.Add(DGV1.Rows(2).Cells(index).Value)
        Next

        My.Settings.Save()
    End Sub

    Private Sub LoadColumns()
        If My.Settings.ColumnsNames Is Nothing Then Exit Sub

        'For Each colName As String In My.Settings.ColumnsNames
        For i = 0 To My.Settings.ColumnsNames.Count - 1
            Dim CustPropHeader As String = My.Settings.ColumnsNames(i)
            If DGV2.Columns.Contains(CustPropHeader) Then Continue For


            Dim col As New DataGridViewTextBoxColumn
            col.HeaderText = CustPropHeader.Split(";")(0)
            col.Name = CustPropHeader
            DGV2.Columns.Add(col)

            col = New DataGridViewTextBoxColumn
            col.HeaderText = CustPropHeader.Split(";")(0)
            col.Name = CustPropHeader
            DGV1.Columns.Add(col)

            DGV1.Rows(0).Cells(CustPropHeader).Value = My.Settings.ColumnsTypes(i)
            DGV1.Rows(1).Cells(CustPropHeader).Value = My.Settings.ColumnsConfigs(i)
            DGV1.Rows(2).Cells(CustPropHeader).Value = My.Settings.ColumnsRules(i)
        Next
    End Sub

    Private Sub MYExHandler(ByVal sender As Object, ByVal e As UnhandledExceptionEventArgs)
        ' catch application error
        MessageBox.Show("An Error has occured." & vbCr & "See Log For more informations")
        log("ExHandler Error: " & e.ExceptionObject.StackTrace)
        Me.Close()
    End Sub

    Private Sub MYThreadHandler(ByVal sender As Object, ByVal e As Threading.ThreadExceptionEventArgs)
        ' catch application error
        MessageBox.Show("An error has occured." & vbCr & "See Log for more informations")
        log("ThreadHandler error: " & e.Exception.StackTrace)
        Me.Close()
    End Sub

    Public Sub log(ByVal logMessage As String)
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
            Dim DiagResult As DialogResult = MessageBox.Show("Do you want to run the program in Fast mode?", "", MessageBoxButtons.YesNo)
            If DiagResult = DialogResult.Yes Then
                MessageBox.Show("Enter Solidworks Document Manager License in file:" & vbCr & LicensePath & vbCr & vbCr & "Then restart the program.")
                Process.Start(LicensePath)
                Exit Sub
            End If
        End If

        Dim LicenseText As List(Of String) = File.ReadAllLines(LicensePath).ToList
        If LicenseText.Count = 0 OrElse LicenseText(0).Length < 80 Then
            Exit Sub
        End If

        Dim LicenseKey As String = LicenseText(0)
        'initialize swdocumentmgr
        Dim swClassFact As SwDMClassFactory = CreateObject("SwDocumentMgr.SwDMClassFactory")
        If swClassFact Is Nothing Then MessageBox.Show("Cannot access Solidworks Document Manager", "Error")

        swDocMgr = swClassFact.GetApplication(LicenseKey)
        SlowMode = False
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
                For Each FileFilter As String In FileFilters
                    ReturnedData.AddRange(Directory.GetFiles(ThisFolder, FileFilter))
                Next
                If Not CbIncludeSubFolders.Checked Then Continue Do
                For Each SubFolder As String In Directory.GetDirectories(ThisFolder)
                    FolderStack.Push(SubFolder)
                Next
            Catch ex As Exception
            End Try
        Loop
        Return ReturnedData
    End Function

    Private Sub SearchForComponentFiles(ByRef swDoc As SwDMDocument, ByRef ReturnedData As List(Of String))

        Dim swConfigMgr As SwDMConfigurationMgr = swDoc.ConfigurationManager
        Dim swConfig As SwDMConfiguration2
        Dim ConfigNames As String() = swConfigMgr.GetConfigurationNames
        For Each ConfigName As String In ConfigNames
            swConfig = swConfigMgr.GetConfigurationByName(ConfigName)

            Dim vComps As Object
            vComps = swConfig.GetComponents
            Dim swComps As Array
            swComps = vComps
            For Each swComp As SwDMComponent6 In swComps
                If Not ReturnedData.Contains(swComp.PathName) Then
                    ReturnedData.Add(swComp.PathName)

                    If Path.GetExtension(swComp.PathName).ToUpper = ".SLDASM" Then
                        Dim swChildDoc As SwDMDocument10
                        Dim swSearchOpt As SwDMSearchOption
                        swSearchOpt = swDocMgr.GetSearchOptionObject
                        swSearchOpt.SearchFilters = SwDmSearchFilters.SwDmSearchExternalReference
                        swChildDoc = swComp.GetDocument2(True, swSearchOpt, Nothing)
                        SearchForComponentFiles(swChildDoc, ReturnedData)
                    End If
                End If
            Next
        Next
        swDoc.CloseDoc()
    End Sub

    Private Sub SearchForComponentFiles2(ByRef swModel As AssemblyDoc, ByRef ReturnedData As List(Of String))

        Dim swAssy As AssemblyDoc = swModel

        'Dim swConfig As Configuration
        'Dim ConfigNames As String() = swModel.GetConfigurationNames
        'Dim swConfigMgr As ConfigurationManager = swModel.ConfigurationManager
        'For Each ConfigName As String In ConfigNames
        '    swConfig = swConfigMgr.GetConfigurationByName(ConfigName)

        Dim vComps As Object = swAssy.GetComponents(False)
        Dim swComps As Array
        swComps = vComps
        For Each swComp As Component2 In swComps
            If Not ReturnedData.Contains(swComp.GetPathName) Then
                ReturnedData.Add(swComp.GetPathName)

                If Path.GetExtension(swComp.GetPathName).ToUpper = ".SLDASM" Then
                    Dim swChildDoc As ModelDoc2 = swComp.GetModelDoc2
                    SearchForComponentFiles2(swChildDoc, ReturnedData)
                End If
            End If
        Next
        ' Next
        swApp.CloseDoc(swModel.GetTitle)
    End Sub



    Private Function OpenDocument(ByVal FilePath As String) As SwDMDocument

        Dim swDocType As Long = GetDocType(FilePath)

        Dim swDoc As SwDMDocument
        Try
            Dim RetVal As Long
            ' open solidworks files with document manager
            swDoc = swDocMgr.GetDocument(FilePath, swDocType, False, RetVal)

            Select Case RetVal
                Case SwDmDocumentOpenError.swDmDocumentOpenErrorNone
                    'log("File " & FilePath & " opened successfully")
                Case SwDmDocumentOpenError.swDmDocumentOpenErrorFail
                    log("File " & FilePath & " failed to open; reasons could be related to permissions or the file is in use by some other application or the file does not exist")

                Case SwDmDocumentOpenError.swDmDocumentOpenErrorNonSW
                    log("Non-SOLIDWORKS file " & FilePath)

                Case SwDmDocumentOpenError.swDmDocumentOpenErrorFileNotFound
                    log("File " & FilePath & " not found")

                Case SwDmDocumentOpenError.swDmDocumentOpenErrorFileReadOnly
                    log("File " & FilePath & " is read only")

                Case SwDmDocumentOpenError.swDmDocumentOpenErrorNoLicense
                    log("No SOLIDWORKS Document Manager API license")

                Case SwDmDocumentOpenError.swDmDocumentOpenErrorFutureVersion
                    log("File " & FilePath & " was created in a version of SOLIDWORKS more recent than the version of Document manager attempting to open the file" & " - Also: The program MUST be compile for x64 CPU only")

            End Select
        Catch ex As Exception
            ' Display error message for unlisted errors
            log("Document Manager API Error. Probable causes:" & vbCrLf &
                " - invalid license number" & vbCrLf &
                " - project not compiled for x64 CPU")
            Return Nothing
        End Try
        Return swDoc
    End Function


    Private Function OpenDocument2(ByVal FilePath As String) As ModelDoc2

        Try
            swApp = GetObject(, "SldWorks.Application")
        Catch ex As Exception
            For Version = 30 To 13 Step -1
                Try
                    swApp = GetObject(, "SldWorks.Application." & Version)
                Catch ex2 As Exception
                End Try
                If swApp IsNot Nothing Then Exit For
            Next
        End Try

        If swApp Is Nothing Then
            Try
                swApp = CreateObject("SldWorks.Application")
            Catch ex As Exception
                log("error opening Solidworks " & ex.Message)
                Me.Close()
            End Try
        End If

        Dim swDocType As Long = GetDocType(FilePath)

        Dim swModel As ModelDoc2 = Nothing
        swApp.CommandInProgress = True
        swApp.Visible = False
        swApp.DocumentVisible(False, swDocType)
        'swApp.UserControl = False

        Try
            ' open solidworks files
            swModel = swApp.OpenDoc6(FilePath, swDocType, swOpenDocOptions_e.swOpenDocOptions_Silent, "", 0, 0)

        Catch ex As Exception
            ' Display error message for unlisted errors
            log("Can't open file:" & vbCrLf & FilePath)
            Return Nothing
        End Try

        If swModel IsNot Nothing Then
            '  swModel.ActiveView.EnableGraphicsUpdate = False
            swModel.Visible = False
        End If


        Return swModel
    End Function

    'Private Sub CloseDocument2(ByVal swModel As ModelDoc2)

    '    Dim swDocType As Long = GetDocType(swModel.GetPathName)

    '    swApp.CommandInProgress = False
    '    swApp.Visible = True
    '    swApp.DocumentVisible(True, swDocType)
    '    'swApp.UserControl = True
    '    If swModel IsNot Nothing Then
    '        '  swModel.ActiveView.EnableGraphicsUpdate = False
    '        swModel.Visible = True
    '    End If


    'End Sub

    Private Function GetDocType(ByVal FilePath As String) As Long
        Dim FileExt As String = Path.GetExtension(FilePath).ToUpper
        Select Case FileExt
            Case ".SLDPRT"
                Return 1 'SwDmDocumentType.swDmDocumentPart & swDocumentTypes_e.swDocPART
            Case ".SLDASM"
                Return 2 'SwDmDocumentType.swDmDocumentAssembly & swDocumentTypes_e.swDocASSEMBLY
            Case ".SLDDRW"
                Return 3 'SwDmDocumentType.swDmDocumentDrawing & swDocumentTypes_e.swDocDRAWING
            Case Else
                Return Nothing
        End Select
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

    Private Function PropTypeToLong2(ByVal StringType As String) As Long
        Select Case StringType
            Case "Text"
                Return swCustomInfoType_e.swCustomInfoText
            Case "Number"
                Return swCustomInfoType_e.swCustomInfoNumber
            Case "Date"
                Return swCustomInfoType_e.swCustomInfoDate
            Case "Yes/No"
                Return swCustomInfoType_e.swCustomInfoYesOrNo
            Case Else
                Return swCustomInfoType_e.swCustomInfoUnknown
        End Select
    End Function


    Private Sub BtBrowse_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles BtBrowseAssy.Click

        OpenFileDialog1.Filter = "Assembly Files|*.sldasm"
        If OpenFileDialog1.ShowDialog = DialogResult.OK Then TbFolder.Text = OpenFileDialog1.FileName

    End Sub


    Private Sub BtBrowseAssy_Click(sender As Object, e As EventArgs) Handles BtBrowseFolder.Click
        Using frm = New OpenFolderDialog()
            If frm.ShowDialog(Me) = DialogResult.OK Then TbFolder.Text = frm.Folder
        End Using
    End Sub


    Private Sub CbIncludeSubFolders_CheckedChanged(sender As Object, e As EventArgs) Handles CbIncludeSubFolders.CheckedChanged
        My.Settings.IncludeSubFolder = CbIncludeSubFolders.Checked
        My.Settings.Save()
    End Sub

    Private Sub ToolStripLoad_Click(sender As Object, e As EventArgs) Handles ToolStripLoadColumns.Click, ToolStripOnlyShowSavedCol.Click, ToolStripLoadCustomProp.Click, ToolStripLoadAllConfig.Click, ToolStripLoadActiveConfig.Click, ToolStripLoadCutList.Click, ToolStripDeleteProp.Click, ToolStripCreateMissingProp.Click
        'prevent menu from closing
        MenuStripOptions.Show()
    End Sub
    Private Sub ToolStripLoadColumns_CheckedChanged(sender As Object, e As EventArgs) Handles ToolStripLoadColumns.CheckedChanged
        My.Settings.LoadColumns = ToolStripLoadColumns.Checked
        My.Settings.Save()
        'disable ToolStripOnlyShowSavedCol if ToolStripLoadColumns is not selected
        ToolStripOnlyShowSavedCol.Enabled = ToolStripLoadColumns.Checked
        ToolStripDeleteProp.Enabled = ToolStripLoadColumns.Checked
        ToolStripCreateMissingProp.Enabled = ToolStripLoadColumns.Checked
    End Sub
    Private Sub ToolStripOnlyShowSavedCol_CheckedChanged(sender As Object, e As EventArgs) Handles ToolStripOnlyShowSavedCol.CheckedChanged
        My.Settings.OnlyShowSavedCol = ToolStripOnlyShowSavedCol.Checked
        My.Settings.Save()
        If ToolStripOnlyShowSavedCol.Checked Then ToolStripDeleteProp.Checked = False
    End Sub

    Private Sub ToolStripDeleteProp_CheckedChanged(sender As Object, e As EventArgs) Handles ToolStripDeleteProp.CheckedChanged
        My.Settings.DeletePropertiesNotInSaved = ToolStripDeleteProp.Checked
        My.Settings.Save()
        If ToolStripDeleteProp.Checked Then ToolStripOnlyShowSavedCol.Checked = False
    End Sub
    Private Sub ToolStripCreateMissingProp_CheckedChanged(sender As Object, e As EventArgs) Handles ToolStripCreateMissingProp.CheckedChanged
        My.Settings.CreateMissingProperties = ToolStripCreateMissingProp.Checked
        My.Settings.Save()
    End Sub

    Private Sub ToolStripLoadCustomProp_CheckedChanged(sender As Object, e As EventArgs) Handles ToolStripLoadCustomProp.CheckedChanged
        My.Settings.LoadCustomProp = ToolStripLoadCustomProp.Checked
        My.Settings.Save()
    End Sub
    Private Sub ToolStripLoadAllConfig_CheckedChanged(sender As Object, e As EventArgs) Handles ToolStripLoadAllConfig.CheckedChanged
        My.Settings.LoadAllConfig = ToolStripLoadAllConfig.Checked
        My.Settings.Save()
        If ToolStripLoadAllConfig.Checked And ToolStripLoadActiveConfig.Checked Then ToolStripLoadActiveConfig.Checked = False
    End Sub
    Private Sub ToolStripLoadActiveConfig_CheckedChanged(sender As Object, e As EventArgs) Handles ToolStripLoadActiveConfig.CheckedChanged
        My.Settings.LoadActiveConfig = ToolStripLoadActiveConfig.Checked
        My.Settings.Save()
        If ToolStripLoadAllConfig.Checked And ToolStripLoadActiveConfig.Checked Then ToolStripLoadAllConfig.Checked = False
    End Sub
    Private Sub ToolStripLoadCutList_CheckedChanged(sender As Object, e As EventArgs) Handles ToolStripLoadCutList.CheckedChanged
        My.Settings.LoadCutList = ToolStripLoadCutList.Checked
        My.Settings.Save()
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
        Dim ColNum, LineNum As Integer
        Dim myRows As List(Of String)
        Dim myRowItems As List(Of String)
        myRows = Clipboard.GetText.Trim.Split(vbCr).ToList

        'Get Paste location
        LineNum = dgv.CurrentCellAddress.Y()


        'get column display order

        Dim DisplayIndex As New List(Of Integer)
        Dim ColumnsInDisplayedOrder = DGV2.Columns.Cast(Of DataGridViewColumn)().OrderBy(Function(column) column.DisplayIndex)
        For Each col As DataGridViewColumn In ColumnsInDisplayedOrder
            If col.DisplayIndex < DGV2.Columns(ColumnClickIndex).DisplayIndex Then Continue For
            DisplayIndex.Add(col.Index)
        Next


        ' Increase the number of rows as necessary (Disabled for this application)
        'If myRows.Count >= (dgv.Rows.Count - LineNum) Then dgv.Rows.Add(myRows.Count - dgv.Rows.Count + LineNum) ' + 1 to add empty row

        'Paste
        For Each myRow As String In myRows
            If Not String.IsNullOrEmpty(myRow) Then
                If LineNum > dgv.Rows.Count - 1 Then Exit Sub
                ColNum = 0
                myRowItems = myRow.Split(vbTab).ToList
                For Each myRowItem As String In myRowItems
                    If ColNum > DisplayIndex.Count - 1 Then Exit For

                    If DisplayIndex(ColNum) > 1 Then
                        dgv.Item(DisplayIndex(ColNum), LineNum).Value = myRowItem.TrimStart
                    End If

                    ColNum += 1
                Next
                LineNum += 1
            End If
        Next

    End Sub

    Private Sub DGV2_MouseUp(ByVal sender As Object, ByVal e As MouseEventArgs) Handles DGV2.MouseUp
        'Prevent first 2 columns from moving
        If DGV2.Rows.Count = 0 Then Exit Sub
        DGV2.Columns(0).DisplayIndex = 0
        DGV2.Columns(1).DisplayIndex = 1
    End Sub
    Private Sub DGV1_Scroll(ByVal sender As Object, ByVal e As ScrollEventArgs) Handles DGV1.Scroll
        'lock scrolling of both DGV
        DGV2.HorizontalScrollingOffset = DGV1.HorizontalScrollingOffset
    End Sub
    Private Sub DGV2_ColumnWidthChanged(ByVal sender As Object, ByVal e As DataGridViewColumnEventArgs) Handles DGV2.ColumnWidthChanged
        'same column width resizing on both DGV
        DGV1.Columns(e.Column.Index).Width = e.Column.Width
        'lock scrolling of both DGV
        DGV2.HorizontalScrollingOffset = DGV1.HorizontalScrollingOffset
    End Sub
    Private Sub DGV2_ColumnDisplayIndexChanged(ByVal sender As Object, ByVal e As DataGridViewColumnEventArgs) Handles DGV2.ColumnDisplayIndexChanged
        'same column order on both DGV
        If DGV1.Columns.Count <> DGV2.Columns.Count Then Exit Sub
        DGV1.Columns(e.Column.Index).DisplayIndex = e.Column.DisplayIndex
        'lock scrolling of both DGV
        DGV2.HorizontalScrollingOffset = DGV1.HorizontalScrollingOffset
    End Sub

    Private Sub ReformatDGVs()
        For Each col As DataGridViewColumn In DGV2.Columns
            Dim index As Integer = col.Index

            'disable columns sort
            DGV1.Columns(index).SortMode = DataGridViewColumnSortMode.NotSortable

            'set width
            AutoSizeColumn(index)

            ' 'rename column with custom prop name only
            'If col.HeaderText.Contains(";") Then col.HeaderText = col.HeaderText.Split(";")(0)
        Next

        'only show columns that have been saved
        If ToolStripLoadColumns.Checked AndAlso My.Settings.ColumnsNames IsNot Nothing AndAlso DGV2.Columns.Count > 2 AndAlso (ToolStripOnlyShowSavedCol.Checked OrElse ToolStripDeleteProp.Checked) Then

            'For Each colName As String In My.Settings.ColumnsNames
            For Each col As DataGridViewColumn In DGV2.Columns
                Dim index As Integer = DGV2.Columns.IndexOf(col)
                If index < 2 Then Continue For
                Dim CustPropHeader As String = col.Name
                If My.Settings.ColumnsNames.Contains(CustPropHeader) Then
                    DGV2.Columns(CustPropHeader).Visible = True
                    DGV1.Columns(CustPropHeader).Visible = True
                Else

                    DGV2.Columns(CustPropHeader).Visible = False
                    DGV1.Columns(CustPropHeader).Visible = False
                End If
            Next

        End If
    End Sub

    Private Sub AutoSizeColumn(index As Integer)
        'lock scrolling of both DGV
        DGV2.HorizontalScrollingOffset = DGV1.HorizontalScrollingOffset

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

    Private StoredCellValue As String
    Private Sub DGV2_CellBeginEdit(ByVal sender As Object, ByVal e As DataGridViewCellCancelEventArgs) Handles DGV2.CellBeginEdit
        'store cell value
        StoredCellValue = DGV2.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
    End Sub

    Private Sub DGV2_CellValueChanged(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) Handles DGV2.CellValueChanged
        'autosize column in both DGV
        AutoSizeColumn(e.ColumnIndex)


        If DontUpdate Then Exit Sub
        If e.ColumnIndex < 2 Then Exit Sub
        If e.RowIndex = -1 Then Exit Sub


        'function will see cells are empty and delete the custom prop in every documents
        Dim DeleteSuccessful As Boolean
        If SlowMode Then
            DeleteSuccessful = UpdateCustomProp2(e.RowIndex, e.ColumnIndex)
        Else
            DeleteSuccessful = UpdateCustomProp(e.RowIndex, e.ColumnIndex)
        End If

        'if delete operation failed, re-write value in cell
        If DGV2.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = "<DEL>" AndAlso Not DeleteSuccessful Then
            DGV2.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = StoredCellValue
        End If

    End Sub

    Private Function UpdateCustomProp2(ByVal RowIndex As Integer, ByVal ColumnIndex As Integer) As Boolean

        ' log("UpdateCustomProp ")
        Dim FilePath As String = DGV2.Rows(RowIndex).Cells(0).Value

        Dim swModel As ModelDoc2 = OpenDocument2(FilePath)
        If swModel Is Nothing Then Return False
        Dim cpm As CustomPropertyManager

        Dim UpdateSuccessful As Boolean = False
        Dim DeleteSuccessful As Boolean = True


        Dim CustPropName As String = DGV2.Columns(ColumnIndex).HeaderText
        Dim CustPropType As String = DGV1.Rows(0).Cells(ColumnIndex).Value
        Dim ConfigName As String = DGV1.Rows(1).Cells(ColumnIndex).Value
        Dim CustPropValue As String = DGV2.Rows(RowIndex).Cells(ColumnIndex).Value

        log("   updating value: " & CustPropName & ", of type: " & CustPropType & " , to value: " & CustPropValue &
            ", in file: " & FilePath & ", config: " & ConfigName)

        'cpm = swModel.Extension.CustomPropertyManager("")
        'DeleteSuccessful = cpm.Delete(CustPropName)
        'If CustPropValue <> String.Empty Then UpdateSuccessful = cpm.Add2(CustPropName, PropTypeToLong2(CustPropType), CustPropValue)

        Dim FileName As String = DGV2.Rows(RowIndex).Cells(1).Value
        If ConfigName = String.Empty AndAlso Not FileName.Contains("\") Then
            'if normal custom prop
            cpm = swModel.Extension.CustomPropertyManager("")
            Dim CustPropNameArr As String() = cpm.GetNames

            If CustPropValue = "<DEL>" Then
                'if user entry is nothing , delete custom prop
                If CustPropNameArr IsNot Nothing AndAlso CustPropNameArr.Contains(CustPropName) Then
                    DeleteSuccessful = cpm.Delete(CustPropName)
                End If
            Else
                'check if user entry is correct
                CustPropValue = ParseCustPropValue(CustPropType, CustPropValue)
                If CustPropValue = "<incorrect>" Then
                    'parse function returned nothing
                    UpdateSuccessful = False
                Else
                    If CustPropNameArr IsNot Nothing AndAlso CustPropNameArr.Contains(CustPropName) Then
                        cpm.Set(CustPropName, CustPropValue)
                        UpdateSuccessful = True
                    Else
                        UpdateSuccessful = cpm.Add2(CustPropName, PropTypeToLong2(CustPropType), CustPropValue)
                    End If
                End If
            End If

        ElseIf Path.GetExtension(FilePath).ToUpper = ".SLDPRT" AndAlso ConfigName = "@CutList" AndAlso FileName.Contains("\") Then
            'if cut list item custom prop

            Dim CutListItemName As String = FileName.Split("\")(1)

            Dim swFeature As Feature = swModel.FirstFeature
            While Not swFeature Is Nothing
                Dim swSubFeature As Feature = swFeature.GetFirstSubFeature
                While swSubFeature IsNot Nothing
                    If swSubFeature.GetTypeName2 = "CutListFolder" Then
                        Dim swCutFolder As BodyFolder = swSubFeature.GetSpecificFeature2
                        If swCutFolder IsNot Nothing AndAlso swCutFolder.GetBodyCount > 0 AndAlso swSubFeature.Name = CutListItemName Then

                            Dim CustPropNameArr As String() = Nothing
                            Dim CustPropEval As String = String.Empty
                            cpm = swSubFeature.CustomPropertyManager

                            If cpm IsNot Nothing Then
                                CustPropNameArr = cpm.GetNames

                                If CustPropValue = "<DEL>" Then
                                    'if user entry is nothing , delete custom prop
                                    If CustPropNameArr IsNot Nothing AndAlso CustPropNameArr.Contains(CustPropName) Then
                                        DeleteSuccessful = cpm.Delete(CustPropName)
                                        '==some cut list properties can't be deleted because they are "linked" to a variable==
                                    End If
                                Else
                                    'check if user entry is correct
                                    CustPropValue = ParseCustPropValue(CustPropType, CustPropValue)
                                    If CustPropValue = "<incorrect>" Then
                                        'parse function returned nothing
                                        UpdateSuccessful = False
                                    Else
                                        If CustPropNameArr IsNot Nothing AndAlso CustPropNameArr.Contains(CustPropName) Then
                                            cpm.Set(CustPropName, CustPropValue)
                                            UpdateSuccessful = True
                                        Else
                                            UpdateSuccessful = cpm.Add2(CustPropName, PropTypeToLong(CustPropType), CustPropValue)
                                        End If
                                    End If
                                End If
                            End If

                        End If
                    End If
                    swSubFeature = swSubFeature.GetNextSubFeature
                End While
                swFeature = swFeature.GetNextFeature
            End While


        ElseIf Path.GetExtension(FilePath).ToUpper <> ".SLDDRW" AndAlso Not FileName.Contains("\") AndAlso Not ConfigName = "@CutList" Then
            'if configuration custom property
            UpdateSuccessful = True
            Dim ConfigNames As String() = swModel.GetConfigurationNames

            If Not ConfigNames.Contains(ConfigName) Then
                'config doesn't exist
                UpdateSuccessful = False
            End If

            Dim swConfig As SolidWorks.Interop.sldworks.Configuration = swModel.GetActiveConfiguration

            If ToolStripLoadActiveConfig.Checked AndAlso ConfigName <> swConfig.Name Then
                ' do not update this config custom property because even if it exists,
                ' the user only requested the active configuration
                UpdateSuccessful = False
            End If

            If ConfigName <> swConfig.Name Then
                swModel.ShowConfiguration2(ConfigName)
            End If

            cpm = swModel.Extension.CustomPropertyManager(ConfigName)

            If UpdateSuccessful AndAlso cpm IsNot Nothing Then
                UpdateSuccessful = False
                Dim CustPropNameArr As String() = cpm.GetNames

                If CustPropValue = "<DEL>" Then
                    'if user entry is nothing , delete custom prop
                    If CustPropNameArr IsNot Nothing AndAlso CustPropNameArr.Contains(CustPropName) Then
                        DeleteSuccessful = cpm.Delete(CustPropName)
                    End If
                Else
                    'check if user entry is correct
                    CustPropValue = ParseCustPropValue(CustPropType, CustPropValue)

                    If CustPropValue = "<incorrect>" Then
                        ' parse function returned nothing
                        UpdateSuccessful = False
                    Else
                        If CustPropNameArr IsNot Nothing AndAlso CustPropNameArr.Contains(CustPropName) Then
                            cpm.Set(CustPropName, CustPropValue)
                            UpdateSuccessful = True
                        Else
                            UpdateSuccessful = cpm.Add2(CustPropName, PropTypeToLong2(CustPropType), CustPropValue)
                        End If
                    End If
                End If
            End If

        End If


        If UpdateSuccessful OrElse DeleteSuccessful Then swModel.SaveAs2(FilePath, 0, False, False)

        If Not UpdateSuccessful AndAlso Not DeleteSuccessful Then
            DontUpdate = True
            DGV2.Rows(RowIndex).Cells(ColumnIndex).Value = ""
            DontUpdate = False
        End If


        swApp.CloseDoc(swModel.GetTitle)
        swApp.CommandInProgress = False
        swApp.Visible = True
        Return DeleteSuccessful

    End Function


    Private Function UpdateCustomProp(ByVal RowIndex As Integer, ByVal ColumnIndex As Integer) As Boolean
        ' log("UpdateCustomProp ")


        Dim FilePath As String = DGV2.Rows(RowIndex).Cells(0).Value
        Dim CustPropName As String = DGV2.Columns(ColumnIndex).HeaderText
        Dim CustPropType As String = DGV1.Rows(0).Cells(ColumnIndex).Value
        Dim ConfigName As String = DGV1.Rows(1).Cells(ColumnIndex).Value
        Dim CustPropValue As String = DGV2.Rows(RowIndex).Cells(ColumnIndex).Value


        Dim swDoc As SwDMDocument = OpenDocument(FilePath)
        If swDoc Is Nothing Then Return False

        log("   updating value: " & CustPropName & ", of type: " & CustPropType & " , to value: " & CustPropValue &
            ", in file: " & FilePath & ", config: " & ConfigName)

        Dim UpdateSuccessful As Boolean = False
        Dim DeleteSuccessful As Boolean = True
        Dim FileName As String = DGV2.Rows(RowIndex).Cells(1).Value
        If ConfigName = String.Empty AndAlso Not FileName.Contains("\") Then
            'if normal custom prop
            Dim CustPropNameArr As String() = swDoc.GetCustomPropertyNames

            If CustPropValue = "<DEL>" Then
                'if user entry is nothing , delete custom prop
                If CustPropNameArr IsNot Nothing AndAlso CustPropNameArr.Contains(CustPropName.ToLower, StringComparer.OrdinalIgnoreCase) Then
                    DeleteSuccessful = swDoc.DeleteCustomProperty(CustPropName)
                End If
            Else
                'check if user entry is correct
                CustPropValue = ParseCustPropValue(CustPropType, CustPropValue)
                If CustPropValue = "<incorrect>" Then
                    'parse function returned nothing
                    UpdateSuccessful = False
                Else
                    If CustPropNameArr IsNot Nothing AndAlso CustPropNameArr.Contains(CustPropName.ToLower, StringComparer.OrdinalIgnoreCase) Then
                        ' log("test 2 " & CustPropName)
                        swDoc.SetCustomProperty(CustPropName, CustPropValue)
                        UpdateSuccessful = True
                    Else
                        ' log("test 3 " & CustPropName)
                        UpdateSuccessful = swDoc.AddCustomProperty(CustPropName, PropTypeToLong(CustPropType), CustPropValue)
                    End If
                End If
            End If

        ElseIf Path.GetExtension(FilePath).ToUpper = ".SLDPRT" AndAlso ConfigName = "@CutList" AndAlso FileName.Contains("\") Then

            'if cut list item custom prop
            Dim swDocument13 As SwDMDocument13
            swDocument13 = swDoc
            Dim vCutListItems As Object
            vCutListItems = swDocument13.GetCutListItems2

            'find cutlistitem
            Dim CutListItemName As String = FileName.Split("\")(1)
            Dim swCutlistItem As SwDMCutListItem2 = Nothing
            For Each vCutListItem In vCutListItems
                swCutlistItem = vCutListItem
                If swCutlistItem.Name = CutListItemName Then Exit For
            Next

            Dim CustPropNameArr As String() = swCutlistItem.GetCustomPropertyNames

            If CustPropValue = "<DEL>" Then
                'if user entry is nothing , delete custom prop
                If CustPropNameArr IsNot Nothing AndAlso CustPropNameArr.Contains(CustPropName.ToLower, StringComparer.OrdinalIgnoreCase) Then
                    DeleteSuccessful = swCutlistItem.DeleteCustomProperty(CustPropName)
                    '==some cut list properties can't be deleted because they are "linked" to a variable==
                End If
            Else
                'check if user entry is correct
                CustPropValue = ParseCustPropValue(CustPropType, CustPropValue)
                If CustPropValue = "" Then
                    'parse function returned nothing
                    UpdateSuccessful = False
                Else
                    If CustPropNameArr IsNot Nothing AndAlso CustPropNameArr.Contains(CustPropName.ToLower, StringComparer.OrdinalIgnoreCase) Then
                        swCutlistItem.SetCustomProperty(CustPropName, CustPropValue)
                        UpdateSuccessful = True
                    Else
                        UpdateSuccessful = swCutlistItem.AddCustomProperty(CustPropName, PropTypeToLong(CustPropType), CustPropValue)
                    End If
                End If
            End If

        ElseIf Path.GetExtension(FilePath).ToUpper <> ".SLDDRW" AndAlso Not FileName.Contains("\") AndAlso Not ConfigName = "@CutList" Then
            UpdateSuccessful = True
            'if configuration custom property
            Dim swConfigMgr As SwDMConfigurationMgr = swDoc.ConfigurationManager
            Dim ConfigNames As String() = swConfigMgr.GetConfigurationNames

            If Not ConfigNames.Contains(ConfigName) Then
                'config doesn't exist
                UpdateSuccessful = False
            End If

            If ToolStripLoadActiveConfig.Checked AndAlso ConfigName <> swConfigMgr.GetActiveConfigurationName Then
                ' do not update this config custom property because even if it exists,
                ' the user only requested the active configuration
                UpdateSuccessful = False
            End If

            If UpdateSuccessful Then
                UpdateSuccessful = False
                Dim swConfig As SwDMConfiguration
                swConfig = swConfigMgr.GetConfigurationByName(ConfigName)
                Dim CustPropNameArr As String() = swConfig.GetCustomPropertyNames

                If CustPropValue = "<DEL>" Then
                    'if user entry is nothing , delete custom prop
                    If CustPropNameArr IsNot Nothing AndAlso CustPropNameArr.Contains(CustPropName.ToLower, StringComparer.OrdinalIgnoreCase) Then
                        DeleteSuccessful = swConfig.DeleteCustomProperty(CustPropName)
                    End If
                Else
                    'check if user entry is correct
                    CustPropValue = ParseCustPropValue(CustPropType, CustPropValue)

                    If CustPropValue = "<incorrect>" Then
                        ' parse function returned nothing
                        UpdateSuccessful = False
                    Else
                        If CustPropNameArr IsNot Nothing AndAlso CustPropNameArr.Contains(CustPropName.ToLower, StringComparer.OrdinalIgnoreCase) Then
                            swConfig.SetCustomProperty(CustPropName, CustPropValue)
                            UpdateSuccessful = True
                        Else
                            UpdateSuccessful = swConfig.AddCustomProperty(CustPropName, PropTypeToLong(CustPropType), CustPropValue)
                        End If
                    End If
                End If
            End If
        End If

        If UpdateSuccessful OrElse DeleteSuccessful Then swDoc.Save()

        If Not UpdateSuccessful AndAlso Not DeleteSuccessful Then
            DontUpdate = True
            DGV2.Rows(RowIndex).Cells(ColumnIndex).Value = ""
            DontUpdate = False
        End If

        swDoc.CloseDoc()
        Return DeleteSuccessful
    End Function


    Private Function ParseCustPropValue(ByVal CustPropType As String, ByVal CustPropValue As String) As String
        Select Case CustPropType
            Case "Number"
                Dim Num As Double
                If Not Double.TryParse(CustPropValue, Num) Then Return "<incorrect>"

            Case "Date"
                Dim result As Date
                If Not DateTime.TryParse(CustPropValue, result) Then Return "<incorrect>"

            Case "Yes/No"
                If CustPropValue = "True" OrElse CustPropValue = "Y" OrElse CustPropValue = "T" OrElse CustPropValue = "1" Then
                    Return "Yes"
                End If
                If CustPropValue = "False" OrElse CustPropValue = "N" OrElse CustPropValue = "F" OrElse CustPropValue = "0" Then
                    Return "No"
                End If
        End Select
        Return CustPropValue
    End Function

    Private ColumnClickIndex As Integer
    Private Sub DGV1_CellMouseClick(ByVal sender As Object, ByVal e As DataGridViewCellMouseEventArgs) Handles DGV1.CellMouseClick

        ColumnClickIndex = e.ColumnIndex

        If e.Button <> MouseButtons.Right Then Exit Sub
        If e.RowIndex = 2 AndAlso e.ColumnIndex > 1 Then
            Dim CustPropType As String = DGV1.Rows(0).Cells(ColumnClickIndex).Value

            'populate menu
            MenuStripRules.Items.Clear()
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

            For Each col As DataGridViewColumn In DGV2.Columns
                Dim index As Integer = DGV2.Columns.IndexOf(col)
                If index < 2 AndAlso CustPropType <> "Text" Then Continue For
                WildCards.Add("<" & col.Name & ">")
            Next

            If CustPropType = "Text" Then
                WildCards.Add("""SW-Material""")
                WildCards.Add("""SW-Mass""")
                WildCards.Add("<Today>")
                WildCards.Add("<UserName>")
                WildCards.Add("<####(1)>")
            End If

            WildCards.Add("<Folder(-2)>")
            WildCards.Add("<Replace(oldTxt,newTxt)>")

            For Each WildCard As String In WildCards
                Dim SubMenuItem = New ToolStripMenuItem
                SubMenuItem.Text = WildCard
                MenuStripRules.Items.Add(SubMenuItem)
            Next

            ' show menu
            MenuStripRules.Show(Cursor.Position)
        End If
    End Sub

    Private Sub DGV2_CellMouseClick(ByVal sender As Object, ByVal e As DataGridViewCellMouseEventArgs) Handles DGV2.CellMouseClick

        ColumnClickIndex = e.ColumnIndex

        If e.Button <> MouseButtons.Right Then Exit Sub

        If e.RowIndex = -1 AndAlso e.ColumnIndex < 2 Then
            MenuApplyRule.Enabled = False
            MenuDeleteColumn.Enabled = False
            MenuDeleteColAndCustProp.Enabled = False
        Else
            MenuApplyRule.Enabled = True
            MenuDeleteColumn.Enabled = True
            MenuDeleteColAndCustProp.Enabled = True
        End If

        If e.RowIndex = -1 AndAlso e.ColumnIndex > 0 Then
            ' show menu
            MenuStripColumn.Show(Cursor.Position)

            ' populate submenu with config
            If DocConfigs.Count > 1 Then
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

    Private Sub MenuApplyRule_Click(sender As Object, e As EventArgs) Handles MenuApplyRule.Click
        ProgressBar1.Value = 0
        ProgressBar1.Maximum = DGV2.RowCount
        ApplyRuleToColumn(ColumnClickIndex)
        ProgressBar1.Value = 0
    End Sub

    Private Sub MenuApplyAllRules_Click(sender As Object, e As EventArgs) Handles MenuApplyAllRules.Click
        ProgressBar1.Value = 0
        ProgressBar1.Maximum = (DGV2.ColumnCount - 2) * DGV2.RowCount
        For Each col As DataGridViewColumn In DGV2.Columns
            Dim index As Integer = DGV2.Columns.IndexOf(col)
            If index < 2 Then Continue For
            ApplyRuleToColumn(index)
        Next
        ProgressBar1.Value = 0
    End Sub

    Private Sub ApplyRuleToColumn(ByVal ColumnIndex As Integer)

        Dim ColRule As String = DGV1.Rows(2).Cells(ColumnIndex).Value
        If ColRule = String.Empty Then
            ProgressBar1.Value += DGV2.RowCount
            Exit Sub
        End If
        For Each row As DataGridViewRow In DGV2.Rows
            ProgressBar1.PerformStep()
            Dim CellValue As String = ColRule
            'Wildcards
            ' <Folder(2)>
            CellValue = FolderWildCard(CellValue, row.Cells(0).Value)
            ' <FilePath> <FileName> <myCustomProp;Type;Config>
            CellValue = ColumnWildCard(CellValue, row)
            ' <#(1)> <##(33)>
            CellValue = CounterWildCard(CellValue, DGV2.Rows.IndexOf(row))
            ' <Today>
            CellValue = CellValue.Replace("<Today>", Date.Today)
            ' <UserName>
            CellValue = CellValue.Replace("<UserName>", System.Environment.UserName)
            ' <Replace(g,h)>
            CellValue = ReplaceWildCard(CellValue, row, ColumnIndex)

            row.Cells(ColumnIndex).Value = CellValue
            If SlowMode Then
                UpdateCustomProp2(DGV2.Rows.IndexOf(row), ColumnIndex)
            Else
                UpdateCustomProp(DGV2.Rows.IndexOf(row), ColumnIndex)
            End If
        Next
        AutoSizeColumn(ColumnIndex)


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

    Private Function ColumnWildCard(ByVal CellValue As String, ByVal row As DataGridViewRow) As String

        For Each col As DataGridViewColumn In DGV2.Columns
            Dim index As Integer = DGV2.Columns.IndexOf(col)

            Dim CellText As String = row.Cells(index).Value
            If CellText Is Nothing Then CellText = ""

            Dim CellTextEval As String

            Dim Pos As Integer = CellText.IndexOf("$$$")
            If Pos <> -1 Then
                CellTextEval = CellText.Substring(Pos + 3)
                CellText = CellText.Substring(0, Pos)
            Else
                CellTextEval = CellText
            End If
            'replace evaluated column value
            CellValue = CellValue.Replace("<<" & col.Name & ">>", CellTextEval)
            'replace standard column value
            CellValue = CellValue.Replace("<" & col.Name & ">", CellText)
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

    Private Function ReplaceWildCard(ByVal CellValue As String, ByVal row As DataGridViewRow, ByVal index As Integer) As String

        Dim Pos As Integer = CellValue.IndexOf("<Replace(")
        If Pos = -1 Then Return CellValue
        Dim Pos2 As Integer = CellValue.IndexOf(")>", Pos + 9)
        If Pos2 = -1 Then Return CellValue
        Dim IndexStr As String = CellValue.Substring(Pos + 9, Pos2 - Pos - 9)
        Dim OldStr As String = IndexStr.Split(",")(0)
        Dim NewStr As String = IndexStr.Split(",")(1)
        ' Console.WriteLine("old " & OldStr)
        ' Console.WriteLine("new " & NewStr)

        ' remove the wildcard string
        ' CellValue = CellValue.Substring(0, Pos) & CellValue.Substring(Pos2 + 2)
        ' log("cell: " & CellValue)

        If IsDBNull(row.Cells(index).Value) Then Return ""
        Dim CellText As String = row.Cells(index).Value

        Return CellText.Replace(OldStr, NewStr)

    End Function

    Private Sub MenuDeleteColumn_Click(sender As Object, e As EventArgs) Handles MenuDeleteColumn.Click
        DeleteColumn(False)
    End Sub
    Private Sub MenuDeleteColAndCustProp_Click(sender As Object, e As EventArgs) Handles MenuDeleteColAndCustProp.Click
        DeleteColumn(True)
    End Sub
    Private Sub MenuDeleteColumnS_Click(sender As Object, e As EventArgs) Handles MenuDeleteColumnS.Click
        DeleteColumn(False, True)
    End Sub
    Private Sub MenuDeleteColAndCustPropS_Click(sender As Object, e As EventArgs) Handles MenuDeleteColAndCustPropS.Click
        DeleteColumn(True, True)
    End Sub
    Private Sub DeleteColumn(ByVal AlsoDeleteCustProp As Boolean, Optional ByVal AlsoDeleteColumnAfter As Boolean = False)
        Dim StoredScrollOffset As Integer = DGV1.HorizontalScrollingOffset
        'app crash if columns are frozen 
        DGV1.Columns(1).Frozen = False
        DGV2.Columns(1).Frozen = False

        Dim StoredIndex As Integer
        Dim ColumnsToDelete As New List(Of Integer)
        If AlsoDeleteColumnAfter Then
            StoredIndex = DGV2.Columns(ColumnClickIndex).DisplayIndex
            For Each col As DataGridViewColumn In DGV2.Columns
                If col.DisplayIndex >= StoredIndex Then
                    ColumnsToDelete.Insert(0, col.Index)
                End If
            Next
        Else
            ColumnsToDelete.Add(ColumnClickIndex)
        End If

        If AlsoDeleteCustProp Then
            If AlsoDeleteColumnAfter Then
                If MessageBox.Show("Are you sure you want to delete this Custom Property" & vbCr & "and also All the Others on the right?", "Confirmation", MessageBoxButtons.YesNo) <> DialogResult.Yes Then Exit Sub
            Else
                If MessageBox.Show("Are you sure you want to delete this Custom Property?", "Confirmation", MessageBoxButtons.YesNo) <> DialogResult.Yes Then Exit Sub
            End If
        End If

        For Each ColumnIndex In ColumnsToDelete
            Dim DeleteColumn As Boolean = True
            If AlsoDeleteCustProp Then
                For Each row As DataGridViewRow In DGV2.Rows
                    'store cell value
                    Dim CellValue As String = row.Cells(ColumnIndex).Value

                    'empty each column's cells
                    row.Cells(ColumnIndex).Value = "<DEL>"

                    'function will see cells are empty and delete the custom prop in every documents
                    Dim DeleteSuccessful As Boolean
                    If SlowMode Then
                        DeleteSuccessful = UpdateCustomProp2(DGV2.Rows.IndexOf(row), ColumnIndex)
                    Else
                        DeleteSuccessful = UpdateCustomProp(DGV2.Rows.IndexOf(row), ColumnIndex)
                    End If

                    'if failed
                    If Not DeleteSuccessful Then
                        row.Cells(ColumnIndex).Value = CellValue
                        DeleteColumn = False
                    End If
                Next
            End If

            ' remove column
            If DeleteColumn Then
                DGV2.Columns.RemoveAt(ColumnIndex)
                DGV1.Columns.RemoveAt(ColumnIndex)
            Else
                MessageBox.Show("The property: " & DGV2.Columns(ColumnIndex).Name & vbCr & " can't be deleted for some files")
            End If

        Next


        'app crash if columns are frozen 
        DGV1.Columns(1).Frozen = True
        DGV2.Columns(1).Frozen = True

        DGV1.HorizontalScrollingOffset = StoredScrollOffset
    End Sub


    Private Sub MenuInsertColumn_DropDownItemClicked(ByVal sender As Object, ByVal e As ToolStripItemClickedEventArgs) Handles MenuInsertColumn.DropDownItemClicked, MenuInsertTypeText.DropDownItemClicked, MenuInsertTypeNumber.DropDownItemClicked, MenuInsertTypeDate.DropDownItemClicked, MenuInsertTypeYesNo.DropDownItemClicked
        Dim StoredScrollOffset As Integer = DGV1.HorizontalScrollingOffset
        Dim StoredIndex As Integer = ColumnClickIndex
        Dim ColName As String = InputBox("Enter a Custom Property Name", "Custom Property Name")
        If ColName = "" Then Exit Sub

        Dim CustPropType As String
        Dim ConfigName As String = String.Empty
        If sender.Text = "Insert Column" Then
            CustPropType = e.ClickedItem.Text
        Else
            CustPropType = sender.Text
            If e.ClickedItem.Text <> "@General" Then ConfigName = e.ClickedItem.Text
        End If

        'check if the name is already attributed to a column
        For Each col2 As DataGridViewColumn In DGV2.Columns
            Dim index As Integer = DGV2.Columns.IndexOf(col2)
            '  If index < 2 Then Continue For
            If String.Compare(DGV2.Columns(index).Name, ColName, True) = 0 AndAlso String.Compare(DGV1.Rows(1).Cells(index).Value, ConfigName) = 0 Then
                MessageBox.Show("Column Name is already taken")
                Exit Sub
            End If
        Next

        Dim CustPropHeader As String = String.Join(";", ColName, CustPropType, ConfigName)


        Dim col As New DataGridViewTextBoxColumn
        col.HeaderText = ColName
        col.Name = CustPropHeader
        DGV2.Columns.Add(col)

        col = New DataGridViewTextBoxColumn
        col.HeaderText = ColName
        col.Name = CustPropHeader
        DGV1.Columns.Add(col)


        DGV1.Rows(0).Cells(CustPropHeader).Value = CustPropType
        DGV1.Rows(1).Cells(CustPropHeader).Value = ConfigName

        'find display index of the clicked column
        StoredIndex = DGV2.Columns(StoredIndex).DisplayIndex
        'if insert on the filename column insert after it
        If StoredIndex < 2 Then StoredIndex = 2
        'move column where user clicked
        DGV2.Columns(DGV2.Columns.Count - 1).DisplayIndex = StoredIndex

        'add to saved columns

        If My.Settings.ColumnsNames IsNot Nothing Then
            My.Settings.ColumnsNames.Add(CustPropHeader)
            My.Settings.ColumnsTypes.Add(CustPropType)
            My.Settings.ColumnsConfigs.Add(ConfigName)
            My.Settings.ColumnsRules.Add("")
            My.Settings.Save()
        End If

        ReformatDGVs()
        DGV1.HorizontalScrollingOffset = StoredScrollOffset
    End Sub

    Private Sub MenuStripRules_ItemClicked(ByVal sender As Object, ByVal e As ToolStripItemClickedEventArgs) Handles MenuStripRules.ItemClicked
        'populate rules cell with the text that the user clicked


        DGV1.Rows(2).Cells(ColumnClickIndex).Value = DGV1.Rows(2).Cells(ColumnClickIndex).Value & e.ClickedItem.Text
        ReformatDGVs()
        DGV2.HorizontalScrollingOffset = DGV1.HorizontalScrollingOffset
    End Sub

    Private Sub BtOptions_Click(sender As Object, e As EventArgs) Handles BtOptions.Click
        ' show menu
        MenuStripOptions.Show(Cursor.Position)
    End Sub

    Private Sub TbFilter_MouseUp(sender As Object, e As MouseEventArgs) Handles TbFilter.MouseUp
        If e.Button <> MouseButtons.Right Then Exit Sub
        ' show menu
        MenuStripFilter.Show(Cursor.Position)
    End Sub

    Private Sub MenuStripFilter_ItemClicked(ByVal sender As Object, ByVal e As ToolStripItemClickedEventArgs) Handles MenuStripFilter.ItemClicked
        If e.ClickedItem.Name = ClearToolStripMenuItem.Name Then
            TbFilter.Text = ""
        Else
            If TbFilter.Text = "" Then
                TbFilter.Text = e.ClickedItem.Text
            Else
                TbFilter.Text &= "," & e.ClickedItem.Text
            End If
        End If
    End Sub

    Private Sub ToolStripHelp_Click(sender As Object, e As EventArgs) Handles ToolStripHelp.Click
        Dim HelpForm As New HelpForm
        HelpForm.Show()
    End Sub

    Private Sub TbFolder_TextChanged(sender As Object, e As EventArgs) Handles TbFolder.TextChanged
        CbIncludeSubFolders.Enabled = (Directory.Exists(TbFolder.Text))
    End Sub

    Private Sub ToolStripSwitchMode_Click(sender As Object, e As EventArgs) Handles ToolStripSwitchMode.Click
        If SlowMode Then
            Dim LicensePath As String = Application.StartupPath & "\LicenseKey.txt"
            Dim LicenseText As List(Of String) = File.ReadAllLines(LicensePath).ToList
            If File.Exists(LicensePath) AndAlso (LicenseText.Count = 0 OrElse LicenseText(0).Length < 80) Then
                File.Delete(LicensePath)
            End If

            LoadLicenseKey()
        Else
            SlowMode = True
        End If
        SetTitle()
    End Sub

    Private Sub SetTitle()
        If SlowMode Then
            ToolStripSwitchMode.Text = "Switch to Fast Mode"
            Me.Text = Title & " - Slow Mode"
        Else
            ToolStripSwitchMode.Text = "Switch to Slow Mode"
            Me.Text = Title
        End If
    End Sub

    Private Sub TbHighlight_TextChanged(sender As Object, e As EventArgs) Handles TbHighlight.TextChanged
        DGV2.ClearSelection()
        'highlight rows
        For Each row As DataGridViewRow In DGV2.Rows
            If row.Cells(1).Value.ToString.ToLower.Contains(TbHighlight.Text.ToLower) Then
                row.DefaultCellStyle.BackColor = Color.LightGreen
            Else
                row.DefaultCellStyle.BackColor = Nothing
            End If
        Next

        'scroll to first row
        For Each row As DataGridViewRow In DGV2.Rows
            If row.Cells(1).Value.ToString.ToLower.Contains(TbHighlight.Text.ToLower) Then
                DGV2.FirstDisplayedScrollingRowIndex = row.Index
                Exit For
            End If
        Next
    End Sub

    Private Sub BtExport_Click(sender As Object, e As EventArgs) Handles BtExport.Click

        SaveFileDialog1.FileName = "CustomProp.txt"
        SaveFileDialog1.Filter = "Txt Files|*.txt"
        If SaveFileDialog1.ShowDialog <> Windows.Forms.DialogResult.OK Then Exit Sub
        Dim CsvFile As String = SaveFileDialog1.FileName


        Dim headers = (From header As DataGridViewColumn In DGV2.Columns.Cast(Of DataGridViewColumn)()
                       Select header.HeaderText).ToArray
        Dim rows = From row As DataGridViewRow In DGV2.Rows.Cast(Of DataGridViewRow)()
                   Where Not row.IsNewRow
                   Select Array.ConvertAll(row.Cells.Cast(Of DataGridViewCell).ToArray, Function(c) If(c.Value IsNot Nothing, c.Value.ToString, ""))
        Using sw As New IO.StreamWriter(CsvFile)
            sw.WriteLine(String.Join(vbTab, headers))
            For Each r In rows
                sw.WriteLine(String.Join(vbTab, r))
            Next
        End Using
    End Sub

    Private Sub BtImport_Click(sender As Object, e As EventArgs) Handles BtImport.Click
        Dim CsvFile As String = Nothing
        OpenFileDialog1.Filter = "Txt Files|*.txt"
        If OpenFileDialog1.ShowDialog = DialogResult.OK Then CsvFile = OpenFileDialog1.FileName
        If Not File.Exists(CsvFile) Then Exit Sub

        Dim Lines As List(Of String) = File.ReadAllLines(CsvFile).ToList

        Dim Headers As List(Of String) = (From header As DataGridViewColumn In DGV2.Columns.Cast(Of DataGridViewColumn)()
                                          Select header.HeaderText).ToList

        If Not Lines(0).Contains("FilePath") Then
            MessageBox.Show("Can not find the columns' headers in that file", "Error")
            Exit Sub
        End If

        Dim ColNames As List(Of String) = Lines(0).Split(vbTab).ToList
        Dim HeaderIndexes As New List(Of Integer)
        For Each Header In Headers
            If ColNames.Contains(Header) Then
                HeaderIndexes.Add(ColNames.IndexOf(Header))
            Else
                HeaderIndexes.Add(0)
            End If
        Next

        ProgressBar1.Value = 0
        ProgressBar1.Maximum = Lines.Count
        For Each Line In Lines.Skip(1)
            ProgressBar1.PerformStep()
            If Line.Length = 0 Then Continue For
            Dim Cells As String() = Line.Split(vbTab)

            For Each row As DataGridViewRow In DGV2.Rows
                If row.Cells(0).Value.ToString.ToLower <> Cells(HeaderIndexes(0)).ToLower Then Continue For

                For i = 2 To row.Cells.Count - 1
                    If Cells.Count - 1 < i OrElse HeaderIndexes(i) = 0 OrElse row.Cells(i).Value = Cells(HeaderIndexes(i)) Then Continue For
                    row.Cells(i).Value = Cells(HeaderIndexes(i))
                Next
            Next
        Next
        ProgressBar1.Value = 0

    End Sub


End Class

