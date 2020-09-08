<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class BatchPropForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.MenuStripColumn = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.MenuApplyRule = New System.Windows.Forms.ToolStripMenuItem()
        Me.MenuApplyAllRules = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator5 = New System.Windows.Forms.ToolStripSeparator()
        Me.MenuInsertColumn = New System.Windows.Forms.ToolStripMenuItem()
        Me.MenuInsertTypeText = New System.Windows.Forms.ToolStripMenuItem()
        Me.MenuInsertTypeNumber = New System.Windows.Forms.ToolStripMenuItem()
        Me.MenuInsertTypeDate = New System.Windows.Forms.ToolStripMenuItem()
        Me.MenuInsertTypeYesNo = New System.Windows.Forms.ToolStripMenuItem()
        Me.MenuDeleteColumn = New System.Windows.Forms.ToolStripMenuItem()
        Me.MenuDeleteColumnS = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.MenuDeleteColAndCustProp = New System.Windows.Forms.ToolStripMenuItem()
        Me.MenuDeleteColAndCustPropS = New System.Windows.Forms.ToolStripMenuItem()
        Me.CbIncludeSubFolders = New System.Windows.Forms.CheckBox()
        Me.BtSearch = New System.Windows.Forms.Button()
        Me.TbFolder = New System.Windows.Forms.TextBox()
        Me.BtBrowseAssy = New System.Windows.Forms.Button()
        Me.DGV1 = New System.Windows.Forms.DataGridView()
        Me.LabelFolder = New System.Windows.Forms.Label()
        Me.TbFilter = New System.Windows.Forms.TextBox()
        Me.DGV2 = New System.Windows.Forms.DataGridView()
        Me.MenuStripRules = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.BtOptions = New System.Windows.Forms.Button()
        Me.LabelSearch = New System.Windows.Forms.Label()
        Me.MenuStripOptions = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ToolStripSaveColumns = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripLoadColumns = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripOnlyShowSavedCol = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripDeleteProp = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripCreateMissingProp = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator3 = New System.Windows.Forms.ToolStripSeparator()
        Me.ToolStripLoadCustomProp = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripLoadAllConfig = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripLoadActiveConfig = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripLoadCutList = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator4 = New System.Windows.Forms.ToolStripSeparator()
        Me.ToolStripSwitchMode = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripHelp = New System.Windows.Forms.ToolStripMenuItem()
        Me.MenuStripFilter = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ClearToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.AddFilterToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SldprtToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SldasmToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SlddrwToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.BtBrowseFolder = New System.Windows.Forms.Button()
        Me.LabelHighlight = New System.Windows.Forms.Label()
        Me.TbHighlight = New System.Windows.Forms.TextBox()
        Me.BtExport = New System.Windows.Forms.Button()
        Me.BtImport = New System.Windows.Forms.Button()
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
        Me.MenuStripColumn.SuspendLayout()
        CType(Me.DGV1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DGV2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.MenuStripOptions.SuspendLayout()
        Me.MenuStripFilter.SuspendLayout()
        Me.SuspendLayout()
        '
        'MenuStripColumn
        '
        Me.MenuStripColumn.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MenuApplyRule, Me.MenuApplyAllRules, Me.ToolStripSeparator5, Me.MenuInsertColumn, Me.MenuDeleteColumn, Me.MenuDeleteColumnS, Me.ToolStripSeparator2, Me.MenuDeleteColAndCustProp, Me.MenuDeleteColAndCustPropS})
        Me.MenuStripColumn.Name = "ContextMenuStrip1"
        Me.MenuStripColumn.Size = New System.Drawing.Size(350, 170)
        '
        'MenuApplyRule
        '
        Me.MenuApplyRule.Name = "MenuApplyRule"
        Me.MenuApplyRule.Size = New System.Drawing.Size(349, 22)
        Me.MenuApplyRule.Text = "Apply Rule"
        '
        'MenuApplyAllRules
        '
        Me.MenuApplyAllRules.Name = "MenuApplyAllRules"
        Me.MenuApplyAllRules.Size = New System.Drawing.Size(349, 22)
        Me.MenuApplyAllRules.Text = "Apply All Rules"
        '
        'ToolStripSeparator5
        '
        Me.ToolStripSeparator5.Name = "ToolStripSeparator5"
        Me.ToolStripSeparator5.Size = New System.Drawing.Size(346, 6)
        '
        'MenuInsertColumn
        '
        Me.MenuInsertColumn.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MenuInsertTypeText, Me.MenuInsertTypeNumber, Me.MenuInsertTypeDate, Me.MenuInsertTypeYesNo})
        Me.MenuInsertColumn.Name = "MenuInsertColumn"
        Me.MenuInsertColumn.Size = New System.Drawing.Size(349, 22)
        Me.MenuInsertColumn.Text = "Insert Column"
        '
        'MenuInsertTypeText
        '
        Me.MenuInsertTypeText.Name = "MenuInsertTypeText"
        Me.MenuInsertTypeText.Size = New System.Drawing.Size(118, 22)
        Me.MenuInsertTypeText.Text = "Text"
        '
        'MenuInsertTypeNumber
        '
        Me.MenuInsertTypeNumber.Name = "MenuInsertTypeNumber"
        Me.MenuInsertTypeNumber.Size = New System.Drawing.Size(118, 22)
        Me.MenuInsertTypeNumber.Text = "Number"
        '
        'MenuInsertTypeDate
        '
        Me.MenuInsertTypeDate.Name = "MenuInsertTypeDate"
        Me.MenuInsertTypeDate.Size = New System.Drawing.Size(118, 22)
        Me.MenuInsertTypeDate.Text = "Date"
        '
        'MenuInsertTypeYesNo
        '
        Me.MenuInsertTypeYesNo.Name = "MenuInsertTypeYesNo"
        Me.MenuInsertTypeYesNo.Size = New System.Drawing.Size(118, 22)
        Me.MenuInsertTypeYesNo.Text = "Yes/No"
        '
        'MenuDeleteColumn
        '
        Me.MenuDeleteColumn.Name = "MenuDeleteColumn"
        Me.MenuDeleteColumn.Size = New System.Drawing.Size(349, 22)
        Me.MenuDeleteColumn.Text = "Delete Column"
        '
        'MenuDeleteColumnS
        '
        Me.MenuDeleteColumnS.Name = "MenuDeleteColumnS"
        Me.MenuDeleteColumnS.Size = New System.Drawing.Size(349, 22)
        Me.MenuDeleteColumnS.Text = "Delete Column and the Other Columns on the Right"
        '
        'ToolStripSeparator2
        '
        Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
        Me.ToolStripSeparator2.Size = New System.Drawing.Size(346, 6)
        '
        'MenuDeleteColAndCustProp
        '
        Me.MenuDeleteColAndCustProp.Name = "MenuDeleteColAndCustProp"
        Me.MenuDeleteColAndCustProp.Size = New System.Drawing.Size(349, 22)
        Me.MenuDeleteColAndCustProp.Text = "Delete Column and Custom Prop"
        '
        'MenuDeleteColAndCustPropS
        '
        Me.MenuDeleteColAndCustPropS.Name = "MenuDeleteColAndCustPropS"
        Me.MenuDeleteColAndCustPropS.Size = New System.Drawing.Size(349, 22)
        Me.MenuDeleteColAndCustPropS.Text = "Delete Column, the ones after, And All Custom Prop"
        '
        'CbIncludeSubFolders
        '
        Me.CbIncludeSubFolders.AutoSize = True
        Me.CbIncludeSubFolders.Location = New System.Drawing.Point(579, 34)
        Me.CbIncludeSubFolders.Name = "CbIncludeSubFolders"
        Me.CbIncludeSubFolders.Size = New System.Drawing.Size(120, 17)
        Me.CbIncludeSubFolders.TabIndex = 59
        Me.CbIncludeSubFolders.Text = "Include Sub Folders"
        Me.CbIncludeSubFolders.UseVisualStyleBackColor = True
        '
        'BtSearch
        '
        Me.BtSearch.Location = New System.Drawing.Point(727, 4)
        Me.BtSearch.Name = "BtSearch"
        Me.BtSearch.Size = New System.Drawing.Size(70, 37)
        Me.BtSearch.TabIndex = 54
        Me.BtSearch.Text = "Search"
        Me.BtSearch.UseVisualStyleBackColor = True
        '
        'TbFolder
        '
        Me.TbFolder.Location = New System.Drawing.Point(271, 5)
        Me.TbFolder.Name = "TbFolder"
        Me.TbFolder.Size = New System.Drawing.Size(426, 20)
        Me.TbFolder.TabIndex = 53
        '
        'BtBrowseAssy
        '
        Me.BtBrowseAssy.Location = New System.Drawing.Point(271, 30)
        Me.BtBrowseAssy.Name = "BtBrowseAssy"
        Me.BtBrowseAssy.Size = New System.Drawing.Size(144, 23)
        Me.BtBrowseAssy.TabIndex = 49
        Me.BtBrowseAssy.Text = "Browse for Assembly File..."
        Me.BtBrowseAssy.UseVisualStyleBackColor = True
        '
        'DGV1
        '
        Me.DGV1.AllowUserToAddRows = False
        Me.DGV1.AllowUserToDeleteRows = False
        Me.DGV1.AllowUserToOrderColumns = True
        Me.DGV1.AllowUserToResizeColumns = False
        Me.DGV1.AllowUserToResizeRows = False
        Me.DGV1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DGV1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DGV1.ColumnHeadersVisible = False
        Me.DGV1.Location = New System.Drawing.Point(1, 84)
        Me.DGV1.Name = "DGV1"
        Me.DGV1.RowHeadersVisible = False
        Me.DGV1.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DGV1.Size = New System.Drawing.Size(884, 85)
        Me.DGV1.TabIndex = 52
        '
        'LabelFolder
        '
        Me.LabelFolder.AutoSize = True
        Me.LabelFolder.Location = New System.Drawing.Point(249, 9)
        Me.LabelFolder.Name = "LabelFolder"
        Me.LabelFolder.Size = New System.Drawing.Size(18, 13)
        Me.LabelFolder.TabIndex = 51
        Me.LabelFolder.Text = "in:"
        '
        'TbFilter
        '
        Me.TbFilter.Location = New System.Drawing.Point(69, 5)
        Me.TbFilter.Name = "TbFilter"
        Me.TbFilter.Size = New System.Drawing.Size(158, 20)
        Me.TbFilter.TabIndex = 50
        '
        'DGV2
        '
        Me.DGV2.AllowUserToAddRows = False
        Me.DGV2.AllowUserToOrderColumns = True
        Me.DGV2.AllowUserToResizeRows = False
        Me.DGV2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DGV2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DGV2.Location = New System.Drawing.Point(1, 168)
        Me.DGV2.Name = "DGV2"
        Me.DGV2.RowHeadersVisible = False
        Me.DGV2.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DGV2.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DGV2.Size = New System.Drawing.Size(884, 300)
        Me.DGV2.TabIndex = 61
        '
        'MenuStripRules
        '
        Me.MenuStripRules.Name = "ContextMenuStrip2"
        Me.MenuStripRules.Size = New System.Drawing.Size(61, 4)
        '
        'BtOptions
        '
        Me.BtOptions.Location = New System.Drawing.Point(807, 4)
        Me.BtOptions.Name = "BtOptions"
        Me.BtOptions.Size = New System.Drawing.Size(70, 37)
        Me.BtOptions.TabIndex = 69
        Me.BtOptions.Text = "Options"
        Me.BtOptions.UseVisualStyleBackColor = True
        '
        'LabelSearch
        '
        Me.LabelSearch.AutoSize = True
        Me.LabelSearch.Location = New System.Drawing.Point(6, 9)
        Me.LabelSearch.Name = "LabelSearch"
        Me.LabelSearch.Size = New System.Drawing.Size(59, 13)
        Me.LabelSearch.TabIndex = 70
        Me.LabelSearch.Text = "Search for:"
        '
        'MenuStripOptions
        '
        Me.MenuStripOptions.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripSaveColumns, Me.ToolStripLoadColumns, Me.ToolStripOnlyShowSavedCol, Me.ToolStripDeleteProp, Me.ToolStripCreateMissingProp, Me.ToolStripSeparator3, Me.ToolStripLoadCustomProp, Me.ToolStripLoadAllConfig, Me.ToolStripLoadActiveConfig, Me.ToolStripLoadCutList, Me.ToolStripSeparator4, Me.ToolStripSwitchMode, Me.ToolStripHelp})
        Me.MenuStripOptions.Name = "ContextMenuStrip2"
        Me.MenuStripOptions.Size = New System.Drawing.Size(328, 258)
        '
        'ToolStripSaveColumns
        '
        Me.ToolStripSaveColumns.Name = "ToolStripSaveColumns"
        Me.ToolStripSaveColumns.Size = New System.Drawing.Size(327, 22)
        Me.ToolStripSaveColumns.Text = "Save Current Columns"
        '
        'ToolStripLoadColumns
        '
        Me.ToolStripLoadColumns.CheckOnClick = True
        Me.ToolStripLoadColumns.Name = "ToolStripLoadColumns"
        Me.ToolStripLoadColumns.Size = New System.Drawing.Size(327, 22)
        Me.ToolStripLoadColumns.Text = "Load Saved Columns On Refresh"
        '
        'ToolStripOnlyShowSavedCol
        '
        Me.ToolStripOnlyShowSavedCol.CheckOnClick = True
        Me.ToolStripOnlyShowSavedCol.Name = "ToolStripOnlyShowSavedCol"
        Me.ToolStripOnlyShowSavedCol.Size = New System.Drawing.Size(327, 22)
        Me.ToolStripOnlyShowSavedCol.Text = "Only Show Saved Columns"
        '
        'ToolStripDeleteProp
        '
        Me.ToolStripDeleteProp.CheckOnClick = True
        Me.ToolStripDeleteProp.Name = "ToolStripDeleteProp"
        Me.ToolStripDeleteProp.Size = New System.Drawing.Size(327, 22)
        Me.ToolStripDeleteProp.Text = "Delete Properties that are Not in Saved Columns"
        '
        'ToolStripCreateMissingProp
        '
        Me.ToolStripCreateMissingProp.CheckOnClick = True
        Me.ToolStripCreateMissingProp.Name = "ToolStripCreateMissingProp"
        Me.ToolStripCreateMissingProp.Size = New System.Drawing.Size(327, 22)
        Me.ToolStripCreateMissingProp.Text = "Create Missing Properties"
        '
        'ToolStripSeparator3
        '
        Me.ToolStripSeparator3.Name = "ToolStripSeparator3"
        Me.ToolStripSeparator3.Size = New System.Drawing.Size(324, 6)
        '
        'ToolStripLoadCustomProp
        '
        Me.ToolStripLoadCustomProp.CheckOnClick = True
        Me.ToolStripLoadCustomProp.Name = "ToolStripLoadCustomProp"
        Me.ToolStripLoadCustomProp.Size = New System.Drawing.Size(327, 22)
        Me.ToolStripLoadCustomProp.Text = "Load General Custom Prop"
        '
        'ToolStripLoadAllConfig
        '
        Me.ToolStripLoadAllConfig.CheckOnClick = True
        Me.ToolStripLoadAllConfig.Name = "ToolStripLoadAllConfig"
        Me.ToolStripLoadAllConfig.Size = New System.Drawing.Size(327, 22)
        Me.ToolStripLoadAllConfig.Text = "Load All Configs Custom Prop"
        '
        'ToolStripLoadActiveConfig
        '
        Me.ToolStripLoadActiveConfig.CheckOnClick = True
        Me.ToolStripLoadActiveConfig.Name = "ToolStripLoadActiveConfig"
        Me.ToolStripLoadActiveConfig.Size = New System.Drawing.Size(327, 22)
        Me.ToolStripLoadActiveConfig.Text = "Load Active Config Custom Prop"
        '
        'ToolStripLoadCutList
        '
        Me.ToolStripLoadCutList.CheckOnClick = True
        Me.ToolStripLoadCutList.Name = "ToolStripLoadCutList"
        Me.ToolStripLoadCutList.Size = New System.Drawing.Size(327, 22)
        Me.ToolStripLoadCutList.Text = "Load Cut List Custom Prop"
        '
        'ToolStripSeparator4
        '
        Me.ToolStripSeparator4.Name = "ToolStripSeparator4"
        Me.ToolStripSeparator4.Size = New System.Drawing.Size(324, 6)
        '
        'ToolStripSwitchMode
        '
        Me.ToolStripSwitchMode.Name = "ToolStripSwitchMode"
        Me.ToolStripSwitchMode.Size = New System.Drawing.Size(327, 22)
        Me.ToolStripSwitchMode.Text = "Switch to Fast Mode"
        '
        'ToolStripHelp
        '
        Me.ToolStripHelp.Name = "ToolStripHelp"
        Me.ToolStripHelp.Size = New System.Drawing.Size(327, 22)
        Me.ToolStripHelp.Text = "Help"
        '
        'MenuStripFilter
        '
        Me.MenuStripFilter.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ClearToolStripMenuItem, Me.ToolStripSeparator1, Me.AddFilterToolStripMenuItem, Me.SldprtToolStripMenuItem, Me.SldasmToolStripMenuItem, Me.SlddrwToolStripMenuItem})
        Me.MenuStripFilter.Name = "ContextMenuStrip2"
        Me.MenuStripFilter.Size = New System.Drawing.Size(129, 120)
        '
        'ClearToolStripMenuItem
        '
        Me.ClearToolStripMenuItem.Name = "ClearToolStripMenuItem"
        Me.ClearToolStripMenuItem.Size = New System.Drawing.Size(128, 22)
        Me.ClearToolStripMenuItem.Text = "Clear"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(125, 6)
        '
        'AddFilterToolStripMenuItem
        '
        Me.AddFilterToolStripMenuItem.Enabled = False
        Me.AddFilterToolStripMenuItem.Name = "AddFilterToolStripMenuItem"
        Me.AddFilterToolStripMenuItem.Size = New System.Drawing.Size(128, 22)
        Me.AddFilterToolStripMenuItem.Text = "Add Filter:"
        '
        'SldprtToolStripMenuItem
        '
        Me.SldprtToolStripMenuItem.Name = "SldprtToolStripMenuItem"
        Me.SldprtToolStripMenuItem.Size = New System.Drawing.Size(128, 22)
        Me.SldprtToolStripMenuItem.Text = "*.sldprt"
        '
        'SldasmToolStripMenuItem
        '
        Me.SldasmToolStripMenuItem.Name = "SldasmToolStripMenuItem"
        Me.SldasmToolStripMenuItem.Size = New System.Drawing.Size(128, 22)
        Me.SldasmToolStripMenuItem.Text = "*.sldasm"
        '
        'SlddrwToolStripMenuItem
        '
        Me.SlddrwToolStripMenuItem.Name = "SlddrwToolStripMenuItem"
        Me.SlddrwToolStripMenuItem.Size = New System.Drawing.Size(128, 22)
        Me.SlddrwToolStripMenuItem.Text = "*.slddrw"
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(727, 46)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(149, 8)
        Me.ProgressBar1.Step = 1
        Me.ProgressBar1.TabIndex = 71
        '
        'BtBrowseFolder
        '
        Me.BtBrowseFolder.Location = New System.Drawing.Point(425, 30)
        Me.BtBrowseFolder.Name = "BtBrowseFolder"
        Me.BtBrowseFolder.Size = New System.Drawing.Size(144, 23)
        Me.BtBrowseFolder.TabIndex = 72
        Me.BtBrowseFolder.Text = "Browse for Folder..."
        Me.BtBrowseFolder.UseVisualStyleBackColor = True
        '
        'LabelHighlight
        '
        Me.LabelHighlight.AutoSize = True
        Me.LabelHighlight.Location = New System.Drawing.Point(6, 64)
        Me.LabelHighlight.Name = "LabelHighlight"
        Me.LabelHighlight.Size = New System.Drawing.Size(51, 13)
        Me.LabelHighlight.TabIndex = 74
        Me.LabelHighlight.Text = "Highlight:"
        '
        'TbHighlight
        '
        Me.TbHighlight.Location = New System.Drawing.Point(69, 60)
        Me.TbHighlight.Name = "TbHighlight"
        Me.TbHighlight.Size = New System.Drawing.Size(158, 20)
        Me.TbHighlight.TabIndex = 73
        '
        'BtExport
        '
        Me.BtExport.Location = New System.Drawing.Point(271, 59)
        Me.BtExport.Name = "BtExport"
        Me.BtExport.Size = New System.Drawing.Size(144, 23)
        Me.BtExport.TabIndex = 75
        Me.BtExport.Text = "Export to Spreadsheet"
        Me.BtExport.UseVisualStyleBackColor = True
        '
        'BtImport
        '
        Me.BtImport.Location = New System.Drawing.Point(425, 59)
        Me.BtImport.Name = "BtImport"
        Me.BtImport.Size = New System.Drawing.Size(144, 23)
        Me.BtImport.TabIndex = 76
        Me.BtImport.Text = "Import Spreadsheet..."
        Me.BtImport.UseVisualStyleBackColor = True
        '
        'SaveFileDialog1
        '
        '
        'BatchPropForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(886, 469)
        Me.Controls.Add(Me.BtImport)
        Me.Controls.Add(Me.BtExport)
        Me.Controls.Add(Me.LabelHighlight)
        Me.Controls.Add(Me.TbHighlight)
        Me.Controls.Add(Me.BtBrowseFolder)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.LabelSearch)
        Me.Controls.Add(Me.BtOptions)
        Me.Controls.Add(Me.DGV2)
        Me.Controls.Add(Me.CbIncludeSubFolders)
        Me.Controls.Add(Me.BtSearch)
        Me.Controls.Add(Me.TbFolder)
        Me.Controls.Add(Me.BtBrowseAssy)
        Me.Controls.Add(Me.DGV1)
        Me.Controls.Add(Me.LabelFolder)
        Me.Controls.Add(Me.TbFilter)
        Me.Name = "BatchPropForm"
        Me.Text = "Batch Custom Properties"
        Me.MenuStripColumn.ResumeLayout(False)
        CType(Me.DGV1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DGV2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.MenuStripOptions.ResumeLayout(False)
        Me.MenuStripFilter.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents MenuStripColumn As ContextMenuStrip
    Friend WithEvents MenuDeleteColumn As ToolStripMenuItem
    Friend WithEvents MenuDeleteColAndCustProp As ToolStripMenuItem
    Friend WithEvents MenuInsertColumn As ToolStripMenuItem
    Friend WithEvents MenuInsertTypeText As ToolStripMenuItem
    Friend WithEvents MenuInsertTypeNumber As ToolStripMenuItem
    Friend WithEvents MenuInsertTypeDate As ToolStripMenuItem
    Friend WithEvents MenuInsertTypeYesNo As ToolStripMenuItem
    Friend WithEvents CbIncludeSubFolders As CheckBox
    Friend WithEvents BtSearch As Button
    Friend WithEvents TbFolder As TextBox
    Friend WithEvents BtBrowseAssy As Button
    Friend WithEvents DGV1 As DataGridView
    Friend WithEvents LabelFolder As Label
    Friend WithEvents TbFilter As TextBox
    Friend WithEvents DGV2 As DataGridView
    Friend WithEvents MenuApplyRule As ToolStripMenuItem
    Friend WithEvents MenuStripRules As ContextMenuStrip
    Friend WithEvents BtOptions As Button
    Friend WithEvents LabelSearch As Label
    Friend WithEvents MenuStripOptions As ContextMenuStrip
    Friend WithEvents ToolStripSaveColumns As ToolStripMenuItem
    Friend WithEvents ToolStripLoadColumns As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator3 As ToolStripSeparator
    Friend WithEvents ToolStripLoadCustomProp As ToolStripMenuItem
    Friend WithEvents ToolStripLoadAllConfig As ToolStripMenuItem
    Friend WithEvents ToolStripLoadActiveConfig As ToolStripMenuItem
    Friend WithEvents ToolStripLoadCutList As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator4 As ToolStripSeparator
    Friend WithEvents ToolStripHelp As ToolStripMenuItem
    Friend WithEvents MenuStripFilter As ContextMenuStrip
    Friend WithEvents AddFilterToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator1 As ToolStripSeparator
    Friend WithEvents SldprtToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents SldasmToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents SlddrwToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ClearToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ProgressBar1 As ProgressBar
    Friend WithEvents MenuDeleteColumnS As ToolStripMenuItem
    Friend WithEvents MenuDeleteColAndCustPropS As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator5 As ToolStripSeparator
    Friend WithEvents ToolStripSeparator2 As ToolStripSeparator
    Friend WithEvents ToolStripOnlyShowSavedCol As ToolStripMenuItem
    Friend WithEvents MenuApplyAllRules As ToolStripMenuItem
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents BtBrowseFolder As Button
    Friend WithEvents ToolStripSwitchMode As ToolStripMenuItem
    Friend WithEvents LabelHighlight As Label
    Friend WithEvents TbHighlight As TextBox
    Friend WithEvents BtExport As Button
    Friend WithEvents BtImport As Button
    Friend WithEvents ToolStripDeleteProp As ToolStripMenuItem
    Friend WithEvents ToolStripCreateMissingProp As ToolStripMenuItem
    Friend WithEvents SaveFileDialog1 As SaveFileDialog
End Class
