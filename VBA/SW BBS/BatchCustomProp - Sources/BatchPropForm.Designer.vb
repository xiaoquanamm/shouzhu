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
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ApplyRuleToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MenuDeleteColumn = New System.Windows.Forms.ToolStripMenuItem()
        Me.MenuDeleteColAndCustProp = New System.Windows.Forms.ToolStripMenuItem()
        Me.MenuInsertColumn = New System.Windows.Forms.ToolStripMenuItem()
        Me.MenuInsertTypeText = New System.Windows.Forms.ToolStripMenuItem()
        Me.MenuInsertTypeNumber = New System.Windows.Forms.ToolStripMenuItem()
        Me.MenuInsertTypeDate = New System.Windows.Forms.ToolStripMenuItem()
        Me.MenuInsertTypeYesNo = New System.Windows.Forms.ToolStripMenuItem()
        Me.CbIncludeSubFolders = New System.Windows.Forms.CheckBox()
        Me.CbLoadColumns = New System.Windows.Forms.CheckBox()
        Me.BtSaveColumns = New System.Windows.Forms.Button()
        Me.BtSearch = New System.Windows.Forms.Button()
        Me.TbFolder = New System.Windows.Forms.TextBox()
        Me.BtBrowse = New System.Windows.Forms.Button()
        Me.DGV1 = New System.Windows.Forms.DataGridView()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TbFilter = New System.Windows.Forms.TextBox()
        Me.DGV2 = New System.Windows.Forms.DataGridView()
        Me.CbLoadCustomProp = New System.Windows.Forms.CheckBox()
        Me.CbLoadAllConfig = New System.Windows.Forms.CheckBox()
        Me.CbLoadActiveConfig = New System.Windows.Forms.CheckBox()
        Me.ContextMenuStrip2 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ContextMenuStrip1.SuspendLayout()
        CType(Me.DGV1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DGV2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ApplyRuleToolStripMenuItem, Me.MenuDeleteColumn, Me.MenuDeleteColAndCustProp, Me.MenuInsertColumn})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(230, 92)
        '
        'ApplyRuleToolStripMenuItem
        '
        Me.ApplyRuleToolStripMenuItem.Name = "ApplyRuleToolStripMenuItem"
        Me.ApplyRuleToolStripMenuItem.Size = New System.Drawing.Size(229, 22)
        Me.ApplyRuleToolStripMenuItem.Text = "Apply Rule"
        '
        'MenuDeleteColumn
        '
        Me.MenuDeleteColumn.Name = "MenuDeleteColumn"
        Me.MenuDeleteColumn.Size = New System.Drawing.Size(229, 22)
        Me.MenuDeleteColumn.Text = "Delete Column"
        '
        'MenuDeleteColAndCustProp
        '
        Me.MenuDeleteColAndCustProp.Name = "MenuDeleteColAndCustProp"
        Me.MenuDeleteColAndCustProp.Size = New System.Drawing.Size(229, 22)
        Me.MenuDeleteColAndCustProp.Text = "Delete Column And Custom Prop"
        '
        'MenuInsertColumn
        '
        Me.MenuInsertColumn.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MenuInsertTypeText, Me.MenuInsertTypeNumber, Me.MenuInsertTypeDate, Me.MenuInsertTypeYesNo})
        Me.MenuInsertColumn.Name = "MenuInsertColumn"
        Me.MenuInsertColumn.Size = New System.Drawing.Size(229, 22)
        Me.MenuInsertColumn.Text = "Insert Column"
        '
        'MenuInsertTypeText
        '
        Me.MenuInsertTypeText.Name = "MenuInsertTypeText"
        Me.MenuInsertTypeText.Size = New System.Drawing.Size(111, 22)
        Me.MenuInsertTypeText.Text = "Text"
        '
        'MenuInsertTypeNumber
        '
        Me.MenuInsertTypeNumber.Name = "MenuInsertTypeNumber"
        Me.MenuInsertTypeNumber.Size = New System.Drawing.Size(111, 22)
        Me.MenuInsertTypeNumber.Text = "Number"
        '
        'MenuInsertTypeDate
        '
        Me.MenuInsertTypeDate.Name = "MenuInsertTypeDate"
        Me.MenuInsertTypeDate.Size = New System.Drawing.Size(111, 22)
        Me.MenuInsertTypeDate.Text = "Date"
        '
        'MenuInsertTypeYesNo
        '
        Me.MenuInsertTypeYesNo.Name = "MenuInsertTypeYesNo"
        Me.MenuInsertTypeYesNo.Size = New System.Drawing.Size(111, 22)
        Me.MenuInsertTypeYesNo.Text = "Yes/No"
        '
        'CbIncludeSubFolders
        '
        Me.CbIncludeSubFolders.AutoSize = True
        Me.CbIncludeSubFolders.Location = New System.Drawing.Point(556, 7)
        Me.CbIncludeSubFolders.Name = "CbIncludeSubFolders"
        Me.CbIncludeSubFolders.Size = New System.Drawing.Size(120, 17)
        Me.CbIncludeSubFolders.TabIndex = 59
        Me.CbIncludeSubFolders.Text = "Include Sub Folders"
        Me.CbIncludeSubFolders.UseVisualStyleBackColor = True
        '
        'CbLoadColumns
        '
        Me.CbLoadColumns.AutoSize = True
        Me.CbLoadColumns.Location = New System.Drawing.Point(103, 33)
        Me.CbLoadColumns.Name = "CbLoadColumns"
        Me.CbLoadColumns.Size = New System.Drawing.Size(182, 17)
        Me.CbLoadColumns.TabIndex = 58
        Me.CbLoadColumns.Text = "Load Saved Columns on Refresh"
        Me.CbLoadColumns.UseVisualStyleBackColor = True
        '
        'BtSaveColumns
        '
        Me.BtSaveColumns.Location = New System.Drawing.Point(4, 30)
        Me.BtSaveColumns.Name = "BtSaveColumns"
        Me.BtSaveColumns.Size = New System.Drawing.Size(90, 20)
        Me.BtSaveColumns.TabIndex = 55
        Me.BtSaveColumns.Text = "Save Columns"
        Me.BtSaveColumns.UseVisualStyleBackColor = True
        '
        'BtSearch
        '
        Me.BtSearch.Location = New System.Drawing.Point(4, 4)
        Me.BtSearch.Name = "BtSearch"
        Me.BtSearch.Size = New System.Drawing.Size(90, 20)
        Me.BtSearch.TabIndex = 54
        Me.BtSearch.Text = "Search"
        Me.BtSearch.UseVisualStyleBackColor = True
        '
        'TbFolder
        '
        Me.TbFolder.Location = New System.Drawing.Point(250, 4)
        Me.TbFolder.Name = "TbFolder"
        Me.TbFolder.Size = New System.Drawing.Size(232, 20)
        Me.TbFolder.TabIndex = 53
        '
        'BtBrowse
        '
        Me.BtBrowse.Location = New System.Drawing.Point(488, 4)
        Me.BtBrowse.Name = "BtBrowse"
        Me.BtBrowse.Size = New System.Drawing.Size(60, 20)
        Me.BtBrowse.TabIndex = 49
        Me.BtBrowse.Text = "Browse..."
        Me.BtBrowse.UseVisualStyleBackColor = True
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
        Me.DGV1.Location = New System.Drawing.Point(1, 58)
        Me.DGV1.Name = "DGV1"
        Me.DGV1.RowHeadersVisible = False
        Me.DGV1.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DGV1.Size = New System.Drawing.Size(675, 85)
        Me.DGV1.TabIndex = 52
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(203, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(47, 13)
        Me.Label1.TabIndex = 51
        Me.Label1.Text = "in folder:"
        '
        'TbFilter
        '
        Me.TbFilter.Location = New System.Drawing.Point(103, 4)
        Me.TbFilter.Name = "TbFilter"
        Me.TbFilter.Size = New System.Drawing.Size(89, 20)
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
        Me.DGV2.Location = New System.Drawing.Point(1, 142)
        Me.DGV2.Name = "DGV2"
        Me.DGV2.RowHeadersVisible = False
        Me.DGV2.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DGV2.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DGV2.Size = New System.Drawing.Size(675, 196)
        Me.DGV2.TabIndex = 61
        '
        'CbLoadCustomProp
        '
        Me.CbLoadCustomProp.AutoSize = True
        Me.CbLoadCustomProp.Location = New System.Drawing.Point(302, 33)
        Me.CbLoadCustomProp.Name = "CbLoadCustomProp"
        Me.CbLoadCustomProp.Size = New System.Drawing.Size(113, 17)
        Me.CbLoadCustomProp.TabIndex = 63
        Me.CbLoadCustomProp.Text = "Load Custom Prop"
        Me.CbLoadCustomProp.UseVisualStyleBackColor = True
        '
        'CbLoadAllConfig
        '
        Me.CbLoadAllConfig.AutoSize = True
        Me.CbLoadAllConfig.Location = New System.Drawing.Point(437, 33)
        Me.CbLoadAllConfig.Name = "CbLoadAllConfig"
        Me.CbLoadAllConfig.Size = New System.Drawing.Size(102, 17)
        Me.CbLoadAllConfig.TabIndex = 62
        Me.CbLoadAllConfig.Text = "Load All Configs"
        Me.CbLoadAllConfig.UseVisualStyleBackColor = True
        '
        'CbLoadActiveConfig
        '
        Me.CbLoadActiveConfig.AutoSize = True
        Me.CbLoadActiveConfig.Location = New System.Drawing.Point(556, 33)
        Me.CbLoadActiveConfig.Name = "CbLoadActiveConfig"
        Me.CbLoadActiveConfig.Size = New System.Drawing.Size(116, 17)
        Me.CbLoadActiveConfig.TabIndex = 64
        Me.CbLoadActiveConfig.Text = "Load Active Config"
        Me.CbLoadActiveConfig.UseVisualStyleBackColor = True
        '
        'ContextMenuStrip2
        '
        Me.ContextMenuStrip2.Name = "ContextMenuStrip2"
        Me.ContextMenuStrip2.Size = New System.Drawing.Size(61, 4)
        '
        'BatchPropForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(677, 339)
        Me.Controls.Add(Me.CbLoadActiveConfig)
        Me.Controls.Add(Me.CbLoadCustomProp)
        Me.Controls.Add(Me.CbLoadAllConfig)
        Me.Controls.Add(Me.DGV2)
        Me.Controls.Add(Me.CbIncludeSubFolders)
        Me.Controls.Add(Me.CbLoadColumns)
        Me.Controls.Add(Me.BtSaveColumns)
        Me.Controls.Add(Me.BtSearch)
        Me.Controls.Add(Me.TbFolder)
        Me.Controls.Add(Me.BtBrowse)
        Me.Controls.Add(Me.DGV1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TbFilter)
        Me.Name = "BatchPropForm"
        Me.Text = "Form1"
        Me.ContextMenuStrip1.ResumeLayout(False)
        CType(Me.DGV1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DGV2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents FolderBrowserDialog1 As FolderBrowserDialog
    Friend WithEvents ContextMenuStrip1 As ContextMenuStrip
    Friend WithEvents MenuDeleteColumn As ToolStripMenuItem
    Friend WithEvents MenuDeleteColAndCustProp As ToolStripMenuItem
    Friend WithEvents MenuInsertColumn As ToolStripMenuItem
    Friend WithEvents MenuInsertTypeText As ToolStripMenuItem
    Friend WithEvents MenuInsertTypeNumber As ToolStripMenuItem
    Friend WithEvents MenuInsertTypeDate As ToolStripMenuItem
    Friend WithEvents MenuInsertTypeYesNo As ToolStripMenuItem
    Friend WithEvents CbIncludeSubFolders As CheckBox
    Friend WithEvents CbLoadColumns As CheckBox
    Friend WithEvents BtSaveColumns As Button
    Friend WithEvents BtSearch As Button
    Friend WithEvents TbFolder As TextBox
    Friend WithEvents BtBrowse As Button
    Friend WithEvents DGV1 As DataGridView
    Friend WithEvents Label1 As Label
    Friend WithEvents TbFilter As TextBox
    Friend WithEvents DGV2 As DataGridView
    Friend WithEvents ApplyRuleToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents CbLoadCustomProp As CheckBox
    Friend WithEvents CbLoadAllConfig As CheckBox
    Friend WithEvents CbLoadActiveConfig As CheckBox
    Friend WithEvents ContextMenuStrip2 As ContextMenuStrip
End Class
