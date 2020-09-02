VERSION 5.00
Begin VB.Form UserForm1 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Solidworks辅助软件"
   ClientHeight    =   1680
   ClientLeft      =   315
   ClientTop       =   510
   ClientWidth     =   10425
   DrawMode        =   1  'Blackness
   DrawStyle       =   4  'Dash-Dot-Dot
   FillColor       =   &H0000FF00&
   FillStyle       =   7  'Diagonal Cross
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   10.5
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FF00FF&
   Icon            =   "20191205改名.frx":0000
   KeyPreview      =   -1  'True
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   2  'Custom
   ScaleHeight     =   84
   ScaleMode       =   2  'Point
   ScaleWidth      =   521.25
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      Caption         =   "参数界面"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Left            =   0
      TabIndex        =   30
      Top             =   2160
      Width           =   10335
      Begin VB.TextBox colure1 
         Height          =   315
         Left            =   120
         TabIndex        =   42
         Text            =   "1"
         Top             =   5880
         Width           =   1575
      End
      Begin VB.ListBox 保存 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2790
         Left            =   6720
         TabIndex        =   36
         Top             =   240
         Width           =   3495
      End
      Begin VB.FileListBox File1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2610
         Left            =   6720
         TabIndex        =   35
         Top             =   3240
         Width           =   3495
      End
      Begin VB.ListBox 判断 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2580
         Left            =   3360
         TabIndex        =   34
         Top             =   3240
         Width           =   3135
      End
      Begin VB.ListBox 同名 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2790
         Left            =   3360
         TabIndex        =   33
         Top             =   240
         Width           =   3135
      End
      Begin VB.ListBox 显参 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2580
         Left            =   120
         TabIndex        =   32
         Top             =   3240
         Width           =   2895
      End
      Begin VB.ListBox 运参 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2790
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FF80FF&
      Caption         =   "文件"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   10680
      TabIndex        =   10
      Top             =   0
      Width           =   10455
      Begin VB.CommandButton 统计数量 
         Caption         =   "数量"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   41
         Top             =   1200
         Width           =   1095
      End
      Begin VB.ComboBox 处理方式 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1320
         TabIndex        =   40
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox 工程图模板位置 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1320
         TabIndex        =   39
         Top             =   1200
         Width           =   9015
      End
      Begin VB.TextBox 左边 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   9120
         TabIndex        =   21
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox 上边 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   7920
         TabIndex        =   20
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox 保存DWG文档位置 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3360
         MultiLine       =   -1  'True
         TabIndex        =   19
         Top             =   720
         Width           =   6975
      End
      Begin VB.CommandButton 录入保存 
         Caption         =   "录入保存"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox 当前日期 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6360
         TabIndex        =   14
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox 批准 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4080
         TabIndex        =   13
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox 审核 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2040
         TabIndex        =   12
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox 设计 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5160
      TabIndex        =   6
      Top             =   -120
      Width           =   5295
      Begin VB.CheckBox 输入2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         Caption         =   "手2"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   2520
         TabIndex        =   38
         Top             =   480
         Width           =   615
      End
      Begin VB.CheckBox 输入1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         Caption         =   "材1"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   37
         Top             =   120
         Width           =   615
      End
      Begin VB.CommandButton 颜色 
         Caption         =   "颜色"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4320
         TabIndex        =   29
         Top             =   240
         Width           =   850
      End
      Begin VB.CommandButton 文件输出 
         Caption         =   "输出"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4320
         TabIndex        =   28
         Top             =   720
         Width           =   850
      End
      Begin VB.CheckBox 桌面 
         BackColor       =   &H00FFFF80&
         Caption         =   "桌面"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3360
         TabIndex        =   27
         Top             =   840
         Width           =   735
      End
      Begin VB.ComboBox 输出文件选择 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1680
         TabIndex        =   26
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox 加工件数量 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1680
         TabIndex        =   23
         Text            =   "1"
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton 属性修改 
         Caption         =   "属性"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3360
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
      Begin VB.ComboBox 零件名字 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1455
      End
      Begin VB.ComboBox 材料 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   5175
      Begin VB.CommandButton 还原 
         Appearance      =   0  'Flat
         Caption         =   "还原"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3240
         TabIndex        =   25
         Top             =   240
         Width           =   850
      End
      Begin VB.ComboBox 改名选择 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3120
         TabIndex        =   24
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton 读取文件 
         Caption         =   "读取"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4200
         TabIndex        =   22
         Top             =   240
         Width           =   850
      End
      Begin VB.CommandButton 零件改名 
         Caption         =   "改名"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4200
         TabIndex        =   5
         Top             =   720
         Width           =   850
      End
      Begin VB.TextBox 项目编号 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   600
         TabIndex        =   4
         Top             =   720
         Width           =   2415
      End
      Begin VB.TextBox 机台名字 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   600
         TabIndex        =   2
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label2 
         BackColor       =   &H0080FF80&
         Caption         =   "项目"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H0080FF80&
         Caption         =   "机台"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   1
         Top             =   255
         Width           =   495
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   16
      Top             =   960
      Width           =   10455
      Begin VB.TextBox 显示内容 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   18
         Top             =   240
         Width           =   10215
      End
      Begin VB.Label Label6 
         Caption         =   "Label6"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   1800
         TabIndex        =   17
         Top             =   720
         Width           =   15
      End
   End
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public swModel2 As SldWorks.ModelDoc2
Public PARTNAME_Value_temp As String
Public MATERIAL_Value2_temp As String

Private Sub Form_Load()

On Error Resume Next

SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, 2 Or 1

Call 界面参数(材料, 处理方式, 零件名字, 加工件类别, 改名选择, 输出文件选择)

Call 界面读取(运参)

Call 外部文档写入界面(运参, 机台名字, 项目编号, 设计, 审核, 批准, 左边, 上边, 保存DWG文档位置, 工程图模板位置, 当前日期)

UserForm1.Left = 左边.Text
UserForm1.Top = 上边.Text
 
End Sub

Private Sub 录入保存_Click()

Call 界面录入保存(机台名字, 项目编号, 设计, 审核, 批准, 左边, 上边, 保存DWG文档位置, 工程图模板位置, 显参)

Call 显示模块(显参, 显示内容, 项目编号)

End Sub

Private Sub cancel_button_Click()
    Unload UserForm1
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then '如果按的键是Esc，
'End '那么退出程序
End If
End Sub
'==============================================================================、

Private Sub 还原_Click()

Call 初始化(运参, 显参, 同名, 判断, 显示内容, 保存)

Call 轻化还原(显参)

Call 显示模块(显参, 显示内容, 项目编号)

End Sub

'=================================================
'读取文件
'=================================================
Private Sub 读取文件_Click()

Call 初始化(运参, 显参, 同名, 判断, 显示内容, 保存)

Call 读取基本信息(显参, 运参, 保存, 保存DWG文档位置, 当前日期, 机台名字)

Call 显示模块(显参, 显示内容, 项目编号)

End Sub

'=================================================
'零件改名
'=================================================
Private Sub 零件改名_Click()

Call 初始化(运参, 显参, 同名, 判断, 显示内容, 保存)

Call 读取基本信息(显参, 运参, 保存, 保存DWG文档位置, 当前日期, 机台名字)

Call 点选数量统计(运参, 同名, 判断, 显参, 保存)

Call 改名(运参, 显参, 同名, 保存, 判断, 项目编号, 改名选择)  '

'Call 保存文件

Call 显示模块(显参, 显示内容, 项目编号)

Set swpart = Nothing
Set swModel = Nothing
Set swapp = Nothing

End Sub

'=================================================
'属性修改
'=================================================

Private Sub 属性修改_Click()

Call 初始化(运参, 显参, 同名, 判断, 显示内容, 保存)

Call 读取基本信息(显参, 运参, 保存, 保存DWG文档位置, 当前日期, 机台名字)

Call 属性值分选(材料, 处理方式, 零件名字, 输入1, 输入2, 同名, 机台名字, 项目编号, 设计, 审核, 批准, 当前日期, 加工件数量)

Call 新建工程图(同名, 运参, 显参, 工程图模板位置)

If (运参.List(10) = "新建" Or 运参.List(3) = "SLDDRW") Then

Call 名字翻译(同名, 零件名字)

Call 清除属性

Call 配置属性(同名, 显参)

Else

显参.List(0) = "错误"

显参.List(1) = "工程图打开，未添加属性"

End If

Call 保存文件

Call 显示模块(显参, 显示内容, 项目编号)

Set swpart = Nothing
Set swModel = Nothing
Set swapp = Nothing

End Sub

'=================================================
'数量统计
'=================================================

Private Sub 统计数量_Click()

Call 初始化(运参, 显参, 同名, 判断, 显示内容, 保存)

Call 数量统计(显参)

Call 显示模块(显参, 显示内容, 项目编号)

End Sub

'=================================================
'颜色修改
'=================================================
Private Sub 颜色_Click()

Call 初始化(运参, 显参, 同名, 判断, 显示内容, 保存)

Call 颜色修改(显参, 运参, colure1)

Call 显示模块(显参, 显示内容, 项目编号)

End Sub

'=================================================
'文件处理
'=================================================

Private Sub 文件输出_Click()

Call 初始化(运参, 显参, 同名, 判断, 显示内容, 保存)

Call 读取基本信息(显参, 运参, 保存, 保存DWG文档位置, 当前日期, 机台名字)

Dim strMMM As String

strMMM = 输出文件选择.Text

Select Case strMMM

Case Is = "项目建立"

Call 项目文件建立(保存DWG文档位置, 机台名字, 显参)

Case Is = "Dwg/Step"

Call DWG和PDF保存(保存DWG文档位置, 显示内容, 显参, 桌面, 运参)

Case Is = "Part"

Call Part保存(保存DWG文档位置, 显示内容, 显参, 桌面, 运参)

Case Is = "属性清除"

Call 批量附件属性值写入(显参)

Case Is = "批量Bom"

Call 读取基本信息(显参, 运参, 保存, 保存DWG文档位置, 当前日期, 机台名字)

Call 新建文档(保存DWG文档位置, 显参, 桌面, 运参)

Call 工程图查找(运参, 保存, 判断, 同名, File1, 桌面, 显参)

Call 生成Bom(运参, 保存, 桌面)

Call 写入表格(同名, 保存)

Case Is = "单个Bom"

Call 读取基本信息(显参, 运参, 保存, 保存DWG文档位置, 当前日期, 机台名字)

Call 单个工程图Bom(保存, 运参, 桌面, 判断, 同名, 显参)

Case Else
       
End Select

Call 显示模块(显参, 显示内容, 项目编号)

End Sub






