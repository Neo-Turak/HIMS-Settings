VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl32.ocx"
Begin VB.MDIForm 系统设置 
   BackColor       =   &H8000000C&
   Caption         =   "系统设置"
   ClientHeight    =   7185
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   11715
   Icon            =   "系统设置.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "系统设置.frx":1082
   StartUpPosition =   1  '所有者中心
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1320
      Top             =   1440
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   6810
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   2469
            MinWidth        =   2469
            TextSave        =   "2016-06-10"
            Object.ToolTipText     =   "当前日期"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "12:51"
            Object.ToolTipText     =   "当前时间"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
            Object.ToolTipText     =   "当前用户"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2469
            MinWidth        =   2469
            Object.ToolTipText     =   "职位"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   6350
            MinWidth        =   6350
            Text            =   "为了一切病人,为了病人一切"
            TextSave        =   "为了一切病人,为了病人一切"
            Object.ToolTipText     =   "医院标语"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "华文新魏"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu 常规设置 
      Caption         =   "常规设置"
      Index           =   1
      NegotiatePosition=   1  'Left
      WindowList      =   -1  'True
      Begin VB.Menu 病区 
         Caption         =   "病区设置"
         Index           =   2
         Shortcut        =   ^B
      End
      Begin VB.Menu 床位 
         Caption         =   "床位设置"
         Index           =   3
      End
      Begin VB.Menu 科室 
         Caption         =   "科室管理"
         Index           =   4
      End
      Begin VB.Menu 医生 
         Caption         =   "医生管理"
         Index           =   5
      End
      Begin VB.Menu 护士 
         Caption         =   "护士管理"
         Index           =   6
      End
      Begin VB.Menu 药库 
         Caption         =   "药库设置"
         Index           =   7
      End
      Begin VB.Menu 项目 
         Caption         =   "检查项目管理"
         Index           =   8
      End
   End
   Begin VB.Menu 用户 
      Caption         =   "用户管理"
      Index           =   10
   End
   Begin VB.Menu 账目查询 
      Caption         =   "账目查询"
      Begin VB.Menu 今日就诊记录 
         Caption         =   "今日就诊记录"
      End
      Begin VB.Menu 今日住院记录 
         Caption         =   "今日住院记录"
      End
      Begin VB.Menu 今日收款记录 
         Caption         =   "今日收款记录"
      End
      Begin VB.Menu 住院病人统计 
         Caption         =   "住院病人统计"
      End
   End
End
Attribute VB_Name = "系统设置"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Timer1_Timer()
StatusBar1.Panels(2) = Time
StatusBar1.Panels(1) = Date
End Sub

Private Sub 病区_Click(Index As Integer)
病区设置.Show
End Sub

Private Sub 床位_Click(Index As Integer)
床位设置.Show
End Sub

Private Sub 护士_Click(Index As Integer)
护士管理.Show
End Sub

Private Sub 今日就诊记录_Click()
系统.今日就诊记录.Show
End Sub

Private Sub 科室_Click(Index As Integer)
科室设置.Show
End Sub

Private Sub 项目_Click(Index As Integer)
检查科室项目设置.Show
End Sub

Private Sub 药库_Click(Index As Integer)
药库设置.Show
End Sub

Private Sub 医生_Click(Index As Integer)
医生设置.Show
End Sub

Private Sub 用户_Click(Index As Integer)
用户管理.Show
End Sub
