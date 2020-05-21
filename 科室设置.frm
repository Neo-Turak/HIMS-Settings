VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form 科室设置 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "科室设置"
   ClientHeight    =   5055
   ClientLeft      =   495
   ClientTop       =   825
   ClientWidth     =   7455
   Icon            =   "科室设置.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   7455
   Begin VB.ComboBox Combo1 
      DataField       =   "是否临床"
      BeginProperty DataFormat 
         Type            =   5
         Format          =   ""
         HaveTrueFalseNull=   1
         TrueValue       =   "否"
         FalseValue      =   "是"
         NullValue       =   "否"
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   7
      EndProperty
      DataSource      =   "Adodc1"
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
      ItemData        =   "科室设置.frx":038A
      Left            =   5760
      List            =   "科室设置.frx":0394
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "Combo1"
      Top             =   3000
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "科室设置.frx":03A0
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   6376
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   0   'False
      BackColor       =   65408
      Enabled         =   0   'False
      HeadLines       =   1
      RowHeight       =   18
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "编码"
         Caption         =   "编码"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "名称"
         Caption         =   "名称"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "是否临床"
         Caption         =   "是否临床"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "助记码"
         Caption         =   "助记码"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   750.047
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1230.236
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1379.906
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   600
      Top             =   4560
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=G:\医院管理系统\系统管理\settings.mdb;Persist Security Info=true"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=G:\医院管理系统\系统管理\settings.mdb;Persist Security Info=true"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "科室"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSForms.CommandButton CommandButton3 
      Height          =   495
      Left            =   2160
      TabIndex        =   10
      Top             =   3960
      Visible         =   0   'False
      Width           =   1335
      Caption         =   "保存"
      Size            =   "2355;873"
      FontName        =   "新宋体"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton CommandButton2 
      Height          =   495
      Left            =   3960
      TabIndex        =   9
      Top             =   3960
      Width           =   1335
      Caption         =   "删除"
      Size            =   "2355;873"
      FontName        =   "新宋体"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   495
      Left            =   360
      TabIndex        =   8
      Top             =   3960
      Width           =   1335
      Caption         =   "添加"
      Size            =   "2355;873"
      FontName        =   "新宋体"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label1 
      Height          =   375
      Index           =   3
      Left            =   5940
      TabIndex        =   7
      Top             =   2640
      Width           =   1095
      VariousPropertyBits=   8388627
      Caption         =   "是否临床"
      Size            =   "1931;661"
      FontName        =   "宋体"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox TextBox3 
      DataField       =   "助记码"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   5760
      TabIndex        =   6
      Top             =   2160
      Width           =   1575
      VariousPropertyBits=   746604575
      Size            =   "2778;661"
      FontName        =   "宋体"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox TextBox2 
      DataField       =   "名称"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   5760
      TabIndex        =   5
      Top             =   1320
      Width           =   1575
      VariousPropertyBits=   746604575
      Size            =   "2778;661"
      FontName        =   "宋体"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox TextBox1 
      DataField       =   "编码"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   5760
      TabIndex        =   4
      Top             =   480
      Width           =   1575
      VariousPropertyBits=   746604575
      Size            =   "2778;661"
      FontName        =   "宋体"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label1 
      Height          =   375
      Index           =   2
      Left            =   6120
      TabIndex        =   3
      Top             =   1800
      Width           =   735
      VariousPropertyBits=   8388627
      Caption         =   "助记码"
      Size            =   "1296;661"
      FontName        =   "宋体"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label1 
      Height          =   375
      Index           =   1
      Left            =   6240
      TabIndex        =   2
      Top             =   120
      Width           =   495
      VariousPropertyBits=   8388627
      Caption         =   "编码"
      Size            =   "873;661"
      FontName        =   "宋体"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label1 
      Height          =   255
      Index           =   0
      Left            =   6000
      TabIndex        =   1
      Top             =   960
      Width           =   975
      VariousPropertyBits=   8388627
      Caption         =   "科室名称"
      Size            =   "1720;450"
      FontName        =   "宋体"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
End
Attribute VB_Name = "科室设置"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mHZtoSM As cHztoSM
Private Sub CommandButton1_Click()
Me.Width = 7545
Adodc1.Recordset.AddNew
TextBox1.Locked = False
TextBox2.Locked = False
TextBox3.Locked = False
Combo1.Locked = False
CommandButton1.Locked = True
CommandButton3.Locked = False
TextBox1.SetFocus
Me.CommandButton3.Visible = True
End Sub

Private Sub CommandButton2_Click()
Adodc1.Recordset.Delete adAffectCurrent
Me.Width = 5685
End Sub

Private Sub CommandButton3_Click()
If TextBox1.Text <> "" And TextBox2.Text <> "" And TextBox3.Text <> "" Then
Adodc1.Recordset.Update
Else
MsgBox "请填写必要内容", vbInformation, "科室设置"
TextBox1.SetFocus
End If
End Sub


Private Sub Form_Load()
Set mHZtoSM = New cHztoSM
    
    mHZtoSM.LoadLibFile App.Path & "\GB2312SM.Lib"
    
    If mHZtoSM.LoadLibSuccess = False Then Unload Me
    Me.Width = 5685
    Adodc1.Recordset.MoveLast
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set mHZtoSM = Nothing
End Sub


Private Sub TextBox3_GotFocus()
TextBox3.Text = mHZtoSM.HZtoSMEx(TextBox2.Text)
End Sub
