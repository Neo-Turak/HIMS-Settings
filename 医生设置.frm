VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form 医生设置 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "医生设置"
   ClientHeight    =   4935
   ClientLeft      =   2175
   ClientTop       =   1890
   ClientWidth     =   8145
   Icon            =   "医生设置.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   8145
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   6360
      Top             =   3720
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=G:\医院管理系统\系统管理\settings.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=G:\医院管理系统\系统管理\settings.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "科室"
      Caption         =   "Adodc2"
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
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "医生设置.frx":08CA
      Height          =   390
      Left            =   6000
      TabIndex        =   3
      Top             =   2160
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   688
      _Version        =   393216
      ListField       =   "名称"
      Text            =   "科室"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ComboBox Combo1 
      BeginProperty DataFormat 
         Type            =   5
         Format          =   ""
         HaveTrueFalseNull=   1
         TrueValue       =   "是"
         FalseValue      =   "否"
         NullValue       =   "否"
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   7
      EndProperty
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
      ItemData        =   "医生设置.frx":08F4
      Left            =   6000
      List            =   "医生设置.frx":08FE
      TabIndex        =   4
      Text            =   "是否主任"
      Top             =   2760
      Width           =   2055
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "医生设置.frx":090A
      Height          =   3615
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   6376
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      BackColor       =   65408
      Enabled         =   -1  'True
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
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "编号"
         Caption         =   "编号"
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
         DataField       =   "姓名"
         Caption         =   "姓名"
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
      BeginProperty Column03 
         DataField       =   "科室"
         Caption         =   "科室"
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
      BeginProperty Column04 
         DataField       =   "是否主任"
         Caption         =   "是否主任"
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
            ColumnWidth     =   659.906
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1454.74
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1035.213
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   945.071
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1214.929
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   6360
      Top             =   4080
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
      Connect         =   "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=NURA\SQLEXPRESS"
      OLEDBString     =   "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=NURA\SQLEXPRESS"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "医生资料"
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
   Begin MSForms.Label Label2 
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   4560
      Width           =   2655
      ForeColor       =   255
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "*注:-1表示:是;0表示:否"
      Size            =   "4683;661"
      BorderColor     =   255
      FontName        =   "宋体"
      FontHeight      =   210
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   495
      Left            =   240
      TabIndex        =   11
      Top             =   3840
      Width           =   1335
      Caption         =   "添加"
      Size            =   "2355;873"
      FontName        =   "新宋体"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton CommandButton2 
      Height          =   495
      Left            =   2280
      TabIndex        =   10
      Top             =   3840
      Width           =   1335
      Caption         =   "删除"
      Size            =   "2355;873"
      FontName        =   "新宋体"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton CommandButton3 
      Height          =   495
      Left            =   4320
      TabIndex        =   9
      Top             =   3840
      Width           =   1335
      Caption         =   "保存"
      Size            =   "2355;873"
      FontName        =   "新宋体"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label1 
      Height          =   255
      Index           =   0
      Left            =   6840
      TabIndex        =   7
      Top             =   720
      Width           =   615
      VariousPropertyBits=   8388627
      Caption         =   "姓名"
      Size            =   "1085;450"
      FontName        =   "宋体"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label1 
      Height          =   375
      Index           =   1
      Left            =   6000
      TabIndex        =   6
      Top             =   240
      Width           =   495
      VariousPropertyBits=   8388627
      Caption         =   "编号"
      Size            =   "873;661"
      FontName        =   "宋体"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label1 
      Height          =   375
      Index           =   2
      Left            =   6720
      TabIndex        =   5
      Top             =   1320
      Width           =   735
      VariousPropertyBits=   8388627
      Caption         =   "助记码"
      Size            =   "1296;661"
      FontName        =   "宋体"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox TextBox1 
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   6480
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   240
      Width           =   1575
      VariousPropertyBits=   750798875
      Size            =   "2778;661"
      Value           =   "<自动生成>"
      FontName        =   "宋体"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox TextBox2 
      DataField       =   "姓名"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   6000
      TabIndex        =   1
      Top             =   960
      Width           =   2055
      VariousPropertyBits=   746604571
      Size            =   "3625;661"
      FontName        =   "宋体"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox TextBox3 
      DataField       =   "助记码"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   6000
      TabIndex        =   2
      Top             =   1560
      Width           =   2055
      VariousPropertyBits=   746604571
      Size            =   "3625;661"
      FontName        =   "宋体"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
End
Attribute VB_Name = "医生设置"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mHZtoSM As cHztoSM
Private Sub CommandButton1_Click()
Me.Width = 8205
Adodc1.Recordset.AddNew
TextBox2.SetFocus
CommandButton2.Enabled = True
CommandButton3.Enabled = True
End Sub

Private Sub CommandButton2_Click()
Adodc1.Recordset.Delete
If Adodc1.Recordset.RecordCount = 0 Then
CommandButton2.Enabled = False
End If
End Sub

Private Sub CommandButton3_Click()
Me.CommandButton3.Enabled = True
Me.CommandButton2.Enabled = True
Adodc1.Recordset.Fields("科室") = DataCombo1.Text
If Combo1.Text = "否" Then
Adodc1.Recordset.Fields("是否主任") = False
Else
Adodc1.Recordset.Fields("是否主任") = True
End If
Adodc1.Recordset.Update
Me.Width = 5925
End Sub

Private Sub Form_Load()
Set mHZtoSM = New cHztoSM
    
    mHZtoSM.LoadLibFile App.Path & "\GB2312SM.Lib"
    
    If mHZtoSM.LoadLibSuccess = False Then Unload Me
   Me.Width = 5925
Adodc1.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set mHZtoSM = Nothing
End Sub

Private Sub TextBox3_GotFocus()
TextBox3.Text = mHZtoSM.HZtoSMEx(TextBox2.Text)
End Sub
