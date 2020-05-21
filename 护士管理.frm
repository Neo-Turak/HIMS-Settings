VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form 护士管理 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000D&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "护士管理"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6450
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "护士管理.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   495
      Left            =   4680
      Top             =   3600
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
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
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "护士管理.frx":1082
      DataField       =   "属于部门"
      DataSource      =   "Adodc1"
      Height          =   390
      Left            =   4680
      TabIndex        =   2
      Top             =   3000
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   688
      _Version        =   393216
      ListField       =   "名称"
      Text            =   "请选择部门"
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      DataField       =   "姓名"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   1440
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "编号"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   4920
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "保存"
      Height          =   495
      Left            =   3240
      TabIndex        =   5
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "删除"
      Height          =   495
      Left            =   1680
      TabIndex        =   4
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "添加"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   4200
      Width           =   1335
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "护士管理.frx":1097
      Height          =   3735
      Left            =   120
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   240
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   6588
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
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
         DataField       =   "编号"
         Caption         =   "编号"
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
         DataField       =   "属于部门"
         Caption         =   "属于部门"
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
            ColumnWidth     =   615.118
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1484.787
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   854.929
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1080
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   720
      Top             =   3480
      Width           =   2295
      _ExtentX        =   4048
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
      Connect         =   "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=NURA\SQLEXPRESS"
      OLEDBString     =   "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=NURA\SQLEXPRESS"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   5040
      TabIndex        =   11
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      DataField       =   "助记码"
      DataSource      =   "Adodc1"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4920
      TabIndex        =   10
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "助记码"
      Height          =   375
      Left            =   5040
      TabIndex        =   9
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "姓名"
      Height          =   375
      Index           =   1
      Left            =   5160
      TabIndex        =   8
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "编号"
      Height          =   375
      Index           =   0
      Left            =   5160
      TabIndex        =   7
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "护士管理"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mHZtoSM As cHztoSM
Private Sub Command1_Click()
Dim B As Integer
B = Adodc1.Recordset.Fields(0).Value

Me.Width = 6600
Adodc1.Recordset.AddNew
Text1.Text = B + 1
Text1.SetFocus
End Sub

Private Sub Command2_Click()
Adodc1.Recordset.Delete
If Adodc1.Recordset.RecordCount = 0 Then
Command2.Enabled = False
End If
End Sub

Private Sub Command3_Click()
If Text1.Text <> "" And Text2.Text <> "" And DataCombo1.Text <> "请选择部门" Then
Adodc1.Recordset.Fields(0) = Text1.Text
Adodc1.Recordset.Fields(1) = Text2.Text
Adodc1.Recordset.Fields(2) = Label3.Caption
Adodc1.Recordset.Fields(3) = DataCombo1.Text
Adodc1.Recordset.Update
Me.Width = 4740
Else
MsgBox "请填写必要内容", vbInformation
Text1.SetFocus
End If
End Sub

Private Sub Form_Load()
Set mHZtoSM = New cHztoSM
    
    mHZtoSM.LoadLibFile App.Path & "\GB2312SM.Lib"
    
    If mHZtoSM.LoadLibSuccess = False Then Unload Me
    Me.Width = 4740
  Adodc1.Recordset.MoveLast
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set mHZtoSM = Nothing
End Sub



Private Sub Text1_GotFocus()
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub Text2_LostFocus()
Label3.Caption = mHZtoSM.HZtoSMEx(Text2.Text)
End Sub
