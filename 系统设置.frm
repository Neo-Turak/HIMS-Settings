VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl32.ocx"
Begin VB.MDIForm ϵͳ���� 
   BackColor       =   &H8000000C&
   Caption         =   "ϵͳ����"
   ClientHeight    =   7185
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   11715
   Icon            =   "ϵͳ����.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "ϵͳ����.frx":1082
   StartUpPosition =   1  '����������
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
            Object.ToolTipText     =   "��ǰ����"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "12:51"
            Object.ToolTipText     =   "��ǰʱ��"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
            Object.ToolTipText     =   "��ǰ�û�"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2469
            MinWidth        =   2469
            Object.ToolTipText     =   "ְλ"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   6350
            MinWidth        =   6350
            Text            =   "Ϊ��һ�в���,Ϊ�˲���һ��"
            TextSave        =   "Ϊ��һ�в���,Ϊ�˲���һ��"
            Object.ToolTipText     =   "ҽԺ����"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "������κ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu �������� 
      Caption         =   "��������"
      Index           =   1
      NegotiatePosition=   1  'Left
      WindowList      =   -1  'True
      Begin VB.Menu ���� 
         Caption         =   "��������"
         Index           =   2
         Shortcut        =   ^B
      End
      Begin VB.Menu ��λ 
         Caption         =   "��λ����"
         Index           =   3
      End
      Begin VB.Menu ���� 
         Caption         =   "���ҹ���"
         Index           =   4
      End
      Begin VB.Menu ҽ�� 
         Caption         =   "ҽ������"
         Index           =   5
      End
      Begin VB.Menu ��ʿ 
         Caption         =   "��ʿ����"
         Index           =   6
      End
      Begin VB.Menu ҩ�� 
         Caption         =   "ҩ������"
         Index           =   7
      End
      Begin VB.Menu ��Ŀ 
         Caption         =   "�����Ŀ����"
         Index           =   8
      End
   End
   Begin VB.Menu �û� 
      Caption         =   "�û�����"
      Index           =   10
   End
   Begin VB.Menu ��Ŀ��ѯ 
      Caption         =   "��Ŀ��ѯ"
      Begin VB.Menu ���վ����¼ 
         Caption         =   "���վ����¼"
      End
      Begin VB.Menu ����סԺ��¼ 
         Caption         =   "����סԺ��¼"
      End
      Begin VB.Menu �����տ��¼ 
         Caption         =   "�����տ��¼"
      End
      Begin VB.Menu סԺ����ͳ�� 
         Caption         =   "סԺ����ͳ��"
      End
   End
End
Attribute VB_Name = "ϵͳ����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Timer1_Timer()
StatusBar1.Panels(2) = Time
StatusBar1.Panels(1) = Date
End Sub

Private Sub ����_Click(Index As Integer)
��������.Show
End Sub

Private Sub ��λ_Click(Index As Integer)
��λ����.Show
End Sub

Private Sub ��ʿ_Click(Index As Integer)
��ʿ����.Show
End Sub

Private Sub ���վ����¼_Click()
ϵͳ.���վ����¼.Show
End Sub

Private Sub ����_Click(Index As Integer)
��������.Show
End Sub

Private Sub ��Ŀ_Click(Index As Integer)
��������Ŀ����.Show
End Sub

Private Sub ҩ��_Click(Index As Integer)
ҩ������.Show
End Sub

Private Sub ҽ��_Click(Index As Integer)
ҽ������.Show
End Sub

Private Sub �û�_Click(Index As Integer)
�û�����.Show
End Sub
