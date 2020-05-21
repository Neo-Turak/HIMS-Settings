VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   0  'None
   Caption         =   "登录"
   ClientHeight    =   2655
   ClientLeft      =   2790
   ClientTop       =   3105
   ClientWidth     =   5025
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin.frx":08CA
   ScaleHeight     =   1568.662
   ScaleMode       =   0  'User
   ScaleWidth      =   4718.203
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1680
      TabIndex        =   0
      Top             =   720
      Width           =   2445
   End
   Begin VB.TextBox txtPassword 
      Height          =   344
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   1800
      Width           =   1455
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
'用于将CreateRoundRectRgn创建的圆角区域赋给窗体
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
'用于创建一个圆角矩形，该矩形由X1，Y1-X2，Y2确定，并由X3，Y3确定的椭圆描述圆角弧度。
'参数 类型及说明：
'X1,Y1 Long，矩形左上角的X，Y坐标
'X2,Y2 Long，矩形右下角的X，Y坐标
'X3 Long，圆角椭圆的宽。其范围从0（没有圆角）到矩形宽（全圆）
'Y3 Long，圆角椭圆的高。其范围从0（没有圆角）到矩形高（全圆）
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Dim outrgn As Long
Option Explicit

Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    '设置全局变量为 false
    '不提示失败的登录
    LoginSucceeded = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    '检查正确的密码
    If txtPassword = "password" Then
        '将代码放在这里传递
        '成功到 calling 函数
        '设置全局变量时最容易的
        LoginSucceeded = True
        Me.Hide
    Else
        MsgBox "无效的密码，请重试!", , "登录"
        txtPassword.SetFocus
       
    End If
End Sub

Private Sub Form_Activate()
Call rgnform(Me, 50, 50)
End Sub
Private Sub rgnform(ByVal frmbox As Form, ByVal fw As Long, ByVal fh As Long) '子过程，改变参数fw和fh的值可实现圆角
Dim w As Long, h As Long
w = frmbox.ScaleX(frmbox.Width, vbTwips, vbPixels)
h = frmbox.ScaleY(frmbox.Height, vbTwips, vbPixels)
outrgn = CreateRoundRectRgn(0, 0, w, h, fw, fh)
Call SetWindowRgn(frmbox.hWnd, outrgn, True)
End Sub

Private Sub Form_Unload(Cancel As Integer) '窗体Unload事件
DeleteObject outrgn '将圆角区域使用的所有系统资源释放
End Sub
End Sub

Private Sub Label1_Click()
Call cmdOK_Click
End Sub

Private Sub Label2_Click()
LoginSucceeded = False
    Me.Hide
End Sub
