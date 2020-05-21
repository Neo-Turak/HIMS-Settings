VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHztoSM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Download by http://www.NewXing.com
Option Explicit

Const ERR_RESULT$ = "?"          ' 函数的错误返回值

Private mGB2312SM$
Private mLoadLibSuccess As Boolean

Private Sub Class_Initialize()
    mGB2312SM$ = ""
    mLoadLibSuccess = False
End Sub

Private Sub Class_Terminate()
    mGB2312SM$ = ""
    mLoadLibSuccess = False
End Sub


'==================================================
' 加载库文件 成功 LoadLibSuccess = True
'            失败 LoadLibSuccess = False
Public Property Get LoadLibSuccess() As Boolean
    LoadLibSuccess = mLoadLibSuccess
End Property
'==================================================


'==================================================
' 方法: LoadLibFile
'
' 功能: 加载库文件
'
' 注意: 将设置加载成功标志 mLoadLibSuccess
'
' 入口: LibFileName     库文件名
'
Public Sub LoadLibFile(ByVal LibFileName$)
    Dim FileNum&
    Dim tmpText$
    
    On Error GoTo ErrLoad:
    FileNum& = FreeFile
    
    Open LibFileName$ For Input As #FileNum
    
    ' 顺序读取库文件,保存到变量 mGB2312SM$ 中
    Do While Not EOF(1)
       Line Input #FileNum, tmpText$
       mGB2312SM$ = mGB2312SM$ & tmpText$
    Loop
    Close #FileNum
    
    ' 加载库文件成功
    mLoadLibSuccess = True
    
    Exit Sub
ErrLoad:
    MsgBox "加载库文件 " & LibFileName$ & " 失败！", vbExclamation, "来自类 cHZtoSM 的错误"
    mGB2312SM$ = ""
    mLoadLibSuccess = False
End Sub
'==================================================


'==================================================
' 函数: HZtoSM
'
' 功能: 返回字符串中第一个字符的声母
'
' 注意: 该函数能处理所有汉字,但需要库文件的支持
'       若待处理的字符并不在库文件中,则函数将返回常数 ERR_RESULT$
'
' 入口: Str     待处理的字符串
'
Public Function HZtoSM$(ByVal Str$)
    Dim tmpStr$, sAscii$
    Dim lAsciiU&, lAsciiL&
    Dim fPos&
    
    ' 取出字符串中的第一个字符
    tmpStr$ = Left(Str$, 1)
    
    ' 若tmpStr长度为 0 ,则函数无返回值
    If Len(tmpStr$) <= 0 Then Exit Function
    
    ' 返回字符型 ASCII 码
    sAscii$ = Hex(Asc(tmpStr$))

    If Len(sAscii$) <> 4 Then GoTo ErrChg:
    
    ' 取出字符高字节和低字节
    lAsciiU& = Val("&H" & Left(sAscii$, 2))
    lAsciiL& = Val("&H" & Right(sAscii$, 2))
    
    ' 公式: ( 高两位 - &H81 ) * ( 16 * 12 - 1 ) +
    '       ( 低两位 - &H40 + 1 )
    ' 即可计算出声母对应的位置
    fPos& = (lAsciiU - &H81) * 191 + (lAsciiL& - 63)
    If fPos& < 0 Or fPos& > Len(mGB2312SM) Then GoTo ErrChg:
    
    ' 在库文件查找对应的声母
    HZtoSM$ = UCase(Mid(mGB2312SM, fPos&, 1))
    If (Asc(HZtoSM$) < &H41) Or (Asc(HZtoSM$) > &H90) Then GoTo ErrChg:
    
    Exit Function
ErrChg:
    ' 函数转换错误,返回常数 ERR_RESULT
    HZtoSM$ = ERR_RESULT$
End Function
'==================================================


'==================================================
' 函数: HZtoSMEx
'
' 功能: 返回字符串中每个字符的声母,由函数 HZtoSM 扩展来
'
' 入口: Str     待处理的字符串
'
Public Function HZtoSMEx$(ByVal Str$)
    Dim nPos&
    
    For nPos& = 1 To Len(Str$)
        HZtoSMEx$ = HZtoSMEx$ & HZtoSM(Mid(Str$, nPos&, 1))
    Next nPos&
End Function
'==================================================
