Attribute VB_Name = "数据库连接"
Public cn As New ADODB.Connection
Public a As String
Public b As String
Public c As String
Public d As String

Public Function 连接() As String
cn.ConnectionString = "Provider=SQLOLEDB.1;User ID=" & c & ";PWD=" & d & ";Initial Catalog=" & b & ";Data Source=" & a & ""

Dim Con As ADODB.Connection
Dim Mrc As ADODB.Recordset
Set Con = New ADODB.Connection
Set Mrc = New ADODB.Recordset
Dim SQL As String
SQL = "Provider=sqloledb.1;Data Source=NURA\SQLEXPRESS;Persist Security Info=true;user id=sa;password=sa;initial catalog=ghgl"
Con.Open SQL
Con.CursorLocation = adUseClient
'Mrc.Open "select * from 检查项目 where 所属科室='" & DataCombo1.Text & "'", Con, adOpenKeyset, adLockOptimistic
'Set DataCombo3.RowSource = Mrc
'SQL = "Insert into testtable(sn,sname,sex) Values('123','ABC','男')"
   ' cn.Execute SQL
'mrc.Close
End Function

