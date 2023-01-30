Attribute VB_Name = "modCommon"
Public strConn As String, Sys As String
'Public SelectedDsn As String
Public oConnectionString As String, strName As String, strSQL As String, genStr As String, strCode As Double
Public reportpath As String
Public User As String
Function Get_Server_Date1() As Date
    On Error GoTo 10
    Dim rs As Recordset
    Set myclass = New cdbase
   Provider = myclass.OpenCon
    cn.Open Provider
    Set rs = New Recordset
    rs.Open "SET DATEFORMAT DMY select GetDate()", cn, adOpenStatic, adLockOptimistic
    Get_Server_Date1 = rs(0)
    getdate = Get_Server_Date1
    rs.Close
    Exit Function
10:    MsgBox Err.description
End Function
Public Function pConnection()
On Error GoTo errorHandler
If oConnectionString = "" Then oConnectionString = cnn
pConnection = cnn
pConnection = oConnectionString
If oConnectionString = "" Then oConnectionString = cnn
If Sys = "" Then Sys = App.EXEName '
 Exit Function
errorHandler:
MsgBox "Data Base connection fails", vbCritical
frmSaveSettings.Show vbModal
End Function
Public Function rptPath()
On Error GoTo errorHandler
Set rst = New Recordset
Provider = cnn
'Dim cn As Connection
Set cn = New ADODB.connection
cn.Open Provider, , ""
'If reportpath = "" Then reportpath = GetSetting("payroll", "AppName", "rptPath", rptPath)
sql = "select path from reportpath"
Set rs = New ADODB.Recordset
rs.Open sql, cn
If Not rs.EOF Then
rptPath = rs.Fields(0)
End If
rptPath = rptPath
If Sys = "" Then Sys = App.EXEName '
Exit Function
errorHandler:
MsgBox "Data Base connection fails", vbCritical
frmSaveSettings.Show vbModal
End Function
