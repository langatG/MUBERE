VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cdbase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim myclass As New cdbase
Public Provider As String
Dim powertoproceed As Boolean
'Public Cn As Object
Dim mvarConnection As String
Dim res
Public Event Connection(bProvider As Boolean)
Public Sub PositionForm(frm As Form)
    
    If Not frm.WindowState = vbNormal Then
        frm.WindowState = vbNormal
    End If
    
End Sub
Sub saveToFOSa(mysql As String)
'On Error GoTo errorhandler
    sql = mysql
    Set cn = CreateObject("adodb.connection")
    Provider = "MAZIWA"
   cn.Open Provider, "atm", "atm"
      cn.Execute sql
    myclass.CloseCon
    Exit Sub
ErrorHandler:
MsgBox err.description
End Sub
Public Sub cnnConnect()
    On Error Resume Next
    Set cn = New Connection
    Set rs = CreateObject("adodb.recordset")
    Set cn = CreateObject("adodb.connection")
    Provider = myclass.OpenCon
   cn.Open Provider, "atm", "atm"
    'cnnPayroll.Open MDBase
End Sub

Public Sub Openrs()
    On Error GoTo ErrorHandler
    Set cn = New Connection
    Set rs = CreateObject("adodb.recordset")
    Set cn = CreateObject("adodb.connection")
    Provider = "MAZIWA"
   cn.Open Provider, "atm", "atm"
    'Set rs = New ADODB.Recordset
    'Set cnnPayroll = New ADODB.connection
    rs.Open sql, cn, adOpenKeyset, adLockOptimistic
    Exit Sub
ErrorHandler:
    MsgBox "Data Base Connection Fails", vbCritical, "Data Base Connection fails"
    'frmSaveSettings.Show vbModal
End Sub

Public Function checkvoucherno(ByVal strvou As String) As Boolean
Dim cn As Connection
Set cn = CreateObject("adodb.connection")
Provider = myclass.OpenCon
cn.Open Provider, "atm", "atm"
sql = ""
sql = "select vno from DailyBalance where vno='" & strvou & "'"
Set rs = CreateObject("adodb.recordset")
rs.Open sql, cn
If rs.EOF Then checkvoucherno = True: Exit Function
If Not rs.EOF Then
checkvoucherno = False
Exit Function
End If
End Function
Public Function validacc(ByVal strAccNo As String) As Boolean

Set cn = CreateObject("adodb.connection")
Provider = myclass.OpenCon
cn.Open Provider, "atm", "atm"
sql = ""
sql = "select accountnumber from CustomerAccount where accountnumber='" & strAccNo & "'"
Set rs = CreateObject("adodb.recordset")
rs.Open sql, cn
If rs.EOF Then validacc = False: Exit Function

validacc = True

End Function
Public Function cantrsact3(ByVal strmemno As String) As Boolean
Dim cn As Object
Set cn = CreateObject("adodb.connection")
Provider = myclass.OpenCon
cn.Open Provider, "atm", "atm"
sql = ""
sql = "select Closed,Frozen from Cub where memberno='" & strmemno & "'"
Set rs = CreateObject("adodb.recordset")
rs.Open sql, cn
If rs.EOF Then cantrsact3 = True: Exit Function
If rs.Fields(0) = True Then
    cantrsact3 = False
    Exit Function
End If
If rs.Fields(1) = True Then
    cantrsact3 = False
    Exit Function
End If

cantrsact3 = True

End Function
Public Function cantrasact2(ByVal PAYNO As String) As Boolean
Set cn = CreateObject("adodb.connection")
Provider = myclass.OpenCon
cn.Open Provider, "atm", "atm"
sql = ""
sql = "select Closed,Frozen from Cub where payno='" & PAYNO & "'"
Set rs = CreateObject("adodb.recordset")
rs.Open sql, cn
If rs.EOF Then cantrasact2 = True: Exit Function
If rs.Fields(0) = True Then
    cantrasact2 = False
    Exit Function
End If
If rs.Fields(1) = True Then
    cantrasact2 = False
    Exit Function
End If

cantrasact2 = True

End Function
Public Function CanTransact1(ByVal stridno As String) As Boolean
Set cn = CreateObject("adodb.connection")
Provider = myclass.OpenCon
cn.Open Provider, "atm", "atm"
sql = ""
sql = "select Closed,Frozen from Cub where idno='" & stridno & "'"
Set rs = CreateObject("adodb.recordset")
rs.Open sql, cn
If rs.EOF Then CanTransact1 = True: Exit Function
If rs.Fields(0) = True Then
    CanTransact1 = False
    Exit Function
End If
If rs.Fields(1) = True Then
    CanTransact1 = False
    Exit Function
End If

CanTransact1 = True
  
End Function
Public Function CanTransact(ByVal strCustNo As String) As Boolean

Set cn = CreateObject("adodb.connection")
Provider = myclass.OpenCon
cn.Open Provider, "atm", "atm"
sql = ""
sql = "select Closed,Frozen from Cub where accno='" & strCustNo & "'"
Set rs = CreateObject("adodb.recordset")
rs.Open sql, cn
If rs.EOF Then CanTransact = True: Exit Function
If rs.Fields(0) = True Then
    CanTransact = False
    Exit Function
End If
If rs.Fields(1) = True Then
    CanTransact = False
    Exit Function
End If

CanTransact = True

End Function

Public Function GetUsers(ByVal strPassword As String, ByVal strUserName As String) As String
    On Error GoTo ErrorHandler
'    Dim Pass As EncryptDecrypt
'    Set Pass = New EncryptDecrypt
    Set cn = CreateObject("adodb.connection")
    Provider = myclass.OpenCon
    cn.Open Provider, "atm", "atm"
'    strPassword = Pass.Encrypt(strPassword)
    strPassword = modsecurity.Decript_String(strPassword)
    sql = "select * from useraccounts where userloginids='" & strUserName & _
    "' and password='" & strPassword & "'"
    Set rs = CreateObject("adodb.recordset")
    rs.Open sql, cn
    Dim hakuna As Recordset
    sql = "select * from useraccounts"
    Set hakuna = New ADODB.Recordset
    hakuna.Open sql, cn
    If hakuna.EOF Then
        strUserName = "ADMIN"
        strPassword = "admin"
        strPassword = modsecurity.Decript_String(strPassword)
        'by default let it insert to the table the following
        sql = "insert into useraccounts (username,userloginids,password)select'ADMIN','" & strUserName & "','" & strPassword & "'"
        myclass.save sql
        '// REOPEN THE DATA BASE AND REFRESH THE DATABASE
        sql = "select * from useraccounts where userloginids='" & strUserName & _
        "' and password='" & strPassword & "'"
        Set rs = CreateObject("adodb.recordset")
        rs.Open sql, cn
        GetUsers = rs!username
        SaveSetting App.title, "Teller", "Name", rs!username
    Else
        sql = "select * from useraccounts where userloginids='" & strUserName & _
        "' and password='" & strPassword & "'"
        Set rs = CreateObject("adodb.recordset")
        rs.Open sql, cn
        If rs.EOF Then
            MsgBox "Your Password is incorrect. Please seek assistance or try again"
            Exit Function
        Else
            ' If strPassword = rs.Fields("password") And strUserName = rs.Fields("username") Then
            GoTo goahead
            AuditName = IIf(IsNull(rs!username), "", rs!username)
            ' End If
        End If
goahead:
        GetUsers = rs!username
        SaveSetting App.title, "Teller", "Name", rs!UserLoginIDs
    End If
    myclass.CloseCon
    Set myclass = Nothing
    Exit Function
ErrorHandler:
    'MsgBox "Check the data base connection", vbCritical, "Data base"
    MsgBox err.description, vbCritical, "Data base"
End Function

Sub Delete(ssql As String, Optional bMsgBox As Boolean)

    sql = ssql

    Set cn = CreateObject("adodb.connection")

    Provider = myclass.OpenCon

   cn.Open Provider, "atm", "atm"

'If bMsgBox = False Then GoTo Skip

    res = MsgBox("Delete Record", vbYesNo + vbQuestion): If res = vbNo Then Exit Sub
    
    If res = vbYes Then

'Skip:
    cn.Execute sql
    Else
    MsgBox "recorde Not Deleted"
    End If

    myclass.CloseCon

End Sub

Sub save(ssql As String)
On Error GoTo ErrorHandler
    sql = ssql
    Set cn = CreateObject("adodb.connection")
    If Provider = "" Then Provider = myclass.OpenCon
   cn.Open Provider, "atm", "atm"
      cn.Execute sql
    myclass.CloseCon
    Exit Sub
ErrorHandler:
MsgBox err.description
End Sub

Public Function OpenCon() As String
    On Error GoTo Reconnect
    Connect
    Set cn = CreateObject("adodb.connection")
Reconnect:
    OpenCon = mvarConnection
End Function

Private Sub Connect()
'kibet
    mvarConnection = "MAZIWA"
End Sub

Public Sub CloseCon()
    
   'Set cn = Nothing

End Sub


