VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSaccoData"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public goConn As New Connection
Public Event ErrOccured()
Public gINIFile As String
Private objRS As New Recordset
Private Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Function GetINIString(strINIFile As String, strSection As String, strKey As String, _
                              strDefault As String) As String

End Function

'This public method is used to write a Key value to the INI file
Public Sub WriteINIString(strINIFile As String, strSection As String, strKey As String, strValue As String)
    Dim indx As Integer
    Dim strTemp As String
    strTemp = strValue
    For indx = 1 To Len(strValue)
        If mid$(strValue, indx, 1) = vbCr Or mid$(strValue, indx, 1) = vbLf Then
            Mid$(strValue, indx) = " "
        End If
    Next indx
    indx = WritePrivateProfileString(strSection, strKey, strTemp, strINIFile)
End Sub

Private Sub Class_Initialize()
    'gINIFile = App.Path & "\Sacco.INI"
    With goConn
        .CursorLocation = adUseClient
        ConnectDatabase
    End With
End Sub

Public Function OpenSQLSVR() As Recordset
    GetRecordset ("exec sp_databases")
    Set OpenSQLSVR = objRS
End Function

Public Function GetConString() As String 'Assemble the ConnectionString and return it
    On Error GoTo SysError
SysError:
    
End Function

Public Function ConnectDatabase() As Boolean
    On Error GoTo ErrTrap
    If goConn.State = adStateOpen Then
        goConn.Close
    End If
    'SelectedDsn = Selecteddsn
    If SelectedDsn = "" Then
        SelectedDsn = "MAZIWA"
    End If
    goConn.Open SelectedDsn, "atm", "atm"
    ConnectDatabase = True
    Exit Function
ErrTrap:
    MsgBox err.description, vbOKOnly, "Error Connecting to database"
    ConnectDatabase = False
End Function

Public Function GetRecordset(strSQL As String) As Recordset
    On Error GoTo ErrTrap
    Dim rsRecordset As New Recordset
    'Class_Initialize
    With rsRecordset
        .ActiveConnection = goConn
        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        success = True
        .Open strSQL
    End With
    Set GetRecordset = rsRecordset
    Exit Function
ErrTrap:
    success = False
    MsgBox err.description, vbOKOnly, "Connection error"
End Function


Public Sub ExecuteThis(strSQL As String)
On Error GoTo ErrTrap
Dim cmdCommand As New Command
With cmdCommand
    .ActiveConnection = goConn
    .CommandType = adCmdText
    .CommandText = strSQL
    .Execute
End With
Exit Sub
ErrTrap:
  ErrorMessage = ""
  MsgBox err.description, vbOKOnly, "Connection error"
  ErrorMessage = err.description
End Sub
Public Function Execute(strSQL As String) As Boolean
 On Error GoTo ErrTrap
    ErrorMessage = ""
    Dim cmdCommand As New Command
    With cmdCommand
        .ActiveConnection = goConn
        .CommandType = adCmdText
        .CommandText = strSQL
        .Execute
        Execute = True
    End With
    Exit Function
ErrTrap:
    Execute = False
    ErrorMessage = err.description
'On Error GoTo ErrTrap
'Dim cmdCommand As New Command
'With cmdCommand
'    .ActiveConnection = goConn
'    .CommandType = adCmdText
'    .CommandText = strSQL
'    .Execute
'    Execute = True
'End With
'Exit Function
'ErrTrap:
'    Execute = False
'  ErrorMessage = ""
'  MsgBox err.description, vbOKOnly, "Connection error"
'  ErrorMessage = err.description
End Function
