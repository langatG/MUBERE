Attribute VB_Name = "ADO_Module"

'Public Cn As ADODB.connection
'Public CnXls As ADODB.Connection
'Public rst As ADODB.Recordset
'Dim OpenRS As New ADODB.Recordset

Public Function OpenMDB() As Long
    Dim strCon As String
    Dim strBuffer As String
    On Error GoTo ConnectError
    Set cn = Nothing
    strBuffer = SelectedDsn
    'strBuffer = IMPORTINFO.accessfile
    'strCon = strBuffer & ";Jet OLEDB:Database Password=" & IMPORTINFO.accesspassword
    Set cn = New ADODB.Connection
    'Cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strCon
    cn.Open strBuffer
    OpenMDB = 0
    Exit Function
ConnectError:
    OpenMDB = Err.number
End Function

Public Function OpenXLS(sFile As String, ErrorMsg As String) As Long
      On Error GoTo fix_err
      Dim sconn As String
      Set CnXls = Nothing
      sconn = "DRIVER=Microsoft Excel Driver (*.xls);" & "DBQ=" & sFile
      Set CnXls = New ADODB.Connection
      CnXls.ConnectionString = sconn
      CnXls.Open 'sconn
      OpenXLS = 0
      Exit Function
fix_err:
      OpenXLS = Err.number
      ErrorMsg = Err.Description
End Function


Public Sub ExecuteQuery(qstring As String)
    cn.Execute qstring
End Sub

Public Function OpenRS(strSQL As String) As ADODB.Recordset
Dim rs As ADODB.Recordset
        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.CursorType = adOpenDynamic
        If IsNull(cn) = False Then
            rs.Open strSQL, cn, adOpenKeyset, adLockOptimistic
        End If
        Set OpenRS = rs
End Function


Public Function OpenRSXLS(strSQL As String) As ADODB.Recordset
Dim rs As ADODB.Recordset
        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.CursorType = adOpenDynamic
        If IsNull(CnXls) = False Then
            rs.Open strSQL, CnXls, adOpenKeyset, adLockPessimistic
        End If
        Set OpenRSXLS = rs
End Function

Public Function QuoteReplace(s As String) As String
Dim tmpstr As String
    'find if the string contains qoutes
    If InStr(s, "'") Then
        tmpstr = Replace(s, "'", "\'")
        QuoteReplace = tmpstr
    ElseIf InStr(s, "\") Then
        tmpstr = Replace(s, "\", "\\")
        QuoteReplace = tmpstr
    Else
        QuoteReplace = s
    End If
End Function
Public Function ClsSql(str As String) As String
    Dim tmpstr As String
    Dim s As String
    s = str
    tmpstr = Replace(s, "'", "''")
    'tmpstr = Replace(s, "\", "\\")
    ClsSql = tmpstr
End Function

