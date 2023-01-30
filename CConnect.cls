VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 2  'vbComplexBound
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public SelectedDsn As String
Public cnnPayroll As ADODB.Connection
Private Sub Class_Initialize()
On Error Resume Next
Provider = "MAZIWA"
Set cnnPayroll = New Connection
'cnnPayroll.Open MDBase
   cnnPayroll.Open Provider, "atm"
    cnnConnect
End Sub

Public Sub cnnConnect()
On Error Resume Next
Provider = "MAZIWA"
Set cnnPayroll = New Connection
'cnnPayroll.Open MDBase
   cnnPayroll.Open Provider, "atm"
End Sub

Public Sub Employees()
On Error Resume Next

Set rsEmployee = New Recordset
Set cnnPayroll = New ADODB.Connection
cnnPayroll.Open modCommon.pConnection
rsEmployee.Open "Select * from Employee order by PayrollNo", cnnPayroll, adOpenKeyset, adLockOptimistic
    
End Sub

Public Sub Kin()
On Error Resume Next

Set RSKIN = New Recordset
Set cnnPayroll = New ADODB.Connection
cnnPayroll.Open modCommon.pConnection
RSKIN.Open "Select * from Kin", cnnPayroll, adOpenKeyset, adLockOptimistic
    
End Sub

Public Sub Payslip()
On Error Resume Next

Set rsPayslip = New Recordset
Set cnnPayroll = New ADODB.Connection
cnnPayroll.Open modCommon.pConnection
rsPayslip.Open "Select * from Payslip", cnnPayroll, adOpenKeyset, adLockOptimistic
    
End Sub

Public Sub Openrs()
Provider = "MAZIWA"
On Error GoTo ErrorHandler
Set rs = New ADODB.Recordset
Set cnnPayroll = New ADODB.Connection
cnnPayroll.Open Provider, "atm", "atm"
rs.Open sql, cnnPayroll, adOpenKeyset, adLockOptimistic
Exit Sub
ErrorHandler:
MsgBox "Data Base Connection Fails", vbCritical, "Data Base Connection fails"
'frmSaveSettings.Show vbModal
End Sub

Public Sub Openrs2()
On Error Resume Next

Set rs2 = New Recordset
Set cnnPayroll = New ADODB.Connection
cnnPayroll.Open modCommon.pConnection
rs2.Open sql, cnnPayroll, adOpenKeyset, adLockOptimistic

End Sub

Public Sub Openrs3()
On Error Resume Next

Set rs3 = New Recordset
Set cnnPayroll = New ADODB.Connection
cnnPayroll.Open modCommon.pConnection
rs3.Open sql, cnnPayroll, adOpenKeyset, adLockOptimistic

End Sub

Public Sub Openrs4()
On Error Resume Next

Set rs4 = New Recordset
Set cnnPayroll = New ADODB.Connection
cnnPayroll.Open modCommon.pConnection
rs4.Open sql, cnnPayroll, adOpenKeyset, adLockOptimistic

End Sub

Public Sub Openrs5()
On Error Resume Next

Set rs5 = New Recordset
Set cnnPayroll = New ADODB.Connection
cnnPayroll.Open modCommon.pConnection
rs5.Open sql, cnnPayroll, adOpenKeyset, adLockOptimistic

End Sub
Public Sub Company()
On Error GoTo ErrorHandler

Set rscompany = New Recordset
Set cnnPayroll = New ADODB.Connection
cnnPayroll.Open modCommon.pConnection
rscompany.Open "Select * from Company", cnnPayroll, adOpenKeyset, adLockOptimistic
Exit Sub
ErrorHandler:
MsgBox err.description
End Sub
Public Sub Openrs1()
On Error Resume Next

Set Rs1 = New Recordset
Set cnnPayroll = New ADODB.Connection
cnnPayroll.Open modCommon.pConnection
Rs1.Open sql, cnnPayroll, adOpenKeyset, adLockOptimistic
    
End Sub

Public Sub Crs()
Set rs = Nothing

End Sub

Public Sub Crs1()
Set Rs1 = Nothing

End Sub

Public Sub CKin()
    Set RSKIN = Nothing
End Sub

Public Sub CEmployees()
    Set rsEmployee = Nothing
End Sub

Public Sub cnnPayrollClose()
  Set cnnPayroll = Nothing
  
End Sub


Public Sub LoadSearch()
    With frmSearch.lstSearch
        .ListItems.Clear
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "Payroll No", 1500
        .ColumnHeaders.Add , , "Surname", 1700
        .ColumnHeaders.Add , , "Other Names", 2800
        .ColumnHeaders.Add , , "ID No", 1500
        .View = lvwReport
        .GridLines = True
    End With
    
    With frmSearch.cboField
        .AddItem "PayrollNo"
        .AddItem "Surname"
        .AddItem "OtherNames"
        .AddItem "IDNo"
    End With
    
'    CConnect.cnnConnect
'    sql = "Select * from Employee order by PayrollNo"
'    CConnect.Openrs
    
    With rs
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                Set li = frmSearch.lstSearch.ListItems.Add(, , !payrollno)
                li.SubItems(1) = !surname & ""
                li.SubItems(2) = !OtherNames & ""
                li.SubItems(3) = !idno & ""
                .MoveNext
                
            Loop
            
        End If
        .Close
    End With
    
'    Set rs = Nothing
'    Set cnnPayroll = Nothing
        

End Sub

Public Function GetRecordset(strSQL As String) As Recordset
    Dim rsRecordset As New Recordset
    Set cnnPayroll = New ADODB.Connection
    cnnPayroll.Open modCommon.pConnection
    rsRecordset.Open strSQL, cnnPayroll, adOpenKeyset, adLockOptimistic
    Set GetRecordset = rsRecordset
    
End Function

