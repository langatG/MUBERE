Attribute VB_Name = "General"
Option Explicit
Public CosCon As New ADODB.Connection
'Public success As Boolean
Public Type Account_Details
    AccountNo As String
    AccountName As String
    OpeningBalance As Double
    currentbalance As Double
    NormalBalance As String
End Type
Public Function GetRecords(str As String)
    On Error GoTo Capture
    Set rst = New ADODB.Recordset
    sql = mysql
    Set cn = CreateObject("adodb.connection")
    Provider = "MAZIWA"
   cn.Open Provider, "atm", "atm"
    rst.Open str, cn
    success = True
    Exit Function
Capture:
    success = False
    MsgBox err.description
End Function

'Public Sub InitSubClass()
'Set colClass = New Collection
'End Sub
Public Function Get_Account_Details(ACCNO As String, DataSource As String, _
errmsg As String) As Account_Details
    On Error GoTo SysError
    Dim rsAccounts As New Recordset
    ''Open_Database DataSource
    Set rsAccounts = oSaccoMaster.GetRecordset("Select * From GLSETUP Where AccNo='" & ACCNO & "'")
    With rsAccounts
        If .State = adStateOpen Then
            If Not .EOF Then
                Get_Account_Details.AccountName = IIf(IsNull(!GlAccName), "", !GlAccName)
                Get_Account_Details.AccountNo = ACCNO
                Get_Account_Details.currentbalance = IIf(IsNull(!CurrentBal), 0, !CurrentBal)
                Get_Account_Details.NormalBalance = IIf(IsNull(!NormalBal), "DR", IIf(!NormalBal <> "Credit", "DR", "CR"))
                Get_Account_Details.OpeningBalance = IIf(IsNull(!OpeningBal), 0, !OpeningBal)
            End If
        End If
    End With
    Exit Function
SysError:
    errmsg = err.description
    Get_Account_Details.AccountNo = ""
End Function




