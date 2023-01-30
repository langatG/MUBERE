VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GeneralLedger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Public Cn As Connection



Private Function Get_Records(strSQL As String, DataSource As String, errmsg As String) As Recordset
    Open_Database DataSource
    Dim Provider As String
    Provider = "maziwa"
    Set cn = New ADODB.Connection
   cn.Open Provider, "atm", "atm"
    Set Get_Records = cn.Execute(strSQL)
End Function

Private Function Save_TrialBalance(ACCNO As String, AccName As String, amount As _
Double, transtype As String, Closed As Long, transdate As Date, auditid As _
String, accType As String, AccGroup As String, Budget As Double, DataSource As _
String, errmsg As String) As Boolean
    On Error GoTo SysError
    With cn
        If .State = adStateClosed Then
            .Open DataSource, "atm"
        End If
        .Execute ("Set DateFormat DMY Exec Save_TrialBalance '" & ACCNO & "','" & _
        AccName & "'," & amount & ",'" & transtype & "'," & Closed & ",'" & transdate & _
        "','" & auditid & "','" & accType & "','" & AccGroup & "'," & Budget)
    End With
    Save_TrialBalance = True
    Exit Function
SysError:
    Save_TrialBalance = False
    MsgBox err.description
End Function

Private Function Print_Trial_Balance(Startdate As Date, FinishDate As Date, DataSource As String, _
errmsg As String) As Boolean
    On Error GoTo SysError
    Dim rsTBBalance As New Recordset, transtype As String, Account As Account_Details, _
    balance As Double, rsAccounts As New Recordset, ACCNO As String, NormalBal As String, _
    OpTransType As String, AccName As String, cnn As New ADODB.Connection
    'XXXXXXXXXXXXXXXXXXXXX Open The Database XXXXXXXXXXXXXXXXXX
    Open_Database DataSource
    Dim mysql As String
    Dim AccCategory As String
    Dim Provider As String
    Provider = "MAZIWA"
    cnn.Open Provider, "atm"
    
    sql = ("Delete From TBBALANCE")
    oSaccoMaster.ExecuteThis (sql)
    'XXXXXXXXXXXXXXXXXXXXX Get_The_Accounts_And_Balances from GLSetUp XXXXXXXXXXXXX
    Set rsAccounts = cnn.Execute("set dateformat dmy Select * From GLSETUP where NewGLOpeningBalDate >='" & Startdate & "' Order By AccNo")
    If Not rsAccounts.EOF Then
    With rsAccounts
    
        If .State = adStateOpen Then
            While Not .EOF
                DoEvents
                ACCNO = IIf(IsNull(!ACCNO), "", !ACCNO)
                Account = Get_Account_Details(ACCNO, DataSource, errmsg)
                AccCategory = IIf(IsNull(!AccCategory), "", !AccCategory)
                If Account.AccountNo <> "" Then
                    'Balance = Account.OpeningBalance
                    balance = !NewGLOpeningBal
                    NormalBal = Account.NormalBalance
                    AccName = Account.AccountName
                    'XXXXXXXXXXXXXX Get Transactions From Temp TBBalance XXXXXXXXXXXXXXXX
                    mysql = "Select TransType,Sum(Amount) as Amount" _
                    & " From TEMTTBBALANCE where AccNo='" & ACCNO & "' Group By TransType"
                    
                    Debug.Print mysql
                    
                    Set rsTBBalance = cnn.Execute(mysql)
                    
                    With rsTBBalance
                        If .State = adStateOpen Then
                            While Not .EOF
                                DoEvents
                                transtype = IIf(IsNull(!transtype), "", !transtype)
                                Select Case NormalBal
                                    Case "DR"
                                    Select Case transtype
                                        Case "DR"
                                        balance = balance + IIf(IsNull(!amount), 0, !amount)
                                        Case "CR"
                                        balance = balance - IIf(IsNull(!amount), 0, !amount)
                                    End Select
                                    Case "CR"
                                    Select Case transtype
                                        Case "DR"
                                        balance = balance - IIf(IsNull(!amount), 0, !amount)
                                        Case "CR"
                                        balance = balance + IIf(IsNull(!amount), 0, !amount)
                                    End Select
                                End Select
                                .MoveNext
                            Wend
                        End If
                    End With
                    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX Save To TBBalance XXXXXXXXXXXXXXXXXXXXXXXXXXXX
                    Select Case balance
                        Case Is > 0
                        Case Is < 0
                        balance = balance * (-1)
                        Select Case NormalBal
                            Case "DR"
                            NormalBal = "CR"
                            Case "CR"
                            NormalBal = "DR"
                        End Select
                    End Select
                    If Not Save_TrialBalance(ACCNO, AccName, balance, NormalBal, 0, FinishDate, _
                    User, AccCategory, "", 0, DataSource, errmsg) Then
                        Print_Trial_Balance = False
                        Exit Function
                    End If
                End If
                .MoveNext
            Wend
            '// do the rest
    
       
        End If
        
        
    End With
    
    Else
    '// work on the items here only with o figures
    'MsgBox "here "
    Set rsAccounts = cnn.Execute("set dateformat dmy Select * From GLSETUP  Order By AccNo")
    With rsAccounts
    
        If .State = adStateOpen Then
            While Not .EOF
                DoEvents
                ACCNO = IIf(IsNull(!ACCNO), "", !ACCNO)
                Account = Get_Account_Details(ACCNO, DataSource, errmsg)
                AccCategory = IIf(IsNull(!AccCategory), "", !AccCategory)
                If Account.AccountNo <> "" Then
                    'Balance = Account.OpeningBalance
                    balance = 0
                    NormalBal = Account.NormalBalance
                    AccName = Account.AccountName
                    'XXXXXXXXXXXXXX Get Transactions From Temp TBBalance XXXXXXXXXXXXXXXX
                    mysql = "Select TransType,Sum(Amount) as Amount" _
                    & " From TEMTTBBALANCE where AccNo='" & ACCNO & "' Group By TransType"
                    
                    Debug.Print mysql
                    
                    Set rsTBBalance = cnn.Execute(mysql)
                    
                    With rsTBBalance
                        If .State = adStateOpen Then
                            While Not .EOF
                                DoEvents
                                transtype = IIf(IsNull(!transtype), "", !transtype)
                                Select Case NormalBal
                                    Case "DR"
                                    Select Case transtype
                                        Case "DR"
                                        balance = balance + IIf(IsNull(!amount), 0, !amount)
                                        Case "CR"
                                        balance = balance - IIf(IsNull(!amount), 0, !amount)
                                    End Select
                                    Case "CR"
                                    Select Case transtype
                                        Case "DR"
                                        balance = balance - IIf(IsNull(!amount), 0, !amount)
                                        Case "CR"
                                        balance = balance + IIf(IsNull(!amount), 0, !amount)
                                    End Select
                                End Select
                                .MoveNext
                            Wend
                        End If
                    End With
                    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX Save To TBBalance XXXXXXXXXXXXXXXXXXXXXXXXXXXX
                    Select Case balance
                        Case Is > 0
                        Case Is < 0
                        balance = balance * (-1)
                        Select Case NormalBal
                            Case "DR"
                            NormalBal = "CR"
                            Case "CR"
                            NormalBal = "DR"
                        End Select
                    End Select
                    If Not Save_TrialBalance(ACCNO, AccName, balance, NormalBal, 0, FinishDate, _
                    User, AccCategory, "", 0, DataSource, errmsg) Then
                        Print_Trial_Balance = False
                        Exit Function
                    End If
                End If
                .MoveNext
            Wend
            '// do the rest
    
       
        End If
        
        
    End With

    End If
    Exit Function
SysError:
    Print_Trial_Balance = False
End Function

Public Function generate_trialbalance(Startdate As Date, FinishDate As Date, DataSource As String, _
errmsg As String) As Boolean
    'On Error GoTo SysError
    Open_Database DataSource
    Dim rsAccounts As New Recordset, ACCNO As String, DocumentNo As String, transdate As Date, _
    amount As Double, transtype As String, Account As Account_Details
    Set rsAccounts = New ADODB.Recordset
    
    sql = ("Delete From TEMTTBBALANCE")
    oSaccoMaster.ExecuteThis (sql)
     
    
    'XXXXXXXXXXXXXXXXXXXXXXXX Get_Direct_GL_Transactions XXXXXXXXXXXXXXXXXXXXXXXXX
    Set rsAccounts = Get_Records("Set DateFormat DMY Select DrAccNo,Sum(Amount) As " _
    & "Amount From GLTRANSACTIONS where TransDate >='" & Startdate & "' and dregard=0 and TransDate" _
    & "<='" & FinishDate & "' Group By DrAccNo", DataSource, errmsg)


'
    With rsAccounts
        If .State = adStateOpen Then
            While Not .EOF
                DoEvents
               ' Accno = IIf(IsNull(!Accno), "", !Accno)
               ACCNO = !DRaccno
                'Accno = IIf(IsNull(!DRaccno), "", !Accno)
                'transtype = IIf(IsNull(!transtype), "DR", UCase(!transtype))
                Account = Get_Account_Details(ACCNO, DataSource, errmsg)
                If Account.AccountNo <> "" Then
                    transdate = FinishDate
                    amount = IIf(IsNull(!amount), 0, !amount)
                    If amount > 0 Then
                        If Not Save_TEMTTBBALANCE(ACCNO, transdate, amount, "DR", "", _
                        DataSource, errmsg) Then
                            errmsg = ""
                        End If
                    End If
                End If
                .MoveNext
            Wend
        End If
    End With
    
    Set rsAccounts = Get_Records("Set DateFormat DMY Select CrAccNo,Sum(Amount) As " _
    & "Amount From GLTRANSACTIONS where TransDate>='" & Startdate & "' and  dregard=0 and TransDate" _
    & "<='" & FinishDate & "' Group By CrAccNo", DataSource, errmsg)

'    Set rsAccounts = Get_Records("Set DateFormat DMY Select AccNo,Sum(Amount) As " _
'    & "Amount From CUSTOMERBALANCE where TransDate>='" & StartDate & "' and TransDate" _
'    & "<='" & FinishDate & "' and Transtype='CR' Group By AccNo", DataSource, ErrMsg)
    
    With rsAccounts
        If .State = adStateOpen Then
            While Not .EOF
                DoEvents
                ACCNO = !Craccno
                'Accno = IIf(IsNull(!Accno), "", !Accno)
                Account = Get_Account_Details(ACCNO, DataSource, errmsg)
                If Account.AccountNo <> "" Then
                    transdate = FinishDate
                    amount = IIf(IsNull(!amount), 0, !amount)
                    If amount > 0 Then
                        If Not Save_TEMTTBBALANCE(ACCNO, transdate, amount, "CR", "", _
                        DataSource, errmsg) Then
                            errmsg = ""
                        End If
                    End If
                End If
                .MoveNext
            Wend
        End If
    End With
    
    '//amount from the payroll only
    sql = ""
    sql = "d_gl_payroll '" & Startdate & "', '" & FinishDate & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    While Not rs.EOF
    
                sql = "SET              dateformat dmy"
                sql = sql & "                  SELECT"
                sql = sql & "  transportaccno , agrovetaccno, AIaccno, tmsharesaccno, fsaaccno, hsharesaccno, advanceaccno, otheraccno, MainAccNo, netaccno"
                sql = sql & "      From d_Payroll"
                sql = sql & "    WHERE     endofperiod >= '" & Startdate & "' AND endofperiod <= '" & FinishDate & "'"
                Set rst = oSaccoMaster.GetRecordset(sql)
                If rs.Fields(0) > 0 Then
                If Not Save_TEMTTBBALANCE(IIf(IsNull(rst.Fields(0)), 0, rst.Fields(0)), FinishDate, rs.Fields(0), "CR", "", DataSource, errmsg) Then
                            errmsg = ""
                        End If
                End If
                If rs.Fields(1) > 0 Then
                If Not Save_TEMTTBBALANCE(IIf(IsNull(rst.Fields(1)), 0, rst.Fields(1)), FinishDate, rs.Fields(1), "CR", "", _
                        DataSource, errmsg) Then
                            errmsg = ""
                        End If
                End If
                If rs.Fields(2) > 0 Then
                If Not Save_TEMTTBBALANCE(IIf(IsNull(rst.Fields(2)), 0, rst.Fields(2)), FinishDate, rs.Fields(2), "CR", "", _
                        DataSource, errmsg) Then
                            errmsg = ""
                        End If
                End If
                If rs.Fields(3) > 0 Then
                If Not Save_TEMTTBBALANCE(IIf(IsNull(rst.Fields(3)), 0, rst.Fields(3)), FinishDate, rs.Fields(3), "CR", "", _
                        DataSource, errmsg) Then
                            errmsg = ""
                        End If
                End If
                If rs.Fields(4) > 0 Then
                If Not Save_TEMTTBBALANCE(IIf(IsNull(rst.Fields(4)), 0, rst.Fields(4)), FinishDate, rs.Fields(4), "CR", "", _
                        DataSource, errmsg) Then
                            errmsg = ""
                        End If
                End If
                If rs.Fields(5) > 0 Then
                If Not Save_TEMTTBBALANCE(IIf(IsNull(rst.Fields(5)), 0, rst.Fields(5)), FinishDate, rs.Fields(5), "CR", "", _
                        DataSource, errmsg) Then
                            errmsg = ""
                        End If
                End If
                If rs.Fields(6) > 0 Then
                If Not Save_TEMTTBBALANCE(IIf(IsNull(rst.Fields(6)), 0, rst.Fields(6)), FinishDate, rs.Fields(6), "CR", "", _
                        DataSource, errmsg) Then
                            errmsg = ""
                        End If
                End If
                If rs.Fields(7) > 0 Then
                If Not Save_TEMTTBBALANCE(IIf(IsNull(rst.Fields(7)), 0, rst.Fields(7)), FinishDate, rs.Fields(7), "CR", "", _
                        DataSource, errmsg) Then
                            errmsg = ""
                        End If
                End If
                If rs.Fields(8) > 0 Then
                If Not Save_TEMTTBBALANCE(IIf(IsNull(rst.Fields(8)), 0, rst.Fields(8)), FinishDate, rs.Fields(8), "DR", "", _
                        DataSource, errmsg) Then
                            errmsg = ""
                        End If
                End If
                If rs.Fields(9) > 0 Then
                If Not Save_TEMTTBBALANCE((IIf(IsNull(rst.Fields(9)), 0, rst.Fields(9))), FinishDate, rs.Fields(9), "CR", "", DataSource, errmsg) Then
                            errmsg = ""
                        End If
                End If
                        
    rs.MoveNext
    Wend
    
    If Not Print_Trial_Balance(Startdate, FinishDate, DataSource, errmsg) Then
        If errmsg <> "" Then
        End If
    End If
    Exit Function
SysError:
    errmsg = err.description
    generate_trialbalance = False
End Function

Private Function Save_TEMTTBBALANCE(ACCNO As String, transdate As Date, amount As Double, _
transtype As String, DocumentNo As String, DataSource As String, errmsg As String) As Boolean
    'On Error GoTo SysError
    Open_Database DataSource
    With cn
        If .State = adStateOpen Then
            .Execute ("Set DateFormat DMY Exec Save_TEMTTBBALANCE '" & ACCNO & "','" & _
            transdate & "'," & amount & ",'" & transtype & "','" & DocumentNo & "'")
        End If
    End With
    Save_TEMTTBBALANCE = True
    Exit Function
SysError:
    Save_TEMTTBBALANCE = False
End Function
Private Function Save_TBBalance() As Boolean

End Function

Private Sub Open_Database(DataSource As String)
    On Error GoTo SysError
    Set cn = New Connection
    DataSource = "MAZIWA"
    With cn
        If .State = adStateClosed Then
            .Open DataSource, "atm"
        End If
    End With
    Exit Sub
SysError:
    
End Sub

Private Sub Class_Terminate()
'    With cn
'        If .State = adStateOpen Then
'            .Close
'        End If
'        Set cn = Nothing
'    End With
End Sub
