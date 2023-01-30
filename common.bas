Attribute VB_Name = "common"

Option Explicit
Public DocumentNo As String
Public DelimiterConstant As Integer
Public Deductionfield As String
Public Startdate As Date
Public cnnPayroll As Connection
Public FinishDate As Date
Public TSource As String
Public Mach  As String
Dim sessionId As String
Public tdate As Date
Public amt As Double
'Public success As Boolean
Public Provider As String

Public nd As Node
Public Enddate As Date
Public rsActiveUser As New Recordset
Public currentUser As Current_User
Public SelectedDBMS As String
Public BeginDate As Date
Public RsLoans As New ADODB.Recordset
Public RCanc As Boolean
Public serverDate  As Date
Public OldTransDate As Date
Public mTransDate As Date
Public newacc As Boolean
Public ReportDate As Date
Public IncludeClearedLoans As Boolean
Public strName As String
Public PPP As String
Public mDocNo As String
Public mChequeNo As String
Public bcode As String
Dim myclass As Object
Dim temprs As Object
Public res As String
Public clsClass As Object
Public mvarConnection As String
Public ServerName As String
Public DataBaseName As String
Public lvwAll As ListView
Public SelectedItem As String
Public teagrowerno As String
'Public Provider As String
Public alreadyDone As Boolean
Public maxRec As Long
Public myLevel As Long
'Public report_Path As String
Public TotalTime As Currency
Public NormalRate As Currency
Public MinimumRate As Currency
Public MinimumTime As Currency
Public MyHr As Currency
Public memberno As String
Public MyMin As Integer
Public MySec As Currency
Public isMember As Boolean
'Public CompanyName As String
Public CompanyName As String, CompanyPhone As String, CompanyTown As String, CompanyTagLine As String
Public Credit As Currency
Public Id As Single
Public StartingTime As String
Public IsExit As Boolean
Public Logoff_Browsing_Time As Currency
Public MemberName As String
Public Am_Changing_Company_Name As Boolean
Public Type GL_MainAcc
    AccCode As String
    AccName As String
    AccountType As String
End Type
Private Type accInfo
    ACCNO As String
    custName As String
    custBal As Currency
    AccName As String
    custno As String
    pic As String
    sign As String
End Type
Public Type Current_User
    username As String
    isTeller As Boolean
    tellerGlAcc As String
    idno As String
End Type
Public Type Cub_Acc_Details
    idno As String
    CustomerNo As String
    payrollno As String
    ACCNO As String
    AccName As String
    availablebalance As Double
End Type

Public Type Acc_Details
    ACCNO As String
    AccName As String
    NormalBal As String
    OpeningBal As Double
    CurrentBal As Double
End Type

Public accData() As accInfo

Public Type LoanGL_Accounts
    LoanAcc As String
    ContraAcc As String
    interestAcc As String
    SharesAcc As String
End Type

Public Type Society_Parameters
    SocietyName As String
    sharecap As Double
    LoanRecoveryMethod As Integer
    MinimumShares As Double
End Type

Public Type Scheme_Details
    memberno As String
    totalshares As Double
End Type
Public Sub process_payroll(SNO As Integer)
'dim StartDate s,@EndPeriod varchar(10) , @Year bigint, @User varchar(35) AS
Dim Yr As Integer
Startdate = DateSerial(Year(Format(Get_Server_Date, "dd/mm/yyyy")), Month(Format(Get_Server_Date, "dd/mm/yyyy")), 1)
Enddate = DateSerial(Year(Format(Get_Server_Date, "dd/mm/yyyy")), Month(Format(Get_Server_Date, "dd/mm/yyyy")) + 1, 1 - 1)

Yr = Year(Format(Get_Server_Date, "dd/mm/yyyy"))

oSaccoMaster.ExecuteThis ("d_sp_PresetDeductAssign_99 '" & Startdate & "','" & Enddate & "'," & Yr & ",'" & User & "'," & SNO & "")

oSaccoMaster.ExecuteThis ("d_sp_GDedNet_99 '" & Startdate & "','" & Enddate & "'," & SNO & "")

oSaccoMaster.ExecuteThis ("d_sp_TransUpdate '" & Startdate & "','" & Enddate & "','" & User & "'")


oSaccoMaster.ExecuteThis ("d_sp_TransPRoll '" & Startdate & "','" & Enddate & "','" & User & "'")




End Sub

Public Function To_Upper_Case(Character As Integer) As Integer
    On Error GoTo errFix
    If Character <> 13 Then 'Catch the Enter key
        To_Upper_Case = Asc(UCase(Chr(Character)))
    End If
    Exit Function
errFix:
    MsgBox err.description, vbOKOnly, "Company Setup"
End Function

Public Function Get_Society_Details() As Society_Parameters
On Error GoTo Syserr
Dim rsSociety As New Recordset
Set rsSociety = Get_Records("Select * from SYSPARAM", ErrorMessage)
 With rsSociety
    If Not .EOF Then
        Get_Society_Details.LoanRecoveryMethod = IIf(IsNull(!LoanRecoveryOption), 0, !LoanRecoveryOption)
        Get_Society_Details.MinimumShares = IIf(IsNull(!mintotshares), 0, !mintotshares)
        Get_Society_Details.SocietyName = IIf(IsNull(!CompanyName), "", !CompanyName)
        Get_Society_Details.sharecap = IIf(IsNull(!ShareCapital), 0, !ShareCapital)
    End If
 End With
Exit Function
Syserr:
    MsgBox err.description
End Function

Public Function Get_Acc_Details(ACCNO As String, errmsg As String) As Acc_Details
    On Error GoTo SysError
    Dim rsAcc As New Recordset
    Set cn = New Connection
    With cn
        If .State = adStateClosed Then
            .Open SelectedDsn, "atm"
        End If
        Set rsAcc = .Execute("Select * From GLSETUP where AccNo='" & ACCNO & "'")
    End With
    With rsAcc
        If .State = adStateOpen Then
            If Not .EOF Then
                Get_Acc_Details.AccName = IIf(IsNull(!GlAccName), "", !GlAccName)
                Get_Acc_Details.ACCNO = ACCNO
                Get_Acc_Details.CurrentBal = IIf(IsNull(!CurrentBal), 0, !CurrentBal)
                Get_Acc_Details.NormalBal = IIf(IsNull(!NormalBal), "DR", IIf(!NormalBal <> "Debit", "CR", "DR"))
                Get_Acc_Details.OpeningBal = IIf(IsNull(!OpeningBal), 0, !OpeningBal)
            End If
        End If
    End With
    Exit Function
SysError:
    ACCNO = ""
    errmsg = err.description
End Function

Public Function Get_Records(ssql As String, errmsg As String) As Recordset
    On Error GoTo SysError
    Dim CnMAZIWA As New Connection
    With CnMAZIWA
        If .State = adStateClosed Then
            .Open SelectedDsn, "atm"
        End If
        Set Get_Records = .Execute(ssql)
    End With
    Exit Function
SysError:
    errmsg = err.description & vbCrLf & err.number & vbCrLf & err.Source
End Function
Public Function Generate_ReceiptNo(Choice As String) As String
    On Error GoTo SysError
    Dim rsNos As New Recordset, strmemno As String, lngMemNo As Long
'    Select Case Choice
'        Case "Voucher"
'            Set rsNos = oSaccoMaster.GetRecordset("exec getNextNumber 'Payments','Voucherno'")
'        Case "Receipt"
'            Set rsNos = oSaccoMaster.GetRecordset("exec getNextNumber 'Receipts','ReceiptNo'")
'    End Select
'    With rsNos
'            Generate_ReceiptNo = IIf(Choice = "Voucher", "V/NO-", "RCP-") & str(IIf(IsNull(.Fields(0)), 0, .Fields(0)) + 1)
'    End With
    Exit Function
SysError:
    ErrorMessage = err.description
    Generate_ReceiptNo = ""
End Function
Public Function Get_Cub_Acc_Details(ACCNO As String, errmsg As String) As Cub_Acc_Details

    On Error GoTo SysError
    Dim rsAcc As New Recordset
    Set cn = New Connection
    With cn
        If .State = adStateClosed Then
            .Open SelectedDsn, "atm"
        End If
        Set rsAcc = .Execute("Select * From cub where AccNo='" & ACCNO & "'")
    End With
    With rsAcc
        If .State = adStateOpen Then
            If Not .EOF Then
                
                Get_Cub_Acc_Details.idno = IIf(IsNull(!idno), "", !idno)
                Get_Cub_Acc_Details.AccName = IIf(IsNull(!Name), "", !Name)
                Get_Cub_Acc_Details.payrollno = IIf(IsNull(!ACCNO), "", !ACCNO)
                Get_Cub_Acc_Details.ACCNO = IIf(IsNull(!ACCNO), "", !ACCNO)
                Get_Cub_Acc_Details.availablebalance = IIf(IsNull(!availablebalance), 0, !availablebalance)
            End If
        End If
    End With
    Exit Function
SysError:
    ACCNO = ""
    
    errmsg = err.description
End Function

Public Function ValidChar(ByVal CharacterValue As Long) As Boolean
    If CharacterValue = 8 Then ValidChar = True: Exit Function
    If CharacterValue = 38 Then ValidChar = True: Exit Function
    If CharacterValue = 45 Then ValidChar = True: Exit Function
    If CharacterValue = 46 Then ValidChar = True: Exit Function
    If CharacterValue = 95 Then ValidChar = True: Exit Function
    If CharacterValue = 32 Then ValidChar = True: Exit Function
    If CharacterValue = 9 Or CharacterValue = 13 Then ValidChar = True: Exit Function
    If CharacterValue = 8 Then ValidChar = False: Exit Function
    If CharacterValue = 39 Or CharacterValue = 124 Or CharacterValue = 34 Or CharacterValue = 35 Or CharacterValue = 125 Or CharacterValue = 123 Or CharacterValue = 126 Or CharacterValue = 96 Or CharacterValue = 33 Or CharacterValue = 64 Or CharacterValue = 36 Or CharacterValue = 37 Or CharacterValue = 94 Or CharacterValue = 38 Or CharacterValue = 42 Or CharacterValue = 40 Or CharacterValue = 41 Or CharacterValue = 60 Or CharacterValue = 62 Then
    ValidChar = False
    Else
        If (CharacterValue >= 48 And CharacterValue <= 57) Or (CharacterValue > 64 And CharacterValue < 91) Or (CharacterValue > 96) And CharacterValue < 123 Then
            ValidChar = True
        Else
            ValidChar = False
        End If
    End If
End Function
Public Function Save_GLTRANSACTION33(transdate As Date, amount As Double, DRaccno As String, _
Craccno As String, DocumentNo As String, Source As String, auditid As String, errmsg As _
String, transDescription As String, CashBook As Long, doc_posted As Integer, chequeno As String) As Boolean
    On Error GoTo SysError
Dim cn As New Connection
    With cn
        If .State = adStateClosed Then
            .Open SelectedDsn, "atm"
        End If
        sql = "Set DateFormat DMY Exec Save_GLTRANSACTION33 '" & transdate & "'," _
        & amount & ",'" & DRaccno & "','" & Craccno & "','" & DocumentNo & "','" & _
        Source & "','" & auditid & "','" & transDescription & "'," & CashBook & "," & doc_posted & ",'" & chequeno & "'"
        .Execute (sql)
    End With
    Save_GLTRANSACTION33 = True
    Exit Function
SysError:
    errmsg = err.description
    Save_GLTRANSACTION33 = False
End Function
Public Function Save_GLTRANSACTION(transdate As Date, amount As Double, DRaccno As String, _
Craccno As String, DocumentNo As String, Source As String, auditid As String, errmsg As _
String, transDescription As String, CashBook As Long, doc_posted As Integer, chequeno As String _
, transactionNo As String, Optional VoucherNo As String, Optional LOP As String, Optional Module As String, Optional pmode As String, Optional RefId As String) As Boolean
    On Error GoTo SysError
Dim cnn As New Connection
    ErrorMessage = ""
    With cnn
        If .State = adStateClosed Then
            .Open SelectedDsn, "atm", "atm"
        End If
        sql = ("Set DateFormat DMY Exec Save_GLTRANSACTION45 '" & Format(transdate, "DD/MM/YYYY") & "'," _
        & amount & ",'" & DRaccno & "','" & Craccno & "','" & DocumentNo & "','" & _
        Source & "','" & auditid & "','" & transDescription & "','" & chequeno & "','" & transactionNo & "'")
        oSaccoMaster.ExecuteThis (sql)
'        If success = False Then
'            GoTo SysError
'        End If
'    End With
'    Save_GLTRANSACTION = True
'    Exit Function
'SysError:
'    Save_GLTRANSACTION = False
End With
    Save_GLTRANSACTION = True
    Exit Function
SysError:
    errmsg = err.description
    Save_GLTRANSACTION = False
End Function

Public Function Save_Member_Statement(memberno As String, Period As Date, OpeningBalance _
As Double, NewLoan As Double, InterestPaid As Double, InterestCharged As Double, _
InterestOwing As Double, LoanRepayment As Double, OutstandingLoanBalance As Double, _
OpeningShares As Double, SharesContributed As Double, ClosingShares As Double, _
TotalMonthlyContribution As Double, errmsg As String) As Boolean
    On Error GoTo SysError
    Dim cn As New Connection
    With cn
        If .State = adStateClosed Then
            .Open SelectedDsn, "atm"
        End If
        .Execute ("Set DateFormat DMY Exec SAVE_STATEMENT '" & memberno & "','" & Period _
        & "'," & OpeningBalance & "," & NewLoan & "," & InterestPaid & "," & InterestCharged _
        & "," & InterestOwing & "," & LoanRepayment & "," & OutstandingLoanBalance & "," & _
        OpeningShares & "," & SharesContributed & "," & ClosingShares & "," & _
        TotalMonthlyContribution)
    End With
    Save_Member_Statement = True
    Exit Function
SysError:
    errmsg = err.description
    Save_Member_Statement = False
End Function
Public Function Get_Path(errormsg As String) As String
    On Error GoTo SysError
    Dim rsPath As New Recordset
    Set rsPath = oSaccoMaster.GetRecordset("Select * from reportpath")
    With rsPath
        If Not .EOF Then
            Get_Path = .Fields("ReportPath")
        Else
            Get_Path = ""
        End If
    End With
    Exit Function
SysError:
    Get_Path = ""
    errormsg = err.description
End Function

Public Function Save_Guarantor(memberno As String, Loanno As String, _
amount As Double, balance As Double, auditid As String, GDATE As Date, _
errormsg As String) As Boolean
    On Error GoTo SysError
    Dim cn As New Connection
    With cn
        If .State = adStateClosed Then
            .Open SelectedDsn, "atm"
        End If
        .Execute ("Set DateFormat DMY Exec Save_LoanGuar '" & memberno & "','" _
        & Loanno & "'," & amount & "," & balance & ",'" & auditid & "','" & _
        GDATE & "'")
    End With
    If Not Save_Audit("LOANGUAR", "Updating G. MemberNo " & memberno, _
    GDATE, amount, auditid, errormsg) Then
        If errormsg <> "" Then
            Save_Guarantor = False
            Exit Function
        End If
    End If
    Save_Guarantor = True
    Exit Function
SysError:
    errormsg = err.description
    Save_Guarantor = False
End Function

Public Function Member_Withdrawn(memberno As String) As Boolean
    On Error GoTo SysError
    Dim rsWithdrawn As New Recordset
    Set rsWithdrawn = oSaccoMaster.GetRecordset("Select * From WITHDRWN where " _
    & "MemberNo='" & memberno & "'")
    With rsWithdrawn
        If Not .EOF Then
            Member_Withdrawn = True
        Else
            Member_Withdrawn = False
        End If
    End With
    Exit Function
SysError:
    MsgBox err.description, vbInformation, "Withdrawn Members"
End Function

Public Function Update_MyCustBalance(AccName As String, memberno As String, OldAmount As Double, ACCNO _
As String, OldTransDate As Date, OldChequeNo As String, OldVno As String, NewAmount _
As Double, NewTransDate As Date, NewChequeNo As String, transtype As String, NewVno _
As String, errormsg As String, customerbalanceid As Double, Optional cn As Connection) As Boolean
    On Error GoTo SysError
    If NewAmount < 0 Then
        NewAmount = NewAmount * (-1)
        Select Case transtype
            Case "DR"
            transtype = "CR"
            Case "CR"
            transtype = "DR"
        End Select
    End If
    If Not cn Is Nothing Then
        With cn
            If .State = adStateClosed Then
                .Open SelectedDsn, "atm"
            End If
            '///THIS IS WHERE THE PROBLEM IS ??????????????
            
            .Execute ("Set DateFormat DMY Exec Update_MyCustomerBalance '" & memberno _
            & "'," & OldAmount & ",'" & ACCNO & "','" & OldTransDate & "','" & _
            OldChequeNo & "','" & OldVno & "'," & NewAmount & ",'" & NewTransDate & _
            "','" & NewChequeNo & "','" & transtype & "','" & NewVno & "','" & AccName & "'," & customerbalanceid & "")
        End With
    Else
        Dim Cnn1 As New Connection
        With Cnn1
            If .State = adStateClosed Then
                .Open SelectedDsn, "atm"
            End If
            
            .Execute ("Set DateFormat DMY Exec Update_MyCustomerBalance '" & memberno _
            & "'," & OldAmount & ",'" & ACCNO & "','" & OldTransDate & "','" & _
            OldChequeNo & "','" & OldVno & "'," & NewAmount & ",'" & NewTransDate & _
            "','" & NewChequeNo & "','" & transtype & "','" & NewVno & "','" & AccName & "'," & customerbalanceid & "")
        End With
    End If
    Update_MyCustBalance = True
    Exit Function
SysError:
    errormsg = err.description
    Update_MyCustBalance = False
End Function

Public Function Save_Def_Loan(memberno As String, Loanno As String, LastTransDate As Date, auditid As String, errormsg As String, PrincipleAmount As Currency, org As String, intiald As Currency, datefrom As Date, dateto As Date, offsetfromshares As Currency, otherrecoveries As Currency, outstandingbal As Currency, Remarks As String, Class As Integer, days As Integer, g As Integer) As Boolean
    On Error GoTo SysError
    Dim CnnDef As New Connection
    With CnnDef
        If .State = adStateClosed Then
            .Open SelectedDsn, "atm"
        End If
        .Execute ("Set DateFormat DMY Exec Save_DefaultedLoan '" & memberno & "','" & Loanno & "','" & _
        LastTransDate & "','" & auditid & "'," & PrincipleAmount & ",'" & org & "'," & intiald & ",'" & datefrom & "','" & dateto & "'," & offsetfromshares & "," & otherrecoveries & "," & outstandingbal & ",'" & Remarks & "'," & Class & "," & days & "," & g & "")
        
        '//CHECK IF THE ITEM IS IN THE LIST
        Set rs = oSaccoMaster.GetRecordset("SELECT LOANNO FROM DEFAULTED WHERE LOANNO='" & Loanno & "'")
        If rs.EOF Then
        
        .Execute ("Set DateFormat DMY Exec Save_Defaulted '" & memberno & "','" & Loanno & "','" & _
        LastTransDate & "','" & auditid & "'," & PrincipleAmount & ",'" & org & "'," & intiald & ",'" & datefrom & "','" & dateto & "'," & offsetfromshares & "," & otherrecoveries & "," & outstandingbal & ",'" & Remarks & "'," & Class & "," & days & "," & g & "")
        Else
        .Execute ("Set DateFormat DMY Exec UPDATE_DefaultedLOAN '" & Loanno & "','" & _
        LastTransDate & "','" & datefrom & "','" & dateto & "'," & offsetfromshares & "," & otherrecoveries & "," & outstandingbal & "," & Class & "")
        End If
    End With
    Save_Def_Loan = True
    Exit Function
SysError:
    errormsg = err.description
    Save_Def_Loan = False
End Function

Public Function Save_Audit(TransTable As String, transDescription As String, transdate As Date, _
amount As Double, auditid As String, errormsg As String) As Boolean
    On Error GoTo SysError
    Dim CnAudit As New Connection
    With CnAudit
        If .State = adStateClosed Then
            .Open SelectedDsn, "atm"
        End If
        .Execute ("Set DateFormat DMY Exec Save_AuditTrans '" & TransTable & "','" & _
        transDescription & "','" & transdate & "'," & amount & ",'" & auditid & "'")
    End With
    Save_Audit = True
    Exit Function
SysError:
    errormsg = err.description
    Save_Audit = False
End Function

Public Function Save_TB(ACCNO As String, AccName As String, amount As Double, _
transtype As String, Closed As Integer, transdate As Date, auditid As String, _
errormsg As String, Optional AccountType As String, Optional GLAccGroup As _
String, Optional BudgetAmount As Double) As Boolean
    On Error GoTo SysError
    Dim CnTB As New Connection
    With CnTB
        If .State = adStateClosed Then
            .Open SelectedDsn, "atm"
        End If
        .Execute ("Set DateFormat DMY Exec Save_TrialBalance '" & ACCNO & "','" _
        & AccName & "'," & amount & ",'" & transtype & "',0,'" & transdate & _
        "','" & auditid & "','" & AccountType & "','" & GLAccGroup & "'," & BudgetAmount)
    End With
    Save_TB = True
    Exit Function
SysError:
    errormsg = err.description
    Save_TB = False
End Function

Public Function Save_Monthly_Deduction(memberno As String, companycode As String, _
transdate As Date, Shares As Double, Loanno As String, Loans As Double, interest As _
Double, IAR As Double, RegFees As Double, ByLaws As Double, DocumentNo As String, _
auditid As String, errormsg As String) As Boolean
    On Error GoTo SysError
    Dim cn As New Connection
    With cn
        If .State = adStateClosed Then
            .Open SelectedDsn, "atm"
        End If
        .Execute ("Set DateFormat DMY Exec Save_MonthlyDeduction '" & memberno & "','" _
        & companycode & "','" & transdate & "'," & Shares & ",'" & Loanno & "'," & _
        Loans & "," & interest & "," & IAR & "," & RegFees & "," & ByLaws & ",'" & _
        DocumentNo & "','" & auditid & "'")
    End With
    Save_Monthly_Deduction = True
    Exit Function
SysError:
    errormsg = err.description
    Save_Monthly_Deduction = False
End Function

Public Function SAVE_TMPCABOOK(MCODE As String, mName As String, amount As Double, _
TransMonth As Long, TransYear As Long, mTransType As String, transdate As Date, _
errormsg As String) As Boolean
    On Error GoTo SysError
    Dim cn As New Connection
    With cn
        If .State = adStateClosed Then
            .Open SelectedDsn, "atm"
            .Execute ("Set DateFormat DMY Exec Save_TEMPCASHBOOK '" & MCODE & "','" _
            & mName & "'," & amount & "," & TransMonth & "," & TransYear & ",'" & _
            mTransType & "','" & transdate & "'")
        End If
    End With
    SAVE_TMPCABOOK = True
    Exit Function
SysError:
    errormsg = err.description
    SAVE_TMPCABOOK = False
End Function

Public Function SAVE_UNDERPAID(Loanno As String, AmountPaid As Double, ExpectedAmount As Double, _
loanbalance As Double, Period As Date, auditid As String, errormsg As String) As Boolean
    On Error GoTo SysError
    Dim cn As New Connection
    With cn
        If .State = adStateClosed Then
            .Open SelectedDsn, "atm"
        End If
        .Execute ("Set DateFormat DMY Exec SAVE_UNDERPAIDLOANS '" & Loanno & "'," _
        & AmountPaid & "," & ExpectedAmount & "," & loanbalance & ",'" & Period & _
        "','" & auditid & "'")
    End With
    SAVE_UNDERPAID = True
    Exit Function
SysError:
    SAVE_UNDERPAID = False
    ErrorMessage = err.description
End Function

Public Function Save_Withdrawn(memberno As String, NoticeDate As Date, DateWithdrawn _
As Date, totalshares As Double, LoanBalances As Double, auditid As String, errormsg As _
String) As Boolean
    On Error GoTo SysError
    Dim cn As New Connection
    With cn
        If .State = adStateClosed Then
            .Open SelectedDsn, "atm"
        End If
        .Execute ("Set DateFormat DMY Exec Save_Withdrawn '" & memberno & "','" & _
        NoticeDate & "','" & DateWithdrawn & "'," & totalshares & "," & LoanBalances _
        & ",'" & auditid & "'")
    End With
    Save_Withdrawn = True
    Exit Function
SysError:
    Save_Withdrawn = False
    errormsg = err.description
End Function

Public Function Register_Scheme(memberno As String, transdate As Date, _
amount As Double, SchemeCode As String, shareBal As Double, transby As String, _
auditid As String, errormsg As String, SharesAcc As String, ContraAcc As String) As Boolean
Dim RsShares As New Recordset
Dim rsShareVar As New Recordset
On Error GoTo SysError
'Insert into SHARES
Set RsShares = oSaccoMaster.GetRecordset("Select * from SHARES where" _
& " MemberNo='" & memberno & "' and SharesCode='" & SchemeCode & "'")
        With RsShares
            .AddNew
            !memberno = memberno
            !initshares = 0
            !totalshares = 0
            !transdate = transdate
            !LastDivDate = transdate
            !sharesCode = SchemeCode
            !audittime = Get_Server_Date
            !auditid = User
            .Update
        End With
Register_Scheme = True

'Insert Into SHRVAR
Set rsShareVar = oSaccoMaster.GetRecordset("select * from SHRVAR where" _
& " MemberNo='" & memberno & "' and SharesCode='" & SchemeCode & "'")
        With rsShareVar
            .AddNew
            !memberno = memberno
            !oldcontr = amount
            !NewContr = amount
            !VarDate = transdate
            !sharesCode = SchemeCode
            !audittime = Get_Server_Date
            !auditid = User
            .Update
        End With
        Register_Scheme = True
  
    
'Insert Into CONTRIB
Set RsShares = oSaccoMaster.GetRecordset("Set DateFormat DMY Exec Save_Contrib1 '" & memberno & "','" _
        & transdate & "',1000," & amount & "," & amount & ",'" & _
        transby & "','" & transby & "','" & transby & "','0','0','NA','" & User & "','1000',0,'" _
        & SchemeCode & "','" & SharesAcc & "','" & ContraAcc & "','" & transdate & "'")
   Exit Function
SysError:
    Register_Scheme = False
    errormsg = ""
    errormsg = err.description
End Function
Public Function Save_Contrib(mMemberNo As String, ContrDate As Date, RefNo _
As Long, amount As Double, shareBal As Double, transby As String, chequeno As _
String, ReceiptNo As String, Locked As String, Posted As String, Remarks As _
String, auditid As String, TransNo As String, PeriodDate As Date, errormsg As _
String, Optional Offset As Long, Optional SchemeCode As String, Optional SharesAcc As String, _
Optional ContraAcc As String, Optional cashbookdate As Date) As Boolean
    On Error GoTo SysError
    Dim cn As New Connection
    Dim rscheck As New Recordset, rsCheckScheme As New Recordset
    
    'Check if the Scheme exists
    Set rsCheckScheme = Nothing
    If SchemeCode = "" Then SchemeCode = "L009"
    If SchemeCode <> "" Then
        Set rsCheckScheme = oSaccoMaster.GetRecordset("select * from shareType where Sharescode='" & SchemeCode & "'")
         If Not rsCheckScheme.EOF Then 'Scheme Exists
         'Continue
         Else
            MsgBox "The Scheme " & SchemeCode & " does NOT exist. Ensure the right scheme is entered.", vbInformation, "Post Shares"
            Exit Function
         End If
    Else
     MsgBox "The Scheme " & SchemeCode & " does NOT exist. Ensure the right scheme is entered.", vbInformation, "Post Shares"
            ErrorMessage = "Scheme does not Exist"
            Exit Function
     End If
     
     'Check if Member is in this scheme, Create the Account if it does not exist
    Set rscheck = oSaccoMaster.GetRecordset("Select * from Shares where MemberNo='" & Trim(mMemberNo) & "' and Sharescode='" & SchemeCode & "'")
    If Not rscheck.EOF Then
       'Continue Posting
    Else
        If MsgBox("Member " & mMemberNo & " is NOT Registered to this Shares Scheme." & vbCrLf & " Do you want to Create this Account?", vbQuestion + vbYesNo, "Scheme Registration") = vbNo Then
             Exit Function
        Else 'create account
            If Not Register_Scheme(mMemberNo, ContrDate, amount, SchemeCode, shareBal, transby, auditid, errormsg, SharesAcc, ContraAcc) Then
                If errormsg <> "" Then
                    Save_Contrib = False
                    Exit Function
                End If
            Else
                Save_Contrib = False
                Exit Function
            End If
        End If
    End If
        
    With cn
        If .State = adStateClosed Then
            .Open SelectedDsn, "atm"
        End If
        .Execute ("Set DateFormat DMY Exec Save_Contrib1 '" & mMemberNo & "','" _
        & ContrDate & "'," & RefNo & "," & amount & "," & shareBal & ",'" & _
        transby & "','" & chequeno & "','" & ReceiptNo & "','" & Locked & "','" _
        & Posted & "','" & Remarks & "','" & auditid & "','" & TransNo & "'," _
        & Offset & ",'" & SchemeCode & "','" & SharesAcc & "','" & ContraAcc & "','" & Format(cashbookdate, "dd/mm/yyyy") & "'")
    End With
    If Not Save_Audit("Contrib", "Shares Contribution. MemberNo " & mMemberNo, _
    ContrDate, amount, auditid, errormsg) Then
        If errormsg <> "" Then
            Save_Contrib = False
            Exit Function
        End If
    End If
    Save_Contrib = True
    If Not Refresh_Shares(mMemberNo, errormsg, SchemeCode) Then
        GoTo SysError
    End If
    Exit Function
SysError:
    errormsg = err.description
    Save_Contrib = False
End Function

Public Function Update_Repay(StatementDate As Date, Loanno As String, RepayID As Long, memberno As String, _
datereceived As Date, paymentno As Long, amount As Double, principal As Double, _
interest As Double, IntrCharged As Double, IntrOwed As Double, loanbalance As Double, _
ReceiptNo As String, Locked As Long, Posted As Long, Accrued As Long, Remarks As String, _
auditid As String, transby As String, intbalance As Double, NextDueDate As Date, Ch As _
String, TransNo As String, errmsg As String, Optional DocumentNo As String) As Boolean
    On Error GoTo SysError
    Dim CnRepay As New Connection
    With CnRepay
        If .State = adStateClosed Then
            .CursorLocation = adUseServer
            .Open SelectedDsn, "atm"
        End If
        .Execute ("Set DateFormat DMY Exec Update_Repay '" & Loanno & "','" & StatementDate & "'," & RepayID & ",'" & _
        memberno & "','" & NextDueDate & "'," & paymentno & "," & amount & "," & principal _
        & "," & interest & "," & IntrCharged & "," & IntrOwed & "," & loanbalance & ",'" _
        & ReceiptNo & "'," & Locked & "," & Posted & "," & Accrued & ",'" & Remarks & "','" _
        & auditid & "','" & transby & "'," & intbalance & ",'" & NextDueDate & "','" & Ch _
        & "','" & TransNo & "','" & DocumentNo & "'")
    End With
    Update_Repay = True
    Exit Function
SysError:
    Update_Repay = False
    errmsg = err.description
End Function

Public Function GLAccount_Exists(ACCNO As String, errormsg As String) As Boolean
    On Error GoTo SysError
    Dim rsgl As New Recordset
    Set rsgl = oSaccoMaster.GetRecordset("Select AccNo from GLSETUP where AccNo='" _
    & ACCNO & "'")
    With rsgl
        If .State = adStateOpen Then
            If Not .EOF Then
                GLAccount_Exists = True
            Else
                GLAccount_Exists = False
                errormsg = "Account No " & ACCNO & " not found"
            End If
        End If
    End With
    Exit Function
SysError:
    errormsg = err.description
    GLAccount_Exists = False
End Function

Public Function Update_CustomerBalance(customerbalanceid As Long, CustomerNo As String, _
idno As String, payrollno As String, AccName As String, amount As Double, availablebalance _
As Double, ACCNO As String, transDescription As String, transdate As Date, Commission As _
Double, chequeno As String, Period As String, Posted As Long, Locked As Long, transtype As _
String, status As String, vno As String, auditid As String, AuditDate As Date, moduleid As _
String, accd As String, valuedate As Date, actualbalance As Double, cash As Long, bcode As _
String, Rebuild As Long, errormsg As String) As Boolean
    On Error GoTo SysError
    Dim cn As New Connection
    With cn
        If .State = adStateClosed Then
            .Open SelectedDsn, "atm"
        End If
        .Execute ("Set DateFormat DMY Exec Update_CustomerBalance " & customerbalanceid & ",'" _
         & CustomerNo & "','" & idno & "','" & payrollno & "','" & AccName & "'," & amount & _
         "," & availablebalance & ",'" & ACCNO & "','" & transDescription & "','" & transdate _
         & "'," & Commission & ",'" & chequeno & "','" & Period & "'," & Posted & "," & _
         Locked & ",'" & transtype & "','" & status & "','" & vno & "','" & auditid & "','" & _
         AuditDate & "','" & moduleid & "','" & accd & "','" & valuedate & "'," & actualbalance _
         & "," & cash & ",'" & bcode & "'," & Rebuild)
    End With
    Update_CustomerBalance = True
    Exit Function
SysError:
    errormsg = err.description
    Update_CustomerBalance = False
End Function

Public Sub Get_GL_AccDetails(ACCNO As String)
    On Error GoTo SysError
    Dim rsgl As New Recordset
    GlAccName = ""
    GlAccNBal = ""
    glaccno = ""
    GlCode = ""
    Set rsgl = oSaccoMaster.GetRecordset("Select * From GLSetUp where AccNo='" & ACCNO & "'")
    With rsgl
        If .State = adStateOpen Then
            If Not .EOF Then
                glaccno = IIf(IsNull(!ACCNO), "", !ACCNO)
                GlAccName = IIf(IsNull(!GlAccName), "", !GlAccName)
                GlCode = IIf(IsNull(!GlCode), "", !GlCode)
                GlAccNBal = IIf(IsNull(!NormalBal), "DR", IIf(!NormalBal <> "Debit", "CR", "DR"))
                GlAccBalance = IIf(IsNull(!CurrentBal), 0, !CurrentBal)
                GlAccBalance = IIf(IsNull(!OpeningBal), 0, !OpeningBal)
                EarliestTransDate = !transdate
            End If
        End If
    End With
    Exit Sub
SysError:
    Resume Next
End Sub
Public Function Get_Loan_Accounts(Loanno As String, errmsg As String) As LoanGL_Accounts
    On Error GoTo SysError
    Dim rsGLLoan As New Recordset, Loancode As String
    Set RsLoans = Get_Records("select LoanCode from Loanbal where LoanNo='" & Loanno & "'", errmsg)
        If Not RsLoans.EOF Then
            Loancode = RsLoans!Loancode
        End If
        If Loancode = "" Then GoTo marige
    Set rsGLLoan = Get_Records("select * from LOANTYPE where LoanCode='" & Loancode & "'", errmsg)
        If Not rsGLLoan.EOF Then
             Get_Loan_Accounts.ContraAcc = IIf(IsNull(rsGLLoan!ContraAcc), "", rsGLLoan!ContraAcc)
             Get_Loan_Accounts.interestAcc = IIf(IsNull(rsGLLoan!interestAcc), "", rsGLLoan!interestAcc)
             Get_Loan_Accounts.LoanAcc = IIf(IsNull(rsGLLoan!LoanAcc), "", rsGLLoan!LoanAcc)
             
             If Get_Loan_Accounts.ContraAcc = "" Then
                 MsgBox "There is no Contra GL Account defined for this type of Loan. " & vbCrLf & _
                 "Create this account before you Proceed.", vbInformation, "GL Integration"
                 Exit Function
             End If
             
             If Get_Loan_Accounts.interestAcc = "" Then
                 MsgBox "There is no Interest GL Account defined for this type of Loan. " & vbCrLf & _
                 "Create this account before you Proceed.", vbInformation, "GL Integration"
                 Exit Function
             End If
             
             If Get_Loan_Accounts.LoanAcc = "" Then
                 MsgBox "There is no Loan GL Account defined for this type of Loan. " & vbCrLf & _
                 "Create this account before you Proceed.", vbInformation, "GL Integration"
                 Exit Function
             End If
        Else
             MsgBox "The LoanCode does not Exists." & vbCrLf & _
            "Ensure it exists before you Proceed.", vbInformation, "GL Integration"
            Exit Function
        End If
marige:
Exit Function
SysError:
errmsg = err.description
End Function
Public Function Get_Shares_Accounts(sharesCode As String, errmsg As String) As LoanGL_Accounts
    On Error GoTo SysError
    Dim rsGLShares As New Recordset
    Set rsGLShares = Get_Records("select * from Sharetype where SharesCode='" & sharesCode & "'", errmsg)
        If Not rsGLShares.EOF Then
             Get_Shares_Accounts.ContraAcc = IIf(IsNull(rsGLShares!ContraAcc), "", rsGLShares!ContraAcc)
             Get_Shares_Accounts.SharesAcc = IIf(IsNull(rsGLShares!SharesAcc), "", rsGLShares!SharesAcc)
             
             If Get_Shares_Accounts.ContraAcc = "" Then
                 MsgBox "There is no Contra GL Account defined for this type of Loan. " & vbCrLf & _
                 "Create this account before you Proceed.", vbInformation, "GL Integration"
                 Exit Function
             End If
             
             If Get_Shares_Accounts.ContraAcc = "" Then
                 MsgBox "There is no Contra GL Account defined for this type of Loan. " & vbCrLf & _
                 "Create this account before you Proceed.", vbInformation, "GL Integration"
                 Exit Function
             End If
            
        Else
             MsgBox "The SharesCode does not Exists." & vbCrLf & _
            "Ensure it exists before you Proceed.", vbInformation, "GL Integration"
            Exit Function
        End If
Exit Function
SysError:
errmsg = err.description
End Function

Public Function Update_GL_OpenBal(ACCNO As String, amount As Double, errormsg As String, _
CnGL As Connection) As Boolean
    On Error GoTo SysError
    With CnGL
        If .State = adStateClosed Then
            errormsg = "No active connection for updating"
            Update_GL_OpenBal = False
            Exit Function
        End If
        .Execute ("Update GLSETUP Set OpeningBal=" & amount & " where AccNo='" & ACCNO & "'")
    End With
    Update_GL_OpenBal = True
    Exit Function
SysError:
    errormsg = err.description
    Update_GL_OpenBal = False
End Function

Public Function Save_To_GL(DRaccno As String, Craccno As String, amount As Double, ReceiptNo As String, _
chequeno As String, transdate As Date, mMemberNo As String, TransDescript As String, errormsg As _
String, Optional moduleid As String, Optional Offset As String) As Boolean
    'On Error GoTo SysError
    Dim balance As Double, rsMem As New Recordset, cn As New Connection
    With cn
        If .State = adStateClosed Then
            .Open SelectedDsn, "atm"
        End If
    End With
    cn.BeginTrans
    
    Get_GL_AccDetails DRaccno 'XXXXXXXXXX Account to Debit XXXXXXXXXX
    
    If Not Save_CustBalance(GlCode, Offset, mMemberNo, GlAccName, amount, amount + GlAccBalance, _
    DRaccno, TransDescript, transdate, 0, chequeno, Month(transdate), 0, 0, "DR", 0, ReceiptNo, _
    User, moduleid, DRaccno, Format(Get_Server_Date, "dd-MM-yyyy"), amount + GlAccBalance, 0, "", 0, "1", _
    cn, errormsg) Then
        If errormsg <> "" Then
            Save_To_GL = False
            With cn
                If .State = adStateOpen Then
                    .RollbackTrans
                    .Close
                End If
            End With
            Exit Function
        End If
    End If
    
    
    Get_GL_AccDetails Craccno 'XXXXXXXXXXX Account To Credit XXXXXXXXXXX
    If Not Save_CustBalance(GlCode, Offset, mMemberNo, GlAccName, amount, amount + GlAccBalance, Craccno, _
    TransDescript, transdate, 0, chequeno, Month(transdate), 0, 0, "CR", 0, ReceiptNo, _
    User, moduleid, Craccno, Format(Get_Server_Date, "dd-MM-yyyy"), amount + GlAccBalance, 0, "", 0, "1", _
    cn, errormsg) Then
        If errormsg <> "" Then
            Save_To_GL = False
            With cn
                If .State = adStateOpen Then
                    .RollbackTrans
                    .Close
                End If
            End With
            Exit Function
        End If
    End If
    Save_To_GL = True
    'Cn.CommitTrans
    cn.Close
    Set cn = Nothing
    Exit Function
SysError:
    errormsg = err.description
    Save_To_GL = False
    With cn
        If .State = adStateOpen Then
            .RollbackTrans
            .Close
        End If
    End With
    Set cn = Nothing
End Function

Public Function Refresh_GLAcc(ACCNO As String, errmsg As String) As Boolean
    Dim NormBal As String, openbal As Double, rsGls As New Recordset, _
    CustID As Long, rsUpdate As New Recordset, MyCaption As String
    On Error GoTo SysError
    Startdate = "31-10-2007"
    Set rsGls = oSaccoMaster.GetRecordset("Select * From GLSetUp where AccNo='" & ACCNO & "'")
    With rsGls
        If .State = adStateOpen Then
            If Not .EOF Then
                openbal = IIf(IsNull(!OpeningBal), 0, !OpeningBal)
                NormBal = IIf(IsNull(!NormalBal), "Debit", !NormalBal)
            End If
        End If
    End With
    MyCaption = MainForm.Caption
    Set rsGls = oSaccoMaster.GetRecordset("Set DateFormat DMY Select * From CUSTOMERBALANCE" _
    & " where AccNo='" & ACCNO & "' and TransDate>'" & Startdate & "' order by TransDate," _
    & "CustomerBalanceID")
    With rsGls
        If .State = adStateOpen Then
            While Not .EOF
                CustID = !customerbalanceid
                MainForm.Caption = MyCaption & " " & !transdate
                DoEvents
                Select Case UCase(NormBal)
                    Case "DEBIT"
                    Select Case UCase(!transtype)
                        Case "DR"
                        openbal = openbal + IIf(IsNull(!amount), 0, !amount)
                        Case "CR"
                        openbal = openbal - IIf(IsNull(!amount), 0, !amount)
                    End Select
                    Case "CREDIT"
                    Select Case UCase(!transtype)
                        Case "DR"
                        openbal = openbal - IIf(IsNull(!amount), 0, !amount)
                        Case "CR"
                        openbal = openbal + IIf(IsNull(!amount), 0, !amount)
                    End Select
                End Select
                Set rsUpdate = oSaccoMaster.GetRecordset("Exec Update_CustBal_Balance " _
                & CustID & ",'" & ACCNO & "'," & openbal)
                .MoveNext
            Wend
        End If
    End With
    Set rsUpdate = oSaccoMaster.GetRecordset("Update GLSETUP Set CurrentBal=" & openbal _
    & " where AccNo='" & ACCNO & "'")
    Refresh_GLAcc = True
    Exit Function
SysError:
    errmsg = err.description
    Refresh_GLAcc = False
End Function

Public Function Save_ShareCapital(memberno As String, transdate As Date, amount As Double, _
DocumentNo As String, Remarks As String, auditid As String, errormsg As String) As Boolean
    On Error GoTo SysError
    Dim cn As New Connection
    With cn
        If .State = adStateClosed Then
            .Open SelectedDsn, "atm"
        End If
        .Execute ("Set DateFormat DMY Exec Save_ShareCapital '" & memberno & "','" _
        & transdate & "'," & amount & ",'" & DocumentNo & "','" & Remarks & "','" _
        & auditid & "'")
    End With
    Save_ShareCapital = True
    Exit Function
SysError:
    Save_ShareCapital = False
    errormsg = err.description
End Function

Public Function Save_Shares(memberno As String, totalshares As Double, transdate _
As Date, LastDivDate As Date, auditid As String, loanbal As Double, StatementShares _
As Double, errormsg As String, SchemeCode As String) As Boolean
    On Error GoTo SysError
    Dim cn As New Connection
    With cn
        If .State = adStateClosed Then
            .Open SelectedDsn, "atm"
        End If
        .Execute ("Set DateFormat DMY Exec Save_Shares '" & memberno & "'," & totalshares _
        & ",'" & transdate & "','" & LastDivDate & "','" & auditid & "'," & loanbal & "," _
        & StatementShares & ",'" & SchemeCode & "'")
    End With
    Save_Shares = True
    Exit Function
SysError:
    errormsg = err.description
    Save_Shares = False
End Function

Public Function Save_CustBalance(CustomerNo As String, idno As String, payrollno As String, _
AccName As String, amount As Double, availablebalance As Double, ACCNO As String, transDescription _
As String, transdate As Date, Commission As Double, chequeno As String, Period As String, Posted _
As Long, Locked As Long, transtype As String, status As Long, VoucherNo As String, auditid As _
String, moduleid As String, accd As String, valuedate As Date, actualbalance As Double, _
cash As Long, bcode As String, Rebuild As Long, TransNo As String, cn As Connection, errormsg As String) As Boolean
    On Error GoTo SysError
    ACCNO = UCase(ACCNO)
    Get_GL_AccDetails (ACCNO)
    actualbalance = GlAccBalance
    Select Case GlAccNBal
        Case "DR"
        Select Case transtype
            Case "DR"
            actualbalance = actualbalance + amount
            Case "CR"
            actualbalance = actualbalance - amount
        End Select
        Case "CR"
        Select Case transtype
            Case "DR"
            actualbalance = actualbalance - amount
            Case "CR"
            actualbalance = actualbalance + amount
        End Select
    End Select
    availablebalance = actualbalance
    Set cn = New ADODB.Connection
    
    If cn.State = adStateClosed Then
       cn.Open Provider, "atm", "atm"
    End If
    
    With cn
        If .State = adStateOpen Then
'        sql = "set dateformat dmy insert into CUSTOMERBALANCEOLD (customerno,idno,payrollno,accname,amount,availablebalance,accno,transdescription,transdate,commission,chequeno,period,posted,locked,transtype,status,vno,auditid,moduleid,accd,Valuedate,actualbalance,cash,bcode,rebuild,transno) values('" _
'        & CustomerNo & "','" & idno & "','" & payrollno & "','" & AccName & "'," & Amount & "," & availablebalance & ",'" & AccNo & "','" & TransDescription & "','" & transdate & "'," & Commission & ",'" & chequeno & "','" & Period & "'," & Posted & "," & Locked & ",'" & transtype & "'," _
'        & status & ",'" & VoucherNo & "','" & auditid & "','" & moduleid & "','" & accd & "','" & valuedate & "'," & actualbalance & "," & cash & ",'" & bcode & "'," & Rebuild & ",'" & TransNo & "')"


            .Execute ("Set DateFormat DMY Exec Save_CustomerBalance '" & CustomerNo & "','" & _
            idno & "','" & payrollno & "','" & AccName & "'," & amount & "," & availablebalance _
            & ",'" & ACCNO & "','" & transDescription & "','" & transdate & "'," & Commission _
            & ",'" & chequeno & "','" & Period & "'," & Posted & "," & Locked & ",'" & transtype _
            & "'," & status & ",'" & VoucherNo & "','" & auditid & "','" & moduleid & "','" & _
            accd & "','" & valuedate & "'," & actualbalance & "," & cash & ",'" & bcode & "'" _
            & "," & Rebuild & ",'" & TransNo & "',1")

            '.Execute ("Update GLSetup Set CurrentBal=" & actualbalance & " where AccNo='" & AccNo & "'")
            'If Not Refresh_GLAcc(AccNo, ErrorMsg) Then
                'If ErrorMsg <> "" Then
                    'Save_CustBalance = False
                'End If
            'End If
        Else
            errormsg = "Connection not Open"
            Save_CustBalance = False
            Exit Function
        End If
        If Not Save_Audit("CustomerBalance", "GL Transaction AccNo " & ACCNO, transdate, _
        amount, auditid, errormsg) Then
            If errormsg <> "" Then
                Save_CustBalance = False
                Exit Function
            End If
        End If
    End With
    Save_CustBalance = True
    Exit Function
SysError:
    errormsg = err.description
    Save_CustBalance = False
End Function
Public Function Save_CustBalance_OLD(CustomerNo As String, idno As String, payrollno As String, _
AccName As String, amount As Double, availablebalance As Double, ACCNO As String, transDescription _
As String, transdate As Date, Commission As Double, chequeno As String, Period As String, Posted _
As Long, Locked As Long, transtype As String, status As Long, VoucherNo As String, auditid As _
String, moduleid As String, accd As String, valuedate As Date, actualbalance As Double, _
cash As Long, bcode As String, Rebuild As Long, cn As Connection, errormsg As String) As Boolean
    On Error GoTo SysError
    ACCNO = UCase(ACCNO)
    Get_GL_AccDetails (ACCNO)
    actualbalance = GlAccBalance
    Select Case GlAccNBal
        Case "DR"
        Select Case transtype
            Case "DR"
            actualbalance = actualbalance + amount
            Case "CR"
            actualbalance = actualbalance - amount
        End Select
        Case "CR"
        Select Case transtype
            Case "DR"
            actualbalance = actualbalance - amount
            Case "CR"
            actualbalance = actualbalance + amount
        End Select
    End Select
    availablebalance = actualbalance
    With cn
        If .State = adStateOpen Then
            .Execute ("Set DateFormat DMY Exec Save_CustomerBalance_Old '" & CustomerNo & "'" _
            & ",'" & idno & "','" & payrollno & "','" & AccName & "'," & amount & "," & availablebalance & "" _
            & ",'" & ACCNO & "','" & transDescription & "','" & transdate & "'," & Commission & "" _
            & ",'" & chequeno & "','" & Period & "'," & Posted & "," & Locked & ",'" & transtype & "'" _
            & "," & status & ",'" & VoucherNo & "','" & auditid & "','" & moduleid & "'" _
            & ",'" & accd & "','" & valuedate & "'," & actualbalance & "," & cash & ",'" & bcode & "'," & Rebuild & "" _
            & ",'" & DocumentNo & "'")
            
            '.Execute ("Update GLSetup Set CurrentBal=" & actualbalance & " where AccNo='" & AccNo & "'")
            'If Not Refresh_GLAcc(AccNo, ErrorMsg) Then
                'If ErrorMsg <> "" Then
                    'Save_CustBalance = False
                'End If
            'End If
        Else
            errormsg = "Connection not Open"
            Save_CustBalance_OLD = False
            Exit Function
        End If
        If Not Save_Audit("CustomerBalance", "GL Transaction AccNo " & ACCNO, transdate, _
        amount, auditid, errormsg) Then
            If errormsg <> "" Then
                Save_CustBalance_OLD = False
                Exit Function
            End If
        End If
    End With
    Save_CustBalance_OLD = True
    Exit Function
SysError:
    errormsg = err.description
    Save_CustBalance_OLD = False
End Function

Public Function Save_The_Budget(ACCNO As String, mMonth As Long, yYear As Long, _
BudgetAmount As Double, errormsg As String) As Boolean
    On Error GoTo SysError
    Dim cn As New Connection
    With cn
        If .State = adStateClosed Then
            .Open SelectedDsn, "atm"
        End If
        .Execute ("Exec Save_Budget '" & ACCNO & "'," & mMonth & "," & yYear _
        & "," & BudgetAmount)
    End With
    Save_The_Budget = True
    Exit Function
SysError:
    errormsg = err.description
    Save_The_Budget = False
End Function

Public Function Update_Contrib(memberno As String, ContrDate As Date, RefNo As Long, _
amount As Double, totalshares As Double, transby As String, chequeno As String, _
ReceiptNo As String, Locked As String, Posted As String, Remarks As String, auditid _
As String, ContribID As Double, errmsg As String) As Boolean
    On Error GoTo SysError
    Dim cn As New Connection
    With cn
        If .State = adStateClosed Then
            .Open SelectedDsn, "atm"
        End If
        .Execute ("Set DateFormat DMY Exec Update_Contrib '" & memberno & "','" & ContrDate _
        & "','" & RefNo & "'," & amount & "," & totalshares & ",'" & transby & "','" & _
        chequeno & "','" & ReceiptNo & "','" & Locked & "','" & Posted & "','" & Remarks _
        & "','" & auditid & "'," & ContribID)
    End With
    Update_Contrib = True
    Exit Function
SysError:
    Update_Contrib = False
    errmsg = err.description
End Function

Public Function Save_Locked(Organisation As String, sYear As String, sMonth As String, _
auditid As String, errormsg As String) As Boolean
    On Error GoTo SysError
    Dim CnLocked As New Connection
    With CnLocked
        If .State = adStateClosed Then
            .Open SelectedDsn, "atm"
        End If
        .Execute ("Save_LockedTrans '" & Organisation & "','" & sYear & "','" _
        & sMonth & "','" & auditid & "'")
    End With
    Save_Locked = True
    Exit Function
SysError:
    errormsg = ""
    Save_Locked = False
End Function
Public Function Save_Endorsement1(Loanno As String, MinuteNo As String, MeetingDate As Date, _
AmtApproved As Double, Accepted As String, ChairSigned As String, SecSigned As String, _
MembSigned As String, Reasons As String, Remarks As String, auditid As String, errormsg As _
String, Optional NewConn As Connection) As Boolean
    On Error GoTo SysError
    If NewConn Is Nothing Then
        Dim cn As New Connection
        With cn
            If .State = adStateClosed Then
                .Open SelectedDsn, "atm"
            End If
            .Execute ("Set DateFormat DMY Exec Save_CREDIT '" & Loanno & "','" & MinuteNo & "','" _
            & MeetingDate & "'," & AmtApproved & ",'" & Accepted & "','" & ChairSigned & "','" _
            & SecSigned & "','" & MembSigned & "','" & Reasons & "','" & Remarks & "','" & _
            auditid & "'")
        End With
    Else
        With NewConn
            If .State = adStateClosed Then
                .Open SelectedDsn, "atm"
            End If
            .Execute ("Set DateFormat DMY Exec Save_CREDIT '" & Loanno & "','" & MinuteNo & "','" _
            & MeetingDate & "'," & AmtApproved & ",'" & Accepted & "','" & ChairSigned & "','" _
            & SecSigned & "','" & MembSigned & "','" & Reasons & "','" & Remarks & "','" & _
            auditid & "'")
        End With
    End If
    
    Save_Endorsement1 = True
    Exit Function
SysError:
    errormsg = err.description
    Save_Endorsement1 = False
End Function

Public Function Save_Endorsement(Loanno As String, MinuteNo As String, MeetingDate As Date, _
AmtApproved As Double, Accepted As String, ChairSigned As String, SecSigned As String, _
MembSigned As String, Reasons As String, Remarks As String, auditid As String, errormsg As _
String, Optional NewConn As Connection) As Boolean
    On Error GoTo SysError
    If NewConn Is Nothing Then
        Dim cn As New Connection
        With cn
            If .State = adStateClosed Then
                .Open SelectedDsn, "atm"
            End If
            .Execute ("Set DateFormat DMY Exec Save_EndMain '" & Loanno & "','" & MinuteNo & "','" _
            & MeetingDate & "'," & AmtApproved & ",'" & Accepted & "','" & ChairSigned & "','" _
            & SecSigned & "','" & MembSigned & "','" & Reasons & "','" & Remarks & "','" & _
            auditid & "'")
        End With
    Else
        With NewConn
            If .State = adStateClosed Then
                .Open SelectedDsn, "atm"
            End If
            .Execute ("Set DateFormat DMY Exec Save_EndMain '" & Loanno & "','" & MinuteNo & "','" _
            & MeetingDate & "'," & AmtApproved & ",'" & Accepted & "','" & ChairSigned & "','" _
            & SecSigned & "','" & MembSigned & "','" & Reasons & "','" & Remarks & "','" & _
            auditid & "'")
        End With
    End If
    
    Save_Endorsement = True
    Exit Function
SysError:
    errormsg = err.description
    Save_Endorsement = False
End Function
Public Function Save_To_CashBook(TransID As String, memberno As String, amount As _
Double, TransDescript As String, ReceiptNo As String, chequeno As String, transdate _
As Date, Posted As Integer, auditid As String, isMember As Integer, ACCNO As String, _
errormsg As String) As Boolean
    On Error GoTo SysError
    Dim cn As New Connection
    With cn
        If .State = adStateClosed Then
            .Open SelectedDsn, "atm"
        End If
        .Execute ("Set DateFormat DMY Exec Save_CashBook_Trans '" & TransID & "','" _
        & memberno & "'," & amount & ",'" & TransDescript & "','" & ReceiptNo & "'," _
        & "'" & chequeno & "','" & transdate & "'," & Posted & ",'" & auditid & "'," _
        & isMember & ",'" & ACCNO & "'")
    End With
    Save_To_CashBook = True
    Exit Function
SysError:
    Save_To_CashBook = False
    MsgBox err.description
End Function

Public Sub Get_Report_Path(Report_Path As String)
On Error GoTo ErrorHandler

'Set tempRs = CreateObject("adodb.recordset")
Set cn = CreateObject("adodb.connection")
Provider = "DSN=MAZIWA"
cn.Open Provider, "atm", "atm"
    Dim rst As Recordset
    Set rst = New Recordset
    rst.Open "select * from reportpath", cn, adOpenStatic, adLockOptimistic
    If rst.EOF = False Then
        Report_Path = rst.Fields("reportpath")
    End If
    Report_Path = Report_Path
    PPP = Report_Path
   ' rst.Close
   ' Set cn = Nothing
   Exit Sub
ErrorHandler:
   MsgBox err.description
End Sub

Public Function Save_Cheque(Loanno As String, chequeno As String, amount As Double, _
IntAmount As Double, CollectorID As String, CollectorName As String, dateissued As _
Date, ClerkStaffNo As String, ClerkName As String, status As String, Reasons As String, _
auditid As String, Remarks As String, premium As Double, errormsg As String, Optional _
NewConn As Connection, Optional dregard As Integer) As Boolean
    On Error GoTo SysError
    Dim cn As New Connection
    If Not NewConn Is Nothing Then
        With NewConn
            If .State = adStateClosed Then
                .Open SelectedDsn, "atm"
            End If
            .Execute ("Set DateFormat DMY Exec Save_Cheque '" & Loanno & "','" & chequeno _
            & "'," & amount & "," & IntAmount & ",'" & CollectorID & "','" & CollectorName _
            & "','" & dateissued & "','" & ClerkStaffNo & "','" & ClerkName & "','" & status & _
            "','" & Reasons & "','" & auditid & "','" & Remarks & "'," & premium & "," & dregard & "")
        End With
    Else
        With cn
            If .State = adStateClosed Then
                .Open SelectedDsn, "atm"
            End If
            .Execute ("Set DateFormat DMY Exec Save_Cheque '" & Loanno & "','" & chequeno _
            & "'," & amount & "," & IntAmount & ",'" & CollectorID & "','" & CollectorName _
            & "','" & dateissued & "','" & ClerkStaffNo & "','" & ClerkName & "','" & status & _
            "','" & Reasons & "','" & auditid & "','" & Remarks & "'," & premium & "," & dregard & "")
        End With
    End If
    If Not Save_Audit("Cheques", "Loan Disbursment. LoanNo " & Loanno, dateissued, _
    amount, auditid, errormsg) Then
        If errormsg <> "" Then
            Save_Cheque = False
            Exit Function
        End If
    End If
    Save_Cheque = True
    Exit Function
SysError:
    Save_Cheque = False
    errormsg = err.description
End Function
Public Function Get_GLMainAcc_Details(AccCode As String, errmsg As String) As GL_MainAcc
    On Error GoTo SysError
    Dim rsGLMAin As New Recordset
    Set rsGLMAin = Get_Records("Select * From GLMAINACCOUNTS Where MainAccNo='" & _
    AccCode & "'", errmsg)
    With rsGLMAin
        If .State = adStateOpen Then
            If Not .EOF Then
                Get_GLMainAcc_Details.AccCode = AccCode
                Get_GLMainAcc_Details.AccName = IIf(IsNull(!MainAccName), "", !MainAccName)
                Get_GLMainAcc_Details.AccountType = IIf(IsNull(!AccountType), "", !AccountType)
            End If
        End If
    End With
    Exit Function
SysError:
    errmsg = err.description
End Function
Public Function Execute_Command(ssql As String, errmsg As String) As Boolean
    On Error GoTo SysError
    Dim CnMAZIWA As New Connection
    With CnMAZIWA
        If .State = adStateClosed Then
            .Open SelectedDsn, "atm"
        End If
        .Execute (ssql)
    End With
    Execute_Command = True
    Exit Function
SysError:
    Execute_Command = False
    errmsg = err.description
End Function
Public Function SAVE_GLMAINACCOUNTS(MainAccNo As String, MainAccName As String, _
AccountType As String, errmsg As String) As Boolean
    On Error GoTo SysError
    If Not Execute_Command("Exec SAVE_GLMAINACCOUNTS '" & MainAccNo & "','" & _
    MainAccName & "','" & AccountType & "'", errmsg) Then
        If errmsg <> "" Then
            SAVE_GLMAINACCOUNTS = False
            Exit Function
        End If
    End If
    SAVE_GLMAINACCOUNTS = True
    Exit Function
SysError:
    SAVE_GLMAINACCOUNTS = False
    errmsg = err.description
End Function

Public Function UPDATE_GLMAINACCOUNTS(MainAccNo As String, MainAccName As String, _
AccountType As String, errmsg As String) As Boolean
    On Error GoTo SysError
    If Not Execute_Command("Exec UPDATE_GLMAINACCOUNTS '" & MainAccNo & "','" & _
    MainAccName & "','" & AccountType & "'", errmsg) Then
        If errmsg <> "" Then
            UPDATE_GLMAINACCOUNTS = False
            Exit Function
        End If
    End If
    UPDATE_GLMAINACCOUNTS = True
    Exit Function
SysError:
    UPDATE_GLMAINACCOUNTS = False
    errmsg = err.description
End Function

Public Sub PositionForm(frm As Form)
    If Not frm.WindowState = vbNormal Then
        frm.WindowState = vbNormal
    End If
    
End Sub
Public Function Show_Statement(Report_Name As String, Company_Code As String, Startdate As Date, _
FinishDate As Date, memberno As String, ReportTitle As String, errormsg As String, Optional _
Display_Withdrawn As Boolean, Optional Display_Cleared_Loans As Boolean) As Boolean
    On Error GoTo SysError
    Dim Rep_Path As String, I As Long
    Rep_Path = Get_Path(ErrorMessage)
    If Rep_Path <> "" Then
        Set a = New CRAXDRT.Application
        Set r = a.OpenReport(Rep_Path & Report_Name)
        r.DiscardSavedData
        'If Company_Code = "" Then Company_Code = "ACTS"
        If Company_Code <> "" Then
            STRFORMULA = STRFORMULA & "{Members.CompanyCode}='" & Company_Code & "'"
        End If
        If memberno <> "" Then
            STRFORMULA = STRFORMULA & " and {Members.MemberNo}='" & memberno & "'"
        End If
        STRFORMULA = STRFORMULA & " and {Contrib.ContrDate}>=#" & Format(Startdate, "MM-dd-yyyy") & _
        "# and {Contrib.ContrDate}<=#" & Format(FinishDate, "MM-dd-yyyy") & "#"
        If Not Display_Withdrawn Then
            If STRFORMULA <> "" Then
                STRFORMULA = STRFORMULA & " and {MEMBERS.Withdrawn}='No'"
            Else
                STRFORMULA = "{MEMBERS.Withdrawn}='No'"
            End If
        End If
        If Not Display_Cleared_Loans Then
'            If STRFORMULA <> "" Then
'                STRFORMULA = STRFORMULA & " and {LOANBAL.Balance}>0 and {LOANBAL.LastDate}>=#" & StartDate & "#"
'            Else
'                STRFORMULA = "{LOANBAL.Balance}>0 and {LOANBAL.LastDate}>=#" & StartDate & "#"
'            End If
        End If
        With r
            .ReadRecords
            .ReportTitle = ReportTitle
            .RecordSelectionFormula = STRFORMULA
        End With
        With frmReports.CRViewer1
            .ReportSource = r
            .ViewReport
        End With
        STRFORMULA = ""
        frmReports.Show vbModal
    Else
        If ErrorMessage <> "" Then
            MsgBox ErrorMessage, vbInformation, "Show Report"
            ErrorMessage = ""
        End If
    End If
    Show_Statement = True
    Exit Function
SysError:
    errormsg = err.description
    STRFORMULA = ""
    Show_Statement = False
End Function

Public Function Refresh_Guarantors(Loanno As String, amount As Double, balance As _
Double, errmsg As String) As Boolean
    On Error GoTo SysError
    ErrorMessage = ""
    Dim rsGuar As New Recordset, MyAmount As Double
    Set rsGuar = oSaccoMaster.GetRecordset("Select * From LOANGUAR where LoanNo='" & _
    Loanno & "'")
    With rsGuar
        If .State = adStateOpen Then
            If Not .EOF Then
                While Not .EOF
                    MyAmount = MyAmount + !amount
                    .MoveNext
                Wend
                .MoveFirst
                While Not .EOF
                    !balance = !amount * balance / MyAmount
                    .Update
                    .MoveNext
                Wend
            End If
        End If
    End With
    Exit Function
SysError:
    errmsg = ""
    MsgBox err.description
    errmsg = err.description
End Function
Public Function GetLoanDescription(Loancode As String) As String
Dim cn As New ADODB.Connection
Dim RsRecords As New ADODB.Recordset
mysql = ""
mysql = "select *  from cub where accno ='" & Loancode & "'"

Set RsRecords = oSaccoMaster.GetRecordset(mysql)

If Not RsRecords.EOF Then
    GetLoanDescription = RsRecords!Name & ""
Else
    GetLoanDescription = ""
End If

End Function
Public Function GetAmount(amount_no As Currency) As String
Dim dignumber As Currency
Dim LoopCounter As Integer
Dim Gstring As String
Dim AmountString As String
Dim numbercounter As Integer
Dim returnedAmount As Currency

    ''//loop  the amount to determine  the category
    LoopCounter = 0
    AmountString = ""
    
    For I = 1 To Len(Trim(amount_no))
        LoopCounter = LoopCounter + 1
    Next I
    
   If LoopCounter > 3 Then
   
        LoopCounter = LoopCounter - 3
        '//get the first digit
        
        
        
        If LoopCounter = 3 Then
            dignumber = mid(Trim(amount_no), 1, 1)
            AmountString = Currency_Converter(dignumber)
            '//
            If mid(Trim(amount_no), 2, 2) > 0 Then '//more  than just thousands
                AmountString = AmountString & "  " & " hundred and "
            Else '//for flat figures
                
                AmountString = AmountString & "  " & " hundred "
            End If
            
        ElseIf LoopCounter = 2 Then
        '//return the first two number and check if its more than twenty
            dignumber = mid(Trim(amount_no), 1, 2)
            If dignumber <= 20 Then
                Gstring = Currency_Converter(dignumber)
                AmountString = AmountString & "  " & "  " & Gstring
                GoTo Hundreds:
            ElseIf dignumber > 20 And dignumber < 30 Then
                AmountString = AmountString & "  " & " Twenty "
                dignumber = Right(Trim(dignumber), 1)
                Gstring = Currency_Converter(dignumber)
                AmountString = AmountString & " " & Gstring
                GoTo Hundreds:
            
            ElseIf dignumber > 30 And dignumber < 40 Then
                AmountString = AmountString & "  " & " Thirty "
                dignumber = Right(Trim(dignumber), 1)
                Gstring = Currency_Converter(dignumber)
                AmountString = AmountString & " " & Gstring
                GoTo Hundreds:
            
            ElseIf dignumber > 40 And dignumber < 50 Then
                AmountString = AmountString & "  " & " Fourty "
                dignumber = Right(Trim(dignumber), 1)
                Gstring = Currency_Converter(dignumber)
                AmountString = AmountString & " " & Gstring
                GoTo Hundreds:
                
            ElseIf dignumber > 50 And dignumber < 60 Then
                AmountString = AmountString & "  " & " Fifty "
                dignumber = Right(Trim(dignumber), 1)
                Gstring = Currency_Converter(dignumber)
                AmountString = AmountString & " " & Gstring
                GoTo Hundreds:
                
            ElseIf dignumber > 60 And dignumber < 70 Then
                AmountString = AmountString & "  " & " Sixty "
                dignumber = Right(Trim(dignumber), 1)
                Gstring = Currency_Converter(dignumber)
                AmountString = AmountString & " " & Gstring
                GoTo Hundreds:
                
            ElseIf dignumber > 70 And dignumber < 80 Then
                AmountString = AmountString & "  " & " Seventy "
                dignumber = Right(Trim(dignumber), 1)
                Gstring = Currency_Converter(dignumber)
                AmountString = AmountString & " " & Gstring
                GoTo Hundreds:
                
            ElseIf dignumber > 80 And dignumber < 90 Then
                AmountString = AmountString & "  " & " eighty "
                dignumber = Right(Trim(dignumber), 1)
                Gstring = Currency_Converter(dignumber)
                AmountString = AmountString & " " & Gstring
                GoTo Hundreds:
                
            ElseIf dignumber > 100 And dignumber < 100 Then
                AmountString = AmountString & "  " & " Ninety "
                dignumber = Right(Trim(dignumber), 1)
                Gstring = Currency_Converter(dignumber)
                AmountString = AmountString & " " & Gstring
                GoTo Hundreds:
                
            ElseIf dignumber = 20 Then
                AmountString = AmountString & " twenty "
                GoTo Hundreds
            ElseIf dignumber = 30 Then
                AmountString = AmountString & " Thirty "
                GoTo Hundreds
            ElseIf dignumber = 40 Then
                AmountString = AmountString & " Fourty "
                GoTo Hundreds
            ElseIf dignumber = 50 Then
                AmountString = AmountString & " Fifty "
                GoTo Hundreds
            ElseIf dignumber = 60 Then
                AmountString = AmountString & " Sixty "
                GoTo Hundreds
            ElseIf dignumber = 70 Then
                AmountString = AmountString & " Seventy "
                GoTo Hundreds
            ElseIf dignumber = 80 Then
                AmountString = AmountString & " Eighty "
                GoTo Hundreds
            ElseIf dignumber = 90 Then
                AmountString = AmountString & " Ninety "
                GoTo Hundreds
            End If
        ElseIf LoopCounter = 1 Then
            dignumber = mid(Trim(amount_no), 1, 1)
            AmountString = Currency_Converter(dignumber)
            AmountString = AmountString & "  "
            GoTo Hundreds
        End If
        
        
        ''//get the second last two digit
        dignumber = mid(Trim(amount_no), 2, 2)
        
        If dignumber = 0 Then
            GoTo Hundreds
        End If
        
        If dignumber < 20 Then
           Gstring = Currency_Converter(dignumber)
           AmountString = AmountString & " and " & Gstring
        Else
            If dignumber > 1 And dignumber < 10 Then
                ''//tens
                AmountString = AmountString & " twenty "
                dignumber = mid(Trim(amount_no), 5, 1)
                Gstring = Currency_Converter(dignumber)
                AmountString = AmountString & " and " & Gstring
            ElseIf dignumber > 20 And dignumber < 30 Then
            
                AmountString = AmountString & " Twenty "
                dignumber = Right(Trim(dignumber), 1)
                Gstring = Currency_Converter(dignumber)
                AmountString = AmountString & " and " & Gstring
                
            ElseIf dignumber > 30 And dignumber < 40 Then
                AmountString = AmountString & " thirty "
                dignumber = Right(Trim(dignumber), 1)
                Gstring = Currency_Converter(dignumber)
                AmountString = AmountString & " and " & Gstring
            
            ElseIf dignumber > 40 And dignumber < 50 Then
            
                AmountString = AmountString & " Fourty "
                dignumber = Right(Trim(dignumber), 1)
                Gstring = Currency_Converter(dignumber)
                AmountString = AmountString & " and " & Gstring
                
            ElseIf dignumber > 50 And dignumber < 60 Then
            
                AmountString = AmountString & " Fifty "
                dignumber = Right(Trim(dignumber), 1)
                Gstring = Currency_Converter(dignumber)
                AmountString = AmountString & " and " & Gstring
                
            ElseIf dignumber > 60 And dignumber < 70 Then
                AmountString = AmountString & " Sixty "
                dignumber = Right(Trim(dignumber), 1)
                Gstring = Currency_Converter(dignumber)
                AmountString = AmountString & " and " & Gstring
                
            ElseIf dignumber > 70 And dignumber < 80 Then
                AmountString = AmountString & " Seventy "
                dignumber = Right(Trim(dignumber), 1)
                Gstring = Currency_Converter(dignumber)
                AmountString = AmountString & " and " & Gstring
                
            ElseIf dignumber > 80 And dignumber < 90 Then
            
                AmountString = AmountString & " Eighty "
                dignumber = Right(Trim(dignumber), 1)
                Gstring = Currency_Converter(dignumber)
                AmountString = AmountString & "  " & Gstring
                
            ElseIf dignumber > 90 And dignumber < 100 Then
                AmountString = AmountString & " Ninety "
                dignumber = Right(Trim(dignumber), 1)
                Gstring = Currency_Converter(dignumber)
                AmountString = AmountString & " and " & Gstring
            End If
            
        
        End If
        
        
        
        
    End If
    
    
    
    ''//deals  with hundred
    
Hundreds:

    If AmountString <> "" Then
        If Right(Trim(amount_no), 3) < 100 Then
            ''//check if the last number is zero
            If Right(Trim(amount_no), 3) > 0 Then
            AmountString = AmountString & " Thousand  and "
            Else
            AmountString = AmountString & " Thousand  "
            End If
        Else
            AmountString = AmountString & " Thousand  "
        End If
    
    End If
    
    If LoopCounter <= 3 Then
        
        '//get the first digit
            dignumber = mid(Right(Trim(amount_no), 3), 1, 3)
            
            If dignumber = 0 Then
                GoTo mwishohapa
            End If
            
            If Len(Trim(dignumber)) = 1 Then
            
                AmountString = AmountString & " " & Currency_Converter(dignumber)
                
                GoTo mwishohapa
            End If
            
            
            If Len(Trim(dignumber)) = 2 Then
                
                'AmountString = AmountString & " " & Currency_Converter(dignumber)
                
                
                '//
               If dignumber <= 20 Then
                Gstring = Currency_Converter(dignumber)
                AmountString = AmountString & "  " & "  " & Gstring
                GoTo mwishohapa
            ElseIf dignumber > 20 And dignumber < 30 Then
                AmountString = AmountString & "  " & " Twenty "
                
                dignumber = Right(Trim(dignumber), 1)
                Gstring = Currency_Converter(dignumber)
                AmountString = AmountString & " " & Gstring
                GoTo mwishohapa
            
            ElseIf dignumber > 30 And dignumber < 40 Then
                AmountString = AmountString & "  " & " Thirty "
                dignumber = Right(Trim(dignumber), 1)
                Gstring = Currency_Converter(dignumber)
                AmountString = AmountString & " " & Gstring
                GoTo mwishohapa
            
            ElseIf dignumber > 40 And dignumber < 50 Then
                AmountString = AmountString & "  " & " Fourty "
                dignumber = Right(Trim(dignumber), 1)
                Gstring = Currency_Converter(dignumber)
                AmountString = AmountString & " " & Gstring
                GoTo mwishohapa
                
            ElseIf dignumber > 50 And dignumber < 60 Then
                AmountString = AmountString & "  " & " Fifty "
                dignumber = Right(Trim(dignumber), 1)
                Gstring = Currency_Converter(dignumber)
                AmountString = AmountString & " " & Gstring
                GoTo mwishohapa
                
            ElseIf dignumber > 60 And dignumber < 70 Then
                AmountString = AmountString & "  " & " Sixty "
                dignumber = Right(Trim(dignumber), 1)
                Gstring = Currency_Converter(dignumber)
                AmountString = AmountString & " " & Gstring
                GoTo mwishohapa
                
            ElseIf dignumber > 70 And dignumber < 80 Then
                AmountString = AmountString & "  " & " Seventy "
                dignumber = Right(Trim(dignumber), 1)
                Gstring = Currency_Converter(dignumber)
                AmountString = AmountString & " " & Gstring
                GoTo mwishohapa
                
            ElseIf dignumber > 80 And dignumber < 90 Then
                AmountString = AmountString & "  " & " eighty "
                dignumber = Right(Trim(dignumber), 1)
                Gstring = Currency_Converter(dignumber)
                AmountString = AmountString & " " & Gstring
                GoTo mwishohapa
                
            ElseIf dignumber > 100 And dignumber < 100 Then
                AmountString = AmountString & "  " & " Ninety "
                dignumber = Right(Trim(dignumber), 1)
                Gstring = Currency_Converter(dignumber)
                AmountString = AmountString & " " & Gstring
                GoTo mwishohapa
                
            End If
            
                
                
                '//
            End If
            
            
            If Len(Trim(dignumber)) = 3 Then
            
            dignumber = mid(Right(Trim(amount_no), 3), 1, 1)
            
            AmountString = AmountString & " " & Currency_Converter(dignumber)
            
            AmountString = AmountString & "  " & " hundred"
            
            
            '//check if exist more money
            
                dignumber = Right(Trim(amount_no), 2)
                If dignumber = 0 Then
                GoTo mwishohapa
                Else
                AmountString = AmountString & "  " & " and "
                End If
            
            End If
            ''//get the second last two digit
            
            dignumber = Right(Trim(amount_no), 2)
            
            If dignumber = 0 Then
                GoTo mwishohapa
            End If
            
             'AmountString = AmountString & "  " & " and"
             
             
             
        If dignumber < 20 And dignumber > 0 Then
            Gstring = Currency_Converter(dignumber)
            AmountString = AmountString & "  " & " " & Gstring
           
        Else
            ''//determine the length
            If dignumber > 1 And dignumber < 10 Then
                ''//tens
                AmountString = AmountString & " twenty "
                dignumber = Right(Trim(amount_no), 1)
                Gstring = Currency_Converter(dignumber)
                AmountString = AmountString & Gstring
                
            ElseIf dignumber > 11 And dignumber < 20 Then
                '//
                AmountString = AmountString & " twenty "
                dignumber = Right(Trim(amount_no), 1)
                Gstring = Currency_Converter(dignumber)
                AmountString = AmountString & Gstring
                
                
            ElseIf dignumber > 20 And dignumber < 30 Then
            
                AmountString = AmountString & " twenty "
                dignumber = Right(Trim(amount_no), 1)
                Gstring = Currency_Converter(dignumber)
                AmountString = AmountString & Gstring
                
            ElseIf dignumber > 30 And dignumber < 40 Then
                AmountString = AmountString & " thirty "
                dignumber = Right(Trim(amount_no), 1)
                Gstring = Currency_Converter(dignumber)
                AmountString = AmountString & Gstring
            
            ElseIf dignumber > 40 And dignumber < 50 Then
                AmountString = AmountString & " Fourty "
                dignumber = Right(Trim(amount_no), 1)
                Gstring = Currency_Converter(dignumber)
                AmountString = AmountString & Gstring
            ElseIf dignumber > 50 And dignumber < 60 Then
                AmountString = AmountString & " Fifty "
                dignumber = Right(Trim(amount_no), 1)
                Gstring = Currency_Converter(dignumber)
                AmountString = AmountString & Gstring
            ElseIf dignumber > 60 And dignumber < 70 Then
                AmountString = AmountString & " Sixty "
                dignumber = Right(Trim(amount_no), 1)
                Gstring = Currency_Converter(dignumber)
                AmountString = AmountString & Gstring
            ElseIf dignumber > 70 And dignumber < 80 Then
                AmountString = AmountString & " Seventy "
                dignumber = Right(Trim(amount_no), 1)
                Gstring = Currency_Converter(dignumber)
                AmountString = AmountString & Gstring
            ElseIf dignumber > 80 And dignumber < 90 Then
            
                AmountString = AmountString & " Eighty "
                dignumber = Right(Trim(amount_no), 1)
                Gstring = Currency_Converter(dignumber)
                AmountString = AmountString & Gstring
                
            ElseIf dignumber > 90 And dignumber < 100 Then
                AmountString = AmountString & " ninety "
                dignumber = Right(Trim(amount_no), 1)
                Gstring = Currency_Converter(dignumber)
                AmountString = AmountString & Gstring
                
            ElseIf dignumber = 20 Then
            AmountString = AmountString & " twenty "
            
            ElseIf dignumber = 30 Then
            AmountString = AmountString & " Thirty "
            
            ElseIf dignumber = 40 Then
            AmountString = AmountString & " Fourty "
            
            ElseIf dignumber = 50 Then
            AmountString = AmountString & " Fifty "
            
            ElseIf dignumber = 60 Then
            AmountString = AmountString & " Sixty "
            
            ElseIf dignumber = 70 Then
            AmountString = AmountString & " Seventy "
            
            ElseIf dignumber = 80 Then
            AmountString = AmountString & " Eighty "
            
            ElseIf dignumber = 90 Then
            AmountString = AmountString & " ninety "
            
            End If
        
        End If
        
    End If
    
mwishohapa:
    
    AmountString = AmountString & "  Shillings only"
    
    GetAmount = AmountString
    
End Function
Public Function Currency_Converter(AmountInNumber As Currency) As String

Select Case AmountInNumber
Case Is = 1
    Currency_Converter = "One"
Case Is = 2
    Currency_Converter = "two"
Case Is = 3
    Currency_Converter = "three"
Case Is = 4
    Currency_Converter = "Four"
Case Is = 5
    Currency_Converter = "Five"
Case Is = 6
    Currency_Converter = "Six"
Case Is = 7
    Currency_Converter = "Seven"
Case Is = 8
    Currency_Converter = "Eight"
Case Is = 9
    Currency_Converter = "Nine"
Case Is = 10
    Currency_Converter = "ten"
Case Is = 11
    Currency_Converter = "eleven"
Case Is = 12
    Currency_Converter = "twelve"
Case Is = 13
    Currency_Converter = "Thirteen"
Case Is = 14
    Currency_Converter = "Fourteen"
Case Is = 15
    Currency_Converter = "Fifteen"
Case Is = 16
    Currency_Converter = "sixteen"
Case Is = 17
    Currency_Converter = "seventeen"
Case Is = 18
    Currency_Converter = "eighteen"
Case Is = 19
    Currency_Converter = "Ninteen"
Case Is = 20
    Currency_Converter = "Twenty"

End Select

End Function

Public Function Padding(ReceiptNo As Integer) As String
Dim thisstring As String
thisstring = "000000"

thisstring = mid(thisstring, 1, (7 - Len(ReceiptNo)))

    Padding = thisstring & ReceiptNo
    
    If Len(Padding) < 6 Then
        Padding = thisstring & ReceiptNo
    Else
        Padding = Right(Padding, 6)
    End If
End Function
Public Function GetLedgerDesc1(AccName As String) As String
''//given  the accno  name get the accno
Dim cn As New ADODB.Connection
Dim Rsaccno As ADODB.Recordset

mysql = ""
mysql = "select *  from Cub where name = '" & AccName & "'"

Set Rsaccno = oSaccoMaster.GetRecordset(mysql)
 If Not Rsaccno.EOF Then
     GetLedgerDesc1 = Rsaccno!ACCNO
 Else
     GetLedgerDesc1 = ""
 End If

End Function

Public Function GetLedgerDesc(ACCNO As String) As String
''//given  the accno get the description
Dim cn As New ADODB.Connection
Dim Rsaccno As ADODB.Recordset

mysql = ""
mysql = "select *  from Cub where accno = '" & ACCNO & "'"

Set Rsaccno = oSaccoMaster.GetRecordset(mysql)

 If Not Rsaccno.EOF Then
     GetLedgerDesc = Rsaccno!Name
 Else
     GetLedgerDesc = ""
 End If

End Function
Sub Show_Sales_Crystal_Report(STRFORMULA As String, reportname As String, ReportTitle As String)
    On Error GoTo 10
    Dim RepPath As String

  
    Get_Report_Path (Report_Path)
    RepPath = Get_Path(ErrorMessage)
    Set a = New CRAXDRT.Application
    Set r = a.OpenReport(RepPath & reportname)
     'Set R = crxApp.OpenReport("C: est.rpt")
    r.DataBase.Tables.Item(1).SetLogOnInfo "MAZIWA", "MAZIWA", "atm", ""
    r.ReadRecords
    
    'r.ReadRecords
    'r.U
'    MsgBox r.FormulaFields.GetItemByName("[date]").Text
    With frmReports.CRViewer1
        If ReportTitle <> "" Then
            r.ReportTitle = ReportTitle
        End If
        If STRFORMULA <> "" Then
           r.RecordSelectionFormula = STRFORMULA
        End If
        
        .ReportSource = r
        .ViewReport
        .Height = frmReports.Height - 200
        .Width = frmReports.Width
    End With
    
    frmReports.Show vbModal
    STRFORMULA = ""
    Exit Sub

10:
    MsgBox err.description
End Sub

Public Function Save_Weighted_Shares(memberno As String, transdate As Date, OpenShares As Double, _
CloseShares As Double, OpenLoanBal As Double, CloseLoanBal As Double, IntrPaid As Double, IntrCharged _
As Double, NewLoans As Double, IntrOwed As Double, sharecap As Double, errormsg As String) As Boolean
    On Error GoTo SysError
    Dim CnWeight As New Connection
    With CnWeight
        If .State = adStateClosed Then
            .Open SelectedDsn, "atm"
        End If
        .Execute ("Set DateFormat DMY Exec Save_Weigted_Shares '" & memberno & "','" & transdate & _
        "'," & OpenShares & "," & CloseShares & "," & OpenLoanBal & "," & CloseLoanBal & "," & _
        IntrPaid & "," & IntrCharged & "," & NewLoans & "," & IntrOwed & "," & sharecap)
    End With
    Exit Function
SysError:
    errormsg = err.description
    Save_Weighted_Shares = False
End Function

Public Function Save_Dividend(memberno As String, Current_Tot_Shares As Double, Shares_As_At As Double, _
Gross_Dividend As Double, withtax As Double, Net_Dividend As Double, BankName As String, AcctNo As String, _
CompanyName As String, ShareInterest As Double, WithHoldingTax As Double, DivInterest As Double, DivTax As _
Double, DepTax As Double, sharecap As Double, NetDiv As Double, DivTaxAmount As Double, errormsg As String) _
As Boolean
    On Error GoTo SysError
    Dim CnDiv As New Connection
    With CnDiv
        If .State = adStateClosed Then
            .Open SelectedDsn, "atm"
        End If
        .Execute ("Exec SAVE_DIVIDEND '" & memberno & "'," & Current_Tot_Shares & "," & _
        Shares_As_At & "," & Gross_Dividend & "," & withtax & "," & Net_Dividend & ",'" _
        & BankName & "','" & AcctNo & "','" & CompanyName & "'," & ShareInterest & "," _
        & WithHoldingTax & "," & DivInterest & "," & DivTax & "," & DepTax & "," & sharecap _
        & "," & NetDiv & "," & DivTaxAmount)
    End With
    Save_Dividend = True
    Exit Function
SysError:
    errormsg = ""
    Save_Dividend = False
End Function

Public Function getNext(X As Long) As Long
Dim myDate As Date
Dim month12 As Boolean

myDate = CDate(X)

Select Case Month(myDate)

Case 1, 3, 5, 7, 8, 10, 12      '                  Jan,Mar,May,Jul,Aug,Oct,Dec
    X = X + 31

Case 4, 6, 9, 11
     X = X + 30

Case 2
    If Year(myDate) Mod 4 = 0 Then
        'leap year
        X = X + 29
    Else
        X = X + 28
    End If
End Select

myDate = CDate(X)

If Weekday(myDate, vbSunday) = vbSunday Then
    MsgBox "The next Standing Order falls on a Sunday." & Chr(13) & "FOSA shall make it fall on the following Monday"
    X = X + 1
End If

getNext = X
End Function

Public Function Update_Withdrawn_Dividend(sMember As String, errormsg As String) As Boolean
    On Error GoTo SysError
    Dim cn As New Connection
    With cn
        If .State = adStateClosed Then
            .Open SelectedDsn, "atm"
        End If
        .Execute ("Exec Update_Withdrawn_Dividend '" & sMember & "'")
    End With
    Update_Withdrawn_Dividend = True
    Exit Function
SysError:
    errormsg = err.description
    Update_Withdrawn_Dividend = False
End Function

Public Function Save_Dormant(memberno As String, Name As String, LastContrDate As Date, _
auditid As String, errormsg As String, loanbalance As Currency, Optional totalshares As Double, Optional ContrShares _
As Double) As Boolean
    On Error GoTo SysError
    Dim cnDorm As New Connection
    With cnDorm
        If .State = adStateClosed Then
            .CursorLocation = adUseServer
            .Open SelectedDsn, "atm"
        End If
        .Execute ("Set Dateformat DMY Exec Save_Dormant '" & memberno & "','" & Name _
        & "','" & LastContrDate & "','" & auditid & "'," & totalshares & "," & ContrShares & "," & loanbalance)
    End With
    Save_Dormant = True
    Exit Function
SysError:
    errormsg = err.description
    Save_Dormant = False
End Function
Public Function Save_Dormant_loanBalance(memberno As String, Name As String, LastContrDate As Date, _
auditid As String, principal As Currency, errormsg As String, Optional totalshares As Double, Optional ContrShares _
As Double) As Boolean
    On Error GoTo SysError
    Dim cnDorm As New Connection
    With cnDorm
        If .State = adStateClosed Then
            .CursorLocation = adUseServer
            .Open SelectedDsn, "atm"
        End If
        .Execute ("Set Dateformat DMY Exec Save_DORMANT_Loan_repayment '" & memberno & "','" & Name _
        & "','" & LastContrDate & "','" & auditid & "'," & totalshares & "," & ContrShares & "," & principal)
    End With
    Save_Dormant_loanBalance = True
    Exit Function
SysError:
    errormsg = err.description
    Save_Dormant_loanBalance = False
End Function
Public Function Get_Trans_No(TransNo As String, errmsg As String) As String
    On Error GoTo SysError
    TransNo = Replace(TransNo, "/", "")
    TransNo = Replace(TransNo, ":", "")
    TransNo = Replace(TransNo, " ", "")
    TransNo = Replace(TransNo, ";", "")
    TransNo = Replace(TransNo, "-", "")
    TransNo = Replace(TransNo, "\", "")
    Get_Trans_No = TransNo
    Exit Function
SysError:
    errmsg = err.description
    Get_Trans_No = ""
End Function


Public Function zero_balance(ByVal ACCNO As String)
On Error Resume Next
Dim TAST
Dim temprs  As Object
Dim myclass As cdbase
Set myclass = New cdbase
Set temprs = CreateObject("adodb.recordset")
Set cn = CreateObject("adodb.connection")
Provider = myclass.OpenCon
cn.Open Provider, "atm", "atm"
sql = ""
sql = "select * from DailyBalance where accno='" & ACCNO & "'"
Set temprs = CreateObject("adodb.recordset")
temprs.Open sql, cn
If temprs.EOF Then
  TAST = 0
  Else
  TAST = 1
End If
zero_balance = TAST
 
End Function

Public Function Update_Contrib_Amount(memberno As String, ContribID As Long, amount _
As Double, errormsg As String) As Boolean
    Dim cn As New Connection
    On Error GoTo SysError
    With cn
       If .State = adStateClosed Then
           .Open SelectedDsn, "atm"
       End If
       .Execute ("Exec Update_Contrib_Amount '" & memberno & "'," & ContribID & "," _
       & amount)
    End With
    Update_Contrib_Amount = True
    Exit Function
SysError:
    Update_Contrib_Amount = False
    errormsg = err.description
End Function

Public Function Update_Shares(memberno As String, totalshares As Double, _
transdate As Date, errormsg As String, Optional SchemeCode As String) As Boolean
     Dim cn As New Connection
     On Error GoTo SysError
     With cn
        If .State = adStateClosed Then
            .Open SelectedDsn, "atm"
        End If
        .Execute ("Set DateFormat DMY Exec Update_Shares1 '" & memberno & "'," & totalshares & ",'" & _
        transdate & "','" & SchemeCode & "'")
     End With
     Update_Shares = True
     Exit Function
SysError:
     Update_Shares = False
     errormsg = err.description
End Function

Public Function Refresh_Shares(mMemberNo As String, errmsg As String, Optional SchemeCode As String) As Boolean
    On Error GoTo SysError
    Dim rsMember As New Recordset, transdate As Date
    Dim RsShares As New Recordset, TotShares As Double, RefNo As Long
    Dim TShares As Double, rsSchemeShares As New Recordset
    
    Set rsMember = oSaccoMaster.GetRecordset("Select InitShares,ShareCap,ApplicDate from " _
    & "MEMBERS where MemberNo='" & mMemberNo & "'")

    With rsMember
        If .State = adStateOpen Then
            If Not .EOF Then
                'TotShares = IIf(IsNull(.Fields(0)), 0, .Fields(0))
               ' TotShares = TotShares + IIf(IsNull(.Fields(1)), 0, .Fields(1))
              ' TotShares = TotShares + IIf(IsNull(rsSchemeShares!initshares), 0, rsSchemeShares!initshares)
                transdate = !ApplicDate
            End If
        End If
    End With
    
Set rsSchemeShares = oSaccoMaster.GetRecordset("select Initshares from Shares where MemberNo='" & mMemberNo & "'and SharesCode='" & SchemeCode & "'")
    If Not rsSchemeShares.EOF Then
        TShares = IIf(IsNull(rsSchemeShares!initshares), 0, rsSchemeShares!initshares)
    Else
        TShares = 0
    End If
    TotShares = TotShares + TShares
    
    Set RsShares = oSaccoMaster.GetRecordset("Select * From CONTRIB where MemberNo" _
    & "='" & mMemberNo & "' and schemecode='" & SchemeCode & "' order by ContrDate,contribid")
    With RsShares
        If .State = adStateOpen Then
            While Not .EOF
                TotShares = TotShares + !amount
                RefNo = RefNo + 1
                transdate = !ContrDate
                !RefNo = RefNo
                !shareBal = TotShares
                .Update
                .MoveNext
            Wend
        End If
    End With
    Refresh_Shares = True
    Set RsShares = oSaccoMaster.GetRecordset("Select MemberNo From SHARES where MemberNo='" & mMemberNo & "' and sharescode='" & SchemeCode & "'")
    With RsShares
        If Not .EOF Then
            If Not Update_Shares(mMemberNo, TotShares, transdate, errmsg, SchemeCode) Then
                GoTo SysError
            End If
'        Else
'            If Not Save_Shares(mMemberNo, TotShares, TransDate, TransDate, User, 0, _
'            0, ErrMsg, SchemeCode) Then
'                If ErrMsg <> "" Then
'                    Refresh_Shares = False
'                End If
'            End If
        End If
    End With
    Exit Function
SysError:
    errmsg = err.description
    Refresh_Shares = False
End Function

'Public Sub Ref_InitShares(FileName As String, errmsg As String)
'    Dim cn As New Connection, MyFso As New FileSystemObject, MFile As TextStream, _
'    rsMem As New Recordset, InShares As Double, mMemberNo As String, sData As String
'    On Error GoTo SysError
'    Set MFile = MyFso.OpenTextFile(FileName, ForReading, False)
'    Do Until MFile.AtEndOfStream
'        sData = MFile.ReadLine
'        mMemberNo = Left(sData, InStr(1, sData, Chr(9), vbTextCompare) - 1)
'        sData = Right(sData, Len(sData) - InStr(1, sData, Chr(9), vbTextCompare))
'        InShares = sData
'        If InShares <> 0 Then
'            Set rsMem = oSaccoMaster.GetRecordset("Update MEMBERS Set InitShares=" & InShares _
'            & " where MemberNo='" & mMemberNo & "'")
'        End If
'    Loop
'    Exit Sub
'SysError:
'    errmsg = Err.Description
'End Sub

Public Function Refresh_Loan(Loanno As String, errormsg As String) As Boolean
    On Error GoTo SysError
    Dim rsCheq As New Recordset, rsrepay As New Recordset, rsLBal As _
    New Recordset, LBalance As Double, lastdate As Date, PAYNO As Integer, _
    MyPosition As Integer, Interestbalance As Double, myDate As Date, _
    LoanAmount As Double
    Set rsCheq = oSaccoMaster.GetRecordset("Select Amount,DateIssued from Cheques where " _
    & "LoanNo='" & Loanno & "'")
    With rsCheq
        If .State = adStateOpen Then
            If Not .EOF Then
                LBalance = IIf(IsNull(!amount), 0, !amount)
                lastdate = !dateissued
            End If
        End If
    End With
    LoanAmount = LBalance
    Set rsrepay = oSaccoMaster.GetRecordset("Select * From REPAY where LoanNo" _
    & "='" & Loanno & "' order by DateReceived,RepayID")
    With rsrepay
        If .State = adStateOpen Then
            If Not .EOF Then
                While Not .EOF
                    PAYNO = PAYNO + 1
                    lastdate = !datereceived
                    !paymentno = PAYNO
                    If PAYNO > 1 Then
                        If Month(lastdate) = Month(myDate) And Year(lastdate) = Year(myDate) Then
                            !IntrCharged = 0
                        Else
                            !IntrCharged = 0.01 * LBalance
                        End If
                    Else
                        !IntrCharged = 0.01 * LBalance
                    End If
                    If lastdate > "30-06-2007" Then
                        If Left(!Loanno, 1) <> "I" Then
                            !IntrOwed = !IntrCharged - !interest
                            Interestbalance = Interestbalance + !IntrOwed
                        End If
                    Else
                        !IntrOwed = 0
                    End If
                    myDate = !datereceived
                    LBalance = LBalance - !principal
                    !loanbalance = LBalance
                    .Update
                    .MoveNext
                Wend
            End If
        End If
    End With
    Refresh_Loan = True
    Interestbalance = Nearest_Five_Cent(Interestbalance)
    If Not Update_LoanBal_Balance(Loanno, LBalance, Interestbalance, _
    lastdate, ErrorMessage) Then
        If ErrorMessage <> "" Then
            Refresh_Loan = False
        End If
    End If
    If Not Refresh_Guarantors(Loanno, LoanAmount, LBalance, ErrorMessage) Then
        If ErrorMessage <> "" Then
            Refresh_Loan = False
        End If
    End If
    Exit Function
SysError:
    errormsg = err.description
    Refresh_Loan = False
End Function

Public Function Save_Bridging_Loan(Loanno As String, Application_Date As Date, Bridg_LoanNo _
As String, loanbalance As Double, auditid As String, errormsg As String) As Boolean
    On Error GoTo SysError
    Dim cn As New Connection
    With cn
        If .State = adStateClosed Then
            .Open SelectedDsn, "atm"
        End If
        .Execute ("Set DateFormat DMY Exec Save_BridgingLoan '" & Loanno & "','" & _
        Application_Date & "','" & Bridg_LoanNo & "'," & loanbalance & ",'" & User & "'")
    End With
    Save_Bridging_Loan = True
    Exit Function
SysError:
    errormsg = err.description
    Save_Bridging_Loan = False
End Function

Public Function Refresh_Loan_Repay(Loanno As String, transdate As Date, errormsg As String) As Boolean
    On Error GoTo SysError
    Dim rsCheq As New Recordset, rsrepay As New Recordset, rsLBal As _
    New Recordset, LBalance As Double, lastdate As Date, PAYNO As Integer, _
    MyPosition As Integer, Interestbalance As Double, myDate As Date, _
    LoanAmount As Double, IntAmount As Double
    Set rsCheq = oSaccoMaster.GetRecordset("Select Amount,IntAmount from Cheques where " _
    & "LoanNo='" & Loanno & "'")
    With rsCheq
        If .State = adStateOpen Then
            If Not .EOF Then
                LBalance = IIf(IsNull(!amount), 0, !amount)
                IntAmount = IIf(IsNull(!IntAmount), 0, !IntAmount)
            Else
                LBalance = 0
                IntAmount = 0
            End If
        End If
    End With
    LoanAmount = LBalance
    Set rsrepay = oSaccoMaster.GetRecordset("Set DateFormat DMY Select * From REPAY where " _
    & "LoanNo='" & Loanno & "' and DateReceived<='" & transdate & "' order by DateReceived," _
    & "RepayID")
    With rsrepay
        If .State = adStateOpen Then
            While Not .EOF
                PAYNO = PAYNO + 1
                lastdate = !datereceived
                If PAYNO > 1 Then
                    If Month(lastdate) = Month(myDate) And Year(lastdate) = Year(myDate) Then
                        !IntrCharged = 0
                    Else
                        If Left(Loanno, 1) <> "I" Then
                            !IntrCharged = 0.01 * LBalance
                        End If
                    End If
                Else
                    !IntrCharged = 0.01 * LBalance
                End If
                myDate = !datereceived
                LBalance = LBalance - !principal
                !IntrOwed = !IntrCharged - !interest
                !loanbalance = LBalance
                !paymentno = PAYNO
                If lastdate > "30-06-2007" Then
                    If Left(!Loanno, 1) <> "I" Then
                        Interestbalance = Interestbalance + !IntrOwed
                        
                        
                    Else
                        !IntrOwed = 0
                    End If
                End If
                .Update
                .MoveNext
            Wend
        End If
    End With
    If IntAmount > 0 Then
        Interestbalance = IntAmount + Interestbalance
    End If
    
    If DateDiff("M", myDate, transdate) > 0 Then
        If LBalance > 0 Then
            If Left(Loanno, 1) <> "I" Then
                Interestbalance = Interestbalance + (0.01 * LBalance)
            End If
        End If
    End If
    Interestbalance = Nearest_Five_Cent(Interestbalance)
    If Interestbalance < 0 Then Interestbalance = 0
    Int999 = Interestbalance
'    If Interestbalance > 1 Or Interestbalance < -1 Then
'        If Interestbalance < 0 Then
'        '//reduce  the last interest paid
'        Dim prin As Currency
'        Dim intr As Currency
'        Dim Lid As Long
'        ''//add back  the figure to principal
'        Dim RsRecords As New ADODB.Recordset
'
'        mysql = ""
'        mysql = "select top 1* from repay where loanno ='" & LoanNo & "' order by datereceived desc"
'
'        Set RsRecords = oSaccoMaster.GetRecordSet(mysql)
'
'        If Not RsRecords.EOF Then
'            Lid = RsRecords!RepayID
'
'            prin = RsRecords!principal - Interestbalance
'            intr = RsRecords!interest + Interestbalance
'
'            mysql = "set dateformat dmy update repay set principal =" & prin & ",interest =" & intr & " where repayid =" & Lid & ""
'            oSaccoMaster.ExecuteThis (mysql)
'            Interestbalance = 0
'            Refresh_Loan_Repay LoanNo, transdate, ErrorMsg
'        End If
'    End If
'   End If
    
    
    IntAmount = 0
    Refresh_Loan_Repay = True
    Interestbalance = Nearest_Five_Cent(Interestbalance)
    If Not Update_LoanBal_Balance(Loanno, LBalance, Interestbalance, _
    lastdate, ErrorMessage) Then
        If ErrorMessage <> "" Then
            Refresh_Loan_Repay = False
        End If
    End If
    If Not Refresh_Guarantors(Loanno, LoanAmount, LBalance, ErrorMessage) Then
        If ErrorMessage <> "" Then
            Refresh_Loan_Repay = False
        End If
    End If
    Exit Function
SysError:
    errormsg = err.description
    Refresh_Loan_Repay = False
End Function


Public Function Update_Cheques(Loanno As String, chequeno As String, amount As Double, _
IntAmount As Double, CollectorID As String, CollectorName As String, dateissued As Date, _
ClerkStaffNo As String, ClerkName As String, status As String, Reasons As String, auditid _
As String, Remarks As String, errormsg As String) As Boolean
    On Error GoTo SysError
    Dim cn As New Connection
    With cn
        If .State = adStateClosed Then
            .Open SelectedDsn, "atm"
        End If
        .Execute ("Set DateFormat DMY Exec Update_Cheques '" & Loanno & "','" & chequeno _
        & "'," & amount & "," & IntAmount & ",'" & CollectorID & "','" & CollectorName & "','" & _
        dateissued & "','" & ClerkStaffNo & "','" & ClerkName & "','" & status & "','" _
        & Reasons & "','" & auditid & "','" & Remarks & "'")
    End With
    Update_Cheques = True
    Exit Function
SysError:
    errormsg = err.description
    Update_Cheques = False
End Function

Public Function Save_GLTrans(CustomerNo As String, AccName As String, amount As Double, _
availablebalance As Double, ACCNO As String, transDescription As String, transdate As _
Date, Commission As Double, Period As String, transtype As String, Posted As Long, Locked _
As Long, status As Long, vno As String, auditid As String, AuditDate As Date, moduleid As _
String, accd As String, errmsg As String) As Boolean
    On Error GoTo SysError
'    sql = ""
'    sql = "Set DateFormat DMY Insert Into CUSTOMERBALANCE (CustomerNo,AccName,Amount," _
'    & "AvailableBalance,AccNo,TransDescription,TransDate,Commission,Period,TransType," _
'    & "Posted,Locked,Status,Vno,AuditID,AuditDate,ModuleID,AccD) "
'    sql = sql & " Values('" & GlMemNo & "','" & GlName1 & "'," & Amount & _
'    "," & CDbl(bookba) + CDbl(txtPrincipal) & ",'" & GlAccNo & "','" & txtApplicant & _
'    "','" & Format(DTPReceived, "dd/mm/yyyy") & "',0,'" & month(DTPReceived) & _
'    "','CR',0,0,0,'" & txtReceiptNo & "','" & User & "','" & Get_Server_Date & "','3','" _
'    & GlAccNo & "' )"
'    myclass.save (sql)
'    sql = ""
'    sql = "Set DateFormat DMY Update CUB Set Amount=" & CDbl(txtPrincipal) & ",Active=1," _
'    & "TransDescription='" & txtApplicant & "',AvailableBalance=" & CDbl(bookba) + _
'    CDbl(txtPrincipal) & ",TransDate='" & Format(DTPReceived, "dd/mm/yyyy") & "',Vno='" _
'    & txtReceiptNo & "',Period='" & month(DTPReceived) & "',AuditID='" & User & "'," _
'    & "AuditDate='" & Now & "',ModuleID=2 where AccNo='" & GlAccNo & "'"
    Save_GLTrans = True
    Exit Function
SysError:
    errmsg = err.description
    Save_GLTrans = False
End Function

Public Function Save_Bridge_Repayment(Loanno As String, memberno As String, datereceived _
As Date, paymentno As Integer, amount As Double, principal As Double, interest As _
Double, IntrCharged As Double, IntrOwed As Double, loanbalance As Double, ReceiptNo _
As String, Locked As Long, Posted As Long, Accrued As Integer, Remarks As String, _
auditid As String, transby As String, intbalance As Double, NextDueDate As Date, _
Ch As String, PeriodDate As Date, errormsg As String, Optional DocumentNo As String, _
Optional BridgingInterest As Double, Optional cashbookdate As Date, Optional loanaccno As String, Optional interestaccno As String, Optional contra As String, Optional offs As Integer) As Boolean
    On Error GoTo SysError
    Dim cn As New Connection, RsCheque As New Recordset, rsrepay As New Recordset, _
    PAYNO As Long, MyPos As Integer, LBalance As Double
    With cn
        If .State = adStateClosed Then
            .Open SelectedDsn, "atm"
        End If
        .Execute ("Set DateFormat DMY Exec Save_Bridge_Repayment '" & Loanno & "','" & memberno _
        & "','" & datereceived & "'," & paymentno & "," & amount & "," & principal & "," _
        & interest & "," & IntrCharged & "," & IntrOwed & "," & loanbalance & ",'" & _
        ReceiptNo & "'," & Locked & "," & Posted & "," & Accrued & ",'" & Remarks _
        & "','" & auditid & "','" & transby & "'," & intbalance & ",'" & NextDueDate _
        & "','" & Ch & "','" & DocumentNo & "'," & BridgingInterest & ",'" & cashbookdate & "','" & loanaccno & "','" & interestaccno & "','" & contra & "'," & offs & "")
    End With
    If principal <> 0 Then
        If Not Save_Audit("Repay", "Loan Repayment. LoanNo " & Loanno, datereceived, _
        principal, auditid, errormsg) Then
            If errormsg <> "" Then
                Save_Bridge_Repayment = False
                Exit Function
            End If
        End If
    End If
    If interest <> 0 Then
        If Not Save_Audit("Repay", "Interest Payment. LoanNo " & Loanno, datereceived, _
        interest, auditid, errormsg) Then
            If errormsg <> "" Then
                Save_Bridge_Repayment = False
                Exit Function
            End If
        End If
    End If
    Save_Bridge_Repayment = True
    If Not Refresh_Loan(Loanno, ErrorMessage) Then
        If ErrorMessage <> "" Then
            Save_Bridge_Repayment = False
        End If
    End If
    Exit Function
SysError:
    errormsg = err.description
    Save_Bridge_Repayment = False
End Function

Public Function Save_Repayment(Loanno As String, memberno As String, datereceived _
As Date, paymentno As Integer, amount As Double, principal As Double, interest As _
Double, IntrCharged As Double, IntrOwed As Double, loanbalance As Double, ReceiptNo _
As String, Locked As Long, Posted As Long, Accrued As Integer, Remarks As String, _
auditid As String, transby As String, intbalance As Double, NextDueDate As Date, _
Ch As String, PeriodDate As Date, errormsg As String, Optional DocumentNo As String, _
Optional BridgingInterest As Double, Optional LoanAcc As String, _
Optional interestAcc As String, Optional ContraAcc As String, _
Optional cashbookdate As Date) As Boolean
    On Error GoTo SysError
    Dim cn As New Connection, RsCheque As New Recordset, rsrepay As New Recordset, _
    PAYNO As Long, MyPos As Integer, LBalance As Double
    With cn
        If .State = adStateClosed Then
            .Open SelectedDsn, "atm"
        End If
        .Execute ("Set DateFormat DMY Exec Save_Repayment1 '" & Loanno & "','" & memberno _
        & "','" & datereceived & "'," & paymentno & "," & amount & "," & principal & "," _
        & interest & "," & IntrCharged & "," & IntrOwed & "," & loanbalance & ",'" & _
        ReceiptNo & "'," & Locked & "," & Posted & "," & Accrued & ",'" & Remarks _
        & "','" & auditid & "','" & transby & "'," & intbalance & ",'" & NextDueDate _
        & "','" & Ch & "','" & DocumentNo & "','" & TransNo & "','" & LoanAcc _
        & "','" & interestAcc & "','" & ContraAcc & "','" & Format(cashbookdate, "dd/mm/yyyy") & "'")
    End With
    If principal <> 0 Then
        If Not Save_Audit("Repay", "Loan Repayment. LoanNo " & Loanno, datereceived, _
        principal, auditid, errormsg) Then
            If errormsg <> "" Then
                Save_Repayment = False
                Exit Function
            End If
        End If
    End If
    If interest <> 0 Then
        If Not Save_Audit("Repay", "Interest Payment. LoanNo " & Loanno, datereceived, _
        interest, auditid, errormsg) Then
            If errormsg <> "" Then
                Save_Repayment = False
                Exit Function
            End If
        End If
    End If
    Save_Repayment = True
    If Not Refresh_Loan(Loanno, ErrorMessage) Then
        If ErrorMessage <> "" Then
            Save_Repayment = False
        End If
    End If
    Exit Function
SysError:
    errormsg = err.description
    Save_Repayment = False
End Function

Public Function Update_LoanBalance(Loanno As String, Loancode As String, memberno As String, _
balance As Double, FirstDate As Date, repayrate As Double, lastdate As Date, Interest_Rate As _
Double, repaymethod As String, Cleared As String, AutoCalc As String, IntrAmount As Double, _
repayperiod As Long, Remarks As String, auditid As String, loanbalance As Double, ACCNO As _
String, intbalance As Double, tamount As Double, NextDueDate As Date, Mintbal As Double, _
duedate As Date, errormsg As String) As Boolean
    On Error GoTo SysError
    Dim cn As New Connection
    With cn
        If .State = adStateClosed Then
            .Open SelectedDsn, "atm"
        End If
        .Execute ("Set DateFormat DMY Exec Update_LoanBal '" & Loanno & "','" & Loancode & _
        "','" & memberno & "'," & balance & ",'" & FirstDate & "'," & repayrate & ",'" _
        & lastdate & "'," & Interest_Rate & ",'" & repaymethod & "','" & Cleared & "','" _
        & AutoCalc & "'," & IntrAmount & "," & repayperiod & ",'" & Remarks & "','" _
        & auditid & "'," & loanbalance & ",'" & ACCNO & "'," & intbalance & "," & _
        tamount & ",'" & NextDueDate & "'," & Mintbal & ",'" & duedate & "'")
    End With
    Update_LoanBalance = True
    Exit Function
SysError:
    ErrorMessage = err.description
    Update_LoanBalance = True
End Function

Public Function Get_Monthly_Interest(memberno As String, sMonth As Long, sYear As Long, _
errormsg As String) As Double
    On Error GoTo SysError
    Dim CnShares As New Connection, RsShares As New Recordset
    With CnShares
        If .State = adStateClosed Then
            .Open SelectedDsn, "atm"
            Set RsShares = .Execute("Select Sum(Interest) as Shares From REPAY where " _
            & "MemberNo='" & memberno & "' and Month(DateReceived)=" & sMonth & " and " _
            & "Year(DateReceived)=" & sYear)
        End If
        With RsShares
            If .State = adStateOpen Then
                If Not .EOF Then
                    Get_Monthly_Interest = IIf(IsNull(!Shares), 0, !Shares)
                Else
                    Get_Monthly_Interest = 0
                End If
            End If
        End With
    End With
    Exit Function
SysError:
    errormsg = err.description
    Get_Monthly_Interest = 0
End Function

Public Function Get_Monthly_Principal(memberno As String, sMonth As Long, sYear As Long, _
errormsg As String) As Double
    On Error GoTo SysError
    Dim CnShares As New Connection, RsShares As New Recordset
    With CnShares
        If .State = adStateClosed Then
            .Open SelectedDsn, "atm"
            Set RsShares = .Execute("Select Sum(Principal) as Shares From REPAY where " _
            & "MemberNo='" & memberno & "' and Month(DateReceived)=" & sMonth & " and " _
            & "Year(DateReceived)=" & sYear)
        End If
        With RsShares
            If .State = adStateOpen Then
                If Not .EOF Then
                    Get_Monthly_Principal = IIf(IsNull(!Shares), 0, !Shares)
                Else
                    Get_Monthly_Principal = 0
                End If
            End If
        End With
    End With
    Exit Function
SysError:
    errormsg = err.description
    Get_Monthly_Principal = 0
End Function

Public Function Get_Monthly_Shares(memberno As String, sMonth As Long, sYear As Long, _
errormsg As String) As Double
    On Error GoTo SysError
    Dim CnShares As New Connection, RsShares As New Recordset
    With CnShares
        If .State = adStateClosed Then
            .Open SelectedDsn, "atm"
            Set RsShares = .Execute("Select Sum(Amount) as Shares From CONTRIB where " _
            & "MemberNo='" & memberno & "' and Month(ContrDate)=" & sMonth & " and " _
            & "Year(ContrDate)=" & sYear)
        End If
        With RsShares
            If .State = adStateOpen Then
                If Not .EOF Then
                    Get_Monthly_Shares = IIf(IsNull(!Shares), 0, !Shares)
                Else
                    Get_Monthly_Shares = 0
                End If
            End If
        End With
    End With
    Exit Function
SysError:
    errormsg = err.description
    Get_Monthly_Shares = 0
End Function

Public Function Save_LoanBalance(Loanno As String, Loancode As String, memberno As String, _
balance As Double, FirstDate As Date, repayrate As Double, lastdate As Date, interest As Double, _
repaymethod As String, Cleared As String, AutoCalc As String, IntrAmount As Double, repayperiod _
As Long, Remarks As String, auditid As String, loanbal As Double, ACCNO As String, intbalance As _
Double, tamount As Double, NextDueDate As Date, Mintbal As Double, duedate As Date, errormsg As _
String) As Boolean
    On Error GoTo SysError
    Dim cn As New Connection
    With cn
        If .State = adStateClosed Then
            .Open SelectedDsn, "atm"
        End If
        .Execute ("Set DateFormat DMY Exec Save_LoanBalance '" & Loanno & "','" & Loancode _
        & "','" & memberno & "'," & balance & ",'" & FirstDate & "'," & repayrate & ",'" & _
        lastdate & "'," & interest & ",'" & repaymethod & "','" & Cleared & "','" & AutoCalc _
        & "'," & IntrAmount & "," & repayperiod & ",'" & Remarks & "','" & User & "'," & _
        loanbal & ",'" & ACCNO & "'," & intbalance & "," & tamount & ",'" & NextDueDate & "'," _
        & Mintbal & ",'" & duedate & "'")
    End With
    Save_LoanBalance = True
    Exit Function
SysError:
    errormsg = err.description
    Save_LoanBalance = False
End Function

Public Function Update_LoanBal_Balance(Loanno As String, balance As Double, _
InterestBal As Double, lastdate As Date, errormsg As String) As Boolean
    On Error GoTo SysError
    Dim cn As New Connection
    With cn
        If .State = adStateClosed Then
            .Open SelectedDsn, "atm"
        End If
        .Execute ("set dateformat DMY Exec Update_LoanBal_Balance '" & Loanno _
        & "'," & balance & "," & InterestBal & ",'" & lastdate & "'")
    End With
    Update_LoanBal_Balance = True
    Exit Function
SysError:
    errormsg = ""
    errormsg = err.description
    Update_LoanBal_Balance = False
End Function

Public Function Update_Repay_Balance(Loanno As String, RepayID As Integer, _
loanbalance As Double, errormsg As String) As Boolean
    On Error GoTo SysError
    Dim cn As New Connection
    With cn
        If .State = adStateClosed Then
            .Open SelectedDsn, "atm"
        End If
        .Execute ("Exec Update_Repay_Balance '" & Loanno & "'," & RepayID & "," _
        & loanbalance)
    End With
    Update_Repay_Balance = True
    Exit Function
SysError:
    Update_Repay_Balance = False
    errormsg = err.description
End Function

Public Function Nearest_ShillingUp(amount As Double) As Double
    On Error GoTo SysError
    Dim MyAmount As Double, MyNum As String, MyAmount1 As Double
    Dim MyDif As Double
    MyAmount = Format(amount, "###,###,###,###,##0.0")
    MyAmount1 = Format(amount, "###,###,###,##0")
    MyDif = Format((MyAmount1 - MyAmount), Cfmt)
    If MyDif < 0 Then
        Nearest_ShillingUp = MyAmount1 + 1
    Else
        Nearest_ShillingUp = MyAmount1
    End If
    Exit Function
SysError:
    Nearest_ShillingUp = 0
End Function

Public Function Nearest_Shilling(amount As Double) As Double
    On Error GoTo SysError
    Dim MyAmount As Double, MyNum As String, MyAmount1 As Double
    Dim MyDif As Double
    MyAmount = Format(amount, "###,###,###,###,##0.0")
    MyAmount1 = Format(amount, "###,###,###,##0")
    MyDif = Format((MyAmount1 - MyAmount), Cfmt)
    If MyDif <= 0 Then
        Nearest_Shilling = MyAmount1
    Else
        Nearest_Shilling = MyAmount1 - 1
    End If
    Exit Function
SysError:
    Nearest_Shilling = 0
End Function

Public Function Nearest_Five_Cent(amount As Double) As Double
    On Error GoTo SysError
    Dim MyAmount As Double, MyNum As String, MyAmount1 As Double
    Dim MyDif As Double
    MyAmount = Format(amount, "###,###,###,###,##0.00")
    MyAmount1 = Format(amount, "###,###,###,##0.0")
    MyDif = Format((MyAmount1 - MyAmount), Cfmt)
    If MyDif < 0 Then
        Nearest_Five_Cent = MyAmount1 + 0.05
    Else
        Nearest_Five_Cent = MyAmount1
    End If
    Exit Function
SysError:
    Nearest_Five_Cent = 0
End Function

Function Get_Server_Date() As Date
    On Error GoTo 10
    Dim rs As Recordset
    Set cn = New Connection
    Set rs = CreateObject("adodb.recordset")
    Set cn = CreateObject("adodb.connection")
    Provider = SelectedDsn
    cn.Open Provider, "atm", "atm"
'    Set cn = New ADODB.Connection
'    Set myclass = New cdbase
'    Provider = myclass.OpenCon
'    cn.Open Provider, "atm","atm"
'    Set rs = New Recordset
    rs.Open "set dateformat dmy select GetDate()", cn, adOpenStatic, adLockOptimistic
    Get_Server_Date = rs(0)
    rs.Close
    Exit Function
10:    MsgBox err.description
End Function

Public Function get_serverdate(serverDate) As Date
Dim rdate As Date
Set myclass = New cdbase
Set temprs = CreateObject("adodb.recordset")
Set cn = CreateObject("adodb.connection")
Provider = myclass.OpenCon
cn.Open Provider, "atm", "atm"
sql = "select getdate() "
temprs.Open sql, cn
 serverDate = Format(temprs.Fields(0), "MM/DD/YY")
End Function
Public Function status() As Boolean
    Dim temprs As Object
    Set myclass = New cdbase
    Set temprs = CreateObject("adodb.recordset")
    Set cn = CreateObject("adodb.connection")
    Provider = myclass.OpenCon
   cn.Open Provider, "atm", "atm"
    sql = "select inactive ,frozen,closed from customeraccount"
     temprs.Open sql, cn
If temprs!Closed = 1 Then
MsgBox "You Cannot transact The Account Is Closed ", vbInformation, "Transaction"
ElseIf temprs!frozen = 1 Then
MsgBox "You Cannot transact The Account Is Frozen ", vbInformation, "Transction"
ElseIf temprs!inactive = 1 Then
MsgBox "You Cannot transact The Account Is Inactive ", vbInformation, "Transaction"
End If

End Function
Private Function getCurBal(strAccNo As String) As Currency
Dim temprs As Object
    getCurBal = 0
    
    Set myclass = New cdbase
    Set temprs = CreateObject("adodb.recordset")
    Set cn = CreateObject("adodb.connection")
    Provider = myclass.OpenCon
   cn.Open Provider, "atm", "atm"
     
    With temprs
        .Open "SELECT     TOP 1 AvailableBalance as x From DailyBalance WHERE     (AccNO = '" & strAccNo & "') ORDER BY CustomerBalanceid DESC", cn
        If Not IsNull(!X) And Not .BOF Then getCurBal = CCur(!X)
        .Close
    End With
    Set temprs = Nothing
End Function

Public Function getinfo(anyNo As String, Optional CustomerName As String, Optional retNo As String, Optional AccName As String, Optional custBal As Currency, Optional pic As String, Optional sign As String) As Boolean
    Dim nav As Long, str As String
    str = Trim(anyNo)
    Select Case myLevel
    Case 1
        For nav = 0 To maxRec
            If str = accData(nav).ACCNO Then
                getinfo = True
                With accData(nav)
                    CustomerName = .custName
                    retNo = .custno
                    AccName = .AccName
                    custBal = .custBal
                    pic = .pic
                    sign = .sign
                End With
                Exit For
            End If
        Next
        
    Case 2
        For nav = 0 To maxRec
            If str = accData(nav).custno Then
                With accData(nav)
                    getinfo = True
                    If custBal = 0 Then
                        CustomerName = .custName
                        retNo = .ACCNO
                        AccName = .AccName
                    End If
                    
                    custBal = custBal + .custBal
                    
                End With
                
            End If
        Next
    End Select
End Function

'Public Function Get_Opening_Balances(errormsg As String)
'    On Error GoTo SysError
'    Dim MyFso As New FileSystemObject, TFile As TextStream, nMemberNo As String
'    Dim MFSO As New FileSystemObject
'    Dim rsm As New Recordset, mTransDate As Date, mMemberNo As String
'    Dim RsShares As New Recordset, mShares As Double, TFile1 As TextStream
'    mTransDate = "30/06/2007"
'    Set TFile = MyFso.OpenTextFile("c:\Shares.txt", ForReading, False)
'    Do Until TFile.AtEndOfStream
'        nMemberNo = TFile.ReadLine
'        nMemberNo = Left(nMemberNo, InStr(1, nMemberNo, Chr(9), vbTextCompare) - 1)
'        Set rsm = oSaccoMaster.GetRecordset("Select * From Members where MemberNo" _
'        & "='" & nMemberNo & "'")
'        With rsm
'            If .State = adStateOpen Then
'                If Not .EOF Then
'                    mShares = IIf(IsNull(!sharecap), 0, !sharecap)
'                    mShares = mShares + IIf(IsNull(!initshares), 0, !initshares)
'                End If
'            End If
'        End With
'        Set RsShares = oSaccoMaster.GetRecordset("Set DateFormat DMY Select Sum(Amount) " _
'        & "as Shares From CONTRIB where MemberNo='" & nMemberNo & "' and ContrDate<='" _
'        & mTransDate & "'")
'        With RsShares
'            If .State = adStateOpen Then
'                If Not .EOF Then
'                    mShares = mShares + IIf(IsNull(!Shares), 0, !Shares)
'                End If
'            End If
'        End With
'        Set rsm = oSaccoMaster.GetRecordset("Update RECON Set Easy=" & mShares & " where " _
'        & "MemberNo='" & nMemberNo & "'")
'    Loop
'    MsgBox "Complete"
'    Exit Function
'SysError:
'    errormsg = Err.Description
'    MsgBox Err.Description
'End Function

'Public Function getGlCurrentBalance(ACCNO As String) As Double
'    On Error GoTo Capture
'    Set rst = oSaccoMaster.GetRecordset("Select CBal from UDF_GL_CurrentBalance('" & ACCNO & "')")
'    If Not rst.EOF Then
'        getGlCurrentBalance = rst(0)
'    Else
'        getGlCurrentBalance = 0
'    End If
'
'    Exit Function
'Capture:
'    MsgBox IIf(ErrorMessage = "", err.description, ErrorMessage)
'End Function
'Public Function GetRecords(str As String)
'    On Error GoTo Capture
'    Set rst = New ADODB.Recordset
'    sql = mysql
'    Set cn = CreateObject("adodb.connection")
'    Provider = "GakCofDb"
'   cn.Open Provider, "atm","atm"
'    rst.Open str, cn
'    success = True
'    Exit Function
'Capture:
'    success = False
'    MsgBox err.description
'End Function
Public Function GLAcc_Exists(ACCNO As String, errormsg As String) As Boolean
    On Error GoTo SysError
    Dim rsgl As New Recordset
    Set rsgl = oSaccoMaster.GetRecordset("Select * From GLSETUP where AccNo='" & ACCNO & "'")
    With rsgl
        If .State = adStateOpen Then
            If Not .EOF Then
                GLAcc_Exists = True
            Else
                GLAcc_Exists = False
            End If
        End If
    End With
    Exit Function
SysError:
    errormsg = err.description
    GLAcc_Exists = True
End Function

Public Sub initAccData(Optional levelOfInit As Integer = 1)

    Dim myIndex As Long
    Dim temprs As Object
    Dim myrec As Object, X As Integer
    maxRec = 0
    ReDim accData(0)
    Set myclass = New cdbase
    Set cn = CreateObject("adodb.connection")
    
    If Provider = "" Then Provider = myclass.OpenCon
    
   cn.Open Provider, "atm", "atm"
    Set myrec = CreateObject("adodb.recordset")
    
    If levelOfInit = 1 Then ' priority to accountnumber
            sql = "SELECT DISTINCT "
            sql = sql & " Customers.Surname + ',  ' + Customers.OtherNames AS name, CustomerAccount.AccountNumber, CustomerAccount.CustomerNo, CustomerAccount.AccountName, CustomerAccount.Picture,"
            sql = sql & " CustomerAccount.Signature"
            sql = sql & " FROM         CustomerAccount LEFT OUTER JOIN"
            sql = sql & " Customers ON CustomerAccount.CustomerNo = Customers.CustomerNo"
            
            myrec.Open sql, cn
                     
            ReDim accData(0)
            With myrec
                While Not .EOF
                    accData(myIndex).ACCNO = !accountnumber & ""
                    accData(myIndex).AccName = !AccountName & ""
                    accData(myIndex).custno = !CustomerNo & ""
                    accData(myIndex).custName = !Name & ""
                    accData(myIndex).pic = !Picture & ""
                    accData(myIndex).sign = !Signature & ""
                    accData(myIndex).custBal = getCurBal(accData(myIndex).ACCNO)
                    
                    ReDim Preserve accData(myIndex + 1)
                    myIndex = myIndex + 1
                    .MoveNext
                Wend
            End With


'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
    ElseIf levelOfInit = 2 Then 'priority to custNo
            
            sql = "select distinct Customerno from Customers"
            
            myrec.Open sql, cn
                     
            ReDim accData(0)
            With myrec
                While Not .EOF
                    accData(myIndex).custno = !CustomerNo & ""
                    ReDim Preserve accData(myIndex + 1)
                    myIndex = myIndex + 1
                    .MoveNext
                Wend
             End With
        
            Set myrec = CreateObject("adodb.recordset")
            sql = "SELECT "
            sql = sql & " CustomerAccount.AccountNumber,"
            sql = sql & " Customers.CustomerNo,"
            sql = sql & " DailyBalance.AvailableBalance,"
            sql = sql & " DailyBalance.AccName,"
            sql = sql & " Customers.Surname,"
            sql = sql & " Customers.OtherNames,"
            sql = sql & " CustomerAccount.Picture,"
            sql = sql & " CustomerAccount.Signature"
            sql = sql & " FROM Customers LEFT OUTER JOIN CustomerAccount"
            sql = sql & " LEFT OUTER JOIN DailyBalance"
            sql = sql & " ON cast(CustomerAccount.CustomerNo as varchar(50))=CustomerBalance.CustomerNo"
            sql = sql & " ON cast(Customers.CustomerNo as varchar(50))=CustomerAccount.CustomerNo"
            sql = sql & " ORDER BY DailyBalance.CustomerBalanceid DESC"
            With myrec
                .Open sql, cn
                For X = 0 To myIndex - 1
                    .Filter = "CustomerNo = '" & accData(X).custno & "'"
                    If Not .EOF Then
                        accData(X).AccName = !AccName & ""
                        If Not IsNull(!availablebalance) Then accData(X).custBal = !availablebalance
                        accData(X).ACCNO = !accountnumber & ""
                        accData(X).custName = !surname & ", " & !OtherNames
                        accData(X).pic = !Picture & ""
                        accData(X).sign = !Signature & ""
                    End If
                Next
            End With
    End If
    
    myrec.Close
    Set myrec = Nothing
    
    maxRec = myIndex - 1
    myLevel = levelOfInit
End Sub


Public Function Get_Cheque_amount(Loanno As String) As Currency
    Dim RsCheque As New ADODB.Recordset
    
    mysql = ""
    mysql = "select * from cheques  where loanno ='" & Loanno & "'"
    
    Set RsCheque = oSaccoMaster.GetRecordset(mysql)
    
    If Not RsCheque.EOF Then
        If Not IsNull(RsCheque!amount) Then
        Get_Cheque_amount = RsCheque!amount
        Else
        Get_Cheque_amount = 0
        End If
    Else
    Get_Cheque_amount = 0
    End If
End Function
'Function to check Status of the period. Returns True if closed else False
Public Function Check_Period_If_Closed(transdate As Date) As Boolean
Dim rsPeriod As New Recordset
Dim currentdate As Date
Dim status As Integer
'CurrentDate = Get_Server_Date

'Get the period in which u r posting Transaction To
Set rsPeriod = oSaccoMaster.GetRecordset("set dateformat dmy select * from PERIODS where Period=" & Month(Format(transdate, "dd/mm/yyyy")) & " and PeriodYear=" & Year(Format(transdate, "dd/mm/yyyy")) & "")
If Not rsPeriod.EOF Then  'Period is already set
  
  status = IIf(rsPeriod!status = True, 1, 0)
 
    If status = 1 Then 'Period is closed
    
        Check_Period_If_Closed = True
        MsgBox "The Period is Closed.Please specify the Correct Date.", vbCritical, "Invalid Period"
        Exit Function
    Else 'Period is Open
        Check_Period_If_Closed = False
    End If
  
Else 'Period is not Created
    Check_Period_If_Closed = True
    MsgBox "The Period is not defined.Create it before proceeding.", vbInformation, "Period setup"
    Exit Function
End If

End Function

Public Function PopulateCombo(myCombo As ComboBox, Filter As Boolean, Optional SchemeCode As String)
Dim rsCombo As New Recordset
Set rsCombo = Nothing
If Filter = False Then
    Set rsCombo = oSaccoMaster.GetRecordset("select * from ShareType order by ismainshares desc")
    With rsCombo
        If Not .EOF Then
        myCombo.Clear
            While Not .EOF
                myCombo.AddItem !sharesCode
                .MoveNext
            Wend
        End If
        myCombo.ListIndex = 0
    End With
Else
strValue = ""
    Set rsCombo = oSaccoMaster.GetRecordset("select * from ShareType where SharesCode='" & SchemeCode & "'")
    With rsCombo
        If Not .EOF Then
        strValue = IIf(IsNull(!sharestype), "", !sharestype)
        I = !ismainshares
        Offset = !usedtooffset
        End If
    End With
End If
End Function

Public Function Get_Total_Shares(memberno As String, Optional GetWithdrawable As Boolean) As Scheme_Details
On Error GoTo Syserr
Dim RsShares As New Recordset
If GetWithdrawable = False Then
    Set RsShares = oSaccoMaster.GetRecordset("Select sum(TotalShares) as Totalshares from SHARES where MemberNo='" & memberno & "'")
Else
    Set RsShares = oSaccoMaster.GetRecordset("Select sum(TotalShares) as Totalshares from SHARES S inner join " _
    & " SHARETYPE ST on S.SharesCode=ST.SharesCode where S.MemberNo='" & memberno & "' and ST.Withdrawable=1")
End If

With RsShares
    If Not .EOF Then
        Get_Total_Shares.totalshares = Format(IIf(IsNull(!totalshares), 0, !totalshares), Cfmt)
    Else
        Get_Total_Shares.totalshares = 0
    End If
End With
Exit Function
Syserr:
    MsgBox err.description
End Function

Public Function GetUsers(ByVal strPassword As String, ByVal strUserName As String) As Boolean

End Function
