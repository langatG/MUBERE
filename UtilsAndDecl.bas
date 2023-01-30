Attribute VB_Name = "UtilsAndDecl"

Option Explicit
Public LvwSelectedItm As String
Public strRangeFrom As String
Public strFrom As String
Public strTo As String
Public strRangeTo As String
Public isItFrom As Boolean
Public Itm As ListItem
Public strMemberNo As String
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Public Const CB_SETDROPPEDWIDTH = &H160
Public Const MF_BYPOSITION = &H400&
'Public oCryptoX As CryptoX
Public Function Save_PSL1(memberno As String, LoanNo As String, transdate As Date, Refno As Integer, _
Description As String, amount As Double, principal As Double, interest As Double, Shares As Double, _
CompanyName As String, ErrorMsg As String, memDescription As String, LastTransactionDate As Date, AuditDate As Date) As Boolean
    Dim Cnn As New Connection
    'On Error GoTo SysError
    With Cnn
        If .State = adStateClosed Then
            .CursorLocation = adUseServer
            .Open SelectedDsn
        End If
        .Execute ("Set DateFormat DMY Exec Save_PSL1 '" & memberno & "','" & LoanNo _
        & "','" & transdate & "'," & Refno & ",'" & Description & "'," & amount & _
        "," & principal & "," & interest & "," & Shares & ",'" & Replace(CompanyName, "'", "") & "','" & memDescription & "','" & LastTransactionDate & "','" & AuditDate & "'")
    End With
    Save_PSL1 = True
    Set Cnn = Nothing
    Exit Function
SysError:
    Save_PSL1 = False
    ErrorMsg = Err.Description
End Function



Public Function Save_PSL(memberno As String, LoanNo As String, transdate As Date, Refno As Integer, _
Description As String, amount As Double, principal As Double, interest As Double, Shares As Double, _
CompanyName As String, ErrorMsg As String, memDescription As String, LastTransactionDate As Date, _
Optional ByLaw As Double, Optional regfee As Double) As Boolean
    Dim Cnn As New Connection
   On Error GoTo SysError
    With Cnn
        If .State = adStateClosed Then
            .CursorLocation = adUseServer
            .Open SelectedDsn
        End If
        .Execute ("Set DateFormat DMY Exec Save_PSL '" & memberno & "','" & LoanNo _
        & "','" & transdate & "'," & Refno & ",'" & Description & "'," & amount & _
        "," & principal & "," & interest & "," & Shares & ",'" & Replace(CompanyName, "'", "") & _
         "','" & memDescription & "','" & LastTransactionDate & "'," & regfee & "," & ByLaw & "")
    End With
    Save_PSL = True
    Set Cnn = Nothing
    Exit Function
SysError:
    Save_PSL = False
    ErrorMsg = Err.Description
End Function
    
Public Sub SetComboWidth(oCombo As ComboBox, lWidth As Long)
    SendMessage oCombo.hwnd, CB_SETDROPPEDWIDTH, lWidth, 0
End Sub

Public Sub Connect()
    Set Cnn = New Connection
    Set rst = New Recordset
    Set rst1 = New Recordset
    Set rst2 = New Recordset
    'DataB = frmLogin.cboDataBase.Text
    Cnn.Open DataB
End Sub

Public Sub LoadMember()
  Connect
  With rst
    '.Open "select memberno,surname,othernames from members where memberno ='" & frmApplic.txtmemberno.Text & "'", Cnn, adOpenKeyset, adLockOptimistic
    If .RecordCount > 0 Then
   ' frmApplic.txtName.Text = !othernames & " " & !surname
    With rst1
      .Open "select totalshares from shares where memberno = '" & rst!memberno & "'", Cnn, adOpenKeyset, adLockOptimistic
      If .RecordCount > 0 Then
        Shares = !totalshares
        'frmApplic.txtShares.Text = Format(Shares, "###,###,###,##0.00")
      End If
      .Close
    End With
    End If
    .Close
  End With
End Sub

Public Sub LoadLoan()
If Cnn.State = adStateClosed Then
  Connect
End If
  With rst
    '.Open "select * from loanbal where memberno='" & frmApplic.txtmemberno.Text & "'", Cnn, adOpenKeyset, adLockOptimistic
    If .RecordCount > 0 Then
      While Not .EOF
        If !balance > 0 Then
          i = i + 1
          loan = loan + !balance
          With rst1
            .Open "select sum(principal) as Prince from repay where loanno='" & rst!LoanNo & "'", Cnn, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
              If Not IsNull(!prince) Then prince = !prince
              If rst!repaymethod = "RBAL" Then
                MaxAmount = Format(MaxAmount + ((prince + rst!balance) / (rst!repayperiod)), "###,###,###,##0.00")
              ElseIf rst!repaymethod = "AMRT" Then
                MaxAmount = Format(MaxAmount + ((prince + rst!balance) / (rst!repayperiod)), "###,###,###,##0.00")
              ElseIf rst!repaymethod = "STL" Then
                MaxAmount = Format(MaxAmount + ((prince + rst!balance) / (rst!repayperiod)), "###,###,###,##0.00")
              End If
            End If
            .Close
          End With
          prince = 0
        End If
        .MoveNext
      Wend
      frmApplic.txtOutBalance.Text = Format(loan, "###,###,###,##0.00")
      frmApplic.txtOutLoan = i
      i = 0
      .MoveFirst
      With rst1
        .Open "select memberno,newcontr from shrvar", Cnn, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
          .MoveFirst
          .Find "memberno ='" & rst!memberno & "'", , adSearchForward, adBookmarkFirst
          If .EOF Then
            .Close
            .Open "select memberno,initshares from members where memberno = '" & rst!memberno & "'", Cnn, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
              initshares = !initshares
            End If
          Else
            initshares = !NewContr
          End If
        End If
        .Close
      End With
      frmApplic.txtTotDeduct = Format((MaxAmount + initshares), "###,###,###,##0.00")
      With rst1
        .Open "select loantoshareratio from sysparam", Cnn, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
          LtSRatio = !loantoshareratio
          frmApplic.txtMaxAmount.Text = Format(((LtSRatio * Shares) - loan), "###,###,###,##0.00")
        End If
        .Close
      End With
    End If
    .Close
  End With
End Sub
Public Sub LoadCheque()
   Connect
   With rst
      .Open "select * from cheques where loanno = '" & frmCheques.txtLoanno.Text & "'", Cnn, adOpenKeyset, adLockOptimistic
      If .RecordCount > 0 Then
         With rst1
            .Open "select * from loans where loanno = '" & rst!LoanNo & "'", Cnn, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
               With rst2
                  .Open "select surname,othernames from members where memberno = '" & rst1!memberno & "'", Cnn, adOpenKeyset, adLockOptimistic
                  If .RecordCount > 0 Then
                    ' frmCheques.txtApplicant.Text = Rst1!othernames & " "
                  End If
                  .Close
               End With
            End If
            .Close
         End With
      End If
      .Close
   End With
End Sub

Public Sub ClearControlsIn(frm As Form)
'Call this to clear all controls on the form
'Pass the form you want to clear
On Error GoTo ErrorTrap
Dim ctrl As Control
    For Each ctrl In frm.Controls
        If TypeOf ctrl Is TextBox Then
            ctrl.Text = ""
        ElseIf TypeOf ctrl Is OptionButton Then
            ctrl.value = 0
        ElseIf TypeOf ctrl Is ComboBox Then
            ctrl.ListIndex = -1
        ElseIf TypeOf ctrl Is ListView Then
            ctrl.ListItems.Clear
        End If
    Next ctrl
    Exit Sub
ErrorTrap:
    MsgBox Err.Description, , "Utilities"
    Exit Sub
End Sub

Public Function checkPriviledges(frm As Form, valToEncr As String)
    Set Rst4 = oSaccoMaster.GetRecordSet("select groupid from users where userid= '" & User & "'")
    If Not Rst4.EOF Then
    Set Rst5 = oSaccoMaster.GetRecordSet("select * from usergrps where groupid= '" & Rst4!groupid & "'")
    If Not Rst5.EOF Then
        valToEncrOrDecr = valToEncr
        EncryptPassword
        If EncryptPass = "View" Then
            acceptmodify = True
            frm.cmdAdd.Enabled = False
            frm.cmdEdit.Enabled = False
            frm.cmdCancel.Enabled = False
            frm.cmdsave.Enabled = False
        Else
            acceptmodify = False
        End If
    End If
    End If
End Function
Public Function EncryptPassword()
    Dim Pwd As Variant
    Dim Temp As String, PwdChr As Long
    Dim EncryptKey As Long
    Pwd = valToEncrOrDecr
    EncryptKey = Int(Sqr(Len(Pwd) * 81)) + 23
    
    For PwdChr = 1 To Len(Pwd)
        Temp = Temp + Chr(Asc(Mid(Pwd, PwdChr, 1)) Xor EncryptKey)
    Next PwdChr
    
    EncryptPass = Temp

End Function
Public Function Export_To_Excel(rsExport As Recordset)

Dim iColumn As Integer
Dim xlApp As Object
Dim xlWb As Object
Dim xlWs As Object
Set xlApp = CreateObject("Excel.Application")
Set xlWb = xlApp.Workbooks.Add
Set xlWs = xlWb.Worksheets("Sheet1")
xlApp.Visible = True
For iColumn = 1 To rsExport.Fields.Count - 3
    xlWs.Cells(1, iColumn) = rsExport.Fields(iColumn).name
Next iColumn
xlApp.UserControl = True
'xlws.cells(2,1)
End Function

Public Function EndMonthDate(ByVal dteMonth As Integer, ByVal intYear As Integer) As Date
'Pass the number signifying the month of the year and the
'function returns the last date of the month
Select Case dteMonth
    Case 1, 3, 5, 7, 8, 10, 12
        EndMonthDate = 31 & "/" & dteMonth & "/" & intYear
    Case 2
        EndMonthDate = 28 & "/" & dteMonth & "/" & intYear
    Case 4, 6, 9, 11
          EndMonthDate = 30 & "/" & dteMonth & "/" & intYear
End Select
End Function
Public Sub PositionForm(frm As Form)
    If Not frm.WindowState = vbNormal Then
        frm.WindowState = vbNormal
    End If
    
End Sub

Public Function Dividend_By_ShareInt(ByVal dblPCTVal As Double, _
ByVal cMemberShares As Currency) As Currency

   Dividend_By_ShareInt = (dblPCTVal / 100) * cMemberShares * (CLng(month(Date)) / 12)

End Function


Public Function Dividends_By_Profit(ByVal cTotalShares As Currency, ByVal cMemberShares As Currency, _
                                    cProfit As Currency, Optional TaxRate As Double) As Currency

Dividends_By_Profit = (cMemberShares / cTotalShares) * cProfit

End Function

Public Function WitholdingTax(ByVal cMemberDividends As Currency, ByVal dblTaxRate As Double) As Currency
   
   WitholdingTax = (dblTaxRate / 100) * cMemberDividends
   
End Function

Public Function NetDividends(cGrossDividends As Currency, cWitholdingTax As Currency) As Currency
   
   NetDividends = cGrossDividends - cWitholdingTax

End Function

Public Sub Main()
Dim strConnection As String
Dim rsRecordset As New Recordset
On Error GoTo ErrHandler
With oSaccoMaster
    strConnection = .GetINIString(.gINIFile, "Settings", "SaccoConnection", "?")
    If strConnection = "?" Then
        MsgBox "Sacco Master could not establish connection to the database" & vbCrLf & _
        " Please carry out database connection configuration." & vbCrLf & _
        " Click OK to carry out the Configuration.", vbInformation, "Initilization"
        frmODBCLogon.Show vbModal
        If .ConnectDatabase Then
            frmLogin.Show
        Else
            MsgBox "Consult your Systems Administrator for help"
        End If
    Else
        If .ConnectDatabase Then
            frmLogin.Show
        Else
            MsgBox "Could not connect to the database"
            frmODBCLogon.Show
            If .ConnectDatabase Then
                frmLogin.Show
            End If
        End If
    End If
End With

Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Public Function ComputePrincipal(ByVal strRepayMethod As String, ByVal cAmount As Currency, _
                           dblRate As Double, dblTime As Double, cBalance As Currency) As Currency


Dim loanbal As Currency
Dim Rate As Integer
Dim dLoanAmount As Double
Dim dInterest As Double
Dim dPrinciple  As Double
Dim LastPayNo As Integer
Dim RepMethod As String
Dim rsrepay As New Recordset
Dim MonthlyPMT As Currency
Select Case strRepayMethod
    Case "AMRT"
              'The loan is in progress
            MonthlyPMT = Pmt(((dblRate / 12) / 100), dblTime, -(cAmount), 0)
            dInterest = (dblRate / 12) / 100 * cBalance
            dPrinciple = MonthlyPMT - dInterest
    Case "STL"
            dPrinciple = cAmount / dblTime
            dInterest = (dblRate / 100 * cAmount) / dblTime
            MonthlyPMT = dPrinciple + dInterest
    Case "RBAL"
            dPrinciple = cAmount / dblTime
            dInterest = (dblRate / 100 * cBalance) / 12   'same as for amortized
            MonthlyPMT = dPrinciple + dInterest
    Case Else
        Exit Function
    End Select
    ComputePrincipal = dPrinciple
End Function

Public Function ComputeInterest(ByVal strRepayMethod As String, ByVal cAmount As Currency, _
                           dblRate As Double, dblTime As Double, cBalance As Currency) As Currency
Dim loanbal As Currency
Dim Rate As Integer
Dim dLoanAmount As Double
Dim dInterest As Double
Dim dPrinciple  As Double
Dim LastPayNo As Integer
Dim RepMethod As String
Dim rsrepay As New Recordset
Dim MonthlyPMT As Currency

Select Case strRepayMethod

    Case "AMRT"
            'The loan is in progress
            MonthlyPMT = Pmt(((dblRate / 12) / 100), dblTime, -(cAmount), 0)
            dInterest = (dblRate / 100) / 12 * cBalance
            dPrinciple = MonthlyPMT - dInterest
      
    Case "STL"
    
            dPrinciple = cAmount / dblTime
            dInterest = (dblRate / 100 * cAmount) / dblTime
            MonthlyPMT = dPrinciple + dInterest
    
    Case "RBAL"
        
            dPrinciple = cAmount / dblTime
            dInterest = (dblRate / 100 * cBalance) / 12   'same as for amortized
            MonthlyPMT = dPrinciple + dInterest
    
    Case Else
        Exit Function
    End Select
    ComputeInterest = dInterest
End Function

Public Function RepayLoan(cLoanAMount As Currency, _
                        StrLoanNo As String, _
                        StringMemberNO As String, _
                        cAmountAwared As Currency, _
                        dblRate As Double, _
                        dblPeriod As Double) As Currency ' Returns balance after servicing LoaNo
                        
    Dim rsLoanInf As New Recordset

End Function

Public Function GetMembersLoans(strMemberNo As String) As Recordset
Dim rsRecordset As New Recordset
strSQL = "SELECT LoanNo, MemberNo FROM LOANBAL WHERE (MemberNo = '" & strMemberNo & "') AND (Balance > 0.1)"
Set rsRecordset = oSaccoMaster.GetRecordSet(strSQL)
Set GetMembersLoans = rsRecordset
End Function

Public Function GetMemberContrib() As Recordset
Dim rsRecordset As New Recordset
strSQL = "Select MemberNew, NewContr From SharVar where MemberNo='" & strMemberNo & "'"
oSaccoMaster.GetRecordSet (strSQL)
Set GetMemberContrib = rsRecordset

End Function
Public Function serviceChristmas(cAmount As Currency, _
                      ExpectedContrib As Currency, _
                      strMemberNo As String, IntRef As Long, _
                      Optional blRequireBalance As Boolean, Optional T As Date) As Currency
'On Error GoTo errorhandler

Dim rsRecordset As New Recordset
Dim rsTemp As New Recordset
Dim cBalance As Currency
Dim cShareBal As Currency
Dim cActualContr As Currency
 If blRequireBalance Then
    If cAmount <= ExpectedContrib Then
        serviceChristmas = 0
    Else
        serviceChristmas = cAmount - ExpectedContrib
    End If
Else
    serviceChristmas = 0
    cActualContr = cAmount
End If

strSQL = "Select * From ChristmusCONTRIB Where MemberNo='" & strMemberNo & "' And RefNo=" & IntRef

Set rsRecordset = oSaccoMaster.GetRecordSet(strSQL)
'// get the total shares from the shares table.
Dim rstotalshares As Recordset

strSQL = "Select totalshares From Christmas Where MemberNo='" & strMemberNo & "'"

Set rstotalshares = oSaccoMaster.GetRecordSet(strSQL)
If Not rstotalshares.EOF Then
cShareBal = rstotalshares.Fields(0)
End If

With rsRecordset
If rsRecordset.EOF Then
cShareBal = cShareBal
Else
   If Not IsNull(!shareBal) Then cShareBal = cShareBal Else cShareBal = !shareBal
   
End If

''// not sure
Dim notsure As Currency
Dim thisamount As Currency


notsure = IIf(Not blRequireBalance, cActualContr, ExpectedContrib)
    If Not blRequireBalance Then
        thisamount = cActualContr
    Else
        thisamount = ExpectedContrib
    End If
    strSQL = ""

    strSQL = "set dateformat dmy Insert into ChristmusCONTRIB(memberno,contrdate,refno,amount,shareBal,transby" _
    & ",chequeno,Locked,posted,remarks,auditid,audittime)VALUES('" & strMemberNo & "'" _
    & ",'" & T & "','" & IntRef + 1 & "'," & thisamount & "," & cShareBal + notsure & "" _
    & ",'Check Off',0,'No',0,'No remarks','" & User & "','" & Now & "')"
    
    oSaccoMaster.ExecuteThis (strSQL)
    

strSQL = ""
strSQL = "set dateformat dmy Update Christmas set TotalShares=" & cShareBal + notsure & _
",transDate=" & "'" & Format(Get_Server_Date, "dd/mm/yyyy") & "'" & " Where MemberNo='" & _
strMemberNo & "'"
 oSaccoMaster.ExecuteThis (strSQL)
 Exit Function
ErrorHandler:
 MsgBox Err.Description
End With

End Function
Public Function ServiceShares(cAmount As Currency, _
                      ExpectedContrib As Currency, _
                      strMemberNo As String, IntRef As Long, _
                      Optional blRequireBalance As Boolean, Optional T As Date) As Currency
'On Error GoTo errorhandler

Dim rsRecordset As New Recordset
Dim rsTemp As New Recordset
Dim cBalance As Currency
Dim cShareBal As Currency
Dim cActualContr As Currency
Dim tt As String
 If blRequireBalance Then
    If cAmount <= ExpectedContrib Then
        ServiceShares = 0
    Else
        ServiceShares = cAmount - ExpectedContrib
    End If
Else
    ServiceShares = 0
    cActualContr = cAmount
End If

strSQL = "Select * From Contrib Where MemberNo='" & strMemberNo & "' And RefNo=" & IntRef

Set rsRecordset = oSaccoMaster.GetRecordSet(strSQL)
'// get the total shares from the shares table.
Dim rrr As New ADODB.Recordset
Set rrr = oSaccoMaster.GetRecordSet("select companycode from members where memberno='" & strMemberNo & "'")
If Not rrr.EOF Then
tt = rrr.Fields(0)
End If
Dim rstotalshares As Recordset

strSQL = "Select totalshares From shares Where MemberNo='" & strMemberNo & "'"

Set rstotalshares = oSaccoMaster.GetRecordSet(strSQL)
If Not rstotalshares.EOF Then
cShareBal = rstotalshares.Fields(0)
End If

With rsRecordset
If rsRecordset.EOF Then
cShareBal = cShareBal
Else
   If Not IsNull(!shareBal) Then cShareBal = cShareBal Else cShareBal = !shareBal
   
End If

''// not sure
Dim notsure As Currency
Dim thisamount As Currency


notsure = IIf(Not blRequireBalance, cActualContr, ExpectedContrib)
    If Not blRequireBalance Then
        thisamount = cActualContr
    Else
        thisamount = ExpectedContrib
    End If
    



'    .AddNew
'    !memberno = strMemberNo
'    !contrdate = t
'    !refno = IntRef + 1
'    If Not blRequireBalance Then
'        !amount = cActualContr
'    Else
'        !amount = ExpectedContrib
'    End If
'    !shareBal = cShareBal + IIf(Not blRequireBalance, cActualContr, ExpectedContrib)
'    !transby = "Cash"
'    !chequeno = ""
'    !Locked = "No"
'    !posted = "No"
'    !remarks = ""
'    !auditid = User
'    '!audittime = Now()
'    '.Update
'
If thisamount <= 1 Then GoTo TAILA
    
    strSQL = ""

    strSQL = "set dateformat dmy Insert into contrib(memberno,contrdate,refno,amount,shareBal,transby" _
    & ",chequeno,Locked,posted,remarks,auditid,audittime,receiptno)VALUES('" & strMemberNo & "'" _
    & ",'" & T & "','" & IntRef + 1 & "'," & thisamount & "," & cShareBal + notsure & "" _
    & ",'" & tt & "',0,'No',0,'" & tt & "','" & User & "','" & Now & "','" & tt & "')"
    
    oSaccoMaster.ExecuteThis (strSQL)
    

strSQL = ""
strSQL = "set dateformat dmy Update Shares set TotalShares=" & cShareBal + notsure & _
",transDate=" & "'" & Format(Get_Server_Date, "dd/mm/yyyy") & "'" & " Where MemberNo='" & _
strMemberNo & "'"
 oSaccoMaster.ExecuteThis (strSQL)
 
 '//update the dataimporta so that no repear job is done.
 sql = ""
 sql = "update dataimporta set posted=1 where memberno='" & strMemberNo & "' and posted=0"
 cn.Execute sql
TAILA:
 Exit Function
ErrorHandler:
 MsgBox Err.Description
End With
End Function


Public Function GetMemberLoanInf(strMemberNo As String) As Recordset
Dim rsRecordset As New Recordset
    Set rsRecordset = oSaccoMaster.GetRecordSet("SELECT MemberNo, LoanNo, Balance, RepayRate, RepayMethod, RepayPeriod, Interest, FirstDate, LastDate, Cleared, AutoCalc, IntrAmount, Remarks, AuditID, AuditTime FROM LOANBAL WHERE (Balance > 0.1) and MemberNO='" & strMemberNo & "' ORDER BY MemberNo")
   Set GetMemberLoanInf = rsRecordset

End Function

Public Function GetMembersLastLoanRepayNo(strMemberNo As String, StrLoanNo As String) As Integer
Dim rsRecordset  As New Recordset
Set rsRecordset = oSaccoMaster.GetRecordSet("SELECT MemberNO, MAX(PaymentNo) AS LastLoanRepNo FROM REPAY where MemberNo ='" & strMemberNo & "' and LoanNo='" & StrLoanNo & "' Group by MemberNO")
With rsRecordset
    If .RecordCount > 0 Then
        .MoveFirst
        GetMembersLastLoanRepayNo = !LastLoanRepNo
    Else
        GetMembersLastLoanRepayNo = 1
    End If
End With
End Function

Public Function GetMembersAwardedLoan(strMemberNo As String, _
                                      StrLoanNo As String) As Currency
                                      
Dim rsRecordset As New Recordset

Set rsRecordset = oSaccoMaster.GetRecordSet("SELECT LoanNo," & _
"LoanBalance AS LoanBal," & _
"Principal FROM REPAY WHERE " & _
"(MemberNo='" & strMemberNo & "') AND (LoanNO='" & StrLoanNo & "') AND (PaymentNo IN " & _
"(SELECT MIN(PAYMENTNO) AS MinNo FROM REPAY))")
If IsNull(rsRecordset!loanbal) Then
    GetMembersAwardedLoan = 0
Else
    GetMembersAwardedLoan = rsRecordset!loanbal + rsRecordset!principal
End If
End Function

Public Sub UpdateLoanBalance(StrLoanNo As String, WithAmount As Currency)
strSQL = "Update LoanBal set Balance=" & WithAmount & " Where LoanNo='" & StrLoanNo & "'"
    oSaccoMaster.ExecuteThis strSQL

End Sub

Public Function GetIntrOwed(strMemberNo As String, StrLoanNo As String, intLastPaymentNo As Integer) As Currency
Dim rsRecordset As New Recordset
Set rsRecordset = oSaccoMaster.GetRecordSet("SELECT MemberNO, PaymentNo, IntrOwed, LoanNo FROM REPAY WHERE (MemberNO ='" & Trim(strMemberNo) & "') AND (LoanNo ='" & StrLoanNo & "') AND (PaymentNo =" & GetMembersLastLoanRepayNo(strMemberNo, StrLoanNo) & ")")
With rsRecordset
    If Not IsNull(rsRecordset!IntrOwed) Then
        GetIntrOwed = Round(!IntrOwed, 0)
    Else
        GetIntrOwed = 0
    End If
End With
End Function

Public Function GetPreviousMonth(dteMonth As Integer, intYear As Integer) As Date

Select Case dteMonth
    Case 1
        GetPreviousMonth = 31 & "/" & 12 & "/" & intYear - 1
    Case 3, 5, 7, 8, 10, 12
        GetPreviousMonth = (31 & "/" & dteMonth - 1 & "/" & intYear)
    Case 2
        GetPreviousMonth = 28 & "/" & dteMonth - 1 & "/" & intYear
    Case 4, 6, 9, 11
        GetPreviousMonth = 30 & "/" & dteMonth - 1 & "/" & intYear
End Select

End Function

Public Sub Generate_Statement(memberno As String, BeginDate As Date, EndDate As Date)
    On Error GoTo ErrorTrap
    Dim RsMembers As New Recordset
    Dim rsPayments As New Recordset
    Dim rsNewLoans As New Recordset
    Dim RsShares As New Recordset
    BeginDate = DateSerial(Year(BeginDate), month(BeginDate), 1)
    EndDate = DateSerial(Year(EndDate), month(EndDate) + 1, 1 - 1)
    Do Until BeginDate > EndDate
        Set li = frmUtilGenMemStatements.lvwGenMemStatements.ListItems.Add(, , Format(DateSerial(Year(BeginDate), month(BeginDate), 1), "MM-YYYY"))
'        Set rsNewLoans = oSaccoMaster.GetRecordSet("select sum(Amount) as NewLoans " _
'        & "from CHEQUES C inner join LOANBAL LB on C.LoanNo=LB.LoanNo where " _
'        & " month(DateIssued)=" & Month(BeginDate) & " and Year(DateIssued)" _
'        & "=" & Year(BeginDate) & " and LB.MemberNo='" & MemberNo & "'")
'        With rsNewLoans
'            Li.SubItems(2) = Format(IIf(IsNull(!NewLoans), 0, !NewLoans), CfMt)
'        End With
'        Set rsPayments = oSaccoMaster.GetRecordSet("select sum(Principal) as Payment," _
'        & "sum(Interest) as IntPaid from REPAY R inner join LOANBAL LB on R.LoanNo=" _
'        & "LB.LoanNo where Month" _
'        & "(R.DateReceived)=" & Month(BeginDate) & " And Year(R.DateReceived)=" _
'        & "" & Year(BeginDate) & " and LB.MemberNo='" & MemberNo & "'")
'        With rsPayments
'            Li.SubItems(3) = Format(IIf(IsNull(!Payment), 0, !Payment), CfMt)
'            Li.SubItems(4) = Format(IIf(IsNull(!IntPaid), 0, !IntPaid), CfMt)
'        End With
'        Set rsShares = oSaccoMaster.GetRecordSet("select InitShares from MEMBERS where " _
'        & "MemberNo='" & MemberNo & "'")
'        Li.SubItems(10) = Format(rsShares!IniTShares, CfMt)
'        Set rsShares = oSaccoMaster.GetRecordSet("select sum(Amount) as IniShares from " _
'        & "CONTRIB where MemberNo='" & MemberNo & "' and ContrDate<#" & BeginDate & "#")
'        With rsShares
'            Li.SubItems(10) = IIf(IsNull(!IniShares), 0, !IniShares)
'        End With
'        Set rsShares = oSaccoMaster.GetRecordSet("select sum(Amount) as Shares from " _
'        & "CONTRIB where month(ContrDate)=" & Month(BeginDate) & " and Year(ContrDate)" _
'        & "=" & Year(BeginDate) & " and MemberNo='" & MemberNo & "'")
'        With rsShares
'            Li.SubItems(9) = Format(IIf(IsNull(!SHARES), 0, !SHARES), CfMt)
'            Li.SubItems(10) = Format(CDbl(Li.SubItems(10)) + CDbl(Li.SubItems(9)), CfMt)
'        End With
        BeginDate = DateSerial(Year(BeginDate), month(BeginDate) + 1, 1)
    Loop
    Set RsShares = Nothing
    Set rsNewLoans = Nothing
    Set rsPayments = Nothing
    frmUtilGenMemStatements.Caption = ""
    Exit Sub
ErrorTrap:
    MsgBox Err.Description, , "Member Statement"
End Sub
Public Sub Gen_Statement(memberno As String, FirstDate As Date, EndDate As Date)
    On Error GoTo ErrorTrap
    Screen.MousePointer = vbHourglass
    Dim TrDate As Date
    TrDate = FirstDate
    Dim RsMembers As New Recordset
    
    '****************************** Get Member Details ****************************'
    
    Set RsMembers = oSaccoMaster.GetRecordSet("select * from members where memberno=" _
    & "'" & memberno & "'")
    With RsMembers
    
        Dim LB As Double
        Dim IntMonths As Integer
        Dim intCount As Integer
        Dim OpeningBal As Currency
        IntMonths = DateDiff("m", FirstDate, EndDate) + 1
        frmUtilGenMemStatements.ProgressBar1.Max = IntMonths
        frmUtilGenMemStatements.ProgressBar1.Visible = True
        FirstDate = DateSerial(Year(FirstDate), month(FirstDate), 1)
        For intCount = 1 To IntMonths
        
        '*********************** Opening LoanBalance **************************'
            
            strSQL = "SELECT Sum(C.Amount) as Amount From CHEQUES C inner " _
            & "join LOANBAL L on C.LoanNo=L.LoanNo WHERE " & _
            "L.FirstDate <=" & Format(DateSerial(Year(FirstDate), month(FirstDate) + 1, 1 - 1), "MM-dd-yyyy") & "" _
            & " AND L.MemberNo='" & memberno & "'"
            Dim rsTotal_Loans_Given_As_At As Recordset
            Set rsTotal_Loans_Given_As_At = oSaccoMaster.GetRecordSet(strSQL)
            If rsTotal_Loans_Given_As_At.EOF Then
                OpeningBal = 0
            Else
                OpeningBal = IIf(IsNull(rsTotal_Loans_Given_As_At!amount), 0, rsTotal_Loans_Given_As_At!amount)
            End If
            Set RsLoans = oSaccoMaster.GetRecordSet("select sum(R.Principal) as Pay" _
            & " FROM REPAY R inner join LOANBAL LB on R.LoanNo=LB.LoanNo WHERE " _
            & "LB.MemberNo='" & memberno & "' and R.Datereceived" _
            & "<=" & Format(DateSerial(Year(FirstDate), month(FirstDate), 1 - 1), "MM-dd-yyyy") & "")
            
            '********************** New Loans Taken ***************************'
            
            Set Itm = frmUtilGenMemStatements.lvwGenMemStatements.ListItems.Add(, , Format(FirstDate, "yyyy-mm"))
            
            strSQL = "SELECT CHEQUES.Amount" & _
            " FROM CHEQUES INNER JOIN LOANBAL ON CHEQUES.LoanNo = LOANBAL.LoanNo" & _
            " Where LoanBal.MemberNo ='" & memberno & "'" _
            & " And Month(LOANBAL.FirstDate) =" & month(FirstDate) _
            & " AND year(LOANBAL.FirstDate)=" & Year(FirstDate)
            Dim rsNewLoans As Recordset
            Set rsNewLoans = oSaccoMaster.GetRecordSet(strSQL)
            
            If rsNewLoans.RecordCount > 0 Then
                Itm.SubItems(2) = Format(rsNewLoans!amount, Cfmt)
            Else
                Itm.SubItems(2) = Format(0, Cfmt)
            End If
            
            strSQL = "SELECT LOANBAL.MemberNo, Sum(REPAY.Principal) AS SumOfPrincipal" & _
            ",sum(REPAY.Interest) as IntPaid,sum(REPAY.IntrCharged) as Charge," _
            & "sum(REPAY.LoanBalance) as BaLnce," _
            & "sum(REPAY.IntrOwed) as Owed,sum(REPAY.LoanBalance) as LBalance FROM LOANBAL INNER JOIN REPAY ON LOANBAL.LoanNo = REPAY.LoanNo" & _
            " Where month(REPAY.DateReceived)=" & month(FirstDate) & " and " _
            & "Year(REPAY.Datereceived)=" & Year(FirstDate) & " and MemberNO='" & memberno & "' GROUP BY LOANBAL.MemberNo"
            Dim rsTotal_Payments_As_At As Recordset
            Set rsTotal_Payments_As_At = oSaccoMaster.GetRecordSet(strSQL)
            
            If RsLoans.EOF Then
                OpeningBal = OpeningBal - CDbl(Itm.SubItems(2))
            Else
                OpeningBal = OpeningBal - CDbl(Itm.SubItems(2)) - IIf(IsNull(RsLoans!Pay), 0, RsLoans!Pay)
            End If
            
            Itm.SubItems(1) = Format(OpeningBal, Cfmt)
            OpeningBal = 0
            Itm.SubItems(3) = Format(IIf(rsTotal_Payments_As_At.EOF, 0, rsTotal_Payments_As_At!SumOfPrincipal), Cfmt)
            Itm.SubItems(4) = Format(IIf(rsTotal_Payments_As_At.EOF, 0, rsTotal_Payments_As_At!intpaid), Cfmt)
            Itm.SubItems(7) = CCur(Itm.SubItems(1)) + CCur(Itm.SubItems(2)) - CCur(Itm.SubItems(3))
            Itm.SubItems(7) = Format(Itm.SubItems(7), Cfmt)
            Itm.SubItems(5) = Format(IIf(rsTotal_Payments_As_At.EOF, 0, rsTotal_Payments_As_At!Charge), Cfmt)
            Itm.SubItems(6) = Format(IIf(rsTotal_Payments_As_At.EOF, 0, rsTotal_Payments_As_At!Owed), Cfmt)
            strSQL = "SELECT Sum(CONTRIB.Amount) AS SumOfAmount" & _
            " ,sum(CONTRIB.Sharebal) as TotShares From CONTRIB" & _
            " WHERE CONTRIB.MemberNo='" & memberno & "'" _
            & " AND Month([ContrDate]) =" & month(FirstDate) & " And Year([ContrDate]) =" & Year(FirstDate)
            Dim rsMonthContr As Recordset
            Set rsMonthContr = oSaccoMaster.GetRecordSet(strSQL)
            If rsMonthContr.RecordCount > 0 And Not IsNull(rsMonthContr!SumofAmount) Then
                Itm.SubItems(9) = Format(rsMonthContr!SumofAmount, Cfmt)
            Else
                Itm.SubItems(9) = Format(0, Cfmt)
            End If
            If Not rsMonthContr.EOF Then
                Itm.SubItems(10) = Format(IIf(IsNull(rsMonthContr!TotShares), Shares, rsMonthContr!TotShares), Cfmt)
                Shares = Itm.SubItems(10)
            Else
                Itm.SubItems(10) = Shares
            End If
            
            Itm.SubItems(8) = Format(CDbl(Itm.SubItems(10)) - CDbl(Itm.SubItems(9)), Cfmt)
            Itm.SubItems(11) = Format(CDbl(Itm.SubItems(3)) + CDbl(Itm.SubItems(4)) + CDbl(Itm.SubItems(9)), Cfmt)
            FirstDate = DateSerial(Year(FirstDate), month(FirstDate) + 1, Day(FirstDate))

        Next intCount
        Screen.MousePointer = vbDefault
        FirstDate = TrDate
    End With
    Exit Sub

ErrorTrap:
    Select Case Err.number
    Case 3704
        MsgBox "Can Not search an empty List"
    Case Else
        'Resume Next
    End Select
Screen.MousePointer = vbNormal

End Sub
Private Function SearchSubItems(lv As ListView, col As Integer, str As String) As Integer
   SearchSubItems = -1
   
   For i = 1 To lv.ListItems.Count
       If lv.ListItems(i).ListSubItems(col).Text = str Then
          SearchSubItems = i
          Exit For
       End If
   Next
End Function
Public Sub setEncryption()
    'Set oCryptoX = New CryptoX
    'oCryptoX.Keyword = "!E%65786guhy*^*()_(+"
    'oCryptoX.Method = StreamEncryption
        
End Sub
