Attribute VB_Name = "Cosmodule"
'Auther: Cosmas Kibet Ngeno
'Date: Friday Jan 2010
'Descriptin:    This module is suppossed to be a global loan repayment schedule
'               that receives a loanno as a sole parameter and returns the principal and interest for that month
'               taking care of three loan repament methods. (STL,RBA and ARMT)
'Note: this should stop any other loan calculation in other module throughout the project
'       it's designed to work with easysacco
'*********************************END OF NOTES*******************************************
Public Forms() As Form
'Public currentUser As Current_User
Public InvoiceNo As String
Public OpeningBal As Double
Public InvoiceBal As Double
Public num_forms As Integer
Public processid As String
Public ActionOnInteretDefaulted As Integer
Public IntBalalance As Double
Public Authority As String
Public memStatus As String
'Public LoanBalance As Double
Public FirstDate As Date, nextdate As Date
Public TotalDr As Double, TotalCr As Double
'Public OpeningBal As Double
Public principal As Double
Public mrepayment As Double
Public interest As Double
Public RepayMode As Integer
Public Penalty As Double
Public intOwed As Double
Public intcharged As Double
Public REarningAcc As String, DRaccno As String, Craccno As String, MPurchaseAcc As String, MSalesAcc As String
Public activeTno As String
'Public success As Boolean
Public LBalance As Currency
Public duedate As Date
Public newMember As Boolean
Public loops As Integer
Public rpInterest As Double
Public totalrepayable As Double
Public RepayableInterest As Double
Public rmethod As String, intRecovery As String
Public Rperiod As Integer, mdtei As Integer
Public rrate As Double
Public initialAmount As Currency
Public transactionNo As String
Public transactionTotal As Double
Public lastrepay As Date, dateissued As Date
'Public success As Boolean
Public daysIntoTheMonth As Integer
Public saveToGl As Boolean
Public sharesCode As String
Public amount As Double
Public serverDate As Date
Public penaltyAcc As String
Public thisMember As Sacco_Member
Public wePenalize As Boolean
Public BPRICE As Double
Public BVat As Double

Public Type Sacco_Member
    NAMES As String
    companycode As String
    idno As String
    sex As String
    CompanyName As String
    payrollno As String
End Type
Public Type Current_User
    username As String
    isTeller As Boolean
    tellerGlAcc As String
    idno As String
End Type

Dim Today As Date
Dim transConn As ADODB.Connection
Dim actualInterest As Double




Public Function Ceiling(value As Double) As Double
    Ceiling = IIf(Int(value) = value, value, Int(value) + 1)
End Function
Public Function GetFirstDate(mdate As Date) As Date
    mdate = DateAdd("m", 1, mdate)
    GetFirstDate = DateSerial(Year(mdate), month(mdate), 1)
End Function
Public Function SelectAllText(tb As TextBox)
    tb.SelStart = 0
    tb.SelLength = Len(tb.Text)
End Function
Public Sub ShowErrorMessage(errmsg As String)
    With frmErrorMsg
        errmsg = IIf(errmsg = "", ErrorMessage, errmsg)
        .lblUsersMsg.Caption = "OOPS! There was an Execution Error! Try to Identify or consult the System Admin"
        .txtDetailedMsg.Text = "Error Details: " & vbNewLine & errmsg
        .Show vbModal
    End With
End Sub
Public Function getGlBalance(ACCNO As String, Startdate As Date, Enddate As Date) As Double
 On Error GoTo Capture
    Dim rsGls As ADODB.Recordset
    Dim OBal As Double
    sql = "set dateformat dmy select gl.Normalbal,op.Cbal, gl.GlAccType from dbo.UDF_GL_OpeningBalance ('" & ACCNO & "','" & Startdate & "') op inner join glsetup gl on op.accno=gl.accno where gl.accno='" & ACCNO & "'"
    Set rst = oSaccoMaster.GetRecordset(sql)
    With rst
    
        OBal = rst(1)
         'OBal = 0
        NormalBal = rst("NormalBal")
'          If ACCNO = "102A" Then
'           MsgBox "hI MUTAHI"
'          End If
          
        'transactions between the dates
        sql = "SET DATEFORMAT DMY Select " _
        & " (select ISNULL(sum(amount),0) " _
        & " from gltransactions " _
        & " where Draccno='" & ACCNO & "' and DATEADD(dd, DATEDIFF(dd, 0, TRANSDATE), 0)>='" & Startdate & "' and DATEADD(dd, DATEDIFF(dd, 0, TRANSDATE), 0)<='" & Enddate & "')DR," _
        & " (select ISNULL(sum(amount),0) " _
        & " from gltransactions " _
        & " where Craccno='" & ACCNO & "' and DATEADD(dd, DATEDIFF(dd, 0, TRANSDATE), 0)>='" & Startdate & "' and DATEADD(dd, DATEDIFF(dd, 0, TRANSDATE), 0)<='" & Enddate & "')CR"
        
        Set rsGls = oSaccoMaster.GetRecordset(sql)
        
        If NormalBal = "Debit" Then
            getGlBalance = rsGls("DR") - rsGls("CR")
            'getGlBalance = OBal + rsGls("DR") - rsGls("CR")
        Else
            'getGlBalance = OBal + rsGls("CR") - rsGls("DR")
            
            getGlBalance = rsGls("CR") - rsGls("DR")
        End If
        
    End With
    success = True
    OBal = True
    Exit Function
Capture:
    success = False
End Function
Public Function getGlBalance1(ACCNO As String, Startdate As Date, Enddate As Date) As Double
 On Error GoTo Capture
    Dim rsGls As ADODB.Recordset
    Dim OBal As Double
    sql = "set dateformat dmy select gl.Normalbal,op.Cbal, gl.GlAccType from dbo.UDF_GL_OpeningBalance ('" & ACCNO & "','" & Startdate & "') op inner join glsetup gl on op.accno=gl.accno where gl.accno='" & ACCNO & "'"
    Set rst = oSaccoMaster.GetRecordset(sql)
    With rst
    
        'OBal = rst(1)
         OBal = 0
        NormalBal = rst("NormalBal")
'          If ACCNO = "102A" Then
'           MsgBox "hI MUTAHI"
'          End If
          
        'transactions between the dates
        sql = "SET DATEFORMAT DMY Select " _
        & " (select ISNULL(sum(amount),0) " _
        & " from gltransactions " _
        & " where Draccno='" & ACCNO & "' and DATEADD(dd, DATEDIFF(dd, 0, TRANSDATE), 0)>='" & Startdate & "' and DATEADD(dd, DATEDIFF(dd, 0, TRANSDATE), 0)<='" & Enddate & "')DR," _
        & " (select ISNULL(sum(amount),0) " _
        & " from gltransactions " _
        & " where Craccno='" & ACCNO & "' and DATEADD(dd, DATEDIFF(dd, 0, TRANSDATE), 0)>='" & Startdate & "' and DATEADD(dd, DATEDIFF(dd, 0, TRANSDATE), 0)<='" & Enddate & "')CR"
        
        Set rsGls = oSaccoMaster.GetRecordset(sql)
        
        If NormalBal = "Debit" Then
            getGlBalance1 = OBal + rsGls("DR") - rsGls("CR")
        Else
            getGlBalance1 = OBal + rsGls("CR") - rsGls("DR")
        End If
        
    End With
    success = True
    Exit Function
Capture:
    success = False
End Function

Public Function SaveTransaction(transactionNo As String, amount As Double, User As String, transdate As Date, description As String, Optional mMemberNo As String) As Boolean

    sql = "set dateformat dmy Insert into transactions(transactionno,amount,auditid,TransDate,transDescription)" _
    & " Values('" & transactionNo & "'," & amount & ",'" & User & "',Convert(Varchar(10), '" & transdate & "', 101),'" & description & "')"
    If Not oSaccoMaster.Execute(sql) Then
        SaveTransaction = False
    Else
        SaveTransaction = True
    End If
End Function
Public Sub GetTransactionNo()
    Dim TimeNow
    TimeNow = Get_Server_Date
    transactionNo = User & Day(TimeNow) & CStr(month(TimeNow)) & CStr(Year(TimeNow)) & CStr(Format(TimeNow, "hh:mm:ss:ampm")) ' & Min & CStr(Second(TimeNow))
End Sub
Public Sub NewTransaction(AmountPaid As Double, transdate As Date, description As String)
        'save TransactionNo
        GetTransactionNo
        transactionTotal = AmountPaid

        If Not SaveTransaction(transactionNo, transactionTotal, User, transdate, description) Then
            GoTo Capture
        End If
        Exit Sub
Capture:
        ErrorMessage = ErrorMessage
End Sub
'Public Function SelectAllText(tb As TextBox)
'    tb.SelStart = 0
'    tb.SelLength = Len(tb.Text)
'End Function

Function UpdateList(lst As ListView, frmName As Form)
    Set li = lst.ListItems.Add(, , frmName.name)
    ''lst.AddItem frmName.name
    num_forms = num_forms + 1
End Function
Public Sub UnlockControls(ByVal frm As Form)
    On Error GoTo Capture
    Dim ctrl As Control
    For Each ctrl In frm
        If TypeOf ctrl Is TextBox Then
            ctrl.Locked = False
        End If
    Next ctrl
    Exit Sub
Capture:
    MsgBox IIf(ErrorMessage = "", err.description, ErrorMessage)
End Sub
Public Sub lockkControls(ByVal frm As Form)
    On Error GoTo Capture
    Dim ctrl As Control
    For Each ctrl In frm
        If TypeOf ctrl Is TextBox Then
            ctrl.Locked = True
        End If
    Next ctrl
    Exit Sub
Capture:
    MsgBox IIf(ErrorMessage = "", err.description, ErrorMessage)
End Sub
Public Function saveReceipt(ReceiptNo As String, ref, RefNo, mMemberNo As String, ccode As String, name As String, transdate As Date, amount As Double, chequeno As String, ptype As String, activity As String) As Boolean
    On Error GoTo Capture
            ErrorMessage = ""
            If activity = "Deposit" Then
                sql = ""
                sql = "set dateformat dmy INSERT INTO ReceiptBooking (ReceiptNo,Ref,Refno,MemberNo,Ccode,Name,Transdate," _
                & "Amount, Chequeno, ptype, auditid,transactionno) VALUES ('" & ReceiptNo & "','" & ref & "','" & RefNo & "','" & _
                mMemberNo & "','" & ccode & "','" & name & "','" & transdate & "'," & amount & ",'" & _
                chequeno & "','" & ptype & "','" & User & "','" & transactionNo & "')"
            Else
                sql = "set dateformat dmy INSERT INTO PaymentBooking (VoucherNo,Memberno,Ccode,Name,Transdate," _
                & "Amount, Chequeno, Ptype, auditid,datedeposited,PayeeDesc,Transactionno) VALUES ('" & ReceiptNo & "','" & _
                mMemberNo & "','" & ccode & "','" & name & "" _
                & "','" & transdate & "'," & amount & ",'" & _
                chequeno & "','" & ptype & "','" & User & "','" & transdate & "','" & name & "','" & transactionNo & "')"
            End If
            If Not oSaccoMaster.Execute(sql) Then
                GoTo Capture
            Else
                saveReceipt = True
            End If
    Exit Function
Capture:
    saveReceipt = False
    ErrorMessage = err.description
End Function

Public Function isValidPassword(Password As String) As Boolean

    'The password rules
    
    '1. Password should be a minimum of eight characters
    '2. Should not contain repeating characters
    '3. Password should contain special characters
    '4. should not begin with numbers
    '5. Password should start with a capital alphabet
    '6. Password should contain a number
    
    'Auther: Cosmas Ngeno
    
    Dim X As Long
    Const SpecialCharacters = "@#$%^&*"
    If Len(Password) >= 8 And Password Like "[A-Z]*[0-9]*" And Password Like "*[" & SpecialCharacters & "]*" Then
        For X = 1 To Len(Password) - 1
            If InStr(mid$(Password, X + 1), mid$(Password, X, 1)) Then
                Exit Function
            End If
        Next X
        isValidPassword = True
    End If
End Function
Public Function SaveGLTRANSACTION(transdate As Date, amount As Double, DRaccno As String, _
Craccno As String, DocumentNo As String, Source As String, transDescription As String, auditid As String, transactionNo As String) As Boolean
    On Error GoTo SysError
Dim cn As New Connection
    With cn
        If .State = adStateClosed Then
            .Open SelectedDsn, "atm", "atm"
        End If
        
        sql = "set dateformat dmy insert into gltransactions (TransDate,Amount,DrAccNo,CrAccNo,DocumentNo,Source,AuditID,TransDescript,transactionno)" _
        & " Values('" & Format(transdate, "DD/MM/YYYY") & "'," & amount & ",'" & DRaccno & "','" & Craccno & "','" & DocumentNo & "','" & Source & "','" & auditid & "','" & transDescription & "','" & transactionNo & "')"
        
        .Execute (sql)
    End With
    SaveGLTRANSACTION = True
    Exit Function
SysError:
    ErrorMessage = err.description
    SaveGLTRANSACTION = False
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
Public Function getActiveTno(sno As Integer, mdate As Date)
    Dim rsTno As ADODB.Recordset
    Set rsTno = oSaccoMaster.GetRecordset("set dateformat dmy select Trans_Code tno,rate from d_Transport where sno='" & sno & "' and DATEADD(d, 0, DATEDIFF(dd, 0, '" & mdate & "'))" _
    & " between startdate and isnull(dateinactivate,getdate())")
    If rsTno.EOF Then
        getActiveTno = "SELF"
    Else
        getActiveTno = rsTno("Tno")
    End If
End Function

Public Sub ShowRights(User As String)
    sql = "select alias from tbl_menus order by id"
    Set rs = oSaccoMaster.GetRecordset(sql)
    While Not rs.EOF
    x1 = rs.Fields(0)
    
    Dim x2 As String
    sql = "select [enable] from tbl_usermenus where UserLoginIDs='" & User & "' and [menu]='" & x1 & "'"
    Set Rs1 = oSaccoMaster.GetRecordset(sql)
    If Not Rs1.EOF Then
        MainForm.Controls(x1).Enabled = Rs1.Fields(0)
    Else
        MainForm.Controls(x1).Enabled = False
    End If
    
    rs.MoveNext
Wend
End Sub
Public Function keyIsValid(kAscii As Integer, vSet As Integer) As Boolean
    Const SpecialCharacters = "@#$%^&*'"""
    If Chr(kAscii) Like "*[" & SpecialCharacters & "]*" Then
        keyIsValid = False
        Exit Function
    End If
    Select Case vSet
        Case 1 '"Numeric"
            If Not ((kAscii > 47 And kAscii < 58) Or kAscii = 8 Or kAscii = 13 Or kAscii = 46) Then
                keyIsValid = False
                Exit Function
            End If
        Case 2 '"AlphaNumeric"
            keyIsValid = True
        Case 3 '"Alpha"
            If Not ((kAscii > 96 And kAscii < 123) Or (kAscii > 64 And kAscii < 91) Or kAscii = 8 Or kAscii = 13) Then
                Exit Function
            End If
        Case Else
    End Select
    keyIsValid = True
End Function







