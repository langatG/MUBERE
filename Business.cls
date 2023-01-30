VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Business"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim RAmnt As String
Dim myclass As cdbase
Dim Provider As String
Dim accountnumber As String
Dim amount As String
Dim tellerName As String
Dim CubicleNumber As String
Dim Account As String
Dim description As String
Dim Credit As String
Dim Debit As String
Dim Period As String

Private Sub GLTransaction(strAmount As String, Period As String, strAccountNumber As String)

        '// First check whether the  table has information. If it has no information then
        '// insert new record else update
    
        Set myclass = New cdbase
         
         Debit = ""
         
         Credit = ""
          
        Set rs = CreateObject("adodb.recordset")
            
            ' // Get Account Code being used
            
            sql = ""
            
            sql = "select * from GLAccount where AccountDescription='Cheque Clearing Suspense Account'"
                                        
             Set cn = CreateObject("adodb.connection")
             
             If Provider = "" Then Provider = myclass.OpenCon
                
           cn.Open Provider, "atm", "atm"
            
            rs.Open sql, cn
                                        
            If rs.EOF Then MsgBox "Setup GL Account ID in Account Classes Setup", vbCritical:  Exit Sub
            
            Account = rs!AccountID
            
            description = rs!AccountDescription
            
            Credit = strAmount
            
'            Description = "Cheque Clearing Suspense"
            
            Debit = ""
                    
SaveAgain:
            sql = ""

            sql = "insert into GLTrans(Account,Description,Date,Period,Debit,Credit)"

            sql = sql & "select'" & Account & "','" & description & "','" & Date & "','" & Period & "','" & _
            Debit & "','" & Credit & "'"
            
            myclass.save sql
            
            rs.Close
            
            If Debit <> "" Then GoTo LeaveGLTransaction

            sql = ""
            
            sql = "select * from GLAccount where AccountDescription='Cheque Clearing Account'"
                                        
             Set cn = CreateObject("adodb.connection")
             
             If Provider = "" Then Provider = myclass.OpenCon
                
           cn.Open Provider, "atm", "atm"
            
            rs.Open sql, cn
                                        
            If rs.EOF Then MsgBox "Setup GL Account ID in Account Classes Setup", vbCritical:  Exit Sub
            
            Account = rs!AccountID
            
            description = rs!AccountDescription
                
            Debit = strAmount
                  
'            Description = "Cheque Clearing"
            
            Credit = ""
             
        GoTo SaveAgain

LeaveGLTransaction:

        
        
        
        '// Cheque Clearance transaction
        
            sql = ""
            
            sql = "select * from GLAccount where AccountDescription='Customer Account'"
                                        
             Set cn = CreateObject("adodb.connection")
             
             If Provider = "" Then Provider = myclass.OpenCon
                
           cn.Open Provider, "atm", "atm"
            
            rs.Open sql, cn
                                        
            If rs.EOF Then MsgBox "Setup GL Account ID in Account Classes Setup", vbCritical:  Exit Sub
        
            description = rs!AccountDescription
            
            Account = rs!AccountID
            
            Credit = strAmount
            
            Debit = ""
                    
ShokaHaha:
            sql = ""

            sql = "insert into GLTrans(Account,Description,Date,Period,Debit,Credit)"

            sql = sql & "select'" & Account & "','" & description & "','" & Date & "','" & Period & "','" & _
            Debit & "','" & Credit & "'"
            
            myclass.save sql
            
            If Debit <> "" Then GoTo LeaveClearanceTransaction

            
            sql = ""
            
            sql = "select * from GLAccount where AccountDescription='Cheque Clearing Suspense Account'"
                                        
             Set cn = CreateObject("adodb.connection")
             
             If Provider = "" Then Provider = myclass.OpenCon
                
           cn.Open Provider, "atm", "atm"
            
            rs.Close
                      
            
            rs.Open sql, cn
                                        
            If rs.EOF Then MsgBox "Setup GL Account ID in Account Classes Setup", vbCritical:  Exit Sub
            
            Account = rs!AccountID
                
            Debit = strAmount
                  
            description = rs!AccountDescription
            
            Credit = ""
             
        GoTo ShokaHaha

LeaveClearanceTransaction:
        
        
        
End Sub
Private Function ClearCheques(strAmount As String, strChequeNo)
    Dim AccNumber  As String
    
    '// Get the amount from CustomerBalance

    Set rs = Nothing
    
    sql = ""
    
    sql = "select * from CustomerBalance where ChequeNo='" & strChequeNo & "'"

    

    Set myclass = New cdbase

    If Provider = "" Then Provider = myclass.OpenCon

    Set cn = CreateObject("adodb.connection")

   cn.Open Provider, "atm", "atm"

    Set rs = CreateObject("adodb.recordset")

    rs.Open sql, cn
   
    sql = ""
    sql = "select * from SavingsAccountsParameters"
   ' rs.Open sql, cn

    
    '// Update the customer's balance
    Dim rsn As Object
    Set rsn = CreateObject("adodb.recordset")
    
    
    Dim AvailableBal As String
    Dim actualbal As String
        
    If rs.EOF Then
    
        AvailableBal = 0
        actualbal = 0
    Else
    
        If Not rs!availablebalance & "" = "" Then
        
            AvailableBal = rs!availablebalance
      
            
        Else
        
            AvailableBal = 0

            
        End If
  
    
        AccNumber = rs!ACCNO
        
    End If
    
    sql = ""

    sql = "update customerbalance set AvailableBalance='" & CStr(CCur(AvailableBal) + CCur(strAmount)) & "' where ChequeNo='" & strChequeNo & "'"
''
'' sql = "insert into Customerbalance(CustomerNo,AccName,Amount,AvailableBalance,ActualBalance,AccNO,TransDescription,TransDate,ChequeNo,Period)Values('"& rs!CustomerNo &"','"& rs!AccName &"','"& CCur(strAmount) &"','"& CStr(CCur(AvailableBal) + CCur(strAmount))&"','"& CStr(CCur(ActualBal) + CCur(strAmount)) &"','"& rs!AccNo &"','Cheque','"& strChequeNo &"','"Month (Date)"')"
'
' sql = " INSERT INTO CustomerBalance(CustomerNo, AccName, Amount, AvailableBalance, ActualBalance, AccNO, TransDescription, TransDate, ChequeNo, Period)"
' sql = sql & " VALUES ('" & rs!CustomerNo & "','" & rs!AccName & "' ,'" & CCur(strAmount) & "' ,'" & CStr(CCur(AvailableBal) + CCur(strAmount)) & "' , '" & CStr(CCur(actualbal) + CCur(strAmount)) & "','" & rs!AccNo & "' ,'" & Date & "','Cheque' ,'" & strChequeNo & "' , Month(Date))"
'
    myclass.save sql
    'Set MyClass = Nothing
  '//put code here for updating the commission for the each and every cheque cleared

    
    '// Update GL Transaction
    
   'If Not rs.EOF Then GLTransaction strAmount, AccNumber
    
    
    '// Delete record from ChequeDeposits using its ChequeID
    
    sql = ""
    
    sql = "delete ChequeDeposits where ChequeNo='" & strChequeNo & "'"
    
    myclass.Delete sql
    
    
    '// teller's amount
    
    Set rs = Nothing
    
    sql = ""
    
    sql = "select * from CustomerBalance where ChequeNo='" & strChequeNo & "'"
    
    Set rs = CreateObject("adodb.recordset")
    
    Set cn = CreateObject("adodb.connection")
    
   cn.Open Provider, "atm", "atm"
    
    rs.Open sql, cn
    
    If rs.EOF Then Exit Function
    
    If rs!transDescription = "Cheque Deposit" Then '// Increase tellers money
    
        sql = ""
        
        sql = "insert into [Teller Transactions]"
        
        sql = sql & "(CubicleNumber,TellerName,Deposits,AccountNumber,TransactionDate)"
        
        sql = sql & "select '" & CubicleNumber & "','" & tellerName & "','" & strAmount & "','" & rs!ACCNO & "','" & Date & "'"
        
        myclass.save sql
    
Else ' // Withdrawal

        sql = ""
        
        sql = "insert into [Teller Transactions]"
        
        sql = sql & "(CubicleNumber,TellerName,Withdrawals,AccountNumber,TransactionDate)"
        
        sql = sql & "select '" & CubicleNumber & "','" & tellerName & "','" & strAmount & "','" & rs!ACCNO & "','" & Date & "'"
        
        myclass.save sql
    
End If
    
    
    myclass.CloseCon
    
     Set rs = Nothing
    
    Set myclass = Nothing

End Function
Public Function GetCubieNumber() As String

    GetCubieNumber = GetSetting(App.title, "Teller", "Cubie Number")

End Function
Public Function GetTellerName() As String

     GetTellerName = GetSetting(App.title, "Teller", "Name")

End Function
Public Function WithdrawCash(strSQL As String, Optional strAccNumber As String, Optional DepositAmount As String)
        
        Dim diff As Integer
        
        Dim withdrawaldate As String
        
        accountnumber = strAccNumber
        
        amount = DepositAmount
        
        sql = strSQL
        
        '// Update Customer Balance
        
        Set myclass = New cdbase
        
        myclass.save sql
        
        ReduceTellerMoney

End Function
Private Sub ReduceTellerMoney()

        Set rs = Nothing
        
        Set rs = CreateObject("adodb.recordset")
        
        sql = ""
        
        sql = "select requisitionamount from TellersTreasuryTransactions where Transactiondate='" & Date & "' and tellername='" & tellerName & "'"
        
        
        Set cn = CreateObject("adodb.recordset")
                
        Set myclass = New cdbase
                
        Provider = myclass.OpenCon
        
       cn.Open Provider, "atm", "atm"
        
        rs.Open sql, cn
        
        If rs.EOF Then MsgBox "Teller Has no money!!??", vbExclamation, "FOSA": Exit Sub '// Teller has no money!! Very Stange
        
        RAmnt = rs!requisitionamount
        
        RAmnt = CCur(RAmnt) - CCur(amount)
        
        sql = "Delete from TellersTreasuryTransactions where transactiondate='" & Date & "'and tellername='" & tellerName & "'"
        
        myclass.Delete sql
        
        sql = ""
        
        sql = "insert into TellersTreasuryTransactions(CubicleNumber,tellername,transactiondate,requisitionamount)"
        sql = sql & "select'" & CubicleNumber & "','" & tellerName & "','" & Date & "','" & RAmnt & "'"
        
        myclass.save sql
        
        Set myclass = Nothing


End Sub
Public Function DepositCash(strSQL As String, Optional strAccNumber As String, Optional DepositAmount As String, _
Optional TransactionType As String)
        
        accountnumber = strAccNumber

        amount = DepositAmount

        If TransactionType = "Reduce" Then
        
            ReduceTellerMoney
        
        Else

            AddTellerMoney
        
        End If
        
End Function
Private Sub AddTellerMoney()
        
        Set myclass = New cdbase
        
        Set rs = CreateObject("adodb.recordset")
        
        sql = ""
        
        sql = "select requisitionamount from TellersTreasuryTransactions where TransactionDate='" & Date & "' and tellername='" & tellerName & "'"
        
        Set cn = CreateObject("adodb.recordset")
        
        Provider = myclass.OpenCon
        
       cn.Open Provider, "atm", "atm"
        
        rs.Open sql, cn
        
        If rs.EOF Then MsgBox "Teller Has no money!!??", vbExclamation, "FOSA": RAmnt = 0: GoTo 10  '// Teller has no money!! Very Stange"
        
        RAmnt = rs.Fields("requisitionamount")
        
10
        If amount = "" Then amount = 0

        RAmnt = CCur(RAmnt) + CCur(amount)
        
        sql = ""
        
        sql = "Delete from TellersTreasuryTransactions where transactiondate='" & Date & "' and tellername='" & tellerName & "'"
        
        myclass.Delete sql
        
        sql = ""
        
        sql = "insert into TellersTreasuryTransactions(CubicleNumber,tellername,transactiondate,requisitionamount)"
        sql = sql & "select'" & CubicleNumber & "','" & tellerName & "','" & Date & "','" & RAmnt & "'"
                
        myclass.save sql
        
        Set myclass = Nothing


End Sub
        
Public Sub ChequesToClear()

    Dim MyRs As Object
    
    Dim MyCn As Object
    
CheckAgain:
        
    Set myclass = New cdbase
    
    If Provider = "" Then Provider = myclass.OpenCon
    
    Set cn = CreateObject("adodb.connection")
    
   cn.Open Provider, "atm", "atm"
    
    Set rs = CreateObject("adodb.recordset")
    
    sql = ""
    
    sql = "select * from ChequeDeposits"
    
    Set MyCn = CreateObject("adodb.connection")
    
    MyCn.Open Provider
    
    Set MyRs = CreateObject("adodb.recordset")
    
    MyRs.Open sql, MyCn
    
    If MyRs.EOF Then Exit Sub '// There are no cheques waiting to be cleared
    
    Dim DepositDate As String
     
     Dim diff As Integer

    While Not MyRs.EOF
    
        DepositDate = MyRs.Fields("DepositDate")
    
        diff = CCur(Day(Date) - Day(DepositDate))
        
        If diff < 0 Then diff = diff * -1
        
        Dim diffMonths As Integer
        
        diffMonths = CCur(month(Date) - month(DepositDate))
        
        If diffMonths < 0 Then diffMonths = diffMonths * 1
        
        If diff >= MyRs.Fields("clearanceDuration") Or diffMonths > 0 Then      '// Clearing days are over
        
        'get amount to be cleared and the cheque number
        
        Dim MyRecordSet As Object
        
        sql = ""
        
        sql = "Select * from ChequeDeposits where AccountNumber='" & MyRs.Fields("AccountNumber") & "'"
        
        Set cn = CreateObject("adodb.connection")
        
       cn.Open Provider, "atm", "atm"
        
        Set MyRecordSet = CreateObject("adodb.Recordset")
        
        MyRecordSet.Open sql, cn
        
        If MyRecordSet.EOF Then Exit Sub 'The cheque has no money! Madness since it can't happen
        
       ClearCheques MyRecordSet!amount, MyRs!chequeno
        
     End If
         
    MyRs.MoveNext
    
    Wend
    
   

End Sub


Public Function DepositCheque(strAccountNumber As String, strDepositAmount As String, _
Optional strChequeType As String, Optional strClearanceDuration As String, Optional strChequeNumber As String)

    accountnumber = strAccountNumber
    
    amount = strDepositAmount

    Dim cTrans As Business
    
    Set myclass = New cdbase
    
    Set cTrans = New Business
    
    sql = ""
    
    sql = "insert into chequeDeposits(AccountNumber,DepositDate,Amount,ChequeType,ClearanceDuration,ChequeNo)"
    sql = sql & "select'" & accountnumber & "','" & Date & "','" & amount & "','" & strChequeType & "','" & _
    strClearanceDuration & "','" & strChequeNumber & "'"
    
    myclass.save sql
        
    AddTellerMoney
    
End Function

Private Sub Class_Initialize()
'Dim reg As regMan
'reg.setDefKey = "localmachine"
'reg.setParentNode = "SOFTWARE"
'TellerName = reg.myReadkey("fosa", "tellername")
    'SaveSetting app.Title, "Teller", "Name", rs!username
    tellerName = GetSetting(App.title, "Teller", "Name")

    CubicleNumber = GetSetting(App.title, "Teller", "Cubie Number")

    If CubicleNumber = "" Then

Retry:
        CubicleNumber = InputBox("Enter the Cubicle Number")

        If CubicleNumber = "" Then GoTo Retry

        SaveSetting App.title, "Teller", "Cubie Number", CubicleNumber

        CubicleNumber = GetSetting(App.title, "Teller", "Cubie Number")

    End If

End Sub


