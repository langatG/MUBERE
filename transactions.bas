Attribute VB_Name = "transactions"
Option Explicit
Public Sub power_to_edit(strUserName As String, edit1 As Boolean)
Dim temprs As Object
   Dim myclass As Object
    
    Set myclass = New cdbase
    
    Set cn = CreateObject("adodb.connection")
    Provider = myclass.OpenCon
   cn.Open Provider, "atm", "atm"
     
       sql = ""
       sql = "select * from useraccounts where UserLoginIDs='" & strUserName & "' and superuser=1"
       Set temprs = CreateObject("adodb.recordset")
     temprs.Open sql, cn
     If temprs.EOF Then
     edit1 = False
     Else
     edit1 = True
     End If

End Sub
Public Sub rebuild_accno1(ACCNO As String)
'On Error Resume Next
'//to rebult all the balances that has not been maintained well
Dim cn As Connection
Dim cn2 As Connection
Dim cn3 As Connection
Dim rs3 As Recordset
Dim rs As Recordset
Dim rs2 As Recordset
Dim sql As String
Dim I As Long

Set cn = New Connection
 Dim rsun1 As Object
Dim uncleared1 As Currency
Dim actual1 As Currency
Dim COMM As Currency
Set rs = New Recordset
Set rs2 = New Recordset
Set rs3 = New Recordset
cn.Open SelectedDsn, "atm"

sql = "SELECT distinct count(accno) From CustomerBalance WHERE AccNO = '" & _
ACCNO & "' and TransDescription <> 'Cheque Deposit(uncleared)' and transdescription <>'Cheque Dep(uncleared)' "
' ORDER BY CustomerBalanceid"
rs2.Open sql, cn
If rs2.EOF Then
 
  MsgBox "No records for rebuilding", vbExclamation
  Exit Sub
Else
  Dim AvailableBal As Currency
  Dim description As String
  Dim amount As Currency
  Dim Total_Records As Long
  Total_Records = rs2.Fields(0)
  rs2.Close
  
  sql = "SELECT distinct accno From CustomerBalance WHERE AccNO = '" & _
  ACCNO & "' and TransDescription <> 'Cheque Deposit(uncleared)' and transdescription <>'Cheque Dep(uncleared)'"  'ORDER BY transdate asc"
  rs2.Open sql, cn
  
  While Not rs2.EOF
      '//loop through all the selected members
      sql = "select customerbalanceid,Amount,AvailableBalance,transType,TransDescription," & _
      "TransDate, Commission, ChequeNo from CustomerBalance WHERE AccNO='" & _
      rs2.Fields("accno") & "' ORDER BY transdate,customerbalanceid asc"
      'and TransDescription <> 'Cheque Deposit(uncleared)' and  (TransDescription <> 'Cheque Dep(uncleared)')ORDER BY transdate asc"
      rs.Open sql, cn
     
      
      While Not rs.EOF
        I = I + 1
        If AvailableBal = 0 Then
          '//means this is the first balance
           If Not IsNull(rs.Fields("AvailableBalance")) Then
          
               If rs.Fields("transdescription") <> "2002 Balance B/F." And rs.Fields("transtype") = "DR" Then
           
               AvailableBal = rs.Fields("Amount")
           
               AvailableBal = -AvailableBal
               If AvailableBal = 0 And I > 1 Then
               AvailableBal = 0
               GoTo hell
               End If
               GoTo saddam
             ElseIf rs.Fields("transdescription") <> "2002 Balance B/F." And rs.Fields("transdescription") = "Cheque Deposit(uncleared)" Or rs.Fields("transdescription") = "Cheque Dep(uncleared)" Then
               GoTo saddam
                            
            ElseIf rs.Fields("transdescription") <> "2002 Balance B/F." Then
               If AvailableBal = 0 And I > 1 Then
               AvailableBal = 0
               GoTo hell
               End If
           AvailableBal = rs.Fields("Amount")
           actual1 = AvailableBal
           GoTo saddam
           Else
             AvailableBal = rs.Fields("AvailableBalance")
             End If
          End If
          rs.MoveNext
        End If
         
        '// check the precision of this (kisumu)
       If rs.EOF Then GoTo KISUMU
       If Not IsNull(rs.Fields("transdescription")) Then description = rs.Fields("transdescription")
        If rs.EOF Then
          rs.MoveFirst
          GoTo saddam
        End If
hell:
        'amount = CCur(rs.Fields("Amount")) + CCur(rs.Fields("Commission"))
          If UCase(Trim(rs.Fields("TransType"))) = "DR" Then
            amount = CCur(rs.Fields("Amount"))
            AvailableBal = AvailableBal - amount
          Else
              If description = "Cheque Deposit(uncleared)" Or description = "Cheque Dep(uncleared)" Then
              
               AvailableBal = AvailableBal
              Else
            amount = CCur(rs.Fields("Amount"))
            AvailableBal = AvailableBal + amount
            End If
          End If
           
            sql = "SELECT     Amount AS unclearedamnt FROM         CustomerBalance  WHERE     (AccNO = '" & rs2.Fields("accno") & "') AND (TransDescription LIKE 'Cheque Dep(uncleared)%') and customerbalanceid='" & rs.Fields("customerbalanceid") & "'"
            Set rsun1 = New ADODB.Recordset
            rsun1.Open sql, cn
            If Not rsun1.EOF Then
            If Not IsNull(rsun1.Fields("unclearedamnt")) Then uncleared1 = rsun1.Fields("unclearedamnt") Else uncleared1 = 0
            uncleared1 = Format(uncleared1, "###,###,###.00")
            End If
            
             actual1 = uncleared1 + AvailableBal
          If actual1 > 0 Then actual1 = uncleared1
        If COMM > 0 Then
saddam1:
      sql = "update customerbalance set availablebalance=" & AvailableBal & " ,commission=" & COMM & ", actualbalance=" & actual1 & "where  customerbalanceid =" & rs.Fields("customerbalanceid") & ""
          
          Set cn3 = New Connection
          cn3.Open SelectedDsn, "atm"
          cn3.Execute sql
          cn3.Close
          COMM = 0
          Set cn3 = Nothing
          End If
saddam:

          sql = "update customerbalance set availablebalance=" & AvailableBal & ",actualbalance=" & actual1 & " where  customerbalanceid =" & rs.Fields("customerbalanceid") & ""
          
          Set cn2 = New Connection
          'cn2.Open Selecteddsn,"atm"
          'cn2.Execute sql
          'cn2.Close
          Set cn2 = Nothing
          '// get the actual balance
           Dim rsun As Object
           Dim uncleared As Currency
           Dim actual As Currency
            sql = "SELECT     SUM(Amount) AS unclearedamnt FROM         CustomerBalance  WHERE     (AccNO = '" & rs2.Fields("accno") & "') AND (TransDescription LIKE 'Cheque Dep(uncleared)%')"
            Set rsun = New ADODB.Recordset
            rsun.Open sql, cn
            If Not rsun.EOF Then
            If Not IsNull(rsun.Fields("unclearedamnt")) Then uncleared = rsun.Fields("unclearedamnt") Else uncleared = 0
            uncleared = Format(uncleared, "###,###,###.00")
            End If
             actual = uncleared + AvailableBal


      sql = "update cub set availablebalance=" & AvailableBal & ",Active=1,actualbalance=" & actual & " where accno='" & rs2.Fields("accno") & "'"
      Set cn2 = New Connection
      cn2.Open SelectedDsn, "atm"
      cn2.Execute sql
      cn2.Close

          'Me.Caption = "Rebuilder    Processing " & i & " of a total " & Total_Records & " records"
        rs.MoveNext
KISUMU:
      Wend
      rs.Close
      rs2.MoveNext
      AvailableBal = 0
  Wend
End If



End Sub
Public Sub rebuild_accno333(ACCNO As String)
On Error Resume Next
'//to rebult all the balances that has not been maintained well

Dim cn2 As Connection
Dim cn3 As Connection
Dim rs3 As Recordset
Dim rs As Recordset
Dim rs2 As Recordset
Dim sql As String
Dim I As Long

Set cn = New Connection
 Dim rsun1 As Object
Dim uncleared1 As Currency
Dim actual1 As Currency
Dim COMM As Currency
Set rs = New Recordset
Set rs2 = New Recordset
Set rs3 = New Recordset
cn.Open SelectedDsn, "atm"

sql = "SELECT distinct count(accno) From CustomerBalance WHERE AccNO = '" & _
ACCNO & "' and TransDescription <> 'Cheque Deposit(uncleared)' and transdescription <>'Cheque Dep(uncleared)' "
' ORDER BY CustomerBalanceid"
rs2.Open sql, cn
If rs2.EOF Then
 
  MsgBox "No records for rebuilding", vbExclamation
  Exit Sub
Else
  Dim AvailableBal As Currency
  Dim description As String
  Dim amount As Currency
  Dim Total_Records As Long
  Total_Records = rs2.Fields(0)
  rs2.Close
  
  sql = "SELECT distinct accno From CustomerBalance WHERE AccNO = '" & _
  ACCNO & "' and TransDescription <> 'Cheque Deposit(uncleared)' and transdescription <>'Cheque Dep(uncleared)'"  'ORDER BY transdate asc"
  rs2.Open sql, cn
  
  While Not rs2.EOF
      '//loop through all the selected members
      sql = "select customerbalanceid,Amount,AvailableBalance,transType,TransDescription," & _
      "TransDate, Commission, ChequeNo from CustomerBalance WHERE AccNO='" & _
      rs2.Fields("accno") & "' ORDER BY transdate,customerbalanceid asc"
      'and TransDescription <> 'Cheque Deposit(uncleared)' and  (TransDescription <> 'Cheque Dep(uncleared)')ORDER BY transdate asc"
      rs.Open sql, cn
     
      
      While Not rs.EOF
        I = I + 1
        If AvailableBal = 0 Then
          '//means this is the first balance
           If Not IsNull(rs.Fields("AvailableBalance")) Then
          
               If rs.Fields("transdescription") <> "2002 Balance B/F." And rs.Fields("transtype") = "DR" Then
           
               AvailableBal = rs.Fields("Amount") - CCur(rs.Fields("Commission"))
                actual1 = rs.Fields("Amount")
           
               AvailableBal = -AvailableBal
               If AvailableBal = 0 And I > 1 Then
               AvailableBal = 0
               GoTo hell
               End If
               GoTo saddam
             ElseIf rs.Fields("transdescription") <> "2002 Balance B/F." And rs.Fields("transdescription") = "Cheque Deposit(uncleared)" Or rs.Fields("transdescription") = "Cheque Dep(uncleared)" Then
               GoTo saddam
            ElseIf rs.Fields("transdescription") <> "2002 Balance B/F." Then
               If AvailableBal = 0 And I > 1 Then
               AvailableBal = 0
               GoTo hell
               End If
           AvailableBal = rs.Fields("Amount") - CCur(rs.Fields("Commission"))
           actual1 = AvailableBal
           GoTo saddam
           Else
             AvailableBal = rs.Fields("AvailableBalance")
             actual1 = rs.Fields("ACTUALBalance")
             End If
          End If
          rs.MoveNext
        End If
         
        '// check the precision of this (kisumu)
       If rs.EOF Then GoTo KISUMU
       If Not IsNull(rs.Fields("transdescription")) Then description = rs.Fields("transdescription")
        If rs.EOF Then
          rs.MoveFirst
          GoTo saddam
        End If
hell:
        'amount = CCur(rs.Fields("Amount")) + CCur(rs.Fields("Commission"))
          If UCase(Trim(rs.Fields("TransType"))) = "DR" Then
            amount = CCur(rs.Fields("Amount")) + CCur(rs.Fields("Commission"))
            AvailableBal = AvailableBal - amount
            actual1 = actual1 - amount
          Else
              If description = "Cheque Deposit(uncleared)" Or description = "Cheque Dep(uncleared)" Then
              sql = "SELECT     Amount AS unclearedamnt FROM         CustomerBalance  WHERE     (AccNO = '" & rs2.Fields("accno") & "') AND (TransDescription LIKE 'Cheque Dep(uncleared)%') and customerbalanceid='" & rs.Fields("customerbalanceid") & "'"
            Set rsun1 = New ADODB.Recordset
            rsun1.Open sql, cn
            'uncleared1 = 0
            If Not rsun1.EOF Then
            If Not IsNull(rsun1.Fields("unclearedamnt")) Then uncleared1 = rsun1.Fields("unclearedamnt") Else uncleared1 = 0
            uncleared1 = Format(uncleared1, "###,###,###.00")
            actual1 = actual1 + uncleared1
            End If
               AvailableBal = AvailableBal
               'actual1 = actual1 + amount
              Else
            amount = CCur(rs.Fields("Amount")) - CCur(rs.Fields("Commission"))
            AvailableBal = AvailableBal + amount
            actual1 = actual1 + amount
            End If
          End If
           
            
            '// CHECK THE STATUS OF THE ACTUAL BALANCE
           
            'uncleared1 = 0
            'If Not rsun1.EOF Then
            'If Not IsNull(rsun1.Fields("unclearedamnt")) Then uncleared1 = rsun1.Fields("unclearedamnt") Else uncleared1 = 0
            'uncleared1 = Format(uncleared1, "###,###,###.00")
            'End If
            
             'actual1 = uncleared1 + AvailableBal
             
          
        If COMM > 0 Then
saddam1:
      sql = "update customerbalance set availablebalance=" & AvailableBal & " ,commission=" & COMM & ", actualbalance=" & actual1 & "where  customerbalanceid =" & rs.Fields("customerbalanceid") & ""
          
          Set cn3 = New Connection
          cn3.Open SelectedDsn, "atm"
          cn3.Execute sql
          cn3.Close
          COMM = 0
          Set cn3 = Nothing
          End If
saddam:

          sql = "update customerbalance set availablebalance=" & AvailableBal & ",actualbalance=" & actual1 & " where  customerbalanceid =" & rs.Fields("customerbalanceid") & ""
          
          Set cn2 = New Connection
          cn2.Open SelectedDsn, "atm"
          cn2.Execute sql
          cn2.Close
          Set cn2 = Nothing
          '// get the actual balance
           Dim rsun As Object
           Dim uncleared As Currency
           Dim actual As Currency
            sql = "SELECT     SUM(Amount) AS unclearedamnt FROM         CustomerBalance  WHERE     (AccNO = '" & rs2.Fields("accno") & "') AND (TransDescription LIKE 'Cheque Dep(uncleared)%')"
            Set rsun = New ADODB.Recordset
            rsun.Open sql, cn
            If Not rsun.EOF Then
            If Not IsNull(rsun.Fields("unclearedamnt")) Then uncleared = rsun.Fields("unclearedamnt") Else uncleared = 0
            uncleared = Format(uncleared, "###,###,###.00")
            End If
             actual = uncleared + AvailableBal


      sql = "update cub set availablebalance=" & AvailableBal & ",Active=1,actualbalance=" & actual & " where accno='" & rs2.Fields("accno") & "'"
      Set cn2 = New Connection
      cn2.Open SelectedDsn, "atm"
      cn2.Execute sql
      cn2.Close

          'Me.Caption = "Rebuilder    Processing " & i & " of a total " & Total_Records & " records"
        rs.MoveNext
KISUMU:
      Wend
      rs.Close
      rs2.MoveNext
      AvailableBal = 0
      actual = 0
      actual1 = 0
  Wend
End If



 'MsgBox "Processing Complete"
Exit Sub
ErrHandler:

MsgBox err.description
End Sub
Public Sub rebuild_accno(ACCNO As String)
'On Error Resume Next
'//to rebult all the balances that has not been maintained well


Dim cn2 As Connection
Dim cn3 As Connection
Dim rs3 As Recordset
Dim rs As Recordset
Dim rs2 As Recordset
Dim sql As String
Dim I As Long

Set cn = New Connection
 Dim rsun1 As Object
Dim uncleared1 As Currency
Dim actual1 As Currency
Dim COMM As Currency
Set rs = New Recordset
Set rs2 = New Recordset
Set rs3 = New Recordset
cn.Open SelectedDsn, "atm"

sql = "SELECT distinct count(accno) From CustomerBalance WHERE AccNO = '" & _
ACCNO & "' and TransDescription <> 'Cheque Deposit(uncleared)' and transdescription <>'Cheque Dep(uncleared)' "
' ORDER BY CustomerBalanceid"
rs2.Open sql, cn
If rs2.EOF Then
 
  MsgBox "No records for rebuilding", vbExclamation
  Exit Sub
Else
  Dim AvailableBal As Currency
  Dim description As String
  Dim amount As Currency
  Dim Total_Records As Long
  Total_Records = rs2.Fields(0)
  rs2.Close
  
  sql = "SELECT distinct accno From CustomerBalance WHERE AccNO = '" & _
  ACCNO & "' and TransDescription <> 'Cheque Deposit(uncleared)' and transdescription <>'Cheque Dep(uncleared)'"  'ORDER BY transdate asc"
  rs2.Open sql, cn
  
  While Not rs2.EOF
      '//loop through all the selected members
      sql = "select customerbalanceid,Amount,AvailableBalance,transType,TransDescription," & _
      "TransDate, Commission, ChequeNo from CustomerBalance WHERE AccNO='" & _
      rs2.Fields("accno") & "' ORDER BY transdate,customerbalanceid asc"
      'and TransDescription <> 'Cheque Deposit(uncleared)' and  (TransDescription <> 'Cheque Dep(uncleared)')ORDER BY transdate asc"
      rs.Open sql, cn
     
      
      While Not rs.EOF
        I = I + 1
        If AvailableBal = 0 Then
          '//means this is the first balance
           If Not IsNull(rs.Fields("AvailableBalance")) Then
          
               If rs.Fields("transdescription") <> "2002 Balance B/F." And rs.Fields("transtype") = "DR" Then
           
               AvailableBal = rs.Fields("Amount") - CCur(rs.Fields("Commission"))
                actual1 = rs.Fields("Amount")
           
               AvailableBal = AvailableBal 'oluoch alisema
               If AvailableBal = 0 And I > 1 Then
               AvailableBal = 0
               GoTo hell
               End If
               GoTo saddam
             ElseIf rs.Fields("transdescription") <> "2002 Balance B/F." And rs.Fields("transdescription") = "Cheque Deposit(uncleared)" Or rs.Fields("transdescription") = "Cheque Dep(uncleared)" Then
               GoTo saddam
            ElseIf rs.Fields("transdescription") <> "2002 Balance B/F." Then
               If AvailableBal = 0 And I > 1 Then
               AvailableBal = 0
               GoTo hell
               End If
           AvailableBal = -rs.Fields("Amount") - CCur(rs.Fields("Commission"))
           actual1 = AvailableBal 'eva sleeping
           GoTo saddam
           Else
             AvailableBal = rs.Fields("AvailableBalance")
             actual1 = rs.Fields("ACTUALBalance")
             End If
          End If
          rs.MoveNext
        End If
         
        '// check the precision of this (kisumu)
       If rs.EOF Then GoTo KISUMU
       If Not IsNull(rs.Fields("transdescription")) Then description = rs.Fields("transdescription")
        If rs.EOF Then
          rs.MoveFirst
          GoTo saddam
        End If
hell:
        'amount = CCur(rs.Fields("Amount")) + CCur(rs.Fields("Commission"))
          If UCase(Trim(rs.Fields("TransType"))) = "DR" Then
            amount = CCur(rs.Fields("Amount")) + CCur(rs.Fields("Commission"))
            AvailableBal = AvailableBal + amount
            actual1 = actual1 + amount ' kibet doubted
          Else
              If description = "Cheque Deposit(uncleared)" Or description = "Cheque Dep(uncleared)" Then
              sql = "SELECT     Amount AS unclearedamnt FROM         CustomerBalance  WHERE     (AccNO = '" & rs2.Fields("accno") & "') AND (TransDescription LIKE 'Cheque Dep(uncleared)%') and customerbalanceid='" & rs.Fields("customerbalanceid") & "'"
            Set rsun1 = New ADODB.Recordset
            rsun1.Open sql, cn
            'uncleared1 = 0
            If Not rsun1.EOF Then
            If Not IsNull(rsun1.Fields("unclearedamnt")) Then uncleared1 = rsun1.Fields("unclearedamnt") Else uncleared1 = 0
            uncleared1 = Format(uncleared1, "###,###,###.00")
            actual1 = actual1 + uncleared1
            End If
               AvailableBal = AvailableBal
               'actual1 = actual1 + amount
              Else
            amount = CCur(rs.Fields("Amount")) - CCur(rs.Fields("Commission"))
            AvailableBal = AvailableBal - amount
            actual1 = actual1 + amount
            End If
          End If
           
            
            '// CHECK THE STATUS OF THE ACTUAL BALANCE
           
            'uncleared1 = 0
            'If Not rsun1.EOF Then
            'If Not IsNull(rsun1.Fields("unclearedamnt")) Then uncleared1 = rsun1.Fields("unclearedamnt") Else uncleared1 = 0
            'uncleared1 = Format(uncleared1, "###,###,###.00")
            'End If
            
             'actual1 = uncleared1 + AvailableBal
             
          
        If COMM > 0 Then
saddam1:
      sql = "update customerbalance set availablebalance=" & AvailableBal & " ,commission=" & COMM & ", actualbalance=" & actual1 & "where  customerbalanceid =" & rs.Fields("customerbalanceid") & ""
          
          Set cn3 = New Connection
          cn3.Open SelectedDsn, "atm"
          cn3.Execute sql
          cn3.Close
          COMM = 0
          Set cn3 = Nothing
          End If
saddam:
sql = ""
          sql = "update customerbalance set availablebalance=" & AvailableBal & ",actualbalance=" & actual1 & " where  customerbalanceid =" & rs.Fields("customerbalanceid") & ""
          
          Set cn2 = New ADODB.Connection
          cn2.Open SelectedDsn, "atm"
          'cn2.Execute sql
          cn2.Close
          Set cn2 = Nothing
          '// get the actual balance
           Dim rsun As Object
           Dim uncleared As Currency
           Dim actual As Currency
            sql = "SELECT     SUM(Amount) AS unclearedamnt FROM         CustomerBalance  WHERE     (AccNO = '" & rs2.Fields("accno") & "') AND (TransDescription LIKE 'Cheque Dep(uncleared)%')"
            Set rsun = New ADODB.Recordset
            rsun.Open sql, cn
            If Not rsun.EOF Then
            If Not IsNull(rsun.Fields("unclearedamnt")) Then uncleared = rsun.Fields("unclearedamnt") Else uncleared = 0
            uncleared = Format(uncleared, "###,###,###.00")
            End If
             actual = uncleared - AvailableBal 'henry thinking


      sql = "update cub set availablebalance=" & AvailableBal & ",Active=1,actualbalance=" & actual & " where accno='" & rs2.Fields("accno") & "'"
      Set cn2 = New Connection
      cn2.Open SelectedDsn, "atm"
      'cn2.Execute sql
      cn2.Close

          frmmembertransactions.Caption = "Rebuilder    Processing " & I & " of a total " & Total_Records & " records"
        rs.MoveNext
KISUMU:
      Wend
      rs.Close
      rs2.MoveNext
      AvailableBal = 0
      actual = 0
      actual1 = 0
  Wend
End If


frmmembertransactions.Caption = "Transactions"
 MsgBox "Processing Complete"
 
Exit Sub
ErrHandler:

MsgBox err.description
End Sub
'Function Get_Report_Path(report_path) As String
'    On Error Resume Next
'    Dim myclass As cdbase
'    Set myclass = New cdbase
''Set tempRs = CreateObject("adodb.recordset")
'Set cn = CreateObject("adodb.connection")
'provider = "DSN=PS"
'cn.Open provider
'    Dim rst As Recordset
'    Set rst = New Recordset
'    rst.Open "select * from reportpath", cn, adOpenStatic, adLockOptimistic
'    If rst.EOF = False Then
'        report_path = rst.Fields("reportpath")
'    End If
'    Get_Report_Path = report_path
'    rst.Close
'    Set cn = Nothing
'End Function
Function cub_balance(Txtaccno As String) As Variant
Dim returnfield
Dim X
Dim Provider As String

Dim balance As Currency
Dim myclass As cdbase
Set myclass = New cdbase
Set cn = CreateObject("adodb.connection")
Provider = myclass.OpenCon
cn.Open Provider, "atm", "atm"
Set myclass = New cdbase
Set cn = CreateObject("adodb.connection")
Provider = myclass.OpenCon
cn.Open Provider, "atm", "atm"
Set myclass = New cdbase
Set rs = CreateObject("adodb.recordset")
sql = ""
sql = "select * from cub where accno='" & Txtaccno & "'"
rs.Open sql, cn
rs.Requery

    returnfield = "availablebalance"
    Set rs = CreateObject("ADODB.Recordset")
    cn.Customer_balance Txtaccno, rs
    If rs.EOF = True Then Exit Function
    If rs("accno") = Txtaccno Then
        X = rs(returnfield)
    End If
    
    Set cn = Nothing
    Set rs = Nothing
   If Not IsNull(X) Then cub_balance = X
   Exit Function
10:
err.Raise err.number, "Balance not correct", err.description
End Function
Function get_rate()
Dim returnfield
Dim X
Dim Provider As String

Dim balance As Currency
Dim myclass As cdbase
Set myclass = New cdbase
Set cn = CreateObject("adodb.connection")
Provider = myclass.OpenCon
cn.Open Provider, "atm", "atm"

Set rs = CreateObject("adodb.recordset")
sql = ""
sql = "select * from ratesofinterest where pkid=2"
rs.Open sql, cn
rs.Requery
    returnfield = "intonoverdrafts"
    Set rs = CreateObject("ADODB.Recordset")
    cn.get_rate rs
    If rs.EOF = True Then Exit Function
    If rs("intonoverdrafts") Then
        X = rs(returnfield)
    End If
    
    Set cn = Nothing
    Set rs = Nothing
   If Not IsNull(X) Then get_rate = X
   Exit Function
10:
err.Raise err.number, "Balance not correct", err.description
End Function
Function To_Nearest_Lower_Number(Numb, ApproxToDigits) As Currency
    Dim X, Subt
    'take off right digits
    Numb = CDbl(Truncate_Number(Numb))
    Subt = Right(Numb, ApproxToDigits)
    X = Numb - Subt
    To_Nearest_Lower_Number = X
End Function
Function Round_Of_Two_Decimal(Numb) As Currency
    On Error Resume Next
    Dim Product, X
    If IsNull(Numb) Then Exit Function
    Numb = CCur(Numb)
    Product = Numb * 20
    Product = CCur(Format(Product, "#,##0"))
    X = CCur(Format(Product / 20, "#,##0.00"))
    Round_Of_Two_Decimal = X
    Exit Function
10:    MsgBox err.description
End Function
Function tens_hundreds_into_words(number As Variant) As String
    On Error GoTo 10
    Dim X
    Select Case number
        Case 20 To 29
            X = tens_number_into_words(20)
            If number - 20 > 0 Then
                X = X & " " & tens_number_into_words(number - 20)
            End If
        Case 30 To 39
            X = tens_number_into_words(30)
            If number - 30 > 0 Then
                X = X & " " & tens_number_into_words(number - 30)
            End If
        Case 40 To 49
            X = tens_number_into_words(40)
            If number - 40 > 0 Then
                X = X & " " & tens_number_into_words(number - 40)
            End If
        Case 50 To 59
            X = tens_number_into_words(50)
            If number - 20 > 0 Then
                X = X & " " & tens_number_into_words(number - 50)
            End If
        Case 60 To 69
            X = tens_number_into_words(60)
            If number - 60 > 0 Then
                X = X & " " & tens_number_into_words(number - 60)
            End If
        Case 70 To 79
            X = tens_number_into_words(70)
            If number - 70 > 0 Then
                X = X & " " & tens_number_into_words(number - 70)
            End If
        Case 80 To 89
            X = tens_number_into_words(80)
            If number - 20 > 0 Then
                X = X & " " & tens_number_into_words(number - 80)
            End If
        Case 90 To 99
            X = tens_number_into_words(90)
            If number - 90 > 0 Then
                X = X & " " & tens_number_into_words(number - 90)
            End If
        Case Else
    End Select
    tens_hundreds_into_words = X
    Exit Function
10:    MsgBox err.description
     
End Function
Function Company_Name(Retfield) As Variant
    On Error GoTo 10
    Dim myclass As Object
    Dim rst As Recordset
  Set myclass = New cdbase
Set cn = CreateObject("adodb.connection")
Provider = myclass.OpenCon
cn.Open Provider, "atm", "atm"
    Set rst = New Recordset
    Set cn = New ADODB.Connection
   cn.Open Provider, "atm", "atm"
    sql = ""
    sql = "select CompanyName  from SYSPARAM"
    rst.Open sql, cn
    If Not rst.EOF Then
    Company_Name = rst(Retfield)
    End If
    rst.Close
    Exit Function
10:    MsgBox err.description
End Function
Public Function val_Date_Format(thisDate As String) As Boolean
    On Error GoTo 10
'    val_Date_Format = False
    
    Dim tmpDD As String
    Dim tmpMM As String
    Dim tmpYYYY As String
     
    tmpDD = Left(thisDate, 2)
    tmpMM = mid(thisDate, 4, 2)
    tmpYYYY = Right(thisDate, 4)
        
        
        If Format(tmpDD & "-" & tmpMM & "-" & tmpYYYY, "yyyymmdd") <> tmpYYYY & tmpMM & tmpDD Then
            val_Date_Format = False
            Exit Function
        Else
            If IsDate(thisDate) Then
                val_Date_Format = True
            Else
                val_Date_Format = False
                Exit Function
            End If
            
        End If
10:
    End Function
    Function Leading_Zero(Digits) As String
    If Digits = 1 Then Leading_Zero = "0"
    If Digits = 2 Then Leading_Zero = "00"
    If Digits = 3 Then Leading_Zero = "000"
    If Digits = 4 Then Leading_Zero = "0000"
    If Digits = 5 Then Leading_Zero = "00000"
    If Digits = 6 Then Leading_Zero = "000000"
    If Digits = 7 Then Leading_Zero = "0000000"
    If Digits = 8 Then Leading_Zero = "00000000"
    If Digits = 9 Then Leading_Zero = "000000000"
    If Digits = 10 Then Leading_Zero = "0000000000"
End Function

Function Month_Begin_Date(mon As Integer, Yr As Integer) As Date
    Dim monEnd As Date, D As Integer
    monEnd = Month_End_Date(mon, Yr)
    D = Days_In_Month(mon, Yr)
    Month_Begin_Date = DateAdd("d", -D + 1, monEnd)
End Function
Function Days_In_Month(mon As Integer, Yr As Integer) As Integer
    On Error GoTo 10
    If mon = 1 Or mon = 3 Or mon = 5 Or mon = 7 Or mon = 8 Or mon = 10 Or mon = 12 Then
        Days_In_Month = 31
        Exit Function
    End If
    If mon = 2 Then
        If Yr Mod 4 = 0 Then
            Days_In_Month = 29
            Else
                Days_In_Month = 28
        End If
        Exit Function
    End If
    Days_In_Month = 30
    Exit Function
10:    MsgBox err.description
End Function
Function Month_End_Date(mon, Yr) As Date
    On Error GoTo 10
    Dim X
    X = "30/" & mon & "/" & Yr
    If mon = 1 Or mon = 3 Or mon = 5 Or mon = 7 Or mon = 8 Or mon = 10 Or mon = 12 Then
        X = "31/" & mon & "/" & Yr
        Month_End_Date = X
        Exit Function
    End If
    If mon = 2 Then
        X = "28/" & mon & "/" & Yr
        If Yr Mod 4 = 0 Then X = "29/" & mon & "/" & Yr
        Month_End_Date = X
        Exit Function
    End If
    Month_End_Date = X
    Exit Function
10:    MsgBox err.description
End Function
Function Month_In_Words(mon) As String
    On Error GoTo 10
    Dim X
    X = "December"
    If mon = 1 Then X = "January"
    If mon = 2 Then X = "February"
    If mon = 3 Then X = "March"
    If mon = 4 Then X = "April"
    If mon = 5 Then X = "May"
    If mon = 6 Then X = "June"
    If mon = 7 Then X = "July"
    If mon = 8 Then X = "August"
    If mon = 9 Then X = "September"
    If mon = 10 Then X = "October"
    If mon = 11 Then X = "November"
    Month_In_Words = X
    Exit Function
10:    MsgBox err.description
End Function
Function One_Decimal(Numb As Currency) As Currency
    One_Decimal = CCur(Format(Numb, "#,##0.0"))
End Function
Function Zero_Decimal(Numb As Currency) As Currency
    Zero_Decimal = CCur(Format(Numb, "#,##0"))
End Function
Function Truncate_Number(Numb) As Variant
    Dim n, D
    n = Format(Numb, "#,##0.000000")
    D = (Right(n, 6)) / 1000000
    n = CCur(n) - CCur(D)
    Truncate_Number = Format(n, "#,##0")
End Function
 
Function Calc_Time_Elapsed(ByVal Start) As String
    Dim Finish
    Dim Minutes
    Finish = Timer
    Minutes = Format((Finish - Start) / 60, "#,##0.00")
    Calc_Time_Elapsed = Minutes & " Minutes"
End Function
Function tens_number_into_words(number As Variant) As String
    On Error GoTo 10
    Dim X
    Select Case number
        Case 1
            X = "one"
        Case 2
            X = "two"
        Case 3
            X = "three"
        Case 4
            X = "four"
        Case 5
            X = "five"
        Case 6
            X = "six"
        Case 7
            X = "seven"
        Case 8
            X = "eight"
        Case 9
            X = "nine"
        Case 10
            X = "ten"
        Case 11
            X = "eleven"
        Case 12
            X = "twelve"
        Case 13
            X = "thirteen"
        Case 14
            X = "fourteen"
        Case 15
            X = "fifteen"
        Case 16
            X = "sixteen"
        Case 17
            X = "seventeen"
        Case 18
            X = "eighteen"
        Case 19
            X = "nineteen"
        Case 20
            X = "twenty"
        Case 30
            X = "thirty"
        Case 40
            X = "fourty"
        Case 50
            X = "fifty"
        Case 60
            X = "sixty"
        Case 70
            X = "seventy"
        Case 80
            X = "eighty"
        Case 90
            X = "ninety"
        Case Else
    End Select
    tens_number_into_words = X
    Exit Function
10:    MsgBox err.description
End Function

