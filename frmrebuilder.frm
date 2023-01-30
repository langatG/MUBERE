VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmrebuilder 
   BackColor       =   &H80000001&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rebuilder"
   ClientHeight    =   2205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5250
   Icon            =   "frmrebuilder.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Txtfrom 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   2520
      MaxLength       =   20
      TabIndex        =   4
      Top             =   120
      Width           =   2655
   End
   Begin VB.TextBox Txtto 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   2520
      MaxLength       =   20
      TabIndex        =   3
      Top             =   1080
      Width           =   2655
   End
   Begin VB.CommandButton cmdrebuild 
      Caption         =   "Rebuild"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   1680
      Width           =   1095
   End
   Begin MSComctlLib.ProgressBar Pb 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   4080
      Visible         =   0   'False
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FF80&
      Caption         =   "FROM ACCOUNT NUMBER"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FF80&
      Caption         =   "TO ACCOUNT NUMBER"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label lblprogress 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   4080
      Width           =   5055
   End
End
Attribute VB_Name = "frmrebuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdrebuild_Click()
On Error Resume Next
'//to rebult all the balances that has not been maintained well
If txtFrom = "" Or txtTo = "" Then
  MsgBox "Enter a range", vbExclamation
  Exit Sub
End If
Me.MousePointer = vbHourglass
Dim cn As Connection
Dim cn2 As Connection
Dim cn3 As Connection
Dim rs3 As Recordset
Dim rs As Recordset
Dim rs2 As Recordset
Dim sql As String
Dim I As Long

Set cn = New Connection

Dim COMM As Currency
Set rs = New Recordset
Set rs2 = New Recordset
Set rs3 = New Recordset
cn.Open SelectedDsn, "atm"

sql = "SELECT distinct count(accno) From CustomerBalance WHERE AccNO between '" & _
txtFrom & "'AND '" & txtTo & "' and TransDescription <> 'Cheque Deposit(uncleared)' and transdescription <>'Cheque Dep(uncleared)' "
' ORDER BY CustomerBalanceid"
rs2.Open sql, cn
If rs2.EOF Then
  Me.MousePointer = 0
  MsgBox "No records for rebuilding", vbExclamation
  Exit Sub
Else
  Dim AvailableBal As Currency
  Dim description As String
  Dim amount As Currency
  Dim Total_Records As Long
  Total_Records = rs2.Fields(0)
  rs2.Close
  
  sql = "SELECT distinct accno From CustomerBalance WHERE AccNO between '" & _
  txtFrom & "'AND '" & txtTo & "' and TransDescription <> 'Cheque Deposit(uncleared)' and transdescription <>'Cheque Dep(uncleared)'"  'ORDER BY transdate asc"
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
           
               AvailableBal = -AvailableBal
               If AvailableBal = 0 And I > 1 Then
               AvailableBal = 0
               GoTo hell
               End If
               GoTo saddam
             ElseIf rs.Fields("transdescription") <> "2002 Balance B/F." And rs.Fields("transdescription") = "Cheque Deposit(uncleared)" Or rs.Fields("transdescription") = "Cheque Dep(uncleared)" Then
               GoTo saddam
               
               
             ElseIf rs.Fields("transdescription") = "CASH WITHDRAWAL" Then
              
               If AvailableBal = 0 And I > 1 Then
               AvailableBal = rs.Fields("Amount") - CCur(rs.Fields("Commission"))
               GoTo saddam
               End If
               
              ElseIf rs.Fields("transdescription") = "CASH DEPOSIT" Then
            
               If AvailableBal = 0 And I > 1 Then
               AvailableBal = rs.Fields("Amount") - CCur(rs.Fields("Commission"))
               GoTo saddam
               End If
            
            ElseIf rs.Fields("transdescription") <> "2002 Balance B/F." And rs.Fields("transdescription") = "CASH DEPOSIT" Then
               
               AvailableBal = rs.Fields("Amount") - 500 - CCur(rs.Fields("Commission"))
               GoTo saddam
               
            
            ElseIf rs.Fields("transdescription") = "2002 Balance B/F." And (rs.Fields("TransType")) = "CR" Then
            
              AvailableBal = rs.Fields("Amount") - 500
              
              GoTo saddam
              
            ElseIf rs.Fields("transdescription") = "CHEQUE DEPOSIT" Then
               If AvailableBal = 0 And I > 1 Then
               AvailableBal = 0
               GoTo hell
               End If
               AvailableBal = rs.Fields("Amount") - CCur(rs.Fields("Commission")) - 500
               GoTo saddam
            
            ElseIf rs.Fields("transdescription") = "Cash deposit, (first)" Then
               If AvailableBal = 0 And I > 1 Then
                AvailableBal = 0
               GoTo hell
               End If
               If CCur(rs.Fields("commission")) = 600 Then
               AvailableBal = rs.Fields("Amount") - 100 - 500
               COMM = 100
               GoTo saddam1
               End If
               AvailableBal = rs.Fields("Amount") - CCur(rs.Fields("Commission")) - 500
               GoTo saddam
               
            ElseIf rs.Fields("transdescription") = "Salary " And rs.Fields("TransType") = "CR" Then
            If AvailableBal = 0 And I > 1 Then
               AvailableBal = 0
               GoTo hell
               End If
            AvailableBal = rs.Fields("Amount") - CCur(rs.Fields("Commission")) - 500
               GoTo saddam
            ElseIf rs.Fields("transdescription") = "Advance " And rs.Fields("TransType") = "CR" Then
            
            If AvailableBal = 0 And I > 1 Then
               AvailableBal = 0
               GoTo hell
               End If
            
            AvailableBal = rs.Fields("Amount") - CCur(rs.Fields("Commission")) - 500
               GoTo saddam
            ElseIf rs.Fields("transdescription") = "KTDA" And rs.Fields("TransType") = "CR" Then
             If AvailableBal = 0 And I > 1 Then
               AvailableBal = 0
               GoTo hell
               End If
             AvailableBal = rs.Fields("Amount") - CCur(rs.Fields("Commission")) - 500
               GoTo saddam
            ElseIf rs.Fields("transdescription") = "NLoan" And rs.Fields("TransType") = "CR" Then
              If AvailableBal = 0 And I > 1 Then
               AvailableBal = 0
               GoTo hell
               End If
            AvailableBal = rs.Fields("Amount") - CCur(rs.Fields("Commission")) - 500
               GoTo saddam
            ElseIf rs.Fields("transdescription") = "BROOKSIDE " And rs.Fields("TransType") = "CR" Then
             If AvailableBal = 0 And I > 1 Then
               AvailableBal = 0
               GoTo hell
               End If
            AvailableBal = rs.Fields("Amount") - CCur(rs.Fields("Commission")) - 500
               GoTo saddam
            ElseIf rs.Fields("transdescription") = "PYHRETHRUM" And rs.Fields("TransType") = "CR" Then
             If AvailableBal = 0 And I > 1 Then
               AvailableBal = 0
               GoTo hell
               End If
            AvailableBal = rs.Fields("Amount") - CCur(rs.Fields("Commission")) - 500
               GoTo saddam
            ElseIf rs.Fields("transdescription") = "FOSA" And rs.Fields("TransType") = "CR" Then
             If AvailableBal = 0 And I > 1 Then
               AvailableBal = 0
               GoTo hell
               End If
            AvailableBal = rs.Fields("Amount") - CCur(rs.Fields("Commission")) - 500
               GoTo saddam
            ElseIf (rs.Fields("transdescription") = "Kernal") And rs.Fields("TransType") = "CR" Then
             If AvailableBal = 0 And I > 1 Then
               AvailableBal = 0
               GoTo hell
               End If
            AvailableBal = rs.Fields("Amount") - CCur(rs.Fields("Commission")) - 500
               GoTo saddam
            ElseIf (rs.Fields("transdescription") = "SHARES REFUND") And rs.Fields("TransType") = "CR" Then
             If AvailableBal = 0 And I > 1 Then
               AvailableBal = 0
               GoTo hell
               End If
            AvailableBal = rs.Fields("Amount") - CCur(rs.Fields("Commission")) - 500
               GoTo saddam
            ElseIf (rs.Fields("transdescription") = "LOAN-DIVIDEND") And rs.Fields("TransType") = "CR" Then
             If AvailableBal = 0 And I > 1 Then
               AvailableBal = 0
               GoTo hell
               End If
            AvailableBal = rs.Fields("Amount") - CCur(rs.Fields("Commission")) - 500
               GoTo saddam
            ElseIf (rs.Fields("transdescription") = "DIVIDENDS") And rs.Fields("TransType") = "CR" Then
             If AvailableBal = 0 And I > 1 Then
               AvailableBal = 0
               GoTo hell
               End If
            AvailableBal = rs.Fields("Amount") - CCur(rs.Fields("Commission")) - 500
               GoTo saddam
            ElseIf (rs.Fields("transdescription") = "INTEREST REFUND") And rs.Fields("TransType") = "CR" Then
             If AvailableBal = 0 And I > 1 Then
               AvailableBal = 0
               GoTo hell
               End If
            AvailableBal = rs.Fields("Amount") - CCur(rs.Fields("Commission")) - 500
               GoTo saddam
            ElseIf (rs.Fields("transdescription") = "KBC") And rs.Fields("TransType") = "CR" Then
             If AvailableBal = 0 And I > 1 Then
               AvailableBal = 0
               GoTo hell
               End If
            AvailableBal = rs.Fields("Amount") - CCur(rs.Fields("Commission")) - 500
               GoTo saddam
            ElseIf (rs.Fields("transdescription")) = "LOAN REFUND" And rs.Fields("TransType") = "CR" Then
             If AvailableBal = 0 And I > 1 Then
               AvailableBal = 0
               GoTo hell
               End If
            AvailableBal = rs.Fields("Amount") - CCur(rs.Fields("Commission")) - 500
               GoTo saddam
            ElseIf rs.Fields("transdescription") = "Shares Refund -out " And rs.Fields("TransType") = "CR" Then
             If AvailableBal = 0 And I > 1 Then
               AvailableBal = 0
               GoTo hell
               End If
            AvailableBal = rs.Fields("Amount") - CCur(rs.Fields("Commission")) - 500
               GoTo saddam
            ElseIf rs.Fields("transdescription") = "Pension " And rs.Fields("TransType") = "CR" Then
             If AvailableBal = 0 And I > 1 Then
               AvailableBal = 0
               GoTo hell
               End If
            AvailableBal = rs.Fields("Amount") - CCur(rs.Fields("Commission")) - 500
               GoTo saddam
            ElseIf rs.Fields("transdescription") = "Shares refund" And rs.Fields("TransType") = "CR" Then
             If AvailableBal = 0 And I > 1 Then
               AvailableBal = 0
               GoTo hell
               End If
             AvailableBal = rs.Fields("Amount") - CCur(rs.Fields("Commission")) - 500
               GoTo saddam
            ElseIf rs.Fields("transdescription") = "Eloan" And rs.Fields("TransType") = "CR" Then
             If AvailableBal = 0 And I > 1 Then
               AvailableBal = 0
               GoTo hell
               End If
            AvailableBal = rs.Fields("Amount") - CCur(rs.Fields("Commission")) - 500
               GoTo saddam
            ElseIf rs.Fields("transdescription") = "STAFF ADVANCE" And rs.Fields("TransType") = "CR" Then
             If AvailableBal = 0 And I > 1 Then
               AvailableBal = 0
               GoTo hell
               End If
            AvailableBal = rs.Fields("Amount") - CCur(rs.Fields("Commission")) - 500
               GoTo saddam
            ElseIf rs.Fields("transdescription") = "STAFF LOAN" And rs.Fields("TransType") = "CR" Then
             If AvailableBal = 0 And I > 1 Then
               AvailableBal = 0
               GoTo hell
               End If
            AvailableBal = rs.Fields("Amount") - CCur(rs.Fields("Commission")) - 500
               GoTo saddam
            ElseIf rs.Fields("transdescription") = "STAFF SALARY " And rs.Fields("TransType") = "CR" Then
             If AvailableBal = 0 And I > 1 Then
               AvailableBal = 0
               GoTo hell
               End If
            AvailableBal = rs.Fields("Amount") - CCur(rs.Fields("Commission")) - 500
               GoTo saddam
            ElseIf rs.Fields("transdescription") = "LEAVE " And rs.Fields("TransType") = "CR" Then
             If AvailableBal = 0 And I > 1 Then
               AvailableBal = 0
               GoTo hell
               End If
            AvailableBal = rs.Fields("Amount") - CCur(rs.Fields("Commission")) - 500
               GoTo saddam
            ElseIf rs.Fields("transdescription") = "Members Welfare " And rs.Fields("TransType") = "CR" Then
             If AvailableBal = 0 And I > 1 Then
               AvailableBal = 0
               GoTo hell
               End If
            AvailableBal = rs.Fields("Amount") - CCur(rs.Fields("Commission")) - 500
               GoTo saddam
            ElseIf rs.Fields("transdescription") = "INTEREST REFUND " And rs.Fields("TransType") = "CR" Then
             If AvailableBal = 0 And I > 1 Then
               AvailableBal = 0
               GoTo hell
               End If
            AvailableBal = rs.Fields("Amount") - CCur(rs.Fields("Commission")) - 500
               GoTo saddam
            ElseIf rs.Fields("transdescription") = "DAIRIES " And rs.Fields("TransType") = "CR" Then
             If AvailableBal = 0 And I > 1 Then
               AvailableBal = 0
               GoTo hell
               End If
            AvailableBal = rs.Fields("Amount") - CCur(rs.Fields("Commission")) - 500
               GoTo saddam
            ElseIf rs.Fields("transdescription") = "CO-OP HOUSE DIVIDENDS " And rs.Fields("TransType") = "CR" Then
             If AvailableBal = 0 And I > 1 Then
               AvailableBal = 0
               GoTo hell
               End If
            AvailableBal = rs.Fields("Amount") - CCur(rs.Fields("Commission")) - 500
               GoTo saddam
            ElseIf rs.Fields("transdescription") = "TEA PR " And rs.Fields("TransType") = "CR" Then
             If AvailableBal = 0 And I > 1 Then
               AvailableBal = 0
               GoTo hell
               End If
            AvailableBal = rs.Fields("Amount") - CCur(rs.Fields("Commission")) - 500
               GoTo saddam
            ElseIf rs.Fields("transdescription") = "COFFEE " And rs.Fields("TransType") = "CR" Then
             If AvailableBal = 0 And I > 1 Then
               AvailableBal = 0
               GoTo hell
               End If
            AvailableBal = rs.Fields("Amount") - CCur(rs.Fields("Commission")) - 500
               GoTo saddam
            ElseIf rs.Fields("transdescription") = "REMMITANCE MWALIMU " And rs.Fields("TransType") = "CR" Then
             If AvailableBal = 0 And I > 1 Then
               AvailableBal = 0
               GoTo hell
               End If
            AvailableBal = rs.Fields("Amount") - CCur(rs.Fields("Commission")) - 500
               GoTo saddam
            ElseIf rs.Fields("transdescription") = "LASDAP " And rs.Fields("TransType") = "CR" Then
             If AvailableBal = 0 And I > 1 Then
               AvailableBal = 0
               GoTo hell
               End If
            AvailableBal = rs.Fields("Amount") - CCur(rs.Fields("Commission")) - 500
               GoTo saddam
            ElseIf rs.Fields("transdescription") = "STO " And rs.Fields("TransType") = "CR" Then
             If AvailableBal = 0 And I > 1 Then
               AvailableBal = 0
               GoTo hell
               End If
            AvailableBal = rs.Fields("Amount") - CCur(rs.Fields("Commission")) - 500
               GoTo saddam
            ElseIf rs.Fields("transdescription") = "HONARARIA " And rs.Fields("TransType") = "CR" Then
             If AvailableBal = 0 And I > 1 Then
               AvailableBal = 0
               GoTo hell
               End If
            AvailableBal = rs.Fields("Amount") - CCur(rs.Fields("Commission")) - 500
               GoTo saddam
            ElseIf rs.Fields("transdescription") = "MANAGEMENT&STAFF DIVIDEND " And rs.Fields("TransType") = "CR" Then
             If AvailableBal = 0 And I > 1 Then
               AvailableBal = 0
               GoTo hell
               End If
            
            
               AvailableBal = rs.Fields("Amount") - CCur(rs.Fields("Commission")) - 500
               GoTo saddam
                
            ElseIf rs.Fields("transdescription") <> "2002 Balance B/F." Then
               If AvailableBal = 0 And I > 1 Then
               AvailableBal = 0
               GoTo hell
               End If
           AvailableBal = rs.Fields("Amount") - CCur(rs.Fields("Commission"))
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
           amount = CCur(rs.Fields("Amount")) + CCur(rs.Fields("Commission"))
            AvailableBal = AvailableBal - amount
          Else
              If description = "Cheque Deposit(uncleared)" Or description = "Cheque Dep(uncleared)" Then
              
               AvailableBal = AvailableBal
              Else
            amount = CCur(rs.Fields("Amount")) - CCur(rs.Fields("Commission"))
            AvailableBal = AvailableBal + amount
            End If
          End If
          
        If COMM > 0 Then
saddam1:
      sql = "update customerbalance set availablebalance=" & AvailableBal & " ,commission=" & COMM & "where  customerbalanceid =" & rs.Fields("customerbalanceid") & ""
          
          Set cn3 = New Connection
          cn3.Open SelectedDsn, "atm"
          cn3.Execute sql
          cn3.Close
          COMM = 0
          Set cn3 = Nothing
          End If
saddam:
          sql = "update customerbalance set availablebalance=" & AvailableBal & " where  customerbalanceid =" & rs.Fields("customerbalanceid") & ""
          
          Set cn2 = New Connection
          cn2.Open SelectedDsn, "atm"
          cn2.Execute sql
          cn2.Close
          Set cn2 = Nothing

      sql = "update cub set availablebalance=" & AvailableBal & ",Active=1 where accno='" & rs2.Fields("accno") & "'"
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
Me.Caption = "Rebuilder"
Me.MousePointer = 0

 MsgBox "Processing Complete"
Exit Sub
ErrHandler:
Me.MousePointer = 0
MsgBox err.description
End Sub
