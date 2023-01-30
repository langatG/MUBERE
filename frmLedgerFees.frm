VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLedgerFees 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Shares statements"
   ClientHeight    =   2025
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5505
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLedgerFees.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   5505
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Print"
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   1440
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "TOTAL  SHARE LIST"
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   2040
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox txtsno 
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   600
      Width           =   1335
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComCtl2.DTPicker dtpPeriod 
      Height          =   315
      Left            =   1770
      TabIndex        =   1
      Top             =   585
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/mm/yyyy"
      Format          =   106954753
      CurrentDate     =   43893
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Load Shares"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   1440
      Width           =   2280
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Supplier No"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   3960
      TabIndex        =   7
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Select Period"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   2085
      TabIndex        =   3
      Top             =   330
      Width           =   1140
   End
End
Attribute VB_Name = "frmLedgerFees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cnn As New Connection, rsAccts As New Recordset
Dim rsUpdate As New Recordset, rsIncome As New Recordset
Dim transdate As Date, LastIncomeDate As Date, lastcashdepositdate As Date, lastcashdeposit As Date, ACCNO As String, Date_Deduc As Date, num, num2 As Integer

Private Sub Deduct_Ledger_Fees()
On Error GoTo SysError
  GetTransactionNo
  Dim sno As Double
'  If txtsno = "" Then
'  MsgBox "Enter supplier no"
'  End If
'  Exit Sub
ProgressBar1.Min = 0
ProgressBar1.Max = 100
 num = ProgressBar1.Max / 2
  sql = ""
sql = "DELETE  FROM CONTRIB"
oSaccoMaster.ExecuteThis (sql)
  sql = "set dateformat dmy select * from d_supplier_deduc where  Description LIKE '%shares%'  and sno=" & txtsno & " order by date_deduc asc  "
  Set rs = oSaccoMaster.GetRecordset(sql)
   With rs
     If Not .EOF Then
       While Not .EOF
          amount = !amount
          memberno = !sno
          Date_Deduc = !Date_Deduc
          ProgressBar1 = 20
'            amount = IIf(amount >= 2000, 2000, amount)
'             sharesCode = "002"
            If amount > 0 Then
                  Set rs = oSaccoMaster.GetRecordset("SELECT count(sno)sno FROM d_supplier_deduc ")
                    If Not rs.EOF Then
                        sno = IIf(IsNull(rs(0)), 0 + 1, rs(0) + 1)
                    Else
                        sno = 1
                    End If
                    End If
                      oSaccoMaster.ExecuteThis ("set dateformat dmy Insert into Contrib(memberno,contrdate,refno,Amount,sharebal,interest,transby,ChequeNo,receiptno,remarks,auditid,sharescode,transactionno,used,fperiod,intRate,MaturityDate)" _
                      & " Values('" & memberno & "','" & Date_Deduc & "'," & RefNo & "," & amount & "," & amount & ",0,'" & User & "','Shares','shares','CHECK-OFF','" & User & "','" & sharesCode & "','" & transactionNo & "',0,1,0,'" & Date & "')")
'         sddd
         
         .MoveNext
         
'         frmLedgerFees.Caption = memberno
       Wend
     End If
   End With
'   aaaa
'sql = ""
   ProgressBar1 = ProgressBar1 + num
   oSaccoMaster.ExecuteThis (sql)
  sql = "set dateformat dmy select * from d_sconribution where  transdescription <>'Registration'  and sno=" & txtsno & " order by transdate asc  "
  Set rs = oSaccoMaster.GetRecordset(sql)
   With rs
     If Not .EOF Then
       While Not .EOF
          amount = !amount
          memberno = !sno
          Date_Deduc = !transdate
            'num = num + 1
'            amount = IIf(amount >= 2000, 2000, amount)
'             sharesCode = "002"
            If amount > 0 Then
                  Set rs = oSaccoMaster.GetRecordset("SELECT count(sno)sno FROM d_sconribution ")
                    If Not rs.EOF Then
                        sno = IIf(IsNull(rs(0)), 0 + 1, rs(0) + 1)
                    Else
                        sno = 1
                    End If
                    End If
                      oSaccoMaster.ExecuteThis ("set dateformat dmy Insert into Contrib(memberno,contrdate,refno,Amount,sharebal,interest,transby,ChequeNo,receiptno,remarks,auditid,sharescode,transactionno,used,fperiod,intRate,MaturityDate)" _
                      & " Values('" & memberno & "','" & Date_Deduc & "'," & RefNo & "," & amount & "," & amount & ",0,'" & User & "','Shares','shares','CASH ','" & User & "','" & sharesCode & "','" & transactionNo & "',0,1,0,'" & Date & "')")
'         sddd
       
         'ProgressBar1 = ProgressBar1 + num
         .MoveNext
         frmLedgerFees.Caption = memberno
       
       Wend
     End If
   End With
   num2 = ProgressBar1.Max - ProgressBar1
   ProgressBar1 = ProgressBar1 + num2
   MsgBox "Statement Loaded successfully", vbInformation
   'reportname = "STATEMENT.rpt"
    'Show_Sales_Crystal_Report STRFORMULA, reportname, ""
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption

End Sub

Private Sub cmdbenevollent_Click()
'frmbenevollent.Show vbModal
End Sub

Private Sub Command1_Click()
    If MsgBox("Do you want to statement share  statement for period Ending " & Format(dtpPeriod, "dd/mm/yyyy"), vbQuestion + vbYesNo, "Ledger Fees") = vbYes Then
        ProgressBar1 = 0
        Deduct_Ledger_Fees
       
    End If
    
End Sub

Private Sub Command2_Click()
On Error GoTo SysError
  GetTransactionNo
  Dim sno As Double
'  If txtsno = "" Then
'  MsgBox "Enter supplier no"
'  End If
'  Exit Sub
  sql = ""
sql = "DELETE  FROM CONTRIB"
oSaccoMaster.ExecuteThis (sql)
  sql = "set dateformat dmy select isnull(sum(amount),0)as Amount,sno,date_deduc from d_supplier_deduc where  Description LIKE '%shares%' group by SNO,Date_Deduc order by SNO asc  "
  Set rs = oSaccoMaster.GetRecordset(sql)
   With rs
     If Not .EOF Then
       While Not .EOF
          amount = !amount
          memberno = !sno
          Date_Deduc = !Date_Deduc
'            amount = IIf(amount >= 2000, 2000, amount)
'             sharesCode = "002"
            If amount > 0 Then
                  Set rs = oSaccoMaster.GetRecordset("SELECT count(sno)sno FROM d_supplier_deduc ")
                    If Not rs.EOF Then
                        sno = IIf(IsNull(rs(0)), 0 + 1, rs(0) + 1)
                    Else
                        sno = 1
                    End If
                    End If
                      oSaccoMaster.ExecuteThis ("set dateformat dmy Insert into Contrib2(memberno,contrdate,refno,Amount,sharebal,interest,transby,ChequeNo,receiptno,remarks,auditid,sharescode,transactionno,used,fperiod,intRate,MaturityDate)" _
                      & " Values('" & memberno & "','" & Date_Deduc & "'," & RefNo & "," & amount & "," & amount & ",0,'" & User & "','Shares','shares','CHECK-OFF','" & User & "','" & sharesCode & "','" & transactionNo & "',0,1,0,'" & Date & "')")
'         sddd
         
         .MoveNext
'         frmLedgerFees.Caption = memberno
       Wend
     End If
   End With
'   aaaa
'sql = ""
   
   oSaccoMaster.ExecuteThis (sql)
  sql = "set dateformat dmy select * from d_sconribution where  transdescription <>'Registration'  and sno=" & txtsno & " order by transdate asc  "
  Set rs = oSaccoMaster.GetRecordset(sql)
   With rs
     If Not .EOF Then
       While Not .EOF
          amount = !amount
          memberno = !sno
          Date_Deduc = !transdate
'            amount = IIf(amount >= 2000, 2000, amount)
'             sharesCode = "002"
            If amount > 0 Then
                  Set rs = oSaccoMaster.GetRecordset("SELECT count(sno)sno FROM d_sconribution ")
                    If Not rs.EOF Then
                        sno = IIf(IsNull(rs(0)), 0 + 1, rs(0) + 1)
                    Else
                        sno = 1
                    End If
                    End If
                      oSaccoMaster.ExecuteThis ("set dateformat dmy Insert into Contrib2(memberno,contrdate,refno,Amount,sharebal,interest,transby,ChequeNo,receiptno,remarks,auditid,sharescode,transactionno,used,fperiod,intRate,MaturityDate)" _
                      & " Values('" & memberno & "','" & Date_Deduc & "'," & RefNo & "," & amount & "," & amount & ",0,'" & User & "','Shares','shares','CASH ','" & User & "','" & sharesCode & "','" & transactionNo & "',0,1,0,'" & Date & "')")
'         sddd
         
         .MoveNext
         frmLedgerFees.Caption = memberno
       Wend
     End If
   End With
   MsgBox "Statement Loaded successfully", vbInformation
'   msgbox(Print Statement)
'   reportname = "STATEMENT.rpt"
'    Show_Sales_Crystal_Report STRFORMULA, reportname, ""
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption

End Sub

Private Sub Command3_Click()
reportname = "STATEMENT.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, ""

End Sub

Private Sub Form_Load()
    dtpPeriod = Format(Get_Server_Date, "dd/mm/yyyy")
End Sub

