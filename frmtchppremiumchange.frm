VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmtchppremiumchange 
   Caption         =   "TCHP PREMIUM CHANGE"
   ClientHeight    =   4065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   5880
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox txtpremiumamount 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2520
      TabIndex        =   24
      Top             =   3120
      Width           =   2175
   End
   Begin VB.ComboBox cbocashmember 
      Height          =   315
      ItemData        =   "frmtchppremiumchange.frx":0000
      Left            =   2520
      List            =   "frmtchppremiumchange.frx":000A
      TabIndex        =   22
      Top             =   6360
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdupdatecashmember 
      Caption         =   "Update"
      Height          =   375
      Left            =   4920
      TabIndex        =   21
      Top             =   6360
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSComCtl2.DTPicker transdate 
      Height          =   375
      Left            =   2520
      TabIndex        =   20
      Top             =   2520
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   97386497
      CurrentDate     =   40929
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      Height          =   375
      Left            =   4920
      TabIndex        =   18
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton cmdupdatestatus 
      Caption         =   "Update"
      Height          =   375
      Left            =   4920
      TabIndex        =   17
      Top             =   6120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox cbostatus 
      Height          =   315
      ItemData        =   "frmtchppremiumchange.frx":001C
      Left            =   2520
      List            =   "frmtchppremiumchange.frx":002F
      TabIndex        =   16
      Top             =   6120
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cndupdate 
      Caption         =   "Update"
      Height          =   375
      Left            =   4920
      TabIndex        =   14
      Top             =   3120
      Width           =   855
   End
   Begin VB.TextBox txtTCHPBALANCE 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   2040
      TabIndex        =   6
      Top             =   1800
      Width           =   2175
   End
   Begin VB.TextBox txtTCHPCurrentStatus 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2040
      TabIndex        =   5
      Top             =   1440
      Width           =   2175
   End
   Begin VB.TextBox txtTCHPMembershipDate 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2040
      TabIndex        =   4
      Top             =   960
      Width           =   2175
   End
   Begin VB.TextBox txtTCHPMonthlyPremium 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2040
      TabIndex        =   3
      Top             =   480
      Width           =   2175
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3240
      TabIndex        =   2
      Top             =   0
      Width           =   2535
   End
   Begin VB.TextBox txtSNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   480
      TabIndex        =   1
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton cmdfind 
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1800
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label6 
      Caption         =   "Cash Member"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   6480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Trans Date"
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Status"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   6240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Change Premium Amount To"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Label Label10 
      Caption         =   "TCHP BALANCE"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label9 
      Caption         =   "TCHP Current status"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label8 
      Caption         =   "TCHP Membership Date"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label7 
      Caption         =   "TCHP Monthly Premium"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "Name:"
      Height          =   255
      Left            =   2400
      TabIndex        =   8
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "SNo:"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   495
   End
End
Attribute VB_Name = "frmtchppremiumchange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim DTPDDeduction As Date
Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdupdatecashmember_Click()
''statusr
On Error GoTo ErrorHandler
sql = ""
Dim cm As Integer
If cbocashmember = "Cash" Then
cm = 1
Else
cm = 1
End If

'tchp_members

sql = "UPDATE    tchp_members  SET            cashm=" & cm & " where sno='" & txtSNo & "'"
oSaccoMaster.ExecuteThis (sql)
Exit Sub
ErrorHandler:
MsgBox err.description
End Sub

Private Sub cmdupdatestatus_Click()
''statusr
On Error GoTo ErrorHandler
sql = ""
sql = ""
sql = "UPDATE    d_Suppliers   SET            statusr='" & cbostatus & "' where sno='" & txtSNo & "'"
oSaccoMaster.ExecuteThis (sql)

'tchp_members

sql = "UPDATE    tchp_members  SET            statusr='" & cbostatus & "' where sno='" & txtSNo & "'"
oSaccoMaster.ExecuteThis (sql)
Exit Sub
ErrorHandler:
MsgBox err.description
End Sub

Private Sub cmdfind_Click()
        Me.MousePointer = vbHourglass
      
        txttchpbalance = 0
        txtTCHPCurrentStatus = ""
        txttchpbalance = 0
        txtTCHPMembershipDate = ""
        txtTCHPMonthlyPremium = 0
        frmSearchSupplier.Show vbModal
        txtSNo = sel
        txtSNo_Validate True
        Me.MousePointer = 0
End Sub

Private Sub cndupdate_Click()
On Error GoTo ErrorHandler
Dim diff As Double
diff = CDbl(txtpremiumamount) - CDbl(txtTCHPMonthlyPremium)
sql = "set dateformat dmy UPDATE    tchp_members  SET   mpremium=" & txtpremiumamount & ",premiumc=" & diff & ",changedate='" & transdate & "' where sno='" & txtSNo & "'"
oSaccoMaster.ExecuteThis (sql)

 'change tchp premium on the supplier screen
sql = "set dateformat dmy UPDATE    d_suppliers  SET  thcppremium=" & txtpremiumamount & ",tmd='" & transdate & "' where sno='" & txtSNo & "'"
oSaccoMaster.ExecuteThis (sql)

Dim txtTCHPBalances As Double, balance As Double

sql = "SELECT     balance   FROM         tchp_trxs  WHERE     sno ='" & txtSNo & "'  ORDER BY transdate DESC, id DESC "
Dim rr As New ADODB.Recordset
Set rr = oSaccoMaster.GetRecordset(sql)
If Not rr.EOF Then
txtTCHPBalances = rr.Fields(0)
End If
balance = txtTCHPBalances + CDbl(diff)

'//add the debits
If diff > 0 Then
'diff = CCur(Day(Date) - Day(DepositDate))
If Day(transdate) <= 28 Then
sql = ""
sql = "set dateformat dmy INSERT INTO tchp_trxs"
sql = sql & "     (sno,transdate, description, Debits, CreditsD, CreditsC, Balance, auditid)"
sql = sql & " VALUES     ('" & txtSNo & "','" & transdate & "','Debit(Modification)'," & diff & ",0,0," & balance & ",'" & User & "')"
oSaccoMaster.ExecuteThis (sql)
End If
End If

'//subtract the debits
If diff < 0 Then
If Day(transdate) <= 28 Then
sql = ""
sql = "set dateformat dmy INSERT INTO tchp_trxs"
sql = sql & "     (sno,transdate, description, Debits, CreditsD, CreditsC, Balance, auditid)"
sql = sql & " VALUES     ('" & txtSNo & "','" & transdate & "','Debit(Modification)'," & diff & ",0,0," & balance & ",'" & User & "')"
oSaccoMaster.ExecuteThis (sql)
End If
End If
MsgBox "Tchp premium changed successfeully"
frmtchppremiumchange.Visible = False
Exit Sub
ErrorHandler:
MsgBox err.description
End Sub

Private Sub Form_Load()
transdate = Format(Get_Server_Date, "dd/mm/yyyy")
Set rs = CreateObject("adodb.recordset")
    
    rs.Open "SELECT RATE FROM Tchp_Rate ", cn
    
    If rs.EOF Then Exit Sub
    
    With rs
        
        While Not .EOF
         
         txtpremiumamount.AddItem rs.Fields(0)
         
         
         .MoveNext
        
        Wend
    
    End With
    
    txtpremiumamount.Locked = False
End Sub

Private Sub txtpremiumamount_Change()
txtpremiumamount.Locked = True
End Sub

Private Sub txtpremiumamount_Validate(Cancel As Boolean)
txtpremiumamount.Locked = True
End Sub

Private Sub txtSNo_Validate(Cancel As Boolean)
On Error GoTo ErrorHandler
Dim tchpa As Integer
Set rs = New ADODB.Recordset
sql = "d_sp_SelectSuppliers '" & txtSNo & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
If Not IsNull(rs.Fields(2)) Then txtName.Text = rs.Fields(2)
If Not IsNull(rs.Fields(24)) Then txtTCHPMembershipDate = rs.Fields(24)
Else
txtName.Text = ""
End If
''tchp_tchpmember
''SELECT     sno, aarno, mpremium, premium, tchpactive,balance    FROM         tchp_members where sno=@sno
Set Rst = New ADODB.Recordset
sql = "tchp_tchpmember '" & txtSNo & "'"
Set Rst = oSaccoMaster.GetRecordset(sql)
If Not Rst.EOF Then
txtTCHPMonthlyPremium = Rst.Fields(2)
'txtTCHPBalances = Rst.Fields(5)
tchpa = Rst.Fields(4)
If tchpa = 1 Then
txtTCHPCurrentStatus = Rst.Fields(6)
Else
txtTCHPCurrentStatus = Rst.Fields(6)
End If
End If

'//get the milk balance at this screen
'//get the milk balance at all the time
Dim Startdate As Date, NetP As Double
Dim Enddate As Date
Startdate = DateSerial(Year(DTPDDeduction), month(DTPDDeduction), 1)
Enddate = DateSerial(Year(DTPDDeduction), month(DTPDDeduction) + 1, 1 - 1)
If txtSNo = "" Then Exit Sub
Set rs = oSaccoMaster.GetRecordset("d_sp_SupNet '" & txtSNo & "','" & Startdate & "','" & Enddate & "', 0")
If rs.EOF Then GoTo jericho
If Not IsNull(rs.Fields(1)) Then
If Not IsNull(rs.Fields(1)) Then
NetP = rs.Fields(1)
Else
NetP = "0.00"
End If
jericho:
Set rs = oSaccoMaster.GetRecordset("d_sp_SupNet '" & txtSNo & "','" & Startdate & "','" & Enddate & "', 1")
If Not IsNull(rs.Fields(0)) Then
NetP = NetP - rs.Fields(0)
Else
NetP = NetP - 0
End If
'txtMilkAccountBalance = NetP
End If
'//get the balance from the trx table
sql = "SELECT     balance   FROM         tchp_trxs  WHERE     sno ='" & txtSNo & "'  ORDER BY transdate DESC, id DESC "
Dim rr As New ADODB.Recordset
Set rr = oSaccoMaster.GetRecordset(sql)
If Not rr.EOF Then
'txtTCHPBalances = rr.Fields(0)
txttchpbalance = rr.Fields(0)
End If
Exit Sub
ErrorHandler:
MsgBox err.description
End Sub
