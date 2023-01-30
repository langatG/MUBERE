VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmsharestransactions 
   Caption         =   "Shares Transactions"
   ClientHeight    =   9135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14190
   Icon            =   "frmsharestransactions.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9135
   ScaleWidth      =   14190
   StartUpPosition =   1  'CenterOwner
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   11760
      TabIndex        =   28
      Top             =   8040
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      Format          =   116457473
      CurrentDate     =   41201
   End
   Begin VB.CommandButton cmdupdate 
      Caption         =   "Update"
      Height          =   375
      Left            =   8880
      TabIndex        =   27
      Top             =   8640
      Width           =   1575
   End
   Begin VB.TextBox txtamount 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   6360
      TabIndex        =   25
      Top             =   8640
      Width           =   2055
   End
   Begin VB.TextBox txtto 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1200
      TabIndex        =   19
      Top             =   8040
      Width           =   2415
   End
   Begin VB.CommandButton cmdsharecerts 
      Caption         =   "shares certs"
      Height          =   375
      Left            =   10440
      TabIndex        =   18
      Top             =   7320
      Width           =   1335
   End
   Begin VB.CommandButton cmdperiodicshares 
      Caption         =   "Periodic Shares Contrib"
      Height          =   375
      Left            =   3240
      TabIndex        =   17
      Top             =   7320
      Width           =   1815
   End
   Begin VB.CommandButton cmdbalancelist 
      Caption         =   "Balances List"
      Height          =   255
      Left            =   6000
      TabIndex        =   16
      Top             =   7440
      Width           =   1455
   End
   Begin VB.TextBox txtsno 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1320
      TabIndex        =   15
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      Height          =   375
      Left            =   8400
      TabIndex        =   14
      Top             =   7320
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker DTPregdate 
      Height          =   375
      Left            =   5160
      TabIndex        =   10
      Top             =   1320
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      Format          =   116457473
      CurrentDate     =   40637
   End
   Begin VB.TextBox txtidno 
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   9480
      TabIndex        =   8
      Top             =   1320
      Width           =   2415
   End
   Begin MSComctlLib.ListView lvwshares 
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   9340
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Transaction Date"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Amount"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Balance"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Description"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Posted By"
         Object.Width           =   4410
      EndProperty
   End
   Begin VB.Label Label12 
      Caption         =   "Transfered amount"
      Height          =   375
      Left            =   4200
      TabIndex        =   26
      Top             =   8640
      Width           =   1815
   End
   Begin VB.Label balanceto 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1920
      TabIndex        =   24
      Top             =   8640
      Width           =   1575
   End
   Begin VB.Label Label11 
      Caption         =   "Balance"
      Height          =   255
      Left            =   240
      TabIndex        =   23
      Top             =   8640
      Width           =   1335
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5040
      TabIndex        =   22
      Top             =   8040
      Width           =   6495
   End
   Begin VB.Label Label9 
      Caption         =   "S No."
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   8040
      Width           =   615
   End
   Begin VB.Label Label8 
      Caption         =   "Name"
      Height          =   255
      Left            =   4080
      TabIndex        =   20
      Top             =   8160
      Width           =   495
   End
   Begin VB.Label txtbal 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   9360
      TabIndex        =   13
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label Label5 
      Caption         =   "Balance"
      Height          =   375
      Left            =   8520
      TabIndex        =   12
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Name"
      Height          =   255
      Left            =   4200
      TabIndex        =   11
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Reg Date"
      Height          =   255
      Left            =   4200
      TabIndex        =   9
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label7 
      Caption         =   "S No."
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1320
      TabIndex        =   6
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label txtxmemberno 
      Caption         =   "Member No."
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label txtlocation 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "Location"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   855
   End
   Begin VB.Label txtname 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5160
      TabIndex        =   2
      Top             =   120
      Width           =   6495
   End
   Begin VB.Label Label1 
      Caption         =   "ID No."
      Height          =   255
      Left            =   8400
      TabIndex        =   1
      Top             =   1440
      Width           =   735
   End
End
Attribute VB_Name = "frmsharestransactions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdbalancelist_Click()
 reportname = "sharesbal.rpt"

 Show_Sales_Crystal_Report "", reportname, ""
End Sub

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdperiodicshares_Click()
 reportname = "Sharescontrib.rpt"

 Show_Sales_Crystal_Report "", reportname, ""
End Sub

Private Sub cmdsharecerts_Click()
frmsharecertificates.Show vbModal, Me
End Sub

Private Sub cmdupdate_Click()
Dim txtTCHPBalances As Double
Set rst = New ADODB.Recordset
sql = "select bal from d_shares where sno= '" & txtTo & "'"
Set rst = oSaccoMaster.GetRecordset(sql)
If Not rst.EOF Then
txtTCHPBalances = rst.Fields(0)

 '//get the balance
 '//receipeint

sql = "SELECT     bal   FROM         d_sconribution  WHERE     sno ='" & txtTo & "'  ORDER BY transdate DESC, id DESC "
Dim rr As New ADODB.Recordset
Set rr = oSaccoMaster.GetRecordset(sql)
If Not rr.EOF Then
txtTCHPBalances = txtamount + CCur(balanceto)
',[sno],[transdate],[amount],[bal],[transdescription],[auditid],[auditdate],[mno]
  'From [EASYTEA].[dbo].[d_sconribution]
  sql = ""
  sql = "set dateformat dmy insert into d_sconribution([sno],[transdate],[amount],[bal],[transdescription],[auditid])"
  sql = sql & " values ('" & txtTo & "','" & DTPicker1 & "'," & txtamount & "," & txtTCHPBalances & ",'Shares-Transfer from" & txtsno & "','" & User & "') "
  oSaccoMaster.ExecuteThis (sql)
  
  sql = ""
  sql = "update d_shares set bal=" & txtTCHPBalances & " where sno='" & txtTo & "' "
  oSaccoMaster.ExecuteThis (sql)
'txtTCHPBALANCE = rr.Fields(0)
End If
End If
'donor of the share

sql = "select bal from d_shares where sno= '" & txtsno & "'"
Set rst = oSaccoMaster.GetRecordset(sql)
If Not rst.EOF Then
txtTCHPBalances = rst.Fields(0)

 '//get the balance
 '//receipeint

sql = "SELECT     bal   FROM         d_sconribution  WHERE     sno ='" & txtsno & "'  ORDER BY transdate DESC, id DESC "
'Dim rr As New ADODB.Recordset
Set rr = oSaccoMaster.GetRecordset(sql)
If Not rr.EOF Then
txtTCHPBalances = CCur(rr.Fields(0)) - CCur(txtamount)
',[sno],[transdate],[amount],[bal],[transdescription],[auditid],[auditdate],[mno]
  'From [EASYTEA].[dbo].[d_sconribution]
  sql = ""
  sql = "set dateformat dmy insert into d_sconribution([sno],[transdate],[amount],[bal],[transdescription],[auditid])"
  sql = sql & " values ('" & txtsno & "','" & DTPicker1 & "'," & ((txtamount) * -1) & "," & txtTCHPBalances & ",'Shares-Transfer to " & txtTo & "','" & User & "') "
  oSaccoMaster.ExecuteThis (sql)
  
  sql = ""
  sql = "update d_shares set bal=" & txtTCHPBalances & " where sno='" & txtsno & "' "
  oSaccoMaster.ExecuteThis (sql)

End If
End If

End Sub

Private Sub Form_Load()
'pick items from deduction from the
DTPicker1 = Format(Get_Server_Date, "dd/mm/yyyy")
End Sub

Private Sub txtidno_Change()


'sql = "SET dateformat dmy SELECT     sno, Name, Sex, Loc, mno,Type, regdate, Cash,bal"
'sql = sql & " From d_Shares WHERE  IdNo = '" & txtidno & "'" 'Period = '" & Enddate & "' AND
'
'Set rs2 = oSaccoMaster.GetRecordset(sql)
'If rs2.RecordCount > 0 Then
'txtSNo = rs2.Fields(0)
'txtname = rs2.Fields(1)
''cboSex = rs2.Fields(2)
'txtLocation = rs2.Fields(3)
'
'
'DTPregdate = IIf(IsNull(rs2.Fields(6)), Date, rs2.Fields(6))
''optCash.value = rs2.Fields(6).value
'txtbal = rs2.Fields(8)
''DTPicker2 = Enddate
'End If
'Dim tamount As Double
''//populate the items on the listview
'Set rs = oSaccoMaster.GetRecordset("SELECT     id, transdate, amount, bal, transdescription, auditid  FROM         d_sconribution where sno='" & txtSNo & "'")
'tamount = 0
'With rs
'While Not rs.EOF
'
'
'   Set li = lvwshares.ListItems.Add(, , IIf(IsNull(!id), 1, !id))
'   If rs.Fields("transdate") <> "" Then li.ListSubItems.Add , , rs.Fields("transdate")
'   If rs.Fields("Amount") <> "" Then li.ListSubItems.Add , , rs.Fields("Amount")
'   If rs.Fields("bal") <> "" Then li.ListSubItems.Add , , rs.Fields("bal")
'   If rs.Fields("transdescription") <> "" Then li.ListSubItems.Add , , rs.Fields("transdescription")
'   If rs.Fields("auditid") <> "" Then li.ListSubItems.Add , , rs.Fields("auditid")
'   tamount = tamount + rs.Fields("Amount")
'   .MoveNext
'
'Wend
'End With
'If tamount = 0 Then
'txtbal = rs.Fields(0)
'Else
'txtbal = tamount
'End If
'lvwshares.View = lvwReport
End Sub

Private Sub txtSNo_Change()


sql = "SET dateformat dmy SELECT     idno, Names, type, Location, mno, regdate"
sql = sql & " From d_suppliers WHERE  sno = '" & txtsno & "'" 'Period = '" & Enddate & "' AND

Set rs2 = oSaccoMaster.GetRecordset(sql)
If rs2.RecordCount > 0 Then
txtIdNo = rs2.Fields(0)
txtName = rs2.Fields(1)
'cboSex = rs2.Fields(2)
txtLocation = rs2.Fields(3)
Dim bal As Double

DTPregdate = IIf(IsNull(rs2.Fields(5)), Date, rs2.Fields(5))
'optCash.value = rs2.Fields(6).value
'Set rs = oSaccoMaster.GetRecordset("select bal from d_shares where sno='" & txtsno & "'")

Set rs = oSaccoMaster.GetRecordset(" SELECT     d_supplier_deduc.SNo, d_supplier_deduc.Description, SUM(d_supplier_deduc.Amount) AS amount, d_Suppliers.Names" _
& " FROM d_Suppliers AS d_Suppliers INNER JOIN d_supplier_deduc AS d_supplier_deduc ON d_Suppliers.SNo = d_supplier_deduc.SNo" _
& " WHERE     (d_supplier_deduc.Description LIKE '%shares%')" _
& "GROUP BY d_supplier_deduc.SNo, d_Suppliers.Names, d_supplier_deduc.Description HAVING      (d_supplier_deduc.SNo = '" & txtsno & "') ORDER BY d_supplier_deduc.SNo")
If Not rs.EOF Then
txtbal = rs.Fields(2)
bal = txtbal
End If
'DTPicker2 = Enddate
End If
Dim tamount As Double
'//populate the items on the listview
'Set rs = oSaccoMaster.GetRecordset("SELECT     id, transdate, amount, bal, transdescription, auditid  FROM         d_sconribution where sno='" & txtsno & "'")
'Set rs = oSaccoMaster.GetRecordset(" SELECT     d_supplier_deduc.SNo, d_supplier_deduc.Description, SUM(d_supplier_deduc.Amount) AS amount, d_Suppliers.Names" _
'& "FROM d_Suppliers AS d_Suppliers INNER JOIN d_supplier_deduc AS d_supplier_deduc ON d_Suppliers.SNo = d_supplier_deduc.SNo" _
'& "WHERE     (d_supplier_deduc.Description LIKE '%shares%')" _
'& "GROUP BY d_supplier_deduc.SNo, d_Suppliers.Names, d_supplier_deduc.Description HAVING      (d_supplier_deduc.SNo = '" & txtsno & "') ORDER BY d_supplier_deduc.SNo")
Set rs = oSaccoMaster.GetRecordset("SELECT    id, SNo, Date_Deduc, Description, Amount,auditid From d_supplier_deduc WHERE     (Description LIKE '%shares%') AND (SNo = '" & txtsno & "') ORDER BY Date_Deduc")
tamount = 0
lvwshares.ListItems.Clear
With rs
While Not rs.EOF
  
   
'   Set li = lvwshares.ListItems.Add(, , IIf(IsNull(!Id), 1, !Id))
'   If rs.Fields("transdate") <> "" Then li.ListSubItems.Add , , rs.Fields("transdate")
'   If rs.Fields("Amount") <> "" Then li.ListSubItems.Add , , rs.Fields("Amount")
'   If rs.Fields("bal") <> "" Then li.ListSubItems.Add , , rs.Fields("bal")
'   If rs.Fields("transdescription") <> "" Then li.ListSubItems.Add , , rs.Fields("transdescription")
'   If rs.Fields("auditid") <> "" Then li.ListSubItems.Add , , rs.Fields("auditid")
'   tamount = tamount + rs.Fields("Amount")
'   .MoveNext
Set li = lvwshares.ListItems.Add(, , IIf(IsNull(!Id), 1, !Id))
   If rs.Fields("Date_Deduc") <> "" Then li.ListSubItems.Add , , rs.Fields("Date_Deduc")
   If rs.Fields("Amount") <> "" Then li.ListSubItems.Add , , rs.Fields("Amount")
   If rs.Fields("amount") <> "" Then li.ListSubItems.Add , , rs.Fields("amount")
   If rs.Fields("Description") <> "" Then li.ListSubItems.Add , , rs.Fields("Description")
   If rs.Fields("auditid") <> "" Then li.ListSubItems.Add , , rs.Fields("auditid")
   tamount = tamount + rs.Fields("Amount")
   'If rs.Fields("tamount") <> "" Then li.ListSubItems.Add , , rs.Fields("tamount")
   .MoveNext
Wend
End With
If tamount = 0 Then
'txtbal = rs.Fields(0)
Else
'txtbal = tamount
End If
lvwshares.View = lvwReport

End Sub

Private Sub txtTo_Change()


sql = "SET dateformat dmy SELECT     idno, Names, type, Location, mno, regdate"
sql = sql & " From d_suppliers WHERE  sno = '" & txtTo & "'" 'Period = '" & Enddate & "' AND

Set rs2 = oSaccoMaster.GetRecordset(sql)
If rs2.RecordCount > 0 Then
txtIdNo = rs2.Fields(0)
Label10 = rs2.Fields(1)

'optCash.value = rs2.Fields(6).value
Set rs = oSaccoMaster.GetRecordset("select bal from d_shares where sno='" & txtTo & "'")
If Not rs.EOF Then
balanceto = rs.Fields(0)
End If
End If
End Sub
