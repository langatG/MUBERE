VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmreceiptadjustment 
   Appearance      =   0  'Flat
   BackColor       =   &H80000004&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "INVOICE ADJUSTMENTS"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10635
   Icon            =   "frmreceiptsadj.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   10635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Height          =   495
      Left            =   4680
      Picture         =   "frmreceiptsadj.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Cancel Process"
      Top             =   3600
      Width           =   495
   End
   Begin VB.CommandButton cmdUpdate 
      Height          =   495
      Left            =   4200
      Picture         =   "frmreceiptsadj.frx":0544
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Save Record"
      Top             =   3600
      Width           =   495
   End
   Begin VB.CommandButton cmdDelete 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      Picture         =   "frmreceiptsadj.frx":0646
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Delete Record"
      Top             =   3600
      Width           =   495
   End
   Begin VB.CommandButton cmdEdit 
      Height          =   495
      Left            =   3240
      Picture         =   "frmreceiptsadj.frx":0748
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Edit Record"
      Top             =   3600
      Width           =   495
   End
   Begin VB.CommandButton cmdAdd 
      Height          =   495
      Left            =   2760
      Picture         =   "frmreceiptsadj.frx":084A
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Add New record"
      Top             =   3600
      Width           =   495
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "CLOSE"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      TabIndex        =   20
      Top             =   3600
      Width           =   1230
   End
   Begin MSComctlLib.ListView lvwSummary 
      Height          =   2295
      Left            =   120
      TabIndex        =   19
      Top             =   4320
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   4048
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   16711680
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Frame fraBank 
      Height          =   2775
      Left            =   120
      TabIndex        =   12
      Top             =   600
      Width           =   9120
      Begin VB.TextBox TxtR_Id 
         Height          =   285
         Left            =   120
         TabIndex        =   28
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtnewreceipts 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7320
         TabIndex        =   26
         Top             =   360
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CheckBox chkreceiptno 
         Caption         =   "Change Invoice No"
         Height          =   255
         Left            =   5520
         TabIndex        =   25
         Top             =   360
         Width           =   1695
      End
      Begin VB.ComboBox txtpcode 
         Height          =   315
         Left            =   3720
         TabIndex        =   24
         Top             =   360
         Width           =   1575
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         DrawStyle       =   2  'Dot
         DrawWidth       =   17015
         Height          =   360
         Left            =   3360
         Picture         =   "frmreceiptsadj.frx":0D7C
         ScaleHeight     =   360
         ScaleWidth      =   240
         TabIndex        =   23
         Top             =   360
         Width           =   240
      End
      Begin MSComCtl2.DTPicker txttransdate 
         Height          =   255
         Left            =   5280
         TabIndex        =   22
         Top             =   1560
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         Format          =   130416641
         CurrentDate     =   38859
      End
      Begin VB.TextBox txtReceiptno 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1800
         TabIndex        =   6
         Top             =   375
         Width           =   1575
      End
      Begin VB.TextBox txtpName 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   5655
      End
      Begin VB.TextBox txtNoOfproducts 
         BackColor       =   &H80000004&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2280
         TabIndex        =   10
         Top             =   2280
         Width           =   1935
      End
      Begin VB.TextBox txtquantity 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   3360
         TabIndex        =   9
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox txtserialno 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   1560
         Width           =   3135
      End
      Begin VB.Label Label8 
         Caption         =   "Invoice Id"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Date Of Transaction"
         Height          =   255
         Left            =   5280
         TabIndex        =   21
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Invoice No"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1785
         TabIndex        =   18
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Customer Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label6 
         Caption         =   "No. of Items in This Invoice"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2280
         TabIndex        =   16
         Top             =   2040
         Width           =   2895
      End
      Begin VB.Label Label5 
         Caption         =   "Customer No."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3720
         TabIndex        =   15
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3315
         TabIndex        =   14
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label3 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   1815
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10635
      _ExtentX        =   18759
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      Begin VB.CheckBox chkPreviewReport 
         Caption         =   "Preview &Report"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3960
         TabIndex        =   11
         Top             =   0
         Width           =   1815
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6360
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmreceiptsadj.frx":0EFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmreceiptsadj.frx":1010
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmreceiptsadj.frx":1122
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmreceiptsadj.frx":1234
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmreceiptadjustment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim disablemodifying As Boolean
Dim pcode As String
Dim M As Boolean
Dim transdate As Date

Private Sub chkreceiptno_Click()
If chkreceiptno = vbChecked Then
txtnewreceipts.Visible = True
Else
txtnewreceipts.Visible = False
End If
End Sub

Private Sub cmdclose_Click()

Unload Me

End Sub
Private Sub cleartext()
TxtR_Id.Text = ""
txtreceiptno.Text = ""
txtpcode.ListIndex = -1
txtpcode.Text = ""
txtquantity.Text = ""
txtpname.Text = ""
End Sub

Private Sub cmddelete_Click()

On Error GoTo ErrorHandler

If txtpcode = "" Then
MsgBox "Product should be selected before you proceed", vbInformation, "ag_receipts Adjustments"
Exit Sub
End If

If txtreceiptno = "" Then
MsgBox "Receipt Number should be selected before you proceed", vbInformation, "ag_receipts Adjustments"
Exit Sub
End If

Provider = cn
Set cn = New ADODB.Connection
cn.Open Provider, "atm", "atm"

sql = ""
sql = "set dateformat dmy delete from  ag_receipts1  where r_id=" & TxtR_Id & " and p_code='" & txtpcode & "'  and t_date='" & txttransdate & "'"
cn.Execute sql

'//update the ag_stockbalance database

sql = ""
sql = "set dateformat dmy delete from ag_stockbalance1  where p_code='" & txtpcode & "' and r_no='" & txtreceiptno & "' and transdate='" & txttransdate & "' "
cn.Execute sql

MsgBox "Item successfully Deleted", vbInformation, "Stocks"
cleartext
Exit Sub
ErrorHandler:
MsgBox err.description

End Sub

Private Sub cmdupdate_Click()
'//update the ag_receipts and stock balance
If txtpcode = "" Then
MsgBox "Product should be selected before you proceed", vbInformation, "ag_receipts Adjustments"
Exit Sub
End If
Provider = cn
Set cn = New ADODB.Connection
cn.Open Provider, "atm", "atm"
sql = ""
'If reportpath = "" Then reportpath = GetSetting("payroll", "AppName", "rptPath", rptPath)
sql = "select r_no,P_CODE,S_NO,Qua,t_date from ag_receipts1 where r_no=" & txtreceiptno & " and p_code='" & txtpcode & "'"
Set rs = New ADODB.Recordset
rs.Open sql, cn
If Not rs.EOF Then

If chkreceiptno = vbChecked Then
sql = ""
sql = "update ag_receipts1 set s_no='" & txtSERIALNO & "',r_no=" & txtnewreceipts & " where t_date=#" & txttransdate & "# and p_code='" & txtpcode & "' and qua=" & txtquantity & ""
cn.Execute sql

'//update the ag_stockbalance database
sql = ""
sql = "update  ag_stockbalance1 set r_no='" & txtnewreceipts & "' where  transdate=#" & txttransdate & "# and  p_code='" & txtpcode & "' and changeinstock=" & (txtquantity * -1) & ""
cn.Execute sql

Else
sql = ""
sql = "set dateformat dmy update ag_receipts1 set s_no='" & txtSERIALNO & "',qua=" & txtquantity & ",t_date='" & txttransdate & "' where r_no=" & txtreceiptno & " and p_code='" & txtpcode & "'"
cn.Execute sql

'//update the ag_stockbalance database
sql = ""
'sql = "update  ag_stockbalance"
sql = "set dateformat dmy UPDATE ag_stockbalance1 SET ag_stockbalance1.changeinstock = " & (txtquantity * -1) & " WHERE (((ag_stockbalance1.p_code)='" & txtpcode & "') AND ((ag_stockbalance1.R_NO)='" & txtreceiptno & "') AND ((ag_stockbalance1.transdate)='" & txttransdate & "'))"
'sql = sql & "  set changeinstock=" & (txtquantity * -1) & ",  transdate='" & txttransdate & "' where p_code='" & txtpcode & "' and r_no='" & txtReceiptno & "' and transdate=#" & txttransdate & "#"
cn.Execute sql
End If
End If
'//call syscronizer

sychronice

End Sub
Private Sub sychronice()
Set rst = New Recordset
Dim rssy As New Recordset
Dim openbal As Double
Dim chg As Double
Dim bal As Double
Dim I As Integer
Dim pcode
Provider = cn
Set cn = New ADODB.Connection
cn.Open Provider, "atm", "atm"
sql = ""
sql = "SELECT p_code  From ag_stockbalance1 ORDER BY p_code"
rst.Open sql, cn, adOpenKeyset, adLockOptimistic
'While Not rst.EOF
'If Not IsNull(rst.Fields(0)) Then pcode = rst.Fields(0)
'// open it again but not order by track id again
sql = ""
sql = "SELECT * From ag_stockbalance1  where p_code='" & txtpcode & "' ORDER BY transdate,   trackid asc"
Set rssy = New ADODB.Recordset
rssy.Open sql, cn, adOpenKeyset, adLockOptimistic
I = 1
While Not rssy.EOF
I = I
'// update the balance
'// if it is the first one then leave it
If I = 1 Then '// get the openning balance
If Not IsNull(rssy.Fields("ag_stockbalance")) Then openbal = rssy.Fields("ag_stockbalance")
Else
chg = rssy.Fields("changeinstock")
bal = openbal + rssy.Fields("changeinstock")
sql = ""
sql = "update ag_stockbalance1 set openningstock=" & openbal & ",changeinstock=" & chg & ",stockbalance=" & bal & " where trackid=" & rssy.Fields("trackid") & ""
cn.Execute sql
openbal = bal
End If
I = I + 1
rssy.MoveNext
Wend
If txtpcode <> "" Then
sql = ""
sql = "update ag_products1 set qout=" & bal & " where p_code='" & txtpcode & "'"
cn.Execute sql
End If
I = 0
'rst.Requery
'Set rssy = Nothing
'rst.MoveNext
'Wend
MsgBox "Process Complete", vbInformation
End Sub
Private Sub Form_Load()
On Error GoTo errFix

'Set rst = oSaccoMaster.GetRecordset("select bankcode from banks order by bankcode")
'With rst
'    If .RecordCount > 0 Then
'        .MoveFirst
'        txtBankCode.Text = !BankCode & ""
'        lvwSummary.Visible = False
'        cmdFirst.Enabled = True
'        cmdPrevious.Enabled = True
'        cmdNext.Enabled = True
'        cmdLast.Enabled = True
'        load_records
'    End If
'End With

'    'If Not Rst4.EOF Then
'        Set Rst5 = oSaccoMaster.GetRecordset("select * from usergrps where groupid= '" & Rst4!groupid & "'")
'        If Not Rst5.EOF Then
'            valToEncrOrDecr = Rst5!banksetup & vbNullString
'            EncryptPassword
'            If EncryptPass = "View" Then
'                disablemodifying = True
'                cmdAdd.Enabled = False
'                cmdUpdate.Enabled = False
'                cmdCancel.Enabled = False
'                cmdDelete.Enabled = False
'                cmdEdit.Enabled = False
'                chkPreviewReport.Enabled = False
'                chkPreviewReport.Value = vbChecked
'            ElseIf EncryptPass = "Mod" Then
'                disablemodifying = False
'                chkPreviewReport.Enabled = True
'                chkPreviewReport.Value = vbChecked
'            End If
'        End If
'   ' Else
'        chkPreviewReport.Value = vbChecked
'        chkPreviewReport.Enabled = True
'
   ' End If
errFix:

End Sub

Public Sub Load_Records()
On Error GoTo errFix
cmdCancel.Enabled = False
cmdupdate.Enabled = False
If disablemodifying = False Then
cmdAdd.Enabled = True
cmdEdit.Enabled = True
End If
fraBank.Enabled = False
Dim ctrl As Control
For Each ctrl In Me
        If TypeOf ctrl Is TextBox Then
            ctrl.Locked = True
        End If
        
        If TypeOf ctrl Is ComboBox Then
            ctrl.Locked = True
        End If
        
        If TypeOf ctrl Is CheckBox Then
        If Not ctrl = chkPreviewReport Then
            ctrl.Enabled = False
        End If
        End If
    Next ctrl
cmdCancel.Enabled = False
Provider = cn
Set cn = New ADODB.Connection
cn.Open Provider, "atm", "atm"
sql = ""
'If reportpath = "" Then reportpath = GetSetting("payroll", "AppName", "rptPath", rptPath)
sql = "select r_id,r_no,P_CODE,S_NO,Qua,t_date from ag_receipts1 where r_no=" & txtreceiptno & " order by p_code"
Set rs = New ADODB.Recordset
rs.Open sql, cn
If Not rs.EOF Then
If Not IsNull(rs.Fields(0)) Then TxtR_Id = (rs.Fields(0))
If Not IsNull(rs.Fields(1)) Then txtreceiptno = (rs.Fields(1))
If Not IsNull(rs.Fields(2)) Then txtpcode = (rs.Fields(2))
If Not IsNull(rs.Fields(3)) Then txtSERIALNO = (rs.Fields(3))
If Not IsNull(rs.Fields(4)) Then txtquantity = (rs.Fields(4))
If Not IsNull(rs.Fields(5)) Then txttransdate = (rs.Fields(5))
'If Not IsNull(rs.Fields(3)) Then lblbalance = (rs.Fields(3))
Call cboname_p
End If

Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Bank Setup"
End Sub

Private Sub Picture2_Click()
On Error Resume Next
frmsearchcustomerpayments.Show vbModal
Dim Y As String
Y = sel
M = False
If Y <> "" Then
Provider = cn
Set cn = New ADODB.Connection
cn.Open Provider, "atm", "atm"
sql = ""
'If reportpath = "" Then reportpath = GetSetting("payroll", "AppName", "rptPath", rptPath)
sql = "select r_no,P_CODE,S_NO,Qua,t_date,r_id from ag_receipts1 where r_id=" & Y & " order by p_code"
Set rs = New ADODB.Recordset
rs.Open sql, cn
If Not rs.EOF Then
txtpcode.Clear
If Not IsNull(rs.Fields(0)) Then txtreceiptno = (rs.Fields(0))
If Not IsNull(rs.Fields(1)) Then txtpcode = (rs.Fields(1))
If Not IsNull(rs.Fields(2)) Then txtSERIALNO = (rs.Fields(2))
If Not IsNull(rs.Fields(3)) Then txtquantity = (rs.Fields(3))
If Not IsNull(rs.Fields(4)) Then transdate = rs.Fields(4)
If Not IsNull(rs.Fields(4)) Then txttransdate = (rs.Fields(4))
If Not IsNull(rs.Fields(5)) Then TxtR_Id = (rs.Fields(5))

Dim NoofProd As String

sql = ""
sql = "SELECT COUNT(R_No) AS NoofProd From ag_receipts1 WHERE   R_No = '" & txtreceiptno & "'"
Set rs = New ADODB.Recordset
rs.Open sql, cn
If Not rs.EOF Then
If Not IsNull(rs.Fields(0)) Then txtNoOfproducts = (rs.Fields(0))

End If
Call cboname_p

End If
End If
End Sub
Private Sub cboname_p()
Provider = cn
pcode = txtpcode
Set cn = New ADODB.Connection
cn.Open Provider, "atm", "atm"
'If reportpath = "" Then reportpath = GetSetting("payroll", "AppName", "rptPath", rptPath)
sql = "select P_NAME from ag_products1 where p_code='" & pcode & "'"
Set rs = New ADODB.Recordset
rs.Open sql, cn
If Not rs.EOF Then
If Not IsNull(rs.Fields(0)) Then txtpname = (rs.Fields(0))
'If Not IsNull(rs.Fields(1)) Then lblbalance = rs.Fields(1)
End If
'//get the number of the items in the ag_receipts
Dim rsf As Recordset
Set rsf = New ADODB.Recordset
sql = ""
sql = "SELECT  p_code From ag_receipts1 WHERE (((ag_receipts1.R_No)=" & txtreceiptno & ")) and t_date='" & transdate & "' ;"
rsf.Open sql, cn, adOpenKeyset, adLockOptimistic

Dim t As Integer
If Not rsf.EOF Then
t = rsf.RecordCount
txtNoOfproducts = t
Else

End If

If M = False Then
While Not rsf.EOF
txtpcode.AddItem rsf.Fields(0)
rsf.MoveNext
M = True
Wend
End If
'txtpcode = pcode
End Sub

Private Sub txtpcode_Change()
On Error GoTo ErrorHandler
Provider = cn
Set cn = New ADODB.Connection
cn.Open Provider, "atm", "atm"
sql = ""
'If reportpath = "" Then reportpath = GetSetting("payroll", "AppName", "rptPath", rptPath)
sql = "select r_no,P_CODE,S_NO,Qua,t_date,r_id from ag_receipts1 where r_no='" & txtreceiptno & "' and p_code='" & txtpcode & "'"
Set rs = New ADODB.Recordset
rs.Open sql, cn
If Not rs.EOF Then
If Not IsNull(rs.Fields(0)) Then txtreceiptno = (rs.Fields(0))
If Not IsNull(rs.Fields(1)) Then txtpcode = (rs.Fields(1))
If Not IsNull(rs.Fields(2)) Then txtSERIALNO = (rs.Fields(2))
If Not IsNull(rs.Fields(3)) Then txtquantity = (rs.Fields(3))
If Not IsNull(rs.Fields(4)) Then txttransdate = (rs.Fields(4))
If Not IsNull(rs.Fields(5)) Then TxtR_Id = (rs.Fields(5))
'txtpcode.Clear
Call cboname_p
End If
Exit Sub
ErrorHandler:
 MsgBox err.description

End Sub

Private Sub txtpcode_Click()
txtpcode_Change
End Sub
