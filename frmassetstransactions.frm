VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmassetstransactions 
   Caption         =   "AC-Asset Register  Transactions."
   ClientHeight    =   8010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9735
   Icon            =   "frmassetstransactions.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   9735
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboassetype 
      Height          =   315
      Left            =   6240
      TabIndex        =   53
      Top             =   1440
      Width           =   3135
   End
   Begin VB.CommandButton cmdbatchprocess 
      Caption         =   "Batch Process"
      Height          =   375
      Left            =   6840
      TabIndex        =   52
      Top             =   2040
      Width           =   1935
   End
   Begin VB.CommandButton cmdfinder 
      Height          =   285
      Left            =   3120
      Picture         =   "frmassetstransactions.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   50
      ToolTipText     =   "Add New record"
      Top             =   960
      Width           =   375
   End
   Begin VB.Frame Frame3 
      Height          =   2895
      Left            =   120
      TabIndex        =   30
      Top             =   3720
      Width           =   9495
      Begin VB.ComboBox cbosourcecode 
         Height          =   315
         ItemData        =   "frmassetstransactions.frx":0704
         Left            =   2160
         List            =   "frmassetstransactions.frx":0717
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   1080
         Width           =   2175
      End
      Begin VB.ComboBox cboyearperiod 
         Height          =   315
         ItemData        =   "frmassetstransactions.frx":0742
         Left            =   6600
         List            =   "frmassetstransactions.frx":07A9
         TabIndex        =   44
         Text            =   "cboyearperiod"
         Top             =   600
         Width           =   2175
      End
      Begin VB.ComboBox cbocurrency 
         Height          =   315
         ItemData        =   "frmassetstransactions.frx":0918
         Left            =   6600
         List            =   "frmassetstransactions.frx":092E
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox acccr 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2160
         MaxLength       =   15
         TabIndex        =   36
         Top             =   2280
         Width           =   1695
      End
      Begin VB.TextBox accdr 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2160
         MaxLength       =   15
         TabIndex        =   35
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox txtamount1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7680
         TabIndex        =   34
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox txtamount2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7680
         TabIndex        =   33
         Top             =   2280
         Width           =   1695
      End
      Begin VB.CommandButton Command7 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         Height          =   285
         Left            =   3840
         MaskColor       =   &H000000FF&
         Picture         =   "frmassetstransactions.frx":0950
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Add New record"
         Top             =   1680
         Width           =   375
      End
      Begin VB.CommandButton Command8 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         Height          =   285
         Left            =   3840
         Picture         =   "frmassetstransactions.frx":0C12
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Add New record"
         Top             =   2280
         Width           =   375
      End
      Begin MSComCtl2.DTPicker txttransdate 
         Height          =   255
         Left            =   2160
         TabIndex        =   43
         ToolTipText     =   "Month Calender"
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         Format          =   162660353
         CurrentDate     =   38856
      End
      Begin VB.Label Label15 
         Caption         =   "End Date"
         Height          =   255
         Left            =   240
         TabIndex        =   49
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label14 
         Caption         =   "Year/Period"
         Height          =   255
         Left            =   5280
         TabIndex        =   48
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label13 
         Caption         =   "Source Code"
         Height          =   255
         Left            =   240
         TabIndex        =   47
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label12 
         Caption         =   "Currency"
         Height          =   255
         Left            =   5280
         TabIndex        =   46
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "Amount"
         Height          =   255
         Left            =   7800
         TabIndex        =   41
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Lblname2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4320
         TabIndex        =   40
         Top             =   2280
         Width           =   3255
      End
      Begin VB.Label lblname1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4320
         TabIndex        =   39
         Top             =   1680
         Width           =   3255
      End
      Begin VB.Label Label10 
         Caption         =   "Credit Asset Account"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Label Label9 
         Caption         =   "Debit Accumulated"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   1680
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdtransreport 
      Caption         =   "Transactions Report"
      Height          =   375
      Left            =   6240
      TabIndex        =   28
      Top             =   7320
      Width           =   1575
   End
   Begin VB.CommandButton cmdcalculate 
      Caption         =   "Calculate"
      Height          =   375
      Left            =   600
      TabIndex        =   27
      Top             =   7320
      Width           =   1215
   End
   Begin VB.TextBox txtserialno 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   26
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      Height          =   375
      Left            =   4800
      TabIndex        =   24
      Top             =   7320
      Width           =   1215
   End
   Begin VB.CommandButton cmdprint 
      Caption         =   "Print"
      Height          =   375
      Left            =   3360
      TabIndex        =   23
      Top             =   7320
      Width           =   1215
   End
   Begin VB.CommandButton cmdpost 
      Caption         =   "Post"
      Height          =   375
      Left            =   2040
      TabIndex        =   22
      Top             =   7320
      Width           =   1095
   End
   Begin VB.TextBox txtvalue 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   21
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox txtassetsname 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   19
      Top             =   1440
      Width           =   3975
   End
   Begin VB.TextBox txtamountvaludep 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1680
      TabIndex        =   17
      Top             =   3240
      Width           =   2415
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   120
      TabIndex        =   13
      Top             =   2520
      Width           =   3975
      Begin VB.CheckBox chkrevaluation 
         Caption         =   "valuation"
         Height          =   255
         Left            =   1800
         TabIndex        =   15
         Top             =   240
         Width           =   1575
      End
      Begin VB.CheckBox chkdepreciation 
         Caption         =   "Depreciation"
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.TextBox txtnetrealisablevalue 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1680
      TabIndex        =   12
      Top             =   3960
      Width           =   2415
   End
   Begin VB.ComboBox cboyear 
      Height          =   315
      ItemData        =   "frmassetstransactions.frx":0ED4
      Left            =   4800
      List            =   "frmassetstransactions.frx":0F1D
      Style           =   1  'Simple Combo
      TabIndex        =   10
      Text            =   "cboyear"
      Top             =   0
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   7800
      TabIndex        =   4
      Top             =   120
      Width           =   1575
      Begin VB.CheckBox chkmonthly 
         Caption         =   "Monthly "
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Value           =   2  'Grayed
         Width           =   1095
      End
      Begin VB.CheckBox chk4thquarter 
         Caption         =   "4 th Quarter"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   2760
         Width           =   1335
      End
      Begin VB.CheckBox chk3quarter 
         Caption         =   "3 rd Quarter"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CheckBox chk2ndquarter 
         Caption         =   "2 nd Quarter"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CheckBox chk1stquarter 
         Caption         =   "1st Quarter"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.TextBox txtassetcode 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Top             =   960
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker DTPtransdate 
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   480
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   162660353
      CurrentDate     =   37982
   End
   Begin VB.Label Label16 
      Caption         =   "Asset Class"
      Height          =   255
      Left            =   6600
      TabIndex        =   51
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Serial No."
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Depre/Valu %"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Assets Name"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Amount Dep/Val"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Net Realisable Value"
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Year"
      Height          =   255
      Left            =   3720
      TabIndex        =   9
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Assets No."
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "End Date"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "frmassetstransactions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim purchase As Currency
Dim revaluation As Currency
Dim depreciation As Currency
Dim NETREALIZABLEVALUE As Currency
Dim aquarter
Dim lab As String
Dim myclass As cdbase
Dim cn As Connection
Dim quarterper
Dim Provider As String

Private Sub Check2_Click()

End Sub

Private Sub Check4_Click()

End Sub

Private Sub chk1stquarter_Click()
If chk1stquarter = vbChecked Then
aquarter = 1
If txtValue = "" Then txtValue = 0
quarterper = txtValue / 4
chk1stquarter = vbChecked
chk2ndquarter = vbUnchecked
chk3quarter = vbUnchecked
chk4thquarter = vbUnchecked
End If
End Sub

Private Sub chk2ndquarter_Click()
If chk2ndquarter = vbChecked Then
aquarter = 2
If txtValue = "" Then txtValue = 0

quarterper = txtValue / 2
chk1stquarter = vbUnchecked
chk2ndquarter = vbChecked
chk3quarter = vbUnchecked
chk4thquarter = vbUnchecked
End If
End Sub

Private Sub chk3quarter_Click()
If chk3quarter = vbChecked Then
aquarter = 3
If txtValue = "" Then txtValue = 0

quarterper = txtValue / 1.333
chk1stquarter = vbUnchecked
chk2ndquarter = vbUnchecked
chk3quarter = vbChecked
chk4thquarter = vbUnchecked
End If
End Sub

Private Sub chk4thquarter_Click()
If chk4thquarter = vbChecked Then
aquarter = 4
If txtValue = "" Then txtValue = 0 Else quarterper = txtValue / 1
chk1stquarter = vbUnchecked
chk2ndquarter = vbUnchecked
chk3quarter = vbUnchecked
chk4thquarter = vbChecked
End If
End Sub

Private Sub chkdepreciation_Click()
If chkdepreciation = vbChecked Then
 chkdepreciation = vbChecked
 Else
  chkrevaluation = vbUnchecked
End If
End Sub

Private Sub chkmonthly_Click()
If chkmonthly = vbChecked Then
chk1stquarter = vbUnchecked
chk2ndquarter = vbUnchecked
chk3quarter = vbUnchecked
chk4thquarter = vbUnchecked
End If
End Sub

Private Sub chkrevaluation_Click()
If chkrevaluation = vbChecked Then
chkrevaluation = vbChecked
Else
chkdepreciation = vbUnchecked
End If
End Sub

Private Sub cmdbatchprocess_Click()
On Error GoTo ErrorHandler
Dim cn As Connection
Dim per As Double
Dim rsf As New ADODB.Recordset
Dim rsd As New ADODB.Recordset
Dim code As String
Dim Name As String
Dim deprate As Double
Dim amountvaludep As Double
Dim netrealisablevalue As Double
Dim depreciation As Double
Dim pprice As Double
Dim transdate As Date, amount As Double, DRaccno As String, Craccno As String, DocumentNo As String
Dim TransSource As String, User1 As String, ErrorMessage As String, transDescription As String, CashBook As Long, doc_posted As Integer, chequeno As String
Set rs = New Recordset
Set cn = New Connection
cn.Open frmODBCLogon.cboDSNList, "atm"

If txtValue = "" Then txtValue = 25
per = txtValue
     
If accdr = "" Then
MsgBox "Gl to Dr must be selected before you proceed", vbInformation
accdr.SetFocus
Exit Sub
End If

If acccr = "" Then
MsgBox "Gl to cr must be selected before you proceed", vbInformation
acccr.SetFocus
Exit Sub
End If
 ''''delete existing
Startdate = DateSerial(Year(DTPTransdate), Month(DTPTransdate), 1)
Enddate = DateSerial(Year(DTPTransdate), Month(DTPTransdate) + 1, 1 - 1)
sql = ""
sql = "set dateformat dmy delete from gltransactions where transdescript='Asset Depreciation' and TransDate>='" & Startdate & "' and TransDate<='" & Enddate & "'"
oSaccoMaster.ExecuteThis (sql)
sql = ""
sql = "set dateformat dmy delete from assetstrans where transdate>='" & Startdate & "' and transdate<='" & Enddate & "'"
oSaccoMaster.ExecuteThis (sql)
''''end
    
sql = ""
sql = "SELECT     AssetsNo, AssetserialNo, AssetsName,depreciation,PurchasePrice from assets where status=0 order by AssetsNo"
Set rs = New ADODB.Recordset
rs.Open sql, cn

While Not rs.EOF

code = rs.Fields(0)
Name = rs.Fields(2)
deprate = rs.Fields(3)
txtValue = deprate
pprice = rs.Fields(4)
'//get accumulated depreciation

sql = ""
sql = "SELECT     SUM(Amountdep_val) AS accum FROM assetstrans WHERE     (Assetcode = '" & code & "')"
Set rsf = New ADODB.Recordset
rsf.Open sql, cn, adOpenKeyset, adLockOptimistic
If Not rsf.EOF Then
Dim tota As Currency
If Not IsNull(rsf.Fields(0)) Then tota = rsf.Fields(0) Else tota = 0
Else
tota = 0
End If
If chkdepreciation = vbChecked Then
depreciation = pprice * deprate / 100 * 1 / 12
amountvaludep = depreciation
'// get net realizable value
NETREALIZABLEVALUE = pprice - depreciation - tota
netrealisablevalue = NETREALIZABLEVALUE
amountvaludep = amountvaludep

End If
sql = ""
sql = "Select * from  assetstrans where assetcode='" & code & "' and year=" & CBOYEAR & " and quaters=0 and mmonth=" & Month(DTPTransdate) & ""
Set rst = New ADODB.Recordset
rst.Open sql, cn
If Not rst.EOF Then
GoTo sargoi

Else

sql = "SET  DATEFORMAT DMY insert into assetstrans ([year],assetcode,assetname,dep_val,amountdep_val,nrv,quaters,transdate,mmonth,posted,auditid,auditdatetime)"
sql = sql & " values (" & CBOYEAR & ",'" & code & "','" & Name & "'," & deprate & "," & amountvaludep & "," & netrealisablevalue & "," & (deprate / 100) & ",'" & txttransdate & "'," & Month(txttransdate) & ",1,'" & User & "','" & Get_Server_Date & "')"
cn.Execute sql
'// update the asset register before you proceed.
sql = ""
sql = "SET DATEFORMAT DMY UPDATE ASSETS SET CURRENTVALUE=" & netrealisablevalue & ",NRVBF=" & netrealisablevalue & " WHERE ASSETSNO='" & code & "'"
cn.Execute sql
'****************************************************************
sql = ""
sql = "Select * from  assets_register where assetcode='" & code & "'"
Set rst = New ADODB.Recordset
rst.Open sql, cn
If Not rst.EOF Then
DRaccno = rst!ContraAccNo
Craccno = rst!ACCNO
Else
 DRaccno = accdr
 Craccno = acccr
End If
transdate = DTPTransdate
  amount = amountvaludep
'  If DRaccno = "" Then accdr = "11-433"
'  If Craccno = "" Then acccr = "40309"
'  DRaccno = Trim(accdr)
'  Craccno = Trim(acccr)
  DocumentNo = code
        TransSource = code
        User1 = User
        transDescription = "Asset Depreciation"
        
        CashBook = 1
        doc_posted = 1
        chequeno = cbosourcecode
    If Not Save_GLTRANSACTION(transdate, amount, DRaccno, Craccno, DocumentNo, _
        TransSource, User1, ErrorMessage, transDescription, CashBook, doc_posted, chequeno, transactionNo) Then
            If ErrorMessage <> "" Then
                MsgBox ErrorMessage, vbInformation, Me.Caption
                ErrorMessage = ""
            End If
    End If
  End If
'sql = ""
'sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & DTPTransdate & "'," & Lvwitems.SelectedItem.SubItems(4) & ",'" & DRaccno & "','" & Craccno & "','" & Lvwitems.SelectedItem & "','" & cboproductname & "' ,'CHECK OFF SALES- " & "" & cboproductname & "','" & User & "',0,0)"
'oSaccoMaster.ExecuteThis (sql)


sargoi:
  netrealisablevalue = 0
  amountvaludep = 0
rs.MoveNext
Wend
MsgBox "You have successfully completed transactions", vbInformation
'// the credit end
txtamountvaludep = ""
txtassetcode = ""
txtASSETSNAME = ""
txtnetrealisablevalue = ""
txtSERIALNO = ""
txtValue = ""
txtamount1 = ""
txtamount2 = ""
Exit Sub




MsgBox "This asset has been posted for this period and cannot be posted again.", vbCritical
Exit Sub
'Else


Exit Sub
ErrorHandler:
MsgBox err.description

End Sub

Private Sub cmdcalculate_Click()
Dim cn As Connection
Set rs = New Recordset
Dim rst As Recordset
Set rst = New ADODB.Recordset
Set cn = New Connection
cn.Open frmODBCLogon.cboDSNList, "atm"
Dim reval As Currency
Dim lastval As Currency
sql = ""
sql = "Select TOP 1 *  from  assets where assetsno='" & txtassetcode & "'"
rst.Open sql, cn

If rst.EOF Then
purchase = purchase
Else
If Not IsNull(rst.Fields("purchaseprice")) Then lastval = rst.Fields("purchaseprice")
purchase = lastval
End If
'// get the accumulated depreciation from the asset trans
sql = ""
sql = "SELECT     SUM(Amountdep_val) AS accum FROM assetstrans WHERE     (Assetcode = '" & txtassetcode & "')"
Set rs = New ADODB.Recordset
rs.Open sql, cn, adOpenKeyset, adLockOptimistic
If Not rs.EOF Then
Dim tota As Currency
If Not IsNull(rs.Fields(0)) Then tota = rs.Fields(0) Else tota = 0
End If
If chkdepreciation = vbChecked Then
depreciation = purchase * txtValue / 100 * 1 / 12
txtamountvaludep = depreciation
'// get net realizable value
NETREALIZABLEVALUE = purchase - depreciation - tota
txtnetrealisablevalue = NETREALIZABLEVALUE
txtamount1 = txtamountvaludep
txtamount2 = txtamountvaludep
End If

If chkrevaluation Then
 revaluation = purchase * quarterper / 100
txtamountvaludep = revaluation
'// get net realizable value
NETREALIZABLEVALUE = purchase + revaluation
txtnetrealisablevalue = NETREALIZABLEVALUE
End If
End Sub

Private Sub cmdclose_Click()
Unload Me
End Sub


Private Sub cboassetype_change()
Dim rsg As Recordset

  Dim cn As Connection
    Set cn = New ADODB.Connection
    cn.Open frmODBCLogon.cboDSNList, "atm"
Dim r As Double
Dim lis As ListItem


Set rs = New Recordset



sql = "Select * from assets where assetsno='" & txtassetcode & "'"
rs.Open sql, cn
   sql = ""
   sql = "select rate from assetcode WHERE ASSETname='" & cboassetype & "'"
   Set rsg = New ADODB.Recordset
   rsg.Open sql, cn, adOpenKeyset, adLockOptimistic
   
 If Not rsg.EOF Then
   If Not IsNull(rsg.Fields(0)) Then r = rsg.Fields(0)
   txtValue = r
 End If
  
End Sub

Private Sub TXTASSETTYPE_Click()
cboassetype_change
End Sub

Private Sub cmdfinder_Click()
On Error GoTo ErrorHandler
frmsearchassets.Show vbModal
Dim Y As String
Y = sel
'm = False
If Y <> "" Then
     Dim cn As Connection
    Set cn = New ADODB.Connection
    
    cn.Open frmODBCLogon.cboDSNList, "atm"
    
sql = ""
'If reportpath = "" Then reportpath = GetSetting("payroll", "AppName", "rptPath", rptPath)
sql = "SELECT     AssetsNo, AssetserialNo, AssetsName,depreciation from assets where assetsno='" & Y & "' and status=0 order by AssetsNo"
Set rs = New ADODB.Recordset
rs.Open sql, cn
If Not rs.EOF Then


If Not IsNull(rs.Fields(1)) Then txtSERIALNO = (rs.Fields(1))
If Not IsNull(rs.Fields(2)) Then txtASSETSNAME = (rs.Fields(2))
If Not IsNull(rs.Fields(3)) Then txtValue = (rs.Fields(3))
If Not IsNull(rs.Fields(0)) Then txtassetcode = (rs.Fields(0))

If Not IsNull(rs.Fields(0)) Then acccr = (rs.Fields(0))
If Not IsNull(rs.Fields(2)) Then Lblname2 = (rs.Fields(2))
Else
MsgBox " Asset you are trying to select is already disposed", vbCritical
Exit Sub
End If
'get the serial no of the item

End If
Exit Sub
ErrorHandler:
MsgBox err.description

End Sub

Private Sub cmdPost_Click()
On Error GoTo ErrorHandler
Dim cn As Connection
Dim per As Double
 Dim transdate As Date, amount As Double, DRaccno As String, Craccno As String, DocumentNo As String, _
        TransSource As String, User1 As String, ErrorMessage As String, transDescription As String, CashBook As Long, doc_posted As Integer, chequeno As String
Set rs = New Recordset
Set cn = New Connection
cn.Open frmODBCLogon.cboDSNList, "atm"
If txtassetcode = "" Then
MsgBox "Assets code required before you proceed", vbInformation
Exit Sub
End If
If txtValue = "" Then txtValue = 0
per = txtValue
If txtamount1 <> txtamount2 Then
MsgBox "Debit and Credit should be equal", vbCritical
Exit Sub
End If

If accdr = "" Then
MsgBox "Gl to Dr must be selected before you proceed", vbInformation
accdr.SetFocus
Exit Sub
End If

If acccr = "" Then
MsgBox "Gl to cr must be selected before you proceed", vbInformation
acccr.SetFocus
Exit Sub
End If

If txtassetcode = "" Then
MsgBox "Asset Code cannot be blank", vbInformation
Else
sql = ""
If aquarter = "" Then aquarter = 0
sql = "Select * from  assetstrans where assetcode='" & txtassetcode & "' and year=" & CBOYEAR & " and quaters=0 and mmonth=" & Month(txttransdate) & ""
rs.Open sql, cn
If rs.EOF Then
If txtassetcode = "" Then
MsgBox "Asset code must be selected before you proceed", vbInformation
Exit Sub
Else

sql = "SET  DATEFORMAT DMY insert into assetstrans ([year],assetcode,assetname,dep_val,amountdep_val,nrv,quaters,transdate,mmonth,posted,auditid,auditdatetime)"
sql = sql & " values (" & CBOYEAR & ",'" & txtassetcode & "','" & txtASSETSNAME & "'," & txtValue & "," & txtamountvaludep & "," & txtnetrealisablevalue & "," & (txtValue.Text / 100) & ",'" & txttransdate & "'," & Month(txttransdate) & ",1,'" & User & "','" & Get_Server_Date & "')"
cn.Execute sql
'// update the asset register before you proceed.
sql = ""
sql = "SET DATEFORMAT DMY UPDATE ASSETS SET CURRENTVALUE=" & txtnetrealisablevalue & ",NRVBF=" & txtnetrealisablevalue & " WHERE ASSETSNO='" & txtassetcode & "'"
cn.Execute sql
'****************************************************************
transdate = DTPTransdate
  amount = txtamount1
  DRaccno = Trim(accdr)
  Craccno = Trim(acccr)
  DocumentNo = txtassetcode
        TransSource = txtassetcode
        User1 = User
        transDescription = "Asset Depreciation"
        CashBook = 1
        doc_posted = 1
        chequeno = txtassetcode
  If Not Save_GLTRANSACTION(transdate, amount, DRaccno, Craccno, DocumentNo, _
        TransSource, User1, ErrorMessage, transDescription, CashBook, doc_posted, chequeno, transactionNo) Then
            If ErrorMessage <> "" Then
                MsgBox ErrorMessage, vbInformation, Me.Caption
                ErrorMessage = ""
            End If
        End If
  End If

'// DO THE CONTRA ACCOUNTS
'Dim batchno As Double
'Dim batchdesc As String
'Dim entryno As Double
'Dim entrydesc As String
'Dim periodyear As String
''// debit the account
''//look for the batch no listing and increment the batch itself
'sql = "SELECT     TOP 1 batchno FROM         j_entry ORDER BY batchno DESC"
'Dim rsj As New ADODB.Recordset
'Set rsj = New ADODB.Recordset
''rsj.Open sql, cn, adOpenKeyset, adLockOptimistic
''If Not rsj.EOF Then
''Dim b As Long
''b = rsj.Fields(0) + 1
''End If
'batchdesc = "Depreciation for " & cboyearperiod & ""
'sql = ""
'sql = "SET DATEFORMAT DMY INSERT INTO j_entry"
'sql = sql & "(batchno, batch_desc, entry_no, entry_Desc, transdate, year_period, source_code, line_ref, reference, trans_desc, Acc_No, DR, CR, FxDR, FxCR,"
'sql = sql & " Quantity, curr_code, curr_rate, F_Curr, S_Curr, AuditOrg, Auditid, AuditDate)"
''sql = sql & "   VALUES     (" & b & ", '" & batchdesc & "', 1, '" & batchdesc & "', '" & txttransdate & "', '" & cboyearperiod & "', '" & cbosourcecode & "', 1, ' " & txtassetcode & " ', ' " & txtASSETSNAME & " ',"
'sql = sql & "  ' " & accdr & " ', " & txtamount1 & ", 0, " & txtamount1 & ", 0, 0, ' " & cbocurrency & " ', 1, " & txtamount1 & ", " & txtamount1 & ", ' IRNET ', ' " & User & " ', '" & Get_Server_Date & " ')"
''cn.Execute sql
''// update the balance of the account
'sql = ""
'sql = "select bal from glsetup where accno='" & accdr & "'"
''Dim rs As Recordset
'Set rs = New ADODB.Recordset
'rs.Open sql, cn, adOpenKeyset, adLockOptimistic
'Dim bal As Currency
'If Not rs.EOF Then
'If Not IsNull(rs.Fields(0)) Then bal = (rs.Fields(0))
'Else
' bal = 0
'End If
'
'sql = ""
'sql = "UPDATE    glsetup SET              bal =" & bal + txtamount1 & "  WHERE     (accno = '" & accdr & "')"
'cn.Execute sql
''//get to know the customer balance
'Dim CUB1 As Recordset
'Dim avail1 As Currency
'Dim mem2 As String
' sql = "select * from cub where accno='" & accdr & "'"
'            Set CUB1 = CreateObject("adodb.recordset")
'            CUB1.Open sql, cn
'            If CUB1.EOF Then
'             Else
'
'             If Not IsNull(CUB1!availablebalance) Then avail1 = CUB1!availablebalance
'             If Not IsNull(CUB1!memberno) Then mem2 = CUB1!memberno
'
'            End If
'
'sql = "set dateformat dmy insert into customerbalance (customerno,accname,amount,availablebalance,accno,transdescription,transdate,commission,period,transType,posted,locked,status,vno,auditid,auditdate,moduleid,accd,s_code) "
'sql = sql & " values ('" & mem2 & "','" & lblname1 & "'," & txtamount1 & "," & avail1 + txtamount1 & ",'" & accdr & "','" & txtASSETSNAME & "','" & Format(Get_Server_Date, "dd/mm/yyyy") & "',0,'" & month(Get_Server_Date) & "','DR',0,0,0,'" & batchdesc & "','" & User & "','" & Get_Server_Date & "','3','" & acccr & "','" & cbosourcecode & "' )"
'cn.Execute sql
'
'sql = ""
'sql = "set dateformat dmy update cub set amount=" & txtamount1 & ",transdescription='" & txtASSETSNAME & "',availablebalance=" & avail1 + txtamount1 & ",transdate='" & txttransdate & "',vno='" & batchdesc & "',period='" & month(Get_Server_Date) & "',auditid='" & User & "',auditdate='" & Get_Server_Date & "',moduleid=2 where accno='" & accdr & "'"
'cn.Execute sql
'Set CUB1 = Nothing
'
''// the credit end
'sql = ""
'sql = "SET DATEFORMAT DMY INSERT INTO j_entry"
'sql = sql & "(batchno, batch_desc, entry_no, entry_Desc, transdate, year_period, source_code, line_ref, reference, trans_desc, Acc_No, DR, CR, FxDR, FxCR,"
'sql = sql & " Quantity, curr_code, curr_rate, F_Curr, S_Curr, AuditOrg, Auditid, AuditDate)"
''sql = sql & "   VALUES     (" & b & ", '" & batchdesc & "', 1, '" & batchdesc & "',  '" & txttransdate & "', '" & cboyearperiod & "', '" & cbosourcecode & "', 1, ' " & txtassetcode & " ', ' " & txtASSETSNAME & " ',"
'sql = sql & "  ' " & acccr & " ', 0," & txtamount1 & ", 0, " & txtamount1 & ", 0, ' " & cbocurrency & " ', 1, " & txtamount1 & ", " & txtamount1 & ", ' IRNET ', ' " & User & " ', '" & Get_Server_Date & " ')"
''cn.Execute sql
''// update the balance of the account
'sql = ""
'sql = "select bal from glsetup where accno='" & acccr & "'"
'
'Set rs = New ADODB.Recordset
'rs.Open sql, cn, adOpenKeyset, adLockOptimistic
'If Not rs.EOF Then
'If Not IsNull(rs.Fields(0)) Then bal = rs.Fields(0)
'Else
'bal = 0
'End If
'
'
'sql = ""
'sql = "UPDATE    glsetup SET              bal = " & bal + txtamount2 & "  WHERE     (accno = '" & acccr & "')"
'cn.Execute sql
'
'sql = "select * from cub where accno='" & acccr & "'"
'            Set CUB1 = CreateObject("adodb.recordset")
'            CUB1.Open sql, cn
'           ' '//If Not IsNull(CUB1!memberno) Then mem2 = "& cub1!memberno &"
'            If CUB1.EOF Then
'             Else
'             If Not IsNull(CUB1!availablebalance) Then avail1 = CUB1!availablebalance
'              If Not IsNull(CUB1!memberno) Then mem2 = CUB1!memberno
'
'            End If
'
'sql = "set dateformat dmy insert into customerbalance (customerno,accname,amount,availablebalance,accno,transdescription,transdate,commission,period,transType,posted,locked,status,vno,auditid,auditdate,moduleid,accd,s_code) "
'sql = sql & " values ('" & mem2 & "','" & Lblname2 & "'," & txtamount2 & "," & avail1 + txtamount2 & ",'" & acccr & "','" & txtASSETSNAME & "','" & Format(Get_Server_Date, "dd/mm/yyyy") & "',0,'" & month(Get_Server_Date) & "','CR',0,0,0,'" & batchdesc & "','" & User & "','" & Get_Server_Date & "','3','" & accdr & "','" & cbosourcecode & "' )"
'cn.Execute sql
'
'sql = ""
'sql = "set dateformat dmy update cub set amount=" & txtamount2 & ",transdescription='" & txtASSETSNAME & "',availablebalance=" & avail1 + txtamount1 & ",transdate='" & txttransdate & "',vno='" & txtassetcode & "',period='" & month(Get_Server_Date) & "',auditid='" & User & "',auditdate='" & Get_Server_Date & "',moduleid=2 where accno='" & acccr & "'"
'cn.Execute sql
MsgBox "You have successfully completed transactions", vbInformation
'// the credit end
txtamountvaludep = ""
txtassetcode = ""
txtASSETSNAME = ""
txtnetrealisablevalue = ""
txtSERIALNO = ""
txtValue = ""
txtamount1 = ""
txtamount2 = ""
Exit Sub
End If

End If

MsgBox "This asset has been posted for this period and cannot be posted again.", vbCritical
Exit Sub
'Else


Exit Sub
ErrorHandler:
MsgBox err.description
End Sub

Private Sub cmdPrint_Click()
'fixedassetsschedule
Dim error As String
Dim path As String
path = Get_Path(error)
Set a = New CRAXDRT.Application

          Set r = a.OpenReport(path & "fixedassetsschedule.rpt")
          r.ReadRecords
          
          With frmReports.CRViewer1
              .ReportSource = r
              .ViewReport
          End With
          
          frmReports.Show vbModal
          
          Set r = Nothing
End Sub

Private Sub cmdtransreport_Click()
On Error GoTo ErrorHandler
Dim errormsg As String
Dim path As String
path = Get_Path(errormsg)
 Set a = New CRAXDRT.Application

          Set r = a.OpenReport(path & "Assetstransactions.rpt")
           r.ReadRecords
            sql = "{assetstrans.year}=" & CBOYEAR & ""
            r.RecordSelectionFormula = sql
          With frmReports.CRViewer1
              .ReportSource = r
              .ViewReport
          End With
          
          frmReports.Show vbModal
          
          Set r = Nothing
          Exit Sub
ErrorHandler:
          MsgBox err.description
End Sub

Private Sub Command7_Click()
Dim Z, S, U
Dim Provider As String
    'Dim sal
    'Dim procode
    Dim myclass As cdbase
    Dim rs As Recordset
Dim cn As Connection
accdr = ""

frmSearchGLAccounts.Show vbModal
'frmsearchnewacc.Show vbModal
    Z = SearchValue
 'Z = strName
    If Z <> "" Then
        lblname1 = Z
    End If

 Set cn = CreateObject("adodb.connection")
      Set myclass = New cdbase
    Provider = myclass.OpenCon
cn.Open Provider, "atm"
   sql = ""
   sql = "select * from glsetup where AccNo='" & Z & "'"
   Set rs = New ADODB.Recordset
   rs.Open sql, cn, adOpenKeyset, adLockOptimistic
   If rs.EOF Then
   Else
       If Not IsNull(rs.Fields("accno")) Then accdr = rs.Fields("accno")
       If Not IsNull(rs.Fields("glaccname")) Then lblname1 = rs.Fields("glaccname")
       'If Not IsNull(rs.Fields("accno")) Then Txtaccno = rs.Fields("accno")
End If
'Else
'frmsearchrecords.Show vbModal
'
'
'    Z = strName
'    If Z <> "" Then
'        txtaccname = Z
'
'        End If
'
'    Set cn = CreateObject("adodb.connection")
'    If txtaccname = "" Then Exit Sub
'      Set cn = CreateObject("adodb.connection")
'      Set myclass = New cdbase
'    Provider = myclass.OpenCon
'cn.Open Provider
'   sql = ""
'   sql = "select * from glsetup where glaccname='" & Z & "'"
'   Set rs = New ADODB.Recordset
'   rs.Open sql, cn, adOpenKeyset, adLockOptimistic
'   If Not rs.EOF Then
'        If Not IsNull(rs.Fields("glcode")) Then txtCode = rs.Fields("glcode")
'       If Not IsNull(rs.Fields("glaccname")) Then txtaccname = rs.Fields("glaccname")
'       If Not IsNull(rs.Fields("glaccno")) Then Txtaccno = rs.Fields("glaccno")
'       If Not IsNull(rs.Fields("glacctype")) Then cboaccoounttype = rs.Fields("glacctype")
'       If Not IsNull(rs.Fields("glaccgroup")) Then cboaccountgroup = rs.Fields("glaccgroup")
'       If Not IsNull(rs.Fields("normalbal")) Then cbonormalbalance = rs.Fields("normalbal")
'   If rs.Fields("glaccStatus") = 0 Then Optactive = True
'   If rs.Fields("glaccStatus") = 1 Then Optinactive = True
'        End If
'  End If
End Sub

Private Sub Command8_Click()
Dim Z, S, U
    'Dim sal
    'Dim procode
    Dim myclass As cdbase
    Dim rs As Recordset
Dim cn As Connection
acccr = ""
frmSearchGLAccounts.Show vbModal
'frmsearchassets.Show vbModal
'frmsearchnewacc.Show vbModal
 'Z = strName
  Z = sel
    If Z <> "" Then
        Lblname2 = Z
    End If

 Set cn = CreateObject("adodb.connection")
      Set myclass = New cdbase
    Provider = myclass.OpenCon
cn.Open Provider, "atm"
   sql = ""
   sql = "SELECT     AssetsNo, AssetsName from assets where assetsno='" & Z & "'"
   'sql = "select * from glsetup where glaccname='" & Z & "'"
   Set rs = New ADODB.Recordset
   rs.Open sql, cn, adOpenKeyset, adLockOptimistic
   If rs.EOF Then
   Else
       If Not IsNull(rs.Fields("AssetsNo")) Then acccr = rs.Fields("AssetsNo")
       If Not IsNull(rs.Fields("AssetsName")) Then Lblname2 = rs.Fields("AssetsName")
''       If Not IsNull(rs.Fields("accno")) Then acccr = rs.Fields("accno")
''       If Not IsNull(rs.Fields("glaccname")) Then Lblname2 = rs.Fields("glaccname")
       'If Not IsNull(rs.Fields("accno")) Then Txtaccno = rs.Fields("accno")
End If
'Else
'frmsearchrecords.Show vbModal
'
'
'    Z = strName
'    If Z <> "" Then
'        txtaccname = Z
'
'        End If
'
'    Set cn = CreateObject("adodb.connection")
'    If txtaccname = "" Then Exit Sub
'      Set cn = CreateObject("adodb.connection")
'      Set myclass = New cdbase
'    Provider = myclass.OpenCon
'cn.Open Provider
'   sql = ""
'   sql = "select * from glsetup where glaccname='" & Z & "'"
'   Set rs = New ADODB.Recordset
'   rs.Open sql, cn, adOpenKeyset, adLockOptimistic
'   If Not rs.EOF Then
'        If Not IsNull(rs.Fields("glcode")) Then txtCode = rs.Fields("glcode")
'       If Not IsNull(rs.Fields("glaccname")) Then txtaccname = rs.Fields("glaccname")
'       If Not IsNull(rs.Fields("glaccno")) Then Txtaccno = rs.Fields("glaccno")
'       If Not IsNull(rs.Fields("glacctype")) Then cboaccoounttype = rs.Fields("glacctype")
'       If Not IsNull(rs.Fields("glaccgroup")) Then cboaccountgroup = rs.Fields("glaccgroup")
'       If Not IsNull(rs.Fields("normalbal")) Then cbonormalbalance = rs.Fields("normalbal")
'   If rs.Fields("glaccStatus") = 0 Then Optactive = True
'   If rs.Fields("glaccStatus") = 1 Then Optinactive = True
'        End If
End Sub

Private Sub DTPtransdate_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
CBOYEAR = Year(DTPTransdate)
End Sub
Private Sub DTPtransdate_Change()
CBOYEAR = Year(DTPTransdate)
txttransdate.value = (DTPTransdate.value)
cboyearperiod = Format(txttransdate, "MM") & "-" & Format(txttransdate, "YYYY")
End Sub
Private Sub Form_Load()
On Error GoTo ErrorHandler
CBOYEAR.Text = DatePart("YYYY", Date)
CBOYEAR.Text = Year(Date)
DTPTransdate = Format(Get_Server_Date, "dd/mm/yyyy")
txttransdate = Format(Get_Server_Date, "dd/mm/yyyy")
cbosourcecode = "GL-DP"
cbocurrency = "KES"
cboyearperiod = Format(Get_Server_Date, "MM") & "-" & Format(Get_Server_Date, "YYYY")
'//get the asset accounts
Set cn = CreateObject("adodb.connection")
      Set myclass = New cdbase
    Provider = myclass.OpenCon
cn.Open Provider, "atm"
   sql = ""
   sql = ""
     sql = "select p_dep,a_dep from param"
     Set rs = New ADODB.Recordset
     rs.Open sql, cn, adOpenKeyset, adLockOptimistic
     
     If Not rs.EOF Then
     acccr = rs.Fields(0)
     accdr = rs.Fields(1)
     'txtredirect = rs.Fields(1)
     get_namecr acccr
     'get_namedr
     Lblname2 = lab
     get_namecr accdr
     lblname1 = lab
     End If
     
   '// populate the asset type in the TXTASSETTYPE
   Dim rsg As Recordset
   sql = ""
   sql = "select Assetname from assetcode"
   Set rsg = New ADODB.Recordset
   rsg.Open sql, cn, adOpenKeyset, adLockOptimistic
   cboassetype.Clear
   While Not rsg.EOF
   cboassetype.AddItem rsg.Fields(0)
   rsg.MoveNext
   Wend
   Exit Sub
ErrorHandler:
   MsgBox err.description, vbCritical, "EASYMA"
     
End Sub
Private Sub get_namecr(acc As String)

    Set cn = CreateObject("adodb.connection")
    Set myclass = New cdbase
    Provider = myclass.OpenCon
    cn.Open Provider, "atm"
    sql = ""
    sql = "select * from glsetup where accno='" & acc & "'"
    Set rs = New ADODB.Recordset
    rs.Open sql, cn, adOpenKeyset, adLockOptimistic
    If rs.EOF Then
    Else
    If Not IsNull(rs.Fields("glaccname")) Then lab = rs.Fields("glaccname")
    End If
End Sub

Private Sub txtassetcode_Change()
Dim lis As ListItem


Dim cn As Connection
Set rs = New Recordset
Set cn = New Connection
cn.Open frmODBCLogon.cboDSNList, "atm"


sql = "Select * from assets where assetsno='" & txtassetcode & "'"
rs.Open sql, cn
If Not rs.EOF Then
 If Not IsNull(rs.Fields("assetsname")) Then txtASSETSNAME = rs.Fields("assetsname")
 If Not IsNull(rs.Fields("assetserialno")) Then txtSERIALNO = rs.Fields("assetserialno")
 If Not IsNull(rs.Fields("purchaseprice")) Then purchase = rs.Fields("purchaseprice")
 If Not IsNull(rs.Fields("Depreciation")) Then txtValue = rs.Fields("Depreciation")
End If

End Sub

Private Sub txttransdate_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
cboyearperiod = Format(Get_Server_Date, "MM") & "-" & Format(Get_Server_Date, "YYYY")

End Sub
Private Sub txttransdate_Change()
cboyearperiod = Format(txttransdate, "MM") & "-" & Format(txttransdate, "YYYY")
End Sub
