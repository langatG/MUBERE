VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmserializationallocations 
   Caption         =   "SERIALIZATION ALLOCATION"
   ClientHeight    =   4545
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8310
   Icon            =   "frmserializationallocations.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4545
   ScaleWidth      =   8310
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cbosupplier 
      Height          =   315
      Left            =   1440
      TabIndex        =   25
      Top             =   1800
      Width           =   4335
   End
   Begin VB.TextBox txtrno 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1440
      TabIndex        =   15
      Top             =   240
      Width           =   2535
   End
   Begin VB.ComboBox cboproductname 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   1200
      Width           =   4335
   End
   Begin VB.TextBox txtamount 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1560
      TabIndex        =   12
      Top             =   3480
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox txtserialno 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1560
      TabIndex        =   11
      Top             =   2520
      Width           =   4215
   End
   Begin VB.TextBox txtquantity 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "1"
      Top             =   3000
      Width           =   1935
   End
   Begin VB.TextBox txtpcode 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1440
      TabIndex        =   9
      Top             =   720
      Width           =   2535
   End
   Begin VB.CommandButton cmdnew 
      Caption         =   "New"
      Height          =   375
      Left            =   1440
      TabIndex        =   8
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   375
      Left            =   3000
      TabIndex        =   7
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      Height          =   375
      Left            =   4560
      TabIndex        =   6
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdnextitem 
      Caption         =   "Next item"
      Height          =   375
      Left            =   5760
      TabIndex        =   5
      Top             =   3480
      Width           =   1455
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      DrawStyle       =   2  'Dot
      DrawWidth       =   17015
      Height          =   360
      Left            =   3960
      Picture         =   "frmserializationallocations.frx":0442
      ScaleHeight     =   360
      ScaleWidth      =   240
      TabIndex        =   4
      Top             =   240
      Width           =   240
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      DrawStyle       =   2  'Dot
      DrawWidth       =   17015
      Height          =   360
      Left            =   3960
      Picture         =   "frmserializationallocations.frx":05C4
      ScaleHeight     =   360
      ScaleWidth      =   240
      TabIndex        =   3
      Top             =   720
      Width           =   240
   End
   Begin VB.Frame fra1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   5520
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2895
      Begin VB.TextBox txtpassword 
         Appearance      =   0  'Flat
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   240
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label8 
         Caption         =   "Enter Password"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   120
         Width           =   1215
      End
   End
   Begin MSComCtl2.DTPicker txtransdate 
      Height          =   375
      Left            =   6360
      TabIndex        =   13
      Top             =   480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   130416641
      CurrentDate     =   38814
   End
   Begin VB.Label lbltransdate 
      Height          =   255
      Left            =   6120
      TabIndex        =   28
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label10 
      Caption         =   "Transdate"
      Height          =   255
      Left            =   6120
      TabIndex        =   27
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Supplier"
      Height          =   255
      Left            =   0
      TabIndex        =   26
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Receipt No."
      Height          =   255
      Left            =   0
      TabIndex        =   24
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Product Name"
      Height          =   255
      Left            =   0
      TabIndex        =   23
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Trans_Date"
      Height          =   255
      Left            =   6360
      TabIndex        =   22
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Amount"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   3480
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblserialno 
      Caption         =   "Serial No."
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "Quantity"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "Product Code"
      Height          =   255
      Left            =   0
      TabIndex        =   18
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label lblbalance 
      Height          =   255
      Left            =   5880
      TabIndex        =   17
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label9 
      Caption         =   "Balance"
      Height          =   255
      Left            =   4440
      TabIndex        =   16
      Top             =   3000
      Width           =   1335
   End
End
Attribute VB_Name = "frmserializationallocations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Provider As String
Dim SelectedDsn As String
Dim DIA
Private Sub cboproductname_Change()
Provider = cn
Set cn = New ADODB.Connection
Dim p As Integer
cn.Open Provider, "atm", "atm"
'If reportpath = "" Then reportpath = GetSetting("payroll", "AppName", "rptPath", rptPath)
'Dim rst As New ADODB.Recordset

sql = ""
sql = "select P_CODE,qout,seria,s_no from ag_products where p_name='" & cboproductname & "'"
Set rs = New ADODB.Recordset
rs.Open sql, cn
If Not rs.EOF Then
txtpcode = rs.Fields(0)
lblbalance = rs.Fields(1)
If Not IsNull(rs.Fields(2)) Then p = (rs.Fields(2))
If p = 1 Then
'If Not IsNull(rs.Fields(3)) Then txtserialno = (rs.Fields(3))
'lblserialno.Visible = True
'txtserialno.Visible = True
'Else
'lblserialno.Visible = False
'txtserialno.Visible = False
End If
End If
End Sub

Private Sub cboproductname_Click()
cboproductname_Change
End Sub

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdNew_Click()
txtrno = ""
 txtpcode = ""
 txtSERIALNO = ""
 txtquantity = ""
 txtamount = ""
End Sub

Private Sub cmdnextitem_Click()
txtpcode = ""
txtquantity = ""
txtSERIALNO = ""

End Sub

Private Sub cmdsave_Click()
Set rst = New Recordset
Dim a, b, X
DIA = 0
Dim U As Double, S As Double
Dim cn As Connection

If txtquantity = "" Then
MsgBox "Quantity cannot be Zero", vbInformation
Exit Sub
End If
If txtamount = "" Then
txtamount = 0
End If
Provider = cn
Set cn = New ADODB.Connection
cn.Open Provider, "atm", "atm"



'If reportpath = "" Then reportpath = GetSetting("payroll", "AppName", "rptPath", rptPath)
'Dim rst As New ADODB.Recordset
sql = ""
sql = "SELECT sisold.ID, sisold.S_no, sisold.p_code, sisold.supplier, sisold.transdate, * FROM sisold  where s_no='" & txtSERIALNO & "' and r_no='" & txtrno & "'"
Set rst = New ADODB.Recordset
rst.Open sql, cn
If rst.EOF Then


'// insert into ag_products
sql = ""
sql = "insert into sisold (s_no,p_code,supplier,transdate,r_no)"
sql = sql & "values(" & txtSERIALNO & ",'" & txtpcode & "','" & cbosupplier & "','" & lblTransDate & "','" & txtrno & "') "
cn.Execute sql

MsgBox "You have successfully updated the serialised item"
Else

MsgBox "The serial number already exist please"

End If
'////////////////////////////

HEREEE:

End Sub

Private Sub Form_Load()
txtransdate = Format(Date, "dd/mm/yyyy")
Provider = cn
Set cn = New ADODB.Connection
cn.Open Provider, "atm", "atm"
'If reportpath = "" Then reportpath = GetSetting("payroll", "AppName", "rptPath", rptPath)
sql = "select P_NAME  from ag_products ORDER BY P_NAME ASC"
Set rs = New ADODB.Recordset
rs.Open sql, cn

While Not rs.EOF
cboproductname.AddItem rs.Fields(0)
rs.MoveNext
Wend
Set rs = Nothing
Set rst = New Recordset
  
    Set cn = New ADODB.Connection
    Provider = cn
   cn.Open Provider, "atm", "atm"
    Set rst = New Recordset
    sql = "Select companyname from Suppliers"
    rst.Open sql, cn, adOpenKeyset, adLockOptimistic
    While Not rst.EOF
    cbosupplier.AddItem rst.Fields(0)
    rst.MoveNext
    Wend
End Sub
Private Sub cboname()
Provider = cn
Set cn = New ADODB.Connection
cn.Open Provider, "atm", "atm"
'If reportpath = "" Then reportpath = GetSetting("payroll", "AppName", "rptPath", rptPath)
sql = "select P_NAME from ag_products where p_code='" & txtpcode & "'"
Set rs = New ADODB.Recordset
rs.Open sql, cn
If Not rs.EOF Then
If Not IsNull(rs.Fields(0)) Then cboproductname = (rs.Fields(0))
If Not IsNull(rs.Fields(1)) Then lblbalance = rs.Fields(1)
End If
End Sub

Private Sub Picture1_Click()
frmSearch.Show vbModal
Dim Y As String
Y = sel
Dim p As Integer
If Y <> "" Then
Provider = cn
Set cn = New ADODB.Connection
cn.Open Provider, "atm", "atm"
'If reportpath = "" Then reportpath = GetSetting("payroll", "AppName", "rptPath", rptPath)
sql = "select P_CODE,P_NAME,S_NO,QOUT,seria,s_no from ag_products where p_code='" & Y & "'"
Set rs = New ADODB.Recordset
rs.Open sql, cn
If Not rs.EOF Then
If Not IsNull(rs.Fields(0)) Then txtpcode = (rs.Fields(0))
If Not IsNull(rs.Fields(4)) Then p = (rs.Fields(4))
'If p = 1 Then
'If Not IsNull(rs.Fields(5)) Then txtserialno = (rs.Fields(5))
'lblserialno.Visible = True
'txtserialno.Visible = True
'Else
'lblserialno.Visible = False
'txtserialno.Visible = False
'End If
If Not IsNull(rs.Fields(1)) Then cboproductname = (rs.Fields(1))

'If Not IsNull(rs.Fields(3)) Then lblbalance = (rs.Fields(3))
'// check if it has the serial numbers
'get_serialno Y

End If
End If

'// check if the product have the serial then show the ag_receipts details



End Sub
Private Sub get_serialno(pcode As String)
Provider = cn
Set cn = New ADODB.Connection
cn.Open Provider, "atm", "atm"
Dim RSSE As Recordset
sql = ""
sql = "select top 1 serialno,p_code,used from serialno where p_code='" & txtpcode & "' and serialno='" & txtSERIALNO & "' order by serialid desc"
Set RSSE = New ADODB.Recordset

RSSE.Open sql, cn, adOpenKeyset, adLockOptimistic
If RSSE.Fields(2) = 1 Then
MsgBox "Serial Number and receipt no used please check again before posting", vbCritical
Exit Sub
End If
End Sub
Private Sub Picture2_Click()
On Error Resume Next
frmsearchre.Show vbModal
Dim Y As String
Y = sel

If Y <> "" Then
Provider = cn
Set cn = New ADODB.Connection
cn.Open Provider, "atm", "atm"
sql = ""
'If reportpath = "" Then reportpath = GetSetting("payroll", "AppName", "rptPath", rptPath)
sql = "select r_no,P_CODE,S_NO,Qua,amount,t_date from ag_receipts where r_no=" & Y & ""
Set rs = New ADODB.Recordset
rs.Open sql, cn
If Not rs.EOF Then
If Not IsNull(rs.Fields(0)) Then txtrno = (rs.Fields(0))
If Not IsNull(rs.Fields(1)) Then txtpcode = (rs.Fields(1))
If Not IsNull(rs.Fields(2)) Then txtSERIALNO = (rs.Fields(2))
If Not IsNull(rs.Fields(3)) Then txtquantity = (rs.Fields(3))
If Not IsNull(rs.Fields(4)) Then txtamount = (rs.Fields(4))
If Not IsNull(rs.Fields(3)) Then lblbalance = (rs.Fields(3))
If Not IsNull(rs.Fields(5)) Then lblTransDate = (rs.Fields(5))

Call cboname
End If
End If
End Sub

Private Sub txtpassword_LostFocus()
'fra1.Visible = True
Dim rsp As Recordset
Set cn = CreateObject("adodb.connection")
Provider = cn
Set cn = New ADODB.Connection
cn.Open Provider, "atm", "atm"
Set rsp = CreateObject("adodb.recordset")
sql = "select *  from useraccounts where UserLoginIDs='" & User & "' and usergroup='administrator'"
rsp.Open sql, cn
Dim pass As String

If Not rsp.EOF Then
pass = modsecurity.Encript_String(txtpassword)
If pass = rsp.Fields("password") Then
fra1.Visible = False
Else
MsgBox "You are not allowed to change the date . Consult administrator only", vbInformation
Exit Sub
txtransdate = Format(Date, "DD/MM/YYYY")
End If
Else
MsgBox "You are not allowed to change the date . Consult administrator only", vbInformation
Exit Sub
txtransdate = Format(Date, "DD/MM/YYYY")
fra1.Visible = True
End If


End Sub
Private Sub txtransdate_click()
fra1.Visible = True
End Sub

Private Sub txtransdate_KeyPress(KeyAscii As Integer)
fra1.Visible = True
End Sub

Private Sub txtransdate_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
fra1.Visible = True
End Sub
Private Sub txtpassword_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
txtpassword_LostFocus
End Sub

Private Sub txtpcode_LostFocus()
Call cboname
End Sub

Private Sub txtserialno_LostFocus()
Dim rss As ADODB.Recordset
Dim rsproduct As ADODB.Recordset
sql = ""
sql = "select * from ag_products where seria=1 AND P_CODE='" & txtpcode & "'"
Set rsproduct = New ADODB.Recordset
rsproduct.Open sql, cn, adOpenKeyset, adLockOptimistic
If Not rsproduct.EOF Then
sql = ""
sql = "select serialno  from serialno where serialno= '" & txtSERIALNO & "'"
Set rss = New ADODB.Recordset
rss.Open sql, cn, adOpenKeyset, adLockOptimistic
If Not rss.EOF Then
'// check if gth
While Not rss.EOF
Dim ser As String
ser = rss.Fields(0)

If ser = txtSERIALNO Then GoTo hererere

rss.MoveNext
Wend
Else
MsgBox "Serial no not in our database", vbInformation

DIA = 1
Exit Sub
End If
End If
hererere:
End Sub

