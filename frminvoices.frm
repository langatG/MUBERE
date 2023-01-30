VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frminvoices 
   Caption         =   "INVOICES PAYMENT"
   ClientHeight    =   4260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8610
   LinkTopic       =   "Form1"
   ScaleHeight     =   4260
   ScaleWidth      =   8610
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtrno 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1425
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
      Left            =   1440
      TabIndex        =   12
      Top             =   5280
      Width           =   1935
   End
   Begin VB.TextBox txtserialno 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1440
      TabIndex        =   11
      Top             =   1680
      Width           =   4215
   End
   Begin VB.TextBox txtquantity 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1440
      TabIndex        =   10
      Top             =   2160
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
      Left            =   1320
      TabIndex        =   8
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   375
      Left            =   2880
      TabIndex        =   7
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      Height          =   375
      Left            =   4440
      TabIndex        =   6
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdnextitem 
      Caption         =   "Next item"
      Height          =   375
      Left            =   5640
      TabIndex        =   5
      Top             =   2640
      Width           =   1455
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      DrawStyle       =   2  'Dot
      DrawWidth       =   17015
      Height          =   360
      Left            =   3960
      Picture         =   "frminvoices.frx":0000
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
      Picture         =   "frminvoices.frx":0182
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
      Format          =   65273857
      CurrentDate     =   38814
   End
   Begin VB.Label Label2 
      Caption         =   "Invoice No."
      Height          =   255
      Left            =   0
      TabIndex        =   24
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Customer Name"
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
      Caption         =   "Amount2"
      Height          =   255
      Left            =   0
      TabIndex        =   21
      Top             =   5280
      Width           =   855
   End
   Begin VB.Label lblserialno 
      Caption         =   "Other details"
      Height          =   255
      Left            =   0
      TabIndex        =   20
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "Amount1"
      Height          =   255
      Left            =   -15
      TabIndex        =   19
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "Customer No."
      Height          =   255
      Left            =   -15
      TabIndex        =   18
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label lblbalance 
      Height          =   255
      Left            =   6600
      TabIndex        =   17
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label9 
      Caption         =   "Balance"
      Height          =   255
      Left            =   4320
      TabIndex        =   16
      Top             =   2160
      Width           =   1335
   End
End
Attribute VB_Name = "frminvoices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim provider As String
Dim SelectedDsn As String
Dim DIA
Private Sub cboproductname_Change()
provider = cnn
Set cn = New ADODB.connection
Dim p As Integer
cn.Open provider, , "pius12"
'If reportpath = "" Then reportpath = GetSetting("payroll", "AppName", "rptPath", rptPath)
'Dim rst As New ADODB.Recordset

sql = ""
sql = "select P_CODE,qout,seria,s_no from products1 where p_name='" & cboproductname & "'"
Set rs = New ADODB.Recordset
rs.Open sql, cn
If Not rs.EOF Then
txtpcode = rs.Fields(0)
lblbalance = rs.Fields(1)
If Not IsNull(rs.Fields(2)) Then p = (rs.Fields(2))
If p = 1 Then
'//CHECK THE ITEM IN THE SERIALS
Dim RSSE As Recordset
sql = ""
sql = "select top 1 serialno,p_code,used from serialno where p_code='" & txtpcode & "' AND USED=0"
Set RSSE = New ADODB.Recordset

RSSE.Open sql, cn, adOpenKeyset, adLockOptimistic
If Not RSSE.EOF Then

If Not IsNull(RSSE.Fields(0)) Then txtserialno = (RSSE.Fields(0))
lblserialno.Visible = True
txtserialno.Visible = True

End If
Else
lblserialno.Visible = False
txtserialno.Visible = False
End If
End If
End Sub

Private Sub cboproductname_Click()
cboproductname_Change
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdnew_Click()
txtrno = ""
 txtpcode = ""
 txtserialno = ""
 txtquantity = ""
 txtamount = ""
End Sub

Private Sub cmdnextitem_Click()
txtpcode = ""
txtquantity = ""
txtserialno = ""

End Sub

Private Sub cmdsave_Click()
Set rst = New Recordset
Dim A, b, X
DIA = 0
Dim U As Double, S As Double
Dim cn As connection
txtserialno_LostFocus
If DIA = 1 Then
Exit Sub
End If
'Dim cn As Connection
If txtquantity = "" Then
MsgBox "Quantity cannot be Zero", vbInformation
Exit Sub
End If
    If txtpcode = "" Then
        MsgBox "Please Enter the Product CODE before You Proceed!", vbCritical
        Exit Sub
    End If
    If txtrno = "" Then
        MsgBox "Please Enter Receipt Number before you Proceed!", vbCritical
        Exit Sub
    End If
    
If txtamount = "" Then
txtamount = 0
End If
provider = cnn
Set cn = New ADODB.connection
cn.Open provider, , "pius12"
'// check if they are in stock.
Dim rsinstock As Recordset
sql = ""
sql = "select P_CODE,qout from products1 where p_code='" & txtpcode & "'"
Set rsinstock = New ADODB.Recordset
rsinstock.Open sql, cn
'// check the stock if it is less than zero
'If rsinstock.Fields(1) <= 0 Then
'MsgBox "Sorry Stock is Zero please re-stock before your proceed", vbInformation
'Exit Sub
'End If
''// check the quanttity being sold versus the balance
'Dim piu As Currency
'piu = rsinstock.Fields(1) - CCur(txtquantity)
'If piu < 0 Then
'MsgBox "Stock will be negative please re-stock before you proceed", vbInformation
'Exit Sub
'End If

b = txtquantity

'If reportpath = "" Then reportpath = GetSetting("payroll", "AppName", "rptPath", rptPath)
'Dim rst As New ADODB.Recordset
sql = ""
sql = "set dateformat DMY select * from receipts1 where r_no=" & txtrno & " and T_Date='" & txtransdate & "' and p_code='" & txtpcode & "'"
Set rst = New ADODB.Recordset
rst.Open sql, cn
If Not rst.EOF Then
'MsgBox "You Cannot sell the same items using the same receipts number on the same date", vbInformation
'Exit Sub

End If

If rst.EOF Then
sql = ""
sql = "select P_CODE,qout from products1 where p_code='" & txtpcode & "'"
Set rs = New ADODB.Recordset
rs.Open sql, cn
If Not rs.EOF Then
'// insert into products
sql = ""
sql = "set dateformat DMY insert into receipts1 (r_no,p_code,t_date,amount,s_no,qua,s_bal)"
sql = sql & "values(" & txtrno & ",'" & txtpcode & "','" & txtransdate & "'," & txtamount & ",'" & txtserialno & "'," & txtquantity & "," & rs.Fields("qout") - txtquantity & ") "
cn.Execute sql
'// update the product table
sql = ""
sql = "set dateformat DMY update Products1 set qout=" & rs.Fields("qout") - txtquantity & ",last_d_updated='" & Date & "',user_id='admin',audit_date='" & Date & "' where p_code='" & txtpcode.Text & "'"
cn.Execute sql

sql = "Update serialno"
sql = sql & "  Set used=1 "
sql = sql & "  WHERE  serialno = '" & txtserialno & "'"
cn.Execute sql


' txtserialno = ""
' txtquantity = ""
' txtamount = ""
Else
'//////////////////////////////////////////////////
'// check if the one is seria and also if it has the same serial no then give a message

    
        Set rst = Nothing
        sql = ""
        sql = "select * from receipts1 where r_no=" & txtrno & ""
        Set rst = New ADODB.Recordset
        rst.Open sql, cn
        
        sql = ""
        sql = "update Products1 set qin=" & txtquantity.Text & ",qout=" & txtquantity.Text + rs.Fields("qout") & ",last_d_updated='" & Date & "',user_id='admin',audit_date='" & Date & "' where p_code='" & txtpcode.Text & "'"
        cn.Execute sql
        
        sql = "Update serialno"
        sql = sql & "  Set used=1 "
        sql = sql & "  WHERE  serialno = '" & txtserialno & "'"
        cn.Execute sql
    

End If
'////////////////////////////
Else
'//IT SHOULD BE HERE

Dim serialised

'// modify the receipts

sql = ""
sql = "select P_CODE,qout,SERIA from products1 where p_code='" & txtpcode & "'"
Set rs = New ADODB.Recordset
rs.Open sql, cn
If rs.Fields(2) = 1 Then serialised = 1 Else serialised = 0
'//WE SHOULD CHECK IF THE SAME RECEIPTS NO IS IN.
If serialised = 1 Then
If Not rs.EOF Then
'// CHECK IF THE SERIAL NUMBER IS ALREADY USED

sql = ""
Dim RSSE As Recordset
sql = ""
sql = "select top 1 serialno,p_code,used from serialno where p_code='" & txtpcode & "' and serialno='" & txtserialno & "' order by serialid desc"
Set RSSE = New ADODB.Recordset

RSSE.Open sql, cn, adOpenKeyset, adLockOptimistic
If RSSE.Fields(2) = 1 Then
MsgBox "Serial Number and receipt no used please check again before posting", vbCritical
Exit Sub
End If
sql = ""
sql = "insert into receipts1 (r_no,p_code,t_date,amount,s_no,qua,s_bal)"
sql = sql & "values(" & txtrno & ",'" & txtpcode & "','" & txtransdate & "'," & txtamount & "," & txtserialno & "," & txtquantity & "," & rs.Fields("qout") - txtquantity & ") "
cn.Execute sql


'// update the product table
sql = ""
sql = "update Products1 set qout=" & rs.Fields("qout") - txtquantity & ",last_d_updated='" & Date & "',user_id='admin',audit_date='" & Date & "' where p_code='" & txtpcode.Text & "'"
cn.Execute sql


sql = "Update serialno"
sql = sql & "  Set used=1 "
sql = sql & "  WHERE  serialno = '" & txtserialno & "'"
cn.Execute sql


'txtserialno = ""
' txtquantity = ""
' txtamount = ""
 End If
Else

If Not rs.EOF Then

sql = ""
sql = "insert into receipts1 (r_no,p_code,t_date,amount,s_no,qua,s_bal)"
sql = sql & "values(" & txtrno & ",'" & txtpcode & "','" & txtransdate & "'," & txtamount & "," & txtserialno & "," & txtquantity & "," & rs.Fields("qout") - txtquantity & ") "
cn.Execute sql


'// update the product table
sql = ""
sql = "update Products1 set qout=" & rs.Fields("qout") - txtquantity & ",last_d_updated='" & Date & "',user_id='admin',audit_date='" & Date & "' where p_code='" & txtpcode.Text & "'"
cn.Execute sql


sql = "Update serialno"
sql = sql & "  Set used=1 "
sql = sql & "  WHERE  serialno = '" & txtserialno & "'"
cn.Execute sql


'txtserialno = ""
'txtquantity = ""
' txtamount = ""
 End If
End If
End If
'// update thes stock balance
Dim h As String
h = CStr("cboname")
Dim rsst As Recordset
Dim stbal As Double
sql = ""
sql = "select top 1 * from stockbalance1 where p_code='" & txtpcode & "' order by trackid desc "
Set rsst = New ADODB.Recordset
rsst.Open sql, cn
If Not rsst.EOF Then
stbal = rsst.Fields("stockbalance")
stbal = stbal - b
b = (b * -1)

sql = ""
sql = "set dateformat DMY INSERT INTO stockbalance1"
sql = sql & " (p_code, productname, openningstock, changeinstock, stockbalance, transdate,companyid,r_no,s_no)"
sql = sql & " VALUES    ('" & txtpcode.Text & "', '" & h & "', " & rsst.Fields("stockbalance") & ", " & b & ", " & stbal & ", '" & txtransdate & "',1,'" & txtrno & "'," & txtrno & ")"
cn.Execute sql
Else
stbal = b
b = (b * -1)

sql = ""
sql = "set dateformat DMY INSERT INTO stockbalance1"
sql = sql & " (p_code, productname, openningstock, changeinstock, stockbalance, transdate,companyid,r_no,s_no)"
sql = sql & " VALUES    ('" & txtpcode.Text & "', '" & h & "', " & b & ", " & b & ", " & stbal & ", '" & txtransdate & "',1,'" & txtrno & "'," & txtrno & ")"
cn.Execute sql
End If
txtrno.Text = ""
txtpcode.Text = ""
txtserialno = ""
txtquantity = ""
 txtamount = ""
HEREEE:

End Sub

Private Sub Form_Load()
txtransdate = Format(Date, "dd/mm/yyyy")
provider = cnn
Set cn = New ADODB.connection
cn.Open provider, , "pius12"
'If reportpath = "" Then reportpath = GetSetting("payroll", "AppName", "rptPath", rptPath)
sql = "select P_NAME  from products1 ORDER BY P_NAME ASC"
Set rs = New ADODB.Recordset
rs.Open sql, cn

While Not rs.EOF
cboproductname.AddItem rs.Fields(0)
rs.MoveNext
Wend
End Sub
Private Sub cboname()
provider = cnn
Set cn = New ADODB.connection
cn.Open provider, , "pius12"
'If reportpath = "" Then reportpath = GetSetting("payroll", "AppName", "rptPath", rptPath)
sql = "select P_NAME from products1 where p_code='" & txtpcode & "'"
Set rs = New ADODB.Recordset
rs.Open sql, cn
If Not rs.EOF Then
If Not IsNull(rs.Fields(0)) Then cboproductname = (rs.Fields(0))
If Not IsNull(rs.Fields(1)) Then lblbalance = rs.Fields(1)
End If
End Sub

Private Sub Picture1_Click()
frmsearchcustomers.Show vbModal
Dim Y As String
Y = Sel
Dim p As Integer
If Y <> "" Then
provider = cnn
Set cn = New ADODB.connection
cn.Open provider, , "pius12"
'If reportpath = "" Then reportpath = GetSetting("payroll", "AppName", "rptPath", rptPath)
sql = "select P_CODE,P_NAME,S_NO,QOUT,seria,s_no from products1 where p_code='" & Y & "'"
Set rs = New ADODB.Recordset
rs.Open sql, cn
If Not rs.EOF Then
If Not IsNull(rs.Fields(0)) Then txtpcode = (rs.Fields(0))
If Not IsNull(rs.Fields(4)) Then p = (rs.Fields(4))
If p = 1 Then
If Not IsNull(rs.Fields(5)) Then txtserialno = (rs.Fields(5))
lblserialno.Visible = True
txtserialno.Visible = True
Else
lblserialno.Visible = False
txtserialno.Visible = False
End If
If Not IsNull(rs.Fields(1)) Then cboproductname = (rs.Fields(1))

'If Not IsNull(rs.Fields(3)) Then lblbalance = (rs.Fields(3))
'// check if it has the serial numbers
'get_serialno Y

End If
End If

'// check if the product have the serial then show the receipts details



End Sub
Private Sub get_serialno(pcode As String)
provider = cnn
Set cn = New ADODB.connection
cn.Open provider, , "pius12"
Dim RSSE As Recordset
sql = ""
sql = "select top 1 serialno,p_code,used from serialno where p_code='" & txtpcode & "' and serialno='" & txtserialno & "' order by serialid desc"
Set RSSE = New ADODB.Recordset

RSSE.Open sql, cn, adOpenKeyset, adLockOptimistic
If RSSE.Fields(2) = 1 Then
MsgBox "Serial Number and receipt no used please check again before posting", vbCritical
Exit Sub
End If
End Sub
Private Sub Picture2_Click()
On Error Resume Next
frmsearchcustomerpayments.Show vbModal
Dim Y As String
Y = Sel

If Y <> "" Then
provider = cnn
Set cn = New ADODB.connection
cn.Open provider, , "pius12"
sql = ""
'If reportpath = "" Then reportpath = GetSetting("payroll", "AppName", "rptPath", rptPath)
sql = "select r_no,P_CODE,S_NO,Qua,amount from receipts1 where r_id=" & Y & ""
Set rs = New ADODB.Recordset
rs.Open sql, cn
If Not rs.EOF Then
If Not IsNull(rs.Fields(0)) Then txtrno = (rs.Fields(0))
If Not IsNull(rs.Fields(1)) Then txtpcode = (rs.Fields(1))
If Not IsNull(rs.Fields(2)) Then txtserialno = (rs.Fields(2))
If Not IsNull(rs.Fields(3)) Then txtquantity = (rs.Fields(3))
If Not IsNull(rs.Fields(4)) Then txtamount = (rs.Fields(4))
If Not IsNull(rs.Fields(3)) Then lblbalance = (rs.Fields(3))
Call cboname
End If
End If
End Sub

Private Sub txtpassword_LostFocus()
'fra1.Visible = True
Dim rsp As Recordset
Set cn = CreateObject("adodb.connection")
provider = cnn
Set cn = New ADODB.connection
cn.Open provider, , "pius12"
Set rsp = CreateObject("adodb.recordset")
sql = "select *  from useraccounts where UserLoginID='" & User & "' and usergroup='administrator'"
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

Private Sub txtquantity_Validate(Cancel As Boolean)
If Not IsNumeric(txtquantity) Then
MsgBox "Please enter a value please", vbCritical
txtquantity = ""
txtquantity.SetFocus
Exit Sub
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
sql = "select * from products1 where seria=1 AND P_CODE='" & txtpcode & "'"
Set rsproduct = New ADODB.Recordset
rsproduct.Open sql, cn, adOpenKeyset, adLockOptimistic
If Not rsproduct.EOF Then
sql = ""
sql = "select serialno  from serialno where serialno= '" & txtserialno & "'"
Set rss = New ADODB.Recordset
rss.Open sql, cn, adOpenKeyset, adLockOptimistic
If Not rss.EOF Then
'// check if gth
While Not rss.EOF
Dim ser As String
ser = rss.Fields(0)

If ser = txtserialno Then GoTo hererere

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

Private Sub txtrno_Validate(Cancel As Boolean)
    If Not IsNumeric(txtrno) Then
        MsgBox "Enter NUMERIC VALUES only", vbCritical
        txtrno = ""
        txtrno.SetFocus
        Exit Sub
    End If
End Sub

