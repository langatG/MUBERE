VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmcustomers 
   Caption         =   "CUSTOMER REGISTRATION"
   ClientHeight    =   4710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8145
   LinkTopic       =   "Form1"
   ScaleHeight     =   4710
   ScaleWidth      =   8145
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtpcode 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1560
      TabIndex        =   17
      Top             =   360
      Width           =   1935
   End
   Begin VB.TextBox txtpname 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1560
      TabIndex        =   16
      Top             =   960
      Width           =   4335
   End
   Begin VB.TextBox txtserialno 
      Appearance      =   0  'Flat
      Height          =   615
      Left            =   1680
      MultiLine       =   -1  'True
      TabIndex        =   15
      Top             =   2280
      Width           =   4335
   End
   Begin VB.TextBox txtquantity 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1560
      TabIndex        =   14
      Top             =   3120
      Width           =   1935
   End
   Begin VB.CommandButton cmdnew 
      Caption         =   "New"
      Height          =   375
      Left            =   1320
      TabIndex        =   13
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   375
      Left            =   2520
      TabIndex        =   12
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      Height          =   375
      Left            =   4920
      TabIndex        =   11
      Top             =   3840
      Width           =   975
   End
   Begin VB.TextBox txtbalance 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   6105
      TabIndex        =   9
      Top             =   3120
      Width           =   1815
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      DrawStyle       =   2  'Dot
      DrawWidth       =   17015
      Height          =   360
      Left            =   3480
      Picture         =   "frmcustomers.frx":0000
      ScaleHeight     =   360
      ScaleWidth      =   240
      TabIndex        =   8
      Top             =   360
      Width           =   240
   End
   Begin VB.CommandButton cmddelete 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   3720
      TabIndex        =   7
      Top             =   3840
      Width           =   975
   End
   Begin VB.ComboBox cbosupplier 
      Height          =   315
      ItemData        =   "frmcustomers.frx":0182
      Left            =   1680
      List            =   "frmcustomers.frx":0192
      TabIndex        =   6
      Top             =   1800
      Width           =   4455
   End
   Begin VB.CheckBox CHKSERIALIZED 
      Caption         =   "Serialized"
      Height          =   255
      Left            =   6480
      TabIndex        =   5
      Top             =   1800
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox chkserialrequired 
      Caption         =   "Serial "
      Height          =   375
      Left            =   6000
      TabIndex        =   4
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Frame fra1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   4200
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   2895
      Begin VB.TextBox txtpassword 
         Appearance      =   0  'Flat
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   240
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label8 
         Caption         =   "Enter Password"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.TextBox txtpassit 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   3600
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   3240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker txtdateenterered 
      Height          =   375
      Left            =   5400
      TabIndex        =   10
      Top             =   480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   130416641
      CurrentDate     =   38814
   End
   Begin VB.Label Label1 
      Caption         =   "Customer No."
      Height          =   255
      Left            =   0
      TabIndex        =   24
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Customer Name"
      Height          =   255
      Left            =   0
      TabIndex        =   23
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Reference"
      Height          =   255
      Left            =   0
      TabIndex        =   22
      Top             =   2505
      Width           =   1800
   End
   Begin VB.Label Label4 
      Caption         =   "Openning Balances"
      Height          =   255
      Left            =   0
      TabIndex        =   21
      Top             =   3120
      Width           =   1500
   End
   Begin VB.Label Label5 
      Caption         =   "Current Balance"
      Height          =   255
      Left            =   6105
      TabIndex        =   20
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Date Entered"
      Height          =   255
      Left            =   4200
      TabIndex        =   19
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "Town"
      Height          =   255
      Left            =   0
      TabIndex        =   18
      Top             =   1800
      Width           =   600
   End
End
Attribute VB_Name = "frmcustomers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Provider As String
Dim seria As Integer
Private Sub CHKSERIALIZED_Click()
If CHKSERIALIZED = vbChecked Then
frmserialization.Show vbModal
End If
End Sub

Private Sub chkserialrequired_Click()
If chkserialrequired = vbChecked Then
txtSERIALNO.Visible = True
seria = 1
Else
seria = 0
txtSERIALNO.Visible = False
End If
End Sub

'Public sel As String
Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmddelete_Click()
txtpassit.Visible = True
If txtpassit = "" Then
MsgBox "Please enter Password on the text above", vbInformation
Exit Sub
End If
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
pass = modsecurity.Encript_String(txtpassit)
If pass = rsp.Fields("password") Then
txtpassit.Visible = False
Else
MsgBox "You are not allowed to delete  . Consult administrator only", vbInformation
Exit Sub

End If
Else
MsgBox "You are not allowed to delete . Consult administrator only", vbInformation
Exit Sub

End If
'*****************************************
Set rst = New Recordset
Dim bo As Boolean
'Dim cn As Connection

Provider = cn
Set cn = New ADODB.Connection
cn.Open Provider, "atm", "atm"

sql = ""
sql = "delete from ag_products1 where p_code='" & txtpcode & "'"
cn.Execute sql

'// delete all the details in the stock balance

sql = ""
sql = "select * from ag_stockbalance1 where p_code='" & txtpcode & "' order by trackid"
Set rst = New ADODB.Recordset
rst.Open sql, cn, adOpenKeyset, adLockOptimistic

If Not rst.EOF Then
While Not rst.EOF
sql = ""
sql = "delete from ag_stockbalance1 where trackid=" & rst.Fields("trackid") & ""
cn.Execute sql

rst.MoveNext
Wend
End If

MsgBox "You have successfully deleted product code", vbInformation
txtbalance = ""
txtpcode = ""
txtpname = ""
txtSERIALNO = ""
txtquantity = ""

End Sub

Private Sub cmdNew_Click()
txtbalance = ""
txtpcode = ""
txtpname = ""
txtquantity = ""
txtSERIALNO = ""
End Sub

Private Sub cmdsave_Click()
Set rst = New Recordset
'
Dim unsera As Integer
'Dim cn As Connection
If txtquantity = "" Then
MsgBox "Quantity cannot be Zero", vbInformation
Exit Sub

End If
If txtbalance = "" Then txtbalance = 0
If chkserialrequired = vbChecked Then

seria = 1
unsera = txtquantity

'// should only be one item
If txtquantity > 1 Then
MsgBox "Serialized items should only be added as one", vbCritical
Exit Sub
End If
Else
seria = 0
unsera = 0
End If

Provider = cn
Set cn = New ADODB.Connection
cn.Open Provider, "atm", "atm"
'If reportpath = "" Then reportpath = GetSetting("payroll", "AppName", "rptPath", rptPath)
sql = "select P_CODE,qout,unserialized from ag_products1 where p_code='" & txtpcode & "'"
Set rs = New ADODB.Recordset
rs.Open sql, cn
If rs.EOF Then
'// insert into ag_products
sql = ""
sql = "set dateformat dmy insert into  ag_products1(p_code,p_name,s_no,qin,qout,date_entered,last_d_updated,user_id,audit_date,o_bal,supplierid,serialized,unserialized,seria )"
sql = sql & "  values('" & txtpcode.Text & "','" & txtpname.Text & "','" & txtSERIALNO.Text & "'," & txtquantity.Text & "," & txtbalance.Text + txtquantity.Text & ",'" & txtdateenterered.value & "','" & txtdateenterered.value & "','Admin','" & Date & "'," & txtquantity.Text & ",'" & cbosupplier & "',0," & unsera & "," & seria & ")"
cn.Execute sql


sql = ""
sql = "set dateformat DMY INSERT INTO ag_stockbalance1"
sql = sql & " (p_code, productname, openningstock, changeinstock, ag_stockbalance, transdate,companyid,s_no)"
sql = sql & " VALUES     ('" & txtpcode.Text & "','" & txtpname & "', " & txtbalance & ", " & txtquantity & ", " & txtbalance.Text + txtquantity.Text & ", '" & txtdateenterered & "',1,'" & txtSERIALNO & "')"
cn.Execute sql



Else
Dim D As Double
If Not IsNull(rs.Fields(2)) Then D = rs.Fields(2)
sql = ""
sql = "set dateformat DMY update ag_products1 set qin=" & txtquantity.Text & ",qout=" & txtquantity.Text + rs.Fields("qout") & ",o_bal=" & txtquantity.Text + rs.Fields("qout") & ",last_d_updated='" & Date & "',user_id='admin',audit_date='" & Date & "',unserialized=" & unsera + D & ",SERIA=" & seria & " where p_code='" & txtpcode.Text & "'"
cn.Execute sql

Dim rsst As Recordset
sql = ""
sql = "set dateformat DMY select top 1 * from ag_stockbalance1 where p_code='" & txtpcode & "' order by trackid desc "
Set rsst = New ADODB.Recordset
rsst.Open sql, cn
If Not rsst.EOF Then
sql = ""
sql = "set dateformat DMY INSERT INTO ag_stockbalance1"
sql = sql & " (p_code, productname, openningstock, changeinstock, ag_stockbalance, transdate,companyid,s_no)"
sql = sql & " VALUES     ('" & txtpcode & "', '" & txtpname & "', '" & txtbalance & "', '" & txtquantity & "', '" & txtquantity.Text + rs.Fields("qout") & "', '" & txtdateenterered & "',1,'" & txtSERIALNO & "')"
cn.Execute sql
'Else
'sql = "Update ag_stockbalance"
'sql = sql & " SET              productname = '" & txtpname & "', openningstock = " & txtbalance & ", changeinstock = " & txtquantity & ", ag_stockbalance = " & txtquantity.Text + rs.Fields("qout") & ", transdate = '" & txtdateenterered & "'"
'sql = sql & " WHERE     (p_code = '" & txtpcode & "') AND trackid=" & rsst.Fields("trackid") & ""
'cn.Execute sql
End If
'// update serialno database


End If
If seria = 1 Then
Set rst = Nothing
    sql = ""
   sql = "select * from serialno where serialno='" & txtSERIALNO & "' AND P_CODE='" & txtpcode & "' and used=0"
   Set rst = New ADODB.Recordset
   rst.Open sql, cn, adOpenKeyset, adLockOptimistic

If rst.EOF Then
sql = ""
sql = "set dateformat DMY INSERT INTO serialno(serialno,p_code,used)"
sql = sql & " values('" & txtSERIALNO & "','" & txtpcode & "',0)"
cn.Execute sql
Else
MsgBox "Item is in place and not yet used", vbInformation
Exit Sub
End If
End If
txtbalance = ""
txtpcode = ""
txtpname = ""
txtSERIALNO = ""
txtquantity = ""
End Sub

Private Sub Form_Load()
txtdateenterered = Format(Date, "dd,mm,yyyy")
 Set rst = New Recordset
    Dim cn As Connection
    Set cn = New ADODB.Connection
    Provider = "MAZIWA"
   cn.Open Provider, "atm", "atm"
    Set rst = New Recordset
    sql = "Select companyname from AG_Supplier1"
    rst.Open sql, cn, adOpenKeyset, adLockOptimistic
    While Not rst.EOF
    cbosupplier.AddItem rst.Fields(0)
    rst.MoveNext
    Wend
    
End Sub

Private Sub Picture2_Click()
frmsearchcustomers.Show vbModal
Dim Y As String
Y = sel

If Y <> "" Then
Provider = cn
Set cn = New ADODB.Connection
cn.Open Provider, "atm", "atm"
'If reportpath = "" Then reportpath = GetSetting("payroll", "AppName", "rptPath", rptPath)
sql = "select P_CODE,P_NAME,S_NO,QOUT from ag_products1 where p_code='" & Y & "'"
Set rs = New ADODB.Recordset
rs.Open sql, cn
If Not rs.EOF Then
If Not IsNull(rs.Fields(0)) Then txtpcode = (rs.Fields(0))
If Not IsNull(rs.Fields(1)) Then txtpname = (rs.Fields(1))
'If Not IsNull(rs.Fields(2)) Then txtserialno = (rs.Fields(2))
If Not IsNull(rs.Fields(3)) Then txtbalance = (rs.Fields(3))


'// check with serial no if it exist


End If
End If
End Sub



Private Sub txtdateenterered_Click()
fra1.Visible = True
End Sub

Private Sub txtdateenterered_KeyPress(KeyAscii As Integer)
fra1.Visible = True
End Sub

Private Sub txtdateenterered_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
fra1.Visible = True
End Sub

Private Sub txtpassword_LostFocus()
'// verufy where you have admin right to change the date
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
txtdateenterered = Date
End If
Else
MsgBox "You are not allowed to change the date . Consult administrator only", vbInformation
Exit Sub
txtdateenterered = Date
fra1.Visible = True
End If


End Sub

Private Sub txtpassword_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
txtpassword_LostFocus
End Sub

Private Sub txtpcode_Change()
'//TWNG001
Provider = cn
Set cn = New ADODB.Connection
cn.Open Provider, "atm", "atm"
'If reportpath = "" Then reportpath = GetSetting("payroll", "AppName", "rptPath", rptPath)
sql = "select P_CODE,P_NAME,S_NO,QOUT from ag_products1 where p_code='" & txtpcode & "'"
Set rs = New ADODB.Recordset
rs.Open sql, cn
If Not rs.EOF Then
 txtpcode = (rs.Fields(0))
If Not IsNull(rs.Fields(1)) Then txtpname = (rs.Fields(1))
'If Not IsNull(rs.Fields(2)) Then txtserialno = (rs.Fields(2))
If Not IsNull(rs.Fields(3)) Then txtbalance = (rs.Fields(3))

If txtbalance <= 0 Then
MsgBox "Your stock is below zero please reorder", vbInformation
End If
End If
'// check with serial no if it exist
End Sub

Private Sub txtquantity_Validate(Cancel As Boolean)
If Not IsNumeric(txtquantity) Then
MsgBox "Enter values please", vbCritical
txtquantity = ""
txtquantity.SetFocus
Exit Sub
End If
End Sub

Private Sub txtserialno_Change()
If Not IsNumeric(txtSERIALNO) Then
'MsgBox "Enter values please", vbCritical
txtSERIALNO = ""
txtSERIALNO.SetFocus
Exit Sub
End If

End Sub
