VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmupdateserialiseditems 
   Caption         =   "STOCK Updating Serialised Items."
   ClientHeight    =   3825
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7740
   Icon            =   "frmupdateserialiseditems.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3825
   ScaleWidth      =   7740
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtserialno 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      TabIndex        =   28
      Top             =   1800
      Width           =   3375
   End
   Begin VB.ComboBox txtserialid 
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   27
      Top             =   1440
      Width           =   3375
   End
   Begin VB.TextBox txtchangeserialnoto 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1560
      MultiLine       =   -1  'True
      TabIndex        =   24
      Top             =   2280
      Width           =   3375
   End
   Begin VB.TextBox txtpcode 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1560
      TabIndex        =   16
      Top             =   360
      Width           =   1935
   End
   Begin VB.TextBox txtpname 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1560
      TabIndex        =   15
      Top             =   960
      Width           =   4335
   End
   Begin VB.TextBox txtquantity 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1560
      TabIndex        =   14
      Text            =   "0"
      Top             =   2880
      Width           =   1935
   End
   Begin VB.CommandButton cmdnew 
      Caption         =   "New"
      Height          =   375
      Left            =   1320
      TabIndex        =   13
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   375
      Left            =   2520
      TabIndex        =   12
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      Height          =   375
      Left            =   4920
      TabIndex        =   11
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox txtbalance 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   5040
      TabIndex        =   9
      Top             =   2880
      Width           =   1815
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      DrawStyle       =   2  'Dot
      DrawWidth       =   17015
      Height          =   360
      Left            =   3480
      Picture         =   "frmupdateserialiseditems.frx":0442
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
      Top             =   3360
      Width           =   975
   End
   Begin VB.ComboBox cbosupplier 
      Height          =   315
      Left            =   1560
      TabIndex        =   6
      Top             =   7440
      Width           =   4455
   End
   Begin VB.CheckBox CHKSERIALIZED 
      Caption         =   "Serialized"
      Height          =   375
      Left            =   6240
      TabIndex        =   5
      Top             =   7320
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox chkserialrequired 
      Caption         =   "Serial Required"
      Height          =   375
      Left            =   6000
      TabIndex        =   4
      Top             =   6360
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
      Top             =   3000
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
      Format          =   66060289
      CurrentDate     =   38814
   End
   Begin VB.Label Label10 
      Caption         =   "Serial ID"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label9 
      Caption         =   "Change Serial No To"
      Height          =   255
      Left            =   0
      TabIndex        =   25
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Product Code"
      Height          =   255
      Left            =   0
      TabIndex        =   23
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Product Name"
      Height          =   255
      Left            =   0
      TabIndex        =   22
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Serial No"
      Height          =   255
      Left            =   0
      TabIndex        =   21
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Quantity"
      Height          =   255
      Left            =   0
      TabIndex        =   20
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Balance In Store"
      Height          =   255
      Left            =   5040
      TabIndex        =   19
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "Date Entered"
      Height          =   255
      Left            =   4200
      TabIndex        =   18
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "Supplier"
      Height          =   255
      Left            =   0
      TabIndex        =   17
      Top             =   7440
      Width           =   975
   End
End
Attribute VB_Name = "frmupdateserialiseditems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public provider As String
Dim seria As Integer
Dim serialid As Long
Private Sub CHKSERIALIZED_Click()
If CHKSERIALIZED = vbChecked Then
frmserialization.Show vbModal
End If
End Sub
Private Sub chkserialrequired_Click()
If chkserialrequired = vbChecked Then
txtserialno.Visible = True
seria = 1
Else
seria = 0
txtserialno.Visible = False
End If
End Sub
Private Sub cmdClose_Click()
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
provider = cnn
Set cn = New ADODB.connection
cn.Open provider, , "pius12"
Set rsp = CreateObject("adodb.recordset")
sql = "select *  from useraccounts where UserLoginID='" & User & "' and usergroup='administrator'"
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

provider = cnn
Set cn = New ADODB.connection
cn.Open provider, , "pius12"

sql = ""
sql = "delete from products where p_code='" & txtpcode & "'"
cn.Execute sql

'// delete all the details in the stock balance

sql = ""
sql = "select * from stockbalance where p_code='" & txtpcode & "' order by trackid"
Set rst = New ADODB.Recordset
rst.Open sql, cn, adOpenKeyset, adLockOptimistic

If Not rst.EOF Then
While Not rst.EOF
sql = ""
sql = "delete from stockbalance where trackid=" & rst.Fields("trackid") & ""
cn.Execute sql

rst.MoveNext
Wend
End If

MsgBox "You have successfully deleted product code", vbInformation
txtbalance = ""
txtpcode = ""
txtpname = ""
txtserialno = ""
txtquantity = ""

End Sub

Private Sub cmdnew_Click()
txtbalance = ""
txtpcode = ""
txtpname = ""
txtquantity = "0"
txtserialno = ""
End Sub

Private Sub cmdsave_Click()
Set rs = New Recordset



provider = cnn
Set cn = New ADODB.connection
cn.Open provider, , "pius12"
sql = "SELECT  serialid,serialno,p_code,used FROM serialno where serialid=" & txtserialid & " "
rs.Open sql, cn
If Not rs.EOF Then
'// insert into products
'If Not IsNull(rs.Fields(0)) Then serialid = rs.Fields(0)
sql = ""
sql = "update serialno set serialno='" & txtchangeserialnoto & "' where serialid=" & txtserialid & ""
cn.Execute sql

Else

MsgBox "The serial Number is either not valid or does not exist", vbInformation, "Updating Serial Numbers"
Exit Sub
End If

End Sub

Private Sub Form_Load()
txtdateenterered = Format(Date, "dd,mm,yyyy")
 Set rst = New Recordset
    Dim cn As connection
    Set cn = New ADODB.connection
    provider = cnn
    cn.Open provider
    Set rst = New Recordset
    sql = "Select companyname from Suppliers"
    rst.Open sql, cn, adOpenKeyset, adLockOptimistic
    While Not rst.EOF
    cbosupplier.AddItem rst.Fields(0)
    rst.MoveNext
    Wend
    
    '//
    
    
End Sub

Private Sub Picture2_Click()
frmSearch.Show vbModal
Dim Y As String
Y = Sel

If Y <> "" Then
provider = cnn
Set cn = New ADODB.connection
cn.Open provider, , "pius12"
'If reportpath = "" Then reportpath = GetSetting("payroll", "AppName", "rptPath", rptPath)
sql = "select P_CODE,P_NAME,S_NO,QOUT from products where p_code='" & Y & "'"
Set rs = New ADODB.Recordset
rs.Open sql, cn, adOpenKeyset, adLockOptimistic
If Not rs.EOF Then
If Not IsNull(rs.Fields(0)) Then txtpcode = (rs.Fields(0))
'If Not IsNull(rs.Fields(1)) Then txtpname.Text = (rs.Fields(1))
'If Not IsNull(rs.Fields(2)) Then txtserialno = (rs.Fields(2))
'If Not IsNull(rs.Fields(3)) Then txtbalance = (rs.Fields("qout"))

If txtbalance <= 0 Then
MsgBox "Your stock is below zero please reorder", vbInformation

End If
'// check with serial no if it exist

'provider = cnn
Set cn = New ADODB.connection
cn.Open provider, , "pius12"

sql = "SELECT  serialid,serialno,p_code,used FROM serialno where p_code='" & txtpcode.Text & "'"
Set rs = New ADODB.Recordset
rs.Open sql, cn
If Not rs.EOF Then
'If Not IsNull(rs.Fields(0)) Then serialid = rs.Fields(0)
While Not rs.EOF
  txtserialid.AddItem rs.Fields(0)
rs.MoveNext
Wend
End If

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
provider = cnn
Set cn = New ADODB.connection
cn.Open provider, , "pius12"
sql = "select P_CODE,P_NAME,S_NO,QOUT from products where p_code='" & txtpcode & "'"
Set rs = New ADODB.Recordset
rs.Open sql, cn
If Not rs.EOF Then
 txtpcode = (rs.Fields(0))
If Not IsNull(rs.Fields(1)) Then txtpname = (rs.Fields(1))
If Not IsNull(rs.Fields(3)) Then txtbalance = (rs.Fields(3))

If txtbalance <= 0 Then
MsgBox "Your stock is below zero please reorder", vbInformation
End If
End If

provider = cnn
Set cn = New ADODB.connection
cn.Open provider, , "pius12"

sql = "SELECT  serialid,serialno,p_code,used FROM serialno where p_code='" & txtpcode.Text & "'"
Set rs = New ADODB.Recordset
rs.Open sql, cn
If Not rs.EOF Then
'If Not IsNull(rs.Fields(0)) Then serialid = rs.Fields(0)
While Not rs.EOF
  txtserialid.AddItem rs.Fields(0)
rs.MoveNext
Wend
End If
'End If

'// check with serial no if it exist
End Sub

Private Sub txtserialid_change()

provider = cnn
Set cn = New ADODB.connection
cn.Open provider, , "pius12"

sql = ""

sql = "SELECT  serialid,serialno,p_code,used FROM serialno where serialid=" & txtserialid.Text & ""
Set rs = New ADODB.Recordset
rs.Open sql, cn
If Not rs.EOF Then
If Not IsNull(rs.Fields(1)) Then txtserialno = rs.Fields(1)
End If
'End If
End Sub

Private Sub txtserialid_click()
txtserialid_change
End Sub
