VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmupdatesellingprice 
   BackColor       =   &H80000002&
   Caption         =   "UPDATE PRICES"
   ClientHeight    =   3585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7725
   LinkTopic       =   "Form1"
   ScaleHeight     =   3585
   ScaleWidth      =   7725
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtqty 
      Height          =   375
      Left            =   4080
      TabIndex        =   36
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Update stock balance"
      Height          =   495
      Left            =   5400
      TabIndex        =   35
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox txtBalance1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   34
      Top             =   1440
      Width           =   855
   End
   Begin VB.TextBox txtppprice 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1680
      TabIndex        =   31
      Top             =   1560
      Width           =   1455
   End
   Begin VB.TextBox txtpcode 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1680
      TabIndex        =   20
      Top             =   360
      Width           =   1935
   End
   Begin VB.TextBox txtpname 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1680
      TabIndex        =   19
      Top             =   960
      Width           =   4335
   End
   Begin VB.TextBox txtserialno 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   9600
      MultiLine       =   -1  'True
      TabIndex        =   18
      Top             =   2520
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.TextBox txtquantity 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   9600
      TabIndex        =   17
      Top             =   3120
      Width           =   1935
   End
   Begin VB.CommandButton cmdnew 
      Caption         =   "New"
      Height          =   375
      Left            =   1440
      TabIndex        =   16
      Top             =   3120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   375
      Left            =   2640
      TabIndex        =   15
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      Height          =   375
      Left            =   5040
      TabIndex        =   14
      Top             =   3120
      Width           =   975
   End
   Begin VB.TextBox txtbalance 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   13080
      TabIndex        =   12
      Top             =   3120
      Width           =   1815
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      DrawStyle       =   2  'Dot
      DrawWidth       =   17015
      Height          =   360
      Left            =   3600
      Picture         =   "frmupdatesellingprice.frx":0000
      ScaleHeight     =   360
      ScaleWidth      =   240
      TabIndex        =   11
      Top             =   360
      Width           =   240
   End
   Begin VB.CommandButton cmddelete 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   3840
      TabIndex        =   10
      Top             =   3120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ComboBox cbosupplier 
      Height          =   315
      Left            =   9360
      TabIndex        =   9
      Top             =   1800
      Width           =   4455
   End
   Begin VB.CheckBox CHKSERIALIZED 
      Caption         =   "Serialized"
      Height          =   375
      Left            =   14040
      TabIndex        =   8
      Top             =   1680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox chkserialrequired 
      Caption         =   "Serial Required"
      Height          =   375
      Left            =   14040
      TabIndex        =   7
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Frame fra1 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   4200
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   2895
      Begin VB.TextBox txtpassword 
         Appearance      =   0  'Flat
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   240
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000002&
         Caption         =   "Enter Password"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.TextBox txtpassit 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   11640
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   3240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtRLevel 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox txtsellingprice 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox txtpprice 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   9600
      TabIndex        =   0
      Top             =   3720
      Width           =   1455
   End
   Begin MSComCtl2.DTPicker txtdateenterered 
      Height          =   375
      Left            =   5400
      TabIndex        =   13
      Top             =   480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   130416641
      CurrentDate     =   38814
   End
   Begin VB.Label Label14 
      BackColor       =   &H80000002&
      Caption         =   "Balance"
      Height          =   255
      Left            =   3360
      TabIndex        =   33
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000002&
      Caption         =   "Product Code"
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000002&
      Caption         =   "Product Name"
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Serial No"
      Height          =   255
      Left            =   8040
      TabIndex        =   28
      Top             =   2520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Quantity"
      Height          =   255
      Left            =   8040
      TabIndex        =   27
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Balance In Store"
      Height          =   255
      Left            =   13080
      TabIndex        =   26
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "Date Entered"
      Height          =   255
      Left            =   4200
      TabIndex        =   25
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "Supplier"
      Height          =   255
      Left            =   7800
      TabIndex        =   24
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label11 
      BackColor       =   &H80000002&
      Caption         =   "Re-Order Level"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label10 
      BackColor       =   &H80000002&
      Caption         =   "Selling Price "
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "Purchase Price "
      Height          =   375
      Left            =   8040
      TabIndex        =   21
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label12 
      BackColor       =   &H80000002&
      Caption         =   "Buying Price "
      Height          =   255
      Left            =   120
      TabIndex        =   32
      Top             =   1560
      Width           =   1215
   End
End
Attribute VB_Name = "frmupdatesellingprice"
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
cn.Open Provider, , "MAZIWA"
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
sql = "delete from products where p_code='" & txtpcode & "'"
cn.Execute sql
sql = ""
sql = "delete from CrditSale  where p_code='" & txtpcode & "'"
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
Dim unsera As Integer
If txtquantity = "" Then
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
cn.Open Provider, , "MAZIWA"
sql = "select P_CODE,qout,unserialized from AG_products where p_code='" & txtpcode & "'"
Set rs = New ADODB.Recordset
rs.Open sql, cn
If rs.EOF Then
'// insert into products

Else
Dim D As Double
If Not IsNull(rs.Fields(2)) Then D = rs.Fields(2)
sql = ""
sql = "set dateformat DMY update AG_Products set sprice=" & txtsellingprice & ",pprice=" & txtppprice & " where p_code='" & txtpcode.Text & "'"
cn.Execute sql

Dim rsst As Recordset
sql = ""
sql = "set dateformat DMY select top 1 * from AG_stockbalance where p_code='" & txtpcode & "' order by trackid desc "
Set rsst = New ADODB.Recordset
rsst.Open sql, cn
If Not rsst.EOF Then
sql = ""
sql = "set dateformat DMY INSERT INTO stockbalance"
End If
End If
MsgBox "Price updated successfully", vbInformation
txtbalance = ""
txtpcode = ""
txtpname = ""
txtSERIALNO = ""
txtpprice = ""
txtsellingprice = ""
txtRLevel = ""
txtquantity = ""
txtppprice = ""
End Sub

Private Sub Command1_Click()
sql = ""
sql = "set dateformat  dmy update ag_products  set Qout= '" & txtqty & "' where p_code='" & txtpcode & "'  "
oSaccoMaster.ExecuteThis (sql)
MsgBox "Updated"
End Sub

Private Sub Form_Load()
txtdateenterered = Format(Date, "dd,mm,yyyy")
 Set rst = New Recordset
    Dim cn As Connection
    Set cn = New ADODB.Connection
    Provider = "MAZIWA"
    cn.Open Provider, "atm", "atm"
    Set rst = New Recordset
    'sql = "Select companyname from Suppliers"
    'rst.Open sql, cn, adOpenKeyset, adLockOptimistic
'    While Not rst.EOF
'    cbosupplier.AddItem rst.Fields(0)
'    rst.MoveNext
'    Wend
    
End Sub
Private Sub Picture2_Click()
Dim ppricess As Double
Dim rspr As New ADODB.Recordset
frmSearch.Show vbModal
Dim Y As String
Y = sel
  Set rs = Nothing
If Y <> "" Then
'Provider = cnn
''Set cn = New ADODB.connection
'cn.Open Provider, , "MAIN"
Set rspr = oSaccoMaster.GetRecordset("select P_CODE,P_NAME,S_NO,QOUT, pprice,sprice from ag_products where p_code='" & Y & "'")

If Not rspr.EOF Then
 txtpcode = IIf(IsNull(rspr.Fields(0)), "", rspr.Fields(0))
txtpname = IIf(IsNull(rspr.Fields(1)), "", rspr.Fields(1)) '(rs.Fields(1))
txtBalance1 = IIf(IsNull(rspr.Fields(3)), 0, rspr.Fields(3))
ppricess = IIf(IsNull(rspr.Fields(4)), 0, rspr.Fields(4))
txtppprice = ppricess
txtsellingprice = IIf(IsNull(rspr.Fields(5)), 0, rspr.Fields(5))
'txtRLevel = IIf(IsNull(rspr.Fields(6)), 0, rspr.Fields(6))
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
Dim rsp As Recordset
Set cn = CreateObject("adodb.connection")
Provider = cn
Set cn = New ADODB.Connection
cn.Open Provider, , "MAZIWA"
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
Provider = cn
Set cn = New ADODB.Connection
cn.Open Provider, "atm", "atm"
sql = "select P_CODE,P_NAME,S_NO,QOUT from AG_products where p_code='" & txtpcode & "'"
Set rs = New ADODB.Recordset
rs.Open sql, cn
If Not rs.EOF Then
 txtpcode = (rs.Fields(0))
If Not IsNull(rs.Fields(1)) Then txtpname = (rs.Fields(1))
If Not IsNull(rs.Fields(3)) Then txtBalance1 = (rs.Fields(3))
If txtBalance1 <= 0 And txtpcode <> "" Then
MsgBox "Your stock is below zero please reorder", vbInformation
End If
End If
End Sub
Private Sub txtquantity_Validate(Cancel As Boolean)
If Not IsNumeric(txtquantity) Then
MsgBox "Enter values please", vbCritical
txtquantity = ""
txtquantity.SetFocus
Exit Sub
End If
End Sub

