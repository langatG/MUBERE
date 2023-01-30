VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmaccountspayable 
   Caption         =   "AC Accounts Payable"
   ClientHeight    =   7485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10230
   Icon            =   "frmAccountspayable.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   10230
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtvatc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Enabled         =   0   'False
      Height          =   285
      Left            =   4800
      TabIndex        =   53
      Top             =   6600
      Width           =   2655
   End
   Begin VB.PictureBox Picture5 
      Height          =   255
      Left            =   3360
      Picture         =   "frmAccountspayable.frx":0442
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   52
      Top             =   6600
      Width           =   255
   End
   Begin VB.TextBox txtvatcontrol 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Height          =   285
      Left            =   3720
      TabIndex        =   51
      Top             =   6600
      Width           =   975
   End
   Begin VB.CheckBox chkvat 
      Caption         =   "Include Vat "
      Height          =   255
      Left            =   4080
      TabIndex        =   50
      Top             =   3480
      Width           =   1455
   End
   Begin VB.TextBox txtvatamount 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      TabIndex        =   49
      Top             =   3480
      Width           =   2295
   End
   Begin VB.TextBox txtinvoicenumber 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      TabIndex        =   46
      ToolTipText     =   "type here invoice no as per the invoice"
      Top             =   1800
      Width           =   3855
   End
   Begin VB.PictureBox Picture3 
      Height          =   285
      Left            =   3960
      Picture         =   "frmAccountspayable.frx":0D0C
      ScaleHeight     =   225
      ScaleWidth      =   195
      TabIndex        =   45
      Top             =   960
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   285
      Left            =   9960
      Picture         =   "frmAccountspayable.frx":0FCE
      ScaleHeight     =   225
      ScaleWidth      =   195
      TabIndex        =   43
      Top             =   2640
      Width           =   255
   End
   Begin VB.PictureBox Picture4 
      Height          =   285
      Left            =   9960
      Picture         =   "frmAccountspayable.frx":1290
      ScaleHeight     =   225
      ScaleWidth      =   195
      TabIndex        =   42
      Top             =   2160
      Width           =   255
   End
   Begin VB.TextBox txtidno 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      TabIndex        =   6
      Top             =   5760
      Width           =   2295
   End
   Begin VB.TextBox txtreceivedby 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7800
      TabIndex        =   15
      Top             =   5040
      Width           =   2175
   End
   Begin VB.TextBox txtpurpose 
      Appearance      =   0  'Flat
      Height          =   645
      Left            =   1680
      TabIndex        =   5
      Top             =   4920
      Width           =   3855
   End
   Begin VB.ComboBox cbocheckstatus 
      Height          =   315
      ItemData        =   "frmAccountspayable.frx":1552
      Left            =   7800
      List            =   "frmAccountspayable.frx":1562
      TabIndex        =   12
      Text            =   "cbocheckstatus"
      Top             =   3600
      Width           =   2175
   End
   Begin VB.TextBox txtfundingaccount 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7800
      TabIndex        =   10
      Top             =   2640
      Width           =   2175
   End
   Begin VB.TextBox txtpaymentref 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      TabIndex        =   23
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox txtpaidto 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Top             =   1440
      Width           =   3855
   End
   Begin VB.TextBox txtparticulars 
      Appearance      =   0  'Flat
      Height          =   645
      Left            =   1680
      TabIndex        =   1
      Top             =   2280
      Width           =   3855
   End
   Begin VB.TextBox txtpayingaccount 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7800
      TabIndex        =   9
      Top             =   2160
      Width           =   2175
   End
   Begin VB.TextBox txtpaidamount 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      Top             =   3120
      Width           =   2295
   End
   Begin VB.TextBox txtchequeno 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Top             =   3960
      Width           =   2295
   End
   Begin VB.TextBox txtexchangerate 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      TabIndex        =   4
      Top             =   4440
      Width           =   2295
   End
   Begin VB.ComboBox cbopaymentmode 
      Height          =   315
      ItemData        =   "frmAccountspayable.frx":1590
      Left            =   7800
      List            =   "frmAccountspayable.frx":15A3
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1680
      Width           =   2175
   End
   Begin VB.ComboBox cbocurrency 
      Height          =   315
      ItemData        =   "frmAccountspayable.frx":15C9
      Left            =   7800
      List            =   "frmAccountspayable.frx":15DF
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   4560
      Width           =   2175
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   6120
      TabIndex        =   22
      Top             =   7080
      Width           =   1095
   End
   Begin VB.CommandButton cmdedits 
      Caption         =   "&Edit"
      Height          =   375
      Left            =   4680
      TabIndex        =   21
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton cmdsave 
      Appearance      =   0  'Flat
      Caption         =   "&Save"
      Height          =   375
      Left            =   3480
      TabIndex        =   20
      Top             =   7080
      Width           =   1095
   End
   Begin VB.CommandButton cmdnew 
      Appearance      =   0  'Flat
      Caption         =   "&New "
      Height          =   375
      Left            =   2400
      TabIndex        =   19
      Top             =   7080
      Width           =   975
   End
   Begin VB.TextBox txtredirect 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Height          =   285
      Left            =   3720
      TabIndex        =   18
      Top             =   6240
      Width           =   975
   End
   Begin VB.PictureBox Picture2 
      Height          =   255
      Left            =   3360
      Picture         =   "frmAccountspayable.frx":1601
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   17
      Top             =   6240
      Width           =   255
   End
   Begin VB.TextBox txtredirectaccno 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Enabled         =   0   'False
      Height          =   285
      Left            =   4800
      TabIndex        =   16
      Top             =   6240
      Width           =   2655
   End
   Begin MSComCtl2.DTPicker txtchequedate 
      Height          =   255
      Left            =   7800
      TabIndex        =   13
      Top             =   4080
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   450
      _Version        =   393216
      Format          =   121831425
      CurrentDate     =   38875
   End
   Begin MSComCtl2.DTPicker txtpaymentdate 
      Height          =   375
      Left            =   7800
      TabIndex        =   7
      Top             =   1080
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      Format          =   121831425
      CurrentDate     =   38875
   End
   Begin MSComCtl2.DTPicker txtdatecollected 
      Height          =   255
      Left            =   7800
      TabIndex        =   11
      Top             =   3120
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   450
      _Version        =   393216
      Format          =   121831425
      CurrentDate     =   38875
   End
   Begin MSComctlLib.ImageList imgEmpTool 
      Left            =   1200
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccountspayable.frx":1ECB
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccountspayable.frx":1FDD
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccountspayable.frx":20EF
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccountspayable.frx":2201
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccountspayable.frx":2743
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccountspayable.frx":2A15
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   44
      Top             =   0
      Width           =   10230
      _ExtentX        =   18045
      _ExtentY        =   1058
      ButtonWidth     =   1826
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgEmpTool"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "AlpaSearch "
            Key             =   "search"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "View  "
            ImageIndex      =   6
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Sum"
                  Text            =   "Summary"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Det"
                  Text            =   "Details"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Reports"
            ImageIndex      =   7
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "LEmp"
                  Text            =   "Agents Summary"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "DEmp"
                  Text            =   "Agent's Details"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "PEmp"
                  Text            =   "Accounts Range By Category"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Label Label21 
      Caption         =   "V.A.T CONTROL"
      Height          =   255
      Left            =   1320
      TabIndex        =   54
      Top             =   6600
      Width           =   1935
   End
   Begin VB.Label Label20 
      Caption         =   "V.A.T Amount"
      Height          =   255
      Left            =   120
      TabIndex        =   48
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label14 
      Caption         =   "Invoice Number"
      Height          =   255
      Left            =   120
      TabIndex        =   47
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label19 
      Caption         =   "ID No."
      Height          =   255
      Left            =   120
      TabIndex        =   41
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label Label18 
      Caption         =   "Received By."
      Height          =   255
      Left            =   6120
      TabIndex        =   40
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Label17 
      Caption         =   "Purpose"
      Height          =   255
      Left            =   120
      TabIndex        =   39
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "Cheque Status."
      Height          =   255
      Left            =   6120
      TabIndex        =   38
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label Label8 
      Caption         =   "Date Collected."
      Height          =   255
      Left            =   6120
      TabIndex        =   37
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label16 
      Caption         =   "Funding Account:"
      Height          =   255
      Left            =   6120
      TabIndex        =   36
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Payment No."
      Height          =   255
      Left            =   120
      TabIndex        =   35
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Paid To:"
      Height          =   255
      Left            =   120
      TabIndex        =   34
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Particulars"
      Height          =   255
      Left            =   120
      TabIndex        =   33
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Payment Date"
      Height          =   255
      Left            =   6120
      TabIndex        =   32
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Paying Account:"
      Height          =   255
      Left            =   6120
      TabIndex        =   31
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "Payment Mode"
      Height          =   255
      Left            =   6120
      TabIndex        =   30
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "Paid Amount."
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "Cheque No."
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   "Cheque Date."
      Height          =   255
      Left            =   6120
      TabIndex        =   27
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Label12 
      Caption         =   "Currency"
      Height          =   255
      Left            =   6120
      TabIndex        =   26
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label Label13 
      Caption         =   "Exchange Rate."
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000040C0&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   4
      DrawMode        =   9  'Not Mask Pen
      X1              =   0
      X2              =   10455
      Y1              =   6960
      Y2              =   6975
   End
   Begin VB.Label Label15 
      Caption         =   "ACCOUNTS PAYABLES"
      Height          =   255
      Left            =   1320
      TabIndex        =   24
      Top             =   6240
      Width           =   1935
   End
End
Attribute VB_Name = "frmaccountspayable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ed As Boolean
Dim sta As Integer
Dim cn As Connection
Dim myclass As Object
Dim Provider As String

Private Sub cmdcancel_Click()

End Sub

Private Sub chkvat_Click()
Dim vat As Currency
Set cn = CreateObject("adodb.connection")
Set myclass = New cdbase
    Set cn = CreateObject("adodb.connection")
    Provider = myclass.OpenCon
cn.Open Provider, "atm", "atm"
    Set rs = CreateObject("ADODB.Recordset")
    sql = ""
   sql = "SELECT vat  from param"
    rs.Open sql, cn, adOpenKeyset, adLockOptimistic
    If Not rs.EOF Then
    If Not IsNull(rs.Fields(0)) Then vat = rs.Fields(0)
    txtvatamount = vat / 100 * txtpaidamount
    End If
    
End Sub

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdedits_Click()
ed = True
End Sub

Private Sub cmdNew_Click()
On Error Resume Next
txtpaidamount = ""
txtpaidto = ""
txtParticulars = ""
txtpayingaccount = ""
txtpaidto.SetFocus
txtChequeno = ""
txtexchangerate = ""
txtfundingaccount = ""
txtidno = ""
txtpurpose = ""
txtreceivedby = ""
txtinvoicenumber = ""
ed = False

Set cn = CreateObject("adodb.connection")

Set myclass = New cdbase
    Set cn = CreateObject("adodb.connection")
    Provider = myclass.OpenCon
cn.Open Provider, "atm", "atm"
    Set rs = CreateObject("ADODB.Recordset")
    
    sql = ""
   sql = "SELECT top 1 p_no  from Ap order by pid desc"
    rs.Open sql, cn
    If Not rs.EOF Then
    txtpaymentref = rs.Fields(0) + 1
    End If
End Sub

Private Sub cmdsave_Click()
On Error GoTo ErrorHandler
Set cn = CreateObject("adodb.connection")

Set myclass = New cdbase
Set cn = CreateObject("adodb.connection")
Provider = myclass.OpenCon
cn.Open Provider, "atm", "atm"




If txtredirect = "" Then
MsgBox "You Have Not input the Control Account As Required; Please Do So Before You Proceed", vbInformation
Exit Sub
End If

If txtinvoicenumber = "" Then
MsgBox "You Have Not input the Invoice Number As Required; Please Do So Before You Proceed", vbInformation
Exit Sub
End If
    Set rs = CreateObject("ADODB.Recordset")
    
    sql = ""
   sql = "SELECT * from Ap WHERE p_NO='" & txtpaymentref & "' and inv_no='" & txtinvoicenumber & "'"
    rs.Open sql, cn, adOpenKeyset, adLockOptimistic
    If txtvatamount = "" Then txtvatamount = 0
If ed = True Then '// update the bank details
'// check if posted then no need to edit again
If rs.Fields("posted") = True Then
MsgBox "Record already posted and cannot be reversed", vbCritical, "AC - AP"
Exit Sub
End If
        '//update here some of the details like cheque no and others
        sql = ""
        sql = "set dateformat dmy update  AP     set     "
        sql = sql & " paidto='" & txtpaidto & "', particulars='" & txtParticulars & "', p_date='" & txtpaymentdate & "', F_AccnO='" & txtfundingaccount & "', exchangerate=" & txtexchangerate & ", "
        sql = sql & " checked_by='" & User & "', chequeno='" & txtChequeno & "', chequestatus='" & cbocheckstatus & "', Date_Collected='" & txtdatecollected & "', Approvedby='" & User & "', ChequeDate='" & txtchequedate & "', Receivedby='" & txtreceivedby & "', IDNo='" & txtidno & "',paYmentmode='" & cbopaymentmode & "',p_accno='" & txtpayingaccount & "',vat=" & txtvatamount & ",vat_accno='" & txtvatcontrol & "'  where p_no='" & txtpaymentref & "'"
        cn.Execute sql
        
txtpaidamount = ""
txtpaidto = ""
txtParticulars = ""
txtpayingaccount = ""
txtpaidto.SetFocus
txtChequeno = ""
txtexchangerate = ""
txtfundingaccount = ""
txtidno = ""
txtpurpose = ""
txtreceivedby = ""
MsgBox "You Have Successfully Update the record", vbInformation, "AC- Payables"
Else

'// check if all the it already exist
        If Not rs.EOF Then
            
            MsgBox "The Payment No already exist Please input a new one.", vbInformation, "Ap"
            txtpaymentref.SetFocus
            Exit Sub
          
        End If
       
    

sql = ""
sql = "set dateformat dmy INSERT INTO AP  "
sql = sql & "             (P_No, paidto, p_amount, curr, Purpose, accno, particulars, p_date, F_AccnO, exchangerate, checked_by, chequeno, chequestatus,"
sql = sql & "               Date_Collected, Approvedby, ChequeDate, Receivedby, IDNo,paYmentmode,auditid,auditdatetime,p_accno,inv_no,vat,vat_accno)"
sql = sql & "  VALUES     (" & txtpaymentref & ", '" & txtpaidto & "', " & txtpaidamount & ", '" & cbocurrency & "', '" & txtpurpose & "', '" & txtredirect & "', '" & txtParticulars & "', '" & txtpaymentdate & "', '" & txtfundingaccount & "', " & txtexchangerate & ", '" & User & "', '" & txtChequeno & "', '" & cbocheckstatus & "', '" & txtdatecollected & "', '" & User & "',"
sql = sql & "           '" & txtchequedate & "', '" & txtreceivedby & "', '" & txtidno & "','" & cbopaymentmode & "','" & User & "','" & Get_Server_Date & "','" & txtpayingaccount & "','" & txtinvoicenumber & "'," & txtvatamount & ",'" & txtvatcontrol & "')"
cn.Execute sql

MsgBox "You successfully added the record", vbInformation, "AC-Payables"
End If

txtpaidamount = ""
txtpaidto = ""
txtParticulars = ""
txtpayingaccount = ""
txtpaymentref = ""
txtChequeno = ""
txtexchangerate = ""
txtfundingaccount = ""
txtidno = ""
txtpurpose = ""
txtreceivedby = ""

cmdNew_Click
Exit Sub
ErrorHandler:
MsgBox err.description
'Form_Load
End Sub
Private Sub get_namevat(vatgl As String)

    Set cn = CreateObject("adodb.connection")
    Set myclass = New cdbase
    Provider = myclass.OpenCon
   cn.Open Provider, "atm", "atm"
    sql = ""
    sql = "select * from glsetup where accno='" & vatgl & "'"
    Set rs = New ADODB.Recordset
    rs.Open sql, cn, adOpenKeyset, adLockOptimistic
    If rs.EOF Then
    Else
   ' If Not IsNull(rs.Fields("accno")) Then acccr = rs.Fields("accno")
    If Not IsNull(rs.Fields("accno")) Then txtvatcontrol = rs.Fields("accno")
    If Not IsNull(rs.Fields("glaccname")) Then txtvatc = rs.Fields("gLaccname")
    End If
End Sub
Private Sub get_namedr()

    Set cn = CreateObject("adodb.connection")
    Set myclass = New cdbase
    Provider = myclass.OpenCon
   cn.Open Provider, "atm", "atm"
    sql = ""
    sql = "select * from glsetup where accno='" & txtredirect & "'"
    Set rs = New ADODB.Recordset
    rs.Open sql, cn, adOpenKeyset, adLockOptimistic
    If rs.EOF Then
    Else
   ' If Not IsNull(rs.Fields("accno")) Then acccr = rs.Fields("accno")
    If Not IsNull(rs.Fields("glaccname")) Then txtredirectaccno = rs.Fields("glaccname")
    'If Not IsNull(rs.Fields("accno")) Then Txtaccno = rs.Fields("accno")
    End If
End Sub
Private Sub get_namecr1(ACCNO As String)

    Set cn = CreateObject("adodb.connection")
    Set myclass = New cdbase
    Provider = myclass.OpenCon
   cn.Open Provider, "atm", "atm"
    sql = ""
    sql = "select name from cub where accno='" & Trim(ACCNO) & "'"
    Set rs = New ADODB.Recordset
    rs.Open sql, cn, adOpenKeyset, adLockOptimistic
    If rs.EOF Then
    Else
    If Not IsNull(rs.Fields(0)) Then txtredirectaccno = Trim(rs.Fields(0))
    End If
End Sub
Private Sub get_namecr()

    Set cn = CreateObject("adodb.connection")
    Set myclass = New cdbase
    Provider = myclass.OpenCon
   cn.Open Provider, "atm", "atm"
    sql = ""
    sql = "select * from cub where accno='" & sel & "'"
    Set rs = New ADODB.Recordset
    rs.Open sql, cn, adOpenKeyset, adLockOptimistic
    If rs.EOF Then
    Else
    If Not IsNull(rs.Fields("name")) Then txtredirectaccno = rs.Fields("name")
    End If
End Sub
Private Sub Form_Load()
On Error GoTo ErrorHandler
cbocurrency = "KES"
txtpaymentdate = Format(Get_Server_Date, "DD/MM/YYYY")
txtdatecollected = Format(Get_Server_Date, "DD/MM/YYYY")
txtchequedate = Format(Get_Server_Date, "DD/MM/YYYY")
Set rs = Nothing
 Set cn = CreateObject("adodb.connection")
Set myclass = New cdbase
  Provider = myclass.OpenCon
 cn.Open Provider, "atm", "atm"
     sql = ""
     sql = "select apcontrol,arcontrol,vat_c from param"
     Set rs = New ADODB.Recordset
     rs.Open sql, cn, adOpenKeyset, adLockOptimistic
     
     If Not rs.EOF Then
    ' txtcomm = rs.Fields(0)
     txtredirect = rs.Fields(0)
     txtvatcontrol = rs.Fields(2)
     get_namecr
     get_namedr
     
     get_namevat txtvatcontrol
     End If
     cmdNew_Click
Exit Sub
ErrorHandler:
MsgBox err.description
End Sub






Private Sub Picture1_Click()
Me.MousePointer = vbHourglass
         frmsearchacc.Show vbModal
        txtfundingaccount = sel
        'Txtaccno_Validate True
        Me.MousePointer = 0
End Sub

Private Sub Picture2_Click()
Me.MousePointer = vbHourglass
         frmsearchacc.Show vbModal
        txtredirect = sel
        'Txtaccno_Validate True
        get_namecr
        Me.MousePointer = 0
End Sub

Private Sub Picture3_Click()
Me.MousePointer = vbHourglass
txtpaymentref = ""
         frmsearchap.Show vbModal
        txtpaymentref = sel
        txtpaymentref_Validate True
        Me.MousePointer = 0
End Sub

Private Sub Picture4_Click()
Me.MousePointer = vbHourglass
         frmsearchacc.Show vbModal
        txtpayingaccount = sel
        'Txtaccno_Validate True
        Me.MousePointer = 0
        
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo errFix
Me.MousePointer = vbHourglass
Select Case Button.Key
  Case "search"
        Me.MousePointer = vbHourglass
         frmsearchacc.Show vbModal
       ' txtcustno = Sel
        'txtcustno_Validate True
        Me.MousePointer = 0
 End Select
Me.MousePointer = 0
   Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Member Registration"
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
On Error GoTo ErrHandler
    Dim myDate As Date
    Dim STRFORMULA As String
    Dim title As String
    Dim reportname As String
    
  
    Select Case ButtonMenu.Text
        Case "Agents Summary"
           
        'STRFORMULA = "month({loans.applicdate})=" & Month(DTPPeriod) & " and year({loans.applicdate})=" & Year(DTPPeriod) & ""
          reportname = "agentssummarylist.rpt"
         'title = UCase(rst!CompanyName)
        Show_Sales_Crystal_Report STRFORMULA, reportname, title

        Case "Periodic Loan Transactions"
        'frmReports.Show vbModal, Me
        If Continue = False Then
            Exit Sub
        End If
   
        STRFORMULA = "month({repay.datereceived})=" & month(SelectedDate) & " and year({repay.datereceived})=" & year(SelectedDate) & ""

        reportname = "Periodic Transactions.rpt"
        title = "LOAN REPAYMENT TRANSACTIONS FOR " & UCase(MonthName(month(SelectedDate))) & " " & year(SelectedDate)
        Show_Sales_Crystal_Report STRFORMULA, reportname, title
        Case "Defaulters List"
        
        
       ' STRFORMULA = "" & DateDiff("m", " & ({LoanBal.lastdate}) &", DTPPeriod) & ">1"
        reportname = "Default.rpt"
        'title = UCase(MonthName(month(DTPPeriod.Value))) & " " & Year(DTPPeriod.Value)
        Show_Sales_Crystal_Report STRFORMULA, reportname, title
        Case "Underpaid Loans"
        
        Case "Loan Balances Summary"
    
        'STRFORMULA = "" & DateDiff("m", " & ({LoanBal.lastdate}) &", DTPPeriod) & ">1"
        reportname = "loan balances summary.rpt"
        title = UCase(rst!CompanyName)
        Show_Sales_Crystal_Report STRFORMULA, reportname, title
        Case "Accrued Transactions"
        
    End Select
    Exit Sub
ErrHandler:
    MsgBox err.description, , "Sacco Master"
    Exit Sub
End Sub

Private Sub txtpaymentref_Validate(Cancel As Boolean)
Dim myrec1 As Object
Dim rss As Object
Dim amt As Currency
Set myclass = New cdbase
Set cn = CreateObject("adodb.connection")
Provider = myclass.OpenCon
cn.Open Provider, "atm", "atm"
Dim MYRE As Recordset
    Set MYRE = CreateObject("adodb.recordset")
    sql = "SELECT * from ap where p_no='" & txtpaymentref & "' "
     MYRE.Open sql, cn
     If txtpaymentref <> "" Then
     If MYRE.EOF Then
      MsgBox "The Payment does not exist Please Seek assistance from the customer services", vbInformation, "Transactions"
     Exit Sub
     End If
     End If
     
     Set myrec1 = New ADODB.Recordset
 sql = "SELECT     paidto, particulars, p_amount, chequeno, exchangerate, Purpose, IDNo, p_date, paymentmode, P_accno, F_AccNo, Date_Collected, chequestatus, ChequeDate, curr, Receivedby,Inv_no ,accno FROM         AP where p_no ='" & sel & "' "
     myrec1.Open sql, cn
     If myrec1.EOF Then
txtpaidamount = ""
txtpaidto = ""
txtParticulars = ""
txtpayingaccount = ""
txtpaidto.SetFocus
txtChequeno = ""
txtexchangerate = ""
txtfundingaccount = ""
txtidno = ""
txtpurpose = ""
txtreceivedby = ""
txtinvoicenumber = ""



     
     Else
     
      If Not IsNull(myrec1.Fields(0)) Then txtpaidto = myrec1.Fields(0)
      If Not IsNull(myrec1.Fields(1)) Then txtParticulars = myrec1.Fields(1)
      If Not IsNull(myrec1.Fields(2)) Then txtpaidamount = myrec1.Fields(2)
      If Not IsNull(myrec1.Fields(3)) Then txtChequeno = Trim(myrec1.Fields(3))
      If Not IsNull(myrec1.Fields(4)) Then txtexchangerate = myrec1.Fields(4)
      If Not IsNull(myrec1.Fields(5)) Then txtpurpose = myrec1.Fields(5)
      If Not IsNull(myrec1.Fields(6)) Then txtidno = Trim(myrec1.Fields(6))
      If Not IsNull(myrec1.Fields(7)) Then txtpaymentdate = myrec1.Fields(7)
      If Not IsNull(myrec1.Fields(8)) Then cbopaymentmode = myrec1.Fields(8)
      If Not IsNull(myrec1.Fields(9)) Then txtpayingaccount = myrec1.Fields(9)
      If Not IsNull(myrec1.Fields(10)) Then txtfundingaccount = myrec1.Fields(10)
      If Not IsNull(myrec1.Fields(11)) Then txtdatecollected = myrec1.Fields(11)
      If Not IsNull(myrec1.Fields(12)) Then cbocheckstatus = myrec1.Fields(12)
      If Not IsNull(myrec1.Fields(13)) Then txtchequedate = myrec1.Fields(13)
      If Not IsNull(myrec1.Fields(14)) Then cbocurrency = myrec1.Fields(14)
      If Not IsNull(myrec1.Fields(15)) Then txtreceivedby = Trim(myrec1.Fields(15))
      If Not IsNull(myrec1.Fields(16)) Then txtinvoicenumber = Trim(myrec1.Fields(16))
       If Not IsNull(myrec1.Fields(17)) Then txtredirect = Trim(myrec1.Fields(17))
      get_namecr1 txtredirect
     End If


End Sub

