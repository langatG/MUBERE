VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPayment 
   Caption         =   "Pay Farmers"
   ClientHeight    =   8550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12135
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8550
   ScaleWidth      =   12135
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      Height          =   405
      Left            =   10680
      TabIndex        =   46
      Top             =   7920
      Width           =   975
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "Show"
      Height          =   285
      Left            =   10680
      TabIndex        =   44
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmdPay 
      Caption         =   "Pay"
      Height          =   375
      Left            =   8760
      TabIndex        =   43
      Top             =   7920
      Width           =   1335
   End
   Begin VB.ComboBox cboTCode 
      Height          =   405
      Left            =   7680
      TabIndex        =   42
      Text            =   "<Select Transporter>"
      Top             =   3960
      Width           =   2895
   End
   Begin VB.ComboBox cboLocation 
      Height          =   405
      Left            =   7680
      TabIndex        =   41
      Text            =   "<Select Route>"
      Top             =   3480
      Width           =   2895
   End
   Begin VB.OptionButton Option3 
      Caption         =   "All Suppliers"
      Height          =   375
      Left            =   6000
      TabIndex        =   40
      Top             =   3960
      Width           =   1575
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Suppliers For Route/Location"
      Height          =   285
      Left            =   2880
      TabIndex        =   39
      Top             =   3960
      Width           =   3015
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Suppliers For Transporter"
      Height          =   285
      Left            =   0
      TabIndex        =   38
      Top             =   3960
      Width           =   2775
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3000
      TabIndex        =   22
      Top             =   120
      Width           =   4575
   End
   Begin VB.TextBox txtBank 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   720
      TabIndex        =   21
      Top             =   1080
      Width           =   2175
   End
   Begin VB.TextBox txtTelNo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   8880
      TabIndex        =   20
      Top             =   600
      Width           =   3015
   End
   Begin VB.TextBox txtBox 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3000
      TabIndex        =   19
      Top             =   600
      Width           =   4575
   End
   Begin VB.TextBox txtTransport 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   8880
      TabIndex        =   18
      Top             =   120
      Width           =   3015
   End
   Begin VB.TextBox txtSNo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   600
      TabIndex        =   17
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtBBranch 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4080
      TabIndex        =   16
      Top             =   1080
      Width           =   2175
   End
   Begin VB.TextBox txtAccNo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   8880
      TabIndex        =   15
      Top             =   1080
      Width           =   3015
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   0
      TabIndex        =   9
      Top             =   2040
      Width           =   7455
      Begin VB.CommandButton cmdShowNet 
         Caption         =   "Show Net"
         Default         =   -1  'True
         Height          =   375
         Left            =   6000
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   375
         Left            =   1320
         TabIndex        =   11
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   131596289
         CurrentDate     =   40157
      End
      Begin MSComCtl2.DTPicker dtpTo 
         Height          =   375
         Left            =   4080
         TabIndex        =   12
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   131596289
         CurrentDate     =   40157
      End
      Begin VB.Label Label9 
         Caption         =   "Date From"
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Date To"
         Height          =   255
         Left            =   3120
         TabIndex        =   13
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.PictureBox Picture5 
      Height          =   255
      Left            =   1680
      Picture         =   "frmPayment.frx":0000
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   8
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox txtLocation 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   600
      TabIndex        =   7
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox TXTIDNO 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   720
      TabIndex        =   6
      Top             =   1560
      Width           =   2175
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   285
      Left            =   4680
      TabIndex        =   5
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox txtAPaid 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1320
      TabIndex        =   4
      Text            =   "0.00"
      Top             =   3480
      Width           =   3015
   End
   Begin MSComctlLib.ListView Lvwitems 
      Height          =   3255
      Left            =   240
      TabIndex        =   37
      Top             =   4440
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   5741
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "SNo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Names"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Quantity(Kgs)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Gross"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Deductions"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "NetPay"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblNetP 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   495
      Left            =   4560
      TabIndex        =   45
      Top             =   7920
      Width           =   4095
   End
   Begin VB.Label lblTKgs 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0.00"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   9000
      TabIndex        =   36
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Total Kgs"
      Height          =   375
      Left            =   7680
      TabIndex        =   35
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Gross"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   34
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label Label19 
      Caption         =   "Bank :"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   33
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label18 
      Caption         =   "Name :"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   32
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label17 
      Caption         =   "Box :"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   31
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label16 
      Caption         =   "Telephone :"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   30
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label15 
      Caption         =   "Transport :"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   29
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label14 
      Caption         =   "SNo :"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   28
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label13 
      Caption         =   "Branch :"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   27
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Account Number :"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6960
      TabIndex        =   26
      Top             =   1080
      Width           =   1800
   End
   Begin VB.Label Label11 
      Caption         =   "Loc :"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   25
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lblNPay 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   495
      Left            =   6720
      TabIndex        =   24
      Top             =   2880
      Width           =   4095
   End
   Begin VB.Label Label12 
      Caption         =   "ID:"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   23
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label8 
      Caption         =   "Amount Paid"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label lblGross 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label lblDeductions 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4320
      TabIndex        =   1
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label Label6 
      Caption         =   "Deductions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   0
      Top             =   3000
      Width           =   855
   End
End
Attribute VB_Name = "frmPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdPay_Click()
Startdate = DateSerial(year(dtpFrom), month(dtpFrom), 1)
Enddate = DateSerial(year(dtpTo), month(dtpTo) + 1, 1 - 1)
Dim j As Integer
j = 1
For j = 1 To Lvwitems.ListItems.Count

Set li = Lvwitems.ListItems(j)

If CCur(Lvwitems.ListItems(j).SubItems(5)) > 0 Then
'//Update deductions
Set cn = New ADODB.Connection
sql = "d_sp_SupplierDeduct " & li & ",'" & dtpTo & "','Advance'," & CCur(Lvwitems.ListItems(j).SubItems(5)) & ",'" & Startdate & "','" & Enddate & "'," & year(dtpTo) & ",'" & User & "','Paid',''"
oSaccoMaster.ExecuteThis (sql)

'//Update Payments
' d_sp_PayFarmer @SNo bigint, @SDate varchar(12),@EDate varchar(12)  AS
oSaccoMaster.ExecuteThis ("d_sp_PayFarmer " & li & ",'" & dtpFrom & "','" & dtpTo & "'")
Else
End If
Next j
MsgBox "Records updated successively!!"
Lvwitems.ListItems.Clear
End Sub

Private Sub cmdsave_Click()
 If Trim(txtSNo) = "" Then
  MsgBox "Please enter supplier number."
    txtSNo.SetFocus
    Exit Sub
End If

If CCur(Trim$(lblNPay)) < 1 Then
 MsgBox "The net pay is not payable."
    txtAPaid.SetFocus
    Exit Sub
End If

If CCur(lblNPay) < CCur(txtAPaid) Then
 MsgBox "The net pay is less than amount you want to pay."
    txtAPaid.SetFocus
    Exit Sub
End If

Startdate = DateSerial(year(dtpFrom), month(dtpFrom), 1)
Enddate = DateSerial(year(dtpTo), month(dtpTo) + 1, 1 - 1)

'//Update deductions
Set cn = New ADODB.Connection
sql = "d_sp_SupplierDeduct " & txtSNo & ",'" & dtpTo & "','Advance'," & CCur(txtAPaid) & ",'" & Startdate & "','" & Enddate & "'," & year(dtpTo) & ",'" & User & "','Paid',''"
oSaccoMaster.ExecuteThis (sql)

'//Update Payments
' d_sp_PayFarmer @SNo bigint, @SDate varchar(12),@EDate varchar(12)  AS
oSaccoMaster.ExecuteThis ("d_sp_PayFarmer " & txtSNo & ",'" & dtpFrom & "','" & dtpTo & "'")

MsgBox "Record Saved Successively!!"
txtSNo = ""
txtSNo_Validate True
End Sub

Private Sub cmdShow_Click()
Dim total As Currency
total = 0
If Option1 = True Then
Set rst = oSaccoMaster.GetRecordset("d_sp_PayTransport '" & cboTCode & "'")
Lvwitems.ListItems.Clear


While Not rst.EOF
'If Not IsNull(rs.Fields(0)) Then Exit Sub


    Set li = Lvwitems.ListItems.Add(, , rst.Fields(0))
                        li.SubItems(1) = rst.Fields(1) & ""
Set rs = oSaccoMaster.GetRecordset("d_sp_SupNet " & rst.Fields(0) & ",'" & dtpFrom & "','" & dtpTo & "', 0")

If Not IsNull(rs.Fields(0)) Then
lblTKgs = rs.Fields(0)
Else
lblTKgs = "0.00"
End If

If Not IsNull(rs.Fields(1)) Then
lblGross = rs.Fields(1)
Else
lblGross = "0.00"
End If



'txtSNo = rst.Fields(0)
Set rs = oSaccoMaster.GetRecordset("d_sp_SupNet " & rst.Fields(0) & ",'" & dtpFrom & "','" & dtpTo & "', 1")
If Not IsNull(rs.Fields(0)) Then
lblDeductions = rs.Fields(0)
Else
lblDeductions = "0.00"
End If

lblNPay = Format((CCur(lblGross) - CCur(lblDeductions)), "#,##0.00")

                        li.SubItems(2) = lblTKgs & ""
                        li.SubItems(3) = lblGross & ""
                        li.SubItems(4) = lblDeductions & ""
                        li.SubItems(5) = lblNPay & ""
                        total = CCur(lblNPay) + total
                rst.MoveNext
            Wend

lblNetP = "Net Pay " & Format(total, "#,##0.00")
End If

If Option2 = True Then
Set rst = oSaccoMaster.GetRecordset("d_sp_PayTransport '" & cboLocation & "'")
Lvwitems.ListItems.Clear
'Dim total As Currency
'total = 0

While Not rst.EOF
'If Not IsNull(rs.Fields(0)) Then Exit Sub


    Set li = Lvwitems.ListItems.Add(, , rst.Fields(0))
                        li.SubItems(1) = rst.Fields(1) & ""
Set rs = oSaccoMaster.GetRecordset("d_sp_SupNet " & rst.Fields(0) & ",'" & dtpFrom & "','" & dtpTo & "', 0")

If Not IsNull(rs.Fields(0)) Then
lblTKgs = rs.Fields(0)
Else
lblTKgs = "0.00"
End If

If Not IsNull(rs.Fields(1)) Then
lblGross = rs.Fields(1)
Else
lblGross = "0.00"
End If



'txtSNo = rst.Fields(0)
Set rs = oSaccoMaster.GetRecordset("d_sp_SupNet " & rst.Fields(0) & ",'" & dtpFrom & "','" & dtpTo & "', 1")
If Not IsNull(rs.Fields(0)) Then
lblDeductions = rs.Fields(0)
Else
lblDeductions = "0.00"
End If

lblNPay = Format((CCur(lblGross) - CCur(lblDeductions)), "#,##0.00")

                        li.SubItems(2) = lblTKgs & ""
                        li.SubItems(3) = lblGross & ""
                        li.SubItems(4) = lblDeductions & ""
                        li.SubItems(5) = lblNPay & ""
                        total = CCur(lblNPay) + total
                rst.MoveNext
            Wend

lblNetP = "Net Pay " & Format(total, "#,##0.00")
End If


End Sub

Private Sub cmdShowNet_Click()
txtSNo_Validate True
End Sub

Private Sub Form_Load()
dtpFrom = DateSerial(year(Get_Server_Date), month(Get_Server_Date), 1)
dtpTo = Format(Get_Server_Date, "dd/mm/yyyy")

Set rs = oSaccoMaster.GetRecordset("SELECT TransCode FROM   d_Transporters WHERE TransCode <> ''")
While Not rs.EOF
'If IsNull(rs.Fields(0)) Then Exit Sub
    If Not IsNull(rs.Fields(0)) Then cboTCode.AddItem (rs.Fields(0))
                rs.MoveNext
            Wend
            
sql = ""
sql = sql & "SELECT DISTINCT Location From d_Suppliers WHERE(Location <> '')ORDER BY Location"
Set rs = oSaccoMaster.GetRecordset(sql)
While Not rs.EOF
'If IsNull(rs.Fields(0)) Then Exit Sub
    If Not IsNull(rs.Fields(0)) Then cboLocation.AddItem (rs.Fields(0))
                rs.MoveNext
            Wend
'Enddate = DateSerial(Year(txtransdate), month(txtransdate) + 1, 1 - 1)
End Sub


Private Sub Option1_Click()
cboLocation.Visible = False
cboTCode.Visible = True
End Sub

Private Sub Option2_Click()
cboLocation.Visible = True
cboTCode.Visible = False
End Sub

Private Sub Option3_Click()
cboLocation.Visible = False
cboTCode.Visible = False
End Sub

Private Sub Picture5_Click()
Me.MousePointer = vbHourglass
        frmSearchSupplier.Show vbModal
        txtSNo = sel
        txtSNo_Validate True
        Me.MousePointer = 0
End Sub

Private Sub txtSNo_Validate(Cancel As Boolean)
Dim a, t As Boolean
If Trim(txtSNo) = "" Then
txtName = ""
txtTransport = ""
txtAccNo = ""
txtBank = ""
txtBBranch = ""
txtLocation = ""
txtTelNo = ""
txtBox = ""
TXTIDNO = ""
lblDeductions = "0.00"
lblGross = "0.00"
lblNPay = "0.00"
lblTKgs = "0.00"
            txtSNo.SetFocus
        Exit Sub
    End If
    
Set rs = New ADODB.Recordset
sql = "d_sp_SupplierEnquiry " & txtSNo & ""
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
' [Names], AccNo, Bcode, BBranch, Location, PhoneNo, Address + ' ' + Town AS ADDRESS
'FROM         d_Suppliers WHERE SNo= @SNo
If Not IsNull(rs.Fields(0)) Then txtName = rs.Fields(0)
If Not IsNull(rs.Fields(1)) Then txtAccNo = rs.Fields(1)
If Not IsNull(rs.Fields(2)) Then txtBank = rs.Fields(2)
If Not IsNull(rs.Fields(3)) Then txtBBranch = rs.Fields(3)
If Not IsNull(rs.Fields(4)) Then txtLocation = rs.Fields(4)
If Not IsNull(rs.Fields(5)) Then txtTelNo = rs.Fields(5)
If Not IsNull(rs.Fields(6)) Then txtBox = rs.Fields(6)
If Not IsNull(rs.Fields(7)) Then TXTIDNO = rs.Fields(7)

 
Set rs = New ADODB.Recordset
sql = "d_sp_TransName " & txtSNo & ""
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
If Not IsNull(rs.Fields(0)) Then txtTransport = rs.Fields(0)

End If


'Startdate = DateSerial(Year(txtransdate), month(txtransdate), 1)
'Enddate = DateSerial(Year(txtransdate), month(txtransdate) + 1, 1 - 1)

Set rs = oSaccoMaster.GetRecordset("d_sp_SupNet " & txtSNo & ",'" & dtpFrom & "','" & dtpTo & "', 0")

If Not IsNull(rs.Fields(0)) Then
lblTKgs = rs.Fields(0)
Else
lblTKgs = "0.00"
End If

If Not IsNull(rs.Fields(1)) Then
lblGross = rs.Fields(1)
Else
lblGross = "0.00"
End If

Set rs = oSaccoMaster.GetRecordset("d_sp_SupNet " & txtSNo & ",'" & dtpFrom & "','" & dtpTo & "', 1")
If Not IsNull(rs.Fields(0)) Then
lblDeductions = rs.Fields(0)
Else
lblDeductions = "0.00"
End If

lblNPay = Format((CCur(lblGross) - CCur(lblDeductions)), "#,##0.00")
txtAPaid = lblNPay
  'SNo
Else
txtName = ""
txtAccNo = ""
txtBank = ""
txtBBranch = ""
txtLocation = ""
txtTelNo = ""
txtBox = ""
TXTIDNO = ""
lblDeductions = ""
lblGross = ""
lblNPay = ""
lblTKgs = ""
End If
End Sub
