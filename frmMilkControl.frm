VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMilkControl 
   BackColor       =   &H00C0FFFF&
   Caption         =   "MILK SALES AND DISPATCH"
   ClientHeight    =   7920
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   13185
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   13185
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "New Debtor"
      Height          =   375
      Left            =   3960
      TabIndex        =   48
      Top             =   7440
      Width           =   1455
   End
   Begin VB.ComboBox cboVehicle 
      Height          =   405
      Left            =   4800
      TabIndex        =   47
      Top             =   3120
      Width           =   2295
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   3255
      Left            =   7560
      TabIndex        =   46
      Top             =   4440
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   5741
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777152
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "DCode"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "DName"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Quantity"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CheckBox chkPay 
      Caption         =   "Make payments"
      Height          =   375
      Left            =   5760
      TabIndex        =   44
      Top             =   6840
      Width           =   1815
   End
   Begin VB.CheckBox chprint 
      Caption         =   "Use LPT1 Printer"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   240
      TabIndex        =   43
      Top             =   5880
      Width           =   3255
   End
   Begin VB.CheckBox chkPrint 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "Print Receipt"
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   6000
      TabIndex        =   41
      Top             =   0
      Width           =   1695
   End
   Begin VB.ComboBox ports 
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "frmMilkControl.frx":0000
      Left            =   1320
      List            =   "frmMilkControl.frx":0010
      TabIndex        =   39
      Text            =   "\\127.0.0.1\E-PoS 80mm Thermal Printer"
      Top             =   6360
      Width           =   2175
   End
   Begin VB.TextBox txtreceiveby 
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   5040
      TabIndex        =   37
      Top             =   4560
      Width           =   2055
   End
   Begin VB.TextBox txttotal 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   4800
      TabIndex        =   36
      Top             =   2040
      Width           =   2055
   End
   Begin VB.TextBox txtamountp 
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   4800
      TabIndex        =   33
      Top             =   1080
      Width           =   2295
   End
   Begin VB.CommandButton cmdstatement 
      Caption         =   "Debtors Statement"
      Height          =   375
      Left            =   960
      TabIndex        =   32
      Top             =   7440
      Width           =   2415
   End
   Begin VB.CommandButton cmdreprint 
      Caption         =   "Reprint"
      Height          =   375
      Left            =   3960
      TabIndex        =   31
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton cmdnewsearch 
      Caption         =   "New "
      Height          =   285
      Left            =   4080
      TabIndex        =   30
      Top             =   120
      Width           =   615
   End
   Begin VB.CheckBox chkapp 
      Caption         =   "Cess Applicable"
      Height          =   285
      Left            =   3000
      TabIndex        =   29
      Top             =   4440
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Height          =   285
      Left            =   3840
      Picture         =   "frmMilkControl.frx":002C
      ScaleHeight     =   225
      ScaleWidth      =   195
      TabIndex        =   19
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox Picture3 
      Height          =   285
      Left            =   2760
      Picture         =   "frmMilkControl.frx":02EE
      ScaleHeight     =   225
      ScaleWidth      =   195
      TabIndex        =   18
      Top             =   3000
      Width           =   255
   End
   Begin VB.TextBox txtdcode 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1560
      TabIndex        =   17
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox txtRefNo 
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   1560
      TabIndex        =   13
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   375
      Left            =   1320
      TabIndex        =   12
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   2640
      TabIndex        =   11
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   6840
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker DTPDispatchDate 
      Height          =   375
      Left            =   1560
      TabIndex        =   9
      Top             =   600
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      Format          =   122224641
      CurrentDate     =   40105
   End
   Begin VB.TextBox txtVariance 
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   1560
      TabIndex        =   8
      Top             =   2520
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox txtIntake 
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   1560
      TabIndex        =   7
      Top             =   2040
      Width           =   2295
   End
   Begin VB.TextBox txtDipping 
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   1560
      TabIndex        =   6
      Top             =   1560
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox txtDispatch 
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   1560
      TabIndex        =   5
      Top             =   1080
      Width           =   2295
   End
   Begin MSComDlg.CommonDialog cdgPrint 
      Left            =   8400
      Top             =   6840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FileName        =   "c:\receipt.txt"
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3855
      Left            =   7200
      TabIndex        =   45
      Top             =   480
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   6800
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   65280
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Dcode"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "DQuantity"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label19 
      Caption         =   "Vehicle No"
      Height          =   255
      Left            =   5040
      TabIndex        =   42
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label18 
      Caption         =   "Printer Port"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   40
      Top             =   6360
      Width           =   1095
   End
   Begin VB.Label Label17 
      Caption         =   "Receive by"
      Height          =   255
      Left            =   5040
      TabIndex        =   38
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label14 
      Caption         =   "Cumulative Variance"
      Height          =   255
      Left            =   4920
      TabIndex        =   35
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Label13 
      Caption         =   "Amounts payable"
      Height          =   375
      Left            =   4800
      TabIndex        =   34
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label16 
      Caption         =   "Cess Acc Dr"
      Height          =   255
      Left            =   360
      TabIndex        =   28
      Top             =   4920
      Width           =   2415
   End
   Begin VB.Label Label15 
      Caption         =   "Debtors Acc Cr"
      Height          =   255
      Left            =   360
      TabIndex        =   27
      Top             =   5400
      Width           =   2415
   End
   Begin VB.Label cessdr 
      Height          =   255
      Left            =   3000
      TabIndex        =   26
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Label cesscr 
      Height          =   255
      Left            =   3000
      TabIndex        =   25
      Top             =   5400
      Width           =   1575
   End
   Begin VB.Label Label12 
      Caption         =   "CESS ACCOUNTS"
      Height          =   255
      Left            =   360
      TabIndex        =   24
      Top             =   4440
      Width           =   2175
   End
   Begin VB.Label Label11 
      Height          =   255
      Left            =   3000
      TabIndex        =   23
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label Label10 
      Height          =   255
      Left            =   3000
      TabIndex        =   22
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label Label9 
      Caption         =   "Acc Cr"
      Height          =   255
      Left            =   360
      TabIndex        =   21
      Top             =   3960
      Width           =   2415
   End
   Begin VB.Label Label6 
      Caption         =   "Acc Dr"
      Height          =   255
      Left            =   360
      TabIndex        =   20
      Top             =   3480
      Width           =   2415
   End
   Begin VB.Label lblDebtors 
      AutoSize        =   -1  'True
      Height          =   285
      Left            =   3000
      TabIndex        =   16
      Top             =   3000
      Width           =   60
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "Debtors Code :"
      Height          =   285
      Left            =   120
      TabIndex        =   15
      Top             =   3000
      Width           =   1410
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "Reference No. :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   1395
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "Variance :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "Intake :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   630
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "Dispatch : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   930
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "Dispatch Date :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1365
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "Dipping :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Menu mnuinvoice 
      Caption         =   "Invoice"
   End
End
Attribute VB_Name = "frmMilkControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Price As Currency
Dim capp As Integer
Dim crate As Double
Dim rsq As New Recordset
Dim milksup As Double
Dim amtpayable As Double
Dim receipno As Double
Dim dispatchby As Double
Dim qty As Double

Private Sub chkPay_Click()
If chkPay.value = vbChecked Then
ListView2.Visible = True
Startdate = DateSerial(year(DTPDispatchDate), month(DTPDispatchDate), 1)
Enddate = DateSerial(year(DTPDispatchDate), month(DTPDispatchDate) + 1, 1 - 1)
sql = ""
sql = "set dateformat dmy SELECT d.DCode, d.DName, m.DispQnty,m.DispDate FROM  d_Debtors AS d INNER JOIN d_MilkControl AS m ON d.DCode = m.dcode WHERE     (DispDate = '" & DTPDispatchDate & "') and Locations='" & cboVehicle & "' and status=0"
Set rs = oSaccoMaster.GetRecordset(sql)
If rs.EOF Then
Exit Sub
End If
ListView2.ListItems.Clear

While Not rs.EOF
If Not IsNull(rs.Fields(0)) Then
Set li = ListView2.ListItems.Add(, , rs.Fields(0))
End If
                    If Not IsNull(rs.Fields(1)) Then li.SubItems(1) = rs.Fields(1) & ""
                        
                    If Not IsNull(rs.Fields(2)) Then li.SubItems(2) = rs.Fields(2) & ""
                     If Not IsNull(rs.Fields(3)) Then li.SubItems(3) = rs.Fields(3) & ""
                    'If Not IsNull(rs.Fields(4)) Then li.SubItems(4) = rs.Fields(4) & ""
                    'If Not IsNull(rs.Fields(5)) Then li.SubItems(5) = rs.Fields(5) & ""
rs.MoveNext

Wend
Else
ListView2.Visible = False
End If
End Sub




Private Sub chprint_Click()
ports.Clear
ports = ""
'//If the drivers are installed it won't matter whether the Port is indicated
' or not it will just work.

If chprint.value = vbChecked Then
ports.AddItem "LPT1"
ports = "LPT1"
ports.AddItem "LPT2"
ports.AddItem "LPT3"
ports.AddItem "LPT4"
ports.AddItem "LPT5"
Else
'Share the printer first the use of 127.0.0.1 which is
'standard IP address for a loopback network connection
'instead of getting the computer name or IP Address
'
Dim prnPrinter As Printer
Dim pr As String
ports.Clear

For Each prnPrinter In Printers
   If InStr(prnPrinter.DeviceName, "\\") Then
    ports.AddItem prnPrinter.DeviceName
    If InStr(prnPrinter.DeviceName, "G") Then
    ports.Text = prnPrinter.DeviceName
    End If
    Else
    ports.AddItem "\\127.0.0.1\" & prnPrinter.DeviceName
    If InStr(prnPrinter.DeviceName, "G") Then
    ports.Text = "\\127.0.0.1\" & prnPrinter.DeviceName
    End If
    End If
   
   
Next
End If
'This code will work only if there is a connection e.g LAN or modem.
'It is not a must that it is an internet connection because
'computer's network interface card has to be functional


End Sub

Private Sub cmdedit_Click()
txtRefNo.Locked = True

    txtDipping.Locked = False
    txtDispatch.Locked = False
    txtIntake.Locked = False
    txtVariance.Locked = False

    cmdnew.Enabled = False
    cmdsave.Enabled = True
    cmdedit.Enabled = False
    
End Sub

Private Sub cmdNew_Click()
    'txtDipping.Locked = False
    txtDispatch.Locked = False
    txtIntake.Locked = True
    txtVariance.Locked = False
    'txtDipping.Locked = True
    txtDispatch = ""
    txtVariance = ""
    txtdcode = ""
    lblDebtors = ""
    DTPDispatchDate = Get_Server_Date
  
    
    
    ListView2.Visible = False
    chkPay.value = vbUnchecked
    cmdnew.Enabled = False
    cmdsave.Enabled = True
    cmdedit.Enabled = False
    
End Sub

Private Sub cmdnewsearch_Click()
Dim rsr As New ADODB.Recordset
Dim rsg As New ADODB.Recordset
Dim I As Object
Dim Mylength As Integer
'//if this record is new then look for receipts no

''//clear all textboxes





'mysql = ""
'mysql = "set dateformat dmy select GenerateReceiptno from param"
sql = ""
sql = "set dateformat dmy select GenerateReceiptno from param"
Set rsg = oSaccoMaster.GetRecordset(sql)
If Not rsg.EOF Then
    ''''check check
    If rsg!GenerateReceiptno = True Then
    
        sql = ""
        sql = "select * from Receiptno where receiptno like 'RF-%' order by Receipthnoid desc"
        
        Set rsr = oSaccoMaster.GetRecordset(sql)
        
        If Not rsr.EOF Then
            Mylength = CInt(Mid(rsr!ReceiptNo, 5, 10))
            Mylength = Mylength + 1
            txtRefNo = Padding(Mylength)
            txtRefNo = "RF-" & txtRefNo
        Else
            Mylength = 1
            txtRefNo = "RF-" & Padding(Mylength)
            
        End If
Else
    ''//receiptno  will be keyed in
End If
End If
End Sub

Private Sub cmdreprint_Click()
STRFORMULA = "{d_MilkControl.RefNo}='" & txtRefNo & "'"
    reportname = "milkinvoice.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, title
End Sub

Private Sub cmdsave_Click()
If txtdcode = "" Then
MsgBox "Debtors code cannot be blank; input an existing one", vbCritical
Exit Sub
End If
If txtDispatch = "" Then
    MsgBox "Please enter the dispatch quantity."
        txtDispatch.SetFocus
    Exit Sub
End If

If txtDipping = "" Then
    MsgBox "Please enter the dipping quantity."
        txtDipping.SetFocus
    Exit Sub
End If

If txtIntake = "" Then
    MsgBox "Please enter the intake quantity."
        txtIntake.SetFocus
    Exit Sub
End If

'If txtVariance = "" Then
'    MsgBox "Please enter the variance quantity."
'        txtVariance.SetFocus
'    Exit Sub
'End If



If txtRefNo = "" Then
    MsgBox "Please enter the reference number."
        txtRefNo.SetFocus
    Exit Sub
End If
'//check if the dispatch is greater than the dipping
If CDbl(txtDipping) < CDbl(txtDispatch) Then 'raiise an alarm
MsgBox "You cannot take more you have in the tank", vbCritical
Exit Sub
End If
Dim Debit As String
Dim Credit As String

sql = ""
    sql = "SET      dateformat dmy     SELECT     *     FROM         d_MilkControl    WHERE     DispDate = '" & DTPDispatchDate & "' and DispQnty = '" & txtDispatch & "'and dcode = '" & txtdcode & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
    MsgBox "You have already dispatch for that day", vbInformation
    Exit Sub
    End If


'Dim Price As Currency

'Set rs = oSaccoMaster.GetRecordset("d_sp_getAccName '" & lblDebtors & "'")
'If IsNull(rs.Fields(0)) Then
'    MsgBox "The debtors account not set. " & vbNewLine & "Please contact the accountant to set GL for " & lblDebtors
'        Exit Sub
'End If
'
Debit = Label10
'
'Set rs = oSaccoMaster.GetRecordset("d_sp_getAccName 'Milk sale'")
'If IsNull(rs.Fields(0)) Then
'    MsgBox "The Creditors account not set. " & vbNewLine & "Please contact the accountant to set GL for milk sales"
'        Exit Sub
'End If
'
Credit = Label11

    

    If Not Save_GLTRANSACTION(Format(DTPDispatchDate, "dd/mm/yyyy"), (CCur(Price) * CCur(txtDispatch)), Debit, Credit, "Milk Sales ", txtRefNo, User, ErrorMessage, "Milk Sales", 1, 1, txtRefNo, transactionNo, "", "", 0) Then
            If ErrorMessage <> "" Then
                MsgBox ErrorMessage, vbInformation, Me.Caption
                ErrorMessage = ""
            End If
    End If
    
    If capp = 1 Then
    
    If Not Save_GLTRANSACTION(Format(DTPDispatchDate, "dd/mm/yyyy"), (CCur(crate) * CCur(txtDispatch)), cessdr, cesscr, "Cess Deductions ", txtRefNo, User, ErrorMessage, "Cess Deductions", 1, 1, txtRefNo, transactionNo, "", "", 0) Then
            If ErrorMessage <> "" Then
                MsgBox ErrorMessage, vbInformation, Me.Caption
                ErrorMessage = ""
            End If
    End If
    
    End If
        
'd_sp_MilkControl @DispDate char(10), @DipsQnty float,@DipQnty float,@InQnty float,@VarQnty float,@Price char(10),@RefNo varchar(35),@CreditAcc varchar(35),@DebitAcc varchar(35),@AuditID varchar (50)
Set rs = New ADODB.Recordset
sql = "d_sp_MilkControl  '" & DTPDispatchDate & "'," & txtDispatch & "," & txtDipping & "," & txtIntake & ",'" & txtVariance & "'," & Price & ",'" & txtRefNo & "','" & Credit & "','" & Debit & "','" & User & "','" & txtdcode & "','" & cboVehicle & "'"
oSaccoMaster.ExecuteThis (sql)

'//subtract from the dispatch table

    sql = ""
    sql = "SET      dateformat dmy     SELECT     ID, Intake,transdate     FROM         d_dispatch    WHERE     transdate = '" & DTPDispatchDate & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If rs.EOF Then
sql = ""
sql = "set dateformat dmy INSERT INTO d_dispatch (Transdate, descrip, Intake, dipping, dispatch, auditid, auditdate)values ('" & DTPDispatchDate.value & "','Dispatch'," & txtDipping & "," & CDbl(txtDipping) - CDbl(txtDispatch) & "," & CDbl(txtDispatch) & ",'" & User & "','" & Get_Server_Date & "')"
oSaccoMaster.ExecuteThis (sql)
'sql = ""
'sql = "set dateformat dmy INSERT INTO d_DetailDispatch (DCode,Name, Transdate, dispatch, auditid,auditdate)values ('" & txtdcode & "','','" & DTPDispatchDate.value & "','" & txtDispatch & "','" & User & "','" & Get_Server_Date & "')"
'oSaccoMaster.ExecuteThis (sql)

'    sql = ""
'    sql = "SET      dateformat dmy     SELECT DCode, DName FROM d_Debtors    WHERE     DCode = " & txtdcode & ""
'    Set rsx = oSaccoMaster.GetRecordset(sql)
'    If rsx.EOF Then
'    sql = ""
'    sql = "set dateformat dmy INSERT INTO d_DetailDispatch (DName)values ('" & DTPDispatchDate.value & "','Dispatch'," & txtDipping & "," & CDbl(txtDipping) - CDbl(txtDispatch) & "," & CDbl(txtDispatch) & ",'" & User & "','" & Get_Server_Date & "')"
'    oSaccoMaster.ExecuteThis (sql)
'
'    End If

Else
sql = ""
sql = "set dateformat dmy UPDATE    d_dispatch  SET   dipping =" & CDbl(txtDipping) - CDbl(txtDispatch) & ",dispatch=" & txtDispatch & "  WHERE     (Transdate = '" & DTPDispatchDate & "')"
oSaccoMaster.ExecuteThis (sql)
'sql = ""
'sql = "set dateformat dmy UPDATE    d_DetailDispatch  SET   dispatch =" & txtDispatch & "  WHERE     (Transdate = '" & DTPDispatchDate & "') and DCode='" & txtdcode & "'"
'oSaccoMaster.ExecuteThis (sql)
End If
'Dim rsd As Recordset
'sql = ""
'sql = "select dispatch1, dispatch2, dispatch3, dispatch4, dispatch5  from d_DetailDispatch where Transdate='" & DTPDispatchDate & "' "
'
'Set rsd = oSaccoMaster.GetRecordset(sql)
'
'Dim DName As Double, two As Double, three As Double, four As Double, five As Double
'one = rsd.Fields(0)
'two = rsd!dispatch2
'three = rsd!dispatch3
'four = rsd!dispatch4
'five = rsd!dispatch5
'If one = "" Then
         Dim DName As String
          Set rs = New ADODB.Recordset
          sql = "SELECT DName from d_Debtors where DCode='" & txtdcode & "'"
          Set rs = oSaccoMaster.GetRecordset(sql)
          If Not rs.EOF Then
          DName = rs!DName
          End If

'sql = ""
'sql = "set dateformat dmy INSERT INTO d_DetailDispatch (DCode, Transdate,Name, dispatch, auditid, auditdate)values ('" & txtdcode & "','" & DTPDispatchDate.value & "','" & DName & "','" & txtDispatch & "','" & User & "','" & Get_Server_Date & "')"
'oSaccoMaster.ExecuteThis (sql)
''Else
'..............INSERT DAILY INTAKE FOR DEBTORS ONLY.........................
'Dim rsd As Recordset
'  sql = ""
'  sql = "set dateformat DMY select  DCode, Transdate,Name, dispatch, auditid, auditdate from d_DetailDispatch where DCode= 'Intake' and dispatch='" & txtIntake & "' and Transdate='" & DTPDispatchDate & "' "
'  Set rsd = New ADODB.Recordset
'  rsd.Open sql, cn
'  If rsd.EOF Then
'    sql = ""
'    sql = "set dateformat dmy INSERT INTO d_DetailDispatch (DCode, Transdate,Name, dispatch, auditid, auditdate)values ('Intake','" & DTPDispatchDate.value & "','Intake','" & txtIntake & "','" & User & "','" & Get_Server_Date & "')"
'    oSaccoMaster.ExecuteThis (sql)
'  Else
'  sql = ""
'  sql = "set dateformat dmy UPDATE d_DetailDispatch SET dispatch='" & txtIntake & "'  WHERE DCode='Intake' and   (Transdate = '" & DTPDispatchDate & "')"
'  oSaccoMaster.ExecuteThis (sql)
'  End If
'..............END OF  DAILY INTAKE INSERT FOR DEBTORS ONLY.........................
mysql = "set dateformat dmy Insert into Receiptno(Receiptno,Auditdate,auditid)values('" & txtRefNo & "','" & Format(Get_Server_Date, "dd/MM/yyyy") & "','" & User & "')"
oSaccoMaster.ExecuteThis (mysql)
If chkPrint = vbChecked Then
    
If chkPrint = vbChecked Then
    
'/*Print out
 Dim fso, chkPrinter, txtFile
    Dim ttt
     Dim escFeedAndCut As String
     Dim escNewLine As String
     Dim escUnerLineON As String
     Dim escUnerLineOnX2 As String
     Dim escUnerLineOff As String
     Dim escBoldOn As String
     Dim escBoldOff As String
     Dim escNegativeOn As String
     Dim escNegativeOff As String
     Dim esc8CpiOn As String
     Dim esc8CPiOff As String
     Dim esc16Cpi As String
     Dim esc20Cpi As String
     Dim escAlignLeft As String
     Dim escAlignCenter As String
     Dim escAlignRight As String
    
     
        escNewLine = Chr(10) '//New Line (LF Line Feed)
        escUnerLineON = Chr(27) + Chr(45) + Chr(1) '//Unerline on
        escUnerLineOnX2 = Chr(27) + Chr(45) + Chr(1) '//Unerline on X2
        escUnerLineOff = Chr(27) + Chr(45) + Chr(0) '//unerline off
        escBoldOn = Chr(27) + Chr(69) + Chr(1) '//Bold on
        escBoldOff = Chr(27) + Chr(69) + Chr(0) '//Bold off
        escNegativeOn = Chr(29) + Chr(66) + Chr(1) '//White on Black on
        escNegativeOff = Chr(29) + Chr(66) + Chr(0) '//white on
        esc8CpiOn = Chr(29) + Chr(33) + Chr(16) '//Font Size X2 on
        esc8CPiOff = Chr(29) + Chr(33) + Chr(0) '//Font size X2 off
        esc16Cpi = Chr(27) + Chr(77) + Chr(48) '//Font A - Normal Size
        esc20Cpi = Chr(27) + Chr(77) + Chr(49) '//Font B - Small Font
        escAlignLeft = Chr(27) + Chr(97) + Chr(48) '//Align text to the left
        escAlignCenter = Chr(27) + Chr(97) + Chr(49) '//Align text to the center
        escAlignRight = Chr(27) + Chr(97) + Chr(50) '//Align text to the right
        escFeedAndCut = Chr(29) + Chr(86) + Chr(65) '//Partial cut and feed
       'cdgPrint.Filter = "*.csv|*.txt"
        'cdgPrint.ShowSave
        Dim PORT As String
   '     PORT = ports
        'ttt = "LPT1" 'LPT1
        ttt = ports
        'ttt = "D:\PROJECTS\FOSA\DAILY" & Date & ""
        Set fso = CreateObject("Scripting.FileSystemObject")
        'Set chkPrinter = fso.GetFile(ttt)
        
    Set txtFile = fso.CreateTextFile(ttt, True)
    txtFile.WriteLine "         " & cname & ""
    txtFile.WriteLine "         Address :" & paddress & ""
    txtFile.WriteLine "         Phone :" & Phone & ""
    txtFile.WriteLine "         Email :" & Email & ""
    'txtfile.WriteLine " " & txtSNo
    
    txtFile.WriteLine "          Delivery Note"
    txtFile.WriteLine "**********************************************"
        
    Set rs2 = New ADODB.Recordset
    sql = "d_sp_ReceiptNumber"
    Set rs2 = oSaccoMaster.GetRecordset(sql)
    
    Dim RNumber As String
    'RNumber = rs2.Fields(0)
    If Not IsNull(rs2.Fields(0)) Then RNumber = rs2.Fields(0)
    'Else
    'RNumber = "0"
    'End If
    
    txtFile.WriteLine "CsNO :" & txtRefNo
    txtFile.WriteLine "To :" & lblDebtors
   txtFile.WriteLine " *********************************************************************"
    txtFile.WriteLine "DESCRIPTION " & vbTab & "" & vbTab & "value"
    sql = "SELECT     d.DCode, d.DName, SUM(m.DispQnty) AS quantity FROM         d_Debtors AS d INNER JOIN d_MilkControl AS m ON d.DCode = m.dcode WHERE     (DispDate = '" & DTPDispatchDate & "') GROUP BY d.DCode, d.DName"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
    'txtamountp = rs!quantity*
   ' Dim milksup As Double
'    Dim amtpayable As Double
'    Dim receipno As Double
'    Dim dispatchby As Double
   ' Exit Sub
   ' End If
'    Set rs = New ADODB.Recordset
'    sql = "d_sp_TotalMonth " & txtSNo & ",'" & Startdate & "','" & DTPMilkDate & "'"
'    Set rs = oSaccoMaster.GetRecordset(sql)
'    If Not rs.EOF Then
'    If Not IsNull(rs.Fields(0)) Then CummulKgs = rs.Fields(0)
'    Else
'    CummulKgs = "0.00"
'    End If
    txtFile.WriteLine "Milk supplied :" & vbTab & "" & vbTab & " " & rs!Quantity & ""
    txtFile.WriteLine "Amount Payable :" & vbTab & "  " & txtamountp
    txtFile.WriteLine "Receipt Number :" & vbTab & "  " & txtRefNo
    txtFile.WriteLine "Dispatched by :" & vbTab & " " & username & ""
    
    txtFile.WriteLine "---------------------------------------"
    End If
'    txtFile.WriteLine "Receipt Number :" & RNumber
'    txtFile.WriteLine "TRANSPORTER :" & TRANSPORTER
    txtFile.WriteLine "Vehicle No :" & cboVehicle
    txtFile.WriteLine "Received by :" & txtreceiveby
    txtFile.WriteLine "---------------------------------------"
    txtFile.WriteLine "     Date :" & Format(DTPDispatchDate, "dd/mm/yyyy") & " ,Time : " & Format(Time, "hh:mm:ss AM/PM")
    txtFile.WriteLine "" & motto & ""
    txtFile.WriteLine "---------------------------------------"
    'If chkComment.value = vbChecked Then
        'txtFile.WriteLine txtComment
        txtFile.WriteLine "---------------------------------------"
        txtFile.WriteLine "********POWERED BY EASYMA***************"
    'End If
    txtFile.WriteLine escFeedAndCut
    
 txtFile.Close
 Reset
End If
txtdcode = ""
txtDispatch = ""
txtIntake = ""
txtDipping = ""
txtRefNo = ""


'* writing to notepad

'If chkNotepad.value = vbChecked Then

'    Dim fso, txtfile
'    Dim ttt
'     Dim escFeedAndCut As String
'     escFeedAndCut = Chr(29) + Chr(86) + Chr(65) '//Partial cut and feed
'       cdgPrint.Filter = "*.csv|*.txt"
'        cdgPrint.ShowSave
'        ttt = cdgPrint.FileName
'        If ttt = "" Then
'        MsgBox "File should not be blank", vbCritical, "Data transfer"
'        Exit Sub
'        End If
'        'ttt = "D:\PROJECTS\FOSA\DAILY" & Date & ""
'        Set fso = CreateObject("Scripting.FileSystemObject")
'        Set txtFile = fso.CreateTextFile(ttt, True)
'        txtFile.WriteLine
'
'    txtFile.WriteLine "---------------------------------------"
'    txtFile.WriteLine "" & cname & ""
'    txtFile.WriteLine " " & paddress & ""
'    txtFile.WriteLine " " & Phone & ""
'   ' Printer.Print Tab(0); "Kimathi House Branch"
'    txtFile.WriteLine " " & paddress & " "
'    txtFile.WriteLine "" & town & ""
'    txtFile.WriteLine "Milk Receipt"
'    txtFile.WriteLine "---------------------------------------"
''    If cbomemtrans = "Shares" Then
''    DESC = bosanames & " -Member No " & memberno
'    txtFile.WriteLine "SNo :" & txtSNo
'    txtFile.WriteLine "Name :" & lblNames
''    Else
'    txtFile.WriteLine "Quantity Supplied :" & txtQnty & " Kgs"
'    Startdate = DateSerial(year(DTPMilkDate), month(DTPMilkDate) - 1, 1)
'    'sql = "d_sp_TotalMonth " & txtSNo & ",'" & StartDate & "','" & DTPMilkDate & "'"
'    Set rs = New ADODB.Recordset
'    sql = "d_sp_TotalMonth " & txtSNo & ",'" & Startdate & "','" & DTPMilkDate & "'"
'    Set rs = oSaccoMaster.GetRecordset(sql)
'    If Not rs.EOF Then
'    If Not IsNull(rs.Fields(0)) Then CummulKgs = rs.Fields(0)
'    Else
'    CummulKgs = "0.00"
'    End If
'    txtFile.WriteLine "Cummulative This Month " & Format(CummulKgs, "#,##0.00" & " Kgs")
''    End If
'    Set rs = New ADODB.Recordset
'    sql = "d_sp_TransName '" & txtSNo & "'"
'    Set rs = oSaccoMaster.GetRecordset(sql)
'    If Not rs.EOF Then
'    If Not IsNull(rs.Fields(0)) Then TRANSPORTER = rs.Fields(0)
'    Else
'    TRANSPORTER = "Self"
'    End If
'    txtFile.WriteLine "---------------------------------------"
'    txtFile.WriteLine "Transporter :" & TRANSPORTER
'    txtFile.WriteLine "Received by :" & username
'    txtFile.WriteLine "---------------------------------------"
'    txtFile.WriteLine "Date :" & Format(Get_Server_Date, "dd/mm/yyyy HH:MM:SS AM/PM")
'    txtFile.WriteLine "     " & motto & " "
'    txtFile.WriteLine "---------------------------------------"
'    txtFile.WriteLine escFeedAndCut
'
'txtFile.Close








End If

MsgBox "Records saved successifully."
'Exit Sub





'//PRINT THE REPORT HERE
'milkinvoice

'd_MilkControl."RefNo"
'    STRFORMULA = "{d_MilkControl.RefNo}='" & txtRefNo & "'"
'    reportname = "milkinvoice.rpt"
'    Show_Sales_Crystal_Report STRFORMULA, reportname, title
    'Form_Load
    ListView2.Visible = False
    chkPay.value = vbUnchecked
    cmdnew.Enabled = True
    cmdsave.Enabled = True
    cmdedit.Enabled = False
    cmdnewsearch_Click
    Exit Sub
ErrorHandler:
        
        MsgBox err.description
End Sub

Private Sub cmdstatement_Click()
'milkstatement

   'reportname = "milkstatement.rpt"
    reportname = "d_DebtorsInvoice.rpt"
    Show_Sales_Crystal_Report "", reportname, ""
    'Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub Command1_Click()
frmDebtorsDetails.Show vbModal
End Sub

Private Sub DTPDispatchDate_Change()
chkPay.value = vbUnchecked
    Set rs = New ADODB.Recordset
    sql = ""
    'sql = "set dateformat dmy sp_dispatch '" & DTPDispatchDate & "'"
'    sql = "set dateformat dmy select intake,dipping from  d_dispatch where transdate='" & DTPDispatchDate & "'"
'    Set rs = oSaccoMaster.GetRecordset(sql)
'    If Not rs.EOF Then
'    'txtIntake = 0
'   ' txtDipping = 0
'    'rs.Fields(0) = 0
'    'rs.Fields(1) = 0
'     rs!intake = txtIntake
'
'    rs!dipping = txtDipping
'    txtIntake = txtDipping
'    'txtDipping = IIf(IsNull(rs.Fields(1)), 0, rs.Fields(1))
'    Else
'    txtIntake = "0.00"
'    txtDipping = 0#
'    End If
    Set rs = New ADODB.Recordset
    sql = "d_sp_DailyTotal '" & DTPDispatchDate & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not IsNull(rs.Fields(0)) Then
    txtIntake = Format(rs.Fields(0), "#0.00")
    txtDipping = txtIntake
    'txtDipping = Format(rs.Fields(0), "#0.00")
    'End If
    
  ' Update Milk Gl
'          Set rs = New ADODB.Recordset
'          sql = "SELECT Price from d_Price"
'          Set rs = oSaccoMaster.GetRecordset(sql)
'          If Not rs.EOF Then
'          Price = rs!Price
'          End If
'          Set cn = New ADODB.Connection
'          cn.Open Provider
'          sql = ""
'          sql = "set dateformat DMY select transdate,amount,draccno,craccno,documentno from gltransactions where draccno='" & lblDrAccM & "' and craccno='" & lblCrAccM & "' and transdate='" & DTPDispatchDate & "' "
'          Set rsy = New ADODB.Recordset
'          rsy.Open sql, cn
'         If rsy.EOF Then
'          sql = ""
'           sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & DTPDispatchDate & "'," & txtIntake & " *" & Price & ",'" & lblDrAccM & "','" & lblCrAccM & "','Milk Purchase','Milk Purchase' ,'Milk Purchase','" & User & "',0,0)"
'           'sql = "set dateformat dmy insert into  ag_products(p_code,p_name,s_no,qin,qout,date_entered,last_d_updated,user_id,audit_date,o_bal,supplierid,serialized,unserialized,seria,pprice,sprice,branch,AI ) values('" & txtpcode.Text & "','" & txtpname.Text & "','0'," & txtquantity.Text & "," & txtquantity.Text & ",'" & DTPdateentered & "','" & DTPdateentered & "','" & User & "','" & Date & "'," & txtquantity.Text & ",'',0,'',''," & txtpprice & "," & txtsellingprice & ",'" & Cbobrancht & "','" & lblStatus & "')"
'          cn.Execute sql
'
''          sql = ""
''          sql = "set dateformat dmy INSERT INTO d_DetailDispatch (DCode, Transdate, dispatch, auditid, auditdate)values ('" & txtdcode & "','" & DTPDispatchDate.value & "','" & txtDispatch & "','" & User & "','" & Get_Server_Date & "')"
''          oSaccoMaster.ExecuteThis (sql)
'
'        Else
'          sql = ""
'          sql = "set dateformat DMY update gltransactions set amount=" & txtIntake & " *" & Price & " where draccno='" & lblDrAccM & "' and craccno='" & lblCrAccM & "' and transdate='" & DTPDispatchDate & "'"
'          cn.Execute sql
'        End If
    ' End of milk update
    
    Else
    txtIntake = "0.00"
    End If
    Set rsq = New ADODB.Recordset
    sql = ""
    sql = "set dateformat dmy SELECT     SUM(DispQnty) AS qty From d_MilkControl WHERE     (DispDate = '" & DTPDispatchDate & "')"
     Set rsq = oSaccoMaster.GetRecordset(sql)
     If Not rsq.EOF Then
     If Not IsNull(rsq!qty) Then
    qty = Format(rsq!qty, "#0.00")
     txttotal = txtIntake - rsq!qty
     End If
     End If
    
    
    sql = "SET dateformat dmy select *  from d_milkcontrol WHERE     (DispDate = '" & DTPDispatchDate & "') AND vehicleno ='" & cboVehicle & "' "
    Set rss = oSaccoMaster.GetRecordset(sql)
    If rss.EOF Then
    ListView1.ListItems.Clear
    Exit Sub
    End If
sql = "set dateformat dmy SELECT     d.DCode, d.DName, SUM(m.DispQnty) AS quantity FROM         d_Debtors AS d INNER JOIN d_MilkControl AS m ON d.DCode = m.dcode WHERE     (DispDate = '" & DTPDispatchDate & "') and Locations='" & cboVehicle & "' GROUP BY d.DCode, d.DName"

Set rs = oSaccoMaster.GetRecordset(sql)
If rs.EOF Then
Exit Sub
End If

ListView1.ListItems.Clear

While Not rs.EOF
If Not IsNull(rs.Fields(0)) Then
Set li = ListView1.ListItems.Add(, , rs.Fields(0))
End If
                    If Not IsNull(rs.Fields(1)) Then li.SubItems(1) = rs.Fields(1) & ""
                        
                    If Not IsNull(rs.Fields(2)) Then li.SubItems(2) = rs.Fields(2) & ""
                   ' If Not IsNull(rs.Fields(3)) Then li.SubItems(3) = rs.Fields(3) & ""
'                    If Not IsNull(rs.Fields(4)) Then li.SubItems(4) = rs.Fields(4) & ""
'                    If Not IsNull(rs.Fields(5)) Then li.SubItems(5) = rs.Fields(5) & ""
rs.MoveNext
Wend
   chkPay.value = vbChecked
   ListView2.Visible = True
End Sub

Private Sub DTPDispatchDate_Click()
DTPDispatchDate_Change
End Sub

Private Sub DTPDispatchDate_Validate(Cancel As Boolean)
DTPDispatchDate_Change
End Sub

Private Sub Form_Load()
'Dim qty As Double
    DTPDispatchDate = Format(Get_Server_Date, "dd/mm/yyyy")
    DTPDispatchDate.MaxDate = Format(Get_Server_Date, "dd/mm/yyyy")
    cmdnewsearch_Click
'    txtCreditAcc.Locked = True
'    txtCreditAccName.Locked = True
'    txtDebitAcc.Locked = True
'    txtDebitAccName.Locked = True
    'txtDipping.Locked = True
    txtDispatch.Locked = True
    txtIntake.Locked = True
    txtVariance.Locked = True
    
'    txtCreditAcc = ""
'    txtCreditAccName = ""
'    txtDebitAcc = ""
'    txtDebitAccName = ""
    txtDipping = ""
    txtDispatch = ""
    txtIntake = ""
    txtVariance = ""
    

    cmdnew.Enabled = True
    cmdsave.Enabled = True
    cmdedit.Enabled = False
    ListView2.Visible = False
    
    
    Set rs = New ADODB.Recordset
    sql = "d_sp_DailyTotal '" & DTPDispatchDate & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not IsNull(rs.Fields(0)) Then
    txtIntake = Format(rs.Fields(0), "#0.00")
    txtDipping = txtIntake
    'txtDipping = Format(rs.Fields(0), "#0.00")
    'End If

'  sql = ""
'  sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & txtdateenterered & "'," & txtquantity & " *" & txtpprice & ",'" & 11 - 400 & "','" & 207 & "','stock intake','" & cbosupplier & "' ,'" & txtpname & "','" & User & "',0,0)"
'  oSaccoMaster.ExecuteThis (sql)

    Else
    txtIntake = "0.00"
    End If
' lblDrAccM = "11-400"
' lblCrAccM = "207"
    
    Set rsq = New ADODB.Recordset
    sql = ""
    sql = "set dateformat dmy SELECT     SUM(DispQnty) AS qty From d_MilkControl WHERE     (DispDate = '" & DTPDispatchDate & "')"
     Set rsq = oSaccoMaster.GetRecordset(sql)
     If Not rsq.EOF Then
     If Not IsNull(rsq!qty) Then
    qty = Format(rsq!qty, "#0.00")
     txttotal = txtIntake - rsq!qty
     End If
     End If

    Set rst = New Recordset
    Dim cn As Connection
    Set cn = New ADODB.Connection
    Provider = "MAZIWA"
    'Provider = cnn
    cn.Open Provider
    Set rst = New Recordset
    
    sql = "Select distinct(Locations) from   d_Debtors"
    rst.Open sql, cn, adOpenKeyset, adLockOptimistic
    While Not rst.EOF
    cboVehicle.AddItem rst.Fields(0)
    rst.MoveNext
    Wend
     
     
     
     
     
     
'sql = "SELECT d_Invoice.InvId, d_Invoice.RNo, d_Invoice.Vendor, d_Invoice.InvDate, d_Invoice.Amount, d_Invoice.[Desc] FROM  dbo.d_Requisition INNER JOIN "
'sql = sql & " d_Invoice ON d_Requisition.RNo = d_Invoice.RNo "
'sql = sql & " WHERE     (d_Requisition.Status <> 'Posted')order by InvDate DESC"

sql = "set dateformat dmy SELECT     d.DCode, d.DName, SUM(m.DispQnty) AS quantity FROM         d_Debtors AS d INNER JOIN d_MilkControl AS m ON d.DCode = m.dcode WHERE     (DispDate = '" & DTPDispatchDate & "') GROUP BY d.DCode, d.DName"

Set rs = oSaccoMaster.GetRecordset(sql)
If rs.EOF Then
Exit Sub
End If

ListView1.ListItems.Clear

While Not rs.EOF
If Not IsNull(rs.Fields(0)) Then
Set li = ListView1.ListItems.Add(, , rs.Fields(0))
End If
                    If Not IsNull(rs.Fields(1)) Then li.SubItems(1) = rs.Fields(1) & ""
                        
                    If Not IsNull(rs.Fields(2)) Then li.SubItems(2) = rs.Fields(2) & ""
                   ' If Not IsNull(rs.Fields(3)) Then li.SubItems(3) = rs.Fields(3) & ""
'                    If Not IsNull(rs.Fields(4)) Then li.SubItems(4) = rs.Fields(4) & ""
'                    If Not IsNull(rs.Fields(5)) Then li.SubItems(5) = rs.Fields(5) & ""
rs.MoveNext

Wend


End Sub

Private Sub lvwCreditAcc_DblClick()
'Dim rsAccount As New ADODB.Recordset
'txtCreditAcc = lvwCreditAcc.SelectedItem
'Set rsAccount = oSaccoMaster.GetRecordset("Select * From GLSETUP where " _
'            & "accno= '" & txtCreditAcc & "'")
'If Not rsAccount.EOF Then
'   txtCreditAccName = IIf(IsNull(rsAccount!GlAccName), "", rsAccount!GlAccName)
'
 
'End If


'lvwCreditAcc.Visible = False

End Sub


Private Sub lvwDebitAcc_DblClick()
Dim rsAccount As New ADODB.Recordset
'txtDebitAcc = lvwDebitAcc.SelectedItem
'Set rsAccount = oSaccoMaster.GetRecordset("Select * From GLSETUP where " _
'            & "accno= '" & txtDebitAcc & "'")
If Not rsAccount.EOF Then
'   txtDebitAccName = IIf(IsNull(rsAccount!GlAccName), "", rsAccount!GlAccName)
  
 
End If


'lvwDebitAcc.Visible = False

End Sub



'Private Sub lblCrAccM_Change()
'    Set rsy = New ADODB.Recordset
'    sql = "set dateformat DMY select glaccname from glsetup where accno='" & lblCrAccM & "'"
'    Set rsy = oSaccoMaster.GetRecordset(sql)
'    'Set rsy = oSaccoMaster.GetRecordset("set dateformat DMY select glaccname from glsetup where accno='" & lblCrAccM & "'")
'    If Not rsy.EOF Then
'    txtcraccm = rsy.Fields("glaccname")
'    End If
'End Sub
'
'Private Sub lblDrAccM_change()
'    Set rsy = New ADODB.Recordset
'    sql = "set dateformat DMY select glaccname from glsetup where accno='" & lblDrAccM & "'"
'    Set rsy = oSaccoMaster.GetRecordset(sql)
'    If Not rsy.EOF Then
'    txtdraccm = rsy.Fields("glaccname")
'    End If
'End Sub

Private Sub listview1_DblClick()
frmmilkdidprev.txtdcode = ListView1.SelectedItem
frmmilkdidprev.txtdesc = ListView1.SelectedItem.ListSubItems(1)
frmmilkdidprev.txtquantity = ListView1.SelectedItem.ListSubItems(2)


Dim q As Double
''//get the quantity for the same first
'
'sql = ""
'sql = "SELECT     qnty,pricing FROM  d_Requisition  where rno='" & frmrequisitionapproval.lblrno & "'"
'Set rs = oSaccoMaster.GetRecordset(sql)
'While Not rs.EOF
'DoEvents
'
'q = rs.Fields(0)
'frmrequisitionapproval.txtestimate = (q * rs.Fields(1))
'rs.MoveNext
'Wend
'lvwrequisition.ListItems.Remove (lvwrequisition.SelectedItem.index)
frmmilkdidprev.Show vbModal
End Sub

Private Sub listview12_DblClick()
frmmilkdidprev.txtdcode = ListView1.SelectedItem
frmmilkdidprev.txtdesc = ListView1.SelectedItem.ListSubItems(1)
frmmilkdidprev.txtquantity = ListView1.SelectedItem.ListSubItems(2)


Dim q As Double
''//get the quantity for the same first
'
'sql = ""
'sql = "SELECT     qnty,pricing FROM  d_Requisition  where rno='" & frmrequisitionapproval.lblrno & "'"
'Set rs = oSaccoMaster.GetRecordset(sql)
'While Not rs.EOF
'DoEvents
'
'q = rs.Fields(0)
'frmrequisitionapproval.txtestimate = (q * rs.Fields(1))
'rs.MoveNext
'Wend
'lvwrequisition.ListItems.Remove (lvwrequisition.SelectedItem.index)
frmmilkdidprev.Show vbModal
End Sub


Private Sub ListView2_DblClick()
frmNominals.chkDebtno = vbChecked
'frmNominals.txtTCode = ListView1.SelectedItem
'frmNominals.txtDNames = ListView1.SelectedItem.ListSubItems(1)
'frmNominals.dtpTransDate = frmMilkControl.DTPDispatchDate
'frmNominals.txtAccNames = ListView1.SelectedItem.ListSubItems(3)
'Dim q As Double
frmNominals.Show vbModal
'Form_Load
'ListView2_load
chkPay.value = vbUnchecked
'chkPay.value = vbChecked
End Sub

Private Sub mnuinvoice_Click()
frmmilkinvoice.Show vbModal
End Sub

Private Sub Picture1_Click()
Me.MousePointer = vbHourglass
         frmSearchMilkControl.Show vbModal
        txtRefNo = sel
        txtRefNo_Validate True
        Me.MousePointer = 0
End Sub

'Private Sub Picture2_Click()
'
'Me.MousePointer = vbHourglass
'        frmSearchGLAccounts.Show vbModal, Me
'    If Continue Then
'        If SearchValue <> "" Then
'            lblCrAccM = SearchValue
'            SearchValue = ""
'        End If
'    End If
'    Me.MousePointer = 0
'
'End Sub

Private Sub Picture3_Click()
Me.MousePointer = vbHourglass
        'sql = "SELECT   dcode,dname   FROM    d_Debtors  where Locations= '" & cboVehicle & "'"
        frmSearchDebtors.Show vbModal
        txtdcode = sel
        txtdcode_Validate True
        Me.MousePointer = 0
End Sub

Private Sub txtCreditAccName_Change()
'On Error GoTo SysError
    Dim rsAccount As New Recordset
'    lvwCreditAcc.ListItems.Clear
    
'    If Trim$(txtCreditAccName) <> "" Then
'        'If Editing = True Then
'            Set rsAccount = oSaccoMaster.GetRecordset("Select * From GLSETUP where " _
'            & "GLAccName Like '%" & txtCreditAccName & "%'")
'            With rsAccount
'                If .State = adStateOpen Then
'                    If Not .EOF Then
'                        'lvwContraAcc.Visible = True
'                        If .RecordCount = 1 Then
'                            txtCreditAcc = IIf(IsNull(!accno), "", !accno)
'                            Editing = True
'                            txtCreditAccName = IIf(IsNull(!GlAccName), "", !GlAccName)
'                            lvwCreditAcc.Visible = False
'                            Else
'                            lvwCreditAcc.Visible = False
'
'                        End If
'                    Else
'                        lvwCreditAcc.Visible = False
'                    End If
'                    'lvwDeductionAcc.Visible = True
'                    While Not .EOF
'                        lvwCreditAcc.Visible = True
'                        Set li = lvwCreditAcc.ListItems.Add(, , IIf(IsNull(!accno), "", !accno))
'                        li.SubItems(1) = IIf(IsNull(!GlAccName), "", !GlAccName)
'                        .MoveNext
'                    Wend
'                    'lvwDeductionAcc.Visible = False
'                End If
'            End With
'        'End If
'    End If
'    Exit Sub
'SysError:
'    MsgBox err.description, vbInformation, Me.Caption
'
'End Sub
'
'
'
'Private Sub txtdcode_Validate(Cancel As Boolean)
'sql = "select dname,Price from d_debtors where dcode='" & txtdcode & "'"
'Set rs = oSaccoMaster.GetRecordset(sql)
'If Not rs.EOF Then
'If Not IsNull(rs.Fields(0)) Then lblDebtors = rs.Fields(0)
'If Not IsNull(rs.Fields(1)) Then Price = rs.Fields(1)
'End If
'End Sub
'
'Private Sub txtDebitAccName_Change()
'On Error GoTo SysError
'    Dim rsAccount As New Recordset
'    lvwDebitAcc.ListItems.Clear
'
'    If Trim$(txtDebitAccName) <> "" Then
'        'If Editing = True Then
'            Set rsAccount = oSaccoMaster.GetRecordset("Select * From GLSETUP where " _
'            & "GLAccName Like '%" & txtDebitAccName & "%'")
'            With rsAccount
'                If .State = adStateOpen Then
'                    If Not .EOF Then
'                        'lvwContraAcc.Visible = True
'                        If .RecordCount = 1 Then
'                            txtDebitAcc = IIf(IsNull(!accno), "", !accno)
'                            Editing = True
'                            txtDebitAccName = IIf(IsNull(!GlAccName), "", !GlAccName)
'                            lvwDebitAcc.Visible = False
'                            Else
'                            lvwDebitAcc.Visible = False
'
'                        End If
'                    Else
'                        lvwDebitAcc.Visible = False
'                    End If
'                    'lvwDeductionAcc.Visible = True
'                    While Not .EOF
'                        lvwDebitAcc.Visible = True
'                        Set li = lvwDebitAcc.ListItems.Add(, , IIf(IsNull(!accno), "", !accno))
'                        li.SubItems(1) = IIf(IsNull(!GlAccName), "", !GlAccName)
'                        .MoveNext
'                    Wend
'                    'lvwDeductionAcc.Visible = False
'                End If
'            End With
'        'End If
'    End If
'    Exit Sub
'SysError:
'    MsgBox err.description, vbInformation, Me.Caption

End Sub

'Private Sub Picture4_Click()
'
'Me.MousePointer = vbHourglass
'        frmSearchGLAccounts.Show vbModal, Me
'    If Continue Then
'        If SearchValue <> "" Then
'            lblDrAccM = SearchValue
'            SearchValue = ""
'        End If
'    End If
'    Me.MousePointer = 0
'
'End Sub

Private Sub txtdcode_Validate(Cancel As Boolean)
Set rs = oSaccoMaster.GetRecordset("SELECT dname,Price,accdr,acccr,drcess,crcess,capp,crate FROM d_Debtors WHERE DCode = '" & txtdcode & "'")
If rs.RecordCount > 0 Then
If Not IsNull(rs.Fields(1)) Then Price = rs.Fields(1)
If Not IsNull(rs.Fields(0)) Then lblDebtors = rs.Fields(0)
If Not IsNull(rs.Fields(2)) Then Label10 = rs.Fields(2)
If Not IsNull(rs.Fields(3)) Then Label11 = rs.Fields(3)
If Not IsNull(rs.Fields(4)) Then cessdr = rs.Fields(4)
If Not IsNull(rs.Fields(5)) Then cesscr = rs.Fields(5)
If Not IsNull(rs.Fields(6)) Then capp = Abs(rs.Fields(6))
If Not IsNull(rs.Fields(7)) Then crate = rs.Fields(7)
txtamountp = txtDispatch * rs.Fields(1)
If capp = 1 Then
chkapp = vbChecked
Else
chkapp = vbUnchecked
End If
Else
lblDebtors = ""
End If
End Sub

Private Sub txtDipping_Change()
If txtIntake = "" Then
txtIntake = "0"
End If
If txtDipping = "" Then
txtDipping = "0"
End If
'txtVariance = Format(txtDipping - txtDispatch, "#0.00")
End Sub

Private Sub txtDipping_Validate(Cancel As Boolean)
txtDispatch_Change
End Sub

Private Sub txtDispatch_Change()
txtDipping = txtDispatch
If txtDispatch = "" Then
txtDispatch = "0"
End If
If txtDipping = "" Then
txtDipping = "0"
End If
'**************PRICE***************'
Set rs = oSaccoMaster.GetRecordset("SELECT dname,Price,accdr,acccr,drcess,crcess,capp,crate FROM d_Debtors WHERE DCode = '" & txtdcode & "'")
If rs.RecordCount > 0 Then
If Not IsNull(rs.Fields(1)) Then Price = rs.Fields(1)
If Not IsNull(rs.Fields(0)) Then lblDebtors = rs.Fields(0)
If Not IsNull(rs.Fields(2)) Then Label10 = rs.Fields(2)
If Not IsNull(rs.Fields(3)) Then Label11 = rs.Fields(3)
If Not IsNull(rs.Fields(4)) Then cessdr = rs.Fields(4)
If Not IsNull(rs.Fields(5)) Then cesscr = rs.Fields(5)
If Not IsNull(rs.Fields(6)) Then capp = Abs(rs.Fields(6))
If Not IsNull(rs.Fields(7)) Then crate = rs.Fields(7)
txtamountp = txtDispatch * rs.Fields(1)
If capp = 1 Then
chkapp = vbChecked
Else
chkapp = vbUnchecked
End If
Else
lblDebtors = ""
End If


'****************END********************'





txtVariance = Format(txtDipping - txtDispatch, "#0.00")

End Sub



Private Sub txtDispatch_Validate(Cancel As Boolean)
txtDipping_Change
End Sub

Private Sub txtdraccm_Change()

End Sub

Private Sub txtIntake_Change()
txtDispatch_Change
End Sub

Private Sub txtIntake_Validate(Cancel As Boolean)
txtDispatch_Change
End Sub

Private Sub txtRefNo_Validate(Cancel As Boolean)
On Error GoTo ErrorHandler
'SELECT TOP 300 DispDate,RefNo,dcode,DispQnty,Price,InQnty,Variance FROM dbo.d_MilkControl"
If Trim(txtRefNo) = "" Then
Exit Sub
End If
 Set rs = oSaccoMaster.GetRecordset("SELECT DispDate,dcode,DispQnty,Price,InQnty,sum(Variance) FROM d_MilkControl WHERE RefNo = '" & txtRefNo & "'")
 
 If rs.RecordCount > 0 Then
    DTPDispatchDate = rs.Fields(0)
    txtDispatch = rs.Fields(2)
    txtDipping = txtDispatch
    txtIntake = rs.Fields(4)
    txtVariance = rs.Fields(5)
    txtdcode = rs.Fields(1)
    
    cmdedit.Enabled = True
Else
    cmdedit.Enabled = False
    
End If
txtdcode_Validate True
Exit Sub
ErrorHandler:
MsgBox err.description
End Sub


Private Sub txtVariance_Change()
'Set rs = oSaccoMaster.GetRecordset("SELECT SUM(Variance) AS Expr1 FROM d_MilkControl WHERE (DispDate ='" & DTPDispatchDate & "')")
'  If Not rs.EOF Then
'   txtVariance = rs!Variance
'   Else
'    txtVariance = "0"
'   End If
'txtVariance = rs
End Sub

Private Sub txtvehicleno_Change()

End Sub
