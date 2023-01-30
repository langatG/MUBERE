VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmpostbookedtransactions 
   Caption         =   "POST BOOKED TRANSACTIONS"
   ClientHeight    =   6765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11985
   Icon            =   "frmpostbookedtransactions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   11985
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Delete Wrongly Posted"
      Height          =   375
      Left            =   240
      TabIndex        =   38
      Top             =   6120
      Width           =   2055
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   390
      Left            =   8025
      TabIndex        =   22
      Top             =   2040
      Width           =   1275
   End
   Begin VB.ComboBox cboTransType 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmpostbookedtransactions.frx":0BC2
      Left            =   6270
      List            =   "frmpostbookedtransactions.frx":0BCC
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   2115
      Width           =   1680
   End
   Begin VB.TextBox txtNarration 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4935
      TabIndex        =   20
      Top             =   945
      Width           =   4365
   End
   Begin VB.TextBox txtCrAccNo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   345
      TabIndex        =   19
      Text            =   "L099"
      Top             =   2130
      Width           =   1440
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   285
      Left            =   30
      TabIndex        =   18
      Top             =   2130
      Width           =   315
   End
   Begin VB.TextBox txtDrAccNo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   345
      TabIndex        =   17
      Top             =   1560
      Width           =   1440
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "..."
      Height          =   285
      Left            =   45
      TabIndex        =   16
      Top             =   1575
      Width           =   300
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5565
      TabIndex        =   15
      Top             =   1560
      Width           =   1665
   End
   Begin VB.TextBox txtSource 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1755
      TabIndex        =   14
      Top             =   945
      Width           =   3135
   End
   Begin VB.CommandButton cmdPost 
      Caption         =   "&Post"
      Height          =   405
      Left            =   6075
      TabIndex        =   13
      Top             =   6150
      Width           =   1275
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   405
      Left            =   10095
      TabIndex        =   12
      Top             =   6150
      Width           =   1275
   End
   Begin VB.TextBox txtDocumentNo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7380
      TabIndex        =   11
      Top             =   1560
      Width           =   1905
   End
   Begin VB.TextBox lblDrAccName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1980
      TabIndex        =   10
      Top             =   1560
      Width           =   3225
   End
   Begin VB.TextBox txtCrAccName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1980
      TabIndex        =   9
      Top             =   2130
      Width           =   3225
   End
   Begin VB.ComboBox cboCompanyCode 
      Height          =   315
      Left            =   5190
      TabIndex        =   7
      Top             =   225
      Width           =   1965
   End
   Begin VB.CommandButton cmdprepost 
      Caption         =   "Print Journals"
      Height          =   390
      Left            =   7440
      TabIndex        =   6
      Top             =   6180
      Width           =   2580
   End
   Begin VB.Frame Frame1 
      Height          =   600
      Left            =   105
      TabIndex        =   3
      Top             =   0
      Width           =   3870
      Begin VB.OptionButton optmember 
         Caption         =   "Supplier"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   255
         Width           =   1590
      End
      Begin VB.OptionButton optgroup 
         Caption         =   "Other Payment"
         Height          =   270
         Left            =   1965
         TabIndex        =   4
         Top             =   210
         Width           =   1695
      End
   End
   Begin VB.TextBox txtchequeno 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9480
      TabIndex        =   2
      Top             =   1575
      Width           =   1905
   End
   Begin VB.CheckBox chknonmemberpostings 
      Caption         =   "Non Member Posting"
      Height          =   285
      Left            =   9645
      TabIndex        =   1
      Top             =   2040
      Value           =   1  'Checked
      Width           =   2190
   End
   Begin VB.CommandButton cmdloadbookings 
      Caption         =   "Refresh Unposted"
      Height          =   375
      Left            =   3960
      TabIndex        =   0
      Top             =   6150
      Width           =   1695
   End
   Begin MSComctlLib.ListView lvwAccName 
      Height          =   1350
      Left            =   1980
      TabIndex        =   8
      Top             =   1860
      Visible         =   0   'False
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   2381
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "AccName"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "AccNo"
         Object.Width           =   18
      EndProperty
   End
   Begin MSComCtl2.DTPicker dtpTransDate 
      Height          =   315
      Left            =   45
      TabIndex        =   23
      Top             =   945
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   " dd-MM-yyyy"
      Format          =   90701827
      CurrentDate     =   39400
   End
   Begin MSComctlLib.ListView lvwTrans 
      Height          =   3390
      Left            =   0
      TabIndex        =   24
      Top             =   2505
      Width           =   11685
      _ExtentX        =   20611
      _ExtentY        =   5980
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "TransDate"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Amount"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Dr AccNo"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Cr AccNo"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "DocumentNo"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Source"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Narration"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "DocPosted"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Cheque No."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "ID"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "TransType"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5340
      TabIndex        =   37
      Top             =   2160
      Width           =   870
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Naration"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4995
      TabIndex        =   36
      Top             =   690
      Width           =   705
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Cr AccNo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   390
      TabIndex        =   35
      Top             =   1905
      Width           =   765
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "AccName"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1995
      TabIndex        =   34
      Top             =   1890
      Width           =   795
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Dr AccNo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   390
      TabIndex        =   33
      Top             =   1350
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "AccName"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1995
      TabIndex        =   32
      Top             =   1335
      Width           =   795
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Amount"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6585
      TabIndex        =   31
      Top             =   1350
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Transaction Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   105
      TabIndex        =   30
      Top             =   690
      Width           =   1410
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Source"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1815
      TabIndex        =   29
      Top             =   690
      Width           =   585
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Document No"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   7395
      TabIndex        =   28
      Top             =   1350
      Width           =   1140
   End
   Begin VB.Label Label11 
      Caption         =   "Source"
      Height          =   255
      Left            =   4005
      TabIndex        =   27
      Top             =   255
      Width           =   1080
   End
   Begin VB.Label lblCompanyName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   7260
      TabIndex        =   26
      Top             =   225
      Width           =   3030
   End
   Begin VB.Label Label12 
      Caption         =   "Cheque No."
      Height          =   195
      Left            =   9465
      TabIndex        =   25
      Top             =   1350
      Width           =   1665
   End
End
Attribute VB_Name = "frmpostbookedtransactions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rscompany As New ADODB.Recordset
Private Sub cboCompanyCode_Click()
If optgroup = True Then
Set rscompany = Nothing
Set rscompany = oSaccoMaster.GetRecordset("select description from transcode where description='" & Trim(cboCompanyCode) & "'")
If Not rscompany.EOF Then
    lblCompanyName.Caption = rscompany.Fields(0)
    txtSource = cboCompanyCode
    txtNarration = lblCompanyName
End If
Else
If optmember = True Then
Set rscompany = oSaccoMaster.GetRecordset("select surname, othernames from members where memberno='" & Trim(cboCompanyCode) & "'")
If Not rscompany.EOF Then
    lblCompanyName.Caption = rscompany.Fields(0) & "  " & rscompany.Fields(1)
    txtSource = cboCompanyCode
    txtNarration = lblCompanyName
End If
End If
End If

End Sub

Private Sub cmdAdd_Click()
    On Error GoTo SysError
    If Trim$(txtAmount) = "" Then
        MsgBox "Please enter the Amount", vbInformation, Me.Caption
        txtAmount.SetFocus
        SendKeys "{Home}+{End}"
        Exit Sub
    End If
    
    If Trim(txtChequeno) = "" Then
        MsgBox "Please Enter The chequne No", vbInformation, Me.Caption
        Exit Sub
    End If
    
    If Val(txtAmount) = 0 Then
        MsgBox "Please enter a valid Amount", vbInformation, Me.Caption
        txtAmount.SetFocus
        SendKeys "{Home}+{End}"
        Exit Sub
    End If
    If Trim$(txtDrAccNo) = "" Then
        MsgBox "Please enter the Account to Debit.", vbInformation, Me.Caption
        txtDrAccNo.SetFocus
        SendKeys "{Home}+{End}"
        Exit Sub
    End If
    If Trim$(lblDrAccName) = "" Then
        MsgBox "Please verify the Debit Account.", vbInformation, Me.Caption
        lblDrAccName.SetFocus
        SendKeys "{Home}+{End}"
        Exit Sub
    End If
    If Trim$(txtCrAccNo) = "" Then
        MsgBox "Please enter the Account to Credit.", vbInformation, Me.Caption
        txtCrAccNo.SetFocus
        SendKeys "{Home}+{End}"
        Exit Sub
    End If
    If Trim$(txtCrAccName) = "" Then
        MsgBox "Please verify the Credit Account", vbInformation, Me.Caption
        txtCrAccName.SetFocus
        SendKeys "{Home}+{End}"
        Exit Sub
    End If
    If Trim$(txtSource) = "" Then
        MsgBox "Please enter the Transaction Source", vbInformation, Me.Caption
        txtSource.SetFocus
        SendKeys "{Home}+{End}"
        Exit Sub
    End If
    If Trim$(txtNarration) = "" Then
        MsgBox "Please enter the Transaction Description", vbInformation, Me.Caption
        txtNarration.SetFocus
        SendKeys "{Home}+{End}"
        Exit Sub
    End If
    If Trim$(txtdocumentno) = "" Then
        MsgBox "Please enter the Amount", vbInformation, Me.Caption
        txtdocumentno.SetFocus
        SendKeys "{Home}+{End}"
        Exit Sub
    End If
    Set li = lvwTrans.ListItems.Add(, , dtpTransDate)
    li.SubItems(1) = Format(CDbl(txtAmount), "#,##0.00")
    li.SubItems(2) = txtDrAccNo
    li.SubItems(3) = txtCrAccNo
    li.SubItems(4) = txtdocumentno
    li.SubItems(5) = txtSource
    li.SubItems(6) = txtNarration
    If chknonmemberpostings = vbChecked Then
    li.SubItems(7) = 1
    Else
    li.SubItems(7) = 0
    End If
    li.SubItems(8) = txtChequeno
    txtAmount = "0"
    txtDrAccNo = ""
    txtCrAccNo = ""
    txtdocumentno = ""
    txtSource = ""
    txtNarration = ""
    lblDrAccName = ""
    txtCrAccName = ""
    txtChequeno = ""
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdloadbookings_Click()
load_Trans
End Sub
Private Sub load_Trans()
Set rs = oSaccoMaster.GetRecordset("set dateformat dmy select * from bookings where glpost=0 ")
While Not rs.EOF
Set li = lvwTrans.ListItems.Add(, , rs.Fields("transdate"))
    li.SubItems(1) = Format(CDbl(rs.Fields("Amount")), "#,##0.00")
    li.SubItems(2) = rs.Fields("DrAccNo")
    li.SubItems(3) = rs.Fields("CrAccNo")
    li.SubItems(4) = rs.Fields("DocumentNo")
    li.SubItems(5) = rs.Fields("Source")
    li.SubItems(6) = rs.Fields("TransDescript")
    If rs.Fields("doc_posted") = 1 Then
    li.SubItems(7) = 1
    Else
    li.SubItems(7) = 0
    End If
    li.SubItems(8) = rs.Fields("ChequeNo")
    li.SubItems(9) = rs.Fields("ID")
    rs.MoveNext
    Wend
  
End Sub
Private Sub cmdPost_Click()
    On Error GoTo SysError
    If Check_Period_If_Closed(dtpTransDate) = True Then
         Exit Sub
     End If
    Dim Cubaccount As Cub_Acc_Details
    Dim Account As Acc_Details
    Dim chequeno As String
    Dim DRaccno As String, Craccno As String, amount As Double, transdate As Date, _
    transDescription As String, TransSource As String, DocumentNo As String, CashBook As Long, doc_posted As Integer, IDENTI As Long
    If lvwTrans.ListItems.Count > 0 Then
        If MsgBox("Do you want post the entry?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
            Exit Sub
        End If
    Else
        MsgBox "There are no transactions to be posted", vbInformation, Me.Caption
        Exit Sub
    End If
    Me.MousePointer = vbHourglass
    For I = 1 To lvwTrans.ListItems.Count
     If lvwTrans.ListItems.Item(I).Checked = True Then
        Set li = lvwTrans.ListItems(I)
        transdate = li
        amount = CDbl(lvwTrans.ListItems(I).SubItems(1))
        DRaccno = lvwTrans.ListItems(I).SubItems(2)
        Craccno = lvwTrans.ListItems(I).SubItems(3)
        DocumentNo = lvwTrans.ListItems(I).SubItems(4)
        TransSource = lvwTrans.ListItems(I).SubItems(5)
        transDescription = lvwTrans.ListItems(I).SubItems(6)
        chequeno = lvwTrans.ListItems(I).SubItems(8)
        doc_posted = lvwTrans.ListItems(I).SubItems(7)
IDENTI = lvwTrans.ListItems(I).SubItems(9)
        CashBook = 1
        If DocumentNo = "" Then DocumentNo = cboCompanyCode

                If chknonmemberpostings = vbChecked Then
                doc_posted = 1
                Else
                doc_posted = 0
                End If
        
        Set rs = oSaccoMaster.GetRecordset("sp_chequeno_used '" & Trim(chequeno) & "','" & Trim(TransSource) & "'")
        If Not rs.EOF Then

         End If

       
        
        If Not Save_GLTRANSACTION(transdate, amount, DRaccno, Craccno, DocumentNo, _
        TransSource, User, ErrorMessage, transDescription, CashBook, doc_posted, chequeno, transactionNo, "", "", 0) Then
            If ErrorMessage <> "" Then
                MsgBox ErrorMessage, vbInformation, Me.Caption
                ErrorMessage = ""
            End If
        End If
        
        
'//UPDATE BOOKING THAT HAVE BEEN POSTED
sql = ""
sql = "Update bookings set glpost=1 where id =" & IDENTI & ""
oSaccoMaster.ExecuteThis (sql)
        
     
        
        
        End If
    Next I
    '//clear listview
    
    lvwTrans.ListItems.Clear
    
    Me.MousePointer = vbDefault
    MsgBox "Posting Successfull", vbInformation, Me.Caption
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
    Me.MousePointer = vbDefault
End Sub

Private Sub cmdprepost_Click()
frmreversalofcashbookentries.Show vbModal, Me
End Sub

Private Sub cmdsearch_Click()
frmSearchGLAccounts.Show vbModal, Me
    If Continue Then
        If SearchValue <> "" Then
            txtDrAccNo = SearchValue
            SearchValue = ""
        End If
    End If
End Sub

Private Sub Command1_Click()
frmSearchGLAccounts.Show vbModal, Me
    If Continue Then
        If SearchValue <> "" Then
            txtCrAccNo = SearchValue
            SearchValue = ""
        End If
    End If
End Sub

Private Sub Command2_Click()
    On Error GoTo SysError
    If Check_Period_If_Closed(dtpTransDate) = True Then
         Exit Sub
     End If
    Dim Cubaccount As Cub_Acc_Details
    Dim Account As Acc_Details
    Dim chequeno As String
    Dim DRaccno As String, Craccno As String, amount As Double, transdate As Date, _
    transDescription As String, TransSource As String, DocumentNo As String, CashBook As Long, doc_posted As Integer, IDENTI As Long
    If lvwTrans.ListItems.Count > 0 Then
        If MsgBox("Do you want to Delete this entry?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
            Exit Sub
        End If
    Else
        MsgBox "There are no transactions to be posted", vbInformation, Me.Caption
        Exit Sub
    End If
    Me.MousePointer = vbHourglass
    For I = 1 To lvwTrans.ListItems.Count
     If lvwTrans.ListItems.Item(I).Checked = True Then
        Set li = lvwTrans.ListItems(I)
        transdate = li
        amount = CDbl(lvwTrans.ListItems(I).SubItems(1))
        DRaccno = lvwTrans.ListItems(I).SubItems(2)
        Craccno = lvwTrans.ListItems(I).SubItems(3)
        DocumentNo = lvwTrans.ListItems(I).SubItems(4)
        TransSource = lvwTrans.ListItems(I).SubItems(5)
        transDescription = lvwTrans.ListItems(I).SubItems(6)
        chequeno = lvwTrans.ListItems(I).SubItems(8)
        doc_posted = lvwTrans.ListItems(I).SubItems(7)
IDENTI = lvwTrans.ListItems(I).SubItems(9)
        CashBook = 1
        If DocumentNo = "" Then DocumentNo = cboCompanyCode

                If chknonmemberpostings = vbChecked Then
                doc_posted = 1
                Else
                doc_posted = 0
                End If
        
        Set rs = oSaccoMaster.GetRecordset("sp_chequeno_used '" & Trim(chequeno) & "','" & Trim(TransSource) & "'")
        If Not rs.EOF Then

         End If

       
        
       
        
        
'//UPDATE BOOKING THAT HAVE BEEN POSTED
sql = ""
sql = "delete from bookings  where id =" & IDENTI & ""
oSaccoMaster.ExecuteThis (sql)
        
     
        
        
        End If
    Next I
    '//clear listview
    
    lvwTrans.ListItems.Clear
    
    Me.MousePointer = vbDefault
    MsgBox "Record Deleted Successfull", vbInformation, Me.Caption
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
    Me.MousePointer = vbDefault

End Sub

Private Sub Form_Load()
    dtpTransDate = Format(Get_Server_Date, " dd-MM-yyyy")
 load_Trans
End Sub

Private Sub lblDrAccName_Change()
    On Error GoTo SysError
    Dim rsAccounts As New Recordset
    TSource = "DR"
    lvwAccName.ListItems.Clear
    If Trim$(lblDrAccName) <> "" Then
        If Not Editing Then
            Set rsAccounts = oSaccoMaster.GetRecordset("Exec Get_Acc_Names '%" & lblDrAccName & "%'")
            With rsAccounts
                If .State = adStateOpen Then
                    If Not .EOF Then
                        lvwAccName.Visible = True
                        lvwAccName.Top = 1365
                    Else
                        lvwAccName.Visible = False
                    End If
                    While Not .EOF
                        Set li = lvwAccName.ListItems.Add(, , !GlAccName)
                        li.SubItems(1) = IIf(IsNull(!AccNo), "", !AccNo)
                        .MoveNext
                    Wend
                End If
            End With
        End If
    End If
    If lvwAccName.ListItems.Count = 1 Then
        lvwAccName_DblClick
    End If
    Editing = False
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub lblDrAccName_KeyPress(KeyAscii As Integer)
    KeyAscii = To_Upper_Case(KeyAscii)
End Sub

Private Sub lvwAccName_DblClick()
    If lvwAccName.ListItems.Count > 0 Then
        Select Case TSource
            Case "DR"
            lblDrAccName = lvwAccName.SelectedItem
            txtDrAccNo = lvwAccName.SelectedItem.SubItems(1)
            Case "CR"
            txtCrAccName = lvwAccName.SelectedItem
            txtCrAccNo = lvwAccName.SelectedItem.SubItems(1)
        End Select
        lvwAccName.Visible = False
    End If
End Sub


Private Sub lvwTrans_DblClick()
    If lvwTrans.ListItems.Count > 0 Then
        If MsgBox("Are you sure you delete  this records " & lvwTrans.SelectedItem.Text & "? ", vbYesNo) = vbYes Then
        lvwTrans.ListItems.Remove (lvwTrans.SelectedItem.Index)  '// removes the selected item
        End If
    End If
End Sub

Private Sub optgroup_Click()
 Set rscompany = Nothing

Set rscompany = oSaccoMaster.GetRecordset("select * from transcode order by description asc")
With rscompany
    If Not .EOF Then
    cboCompanyCode.Clear
        While Not .EOF
            cboCompanyCode.AddItem !description
            .MoveNext
        Wend
    End If
End With
End Sub

Private Sub optmember_Click()
 Set rscompany = Nothing

Set rscompany = oSaccoMaster.GetRecordset("select memberno from members order by memberno asc")
With rscompany
    If Not .EOF Then
    cboCompanyCode.Clear
        While Not .EOF
            cboCompanyCode.AddItem !memberno
            .MoveNext
        Wend
    End If
End With
End Sub

Private Sub txtCrAccName_Change()
    On Error GoTo SysError
    Dim rsAccounts As New Recordset
    TSource = "CR"
    lvwAccName.ListItems.Clear
    If Trim$(txtCrAccName) <> "" Then
        If Not Editing Then
            Set rsAccounts = oSaccoMaster.GetRecordset("Exec Get_Acc_Names '%" & txtCrAccName & "%'")
            With rsAccounts
                If .State = adStateOpen Then
                    If Not .EOF Then
                        lvwAccName.Visible = True
                        lvwAccName.Top = 1935
                    Else
                        lvwAccName.Visible = False
                    End If
                    While Not .EOF
                        Set li = lvwAccName.ListItems.Add(, , !GlAccName)
                        li.SubItems(1) = IIf(IsNull(!AccNo), "", !AccNo)
                        .MoveNext
                    Wend
                End If
            End With
        End If
    End If
    If lvwAccName.ListItems.Count = 1 Then
        lvwAccName_DblClick
    End If
    Editing = False
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub txtCrAccName_KeyPress(KeyAscii As Integer)
    KeyAscii = To_Upper_Case(KeyAscii)
End Sub

Private Sub txtCrAccNo_Change()
    On Error GoTo SysError
    Dim Account As Acc_Details
        
        Editing = True
    Account = Get_Acc_Details(txtCrAccNo, ErrorMessage)
    If Account.AccNo <> "" Then
        txtCrAccName = Account.AccName
    Else
        If ErrorMessage <> "" Then
            MsgBox ErrorMessage, vbInformation, Me.Caption
            ErrorMessage = ""
        End If
        txtCrAccName = ""
    End If
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub txtCrAccNo_KeyPress(KeyAscii As Integer)
    KeyAscii = To_Upper_Case(KeyAscii)
End Sub

Private Sub txtDrAccNo_Change()
    On Error GoTo SysError
    Dim Account As Acc_Details
    Editing = True
    Account = Get_Acc_Details(txtDrAccNo, ErrorMessage)
    If Account.AccNo <> "" Then
        lblDrAccName = Account.AccName
    Else
        If ErrorMessage <> "" Then
            MsgBox ErrorMessage, vbInformation, Me.Caption
            ErrorMessage = ""
        End If
        lblDrAccName = ""
    End If
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub txtDrAccNo_KeyPress(KeyAscii As Integer)
    KeyAscii = To_Upper_Case(KeyAscii)
End Sub

Private Sub txtSource_KeyPress(KeyAscii As Integer)
    KeyAscii = To_Upper_Case(KeyAscii)
End Sub



