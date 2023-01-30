VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmJournals 
   Caption         =   "JOURNAL POSTING"
   ClientHeight    =   7860
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   9615
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmJournals.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   9615
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox txtBalance 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   44
      Top             =   2160
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      Height          =   330
      Left            =   1800
      TabIndex        =   43
      Top             =   3480
      Width           =   3495
   End
   Begin VB.CommandButton cmdRemoveu 
      Caption         =   "Remove (Unposted)"
      Height          =   345
      Left            =   7230
      TabIndex        =   38
      Top             =   4320
      Width           =   1590
   End
   Begin VB.CommandButton cmdNewJournal 
      Caption         =   "New Journal"
      Height          =   360
      Left            =   120
      TabIndex        =   37
      Top             =   7440
      Width           =   1185
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print JV"
      Height          =   360
      Left            =   3105
      TabIndex        =   36
      Top             =   150
      Width           =   1095
   End
   Begin MSComctlLib.ListView lvwUnpostedjvs 
      Height          =   2730
      Left            =   120
      TabIndex        =   35
      Top             =   4785
      Visible         =   0   'False
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   4815
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "VoucherNo"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "JV Amount"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Narration"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdPostJournal 
      Caption         =   "&Post"
      Height          =   375
      Left            =   2640
      TabIndex        =   34
      Top             =   7425
      Width           =   1425
   End
   Begin VB.CommandButton CmdUnpostedJV 
      Caption         =   "Unposted Journals"
      Height          =   345
      Left            =   240
      TabIndex        =   33
      Top             =   4200
      Width           =   1770
   End
   Begin RichTextLib.RichTextBox rtpNarration 
      Height          =   270
      Left            =   3645
      TabIndex        =   5
      Top             =   3960
      Visible         =   0   'False
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   476
      _Version        =   393217
      TextRTF         =   $"frmJournals.frx":030A
   End
   Begin VB.ComboBox cboLoanno 
      Height          =   330
      Left            =   1770
      TabIndex        =   4
      Top             =   3000
      Width           =   1485
   End
   Begin VB.TextBox txtJournaNo 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1860
      TabIndex        =   29
      Top             =   165
      Width           =   1230
   End
   Begin VB.ComboBox cboShareType 
      Height          =   330
      Left            =   1785
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   2625
      Width           =   1455
   End
   Begin VB.CommandButton cmdSearchLoan 
      Caption         =   "<>"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3240
      TabIndex        =   23
      Top             =   2970
      Width           =   345
   End
   Begin VB.TextBox txtCr 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6420
      TabIndex        =   2
      Text            =   "0"
      Top             =   840
      Width           =   1275
   End
   Begin VB.TextBox txtDr 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6420
      TabIndex        =   1
      Text            =   "0"
      Top             =   510
      Width           =   1275
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Remove All"
      Height          =   345
      Left            =   6000
      TabIndex        =   20
      Top             =   4320
      Width           =   1230
   End
   Begin VB.CommandButton cmdProcessJournal 
      Caption         =   "&Process"
      Height          =   375
      Left            =   1320
      TabIndex        =   19
      Top             =   7425
      Width           =   1425
   End
   Begin VB.TextBox txtTotalCr 
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
      Height          =   285
      Left            =   8220
      TabIndex        =   18
      Text            =   "0"
      Top             =   3525
      Width           =   1155
   End
   Begin VB.TextBox txtTotalDr 
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
      Height          =   315
      Left            =   8220
      TabIndex        =   17
      Text            =   "0"
      Top             =   2820
      Width           =   1155
   End
   Begin VB.CommandButton cmdAcctsSearch 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1965
      Picture         =   "frmJournals.frx":0381
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   675
      Width           =   330
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   345
      Left            =   4710
      TabIndex        =   15
      Top             =   4320
      Width           =   1230
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Enabled         =   0   'False
      Height          =   345
      Left            =   3525
      TabIndex        =   14
      Top             =   4320
      Width           =   1140
   End
   Begin VB.TextBox txtAccNames 
      Height          =   315
      Left            =   2340
      TabIndex        =   13
      Top             =   675
      Width           =   3300
   End
   Begin VB.ComboBox cboAccno 
      Height          =   330
      Left            =   765
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   690
      Width           =   1200
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "<>"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3240
      TabIndex        =   9
      Top             =   1800
      Width           =   345
   End
   Begin VB.TextBox txtMemberNo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1785
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1800
      Width           =   1455
   End
   Begin MSComctlLib.ListView lvwTrans 
      Height          =   2520
      Left            =   135
      TabIndex        =   8
      Top             =   4770
      Width           =   9285
      _ExtentX        =   16378
      _ExtentY        =   4445
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "AccNo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "DR"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "CR"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "MemberNo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "RefCode"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlg9 
      Left            =   10455
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6555
      TabIndex        =   7
      Top             =   7440
      Width           =   1005
   End
   Begin MSComCtl2.DTPicker dtpReceiptDate 
      Height          =   300
      Left            =   8040
      TabIndex        =   12
      Top             =   225
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   529
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
      CustomFormat    =   "  dd-MM-yyyy"
      Format          =   123928577
      CurrentDate     =   40336
   End
   Begin VB.Frame Frame5 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   1680
      TabIndex        =   39
      Top             =   2040
      Width           =   1935
      Begin VB.OptionButton optSavings 
         Caption         =   "Savings"
         Height          =   315
         Left            =   960
         TabIndex        =   41
         Top             =   120
         Width           =   975
      End
      Begin VB.OptionButton optShares 
         Caption         =   "Shares"
         Height          =   315
         Left            =   0
         TabIndex        =   40
         Top             =   120
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.Label Label6 
      Caption         =   "Ref Doc:"
      Height          =   255
      Left            =   240
      TabIndex        =   42
      Top             =   3240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Total DR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   8475
      TabIndex        =   32
      Top             =   2535
      Width           =   660
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Total CR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   8475
      TabIndex        =   31
      Top             =   3225
      Width           =   675
   End
   Begin VB.Label Label4 
      Caption         =   "Journal Narration"
      Height          =   285
      Left            =   480
      TabIndex        =   30
      Top             =   3480
      Width           =   1305
   End
   Begin VB.Label Label2 
      Caption         =   "Document No"
      Height          =   270
      Left            =   480
      TabIndex        =   28
      Top             =   180
      Width           =   1125
   End
   Begin VB.Label ShareType 
      AutoSize        =   -1  'True
      Caption         =   "ShareType"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   780
      TabIndex        =   27
      Top             =   2655
      Width           =   915
   End
   Begin VB.Label lblShareType 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3600
      TabIndex        =   26
      Top             =   2625
      Width           =   4080
   End
   Begin VB.Label Loanno 
      AutoSize        =   -1  'True
      Caption         =   "Loanno"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   780
      TabIndex        =   25
      Top             =   3000
      Width           =   915
   End
   Begin VB.Label lblLoantype 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3570
      TabIndex        =   24
      Top             =   2970
      Width           =   4110
   End
   Begin VB.Label Label5 
      Caption         =   "CR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5715
      TabIndex        =   22
      Top             =   900
      Width           =   585
   End
   Begin VB.Label Label3 
      Caption         =   "DR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5715
      TabIndex        =   21
      Top             =   525
      Width           =   645
   End
   Begin VB.Label lblfullnames 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3585
      TabIndex        =   11
      Top             =   1800
      Width           =   4095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "MemberNo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   780
      TabIndex        =   10
      Top             =   1830
      Width           =   915
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuAllJournals 
         Caption         =   "All Journals"
      End
      Begin VB.Menu mnuGltrans 
         Caption         =   "GL Correction"
      End
      Begin VB.Menu mnuShareAdjust 
         Caption         =   "Share Capital Adjustment"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmJournals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim totalamount As Currency
Dim pushed As Currency
Dim objLabelEdit As LabelEdit
Dim interestAcc As String, LoanAcc As String
Dim k As Integer
Dim shareBal As Double, Balance As Double
Dim sharesCode As String, VoucherNo As String
Dim isMember As Boolean
Dim memberId As String
Dim subType As String
Dim totalCr As Double, totalDr As Double


Private Sub cboAccno_Change()
    Dim AccNo As String
    AccNo = cboAccno.Text
    sql = "select GLACCNAME,TYPE,SUBTYPE from glsetup where accno='" & AccNo & "'"
    Set rs = oSaccoMaster.GetRecordSet(sql)
    If Not rs.EOF Then
        txtAccNames.Text = rs(0)
        subType = UCase(rs.Fields("SubType"))
        If UCase(rs.Fields(1)) = UCase("MEMBER") Then
            isMember = True
            txtMemberNo.Locked = False
            cmdFind.Enabled = True
            cboLoanno.Locked = False
            cboAccno.Locked = False
            cmdSearchLoan.Enabled = True
            cboAccno_KeyPress 13
        Else
            isMember = False
            txtMemberNo.Locked = True
            txtMemberNo.Text = ""
            cboShareType.Text = " "
            cboLoanno.clear
            cboLoanno.Text = " "
            cboLoanno.Locked = True
            cmdFind.Enabled = False
            cmdSearchLoan.Enabled = False
        End If
    Else
        txtAccNames.Text = ""
    End If
End Sub

Private Sub cboAccno_Click()
    cboAccno_Change
End Sub

Private Sub cboAccno_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtMemberNo.Locked = True Then
            txtDr.SetFocus
        Else
            txtMemberNo.SetFocus
        End If
    End If
End Sub

Private Sub cboLoanno_Change()
    If Trim(cboLoanno.Text) = "" Then
        Exit Sub
    End If
    sql = "select lt.loantype from loantype lt inner join loanbal lb on lt.loancode=lb.loancode where lb.Loanno='" & cboLoanno & "'"
    Set rs = oSaccoMaster.GetRecordSet(sql)
    If rs.EOF Then
        lblLoantype.Caption = ""
    Else
        lblLoantype.Caption = rs(0)
    End If
End Sub

Private Sub cboLoanno_Click()
    cboLoanno_Change
End Sub

Private Sub cboLoanno_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtDr.SetFocus
    End If
End Sub

Private Sub cboShareType_Change()
  Dim bala As String
  
    On Error GoTo Capture
    If optShares.Value = True Then
        sql = "select st.sharestype from sharetype ST where ST.sharescode='" & cboShareType & "'"
        Set rs = oSaccoMaster.GetRecordSet(sql)
        With rs
            If Not .EOF Then
                lblShareType.Caption = rs(0)
            Else
                lblShareType.Caption = ""
            End If
        End With
    Else
        bala = getcubaccbalance(cboShareType)
        bala = bala - 200
        txtBalance = Format(bala, Cfmt)
        sql = "Select ac.AccountName AccName  " _
        & " from AccountCodes ac inner join cub on cub.accountcode=ac.accountcode " _
        & " where cub.Accno='" & cboShareType & "'"
        
        Set rst = oSaccoMaster.GetRecordSet(sql)
        If Not rst.EOF Then
            lblShareType.Caption = rst("AccName")
        Else
            lblShareType.Caption = ""
        End If
    End If
    Exit Sub
Capture:
    ShowErrorMessage Err.Description
End Sub

Private Sub cboShareType_Click()
    cboShareType_Change
End Sub

Private Sub cboShareType_KeyPress(KeyAscii As Integer)
    If cboShareType.Text = "" Then
        cboLoanno.SetFocus
    Else
        txtDr.SetFocus
    End If
End Sub

Private Sub cmdAcctsSearch_Click()
    On Error Resume Next
    frmAcctsSearch.Show vbModal, Me
    If Continue Then
        If SearchValue <> "" Then
            cboAccno = SearchValue
            SearchValue = ""
            Continue = False
        End If
    End If
End Sub

Private Sub cmdadd_Click()
    On Error GoTo SysError
    If cboAccno.Text = "" Then
        Exit Sub
    End If
    
    If Not ValidateEntry(0, 0) Then
        MsgBox ErrorMessage, vbCritical
        Exit Sub
    End If
    
    Set li = lvwTrans.ListItems.Add(, , cboAccno)
     li.SubItems(1) = txtAccNames
    li.SubItems(2) = "0.00"
    li.SubItems(3) = txtMemberNo.Text

    
    Exit Sub
SysError:
    MsgBox Err.Description, vbInformation, Me.Caption
End Sub


Private Sub cmdbookedreceipts_Click()
'//bookedreceipts
    reportname = "bookedreceipts.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, title
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    frmSearchMembers.Show vbModal
    txtMemberNo.Text = SearchValue
End Sub

Private Sub cmdNewJournal_Click()
    Form_Load
End Sub

Private Sub cmdPrint_Click()
    reportname = "JV.rpt"
    STRFORMULA = "{journals.vno}='" & txtJournaNo.Text & "'"
    Show_Sales_Crystal_Report STRFORMULA, reportname, companyname
End Sub



Private Sub cmdRemove_Click()
    On Error GoTo SysError
    With lvwTrans
        If .ListItems.count > 0 Then
            If MsgBox("Do you want to remove " & lvwTrans.SelectedItem.SubItems(1) & _
            " From the list?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
                Exit Sub
            End If
            .ListItems.Remove (.SelectedItem.index)
            Recalculate
            
        End If
    End With
    
    Exit Sub
SysError:
    MsgBox Err.Description, vbInformation, Me.Caption
End Sub


Private Sub cmdRemoveu_Click()
    On Error GoTo Capture
    Dim JvNos As String
    Dim Sel As Integer
    With lvwUnpostedjvs
        If .ListItems.count > 0 Then
            For I = 1 To .ListItems.count
                If .ListItems(I).Checked Then
                    If I = 1 Then
                        JvNos = "'" & .ListItems(I).Text & "'"
                        Sel = 1
                    Else
                        JvNos = JvNos & "," & "'" & .ListItems(I).Text & "'"
                        Sel = Sel + 1
                    End If
                End If
            Next I
            If JvNos <> "" Then
                If MsgBox("Delete the selected " & Sel & " Unposted Records?", vbQuestion + vbYesNo) = vbNo Then
                    Exit Sub
                End If
                If Not oSaccoMaster.Execute("delete from journals where vNo in (" & JvNos & ")") Then
                    MsgBox ErrorMessage
                End If
                CmdUnpostedJV_Click
            End If
        End If
    End With
    Exit Sub
Capture:
    MsgBox IIf(ErrorMessage = "", Err.Description, ErrorMessage)
End Sub

Private Sub cmdProcessJournal_Click()
    On Error GoTo SysError
    'Check that the DR equals the CR
    Dim sourceAcc As String
        If CDbl(txtTotalDr) <> CDbl(txtTotalCr) Then
            MsgBox "The journal is not balancing, please rectify", vbCritical + vbOKOnly
                Exit Sub
        End If
        If Combo1.Text = "" Then
            MsgBox "The Naration is Required", vbCritical
            Combo1.SetFocus
            Exit Sub
        End If
        Set rst = oSaccoMaster.GetRecordSet("select vno from journals where vno='" & txtJournaNo & "'")
        If Not rst.EOF Then
            MsgBox "The Voucherno is already Processed, maybe awaiting Posting"
            Exit Sub
        End If
        oSaccoMaster.ConnectDatabase
        With oSaccoMaster.goConn
        .BeginTrans
            With lvwTrans
                For I = 1 To lvwTrans.ListItems.count
                    If I = 1 Then
                        If CDbl(.ListItems(I).ListSubItems(2)) > 0 Then
                            mMemberno = .ListItems(I).ListSubItems(4)
                            dracc = .ListItems(I).Text
                            CRAcc = ""
                        Else
                            mMemberno = .ListItems(I).ListSubItems(4)
                            CRAcc = .ListItems(I).Text
                            dracc = ""
                        End If
                        sql = ""
                        'i = IIf(DRAcc = "", CDbl(.ListItems(i).ListSubItems(3)), CDbl(.ListItems(i).ListSubItems(2)))
                        sql = "set dateformat dmy insert into Journals(accno,name,Naration,memberno,vno,Amount,Transtype,TRANSDATE,AuditId,Posted,Loanno,sharetype,transactionno)" _
                        & " Values('" & .ListItems(I).Text & "','" & .ListItems(I).ListSubItems(1) & "','" & Combo1.Text & "','" & mMemberno & "'," _
                        & " '" & txtJournaNo & "'," & IIf(dracc = "", CDbl(.ListItems(I).ListSubItems(3)), CDbl(.ListItems(I).ListSubItems(2))) & ",'" & IIf(dracc <> "", "DR", "CR") & "'," _
                        & " '" & dtpReceiptDate & "','" & User & "',0,'" & .ListItems(I).ListSubItems(5) & "','" & .ListItems(I).ListSubItems(5) & "','" & transactionNo & "')"
                        oSaccoMaster.goConn.Execute sql
                        GoTo NNext
                    End If
                    If CRAcc = "" Then
                        CRAcc = .ListItems(I).Text
'                        If Not Save_GLTRANSACTION(dtpReceiptDate, .ListItems(i).ListSubItems(3), DRAcc, CRAcc, txtJournaNo, "JV", User, ErrorMessage, rtpNarration, 0, 1, txtJournaNo) Then
'                            GoTo SysError
'                        End If
                        'CRAcc = ""
                    End If
                    If dracc = "" Then
                        dracc = .ListItems(I).Text
'                        If Not Save_GLTRANSACTION(dtpReceiptDate, .ListItems(i).ListSubItems(2), DRAcc, CRAcc, txtJournaNo, "JV", User, ErrorMessage, rtpNarration, 0, 1, txtJournaNo) Then
'                            GoTo SysError
'                        End If
                        dracc = ""
                    End If
                    'save the jv
                    
                    sql = ""
                    sql = "set dateformat dmy insert into Journals(accno,name,Naration,vno,Amount,Transtype,Memberno,Sharetype,Loanno,TRANSDATE,AuditId,posted)" _
                    & " Values('" & .ListItems(I).Text & "','" & .ListItems(I).ListSubItems(1) & "','" & Combo1.Text & "'," _
                    & " '" & txtJournaNo & "'," & IIf(dracc = "", CDbl(.ListItems(I).ListSubItems(2)), CDbl(.ListItems(I).ListSubItems(3))) & ",'" & IIf(dracc = "", "DR", "CR") & "'," _
                    & " '" & .ListItems(I).ListSubItems(4) & "','" & .ListItems(I).ListSubItems(5) & "','" & .ListItems(I).ListSubItems(5) & "','" & dtpReceiptDate & "','" & User & "',0)"
                    oSaccoMaster.goConn.Execute sql
NNext:
                Next I
            End With
        oSaccoMaster.goConn.CommitTrans
        MsgBox "Journal Processed Successfully"
        lvwTrans.ListItems.clear
        txtJournaNo.Text = JVnumber
        txtTotalCr = 0
        txtTotalDr = 0
    Exit Sub
SysError:
        'If oSaccoMaster.goConn.IsolationLevel = adXactReadUncommitted Then
            .RollbackTrans
        'End If
        End With
    MsgBox Err.Description, vbInformation, Me.Caption
End Sub
Private Function JVnumber()
Dim jvid
    Set rs = oSaccoMaster.GetRecordSet("select top 1 vno from journals order by jvid desc ")
    If Not rs.EOF Then
        jvid = CInt(Mid(rs(0), 4, 5))
        jvid = jvid + 1
    Else
        jvid = 1
    End If
    Select Case jvid
        Case Is < 10
            JVnumber = "JV-0000" & CStr(jvid)
        Case Is < 100
            JVnumber = "JV-000" & CStr(jvid)
        Case Is < 1000
            JVnumber = "JV-00" & CStr(jvid)
        Case Is < 10000
            JVnumber = "JV-0" & CStr(jvid)
        Case Else
            JVnumber = "JV-" & CStr(jvid)
    End Select
    
    
End Function
Public Sub getrefno(memberno As String)

    Set rst = oSaccoMaster.GetRecordSet("select * from CONTRIB where " _
        & "MemberNo='" & txtMemberNo.Text & "' and schemecode='select sharescode from sharetype where UsedToGuarantee=1' order by RefNo desc")
        k = 0
        With rst
            If Not .EOF Then
                If IsNull(!RefNo) Then
                    k = 1
                Else
                    k = !RefNo
                End If
            End If
            k = k + 1
        End With
        'get sharebal to update
        Set rst = oSaccoMaster.GetRecordSet("select * from CONTRIB where " _
        & "MemberNo='" & txtMemberNo.Text & "' and schemecode='L009' order by RefNo desc")
        With rst
            If Not .EOF Then
                shareBal = !shareBal
            End If
        End With

End Sub
Private Function saveReceipt(Receiptno As String, mMemberno As String, ccode As String, name As String, transdate As Date, Amount As Double, chequeno As String, ptype As String) As Boolean
    On Error GoTo Capture
            ErrorMessage = ""
            sql = ""
            sql = "set dateformat dmy INSERT INTO ReceiptBooking (ReceiptNo,MemberNo,Ccode,Name,Transdate," _
            & "Amount, Chequeno, ptype, auditid,datedeposited) VALUES ('" & Receiptno & "','" & _
            mMemberno & "','" & ccode & "','" & name & "','" & transdate & "'," & Amount & ",'" & _
            chequeno & "','" & ptype & "','" & User & "','" & Get_Server_Date & "')"
            oSaccoMaster.ExecuteThis (sql)
            saveReceipt = True
    Exit Function
Capture:
    saveReceipt = False
    ErrorMessage = Err.Description
End Function


Private Sub cmdClear_Click()
    With lvwTrans
        If .ListItems.count > 0 Then
            If MsgBox("Are you sure you want to clear the entire list?", vbQuestion + vbYesNo) = vbYes Then
                Recalculate
                .ListItems.clear
            End If
        End If
    End With
End Sub



Private Sub Command3_Click()
    With lvwTrans
        If .ListItems.count > 0 Then
           pushed = pushed - .SelectedItem.ListSubItems(2)
            .ListItems.Remove (.SelectedItem.index)
        End If
    End With
End Sub

Private Sub cmdPostJournal_Click()
    On Error GoTo SysError
    Dim debitJournal As Boolean, creditJournal As Boolean
    Dim NormalBal As String, Effect As String, Source() As String
    Dim jvSubAmount As Double
    Dim mValue As Double
    Dim Dr As Integer, Cr As Integer
    dracc = ""
    CRAcc = ""
    Loanno = ""
    sharesCode = ""
    
    Set rst = oSaccoMaster.GetRecordSet("select vno from journals where vno='" & txtJournaNo.Text & "'")
    If rst.EOF Then
        MsgBox "The Above Journal has not been processed", vbCritical
        Exit Sub
    End If
    
'    If currentUser.IDNo = memberId Then
'        MsgBox "You Can't Operate your own account", vbCritical
'        Exit Sub
'    End If
    
    If CDbl(txtTotalDr) <> CDbl(txtTotalCr) Then
        MsgBox "The journal is not balancing, please rectify", vbCritical + vbOKOnly
            Exit Sub
    ElseIf MsgBox("Confirm the posting date as " & dtpReceiptDate & "?", vbQuestion + vbYesNo) = vbNo Then
        Exit Sub
    Else
    
    End If
    
    
    With lvwTrans
        For I = 1 To .ListItems.count
            If .ListItems(I).ListSubItems(2) > 0 Then
                Dr = Dr + 1
            Else
                Cr = Cr + 1
            End If
        Next I
        
        If Dr = 1 Then
            debitJournal = True
        Else
            creditJournal = True
        End If
        
        'so which is this Debit/Credit account (Contra)
        For I = 1 To .ListItems.count
            If debitJournal = True Then
                If .ListItems(I).ListSubItems(2) > 0 Then
                    ContraAcc = .ListItems(I).Text
                    GoTo moveOn
                End If
            Else
                If .ListItems(I).ListSubItems(3) > 0 Then
                    ContraAcc = .ListItems(I).Text
                    GoTo moveOn
                End If
            End If
        Next I

        
    End With
        
moveOn:

    'oSaccoMaster.ConnectDatabase
    With oSaccoMaster.goConn
        '// get transactionno
        saveTransno (User)
       transno = getTransno
    memberno = txtMemberNo.Text
    .BeginTrans
    
    'Create a TransactionNo
    transactionTotal = CDbl(txtTotalCr.Text)
    NewTransaction transactionTotal, dtpReceiptDate, "Journal Posting", Mach, transno
 
        With lvwTrans
            If lvwTrans.ListItems.count > 0 Then
                ReDim Source(lvwTrans.ListItems.count)
            End If
        
            For I = 1 To lvwTrans.ListItems.count
                Source(I - 1) = .ListItems(I).Text
            Next I
            For I = 1 To lvwTrans.ListItems.count
                saveToGl = False
                'totalamount = IIf(.ListItems(I).ListSubItems(2).Text > 0, .ListItems(I).ListSubItems(2).Text, .ListItems(I).ListSubItems(3).Text)
                
                'If creditJournal = True Then
                If CDbl(.ListItems(I).ListSubItems(2)) > 0 Then
                    dracc = .ListItems(I).Text
                    jvSubAmount = .ListItems(I).ListSubItems(2)
                    mValue = jvSubAmount
                    Set rst = oSaccoMaster.GetRecordSet("SELECT NORMALBAL,TYPE FROM GLSETUP WHERE ACCNO='" & dracc & "'")
                    If Not success Then
                        GoTo SysError
                    End If
                    If UCase(rst!Type) = UCase("MEMBER") Then
                        isMember = True
                        If rst!NormalBal = "Debit" Then
                            memberno = .ListItems(I).ListSubItems(4).Text
                            Loanno = .ListItems(I).ListSubItems(5).Text
                            sharesCode = Loanno
                            Effect = "Add"
                            If Not Loanno = "" Then
                                If Not effectOnMember(memberno, dracc, Source(IIf(I = 1, 1, 0)), Loanno, jvSubAmount, sharesCode, Effect, transactionNo, transno, Combo1.Text) Then
                                    GoTo SysError
                                End If
                            End If
                        Else
                            Effect = "Less"
                            memberno = .ListItems(I).ListSubItems(4).Text
                            Loanno = .ListItems(I).ListSubItems(5).Text
                            sharesCode = .ListItems(I).ListSubItems(5).Text

                            If Not effectOnMember(memberno, dracc, Source(IIf(I = 1, 1, 0)), Loanno, jvSubAmount, sharesCode, Effect, transactionNo, transno, Combo1.Text) Then
                                GoTo SysError
                            End If
                        End If
                    Else
                        isMember = False
                    End If
                End If
            'Else
                If CDbl(.ListItems(I).ListSubItems(3)) > 0 Then
                    CRAcc = .ListItems(I).Text
                    jvSubAmount = .ListItems(I).ListSubItems(3)
                    mValue = jvSubAmount
                    Set rst = oSaccoMaster.GetRecordSet("SELECT NORMALBAL,TYPE FROM GLSETUP WHERE ACCNO='" & CRAcc & "'")
                    If Not success Then
                        GoTo SysError
                    End If
                    If UCase(rst!Type) = UCase("MEMBER") Then
                        isMember = True
                        If UCase(rst!NormalBal) = "DEBIT" Then
                            memberno = .ListItems(I).ListSubItems(4).Text
                            Loanno = .ListItems(I).ListSubItems(5).Text
    
                            Effect = "Less"
                            If Not effectOnMember(memberno, CRAcc, Source(IIf(I = 1, 1, 0)), Loanno, jvSubAmount, sharesCode, Effect, transactionNo, transno) Then
                                GoTo SysError
                            End If
                        Else
                            Effect = "Add"
                            memberno = .ListItems(I).ListSubItems(4).Text
                            sharesCode = .ListItems(I).ListSubItems(5).Text
                            Loanno = sharesCode
                            If Not effectOnMember(memberno, CRAcc, Source(IIf(I = 1, 1, 0)), Loanno, jvSubAmount, sharesCode, Effect, transactionNo, transno, Combo1.Text) Then
                                GoTo SysError
                            End If
                        End If
                    Else
                        isMember = False
                    End If
                End If
                
                'End If
                
                If debitJournal = True And CRAcc <> "" Then
                        If CRAcc <> ContraAcc Then
                    If Not Save_GLTRANSACTION(dtpReceiptDate, mValue, ContraAcc, CRAcc, txtJournaNo, memberno, User, ErrorMessage, Combo1.Text, 0, 1, txtJournaNo, transactionNo, transno) Then
                        GoTo SysError
                    End If
                   Else
                  End If
                    CRAcc = ""
                ElseIf creditJournal = True And dracc <> "" Then
                     If dracc <> ContraAcc Then
                    If Not Save_GLTRANSACTION(dtpReceiptDate, mValue, dracc, ContraAcc, txtJournaNo, memberno, User, ErrorMessage, Combo1.Text, 0, 1, txtJournaNo, transactionNo, transno) Then
                        GoTo SysError
                    End If
                   Else
                 End If
                    dracc = ""
                End If
                
NNext:
            Next I
            oSaccoMaster.ExecuteThis ("update journals set posted=1 where vno='" & txtJournaNo.Text & "'")
        End With
        
        oSaccoMaster.goConn.CommitTrans
        MsgBox "Journal Posted Successfully"
        lvwTrans.ListItems.clear
        txtJournaNo.Text = JVnumber
        txtTotalCr = 0
        txtTotalDr = 0
        saveToGl = True
    Exit Sub
SysError:
        If oSaccoMaster.goConn.State = adStateOpen Then
            oSaccoMaster.goConn.RollbackTrans
        End If
        saveToGl = True
        End With
    MsgBox IIf(ErrorMessage = "", Err.Description, ErrorMessage), vbInformation, Me.Caption
End Sub
Private Function effectOnMember(mMemberno As String, AccNo As String, Source As String, Loanno As String, Amount As Double, sharesCode As String, Effect As String, Optional transactionNo As String, Optional transno As Integer, Optional Remarks As String) As Boolean
    On Error GoTo Capture
    Dim SomethingDone As Boolean
    SomethingDone = False
    sql = "SELECT AccNo, GlAccName, NormalBal,subType FROM GLSETUP WHERE (Type = 'MEMBER') and accno ='" & AccNo & "'"
    Set rst = oSaccoMaster.GetRecordSet(sql)
    With rst
        If Not .EOF Then
            .Filter = "SubType='LOAN'"
            If Not .EOF Then ' it's a loan account
                If Effect = "Add" Then
                      Transcode = Combo1.Text
                    If Not SaveRepay(Loanno, Format(dtpReceiptDate, "dd/mm/yyyy"), Amount * -1, Source, txtJournaNo, 0, 0, Remarks, User, mMemberno, transactionNo, transno, "", False, , , Transcode, Effect) Then
                        effectOnMember = False
                        Exit Function
                    End If
                Else
                    Transcode = Combo1.Text
                    If Not SaveRepay(Loanno, Format(dtpReceiptDate, "dd/mm/yyyy"), Amount, Source, txtJournaNo, 0, 0, Remarks, User, mMemberno, transactionNo, transno, False, False, , , Transcode, Effect) Then
                        effectOnMember = False
                        Exit Function
                    End If
                End If
                SomethingDone = True
                GoTo proceed
            End If
            .Filter = adFilterNone
            .Filter = "SubType='INTEREST'"
            If Not .EOF Then ' it's a interest account
                
                sql = " select max(paymentno) from repay where loanno='" & Loanno & "' "
                Set rs = oSaccoMaster.GetRecordSet(sql)
                If Not IsNull(rs.EOF) Then
                    PaymentNo = IIf(IsNull(rs(0)) = True, 0, rs(0)) + 1
                Else
                    PaymentNo = 1
                End If
                
                

                If Effect = "Add" Then
                    sql = "set dateformat dmy insert into Repay(loanno,datereceived,paymentno,amount,principal,interest,penalty,intrcharged,introwed,intbalance,loanbalance,receiptno,remarks,auditid,transactionno,Transno)" _
                    & " Values ('" & Loanno & "','" & Format(dtpReceiptDate, "dd/mm/yyyy") & "'," & PaymentNo & "," & Amount & ",0," & Amount & ",0,0,0," & IntBalalance - Amount & ",0,'" & txtJournaNo & "','" & Combo1.Text & "','" & auditid & "','" & transactionNo & "','" & transno & "')"
                
                
                Else
                    sql = "set dateformat dmy insert into Repay(loanno,datereceived,paymentno,amount,principal,interest,penalty,intrcharged,introwed,intbalance,loanbalance,receiptno,remarks,auditid,transactionno,transno)" _
                    & " Values ('" & Loanno & "','" & Format(dtpReceiptDate, "dd/mm/yyyy") & "'," & PaymentNo & "," & Round(Amount, 2) * (-1) & ",0," & Amount * (-1) & ",0,0,0," & IntBalalance - Amount & ",0,'" & txtJournaNo & "','" & Combo1.Text & "','" & auditid & "','" & transactionNo & "','" & transno & "')"
                    
                End If
                If Not oSaccoMaster.Execute(sql) Then
                    GoTo Capture
                ElseIf Not oSaccoMaster.Execute("update loanbal set intbalance=" & IntBalalance - Amount & " where loanno='" & Loanno & "'") Then
                    GoTo Capture
                End If
                
                SomethingDone = True
'                If Not Refresh_Loan(Loanno) Then
'
'                End If
                GoTo proceed
            Else
                .Filter = adFilterNone
                .Filter = "SubType= 'SHARE'"
                If Not .EOF Then 'Affects the share/deposit account
                    If Effect = "Add" Then
                        If Not SaveContrib(mMemberno, dtpReceiptDate, sharesCode, Amount, "", txtJournaNo, txtJournaNo, User, txtJournaNo, transactionNo, Combo1, , , Effect) Then
                                effectOnMember = False
                            Exit Function
                        End If
                    Else
                        Effect = "Less"
                        If Not SaveContrib(mMemberno, dtpReceiptDate, sharesCode, Amount * (-1), "", txtJournaNo, txtJournaNo, "Journal", txtJournaNo, transactionNo, Combo1, , , Effect) Then
                                effectOnMember = False
                            Exit Function
                        End If
                    End If
                End If
                SomethingDone = True
                
                .Filter = adFilterNone
                .Filter = "SubType= 'SAVING'"
                If Not .EOF Then 'Affects the savings account
                    saveToGl = False
                    If Effect = "Add" Then
                        If Not Deposit("", sharesCode, dtpReceiptDate, Amount, "JE", txtJournaNo, "", transactionNo, transno, "JE", Remarks, Combo1.Text) Then
                            effectOnMember = False
                            SomethingDone = False
                        End If
                    Else
                        If Not Withdraw("", sharesCode, dtpReceiptDate, Amount, "JE", txtJournaNo, "", transactionNo, transno, Remarks, False, Combo1.Text) Then
                            effectOnMember = False
                            SomethingDone = False
                        End If
                    End If
                End If
                SomethingDone = True
                
            End If
proceed:
        End If
        If SomethingDone = False Then
            ErrorMessage = "Effect on member could not be done. Check the member accounts namings"
            GoTo Capture
        End If
        effectOnMember = True
    End With
    Exit Function
Capture:
    effectOnMember = False
End Function

Private Sub CmdUnpostedJV_Click()
    On Error Resume Next
    lvwUnpostedjvs.ListItems.clear
    sql = "select vno,transdate,naration,(sum(amount)/2)Amount from journals where posted =0 group by vno,transdate,naration"
    Set rst = oSaccoMaster.GetRecordSet(sql)
    With rst
        If .EOF Then Exit Sub
        lvwUnpostedjvs.Visible = True
        While Not .EOF
            Set li = lvwUnpostedjvs.ListItems.Add(, , .Fields(0))
            li.ListSubItems.Add , , .Fields(1)
            li.ListSubItems.Add , , .Fields(2)
            li.ListSubItems.Add , , .Fields(3)
        .MoveNext
        Wend
    End With
End Sub

Private Sub Combo1_Change()
'Set rs = oSaccoMaster.GetRecordSet("select TransactionType from TransactionCode where TranscationCode ='" & Combo1 & "'")
'     If Not rs.EOF Then
'       rtpNarration.Text = rs(0)
'     End If
End Sub

Private Sub Combo1_Click()
Combo1_Change
End Sub
Private Sub Form_Load()
    On Error Resume Next
    Dim rscompany As New ADODB.Recordset
    'Load Gl's
    txtBalance.Visible = False
      txtBalance.Text = 0#
    sql = "Select accno from glsetup order by accno asc"
    Set rst = oSaccoMaster.GetRecordSet(sql)
    While Not rst.EOF
        cboAccno.AddItem (rst(0))
        rst.MoveNext
    Wend
    'load hareCodes
     sql = "Select Description from TransCodebosa  order by Transactioncode "
    Set rst1 = oSaccoMaster.GetRecordSet(sql)
    While Not rst1.EOF
        Combo1.AddItem (rst1(0))
        rst1.MoveNext
    Wend
    cboShareType.clear
    Set rst = oSaccoMaster.GetRecordSet("select sharescode from sharetype")
    While Not rst.EOF
        cboShareType.AddItem rst(0)
        rst.MoveNext
    Wend
    'initialization
    totalamount = 0
    pushed = 0
    txtJournaNo.Text = JVnumber
    
    dtpReceiptDate.Value = Get_Server_Date
    
End Sub

Private Sub optotherpayments_Click()
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Stop subclassing
    CloseSubClass
    'Clean up by setting the classes to Nothing
    Set objLabelEdit = Nothing
    'Set objLabelEdit2 = Nothing
End Sub

Private Sub lvwTrans_Click()
    Dim total As Double, amt As Double
    Dim ccount As Integer
    total = 0
    With lvwTrans
        If .ListItems.count > 0 Then
            ccount = .ListItems.count
            For I = 1 To ccount
                With .ListItems(I)
                        amt = CDbl(.ListSubItems(2))
                        total = total + amt
                End With
            Next I

        Else
            total = 0
        End If
    End With
'    txtDistributed.Text = total
    

End Sub

Private Sub lvwTrans_ItemClick(ByVal Item As MSComctlLib.ListItem)
'    On Error GoTo SysError
'    If lvwTrans.ListItems.count > 0 Then
'        txtAmount = lvwTrans.SelectedItem.SubItems(2)
'    End If
'    Exit Sub
'SysError:
'    MsgBox Err.Description, vbInformation, Me.Caption
End Sub

Private Sub lvwTrans_KeyPress(KeyAscii As Integer)
    MsgBox KeyAscii
End Sub




Private Sub txtStaffNo_Change()
Call txtStaffNo_Change
End Sub

Private Sub lvwUnpostedjvs_DblClick()

    Dim vno As String
    txtTotalCr = 0
    txtTotalDr = 0
    With lvwUnpostedjvs
        If .ListItems.count > 0 Then
            vno = .SelectedItem.Text
            VoucherNo = vno
            txtJournaNo.Text = vno
            sql = "select accno,name,Naration,vno,memberno,sharetype,loanno,Amount,Transtype,TRANSDATE,AuditId,Posted from journals where vno='" & vno & "'"
            Set rst = oSaccoMaster.GetRecordSet(sql)
            If Not rst.EOF Then
                rtpNarration = rst!naration
                With lvwTrans
                    .ListItems.clear
                    While Not rst.EOF
                        Set li = .ListItems.Add(, , rst!AccNo)
                        li.ListSubItems.Add , , rst!name
                        li.ListSubItems.Add , , IIf(Trim(rst!transtype) = "DR", rst!Amount, 0)
                        li.ListSubItems.Add , , IIf(Trim(rst!transtype) = "CR", rst!Amount, 0)
                        li.ListSubItems.Add , , rst!memberno
                        li.ListSubItems.Add , , IIf(IsNull(rst!ShareType), "", rst!ShareType)
                        li.ListSubItems.Add , , IIf(IsNull(rst!Loanno), "", rst!Loanno)
                        txtDr.Text = 0
                        txtCr.Text = 0
                        rst.MoveNext
                    Wend
                    Recalculate
                    lvwUnpostedjvs.Visible = False
                End With
            End If
        End If
    End With
End Sub

Private Sub mnuAllJournals_Click()
    reportname = "JV2.rpt"
    Show_Sales_Crystal_Report "", reportname, companyname
End Sub

Private Sub mnuGltrans_Click()
 frmGltransactionview.Show vbModal, Me
End Sub

Private Sub mnuShareAdjust_Click()
    'frmShareCapUpdate.Show vbModal
End Sub

 Private Sub optSavings_Click()
    cboShareType.clear
    lblShareType.Caption = ""
    If OptSavings.Value = True Then
    
    txtBalance.Visible = True
    Else
     txtBalance.Visible = False
     txtBalance.Text = 0
    End If
    
    sql = "select Accno from cub where memberno='" & txtMemberNo & "' order by 1"
    Set rst = oSaccoMaster.GetRecordSet(sql)
    If Not rst.EOF Then
        While Not rst.EOF
            cboShareType.AddItem rst(0)
            rst.MoveNext
        Wend
    End If
End Sub

Private Sub optShares_Click()
   If optShares.Value = True Then
     txtBalance.Visible = False
      txtBalance.Text = 0#
     Else
     txtBalance.Visible = True
   End If
    cboShareType.clear
    lblShareType.Caption = ""
    sql = "select sharescode from sharetype"
    Set rst = oSaccoMaster.GetRecordSet(sql)
    If Not rst.EOF Then
        While Not rst.EOF
            cboShareType.AddItem rst(0)
            rst.MoveNext
        Wend
    End If
End Sub

Private Sub txtCr_GotFocus()
     txtCr = SelectAllText(txtCr)
End Sub

Private Sub txtCr_keypress(KeyAscii As Integer)
   Dim balances As Double
    Dim Source As String
    If Not keyIsValid(KeyAscii, 1) Then
        Beep
        KeyAscii = 0
    End If
    balances = Format(txtBalance, Cfmt)
      If OptSavings.Value = True Then
        If txtCr = "" Then
        Exit Sub
        End If
'      If CDbl(txtCr) > CDbl(txtBalance) Then
'       MsgBox "The amount you Enter is great than the Available", vbInformation
'       Exit Sub
'       End If
    End If
      
    Select Case KeyAscii
        Case 13
            
            If Not ValidateEntry(0, 1) Then
                MsgBox ErrorMessage, vbCritical, "ENTRY NOT VALID"
                cboLoanno.SetFocus
                Exit Sub
            End If
            
            Source = IIf(cboShareType.Text = "", cboLoanno.Text, cboShareType.Text)
            If Source = "" And isMember = True And subType <> "OTHERS" Then
                MsgBox "You must select the Share/Saving or Loan Account", vbCritical
                Exit Sub
            End If
            
            If CDbl(txtCr) > 0 Then
                With lvwTrans
                    Set li = .ListItems.Add(, , cboAccno)
                    li.ListSubItems.Add , , txtAccNames
                    li.ListSubItems.Add , , 0
                    li.ListSubItems.Add , , txtCr
                    li.ListSubItems.Add , , txtMemberNo
                    li.ListSubItems.Add , , IIf(cboShareType.Text = "", cboLoanno.Text, cboShareType.Text)
                    'li.ListSubItems.Add , , cboLoanno
                    txtTotalCr = CDbl(txtTotalCr) + CDbl(txtCr)
                    txtCr.Text = 0
                End With
            End If
            cboAccno.SetFocus
            cboShareType.Text = ""
            cboLoanno.Text = ""
            Recalculate
        Case Else
            
    End Select
End Sub



Private Sub txtCr_LostFocus()
    txtCr = IIf(txtCr.Text = "", 0, txtCr)
    
End Sub

Private Sub txtDr_GotFocus()
 txtDr = SelectAllText(txtDr)
End Sub

Private Sub txtDr_KeyPress(KeyAscii As Integer)
    Dim Source As String
    Dim balances As Double
    If Not keyIsValid(KeyAscii, 1) Then
        Beep
        KeyAscii = 0
    End If
    Select Case KeyAscii
        Case 13
        
            If Not ValidateEntry(1, 0) Then
                MsgBox ErrorMessage, vbCritical
                cboLoanno.SetFocus
                Exit Sub
            End If
         balances = Format(txtBalance, Cfmt)
      If OptSavings.Value = True Then
        If txtDr = "" Then
        Exit Sub
        End If
      If CDbl(txtDr) > CDbl(txtBalance) Then
       MsgBox "The amount you Enter is great than the Available Balance", vbInformation
       Exit Sub
       End If
      End If
        Source = IIf(cboShareType.Text = "", cboLoanno.Text, cboShareType.Text)
            If Source = "" And isMember = True And subType <> "OTHERS" Then
                MsgBox "You must select the Share/Saving or Loan Account", vbCritical
                Exit Sub
            End If
            
            If CDbl(txtDr) > 0 Then
                With lvwTrans
                    Set li = .ListItems.Add(, , cboAccno)
                    li.ListSubItems.Add , , txtAccNames
                    li.ListSubItems.Add , , txtDr
                    li.ListSubItems.Add , , 0
                    li.ListSubItems.Add , , txtMemberNo
                    li.ListSubItems.Add , , IIf(cboShareType.Text = "", cboLoanno.Text, cboShareType.Text)
                    txtDr.Text = 0
                End With
            End If
            txtCr.SetFocus
            cboShareType.Text = ""
            cboLoanno.Text = ""
            Recalculate
        Case Else
            
    End Select
End Sub

Function ValidateEntry(addedDR As Double, addedCR As Double) As Boolean

    Dim Dr As Integer, Cr As Integer

    If subType = "LOAN" Or subType = "INTEREST" Then
        If cboLoanno.Text = "" Then
            ErrorMessage = "Choose the Loan that is affected by this Journal"
            txtDr_LostFocus
            Exit Function
        End If
    ElseIf subType = "SHARE" Then
        If cboShareType.Text = "" Then
            ErrorMessage = "Choose the share/Saving type that is affected by the Journal"
            txtDr_LostFocus
            Exit Function
        End If
    End If
    
    With lvwTrans
        For I = 1 To .ListItems.count
'            If cboAccno.Text = .ListItems(I).Text Then
'                ErrorMessage = "This account have already been added to the list"
'                Exit Function
'            End If
            
            If .ListItems(I).ListSubItems(2) > 0 Then
                Dr = Dr + 1
            Else
                Cr = Cr + 1
            End If
        Next I
        
        If Dr + addedDR > 1 And Cr + addedCR > 1 Then
            ErrorMessage = "A journal must have a single entry in atleast one side!"
            Exit Function
        End If
    End With
    ValidateEntry = True
    Exit Function
    
End Function

Private Sub txtDr_LostFocus()
    txtDr.Text = IIf(txtDr.Text = "", 0, txtDr.Text)
End Sub

Private Sub txtMemberNo_Change()
    mysql = ""
    mysql = "select surname,othernames,HomeAddr,companycode,idno  from members  where memberno ='" & txtMemberNo & "'"
    Set rs = oSaccoMaster.GetRecordSet(mysql)
    If Not rs.EOF Then
        lblfullnames = rs!surname & "  " & rs!OtherNames
        memberId = rs!idno
        cmdAdd.Enabled = True
    Else
        lblfullnames = ""
        cmdAdd.Enabled = False
        Exit Sub
    End If
End Sub

Private Sub txtMemberNo_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13
            If UCase(subType) = "SHARE" Then
                cboShareType.Locked = False
                cboShareType.SetFocus
                cboLoanno.Text = ""
                cboLoanno.Locked = True
                optShares.Value = True
                optShares_Click
            ElseIf UCase(subType) = "SAVING" Then
                cboShareType.Locked = False
                cboShareType.SetFocus
                cboLoanno.Text = ""
                cboLoanno.Locked = True
                OptSavings.Value = True
                optSavings_Click
            ElseIf UCase(subType) = "LOAN" Or UCase(subType) = "INTEREST" Then
                sql = "select loanno from loanbal where memberno='" & txtMemberNo & "'"
                Set rst = oSaccoMaster.GetRecordSet(sql)
                With rst
                    cboLoanno.clear
                    While Not .EOF
                        cboLoanno.AddItem rst(0)
                        .MoveNext
                    Wend
                End With
                cboShareType.Locked = True
                cboShareType.Text = ""
                cboLoanno.Locked = False
                cboLoanno.SetFocus
            End If
        Case Else
            Exit Sub
    End Select
End Sub
Sub Recalculate()
    totalCr = 0
    totalDr = 0
    With lvwTrans
        For I = 1 To .ListItems.count
            If .ListItems(I).ListSubItems(2) > 0 Then
                totalDr = totalDr + .ListItems(I).ListSubItems(2)
            Else
                totalCr = totalCr + .ListItems(I).ListSubItems(3)
            End If
        Next I
        txtTotalCr = totalCr
        txtTotalDr = totalDr

    End With
End Sub



