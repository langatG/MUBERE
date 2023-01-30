VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPostCashBook 
   Caption         =   "Post Cash Book Entries"
   ClientHeight    =   5985
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9780
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPostCashBook.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   9780
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker DTPPeriod 
      Height          =   315
      Left            =   6195
      TabIndex        =   20
      Top             =   2505
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   " MMM-yyyy"
      Format          =   71499779
      CurrentDate     =   39506
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4320
      TabIndex        =   17
      Top             =   2520
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ComboBox cboTransType 
      Height          =   330
      ItemData        =   "frmPostCashBook.frx":0442
      Left            =   2025
      List            =   "frmPostCashBook.frx":044F
      TabIndex        =   16
      Top             =   2520
      Visible         =   0   'False
      Width           =   2295
   End
   Begin MSComctlLib.ListView lvwEntries 
      Height          =   240
      Left            =   105
      TabIndex        =   14
      Top             =   5265
      Visible         =   0   'False
      Width           =   8220
      _ExtentX        =   14499
      _ExtentY        =   423
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
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "MemberNo"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Receipt No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Names"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Amount"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Transdate"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdUnposted 
      Caption         =   "Unposted Entries"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6225
      TabIndex        =   15
      Top             =   525
      Width           =   2010
   End
   Begin VB.CommandButton cmdPost 
      Caption         =   "&Post"
      Height          =   345
      Left            =   6840
      TabIndex        =   13
      Top             =   5520
      Width           =   1455
   End
   Begin MSComctlLib.ListView lvwNames 
      Height          =   1395
      Left            =   1785
      TabIndex        =   4
      Top             =   855
      Visible         =   0   'False
      Width           =   3765
      _ExtentX        =   6641
      _ExtentY        =   2461
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
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Names"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "MemberNo"
         Object.Width           =   9
      EndProperty
   End
   Begin VB.TextBox txtExpectedTotals 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6315
      TabIndex        =   11
      Top             =   5040
      Width           =   2000
   End
   Begin VB.TextBox txtInterest 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3570
      TabIndex        =   10
      Top             =   5040
      Width           =   1590
   End
   Begin VB.TextBox txtPrincipal 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2010
      TabIndex        =   9
      Top             =   5040
      Width           =   1575
   End
   Begin MSComctlLib.ListView lvwUnposted 
      Height          =   1005
      Left            =   90
      TabIndex        =   5
      Top             =   1260
      Width           =   8220
      _ExtentX        =   14499
      _ExtentY        =   1773
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "TransDate"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ReceiptNo"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Amount"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Cheque No"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox txtNames 
      Height          =   315
      Left            =   1785
      TabIndex        =   2
      Top             =   540
      Width           =   3765
   End
   Begin VB.TextBox txtMemberNo 
      Height          =   315
      Left            =   105
      TabIndex        =   0
      Top             =   540
      Width           =   1695
   End
   Begin MSComctlLib.ListView lvwMemberDeductions 
      Height          =   1905
      Left            =   90
      TabIndex        =   6
      Top             =   2850
      Width           =   9630
      _ExtentX        =   16986
      _ExtentY        =   3360
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Deduction"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Amount / Principal"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Interest"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Balance"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "LoanNo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "LoanCode"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CheckBox chkTransactions 
      Caption         =   "Other Transactions"
      Height          =   210
      Left            =   2025
      TabIndex        =   19
      Top             =   2280
      Width           =   2325
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Period"
      Height          =   210
      Left            =   6330
      TabIndex        =   21
      Top             =   2265
      Width           =   510
   End
   Begin VB.Label lblAmount 
      AutoSize        =   -1  'True
      Caption         =   "Amount"
      Height          =   210
      Left            =   5175
      TabIndex        =   18
      Top             =   2280
      Width           =   660
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Expected Totals"
      Height          =   210
      Left            =   6945
      TabIndex        =   12
      Top             =   4770
      Width           =   1335
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Expected Deductions"
      Height          =   210
      Left            =   120
      TabIndex        =   8
      Top             =   2610
      Width           =   1755
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Unposted Entries"
      Height          =   210
      Left            =   105
      TabIndex        =   7
      Top             =   1020
      Width           =   1410
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Member Names"
      Height          =   210
      Left            =   1800
      TabIndex        =   3
      Top             =   285
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "MemberNo"
      Height          =   210
      Left            =   120
      TabIndex        =   1
      Top             =   285
      Width           =   885
   End
End
Attribute VB_Name = "frmPostCashBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TotalDed As Double

Private Sub chkTransactions_Click()
    On Error GoTo SysError
    If chktransactions.value = vbChecked Then
        lblAmount.Visible = True
        txtAmount.Visible = True
        cboTransType.Visible = True
        txtAmount = "0.00"
    Else
        lblAmount.Visible = False
        txtAmount.Visible = False
        cboTransType.Visible = False
    End If
    Exit Sub
SysError:
    MsgBox Err.description, vbInformation, Me.Caption
End Sub

Private Sub cmdPost_Click()
    On Error GoTo SysError
    If Trim(txtmemberno) <> "" Then
        If lvwUnposted.ListItems.Count > 0 Then
            If CDbl(txtExpectedTotals) <> CDbl(lvwUnposted.ListItems(1).SubItems(2)) Then
                MsgBox "The Distributed Amount should same as Received Amount", _
                vbInformation, Me.Caption
                Exit Sub
            Else
                'Post The Transaction
                For i = 1 To lvwMemberDeductions.ListItems.Count
                    If lvwMemberDeductions.ListItems(i).Checked = True Then
                        Set li = lvwMemberDeductions.ListItems(i)
                        'XXXXXXXXXXX Post Shares XXXXXXXXXXXXXXX
                        TransNo = Get_Trans_No(txtmemberno & Get_Server_Date, ErrorMessage)
                        If li = "Shares Contribution" Then
                            If Not Save_Contrib(txtmemberno, lvwUnposted.ListItems(1), _
                            1000, CDbl(li.SubItems(1)), 1000, txtmemberno, lvwUnposted.ListItems(1).SubItems(1), _
                            lvwUnposted.ListItems(1).SubItems(1), "No", "No", "Cash Receipt", User, TransNo, _
                            DTPPeriod, ErrorMessage) Then
                                If ErrorMessage <> "" Then
                                    MsgBox ErrorMessage, vbInformation, Me.Caption
                                    ErrorMessage = ""
                                End If
                            End If
                            'XXXXXXXXXXXXXXXXXX UPDATE GL WITH THE SHARES AMOUNT XXXXXXXXXXXXXXXXXXXXX
                            Select Case CDbl(li.SubItems(1))
                                Case Is > 0
                                If Not Save_To_GL("A001", "L009", CDbl(li.SubItems(1)), txtmemberno, _
                                lvwUnposted.ListItems(1).SubItems(1), lvwUnposted.ListItems(1), _
                                txtmemberno, txtNames, ErrorMessage, "Cash ag_receipts") Then
                                    If ErrorMessage <> "" Then
                                        MsgBox ErrorMessage, vbInformation, Me.Caption
                                        ErrorMessage = ""
                                    End If
                                End If
                                Case Is < 0
                                If Not Save_To_GL("L009", "A001", CDbl(li.SubItems(1)) * (-1), txtmemberno, _
                                lvwUnposted.ListItems(1).SubItems(1), lvwUnposted.ListItems(1), _
                                txtmemberno, txtNames, ErrorMessage, "Cash ag_receipts") Then
                                    If ErrorMessage <> "" Then
                                        MsgBox ErrorMessage, vbInformation, Me.Caption
                                        ErrorMessage = ""
                                    End If
                                End If
                            End Select
                        'XXXXXXXXXXX Post the Loan Repayment XXXXXXXXXXXXXXXXX
                        ElseIf li = "Loan Repayment" Then
                            If Not Save_Repayment(li.SubItems(4), txtmemberno, _
                            lvwUnposted.ListItems(1), 1000, CDbl(li.SubItems(1)) _
                            + CDbl(li.SubItems(2)), CDbl(li.SubItems(1)), CDbl(li.SubItems(2)), _
                            0, 0, 1000, txtmemberno, 0, 0, 0, "Cash Receipt", _
                            User, "", 0, Get_Server_Date, "", DTPPeriod, ErrorMessage) Then
                                If ErrorMessage <> "" Then
                                    MsgBox ErrorMessage, vbInformation, Me.Caption
                                    ErrorMessage = ""
                                End If
                            End If
                            Select Case CDbl(li.SubItems(1)) 'PRINCIPAL
                                Case Is > 0
                                If Not Save_To_GL("A001", "A007", CDbl(li.SubItems(1)), txtmemberno, _
                                lvwUnposted.ListItems(1).SubItems(1), lvwUnposted.ListItems(1), _
                                txtmemberno, txtNames, ErrorMessage, _
                                "Cash ag_receipts") Then
                                    If ErrorMessage <> "" Then
                                        MsgBox ErrorMessage, vbInformation, Me.Caption
                                        ErrorMessage = ""
                                    End If
                                End If
                                Case Is < 0
                                If Not Save_To_GL("A007", "A001", CDbl(li.SubItems(1)) * (-1), txtmemberno, _
                                lvwUnposted.ListItems(1).SubItems(1), lvwUnposted.ListItems(1), _
                                txtmemberno, txtNames, ErrorMessage, _
                                "Cash ag_receipts") Then
                                    If ErrorMessage <> "" Then
                                        MsgBox ErrorMessage, vbInformation, Me.Caption
                                        ErrorMessage = ""
                                    End If
                                End If
                            End Select
                            Select Case CDbl(li.SubItems(2)) 'INTEREST
                                Case Is > 0
                                If Not Save_To_GL("A001", "I001", CDbl(li.SubItems(2)), txtmemberno, _
                                lvwUnposted.ListItems(1).SubItems(1), lvwUnposted.ListItems(1), _
                                txtmemberno, txtNames, ErrorMessage, "Cash ag_receipts") Then
                                    If ErrorMessage <> "" Then
                                        MsgBox ErrorMessage, vbInformation, Me.Caption
                                        ErrorMessage = ""
                                    End If
                                End If
                                Case Is < 0
                            End Select
                            'XXXXXXXXXXXXXXXX UPDATE THE GL ACCOUNTS INTEREST XXXXXXXXXXXXXXXXX

                            'XXXXXXXXXXXXXXXX UPDATE THE GL ACCOUNTS PRINCIPAL XXXXXXXXXXXXXXXXX
'
                        End If
                    End If
                Next i
                Set rsMembership = oSaccoMaster.GetRecordset("Update CashBook Set Posted=1 where " _
                & "MemberNo='" & txtmemberno & "' and Posted=0 and ReceiptNo='" & _
                lvwUnposted.ListItems(1).SubItems(1) & "'")
                txtMemberNo_Change
            End If
        Else
            MsgBox "There is no unposted Receipt for this Member", vbInformation, Me.Caption
            Exit Sub
        End If
    Else
        MsgBox "Please Select A Member", vbInformation, Me.Caption
        Exit Sub
    End If
    MsgBox "Record Posted Successfully", vbInformation, Me.Caption
    lblAmount.Visible = False
    txtAmount.Visible = False
    txtAmount = "0.00"
    cboTransType.Visible = False
    Exit Sub
SysError:
    MsgBox Err.description, vbInformation, Me.Caption
End Sub

Private Sub cmdUnposted_Click()
    On Error GoTo SysError
    With lvwEntries
        .Move 105, 270, .Width, 5190
        .Visible = True
    End With
    Dim rsEntries As New Recordset
    lvwEntries.ListItems.Clear
    Set rsEntries = oSaccoMaster.GetRecordset("Select * From CASHBOOK where " _
    & "Posted=0")
    With rsEntries
        If .State = adStateOpen Then
            While Not .EOF
                Set li = lvwEntries.ListItems.Add(, , IIf(IsNull(!memberno), "", !memberno))
                li.SubItems(1) = IIf(IsNull(!Receiptno), "", !Receiptno)
                li.SubItems(3) = Format(IIf(IsNull(!amount), 0, !amount), Cfmt)
                li.SubItems(4) = Format(!transdate, "dd-MM-yyyy")
                .MoveNext
            Wend
        End If
    End With
    Exit Sub
SysError:
    MsgBox Err.description, vbInformation, Me.Caption
End Sub

Private Sub Form_Activate()
    On Error GoTo SysError
    cboTransType.ListIndex = 0
    If MyRecord <> "" Then
        txtmemberno = MyRecord
    End If
    memnum = ""
    Exit Sub
SysError:
    MsgBox Err.description, vbInformation, Me.Caption
End Sub

Private Sub Form_Load()
    DTPPeriod = Get_Server_Date
End Sub

Private Sub lvwEntries_DblClick()
    On Error GoTo SysError
    If lvwEntries.ListItems.Count > 0 Then
        txtmemberno = lvwEntries.SelectedItem
    End If
    lvwEntries.Visible = False
    Exit Sub
SysError:
    MsgBox Err.description, vbInformation, Me.Caption
End Sub

Private Sub lvwMemberDeductions_Click()
    On Error GoTo SysError
    If lvwMemberDeductions.ListItems.Count > 0 Then
        Set li = lvwMemberDeductions.SelectedItem
        txtPrincipal = li.SubItems(1)
        If li.SubItems(3) <> "" Then
            txtInterest = li.SubItems(2)
        Else
            txtInterest = "0.00"
        End If
    End If
    Get_Totals
    txtPrincipal.SetFocus
    SendKeys "{Home}+{End}"
    Exit Sub
SysError:
    MsgBox Err.description, vbInformation, Me.Caption
End Sub

Private Sub lvwNames_Click()
    On Error GoTo SysError
    If lvwNames.ListItems.Count > 0 Then
        Editing = True
        Set li = lvwNames.SelectedItem
        txtmemberno = li.SubItems(1)
        txtNames = li
        Editing = False
        lvwNames.Visible = False
    End If
    Exit Sub
SysError:
    MsgBox Err.description, vbInformation, Me.Caption
End Sub

Private Sub txtAmount_KeyPress(KeyAscii As Integer)
    On Error GoTo SysError
    Select Case KeyAscii
        Case Is = 8
        Case Is = 46
        Case 48 To 57
        Case 13
        If Trim$(txtAmount) <> "" Then
            Set li = lvwMemberDeductions.ListItems.Add(, , cboTransType)
            li.SubItems(1) = Format(txtAmount, Cfmt)
            li.SubItems(2) = "0.00"
            txtAmount = "0.00"
            SendKeys "{Home}+{End}"
        End If
    End Select
    Get_Totals
    Exit Sub
SysError:
    MsgBox Err.description, vbInformation, Me.Caption
End Sub

Private Sub txtInterest_KeyPress(KeyAscii As Integer)
    On Error GoTo SysError
    Select Case KeyAscii
        Case Is = 8
        Case Is = 46
        Case 48 To 57
        Case 13
        If Trim$(txtInterest) <> "" Then
            Set li = lvwMemberDeductions.SelectedItem
            If li.SubItems(3) <> "" Then
                li.SubItems(2) = Format(txtInterest, Cfmt)
            End If
        End If
    End Select
    Get_Totals
    Exit Sub
SysError:
    MsgBox Err.description, vbInformation, Me.Caption
End Sub

Private Sub txtMemberNo_Change()
    On Error GoTo SysError
    lvwMemberDeductions.ListItems.Clear
    lvwUnposted.ListItems.Clear
    If Trim$(txtmemberno) <> "" Then
        Editing = True
        Load_Details txtmemberno
        Editing = False
    End If
    Exit Sub
SysError:
    MsgBox Err.description, vbInformation, Me.Caption
End Sub

Private Sub txtMemberNo_KeyPress(KeyAscii As Integer)
    On Error GoTo SysError
    Select Case KeyAscii
        Case 13
        If Trim$(txtmemberno) <> "" Then
            Editing = True
            Load_Details txtmemberno
            Editing = False
        End If
    End Select
    Exit Sub
SysError:
    MsgBox Err.description, vbInformation, Me.Caption
End Sub

Private Sub Load_Details(memberno As String)
    On Error GoTo SysError
    Dim rsMember As New Recordset
    Set rsMember = oSaccoMaster.GetRecordset("Select MemberNo,OtherNames + ' ' " _
    & "+ SurName as [Names] from MEMBERS where MemberNo='" & txtmemberno & "'")
    With rsMember
        If .State = adStateOpen Then
            If Not .EOF Then
                txtmemberno = memberno
                txtNames = IIf(IsNull(![names]), "", ![names])
                Get_Unposted_Entries memberno
            Else
                txtNames = ""
            End If
        End If
    End With
    Exit Sub
SysError:
    MsgBox Err.description, vbInformation, Me.Caption
End Sub

Private Sub Get_Expected_Deductions(memberno As String)
    On Error GoTo SysError
    Dim rsDeduct As New Recordset, RSPARAM As New Recordset, Shares As Double, _
    LoanNo As String, loan As Double, interest As Double, minshares As Double
    lvwMemberDeductions.ListItems.Clear
    Set RSPARAM = oSaccoMaster.GetRecordset("Select MinTotShares From SYSPARAM")
    With RSPARAM
        If .State = adStateOpen Then
            If Not .EOF Then
                minshares = IIf(IsNull(.Fields(0)), 0, .Fields(0))
            End If
        End If
    End With
    Set rsDeduct = oSaccoMaster.GetRecordset("Select NewContr From SHRVAR " _
    & "where MemberNo='" & memberno & "'")
    With rsDeduct
        If .State = adStateOpen Then
            If Not .EOF Then
                Shares = IIf(IsNull(.Fields(0)), 0, .Fields(0))
            End If
        End If
    End With
    Set li = lvwMemberDeductions.ListItems.Add(, , "Shares Contribution")
    li.SubItems(1) = Format(IIf(Shares > minshares, Shares, minshares), Cfmt)
    li.SubItems(2) = "0.00"
    Set rsDeduct = oSaccoMaster.GetRecordset("Select * From LoanBal where MemberNo" _
    & "='" & memberno & "' and Balance>1")
    With rsDeduct
        If .State = adStateOpen Then
            While Not .EOF
                Set li = lvwMemberDeductions.ListItems.Add(, , "Loan Repayment")
                li.SubItems(1) = Format(IIf(!balance > !repayrate, !repayrate, !balance), Cfmt)
                li.SubItems(2) = Format(Nearest_Five_Cent((!interest / 12 / 100) * !balance), Cfmt)
                li.SubItems(3) = Format(IIf(IsNull(!balance), 0, !balance), Cfmt)
                li.SubItems(4) = IIf(IsNull(!LoanNo), "", !LoanNo)
                li.SubItems(5) = IIf(IsNull(!Loancode), "", !Loancode)
                .MoveNext
            Wend
        End If
    End With
    Get_Totals
    Exit Sub
SysError:
    MsgBox Err.description, vbInformation, Me.Caption
End Sub

Private Sub Get_Totals()
    On Error GoTo SysError
    Dim i As Long
    TotalDed = 0
    If lvwMemberDeductions.ListItems.Count > 0 Then
        For i = 1 To lvwMemberDeductions.ListItems.Count
            If lvwMemberDeductions.ListItems(i).Checked = True Then
                Set li = lvwMemberDeductions.ListItems(i)
                If li.SubItems(3) = "" Then
                    TotalDed = TotalDed + CDbl(li.SubItems(1))
                Else
                    TotalDed = TotalDed + CDbl(li.SubItems(1)) + CDbl(li.SubItems(2))
                End If
            End If
        Next
    End If
    txtExpectedTotals = Format(TotalDed, Cfmt)
    Exit Sub
SysError:
    MsgBox Err.description, vbInformation, Me.Caption
End Sub

Private Sub Get_Unposted_Entries(memberno As String)
    On Error GoTo SysError
    Dim rsCash As New Recordset
    lvwUnposted.ListItems.Clear
    Set rsCash = oSaccoMaster.GetRecordset("Select * From CashBook where " _
    & "MemberNo='" & memberno & "' and Posted=0")
    With rsCash
        If .State = adStateOpen Then
            If Not .EOF Then
                Get_Expected_Deductions memberno
            End If
            While Not .EOF
                Set li = lvwUnposted.ListItems.Add(, , !transdate)
                li.SubItems(1) = IIf(IsNull(!Receiptno), "", !Receiptno)
                li.SubItems(2) = Format(IIf(IsNull(!amount), 0, !amount), Cfmt)
                li.SubItems(3) = IIf(IsNull(!chequeno), "", !chequeno)
                .MoveNext
            Wend
        End If
    End With
    Exit Sub
SysError:
    MsgBox Err.description, vbInformation, Me.Caption
End Sub

Private Sub txtNames_Change()
    On Error GoTo SysError
    Dim rsNames As New Recordset
    lvwNames.ListItems.Clear
    'txtMemberNo = ""
    If Not Editing Then
        If Trim$(txtNames) <> "" Then
            Set rsNames = oSaccoMaster.GetRecordset("Select MemberNo,OtherNames + ' ' + " _
            & "SurName as [Names] From MEMBERS where OtherNames Like '" & txtNames & "%'")
            With rsNames
                If .State = adStateOpen Then
                    While Not .EOF
                        Set li = lvwNames.ListItems.Add(, , IIf(IsNull(![names]), "", ![names]))
                        li.SubItems(1) = IIf(IsNull(!memberno), "", !memberno)
                        .MoveNext
                    Wend
                End If
            End With
            lvwNames.Visible = True
        Else
            lvwNames.Visible = False
        End If
    End If
    Exit Sub
SysError:
    MsgBox Err.description, vbInformation, Me.Caption
    lvwNames.Visible = False
End Sub

Private Sub txtPrincipal_KeyPress(KeyAscii As Integer)
    On Error GoTo SysError
    Select Case KeyAscii
        Case Is = 8
        Case Is = 46
        Case 48 To 57
        Case 13
        If Trim$(txtPrincipal) <> "" Then
            Set li = lvwMemberDeductions.SelectedItem
            li.SubItems(1) = Format(txtPrincipal, Cfmt)
        End If
    End Select
    Get_Totals
    Exit Sub
SysError:
    MsgBox Err.description, vbInformation, Me.Caption
End Sub
