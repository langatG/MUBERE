VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form frmUtilExportToGl 
   Caption         =   "Export Transactions To General Ledger"
   ClientHeight    =   4065
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9000
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmExportToGl.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4065
   ScaleWidth      =   9000
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   1560
      TabIndex        =   40
      Top             =   3600
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4680
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraLoansLedgerAccounts 
      Caption         =   "Ledger Accounts"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2535
      Left            =   4560
      TabIndex        =   13
      Top             =   960
      Width           =   4335
      Begin VB.TextBox txtIntrestDRAccount 
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   1800
         TabIndex        =   17
         Top             =   1560
         Width           =   2415
      End
      Begin VB.TextBox txtLoanDRAccount 
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   1800
         TabIndex        =   16
         Top             =   1200
         Width           =   2415
      End
      Begin VB.TextBox txtIntrestAccount 
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   1800
         TabIndex        =   15
         Top             =   720
         Width           =   2415
      End
      Begin VB.TextBox txtLoansAccount 
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   1800
         TabIndex        =   14
         Top             =   360
         Width           =   2415
      End
      Begin VB.TextBox txtDecPlaces1 
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   18
         Text            =   "2"
         Top             =   2040
         Width           =   735
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   330
         Left            =   2505
         TabIndex        =   19
         Top             =   2040
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Label Label10 
         Caption         =   "Intrest DR Account"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label9 
         Caption         =   "Loan DR Account"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Intrest Account"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Loans Account"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "&Decimal Places"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   2040
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "&Export"
      Height          =   375
      Left            =   120
      TabIndex        =   24
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Frame fraExportOptions 
      Caption         =   "Export Options"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2535
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   4335
      Begin VB.Frame fraSearch 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Visible         =   0   'False
         Width           =   3855
         Begin VB.TextBox txtFromNumber 
            ForeColor       =   &H00FF0000&
            Height          =   330
            Left            =   1680
            TabIndex        =   7
            Top             =   240
            Width           =   1575
         End
         Begin VB.CommandButton cmdFindFrom 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3240
            Picture         =   "frmExportToGl.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox txtToNumber 
            ForeColor       =   &H00FF0000&
            Height          =   330
            Left            =   1680
            TabIndex        =   9
            Top             =   600
            Width           =   1575
         End
         Begin VB.CommandButton cmdFindMemberTo 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3240
            Picture         =   "frmExportToGl.frx":040C
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   600
            Width           =   375
         End
         Begin MSComCtl2.DTPicker dtpFromDate 
            Height          =   330
            Left            =   1680
            TabIndex        =   11
            Top             =   960
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   582
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   69271553
            CurrentDate     =   37799
         End
         Begin MSComCtl2.DTPicker dtpToDate 
            Height          =   330
            Left            =   1680
            TabIndex        =   12
            Top             =   1320
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   582
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   69271553
            CurrentDate     =   37799
         End
         Begin VB.Label lblFromNumber 
            Caption         =   "From Member No."
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label lblToNumber 
            Caption         =   "To Member No."
            Height          =   375
            Left            =   120
            TabIndex        =   31
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label6 
            Caption         =   "From Date"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label7 
            Caption         =   "To Date"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   1320
            Width           =   1335
         End
      End
      Begin VB.CheckBox chkAllTransactions 
         Caption         =   "All Transactions"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame fraSharesLedgerAccounts 
      Caption         =   "Ledger Accounts"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2535
      Left            =   4560
      TabIndex        =   26
      Top             =   960
      Width           =   4335
      Begin VB.TextBox txtDecPlaces2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "2"
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox txtShareAcc 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   1680
         TabIndex        =   20
         Top             =   360
         Width           =   2535
      End
      Begin VB.TextBox txtCreditAcc 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   1680
         TabIndex        =   21
         Top             =   840
         Width           =   2535
      End
      Begin MSComCtl2.UpDown UpDown2 
         Height          =   330
         Left            =   2400
         TabIndex        =   23
         Top             =   1320
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Label Label3 
         Caption         =   "&Decimal Places"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   1350
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Shares Account"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Credit Account"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   840
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   7560
      TabIndex        =   25
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Frame fraExportRecords 
      Caption         =   "Export Records"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin VB.OptionButton optLoansAndInt 
         Caption         =   "Loans And Interest"
         Height          =   255
         Left            =   2520
         TabIndex        =   3
         Top             =   360
         Width           =   1935
      End
      Begin VB.OptionButton optNewLoans 
         Caption         =   "New Loans"
         Height          =   255
         Left            =   1200
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton optShares 
         Caption         =   "Shares"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
   End
   Begin MSComctlLib.ListView lvwACCPAC 
      Height          =   2175
      Left            =   120
      TabIndex        =   39
      Top             =   4200
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   3836
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   16711680
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Account"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Reference"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Description"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Debit"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Credit"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Menu mnuExportFormat 
      Caption         =   "Export Format"
      Visible         =   0   'False
      Begin VB.Menu mnuStandardFormat 
         Caption         =   "Using Standard Export File Format"
      End
      Begin VB.Menu mnuSimplyFormat 
         Caption         =   "Using Simply Accounting File Format"
      End
      Begin VB.Menu mnuAccpacFormat 
         Caption         =   "Export to ACCPAC"
      End
   End
End
Attribute VB_Name = "frmUtilExportToGl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public SourceNo As String
Public SourceInputed As Boolean
Public ACCPACSessionDate As String
Public ACCPACDataBase As String
Public ACCPACUserName As String
Public ACCPACPassword As String
Private searchFrom As Boolean
Private intFileFormat As Byte
Private strCompanyName As String

Private Sub chkAllTransactions_Click()
Me.txtFromNumber.Text = ""
Me.txtToNumber.Text = ""
If Me.chkAllTransactions Then
 Me.fraSearch.Visible = False
Else
 Me.fraSearch.Visible = True
End If
End Sub

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdExport_Click()
'============Validate Text=========================

'============ Validate Share Accounts ===============

    If Me.optShares.Value Then
        If Trim(Me.txtCreditAcc.Text) = "" Or Trim(Me.txtShareAcc.Text) = "" Then
            MsgBox "Please enter accounts to export to", vbInformation, "Account Numbers"
            Me.txtShareAcc.SetFocus
            Exit Sub
        End If

    Else
        If Trim(Me.txtIntrestAccount.Text) = "" Or Trim(Me.txtIntrestDRAccount.Text) = "" Or Trim(Me.txtLoansAccount.Text) = "" Or Me.txtLoanDRAccount.Text = "" Then
            MsgBox "Please enter accounts to export to", vbInformation, "Account Numbers"
            Me.txtLoansAccount.SetFocus
            Exit Sub
        End If
    End If

'============== Validate Transaction Ranges ==================

    If Not Me.chkAllTransactions.Value = 1 Then
        If Trim(Me.txtFromNumber.Text) = "" Or Trim(Me.txtToNumber.Text) = "" Then
            MsgBox "Please enter search range", vbInformation, "Search Range"
            Me.txtFromNumber.SetFocus
        Exit Sub
        End If
    End If
    
'==========End of validate Text==========================


PopupMenu mnuExportFormat
End Sub

Private Sub cmdFindFrom_Click()
Set formCallingSearch = Me
searchFrom = True
onFormLoadOfSearchFrm
End Sub

Public Sub onRefreshOfSearchFrm()
frmSearch.LstSearch.ListItems.Clear
    
If Me.optShares.Value Then
'--------------Load members in form search-----------------------------
    strSql = "Select * from members order by memberno"
    Set rst = oSaccoMaster.GetRecordSet(strSql)
    With rst
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                Set li = frmSearch.LstSearch.ListItems.Add(, , !MemberNo)
                li.SubItems(1) = !staffno & ""
                li.SubItems(2) = !surname & ""
                li.SubItems(3) = !othernames & ""
                li.SubItems(4) = !idNo & ""
                li.SubItems(5) = !employer & ""
                li.SubItems(6) = !CompanyCode & ""
                .MoveNext
            Loop
        End If
        .Close
    End With
    Set rst = Nothing
 frmSearch.cboField.Text = "Member No"
'End If
 
 Else
  '---------------------------Load loan numbers---------------------------
   strSql = "Select * from loans order by loanno"
    Set rst = oSaccoMaster.GetRecordSet(strSql)
    With rst
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                Set li = frmSearch.LstSearch.ListItems.Add(, , !LoanNo & "")
                li.ListSubItems.Add , , !MemberNo & ""
                li.ListSubItems.Add , , !purpose & ""
                li.ListSubItems.Add , , !applicdate & ""
                li.ListSubItems.Add , , !loancode & ""
                li.ListSubItems.Add , , !LoanAmt & ""
                .MoveNext
            Loop
        End If
        .Close
    End With
    Set rst = Nothing
    frmSearch.cboField.Text = "Loan No"
 End If
 '-------------End of records to load---------------------
 
frmSearch.Txtfrom.Text = ""
frmSearch.Txtto.Text = ""
frmSearch.cboCrieria.Text = "="
End Sub
Public Sub onFindOfSearchFrmClick()
    Dim Find As Integer

If Me.optShares.Value Then
'--------------------Load members in search box---------
    Select Case frmSearch.cboField.Text
        Case "Member No"
            searchField = "memberno"
        Case "Surname"
            searchField = "surname"
        Case "Other Names"
            searchField = "othernames"
        Case "ID No"
            searchField = "idno"
        Case "Staff No"
            searchField = "staffno"
        Case "Employer"
            searchField = "employer"
        Case "Company Code"
            searchField = "companycode"
   End Select
   frmSearch.LstSearch.ListItems.Clear
    If Not frmSearch.cboField.Text = "" Then
        If Not frmSearch.cboCrieria.Text = "" Then
            If Not frmSearch.cboCrieria.Text = "Between" And Not frmSearch.cboCrieria.Text = "Like" Then
                strSql = "Select * from members where " & searchField & " " & frmSearch.cboCrieria.Text & " '" & frmSearch.Txtfrom.Text & "'"
                Set rst = oSaccoMaster.GetRecordSet(strSql)
                
                With rst
                    If .RecordCount > 0 Then
                        .MoveFirst
                        Do While Not .EOF
                            Set li = frmSearch.LstSearch.ListItems.Add(, , !MemberNo & "")
                            li.ListSubItems.Add , , !surname & ""
                            li.ListSubItems.Add , , !othernames & ""
                            
                            .MoveNext
                        Loop
                    End If
                End With
                Set rst = Nothing
            ElseIf frmSearch.cboCrieria.Text = "Like" Then
                sql = "Select * from members order by memberno"
                Set rst = oSaccoMaster.GetRecordSet(strSql)
                With rst
                    If .RecordCount > 0 Then
                        .MoveFirst
                        Do While Not .EOF
                            .Find "" & searchField & " " & frmSearch.cboCrieria.Text & " '" & frmSearch.Txtfrom.Text & "%'", , adSearchForward
                            If Not .EOF Then
                                Set li = frmSearch.LstSearch.ListItems.Add(, , !MemberNo & "")
                                li.ListSubItems.Add , , !surname & ""
                                li.ListSubItems.Add , , !othernames & ""
                                .MoveNext
                            End If
                        Loop
                    End If
                End With
                Set rst = Nothing
            Else
                    strSql = "select * from members where " & searchField & " between'" & Txtfrom.Text & "' And '" & Txtto.Text & " '"
                    Set rst = oSaccoMaster.GetRecordSet(strSql)
                With rst
                    If .RecordCount > 0 Then
                        .MoveFirst
                        Do While Not .EOF
                            Set li = frmSearch.LstSearch.ListItems.Add(, , !MemberNo & "")
                            li.SubItems(1) = !staffno & ""
                            li.SubItems(2) = !surname & ""
                            li.SubItems(3) = !othernames & ""
                            li.SubItems(4) = !idNo & ""
                            li.SubItems(5) = !employer & ""
                            li.SubItems(6) = !CompanyCode & ""
                            .MoveNext
                        Loop
                    End If
                End With
                Set rst = Nothing
            End If
        Else
            MsgBox "Select the search criteria.", vbExclamation
        End If
    Else
        MsgBox "Select the search field.", vbExclamation
    End If

 Else
 '----------------------------------Loan search fields------------------
   Select Case frmSearch.cboField.Text
        Case "Loan No"
            searchField = "loanno"
        Case "Member No"
            searchField = "memberno"
        Case "Purpose"
            searchField = "purpose"
        Case "Applic Date"
            searchField = "applicdate"
        Case "Loan Code"
            searchField = "loancode"
        Case "Loan Amount"
            searchField = "loanamt"
    End Select
    frmSearch.LstSearch.ListItems.Clear
    If Not frmSearch.cboField.Text = "" Then
        If Not frmSearch.cboCrieria.Text = "" Then
            If Not frmSearch.cboCrieria.Text = "Between" And Not frmSearch.cboCrieria.Text = "Like" Then
                strSql = "Select * from loans where " & searchField & " " & cboCrieria.Text & " '" & frmSearch.Txtfrom.Text & "'"
                Set rst = oSaccoMaster.GetRecordSet(strSql)
                With rst
                    If .RecordCount > 0 Then
                        .MoveFirst
                        Do While Not .EOF
                            Set li = frmSearch.LstSearch.ListItems.Add(, , !LoanNo & "")
                            li.ListSubItems.Add , , !MemberNo & ""
                            li.ListSubItems.Add , , !purpose & ""
                            li.ListSubItems.Add , , !applicdate & ""
                            li.ListSubItems.Add , , !loancode & ""
                            li.ListSubItems.Add , , !LoanAmt & ""
                            .MoveNext
                        Loop
                    End If
                End With
                Set rst = Nothing
            ElseIf frmSearch.cboCrieria.Text = "Like" Then
                sql = "Select * from loans order by loanno"
                Set rst = oSaccoMaster.GetRecordSet(strSql)
                With rst
                    If .RecordCount > 0 Then
                        .MoveFirst
                        Do While Not .EOF
                            .Find "" & searchField & " " & frmSearch.cboCrieria.Text & " '" & frmSearch.Txtfrom.Text & "%'", , adSearchForward
                            If Not .EOF Then
                                Set li = frmSearch.LstSearch.ListItems.Add(, , !LoanNo & "")
                                li.ListSubItems.Add , , !MemberNo & ""
                                li.ListSubItems.Add , , !purpose & ""
                                li.ListSubItems.Add , , !applicdate & ""
                                li.ListSubItems.Add , , !loancode & ""
                                li.ListSubItems.Add , , !LoanAmt & ""
                                .MoveNext
                            End If
                        Loop
                    End If
                End With
                Set rst = Nothing
            Else
                    strSql = "select * from loans where " & searchField & " between'" & frmSearch.Txtfrom.Text & "' And '" & frmSearch.Txtto.Text & " '"
                    Set rst = oSaccoMaster.GetRecordSet(strSql)
                With rst
                    If .RecordCount > 0 Then
                        .MoveFirst
                        Do While Not .EOF
                            Set li = frmSearch.LstSearch.ListItems.Add(, , !LoanNo & "")
                            li.ListSubItems.Add , , !MemberNo & ""
                            li.ListSubItems.Add , , !purpose & ""
                            li.ListSubItems.Add , , !applicdate & ""
                            li.ListSubItems.Add , , !loancode & ""
                            li.ListSubItems.Add , , !LoanAmt & ""
                            .MoveNext
                        Loop
                    End If
                End With
                Set rst = Nothing
            End If
        Else
            MsgBox "Select the search criteria.", vbExclamation
        End If
    Else
        MsgBox "Select the search field.", vbExclamation
    End If
End If

End Sub
Public Sub onFormLoadOfSearchFrm()
If Me.optShares.Value Then
'--------------------Load members in search box---------
    frmSearch.Caption = "Search Members"
    PositionForm frmSearch
        frmSearch.Show vbModal
        With frmSearch.LstSearch
            .ListItems.Clear
            .ColumnHeaders.Clear
            .ColumnHeaders.Add 1, , "Member No", 2000
            .ColumnHeaders.Add 2, , "Staff No", 2000
            .ColumnHeaders.Add 3, , "Surname", 3000
            .ColumnHeaders.Add 4, , "Other Names", 3000
            .ColumnHeaders.Add 5, , "ID No", 2000
            .ColumnHeaders.Add 6, , "Employer", 3000
            .ColumnHeaders.Add 7, , "Company Code", 2000
            .View = lvwReport
            .Gridlines = True
        End With
        With frmSearch.cboField
            .AddItem ("Member No")
            .AddItem ("Surname")
            .AddItem ("Other Names")
            .AddItem ("Company")
            .AddItem ("Employer")
            .AddItem ("ID No")
            .AddItem ("Staff No")
        End With
        strSql = "Select * from members order by memberno"
        Set rst = oSaccoMaster.GetRecordSet(strSql)
        With rst
            If .RecordCount > 0 Then
                .MoveFirst
                Do While Not .EOF
                    Set li = frmSearch.LstSearch.ListItems.Add(, , !MemberNo)
                    li.SubItems(1) = !staffno & ""
                    li.SubItems(2) = !surname & ""
                    li.SubItems(3) = !othernames & ""
                    li.SubItems(4) = !idNo & ""
                    li.SubItems(5) = !employer & ""
                    li.SubItems(6) = !CompanyCode & ""
                    .MoveNext
                Loop
            End If
            .Close
        End With
        Set rst = Nothing
        frmSearch.cboCrieria.Text = frmSearch.cboCrieria.List(0)
        frmSearch.cboField.Text = frmSearch.cboField.List(0)

Else
 '-----------------------Search Loans----------------------
 frmSearch.Caption = "Search Loans"
    PositionForm frmSearch
    frmSearch.Show vbModal
    With frmSearch.LstSearch
        .ListItems.Clear
        .ColumnHeaders.Clear
        .ColumnHeaders.Add 1, , "Loan No", 2000
        .ColumnHeaders.Add 2, , "Member No", 2000
        .ColumnHeaders.Add 3, , "Purpose", 3000
        .ColumnHeaders.Add 4, , "Applic Date", 3000
        .ColumnHeaders.Add 5, , "Loan Code", 2000
        .ColumnHeaders.Add 6, , "Loan Amount", 3000
        .View = lvwReport
        .Gridlines = True
    End With
 With frmSearch.cboField
    .AddItem ("Loan No")
    .AddItem ("Member No")
    .AddItem ("Purpose")
    .AddItem ("Applic Date")
    .AddItem ("Loan Code")
    .AddItem ("Loan Amount")
 End With
    strSql = "Select * from loans order by loanno"
    Set rst = oSaccoMaster.GetRecordSet(strSql)
    With rst
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                Set li = frmSearch.LstSearch.ListItems.Add(, , !LoanNo & "")
                li.ListSubItems.Add , , !MemberNo & ""
                li.ListSubItems.Add , , !purpose & ""
                li.ListSubItems.Add , , !applicdate & ""
                li.ListSubItems.Add , , !loancode & ""
                li.ListSubItems.Add , , !LoanAmt & ""
                .MoveNext
            Loop
        End If
        .Close
    End With
    Set rst = Nothing
    frmSearch.cboCrieria.Text = frmSearch.cboCrieria.List(0)
    frmSearch.cboField.Text = frmSearch.cboField.List(0)
End If

End Sub
Public Sub searchSelect()
 If searchFrom Then
  formCallingSearch.txtFromNumber.Text = Sel
 Else
  formCallingSearch.txtToNumber.Text = Sel
 End If
   frmSearch.Visible = False
 
  
End Sub

Private Sub cmdFindMemberFrom_Click()

End Sub

Private Sub cmdFindMemberTo_Click()
Set formCallingSearch = Me
searchFrom = False
onFormLoadOfSearchFrm
End Sub

'Private Sub Command1_Click()
'Dim PostAccpac As ACCPAClinks.CAccpac
'Set PostAccpac = New ACCPAClinks.CAccpac
''PostAccpac.PostToAccpac Me.lvwACCPAC, Me.ACCPACUserName, Me.ACCPACPassword, Me.ACCPACDataBase, Format(Me.ACCPACSessionDate, Dfmt), "GL Shares", "Share Contrib", "JE", Format(Date, Dfmt), "06"
'PostAccpac.PostToAccpac Me.lvwACCPAC, "ADMIN", "ADMIN", "SAMLTD", Date, "Batch Desc", "Journal Desc", "JE", "25/06/2003", "06"
'End Sub

Private Sub Form_Load()
Dim rst As ADODB.Recordset
Dim strSql As String

Set formCallingSearch = Me
'------------------Position this form--------------
 PositionForm Me
 
'-------------------Set default startup options----------
Me.chkAllTransactions.Value = 1
Me.optShares.Value = True
Me.fraSearch.Visible = False
Me.fraSharesLedgerAccounts.Visible = True
Me.fraLoansLedgerAccounts.Visible = False
Me.fraSharesLedgerAccounts.Enabled = True
Me.fraLoansLedgerAccounts.Enabled = False
Me.lblFromNumber.Caption = "From Member No."
Me.lblToNumber.Caption = "To Member No."
ProgressBar1.Visible = False

'--------------------Get export settings---------------------

strSql = "SELECT CompanyName FROM SYSPARAM"
Set rst = oSaccoMaster.GetRecordSet(strSql)
If Not rst.EOF Then
strCompanyName = rst!CompanyName & ""
End If
rst.Close

End Sub

'Private Sub mnuAccpacFormat_Click()
'On Error GoTo myTrap
'Dim ItmX As ListItem
'Dim counter As Long
'Dim dblShareSum As Double
'Dim PostAccpac As ACCPAClinks.CAccpac
'SourceInputed = False
'frmExportToAccpac.Show vbModal
'If Not SourceInputed Then
'    Exit Sub
'End If
'
'Me.lvwACCPAC.ListItems.Clear
'
'    If Me.optShares.Value Then
'        '=============== if the Shares option was clicked================
'        '------------------------------if All transactions---------------
'        If Me.chkAllTransactions.Value = 1 Then
'            strSQL = "SELECT TotalShares,MemberNo FROM SHARES ORDER BY TransDate"
'            Set Rst = oSaccoMaster.GetRecordSet(strSQL)
'            If Not Rst.EOF Then
'                Do While Not Rst.EOF
'                    counter = counter + 1
'                    dblShareSum = dblShareSum + FormatNumber(Rst!TotalShares, Me.txtDecPlaces2.Text, , , vbFalse)
'                    Set ItmX = Me.lvwACCPAC.ListItems.Add(, , Me.txtShareAcc.Text)
'                    ItmX.SubItems(1) = Rst!MemberNo & ""
'                    ItmX.SubItems(2) = "Contribution By " & Rst!MemberNo
'                    ItmX.SubItems(3) = "-"
'                    ItmX.SubItems(4) = Rst!TotalShares & ""
'                    Rst.MoveNext
'                Loop
'                Set ItmX = Me.lvwACCPAC.ListItems.Add(, , Me.txtCreditAcc.Text)
'                ItmX.SubItems(1) = "Contributions"
'                ItmX.SubItems(2) = "Total Contributions"
'                ItmX.SubItems(3) = dblShareSum
'                ItmX.SubItems(4) = "-"
'            Else
'                MsgBox "There are no records within the range to export", vbInformation, "No Records"
'            End If
'        '---------------if not all transations----------------------------
'        Else
'            strSQL = "SELECT Amount FROM CONTRIB WHERE (MemberNo BETWEEN '" & Me.txtFromNumber.Text & "' AND '" & Me.txtToNumber.Text & "') AND (ContrDate BETWEEN #" & Me.dtpFromDate.Value & "# AND #" & Me.dtpToDate.Value & "#)"
'            Set Rst = oSaccoMaster.GetRecordSet(strSQL)
'            If Not Rst.EOF Then
'                Do While Not Rst.EOF
'                    counter = counter + 1
'                    dblShareSum = dblShareSum + FormatNumber(Rst!Amount, Me.txtDecPlaces2.Text, , , vbFalse)
'                    Set ItmX = Me.lvwACCPAC.ListItems.Add(, , Me.txtShareAcc.Text)
'                    ItmX.SubItems(1) = Rst!MemberNo & ""
'                    ItmX.SubItems(2) = "Contribution By " & Rst!MemberNo
'                    ItmX.SubItems(3) = "-"
'                    ItmX.SubItems(4) = Rst!Amount & ""
'                    Rst.MoveNext
'                Loop
'                Set ItmX = Me.lvwACCPAC.ListItems.Add(, , Me.txtCreditAcc.Text)
'                ItmX.SubItems(1) = "Contributions"
'                ItmX.SubItems(2) = "Total Contributions"
'                ItmX.SubItems(3) = dblShareSum
'                ItmX.SubItems(4) = "-"
'                Rst.MoveNext
'            Else
'                MsgBox "There are no records within the range to export", vbInformation, "No Records"
'            End If
'        End If
'
'        Set PostAccpac = New ACCPAClinks.CAccpac
'        PostAccpac.PostToAccpac Me.lvwACCPAC, Me.ACCPACUserName, Me.ACCPACPassword, Me.ACCPACDataBase, Format(Me.ACCPACSessionDate, Dfmt), "GL Shares", "Share Contrib", "JE", Format(Date, Dfmt), "06"
'        Rst.Close
'        MsgBox counter & " Transations Exported ", vbInformation, "Exported"
'        '====================End Shares option============================
'    End If
'Exit Sub
'myTrap:
' MsgBox "Export operation failed", vbCritical, "Failed"
'End Sub

Private Sub mnuSimplyFormat_Click()
Dim fso
Dim fs
Dim rst As ADODB.Recordset
Dim counter As Long
Dim dblShareSum As Double
    SourceInputed = False
    frmSourceNo.Show vbModal
    If Not SourceInputed Then
        Exit Sub
    End If
    Me.CommonDialog1.Filter = ".CSV"
    Me.CommonDialog1.ShowSave
    strFileName = CommonDialog1.FileName
    If Trim(strFileName) = "" Then
        Exit Sub
    End If
    strFileName = strFileName & ".CSV"
    counter = 0
    ProgressBar1.Visible = True
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fs = fso.CreateTextFile(strFileName, True)

    If Me.optShares.Value Then
        '=============== if the Shares option was clicked================
        '------------------------------if All transactions---------------
        If Me.chkAllTransactions.Value = 1 Then
            strSql = "SELECT TotalShares FROM SHARES ORDER BY TransDate"
            Set rst = oSaccoMaster.GetRecordSet(strSql)
            
            If Not rst.EOF Then
                fs.WriteLine (Format(Date, "mm-dd-yy") & "," & Me.SourceNo & ",Share Contributions")
                Me.ProgressBar1.max = rst.RecordCount
                
                Do While Not rst.EOF
                    counter = counter + 1
                    dblShareSum = dblShareSum + FormatNumber(rst!TotalShares, Me.txtDecPlaces2.Text, , , vbFalse)
                    fs.WriteLine (Me.txtShareAcc.Text & ",-" & FormatNumber(rst!TotalShares, Me.txtDecPlaces2.Text, , , vbFalse))
                    ProgressBar1.Value = counter
                    rst.MoveNext
                Loop
                fs.WriteLine (Me.txtCreditAcc.Text & "," & FormatNumber(dblShareSum, Me.txtDecPlaces2.Text, , , vbFalse))
            Else
                MsgBox "There are no records within the range to export", vbInformation, "No Records"
            End If
            '---------------if not all transations----------------------------
        Else
            strSql = "SELECT Amount FROM CONTRIB WHERE (MemberNo BETWEEN '" & Me.txtFromNumber.Text & "' AND '" & Me.txtToNumber.Text & "') AND (ContrDate BETWEEN #" & Me.dtpFromDate.Value & "# AND #" & Me.dtpToDate.Value & "#)"
            Set rst = oSaccoMaster.GetRecordSet(strSql)
            If Not rst.EOF Then
                fs.WriteLine (Format(Date, "mm-dd-yy") & "," & Me.SourceNo & ",Share Contributions")
                Me.ProgressBar1.max = rst.RecordCount
                
                Do While Not rst.EOF
                    counter = counter + 1
                    dblShareSum = dblShareSum + FormatNumber(rst!amount, Me.txtDecPlaces2.Text, , , vbFalse)
                    fs.WriteLine (Me.txtShareAcc.Text & ",-" & FormatNumber(rst!amount, Me.txtDecPlaces2.Text, , , vbFalse))
                    ProgressBar1.Value = counter
                    rst.MoveNext
                Loop
                fs.WriteLine (Me.txtCreditAcc.Text & "," & FormatNumber(dblShareSum, Me.txtDecPlaces2.Text, , , vbFalse))
            Else
                MsgBox "There are no records within the range to export", vbInformation, "No Records"
            End If
        End If
        rst.Close
        fs.Close
        ProgressBar1.Visible = False
        MsgBox counter & " Transations Exported ", vbInformation, "Exported"
        '====================End Shares option============================
    End If
End Sub

Private Sub mnuStandardFormat_Click()
Dim fso
Dim fs
Dim rst As ADODB.Recordset
Dim counter As Long
Dim dblShareSum As Double

    Me.CommonDialog1.Filter = ".CSV"
    Me.CommonDialog1.ShowSave
    strFileName = CommonDialog1.FileName
    If Trim(strFileName) = "" Then
        Exit Sub
    End If

    strFileName = strFileName & ".CSV"
    counter = 0
    ProgressBar1.Visible = True
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fs = fso.CreateTextFile(strFileName, True)
    
    If Me.optShares.Value Then
    
        '=============== if the Shares option was clicked================
        
        '=================== if All transactions ========================
        
        If Me.chkAllTransactions.Value = 1 Then
            strSql = "SELECT TotalShares,MemberNo FROM SHARES ORDER BY TransDate"
            Set rst = oSaccoMaster.GetRecordSet(strSql)
            
            If Not rst.EOF Then
                fs.WriteLine (strCompanyName)
                fs.WriteLine ("Member Share Contributions Report Printed on " & Format(Date, Dfmt))
                fs.WriteLine ("ACCOUNTID,TRANSDESC,TRANSAMT")
                Me.ProgressBar1.max = rst.RecordCount
                
                Do While Not rst.EOF
                    counter = counter + 1
                    dblShareSum = dblShareSum + FormatNumber(rst!TotalShares, Me.txtDecPlaces2.Text, , , vbFalse)
                    fs.WriteLine (Me.txtShareAcc.Text & ",Contrib. by Member No. " & rst!MemberNo & ",-" & FormatNumber(rst!TotalShares, Me.txtDecPlaces2.Text, , , vbFalse))
                    ProgressBar1.Value = counter
                    rst.MoveNext
                Loop
                
                fs.WriteLine (Me.txtCreditAcc.Text & ",Total Contributions," & FormatNumber(dblShareSum, Me.txtDecPlaces2.Text, , , vbFalse))
            Else
                MsgBox "There are no records within the range to export", vbInformation, "No Records"
            End If
            
            '========================= if not all transations ==============================
            
        Else
            strSql = "SELECT Amount,MemberNo FROM CONTRIB WHERE (MemberNo BETWEEN '" & Me.txtFromNumber.Text & "' AND '" & Me.txtToNumber.Text & "') AND (ContrDate BETWEEN #" & Me.dtpFromDate.Value & "# AND #" & Me.dtpToDate.Value & "#)"
            Set rst = oSaccoMaster.GetRecordSet(strSql)
            If Not rst.EOF Then
                fs.WriteLine (strCompanyName)
                fs.WriteLine ("Member Share Contributions Report Printed on " & Format(Date, Dfmt))
                fs.WriteLine ("ACCOUNTID,TRANSDESC,TRANSAMT")
                Me.ProgressBar1.max = rst.RecordCount
                
                Do While Not rst.EOF
                    counter = counter + 1
                    dblShareSum = dblShareSum + FormatNumber(rst!amount, Me.txtDecPlaces2.Text, , , vbFalse)
                    fs.WriteLine (Me.txtShareAcc.Text & ",Contrib. by Member No. " & rst!MemberNo & ",-" & FormatNumber(rst!amount, Me.txtDecPlaces2.Text, , , vbFalse))
                    ProgressBar1.Value = counter
                    rst.MoveNext
                Loop
                
                fs.WriteLine (Me.txtCreditAcc.Text & ",Total Contributions," & FormatNumber(dblShareSum, Me.txtDecPlaces2.Text, , , vbFalse))
            Else
                MsgBox "There are no records within the range to export", vbInformation, "No Records"
            End If
        End If
        
        rst.Close
        fs.Close
        ProgressBar1.Visible = False
        MsgBox counter & " Transations Exported ", vbInformation, "Exported"
        '====================End Shares option============================
    End If

End Sub

Private Sub optLoansAndInt_Click()
    Me.fraSharesLedgerAccounts.Visible = False
    Me.fraLoansLedgerAccounts.Visible = True
    Me.fraSharesLedgerAccounts.Enabled = False
    Me.fraLoansLedgerAccounts.Enabled = True
    Me.lblFromNumber.Caption = "From Loan No."
    Me.lblToNumber.Caption = "To Loan No."
    Me.txtFromNumber.Text = ""
    Me.txtToNumber.Text = ""
End Sub

Private Sub optNewLoans_Click()
    Me.fraSharesLedgerAccounts.Visible = False
    Me.fraLoansLedgerAccounts.Visible = True
    Me.fraSharesLedgerAccounts.Enabled = False
    Me.fraLoansLedgerAccounts.Enabled = True
    Me.lblFromNumber.Caption = "From Loan No."
    Me.lblToNumber.Caption = "To Loan No."
    Me.txtFromNumber.Text = ""
    Me.txtToNumber.Text = ""
End Sub

Private Sub optShares_Click()
    Me.fraSharesLedgerAccounts.Visible = True
    Me.fraLoansLedgerAccounts.Visible = False
    Me.fraSharesLedgerAccounts.Enabled = True
    Me.fraLoansLedgerAccounts.Enabled = False
    Me.lblFromNumber.Caption = "From Member No."
    Me.lblToNumber.Caption = "To Member No."
    Me.txtFromNumber.Text = ""
    Me.txtToNumber.Text = ""
End Sub

Private Sub UpDown1_DownClick()
    If CLng(Me.txtDecPlaces1.Text) > 0 Then
        Me.txtDecPlaces1.Text = Me.txtDecPlaces1.Text - 1
    End If
End Sub

Private Sub UpDown1_UpClick()
    If CLng(Me.txtDecPlaces1.Text) < 2 Then
        Me.txtDecPlaces1.Text = Me.txtDecPlaces1.Text + 1
    End If
End Sub

Private Sub UpDown2_DownClick()
    If CLng(Me.txtDecPlaces2.Text) > 0 Then
        Me.txtDecPlaces2.Text = Me.txtDecPlaces2.Text - 1
    End If
End Sub

Private Sub UpDown2_UpClick()
    If CLng(Me.txtDecPlaces2.Text) < 2 Then
        Me.txtDecPlaces2.Text = Me.txtDecPlaces2.Text + 1
    End If
End Sub
