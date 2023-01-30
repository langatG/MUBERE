VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmJournalEntry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Journal Entry"
   ClientHeight    =   9540
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10965
   Icon            =   "frmJournalEntry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9540
   ScaleWidth      =   10965
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7800
      TabIndex        =   77
      Top             =   9000
      Width           =   1335
   End
   Begin VB.CommandButton cmdPost 
      Caption         =   "Post"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5145
      TabIndex        =   56
      Top             =   9015
      Width           =   1335
   End
   Begin VB.CommandButton cmdPostShares 
      Caption         =   "Post"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   55
      Top             =   9000
      Width           =   1335
   End
   Begin VB.Frame grpShares 
      Caption         =   "GL Accounts"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   46
      Top             =   7320
      Width           =   10695
      Begin VB.PictureBox Picture2 
         Height          =   255
         Left            =   3600
         Picture         =   "frmJournalEntry.frx":0442
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   48
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Height          =   255
         Left            =   3600
         Picture         =   "frmJournalEntry.frx":0704
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   47
         Top             =   360
         Width           =   255
      End
      Begin VB.Label lblSharesContra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3960
         TabIndex        =   54
         Top             =   840
         Width           =   3855
      End
      Begin VB.Label Label5 
         Caption         =   "Contra Account"
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Shares Control Account"
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label lblsharesAcc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3960
         TabIndex        =   51
         Top             =   360
         Width           =   3855
      End
      Begin VB.Label lblsharesAccNo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2040
         TabIndex        =   50
         Top             =   345
         Width           =   1575
      End
      Begin VB.Label lblContraAccNo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2025
         TabIndex        =   49
         Top             =   840
         Width           =   1575
      End
   End
   Begin VB.Frame grpLoans 
      Caption         =   "GL Accounts"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   33
      Top             =   7320
      Width           =   10695
      Begin VB.PictureBox Picture14 
         Height          =   255
         Left            =   3600
         Picture         =   "frmJournalEntry.frx":09C6
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   36
         Top             =   240
         Width           =   255
      End
      Begin VB.PictureBox Picture15 
         Height          =   255
         Left            =   3600
         Picture         =   "frmJournalEntry.frx":0C88
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   35
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox Picture16 
         Height          =   255
         Left            =   3600
         Picture         =   "frmJournalEntry.frx":0F4A
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   34
         Top             =   960
         Width           =   255
      End
      Begin VB.Label lblloans 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4320
         TabIndex        =   45
         Top             =   240
         Width           =   3855
      End
      Begin VB.Label Label50 
         Caption         =   "Loan Control Acc"
         Height          =   255
         Left            =   360
         TabIndex        =   44
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblinterst2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4320
         TabIndex        =   43
         Top             =   600
         Width           =   3855
      End
      Begin VB.Label Label53 
         Caption         =   "Interest Control"
         Height          =   255
         Left            =   360
         TabIndex        =   42
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label54 
         Caption         =   "Contra Account"
         Height          =   255
         Left            =   360
         TabIndex        =   41
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label lblContraAccount 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4320
         TabIndex        =   40
         Top             =   960
         Width           =   3855
      End
      Begin VB.Label lblLoanContraAccno 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2040
         TabIndex        =   39
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label lblInterestAccno 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2040
         TabIndex        =   38
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label lblLoanAccNo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2040
         TabIndex        =   37
         Top             =   240
         Width           =   1575
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7095
      Left            =   120
      TabIndex        =   0
      Top             =   105
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   12515
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "SHARES"
      TabPicture(0)   =   "frmJournalEntry.frx":120C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame20"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "grpSharesDest"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "LOANS"
      TabPicture(1)   =   "frmJournalEntry.frx":1228
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame35"
      Tab(1).Control(1)=   "grpLoanDest"
      Tab(1).ControlCount=   2
      Begin VB.Frame grpSharesDest 
         Caption         =   "Destination Member Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   120
         TabIndex        =   67
         Top             =   5280
         Width           =   10335
         Begin VB.TextBox txtSharesDestMemberNo 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   120
            TabIndex        =   72
            Top             =   600
            Width           =   2175
         End
         Begin VB.ComboBox cboSharesLoanNo 
            Height          =   315
            Left            =   5040
            Style           =   2  'Dropdown List
            TabIndex        =   71
            Top             =   1200
            Width           =   1695
         End
         Begin VB.Frame Frame6 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            TabIndex        =   68
            Top             =   960
            Width           =   3735
            Begin VB.OptionButton optSharesDestLoan 
               Caption         =   "Transfer To Loan"
               Height          =   255
               Left            =   120
               TabIndex        =   70
               Top             =   240
               Width           =   1695
            End
            Begin VB.OptionButton optSharesDestShares 
               Caption         =   "Transfer to Shares"
               Height          =   255
               Left            =   1920
               TabIndex        =   69
               Top             =   240
               Width           =   1695
            End
         End
         Begin VB.Label Label6 
            Caption         =   "Full Name"
            Height          =   255
            Left            =   4080
            TabIndex        =   76
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label23 
            Caption         =   "Loan No"
            Height          =   255
            Left            =   4080
            TabIndex        =   75
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label lblSharesFullName 
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000001&
            Height          =   375
            Left            =   5040
            TabIndex        =   74
            Top             =   480
            Width           =   4935
         End
         Begin VB.Label Label21 
            Caption         =   "Memberno"
            Height          =   375
            Left            =   120
            TabIndex        =   73
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame grpLoanDest 
         Caption         =   "Destination Member Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   -74880
         TabIndex        =   57
         Top             =   5160
         Width           =   10335
         Begin VB.TextBox txtLoanDestMemberNo 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1080
            TabIndex        =   62
            Top             =   360
            Width           =   2295
         End
         Begin VB.ComboBox cboLoanLoanNo 
            Height          =   315
            Left            =   4800
            Style           =   2  'Dropdown List
            TabIndex        =   61
            Top             =   960
            Width           =   1695
         End
         Begin VB.Frame Frame5 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            TabIndex        =   58
            Top             =   720
            Width           =   3735
            Begin VB.OptionButton optLoanDestLoan 
               Caption         =   "Transfer To Loan"
               Height          =   255
               Left            =   120
               TabIndex        =   60
               Top             =   240
               Width           =   1695
            End
            Begin VB.OptionButton optLoanDestShares 
               Caption         =   "Transfer to Shares"
               Height          =   255
               Left            =   1920
               TabIndex        =   59
               Top             =   240
               Width           =   1695
            End
         End
         Begin VB.Label Label15 
            Caption         =   "Loan No"
            Height          =   255
            Left            =   3960
            TabIndex        =   66
            Top             =   960
            Width           =   975
         End
         Begin VB.Label lblLoanfullName 
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000001&
            Height          =   375
            Left            =   4800
            TabIndex        =   65
            Top             =   360
            Width           =   4095
         End
         Begin VB.Label Label18 
            Caption         =   "Member No"
            Height          =   255
            Left            =   120
            TabIndex        =   64
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label19 
            Caption         =   "Full Name"
            Height          =   255
            Left            =   3960
            TabIndex        =   63
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame Frame20 
         Height          =   4935
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   10335
         Begin VB.ComboBox cboShareType 
            Height          =   315
            Left            =   3840
            Style           =   2  'Dropdown List
            TabIndex        =   84
            Top             =   960
            Width           =   1695
         End
         Begin VB.TextBox txtSharesTotalAmount 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   82
            Top             =   4560
            Width           =   1575
         End
         Begin VB.TextBox txtInterestAmount 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4080
            TabIndex        =   78
            Top             =   4560
            Width           =   1575
         End
         Begin VB.Frame Frame2 
            Caption         =   "Transaction Type"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            TabIndex        =   27
            Top             =   720
            Width           =   3615
            Begin VB.OptionButton optSharesReversal 
               Caption         =   "Reversal"
               Height          =   255
               Left            =   360
               TabIndex        =   29
               Top             =   240
               Width           =   1575
            End
            Begin VB.OptionButton optSharesTransfer 
               Caption         =   "Transfer"
               Height          =   255
               Left            =   2040
               TabIndex        =   28
               Top             =   240
               Width           =   1215
            End
         End
         Begin VB.TextBox txtSharesVoucherNo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   6240
            TabIndex        =   25
            Top             =   4560
            Width           =   1575
         End
         Begin VB.TextBox txtAmount 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1920
            TabIndex        =   22
            Top             =   4560
            Width           =   1695
         End
         Begin VB.TextBox txtShareMemberNo 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1320
            TabIndex        =   16
            Top             =   240
            Width           =   1935
         End
         Begin MSComCtl2.DTPicker dtpSharesTransDate 
            Height          =   345
            Left            =   5640
            TabIndex        =   19
            Top             =   960
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   609
            _Version        =   393216
            Format          =   122355713
            CurrentDate     =   39863
         End
         Begin MSComctlLib.ListView lvwShareContrib 
            Height          =   2775
            Left            =   120
            TabIndex        =   21
            Top             =   1440
            Width           =   9975
            _ExtentX        =   17595
            _ExtentY        =   4895
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483624
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   0
         End
         Begin VB.Label Label10 
            Caption         =   "Share Type"
            Height          =   255
            Left            =   3960
            TabIndex        =   85
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label9 
            Caption         =   "Total Amount"
            Height          =   255
            Left            =   120
            TabIndex        =   83
            Top             =   4320
            Width           =   1215
         End
         Begin VB.Label Label8 
            Caption         =   "Interest"
            Height          =   255
            Left            =   4080
            TabIndex        =   79
            Top             =   4320
            Width           =   1215
         End
         Begin VB.Label Label3 
            Caption         =   "Voucher No"
            Height          =   255
            Left            =   6360
            TabIndex        =   26
            Top             =   4320
            Width           =   1455
         End
         Begin VB.Label Label7 
            Caption         =   "Amount/Principal"
            Height          =   255
            Left            =   1920
            TabIndex        =   23
            Top             =   4320
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Transaction Date"
            Height          =   255
            Left            =   5640
            TabIndex        =   20
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label Label17 
            Caption         =   "Member No"
            Height          =   255
            Left            =   240
            TabIndex        =   18
            Top             =   240
            Width           =   855
         End
         Begin VB.Label lblNames 
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000001&
            Height          =   375
            Left            =   4200
            TabIndex        =   17
            Top             =   240
            Width           =   4695
         End
      End
      Begin VB.Frame Frame35 
         Height          =   4695
         Left            =   -74880
         TabIndex        =   1
         Top             =   480
         Width           =   10335
         Begin VB.TextBox txtLoansTotalAmount 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   80
            Top             =   4320
            Width           =   1575
         End
         Begin VB.Frame Frame1 
            Caption         =   "Transaction Type"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            TabIndex        =   30
            Top             =   720
            Width           =   3615
            Begin VB.OptionButton optLoanTransfer 
               Caption         =   "Transfer"
               Height          =   255
               Left            =   2040
               TabIndex        =   32
               Top             =   240
               Width           =   1215
            End
            Begin VB.OptionButton optLoanReversal 
               Caption         =   "Reversal"
               Height          =   255
               Left            =   240
               TabIndex        =   31
               Top             =   240
               Width           =   1575
            End
         End
         Begin VB.TextBox txtLoanMemberno 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1080
            TabIndex        =   24
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox txtLoansVoucherNo 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   6960
            TabIndex        =   5
            Top             =   4320
            Width           =   1575
         End
         Begin VB.TextBox txtInterest 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4545
            TabIndex        =   4
            Top             =   4320
            Width           =   1575
         End
         Begin VB.TextBox txtPrincipal 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2280
            TabIndex        =   3
            Top             =   4320
            Width           =   1575
         End
         Begin VB.ComboBox cboLoanno 
            Height          =   315
            Left            =   3960
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   960
            Width           =   1695
         End
         Begin MSComCtl2.DTPicker dtpLoanTransDate 
            Height          =   345
            Left            =   6000
            TabIndex        =   6
            Top             =   960
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   609
            _Version        =   393216
            Format          =   122355713
            CurrentDate     =   39863
         End
         Begin MSComctlLib.ListView lsvLoanTrans 
            Height          =   2535
            Left            =   120
            TabIndex        =   14
            Top             =   1440
            Width           =   9975
            _ExtentX        =   17595
            _ExtentY        =   4471
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483624
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   0
         End
         Begin VB.Label Label2 
            Caption         =   "Total Amount"
            Height          =   255
            Left            =   120
            TabIndex        =   81
            Top             =   4080
            Width           =   1215
         End
         Begin VB.Label Label45 
            Caption         =   "Transaction Date"
            Height          =   255
            Left            =   6000
            TabIndex        =   13
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label Label44 
            Caption         =   "Voucher No"
            Height          =   255
            Left            =   6960
            TabIndex        =   12
            Top             =   4080
            Width           =   1455
         End
         Begin VB.Label Label43 
            Caption         =   "Interest"
            Height          =   255
            Left            =   4920
            TabIndex        =   11
            Top             =   4080
            Width           =   1215
         End
         Begin VB.Label Label42 
            Caption         =   "Principal/Amount"
            Height          =   255
            Left            =   2280
            TabIndex        =   10
            Top             =   4080
            Width           =   1215
         End
         Begin VB.Label Label41 
            Caption         =   "Loan No"
            Height          =   255
            Left            =   4080
            TabIndex        =   9
            Top             =   720
            Width           =   975
         End
         Begin VB.Label lblfullnames 
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000001&
            Height          =   375
            Left            =   3360
            TabIndex        =   8
            Top             =   240
            Width           =   4935
         End
         Begin VB.Label Label39 
            Caption         =   "Member No"
            Height          =   375
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   975
         End
      End
   End
End
Attribute VB_Name = "frmJournalEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NA As String
Dim li As ListItem
'Dim txtPrincipal As Currency
'Dim txtInterest As Currency

Private Sub cboLoanno_Change()
    Dim rsrepay As New ADODB.Recordset
    
    
    mysql = ""
    mysql = "select *  from Repay where  loanno  ='" & cboLoanno & "' order by DateReceived"
    txtLoansTotalAmount = 0
    txtPrincipal = 0
    txtInterest = 0
    Set rsrepay = oSaccoMaster.GetRecordset(mysql)
    
    If Not rsrepay.EOF Then
        lsvLoanTrans.ListItems.Clear
    
        Do While Not rsrepay.EOF
            Set li = lsvLoanTrans.ListItems.Add(, , rsrepay!datereceived & "")
            li.ListSubItems.Add , , rsrepay!principal & ""
            li.ListSubItems.Add , , rsrepay!interest & ""
            li.ListSubItems.Add , , rsrepay!IntrOwed & ""
            li.ListSubItems.Add , , rsrepay!ReceiptNo & ""
            li.ListSubItems.Add , , rsrepay!auditid & ""
            li.ListSubItems.Add , , rsrepay!transby & ""
            rsrepay.MoveNext
        Loop
        Else
        lsvLoanTrans.ListItems.Clear
        End If
End Sub

Private Sub cboLoanno_Click()

cboLoanno_Change
End Sub


Private Sub cboShareType_Click()
Call PopulateList
End Sub

Private Sub cmdclose_Click()
Unload Me

End Sub

Private Sub cmdMatransfer_Click()
''// transfer this money  from the original owner which might be a wrong number  to coreect number  then it be refunded back to rightfull owner
Dim myclass As New cdbase
Dim RsOrig As New ADODB.Recordset
Dim RsTran As New ADODB.Recordset
Dim sql As String

If txtMembernoOverRec.Text = "" Then
    MsgBox "Please select memberno", vbInformation
End If

For I = 1 To lsvOverRecoveryTransfer.ListItems.Count
    If lsvOverRecoveryTransfer.ListItems.Item(I).Checked = True Then
        
        sql = ""
        sql = "Update OVERRECOVERY set memberno ='" & txtMembernoOverRec.Text & "' where memberno ='" & lsvOverRecoveryTransfer.ListItems(I).Text & "'"
        
        oSaccoMaster.ExecuteThis (sql)
    End If
Next I
MsgBox "done"

End Sub

Private Sub cmdNew_Click()
Get_Vno
End Sub

Private Sub cmdOffsetloan_Click()
Dim myclass As New cdbase
Dim I As Integer
Dim Loans As Currency
Dim selection As Boolean
Dim Loanno As String
Dim loanbalance As Currency
Dim rst1 As New ADODB.Recordset
Dim rst2 As New ADODB.Recordset
Dim Offsettingamt As Currency
Dim RefNo As Integer
Dim InteretAmt As Currency
Dim interest As Currency
Dim txtreceiptno As String
Dim rst As New ADODB.Recordset
 selection = False
 If lblAmountOffsetting = "" Then
    MsgBox "No money to off set this loan", vbInformation
    Exit Sub
    
 End If
''//check if loan has been selected
For I = 1 To lsvloans.ListItems.Count
    If lsvloans.ListItems(I).Checked = True Then
    selection = True
    
    Loanno = lsvloans.ListItems.Item(I).Text
    loanbalance = lsvloans.ListItems.Item(I).ListSubItems(2).Text
    Offsettingamt = lblAmountOffsetting
    
    GoTo Continue
    Else
    selection = False
    End If
    
Next I
If selection = False Then
    MsgBox "select loan to off set", vbInformation
    Exit Sub
End If

Continue:
''//continue deduction loan balance
If loanbalance <= 0 Then
    MsgBox "Loan balance is less than zero", vbInformation
    'Exit Sub
End If

Set rst2 = oSaccoMaster.GetRecordset("select * from repay where loanno=" _
    & "'" & Loanno & "' order by datereceived desc,paymentno desc")
    
    
    Set rst1 = oSaccoMaster.GetRecordset("select c.amount,l.* from loanbal " _
    & "l inner join cheques c on l.loanno=c.loanno where l.loanno=" _
    & "'" & Loanno & "'")
    
    If Not rst1.EOF Then
    RepMethod = rst1.Fields("repaymethod")
    '//get interest of the amount
    
    'Transact LoanNo
    interest = lbloffsetinterest
    
    End If
    
    If Not rst2.EOF Then ''//if records are in repay table
        RefNo = rst2!paymentno + 1
        loanbal = rst1!balance
            If loanbal > 0 Then
                If Offsettingamt > loanbal Then
                    
                    Offsettingamt = loanbal - Offsettingamt
                    loanbal = 0
                ElseIf Offsettingamt < loanbal Then
                
                    Offsettingamt = Offsettingamt - interest
                    loanbal = loanbal - Offsettingamt
                ElseIf Offsettingamt = loanbal Then
                loanbal = loanbal - Offsettingamt
                    
                End If
                ''// let put into history what we have achieved
                sql = ""
                    sql = "set dateformat dmy Insert into Repay(LoanNo,MemberNo,DateReceived,PaymentNo,Amount,Principal,Interest" _
                           & ",LoanBalance,ReceiptNo,Locked,Posted,Remarks,AuditID" _
                           & ",AuditTime,Transby)values('" & Loanno & "','" & txtmembernoOff & "','" & Format(dtptransdate1, "dd/MM/yyyy") & "'" _
                           & "," & RefNo & "," & CCur(Offsettingamt) + CCur(interest) & "," & CCur(Offsettingamt) & "," & interest & "" _
                           & "," & loanbal & ",'Over Recovery',0,0,'Over recovery','" & User & "','" & Now & "','Over Recovery')"
                       
                       oSaccoMaster.ExecuteThis (sql)
                       
                       sql = ""
                       sql = "set dateformat dmy Update loanbal set balance =" & loanbal & ",lastdate ='" & Format(dtptransdate1, "dd/MM/yyyy") & "',auditId ='" & User & "' where loanno ='" & Loanno & "'"
                       oSaccoMaster.ExecuteThis sql
                       
                        sql = "Update OVERRECOVERY set posted =1 where memberno ='" & txtmembernoOff & "' and  OverRecoveryid ='" & lbloverrecoverid & "'"
                        oSaccoMaster.ExecuteThis sql
            End If
    Else ''//if records are in repay table
        ''//get loan balance  from loanbal table
        mysql = ""
        mysql = "select * from loanbal  where loanno  ='" & Loanno & "'"
        
        Set rst = oSaccoMaster.GetRecordset(mysql)
        
        If Not rst.EOF Then
        
        
            RefNo = 1
            loanbal = rst!balance
            
            If loanbal > 0 Then
                If Offsettingamt >= loanbal Then
                    Loans = rst!balance
                    Offsettingamt = Offsettingamt - loanbal
                    loanbal = 0
                ElseIf Offsettingamt < loanbal Then
                    Loans = Offsettingamt
                    Offsettingamt = Offsettingamt - interest
                    loanbal = loanbal - Offsettingamt
                      
                End If
                ''// let put into history what we have achieved
                sql = ""
                sql = "set dateformat dmy Insert into Repay(LoanNo,MemberNo,DateReceived,PaymentNo,Amount,Principal,Interest" _
                       & ",LoanBalance,ReceiptNo,Locked,Posted,Remarks,AuditID" _
                       & ",AuditTime,Transby)values('" & Loanno & "','" & txtmembernoOff & "','" & Format(dtptransdate1, "dd/MM/yyyy") & "'" _
                       & "," & RefNo & "," & CCur(Loans) + CCur(interest) & "," & CCur(Loans) & "," & interest & "" _
                       & "," & loanbal & ",'Over Recovery',0,0,'Over recovery','" & User & "','" & Now & "','Over Recovery')"
                       
                       oSaccoMaster.ExecuteThis (sql)
                       
                       sql = ""
                       sql = "set dateformat dmy Update loanbal set balance =" & loanbal & ",lastdate ='" & Format(dtptransdate1, "dd/MM/yyyy") & "',auditId ='" & User & "' where loanno ='" & Loanno & "'"
                       oSaccoMaster.ExecuteThis sql
                       
                        sql = "Update OVERRECOVERY set amount =" & Offsettingamt & " where memberno ='" & txtmembernoOff & "' and  OverRecoveryid ='" & lbloverrecoverid & "'"
                        oSaccoMaster.ExecuteThis sql
            End If
    End If
    End If
    ''//remove that loan from existance
    mysql = "select loanno  from overrecovery  where   OverRecoveryid ='" & lbloverrecoverid & "'"
    
    
    Set rsrepay = oSaccoMaster.GetRecordset(mysql)
    
    If Not rsrepay.EOF Then
    Loanno = rsrepay!Loanno & ""
    End If
    mysql = ""
    mysql = "delete  from overrecovery where  memberno ='" & txtmembernoOff & "' and  OverRecoveryid ='" & lbloverrecoverid & "'"
    oSaccoMaster.ExecuteThis (mysql)
    
    mysql = "delete  from loanbal  where loanno ='" & Loanno & "'"
    
    oSaccoMaster.ExecuteThis (mysql)
    
    mysql = "delete  from cheques  where loanno ='" & Loanno & "'"
    
    oSaccoMaster.ExecuteThis (mysql)
    
    mysql = "delete from repay where loanno ='" & Loanno & "'"
    
        txtreceiptno = "Over Recovery from " & txtmembernoOff
        
        If Label13 <> "" Then '  LOAN ACCOUNT FIRST
                        '// put the gl for the customers
                        sql = ""
                        
                        NA = Label13
                        
                        getde NA
                        
                        sql = "set dateformat dmy insert into customerbalance (customerno,accname,amount,availablebalance,accno,transdescription,transdate,commission,period,transType,posted,locked,status,vno,auditid,auditdate,moduleid,accd) "
                        sql = sql & " values ('" & glmemno & "','" & glnamE1 & "'," & Offsettingamt & "," & bookba + Offsettingamt & ",'" & glaccno & "','" & txtmembernoOff & "','" & Format(DTPTransdate, "dd/mm/yyyy") & "',0,'" & month(DTPTransdate) & "','CR',0,0,0,'" & txtreceiptno & "','" & User & "','" & DTPTransdate & "','3','" & glaccno & "' )"
                        oSaccoMaster.ExecuteThis (sql)
                        
                        sql = ""
                        sql = "set dateformat dmy update cub set amount=" & Offsettingamt & ",transdescription='Over Recovery',availablebalance=" & bookba + Offsettingamt & ",transdate='" & Format(DTPTransdate, "dd/mm/yyyy") & "',vno='" & txtreceiptno & "',period='" & month(DTPTransdate) & "',auditid='" & User & "',auditdate='" & Now & "',moduleid=2 where accno='" & glaccno & "'"
                        
                        oSaccoMaster.ExecuteThis (sql)
            End If
            If Label11 <> "" Then
            
                        sql = ""
                        
                        NA = Label11
                        
                        getde NA
                        
                        sql = "set dateformat dmy insert into customerbalance (customerno,accname,amount,availablebalance,accno,transdescription,transdate,commission,period,transType,posted,locked,status,vno,auditid,auditdate,moduleid,accd) "
                        sql = sql & " values ('" & glmemno & "','" & glnamE1 & "'," & (interest) & "," & bookba + interest & ",'" & glaccno & "','" & txtreceiptno & "','" & Format(DTPTransdate, "dd/mm/yyyy") & "',0,'" & month(DTPTransdate) & "','CR',0,0,0,'" & txtreceiptno & "','" & User & "','" & DTPTransdate & "','3','" & glaccno & "' )"
                        
                       oSaccoMaster.ExecuteThis (sql)
                        
                        sql = ""
                        sql = "set dateformat dmy update cub set amount=" & interest & ",transdescription='" & txtreceiptno & "',availablebalance=" & bookba + interest & ",transdate='" & Format(DTPTransdate, "dd/mm/yyyy") & "',vno='" & txtreceiptno & "',period='" & month(DTPTransdate) & "',auditid='" & User & "',auditdate='" & Now & "',moduleid=2 where accno='" & glaccno & "'"
                        oSaccoMaster.ExecuteThis (sql)
                        
            End If
    
            If lblPettyCash <> "" Then '  interest control
            
                        sql = ""
                
                        '// put the gl to the control account
                        sql = ""
                        NA = lblPettyCash
                         
                        getde NA
                        
                        sql = "set dateformat dmy insert into customerbalance (customerno,accname,amount,availablebalance,accno,transdescription,transdate,commission,period,transType,posted,locked,status,vno,auditid,auditdate,moduleid,accd) "
                        sql = sql & " values ('" & glmemno & "','" & glnamE1 & "'," & CCur(Offsettingamt) + CCur(interest) & "," & bookba + CCur(Offsettingamt) + CCur(interest) & ",'" & glaccno & "','" & txtreceiptno & "','" & Format(DTPTransdate, "dd/mm/yyyy") & "',0,'" & month(DTPTransdate) & "','CR',0,0,0,'" & txtreceiptno & "','" & User & "','" & DTPTransdate & "','3','" & glaccno & "' )"
                        
                        oSaccoMaster.ExecuteThis (sql)
                        sql = ""
                        sql = "set dateformat dmy update cub set amount=" & CCur(Offsettingamt) + CCur(interest) & ",transdescription='" & txtreceiptno & "',availablebalance=" & bookba + txtInterest & ",transdate='" & Format(DTPTransdate, "dd/mm/yyyy") & "',vno='" & txtreceiptno & "',period='" & month(DTPTransdate) & "',auditid='" & User & "',auditdate='" & Now & "',moduleid=2 where accno='" & glaccno & "'"
                       oSaccoMaster.ExecuteThis (sql)
                
                End If
    
    
End Sub

Private Sub cmdPost_Click()
Dim rstRef As New ADODB.Recordset
Dim loanRefbal As Currency
Dim rstref1 As New ADODB.Recordset, Account As LoanGL_Accounts
On Error GoTo SysError

If txtLoanMemberno = "" Then
MsgBox "Enter the MemberNo", vbInformation + vbOKOnly
txtLoanMemberno.SetFocus
Exit Sub
End If

If txtPrincipal = "" Then
MsgBox "Enter the Principal amount to refund", vbInformation + vbOKOnly
txtPrincipal.SetFocus
Exit Sub
End If

If txtInterest = "" Then
MsgBox "Enter the Interest amount to refund", vbInformation + vbOKOnly
txtInterest.SetFocus
Exit Sub
End If

If txtInterest <= 0 And txtPrincipal <= 0 Then
MsgBox "The Principal and Interest amount should not be equal to Zero", vbInformation + vbOKOnly
Exit Sub
End If

If txtLoansVoucherNo = "" Then
MsgBox "Enter the Voucher Number", vbInformation + vbOKOnly
txtLoansVoucherNo.SetFocus
Exit Sub
End If

If txtLoansTotalAmount = "" Or txtLoansTotalAmount < 0 Then
MsgBox "Enter the TotalAmount", vbInformation + vbOKOnly
txtLoansTotalAmount.SetFocus
Exit Sub
End If


    'Check if the Period is closed
    If Check_Period_If_Closed(dtpLoanTransDate) = True Then
        Exit Sub
    End If
    
    If txtLoansTotalAmount <> "" Then
        If CDbl(txtLoansTotalAmount) <> (CDbl(txtInterest) + CDbl(txtPrincipal)) Then
            MsgBox "The Total Amount does not Match the Amount to Transfer/Reverse", vbExclamation, Me.Caption
            Exit Sub
        End If
    End If
    
If optLoanReversal = True Then
        
        If cboLoanno = "" Then
        MsgBox "Select the Loan to Reverse", vbInformation, Me.Caption
        Exit Sub
        End If
        
        If txtLoansTotalAmount = "" Or txtLoansTotalAmount <= 0 Then
        MsgBox "Enter the Total Amount to Reverse", vbInformation, Me.Caption
        txtLoansTotalAmount.SetFocus
        Exit Sub
        End If
        
        If txtPrincipal = "" And txtInterest = "" Then
        MsgBox "Enter the Amount To Reverse(Interest and/or Principal)", vbInformation, Me.Caption
        Exit Sub
        End If
        
        
        If MsgBox("Are you sure you want to Reverse the Transaction?", vbQuestion + vbYesNo) = vbNo Then
            Exit Sub
        End If
               
        
        Set rstRef = oSaccoMaster.GetRecordset("select * from repay where loanno=" _
            & "'" & cboLoanno & "' order by datereceived desc,paymentno desc")
            
        Set rstref1 = oSaccoMaster.GetRecordset("select c.amount,l.* from loanbal " _
            & "l inner join cheques c on l.loanno=c.loanno where l.loanno=" _
            & "'" & cboLoanno & "'")
        
            If Not rstRef.EOF Then
                RefNo = rstRef!paymentno + 1
                loanRefbal = rstref1!balance
                loanRefbal = loanRefbal + txtPrincipal
                          
             End If
             
              Account = Get_Loan_Accounts(cboLoanno, ErrorMessage)
        '//addnew item to repay
        
        sql = ""
        sql = "select top 1 paymentno from repay where loanno='" & cboLoanno & " ' order by datereceived desc, paymentno desc,repayid desc"
        Set rs = oSaccoMaster.GetRecordset(sql)
        If Not rs.EOF Then
        Dim paymentno As Integer
        paymentno = rs.Fields(0)
        End If
        sql = ""
                        sql = "set dateformat dmy Insert into Repay(LoanNo,MemberNo,DateReceived,PaymentNo,Amount,Principal,Interest" _
                               & ",LoanBalance,ReceiptNo,Locked,Posted,Remarks,AuditID" _
                               & ",AuditTime,Transby,LoanAcc,InterestAcc,ContraAcc,cashbookdate,chequeno,dregard)values('" & cboLoanno & "','" & txtLoanMemberno & "','" & Format(dtpLoanTransDate, "dd/MM/yyyy") & "'" _
                               & "," & paymentno + 1 & "," & CCur((txtPrincipal * (-1))) + CCur((txtInterest * (-1))) & "," & CCur((txtPrincipal * (-1))) & "," & CCur((txtInterest * (-1))) & "" _
                               & "," & loanRefbal & ",' Reversal-" & txtLoansVoucherNo & "',0,0,'Loan Reversal" & txtLoansVoucherNo & "','" & User & "','" & Now & "','Loan Reversal" & txtLoansVoucherNo & "','" & Account.LoanAcc & "','" _
                               & Account.interestAcc & "','" & Account.ContraAcc & "','" & Format(dtpLoanTransDate, "dd/MM/yyyy") & "','Reversal" & txtLoansVoucherNo & "',1)"
                               
                               oSaccoMaster.ExecuteThis (sql)
                               
        '//update loanbal
                               sql = ""
                               sql = "set dateformat dmy Update loanbal set balance =" & loanRefbal & ",lastdate ='" & Format(dtpLoanTransDate, "dd/MM/yyyy") & "',auditId ='" & User & "' where loanno ='" & cboLoanno & "'"
                               oSaccoMaster.ExecuteThis sql
                               
                               '// POST IT TO THE GL FOR POSTING LATER
                               Dim DocumentNo As String, TransSource As String, transDescription As String, CashBook As Long, doc_posted As Integer, chequeno As String
                               DocumentNo = "RL" & txtLoansVoucherNo & Format(Time, "hh:mm:ss")
                               transDescription = DocumentNo
                               TransSource = DocumentNo
                               CashBook = 0
                               doc_posted = 0
                               chequeno = DocumentNo
        If Not Save_GLTRANSACTION(Format(dtpLoanTransDate, "dd/mm/yyyy"), CDbl(txtPrincipal), "A007", "L099", DocumentNo, TransSource, User, ErrorMessage, transDescription, CashBook, doc_posted, chequeno, transactionNo) Then
            If ErrorMessage <> "" Then
                MsgBox ErrorMessage, vbInformation, Me.Caption
                ErrorMessage = ""
            End If
        End If
        
                            DocumentNo = "RI" & txtLoansVoucherNo & Format(Time, "hh:mm:ss")
                               transDescription = DocumentNo
                               transource = DocumentNo
                               CashBook = 0
                               doc_posted = 0
                               chequeno = DocumentNo
                If Not Save_GLTRANSACTION(Format(dtpLoanTransDate, "dd/mm/yyyy"), CDbl(txtInterest), "I001", "L099", DocumentNo, TransSource, User, ErrorMessage, transDescription, CashBook, doc_posted, chequeno, TransNo) Then
            If ErrorMessage <> "" Then
                MsgBox ErrorMessage, vbInformation, Me.Caption
                ErrorMessage = ""
            End If
        End If
            
            '//Refresh Loan
            If Not Refresh_Loan(cboLoanno, ErrorMessage) Then
                If ErrorMessage <> "" Then
                    MsgBox ErrorMessage, vbInformation, Me.Caption
                    ErrorMessage = ""
                End If
            End If
                               
        '//update  gl
                    'Principal
                    Select Case CDbl(txtPrincipal)
                    Case Is > 0
                        If Left(cboLoanno, 1) = "I" Then
                            If Not Save_To_GL(Account.interestAcc, Account.ContraAcc, CDbl(txtPrincipal), txtLoansVoucherNo, _
                            txtLoansVoucherNo, dtpLoanTransDate, txtLoanMemberno, lblfullnames.Caption, ErrorMessage, _
                            "Loan Reversal") Then
                                If ErrorMessage <> "" Then
                                    MsgBox ErrorMessage, vbInformation, Me.Caption
                                    ErrorMessage = ""
                                End If
                            End If
                        Else
                              
                            If Not Save_To_GL(Account.LoanAcc, Account.ContraAcc, CDbl(txtPrincipal), txtLoansVoucherNo, _
                            txtLoansVoucherNo, dtpLoanTransDate, txtLoanMemberno, lblfullnames.Caption, ErrorMessage, _
                            "Loan Reversal") Then
                                If ErrorMessage <> "" Then
                                    MsgBox ErrorMessage, vbInformation, Me.Caption
                                    ErrorMessage = ""
                                End If
                            End If
                        End If
                    Case Is < 0
                            If Not Save_To_GL(Account.ContraAcc, Account.LoanAcc, CDbl(txtPrincipal) * (-1), txtLoansVoucherNo, _
                            txtLoansVoucherNo, dtpLoanTransDate, txtLoanMemberno, lblfullnames.Caption, ErrorMessage, _
                            "Loan Reversal") Then
                                If ErrorMessage <> "" Then
                                    MsgBox ErrorMessage, vbInformation, Me.Caption
                                    ErrorMessage = ""
                                End If
                            End If
                    End Select
                    
                'XXXXXXXXXXXXXXX Finish With Interest XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
                Select Case CDbl(txtInterest)
                    Case Is > 0
                        If Not Save_To_GL(Account.interestAcc, Account.ContraAcc, CDbl(txtInterest), txtLoansVoucherNo, _
                        txtLoansVoucherNo, dtpLoanTransDate, txtLoanMemberno, lblfullnames.Caption, ErrorMessage, _
                        "Loan Reversal") Then
                            If ErrorMessage <> "" Then
                                MsgBox ErrorMessage, vbInformation, Me.Caption
                                ErrorMessage = ""
                            End If
                        End If
                    Case Is < 0
                        If Not Save_To_GL(Account.ContraAcc, Account.interestAcc, CDbl(txtInterest) * (-1), txtLoansVoucherNo, _
                        txtLoansVoucherNo, dtpLoanTransDate, txtLoanMemberno, lblfullnames.Caption, ErrorMessage, _
                        "Loan Reversal") Then
                            If ErrorMessage <> "" Then
                                MsgBox ErrorMessage, vbInformation, Me.Caption
                                ErrorMessage = ""
                            End If
                        End If
                End Select
Else 'TRANSFER LOAN TO:
        
    'Case 1: Loan to Shares
    If optLoanDestShares = True Then
    
        If txtLoanDestMemberNo = "" Then
            MsgBox "Enter the MemberNo to Transfer to.", vbExclamation, Me.Caption
            txtLoanDestMemberNo.SetFocus
            Exit Sub
        End If
    
        If lblContraAccNo = "" Then
            MsgBox "Enter the Contra Account.", vbExclamation, Me.Caption
            Exit Sub
        End If
        
        If lblContraAccNo = "" Then
            MsgBox "Enter the Contra Account.", vbExclamation, Me.Caption
            Exit Sub
        End If
    
        If MsgBox("Are you sure you want to make a Transfer to Shares?", vbQuestion + vbYesNo) = vbNo Then
            Exit Sub
        End If

    
        'Update the Repay
        Set rstRef = oSaccoMaster.GetRecordset("select * from repay where loanno=" _
            & "'" & cboLoanno & "' order by datereceived desc,paymentno desc")
            
        Set rstref1 = oSaccoMaster.GetRecordset("select c.amount,l.* from loanbal " _
            & "l inner join cheques c on l.loanno=c.loanno where l.loanno=" _
            & "'" & cboLoanno & "'")
        
            If Not rstRef.EOF Then
                RefNo = rstRef!paymentno + 1
                loanRefbal = rstref1!balance
                loanRefbal = loanRefbal + txtPrincipal
                          
             End If
        '//addnew item to repay
        sql = ""
                        sql = "set dateformat dmy Insert into Repay(LoanNo,MemberNo,DateReceived,PaymentNo,Amount,Principal,Interest" _
                               & ",LoanBalance,ReceiptNo,Locked,Posted,Remarks,AuditID" _
                               & ",AuditTime,Transby,ContraAcc)values('" & cboLoanno & "','" & txtLoanMemberno & "','" & Format(dtpLoanTransDate, "dd/MM/yyyy") & "'" _
                               & "," & txtLoansVoucherNo & "," & CDbl(txtPrincipal) + CDbl(txtInterest) & "," & CDbl(txtPrincipal) * (-1) & "," & CDbl(txtInterest) * (-1) & "" _
                               & "," & loanRefbal & ",'" & txtLoansVoucherNo & "',0,0,'Transfer To " & txtLoanDestMemberNo & "','" & User & "'," _
                               & "'" & Now & "','Transfer To " & txtLoanDestMemberNo & "','" & IIf(lblsharesAccNo <> "", lblsharesAccNo, "") & "')"
                               
                               oSaccoMaster.ExecuteThis (sql)
                               
        '//update loanbal
                        sql = ""
                        sql = "set dateformat dmy Update loanbal set balance =" & loanRefbal & ",lastdate ='" & Format(dtpLoanTransDate, "dd/MM/yyyy") & "',auditId ='" & User & "' where loanno ='" & cboLoanno & "'"
                        oSaccoMaster.ExecuteThis sql
        
        '//Refresh Loan
            If Not Refresh_Loan(cboLoanno, ErrorMessage) Then
                If ErrorMessage <> "" Then
                    MsgBox ErrorMessage, vbInformation, Me.Caption
                    ErrorMessage = ""
                End If
            End If
                                        
        'Update shares of the Destination Member/Account
        Set rstRef = oSaccoMaster.GetRecordset("select * from contrib where memberno=" _
            & "'" & txtLoanDestMemberNo & "' order by contrdate desc,RefNo desc")
            
            If Not rstRef.EOF Then
                RefNo = rstRef!RefNo + 1
                          
            End If
        
        Set rstRef = Nothing
        Set rstRef = oSaccoMaster.GetRecordset("set dateformat dmy insert into contrib(MemberNo,ContrDate," _
        & "RefNo,Amount,ShareBal,TransBy,ChequeNo,ReceiptNo,Locked,Posted,Remarks,AuditID,AuditTime,TransNo,Offset,ContraAcc) " _
        & "values('" & txtLoanDestMemberNo & "','" & dtpLoanTransDate & "'," & RefNo & "," & CDbl(txtLoansTotalAmount) & ", 10000000,'" _
        & txtLoanDestMemberNo & "','" & txtLoansVoucherNo & "','" & txtLoansVoucherNo & "', 'No', 'No', 'Transfer From Loan " & cboLoanno & "','" _
        & User & "','" & Get_Server_Date & "','" & txtShareMemberNo & "',0,'" & IIf(lblContraAccNo <> "", lblContraAccNo, "") & "')")
        
            If Not Save_Audit("Contrib", "Shares Contribution. MemberNo " & txtLoanDestMemberNo, _
    dtpLoanTransDate, txtLoansTotalAmount, User, ErrorMessage) Then
                If errormsg <> "" Then
                    MsgBox errormsg
                    Exit Sub
                End If
         End If
    
    If Not Refresh_Shares(txtLoanDestMemberNo, ErrorMessage) Then
        GoTo SysError
    End If
        
        
'        '//addnew item to  contrib
'             If Not Save_Contrib(txtLoanDestMemberNo, dtpLoanTransDate, 1000, CDbl(txtLoansTotalAmount), _
'                10000000, txtLoanDestMemberNo, txtLoansVoucherNo, txtLoansVoucherNo, "No", "No", "Transfer From Loan " & cboLoanno, _
'                User, "", dtpLoanTransDate, ErrorMessage) Then
'                    If ErrorMessage <> "" Then
'                        MsgBox ErrorMessage, vbInformation, Me.Caption
'                        ErrorMessage = ""
'                        Exit Sub
'                    End If
'                End If
'
        'XXXXXXXXXXXX Update The General Ledger with the Transactions XXXXXXXXXXXXXXX'
                'Principal
                If Not Save_To_GL(lblContraAccNo, lblsharesAccNo, CDbl(txtPrincipal), txtLoanDestMemberNo, txtLoansVoucherNo, _
                dtpLoanTransDate, txtLoanDestMemberNo, "Transfer From Loan " & cboLoanno, ErrorMessage, "Transfer From Loan " & cboLoanno) Then
                    If ErrorMessage <> "" Then
                        If ErrorMessage <> "" Then
                            MsgBox ErrorMessage, vbInformation, Me.Caption
                            ErrorMessage = ""
                        End If
                    End If
                End If
                
                'Interest
            If txtInterest > 0 Then
               If Not Save_To_GL("I001", lblsharesAccNo, CDbl(txtInterest), txtLoanDestMemberNo, txtLoansVoucherNo, _
                dtpLoanTransDate, txtLoanDestMemberNo, "Transfer From Loan " & cboLoanno, ErrorMessage, "Transfer From Loan " & cboLoanno) Then
                    If ErrorMessage <> "" Then
                        If ErrorMessage <> "" Then
                            MsgBox ErrorMessage, vbInformation, Me.Caption
                            ErrorMessage = ""
                        End If
                    End If
                End If
            End If
                
    Else
    'Case 2: Loan to Loan
    'Update the Source Member's Account
                
        If MsgBox("Are you sure you want to make a Transfer to Loan?", vbQuestion + vbYesNo) = vbNo Then
            Exit Sub
        End If
    
        Set rstRef = oSaccoMaster.GetRecordset("select * from repay where loanno=" _
        & "'" & cboLoanno & "' order by datereceived desc,paymentno desc")
            
        Set rstref1 = oSaccoMaster.GetRecordset("select c.amount,l.* from loanbal " _
            & "l inner join cheques c on l.loanno=c.loanno where l.loanno=" _
            & "'" & cboLoanno & "'")
        
            If Not rstRef.EOF Then
                RefNo = rstRef!paymentno + 1
                loanRefbal = rstref1!balance
                loanRefbal = loanRefbal + txtPrincipal
                          
             End If
        '//addnew item to repay
        sql = ""
                        sql = "set dateformat dmy Insert into Repay(LoanNo,MemberNo,DateReceived,PaymentNo,Amount,Principal,Interest" _
                               & ",LoanBalance,ReceiptNo,Locked,Posted,Remarks,AuditID" _
                               & ",AuditTime,Transby)values('" & cboLoanno & "','" & txtLoanMemberno & "','" & Format(dtpLoanTransDate, "dd/MM/yyyy") & "'" _
                               & "," & txtLoansVoucherNo & "," & CDbl(txtPrincipal) + CDbl(txtInterest) & "," & CDbl(txtPrincipal) * (-1) & "," & CDbl(txtInterest) * (-1) & "" _
                               & "," & loanRefbal & ",'" & txtLoansVoucherNo & "',0,0,'Transfer To " & cboLoanLoanNo & "','" & User & "','" & Now & "','Transfer To " & cboLoanLoanNo & "')"
                               
                               oSaccoMaster.ExecuteThis (sql)
                               
        '//update loanbal
                        sql = ""
                        sql = "set dateformat dmy Update loanbal set balance =" & loanRefbal & ",lastdate ='" & Format(dtpLoanTransDate, "dd/MM/yyyy") & "',auditId ='" & User & "' where loanno ='" & cboLoanno & "'"
                        oSaccoMaster.ExecuteThis sql
        '//Refresh Loan
        ErrorMessage = ""
            If Not Refresh_Loan(cboLoanno, ErrorMessage) Then
                If ErrorMessage <> "" Then
                    MsgBox ErrorMessage, vbInformation, Me.Caption
                    ErrorMessage = ""
                End If
            End If
        
        
        'Update the Destination Member's Account
        Set rstRef = Nothing
        Set rstRef = oSaccoMaster.GetRecordset("select * from repay where loanno=" _
        & "'" & cboLoanLoanNo & "' order by datereceived desc,paymentno desc")
            
        Set rstref1 = oSaccoMaster.GetRecordset("select c.amount,l.* from loanbal " _
            & "l inner join cheques c on l.loanno=c.loanno where l.loanno=" _
            & "'" & cboLoanLoanNo & "'")
        
            If Not rstRef.EOF Then
                RefNo = rstRef!paymentno + 1
                loanRefbal = rstref1!balance
                loanRefbal = loanRefbal - txtPrincipal
                          
             End If
        '//addnew item to repay
        sql = ""
                        sql = "set dateformat dmy Insert into Repay(LoanNo,MemberNo,DateReceived,PaymentNo,Amount,Principal,Interest" _
                               & ",LoanBalance,ReceiptNo,Locked,Posted,Remarks,AuditID" _
                               & ",AuditTime,Transby)values('" & cboLoanLoanNo & "','" & txtLoanDestMemberNo & "','" & Format(dtpLoanTransDate, "dd/MM/yyyy") & "'" _
                               & "," & txtLoansVoucherNo & "," & CDbl(txtPrincipal) + CDbl(txtInterest) & "," & CDbl(txtPrincipal) & "," & CDbl(txtInterest) & "" _
                               & "," & loanRefbal & ",'" & txtLoansVoucherNo & "',0,0,'Transfer - " & cboLoanno & "','" & User & "','" & Now & "','Transfer - " & cboLoanno & "')"
                               
                               oSaccoMaster.ExecuteThis (sql)
                               
        '//update loanbal
                        sql = ""
                        sql = "set dateformat dmy Update loanbal set balance =" & loanRefbal & ",lastdate ='" & Format(dtpLoanTransDate, "dd/MM/yyyy") & "',auditId ='" & User & "' where loanno ='" & cboLoanLoanNo & "'"
                        oSaccoMaster.ExecuteThis sql
    
            '//Refresh Loan
            If Not Refresh_Loan(cboLoanLoanNo, ErrorMessage) Then
                If ErrorMessage <> "" Then
                    MsgBox ErrorMessage, vbInformation, Me.Caption
                    ErrorMessage = ""
                End If
            End If
    
    End If
    

End If
If Not Refresh_Loan(cboLoanLoanNo, ErrorMessage) Then
                If ErrorMessage <> "" Then
                    MsgBox ErrorMessage, vbInformation, Me.Caption
                    ErrorMessage = ""
                End If
            End If
    MsgBox "Process complete", vbInformation + vbOKOnly
    cboLoanno_Change
    txtPrincipal = ""
    txtLoansTotalAmount = ""
    txtInterest = ""
    txtLoansVoucherNo = ""
    Exit Sub
SysError:
MsgBox err.description
End Sub

Private Sub cmdRefund_Click()
Dim myclass As New cdbase
Dim Rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rsGl1 As New ADODB.Recordset
Dim rsgl2 As New ADODB.Recordset
Dim Rsaccno As New ADODB.Recordset
Dim cn As New ADODB.Connection
Dim sql As String
Dim amt As Currency

''// open front office and check if then account number exist
''// if not then prompt the user
If lblAmount = "" Then lblAmount = 0
If lblint = "" Then lblint = 0
        If lblLoanNo.Caption = "" Then
            MsgBox "Please very this refund,There is no loan No", vbCritical
            Exit Sub
        End If
        If lblAccNo.Caption = "" Then '// for chepsol members
            MsgBox "There should account number to refund", vbInformation
            Exit Sub
        End If
        If lsvOverRecovery.ListItems.Count < 1 Then  'test of  the listview
            MsgBox "Please  check  Records if exists", vbCritical
            Exit Sub
        End If
        
        If lblMemberNo = "" Then
            MsgBox "Select memberno", vbCritical
            Exit Sub
        End If
        If lblAmount = "" Then
            MsgBox "There is no amount to post", vbInformation
            Exit Sub
        End If
        ''//save  the  data to repay table  for  reporting purpose
        ''/get refno
        Dim RsRefno As New ADODB.Recordset
        Dim RefNo As Integer
        
        mysql = ""
        mysql = "select * from repay where loanno ='" & lblLoanNo & "'"
        
        Set RsRefno = oSaccoMaster.GetRecordset(mysql)
        
        If Not RsRefno.EOF Then
            RefNo = RsRefno!paymentno + 1
            mysql = ""
        
            sql = "set dateformat dmy Insert into Repay(LoanNo,MemberNo,DateReceived,PaymentNo,Amount,Principal,Interest" _
                & ",LoanBalance,ReceiptNo,Locked,Posted,Remarks,AuditID" _
                & ",AuditTime,Transby)values('" & lblLoanNo & "','" & lblMemberNo & "','" & Format(DTPTransdate, "dd/MM/yyyy") & "'" _
                & "," & RefNo & "," & CCur(Format(lblAmount, "############.00")) + CCur(Format(lblint, "########.00")) & "," & CCur(Format(lblAmount, "#########.00")) & "," & CCur(Format(lblint, "##########.00")) & "" _
                & ",0,'" & txtvoucherno & "',0,0,'Over recovery','" & User & "','" & Now & "','Over Recovery')"
                       
                oSaccoMaster.ExecuteThis (sql)
            ''//update loanbal table
            sql = ""
            sql = "Update loanbal set balance =0 where loanno ='" & lblLoanNo & "' and memberno ='" & lblMemberNo & "'"
            oSaccoMaster.ExecuteThis (sql)
        Else
        ''//which is very rare
            RefNo = 1
            mysql = ""
        
            sql = "set dateformat dmy Insert into Repay(LoanNo,MemberNo,DateReceived,PaymentNo,Amount,Principal,Interest" _
                & ",LoanBalance,ReceiptNo,Locked,Posted,Remarks,AuditID" _
                & ",AuditTime,Transby)values('" & lblLoanNo & "','" & lblMemberNo & "','" & Format(DTPTransdate, "dd/MM/yyyy") & "'" _
                & "," & RefNo & "," & CCur(lblAmount) + CCur(lblint) & "," & CCur(lblAmount) & "," & lblint & "" _
                & ",0,'Over Recovery',0,0,'Over recovery','" & User & "','" & Now & "','Over Recovery')"
                       
                oSaccoMaster.ExecuteThis (sql)
                
            ''//update loanbal table
            sql = ""
            sql = "Update loanbal set balance =0 where loanno ='" & lblLoanNo & "' and memberno ='" & lblMemberNo & "'"
            oSaccoMaster.ExecuteThis (sql)
            
            
        End If
        
        
                       
        
        ''// having said and done let us post this monies
        sql = "select * from OverRecovery where memberno ='" & lblMemberNo & "'"
        
        Set Rs1 = oSaccoMaster.GetRecordset(sql)
    
    
    If Not Rs1.EOF Then
        sql = ""
        sql = "set dateformat dmy Update OverRecovery set paid =1 where memberno ='" & lblMemberNo & "'  and refno ='" & lblRefno.Caption & "' and loanno ='" & lblLoanNo & "'"
        myclass.save sql
        
        ''//update gls
    
        sql = ""
        sql = "select * from param"
    
    Set rs2 = oSaccoMaster.GetRecordset(sql)
    
    If Not rs2.EOF Then
        If LCase(rs2!GeneralLegerOpt) = LCase("MAZIWA") Then
            sql = ""
            If lblintcontrol <> "" Then 'test   Over Recovery  control account
                sql = ""
                txtreceiptno = txtvoucherno
            If txtreceiptno = "" Then txtreceiptno = "Over Recovery Refund from accno " & lblAccNo & ""
                sql = ""
                NA = lblintcontrol
             
                getde NA
            
            
                sql = "set dateformat dmy insert into customerbalance (customerno,accname,amount,availablebalance,accno,transdescription,transdate,commission,period,transType,posted,locked,status,vno,auditid,auditdate,moduleid,accd) "
                sql = sql & " values ('" & glmemno & "','" & glnamE1 & "'," & Format(CCur(lblAmount) + CCur(lblint), "#########.00") & "," & bookba + Format(lblAmount, "#########.00") & ",'" & glaccno & "','" & lblname & "','" & Format(DTPTransdate, "dd/mm/yyyy") & "',0,'" & month(DTPTransdate) & "','DR',0,0,0,'" & txtvoucherno & "','" & User & "','" & DTPTransdate & "','3','" & glaccno & "' )"
                
                oSaccoMaster.ExecuteThis (sql)
            
                sql = ""
                sql = "set dateformat dmy update cub set amount=" & Format(lblAmount, "#########.00") & ",transdescription='" & lblname & "',availablebalance=" & bookba + lblAmount & ",transdate='" & Format(DTPTransdate, "dd/mm/yyyy") & "',vno='" & txtreceiptno & "',period='" & month(DTPTransdate) & "',auditid='" & User & "',auditdate='" & Now & "',moduleid=2 where accno='" & glaccno & "'"
                oSaccoMaster.ExecuteThis (sql)
            End If
            
            
            
            If lblloancontrolacc <> "" Then 'test   Over Recovery  control account
    
            
       
                If txtreceiptno = "" Then txtreceiptno = "Over Recovery Refund from accno " & lblAccNo & ""
                    If lblAmount <> 0 Then
                    
                      sql = ""
                      NA = lblloancontrolacc
             
                      getde NA
            
                      sql = "set dateformat dmy insert into customerbalance (customerno,accname,amount,availablebalance,accno,transdescription,transdate,commission,period,transType,posted,locked,status,vno,auditid,auditdate,moduleid,accd) "
                      sql = sql & " values ('" & glmemno & "','" & glnamE1 & "'," & Format(lblAmount, "#########.00") & "," & bookba + Format(lblAmount, "#########.00") & ",'" & glaccno & "','" & lblname & "','" & Format(DTPTransdate, "dd/mm/yyyy") & "',0,'" & month(DTPTransdate) & "','DR',0,0,0,'" & txtvoucherno & "','" & User & "','" & DTPTransdate & "','3','" & glaccno & "' )"
                
                      oSaccoMaster.ExecuteThis (sql)
            
                     sql = ""
                     sql = "set dateformat dmy update cub set amount=" & Format(lblAmount, "#########.00") & ",transdescription='" & lblname & "',availablebalance=" & bookba + lblAmount & ",transdate='" & Format(DTPTransdate, "dd/mm/yyyy") & "',vno='" & txtreceiptno & "',period='" & month(DTPTransdate) & "',auditid='" & User & "',auditdate='" & Now & "',moduleid=2 where accno='" & glaccno & "'"
                     oSaccoMaster.ExecuteThis (sql)
                     
                     End If
                End If
        
            End If 'test   Over Recovery  control account
        
            If lblInterest <> "" Then 'test   Over Recovery  control account
            
                If txtreceiptno = "" Then txtreceiptno = "Over Recovery Refund from accno " & lblAccNo & ""
                    If lblint <> 0 Then
                    
                    sql = ""
                    NA = lblInterest
             
                    getde NA
            
                    sql = "set dateformat dmy insert into customerbalance (customerno,accname,amount,availablebalance,accno,transdescription,transdate,commission,period,transType,posted,locked,status,vno,auditid,auditdate,moduleid,accd) "
                    sql = sql & " values ('" & glmemno & "','" & glnamE1 & "'," & Format(lblint, "#########.00") & "," & bookba + Format(lblAmount, "#########.00") & ",'" & glaccno & "','" & lblname & "','" & Format(DTPTransdate, "dd/mm/yyyy") & "',0,'" & month(DTPTransdate) & "','DR',0,0,0,'" & txtvoucherno & "','" & User & "','" & DTPTransdate & "','3','" & glaccno & "' )"
                
                    oSaccoMaster.ExecuteThis (sql)
            
                    sql = ""
                    sql = "set dateformat dmy update cub set amount=" & Format(lblint, "#########.00") & ",transdescription='" & lblname & "',availablebalance=" & bookba + lblAmount & ",transdate='" & Format(DTPTransdate, "dd/mm/yyyy") & "',vno='" & txtreceiptno & "',period='" & month(DTPTransdate) & "',auditid='" & User & "',auditdate='" & Now & "',moduleid=2 where accno='" & glaccno & "'"
                    oSaccoMaster.ExecuteThis (sql)
                    
                    End If
                    
                End If
            If lblcashcontrol <> "" Then 'test   Over Recovery  control account
                If txtreceiptno = "" Then txtreceiptno = "Over Recovery Refund from accno " & lblAccNo & ""
                    sql = ""
                    NA = lblcashcontrol
                    getde NA
                    
                    sql = "set dateformat dmy insert into customerbalance (customerno,accname,amount,availablebalance,accno,transdescription,transdate,commission,period,transType,posted,locked,status,vno,auditid,auditdate,moduleid,accd) "
                    sql = sql & " values ('" & glmemno & "','" & glnamE1 & "'," & Format(CCur(lblAmount) + CCur(lblint), "#########.00") & "," & bookba + Format(lblAmount, "#########.00") & ",'" & glaccno & "','" & lblname & "','" & Format(DTPTransdate, "dd/mm/yyyy") & "',0,'" & month(DTPTransdate) & "','CR',0,0,0,'" & txtvoucherno & "','" & User & "','" & DTPTransdate & "','3','" & glaccno & "' )"
                
                    oSaccoMaster.ExecuteThis (sql)
            
                    sql = ""
                    sql = "set dateformat dmy update cub set amount=" & Format(lblAmount, "#########.00") & ",transdescription='" & lblname & "',availablebalance=" & bookba + lblAmount & ",transdate='" & Format(DTPTransdate, "dd/mm/yyyy") & "',vno='" & txtreceiptno & "',period='" & month(DTPTransdate) & "',auditid='" & User & "',auditdate='" & Now & "',moduleid=2 where accno='" & glaccno & "'"
                    oSaccoMaster.ExecuteThis (sql)
                End If
            End If 'test   Over Recovery  control account
    
    End If 'test of  the listview
    
    MsgBox "Posted complete", vbInformation, "Succeeded"
    Call LoadData
    
End Sub

Private Sub cmdPostShares_Click()
Dim rstRef As New ADODB.Recordset
Dim shareRefbal As Double
Dim shareBal As Double
Dim shareAcc As String, ContraAcc As String
Dim rstref1 As New ADODB.Recordset
On Error GoTo SysError

If txtShareMemberNo = "" Then
MsgBox "Enter the MemberNo.", vbInformation + vbOKOnly
txtShareMemberNo.SetFocus
Exit Sub
End If

If txtamount = "" Then
MsgBox "Enter the amount to Reverse", vbInformation + vbOKOnly
txtamount.SetFocus
Exit Sub
End If

If txtamount <= 0 Then
MsgBox "Enter the amount to Reverse", vbInformation + vbOKOnly
txtamount.SetFocus
Exit Sub
End If

If txtSharesTotalAmount = "" Then
MsgBox "Enter the Total amount to Reverse", vbInformation + vbOKOnly
txtSharesTotalAmount.SetFocus
Exit Sub
End If

If txtSharesTotalAmount <= 0 Then
MsgBox "The Total amount to Reverse should not be Zero", vbInformation + vbOKOnly
txtSharesTotalAmount.SetFocus
Exit Sub
End If

If txtSharesVoucherNo = "" Then
MsgBox "Enter the Voucher Number", vbInformation + vbOKOnly
txtSharesVoucherNo.SetFocus
Exit Sub
End If

'Check if the Period is closed
If Check_Period_If_Closed(dtpSharesTransDate) = True Then
  Exit Sub
End If
    
If txtSharesTotalAmount <> "" Then
    If CDbl(txtSharesTotalAmount) <> (CDbl(txtInterestAmount) + CDbl(txtamount)) Then
        MsgBox "The Total Amount does not Match the Amount to Transfer/Reverse", vbExclamation, Me.Caption
        Exit Sub
    End If
End If
    
    
 'Post Share Reversal
If optSharesReversal = True Then
        If MsgBox("Are you sure you want to Reverse the Transaction?", vbQuestion + vbYesNo) = vbNo Then
            Exit Sub
        End If
        
        Set rstRef = oSaccoMaster.GetRecordset("select * from contrib where memberno=" _
            & "'" & txtShareMemberNo & "' order by contrdate desc,RefNo desc")
            
            If Not rstRef.EOF Then
                RefNo = rstRef!RefNo + 1
                          
             End If
             
            mysql = ""
    mysql = "select * from sharetype where sharescode ='" & cboShareType & "'"
    
    Set rs = Nothing
    Set rs = oSaccoMaster.GetRecordset(mysql)
    
    If rs.EOF Then
        MsgBox "This scheme code does not exist,Please check again", vbInformation
        Exit Sub
    Else
      shareAcc = rs!SharesAcc
      lblsharesAcc.Caption = shareAcc
      ContraAcc = rs!ContraAcc
      lblSharesContra.Caption = ContraAcc
    End If
    
        '//addnew item to  contrib
             If Not Save_Contrib(txtShareMemberNo, dtpSharesTransDate, 1000, CDbl(txtSharesTotalAmount) * (-1), _
                10000000, txtShareMemberNo, txtSharesVoucherNo, txtSharesVoucherNo, "No", "No", "Share Reversal", _
                User, "", dtpSharesTransDate, ErrorMessage, , cboShareType, shareAcc, ContraAcc, dtpSharesTransDate) Then
                    If ErrorMessage <> "" Then
                        MsgBox ErrorMessage, vbInformation, Me.Caption
                        ErrorMessage = ""
                        Exit Sub
                    End If
                End If
                sql = ""
                sql = "update contrib set dregard=1 where memberno='" & txtShareMemberNo & "' and contrdate='" & dtpSharesTransDate & "' and receiptno='" & txtSharesVoucherNo & "'"
                oSaccoMaster.ExecuteThis (sql)
                'XXXXXXXXXXXX Update The General Ledger with the Transactions XXXXXXXXXXXXXXX'
                If Not Save_To_GL(shareAcc, ContraAcc, CDbl(txtSharesTotalAmount), txtShareMemberNo, txtSharesVoucherNo, _
                dtpSharesTransDate, txtShareMemberNo, "Shares Reversal", ErrorMessage, "Shares Reversal") Then
                    If ErrorMessage <> "" Then
                        If ErrorMessage <> "" Then
                            MsgBox ErrorMessage, vbInformation, Me.Caption
                            Exit Sub
                            ErrorMessage = ""
                        End If
                    End If
                End If
                
                '//put the gl
                Dim DocumentNo As String, TransSource As String, transDescription As String, CashBook As Long, doc_posted As Integer, chequeno As String
                               DocumentNo = "R" & txtShareMemberNo & Format(Time, "hh:mm:ss")
                               transDescription = DocumentNo
                               TransSource = DocumentNo
                               CashBook = 0
                               doc_posted = 0
                               chequeno = DocumentNo
        If Not Save_GLTRANSACTION(Format(dtpSharesTransDate, "dd/mm/yyyy"), CDbl(txtSharesTotalAmount), "L009", "L099", DocumentNo, TransSource, User, ErrorMessage, transDescription, CashBook, doc_posted, chequeno, transactionNo) Then
            If ErrorMessage <> "" Then
                MsgBox ErrorMessage, vbInformation, Me.Caption
                ErrorMessage = ""
            End If
        End If
    
Else 'POST TRANSFER
        'Case 1 : Transfer to Shares - another Member
        
    If optSharesDestShares = True Then
                
        If txtSharesDestMemberNo = "" Then
            MsgBox "Enter the MemberNo to Transfer the Amount To.", vbExclamation, Me.Caption
            txtSharesDestMemberNo.SetFocus
            Exit Sub
        End If
                
                
        If MsgBox("Are you sure you want to make a Transfer to Shares?", vbQuestion + vbYesNo) = vbNo Then
            Exit Sub
        End If
                        
                'Update the SOURCE Member Records
                '//addnew item to  contrib
             If Not Save_Contrib(txtShareMemberNo, dtpSharesTransDate, 1000, CDbl(txtSharesTotalAmount) * (-1), _
                10000000, txtShareMemberNo, txtSharesVoucherNo, txtSharesVoucherNo, "No", "No", "Share Transfer to MemberNo " & txtSharesDestMemberNo, _
                User, "", dtpSharesTransDate, ErrorMessage) Then
                    If ErrorMessage <> "" Then
                        MsgBox ErrorMessage, vbInformation, Me.Caption
                        ErrorMessage = ""
                        Exit Sub
                    End If
                End If
                
                '====================================================================
                ' DESTINATION MEMBER RECORDS
                '=====================================================================
                Set rstRef = Nothing
                Set rstRef = oSaccoMaster.GetRecordset("select * from contrib where memberno=" _
                & "'" & txtSharesDestMemberNo & "' order by contrdate desc,RefNo desc")
                
                If Not rstRef.EOF Then
                    RefNo = rstRef!RefNo + 1
                              
                End If
            
            '//addnew item to  contrib
                 If Not Save_Contrib(txtSharesDestMemberNo, dtpSharesTransDate, 1000, CDbl(txtSharesTotalAmount), _
                    10000000, txtSharesDestMemberNo, txtSharesVoucherNo, txtSharesVoucherNo, "No", "No", "Share Transfer from MemberNo " & txtShareMemberNo, _
                    User, "", dtpSharesTransDate, ErrorMessage) Then
                        If ErrorMessage <> "" Then
                            MsgBox ErrorMessage, vbInformation, Me.Caption
                            ErrorMessage = ""
                            Exit Sub
                        End If
                    End If
            
         'Case 2 : Transfer Shares to Loan for Same or another Member
    Else
        
        If txtInterestAmount = "" Then
        MsgBox "Enter the Interest amount.", vbInformation + vbOKOnly
        txtInterestAmount.SetFocus
        Exit Sub
        End If
             
       
        If txtSharesDestMemberNo = "" Then
            MsgBox "Enter the Member No and Select the Loan to Transfer the Amount To.", vbInformation, Me.Caption
            txtSharesDestMemberNo.SetFocus
            Exit Sub
        End If
        
        If cboSharesLoanNo = "" Then
            MsgBox "Select the Loan to Transfer To.", vbInformation, Me.Caption
            cboSharesLoanNo.SetFocus
            Exit Sub
        End If
    
        If lblLoanAccNo = "" Then
            MsgBox "Enter the Loan GL Account No.", vbInformation, Me.Caption
            Exit Sub
        End If
        
        If lblLoanContraAccno = "" Then
            MsgBox "Enter the Contra Account.", vbInformation, Me.Caption
            Exit Sub
        End If
        
        If lblInterestAccno = "" Then
            MsgBox "Enter the Interest Account.", vbInformation, Me.Caption
            Exit Sub
        End If
        
        If MsgBox("Are you sure you want to make a Transfer to Loan?", vbQuestion + vbYesNo) = vbNo Then
            Exit Sub
        End If
        
        'Update Shares of the Source Member
        Set rstRef = oSaccoMaster.GetRecordset("select * from contrib where memberno=" _
        & "'" & txtShareMemberNo & "' order by contrdate desc,RefNo desc")
            
            If Not rstRef.EOF Then
                RefNo = rstRef!RefNo + 1
                          
            End If
        
        '//addnew item to  contrib
        Set rstRef = Nothing
        Set rstRef = oSaccoMaster.GetRecordset("set dateformat dmy insert into contrib(MemberNo,ContrDate," _
        & "RefNo,Amount,ShareBal,TransBy,ChequeNo,ReceiptNo,Locked,Posted,Remarks,AuditID,AuditTime,TransNo,Offset,ContraAcc) " _
        & "values('" & txtShareMemberNo & "','" & dtpSharesTransDate & "'," & RefNo & "," & CDbl(txtamount) * (-1) & ", 10000000,'" _
        & txtSharesDestMemberNo & "','" & txtSharesVoucherNo & "','" & txtSharesVoucherNo & "', 'No', 'No', 'Transfer to MemberNo " & txtSharesDestMemberNo & "','" _
        & User & "','" & Get_Server_Date & "', '" & txtShareMemberNo & "',0,'" & IIf(lblLoanAccNo <> "", lblLoanAccNo, "") & "')")
        
            If Not Save_Audit("Contrib", "Shares Contribution. MemberNo " & txtShareMemberNo, _
    dtpSharesTransDate, txtSharesTotalAmount, User, ErrorMessage) Then
                If errormsg <> "" Then
                    MsgBox ErrorMessage
                    Exit Sub
                End If
         End If
    If txtInterestAmount <> "" Then
        If txtInterestAmount > 0 Then
                Set rstRef = Nothing
        Set rstRef = oSaccoMaster.GetRecordset("set dateformat dmy insert into contrib(MemberNo,ContrDate," _
        & "RefNo,Amount,ShareBal,TransBy,ChequeNo,ReceiptNo,Locked,Posted,Remarks,AuditID,AuditTime,TransNo,Offset,ContraAcc) " _
        & "values('" & txtShareMemberNo & "','" & dtpSharesTransDate & "'," & RefNo & "," & CDbl(txtInterestAmount) * (-1) & ", 10000000,'" _
        & txtSharesDestMemberNo & "','" & txtSharesVoucherNo & "','" & txtSharesVoucherNo & "', 'No', 'No', 'Transfer to Interest M/No " & txtSharesDestMemberNo & "','" _
        & User & "','" & Get_Server_Date & "', '" & txtShareMemberNo & "',0,'" & IIf(lblInterestAccno <> "", lblInterestAccno, "") & "')")
        
            If Not Save_Audit("Contrib", "Shares Contribution. MemberNo " & txtShareMemberNo, _
    dtpSharesTransDate, txtInterestAmount, User, ErrorMessage) Then
                If errormsg <> "" Then
                    MsgBox ErrorMessage
                    Exit Sub
                End If
         End If
        End If
    End If
    
    If Not Refresh_Shares(txtShareMemberNo, ErrorMessage) Then
        GoTo SysError
    End If

        
'             If Not Save_Contrib(txtShareMemberNo, dtpSharesTransDate, 1000, CDbl(txtSharesTotalAmount) * (-1), _
'                10000000, txtShareMemberNo, txtSharesVoucherNo, txtSharesVoucherNo, "No", "No", "Transfer to MemberNo " & txtSharesDestMemberNo, _
'                User, "", dtpSharesTransDate, ErrorMessage) Then
'                    If ErrorMessage <> "" Then
'                        MsgBox ErrorMessage, vbInformation, Me.Caption
'                        ErrorMessage = ""
'                        Exit Sub
'                    End If
'                End If
'
                'XXXXXXXXXXXX Update The General Ledger with the Transactions XXXXXXXXXXXXXXX'
                If Not Save_To_GL(lblLoanContraAccno, lblContraAccNo, CDbl(txtSharesTotalAmount), txtShareMemberNo, txtSharesVoucherNo, _
                dtpSharesTransDate, txtShareMemberNo, "Transfer to MemberNo " & txtSharesDestMemberNo, ErrorMessage, "Transfer to MemberNo " & txtSharesDestMemberNo) Then
                    If ErrorMessage <> "" Then
                        If ErrorMessage <> "" Then
                            MsgBox ErrorMessage, vbInformation, Me.Caption
                            ErrorMessage = ""
                            Exit Sub
                        End If
                    End If
                End If
        
        'Update Loan of the Destination Member
        Set rstRef = oSaccoMaster.GetRecordset("select * from repay where loanno=" _
        & "'" & cboSharesLoanNo & "' order by datereceived desc,paymentno desc")
        
        Set rstref1 = oSaccoMaster.GetRecordset("select c.amount,l.* from loanbal " _
            & "l inner join cheques c on l.loanno=c.loanno where l.loanno=" _
            & "'" & cboSharesLoanNo & "'")
    
        If Not rstRef.EOF Then
            RefNo = rstRef!paymentno + 1
            loanRefbal = rstref1!balance
            loanRefbal = loanRefbal + txtamount
                      
         End If
         
    '//addnew item to repay
                sql = ""
                sql = "set dateformat dmy Insert into Repay(LoanNo,MemberNo,DateReceived,PaymentNo,Amount,Principal,Interest" _
                & ",LoanBalance,ReceiptNo,Locked,Posted,Remarks,AuditID" _
                & ",AuditTime,Transby,ContraAcc)values('" & cboSharesLoanNo & "','" & txtSharesDestMemberNo & "','" & Format(dtpSharesTransDate, "dd/MM/yyyy") & "'" _
                & "," & txtSharesVoucherNo & "," & CDbl(txtamount) + CDbl(txtInterestAmount) & "," & CDbl(txtamount) & "," & CDbl(txtInterestAmount) & "" _
                & "," & loanRefbal & ",'" & txtSharesVoucherNo & "',0,0,'Transfer From " & txtShareMemberNo & "','" _
                & User & "','" & Now & "','" & txtShareMemberNo & "','" & IIf(lblLoanContraAccno <> "", lblLoanContraAccno, "") & "')"
                       
                oSaccoMaster.ExecuteThis (sql)
                       
'//update loanbal
                sql = ""
                sql = "set dateformat dmy Update loanbal set balance =" & loanRefbal & ",lastdate ='" & Format(dtpSharesTransDate, "dd/MM/yyyy") & "',auditId ='" & User & "' where loanno ='" & cboSharesLoanNo & "'"
                oSaccoMaster.ExecuteThis sql
                
       '//Refresh Loan
            If Not Refresh_Loan(cboSharesLoanNo, ErrorMessage) Then
                If ErrorMessage <> "" Then
                    MsgBox ErrorMessage, vbInformation, Me.Caption
                    ErrorMessage = ""
                    Exit Sub
                End If
            End If
                
'//update  gl
            'Principal
            Select Case CDbl(txtamount)
            Case Is > 0
                If Left(cboSharesLoanNo, 1) = "I" Then
                    If Not Save_To_GL(lblContraAccNo, "I001", CDbl(txtamount), txtSharesVoucherNo, _
                    txtSharesVoucherNo, dtpSharesTransDate, txtShareMemberNo, lblSharesFullName.Caption, ErrorMessage, _
                    "Transfer From MemberNo " & txtShareMemberNo) Then
                        If ErrorMessage <> "" Then
                            MsgBox ErrorMessage, vbInformation, Me.Caption
                            ErrorMessage = ""
                        End If
                    End If
                Else
                      
                    If Not Save_To_GL(lblLoanContraAccno, lblLoanAccNo, CDbl(txtamount), txtSharesVoucherNo, _
                    txtSharesVoucherNo, dtpSharesTransDate, txtShareMemberNo, lblSharesFullName.Caption, ErrorMessage, _
                    "Transfer From MemberNo " & txtShareMemberNo) Then
                        If ErrorMessage <> "" Then
                            MsgBox ErrorMessage, vbInformation, Me.Caption
                            ErrorMessage = ""
                            Exit Sub
                        End If
                    End If
                End If
            Case Is < 0
                    If Not Save_To_GL(lblLoanAccNo, lblLoanContraAccno, CDbl(txtamount) * (-1), txtSharesVoucherNo, _
                    txtSharesVoucherNo, dtpSharesTransDate, txtShareMemberNo, lblSharesFullName.Caption, ErrorMessage, _
                    "Transfer From MemberNo " & txtShareMemberNo) Then
                        If ErrorMessage <> "" Then
                            MsgBox ErrorMessage, vbInformation, Me.Caption
                            ErrorMessage = ""
                            Exit Sub
                        End If
                    End If
            End Select
            
        'XXXXXXXXXXXXXXX Finish With Interest XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
        Select Case CDbl(txtInterestAmount)
            Case Is > 0
                If Not Save_To_GL(lblLoanContraAccno, lblInterestAccno, CDbl(txtInterestAmount), txtSharesVoucherNo, _
                txtSharesVoucherNo, dtpSharesTransDate, txtShareMemberNo, lblSharesFullName.Caption, ErrorMessage, _
                "Transfer From MemberNo " & txtShareMemberNo) Then
                    If ErrorMessage <> "" Then
                        MsgBox ErrorMessage, vbInformation, Me.Caption
                        ErrorMessage = ""
                        Exit Sub
                    End If
                End If
            Case Is < 0
                If Not Save_To_GL(lblInterestAccno, lblContraAccNo, CDbl(txtInterestAmount), txtSharesVoucherNo, _
                txtSharesVoucherNo, dtpSharesTransDate, txtShareMemberNo, lblSharesFullName.Caption, ErrorMessage, _
                "Transfer From MemberNo " & txtShareMemberNo) Then
                    If ErrorMessage <> "" Then
                        MsgBox ErrorMessage, vbInformation, Me.Caption
                        ErrorMessage = ""
                    End If
                End If
        End Select
    End If
        
End If
    
    MsgBox "Process complete", vbInformation + vbOKOnly
    txtShareMemberNo_Change
                sql = ""
                sql = "update contrib set dregard=1 where memberno='" & txtShareMemberNo & "' and contrdate='" & dtpSharesTransDate & "' and receiptno='" & txtSharesVoucherNo & "'"
                oSaccoMaster.ExecuteThis (sql)
                If Not Refresh_Shares(txtShareMemberNo, ErrorMessage) Then
                GoTo SysError
                End If
    txtamount = ""
    txtSharesTotalAmount = ""
    txtSharesVoucherNo = ""
    Exit Sub
SysError:
MsgBox err.description
End Sub

Private Sub Form_Load()
''//get schemes
Dim DTPTransRefDate As Date
Dim RsSchemes As New ADODB.Recordset

dtpSharesTransDate = Format(Get_Server_Date, "dd/MM/yyyy")
dtpLoanTransDate = Format(Get_Server_Date, "dd/MM/yyyy")

mysql = ""
mysql = "select * from sharetype"

Set RsSchemes = oSaccoMaster.GetRecordset(mysql)

If Not RsSchemes.EOF Then
    Do While Not RsSchemes.EOF
        With cboShareType
            .AddItem (RsSchemes!sharesCode & "")
        End With
        RsSchemes.MoveNext
    Loop
Else
    cboShareType.Clear
End If
'optCash_Click
cboShareType.ListIndex = 0
LOADHEADER

If SSTab1.Caption = "SHARES" Then
cmdPostShares.Visible = True
cmdPost.Visible = False
ElseIf SSTab1.Caption = "LOANS" Then
cmdPost.Visible = True
cmdPostShares.Visible = False
End If
'optSharesDestLoan
grpShares.Visible = True
grpLoans.Visible = False
optSharesDestLoan.value = True
optLoanDestShares.value = True
optLoanReversal.value = True
optSharesReversal.value = True
optLoanTransfer.Enabled = False
optSharesTransfer.Enabled = True
txtPrincipal.Enabled = True
txtInterest.Enabled = True
txtamount.Enabled = True
txtInterestAmount.Enabled = True
End Sub

Private Sub LOADHEADER()
With lvwShareContrib
   .ColumnHeaders.Add , , "Date"
   .ColumnHeaders.Add , , "Amount"
   .ColumnHeaders.Add , , "ChequeNo"
   .ColumnHeaders.Add , , "ReceiptNo"
   .ColumnHeaders.Add , , "Trans By"
   .ColumnHeaders.Add , , "Balance"
   .ColumnHeaders.Add , , "AuditId"
   .FullRowSelect = True
   .View = lvwReport
   .GridLines = True
   .LabelEdit = lvwManual
End With

With lsvLoanTrans
    .ColumnHeaders.Add , , "Date Received"
    .ColumnHeaders.Add , , "Principal"
    .ColumnHeaders.Add , , "Interest"
    .ColumnHeaders.Add , , "Interest owed"
    .ColumnHeaders.Add , , "Receiptno"
    .ColumnHeaders.Add , , "AuditID"
    .ColumnHeaders.Add , , "Trans By"
    .FullRowSelect = True
    .View = lvwReport
    .GridLines = True
    .LabelEdit = lvwManual
    
End With

End Sub

Private Sub LoadData()
    Dim li As ListItem
    Dim myclass As New cdbase
    Dim cn As New ADODB.Connection
    Dim Rs1 As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    
    Provider = "MAZIWA"
   cn.Open Provider, "atm", "atm"
    
    Rs1.Open "select * from OverRecovery  where (amount +interest) > 1  and paid  =0 order by memberno asc", cn
    
        lsvOverRecovery.ListItems.Clear
        
    If Not Rs1.EOF Then
        Do While Not Rs1.EOF
            Set li = lsvOverRecovery.ListItems.Add(, , Rs1!memberno)
                li.ListSubItems.Add , , Rs1!ContrDate
                li.ListSubItems.Add , , Format(Rs1!amount, "#########.00")
                li.ListSubItems.Add , , Format(Rs1!interest, "#########.00")
                li.ListSubItems.Add , , Rs1!Loancode & ""
                li.ListSubItems.Add , , Rs1!Remarks & ""
                li.ListSubItems.Add , , Rs1!RefNo & ""
                li.ListSubItems.Add , , Rs1!auditid & ""
                li.ListSubItems.Add , , Rs1!Loanno & ""
                
            Rs1.MoveNext
        Loop
        
    End If
 
End Sub

Private Sub lblmemberno_Change()
Dim myclass As New cdbase
    Dim cn As New ADODB.Connection
    Dim Rs1 As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    
    Provider = "MAZIWA"
   cn.Open Provider, "atm", "atm"
    
    Rs1.Open "select * from members  where memberno ='" & lblMemberNo & "'", cn
        If Not Rs1.EOF Then
            lblname = Rs1!surname & "  " & Rs1!OtherNames & ""
                If Not IsNull(Rs1!ACCNO) Then
                    lblAccNo = Rs1!ACCNO
                End If
        Else
            lblname = ""
            lblAccNo = ""
        End If
End Sub

Private Sub lblmno_Change()
lblmno_Click
End Sub

Private Sub lblmno_Click()
    Dim RsRecords As New ADODB.Recordset
    
    mysql = ""
    mysql = "select * from members where memberno ='" & lblmno & "'"
    
    Set RsRecords = oSaccoMaster.GetRecordset(mysql)
    
    If Not RsRecords.EOF Then
        lblnmes = RsRecords!surname & "" & RsRecords!OtherNames & ""
        lblacc = RsRecords!ACCNO
        
    Else
    End If
    
End Sub

Private Sub lsvo_Click()
    If lsvo.ListItems.Count > 0 Then
        lblamount1.Caption = lsvo.SelectedItem.ListSubItems(2).Text
        lbldate1.Caption = lsvo.SelectedItem.ListSubItems(1).Text
        lblsharesid.Caption = lsvo.SelectedItem.ListSubItems(6).Text
    End If
End Sub
Private Sub lsvOverRecovery_Click()
    If lsvOverRecovery.ListItems.Count > 0 Then
        txtRMemberno = lsvOverRecovery.SelectedItem.Text
        lblMemberNo = lsvOverRecovery.SelectedItem.Text
        DTPTransdate = lsvOverRecovery.SelectedItem.ListSubItems.Item(1).Text
        lblAmount = Format(lsvOverRecovery.SelectedItem.ListSubItems.Item(2).Text, "###,###,###")
        lblMonth = lsvOverRecovery.SelectedItem.ListSubItems.Item(1).Text
        lblRefno = lsvOverRecovery.SelectedItem.ListSubItems.Item(6).Text
        lblint = lsvOverRecovery.SelectedItem.ListSubItems.Item(3).Text
        lblLoanNo = lsvOverRecovery.SelectedItem.ListSubItems.Item(8).Text
        ''//get loancode  from loanno
        ''//given the loancode  you can get description
        lblloancontrolacc = GetLedgerDesc(lsvOverRecovery.SelectedItem.ListSubItems.Item(4).Text)
        lblInterest = GetLedgerDesc("074")
    End If
End Sub
Private Sub lsvOverRecoveryoffsetting_Click()
If lsvOverRecoveryoffsetting.ListItems.Count > 0 Then
       lblAmountOffsetting = lsvOverRecoveryoffsetting.SelectedItem.ListSubItems(2).Text
       lbloffsetinterest = lsvOverRecoveryoffsetting.SelectedItem.ListSubItems(3).Text
       lbloverrecoverid = lsvOverRecoveryoffsetting.SelectedItem.ListSubItems(7).Text
       dtptransdate1.value = lsvOverRecoveryoffsetting.SelectedItem.ListSubItems(1).Text
       
    End If
End Sub



Private Sub Optcash_Click()
lblcashcontrol = "FLOAT AT HAND"
End Sub
Private Sub Optcheque_Click()
lblcashcontrol = "CURRENT A/C CO-OP BANK"
End Sub

Private Sub lsvLoanTrans_Click()
With lsvLoanTrans
    If .ListItems.Count >= 1 Then
        txtPrincipal = .SelectedItem.ListSubItems(1).Text
        txtInterest = .SelectedItem.ListSubItems(2).Text
        txtLoansTotalAmount = CDbl(txtPrincipal) + CDbl(txtInterest)
    End If
End With

End Sub

Private Sub lvwShareContrib_Click()
With lvwShareContrib
    If .ListItems.Count >= 1 Then
        txtamount = .SelectedItem.ListSubItems(1).Text
        txtSharesTotalAmount = .SelectedItem.ListSubItems(1).Text
    End If
End With
txtInterestAmount = "0.00"
End Sub

Private Sub optLoanDestLoan_Click()
If optLoanDestLoan.value = True Then
        grpLoans.Enabled = False
        grpShares.Enabled = False
End If
End Sub

Private Sub optLoanDestShares_Click()
If optLoanDestShares.value = True Then
        grpLoans.Visible = False
        grpShares.Enabled = True
        grpShares.Visible = True
        txtInterest = "0.00"
End If
End Sub

Private Sub optLoanReversal_Click()
If optLoanReversal.value = True Then
grpShares.Visible = False
grpLoans.Visible = True
grpLoanDest.Enabled = False
grpShares.Enabled = False
grpLoans.Enabled = False
End If
End Sub

Private Sub optLoanTransfer_Click()
If optLoanTransfer.value = True Then
grpLoanDest.Enabled = True
grpShares.Enabled = True
grpLoans.Enabled = True
optLoanDestShares_Click
optLoanDestLoan_Click
End If
End Sub

Private Sub optSharesDestLoan_Click()
If optSharesDestLoan.value = True Then
grpLoans.Visible = True
grpLoans.Enabled = True
grpShares.Visible = False
End If
End Sub

Private Sub optSharesDestShares_Click()
If optSharesDestShares.value = True Then
grpShares.Enabled = False
grpLoans.Enabled = False
txtInterestAmount = "0.00"
End If
End Sub

Private Sub optSharesReversal_Click()
If optSharesReversal.value = True Then
grpShares.Visible = True
grpLoans.Visible = False
grpSharesDest.Enabled = False
grpShares.Enabled = False
grpLoans.Enabled = False
End If
End Sub

Private Sub optSharesTransfer_Click()
If optSharesTransfer.value = True Then
grpSharesDest.Enabled = True
grpShares.Enabled = True
grpLoans.Enabled = True
optSharesDestShares_Click
optSharesDestLoan_Click
End If
End Sub

Private Sub Picture1_Click()
Dim Z
Dim rs As New ADODB.Recordset
frmsearchrecords.Show vbModal
 Z = strName
    If Z <> "" Then
     lblintcontrol = Z
        
        End If
        
        Set cn = CreateObject("adodb.connection")
    'If accdr = "" Then Exit Sub
  cn.Open Provider, "atm", "atm"
   
   ''//check where the gls are coming from
   
   rs.Open "select * from param", cn
   
   If Not rs.EOF Then
    If rs!GeneralLegerOpt = "FOSA" Then
        If cn.State = adStateOpen Then
            cn.Close
        End If
            Provider = "dsn_FOSA"
           cn.Open Provider, "atm", "atm"
            sql = ""
            sql = "select * from cub where accno='" & Z & "'"
            Set rs = New ADODB.Recordset
            rs.Open sql, cn, adOpenKeyset, adLockOptimistic
            If Not rs.EOF Then
            If Not IsNull(rs.Fields("name")) Then Label16 = rs.Fields("name")
            If Not IsNull((rs.Fields("ACCNO"))) Then ACCNO1 = rs.Fields("ACCNO")
            If Not IsNull(rs.Fields("name")) Then lblintcontrol = rs.Fields("name")
            
            End If
            
    Else
    If cn.State = adStateOpen Then
            cn.Close
   End If
            Provider = "MAZIWA"
           cn.Open Provider, "atm", "atm"
            
            sql = ""
            sql = "select * from glsetup where glaccname='" & Z & "'"
            Set rs = New ADODB.Recordset
            rs.Open sql, cn, adOpenKeyset, adLockOptimistic
            If Not rs.EOF Then
            If Not IsNull(rs.Fields("glaccname")) Then lblsharesAcc = rs.Fields("glaccname")
            If Not IsNull((rs.Fields("ACCNO"))) Then ACCNO1 = rs.Fields("ACCNO")
            lblsharesAccNo = rs.Fields("ACCNO")
            If Not IsNull((rs.Fields("glaccname"))) Then lblsharesc = rs.Fields("glaccname")
            End If
    End If
    
   End If
   

End Sub

Private Sub Picture14_Click()
Dim Z
Dim rs As New ADODB.Recordset
frmsearchrecords.Show vbModal
 Z = strName
If Z <> "" Then
     lblloans = Z
        
End If
        
        Set cn = CreateObject("adodb.connection")
  cn.Open Provider, "atm", "atm"
   

  If cn.State = adStateOpen Then
            cn.Close
  End If
            Provider = SelectedDsn
           cn.Open Provider, "atm", "atm"
            
            sql = ""
            sql = "select * from glsetup where glaccname='" & Z & "'"
            Set rs = New ADODB.Recordset
            rs.Open sql, cn, adOpenKeyset, adLockOptimistic
            If Not rs.EOF Then
            If Not IsNull(rs.Fields("glaccname")) Then Label16 = rs.Fields("glaccname")
            If Not IsNull((rs.Fields("ACCNO"))) Then ACCNO1 = rs.Fields("ACCNO")
            lblLoanAccNo = rs.Fields("ACCNO")
            End If
End Sub

Private Sub Picture15_Click()
Dim Z
Dim rs As New ADODB.Recordset
frmsearchrecords.Show vbModal
 Z = strName
    If Z <> "" Then
     lblinterst2 = Z
        
        End If
        
        Set cn = CreateObject("adodb.connection")
    'If accdr = "" Then Exit Sub
  cn.Open Provider, "atm", "atm"
   
   ''//check where the gls are coming from
   
   rs.Open "select * from param", cn
   
   If Not rs.EOF Then
    If rs!GeneralLegerOpt = "FOSA" Then
        If cn.State = adStateOpen Then
            cn.Close
        End If
            Provider = "dsn_FOSA"
           cn.Open Provider, "atm", "atm"
            sql = ""
            sql = "select * from cub where accno='" & Z & "'"
            Set rs = New ADODB.Recordset
            rs.Open sql, cn, adOpenKeyset, adLockOptimistic
            If Not rs.EOF Then
            If Not IsNull(rs.Fields("name")) Then Label13 = rs.Fields("name")
            If Not IsNull((rs.Fields("ACCNO"))) Then ACCNO1 = rs.Fields("ACCNO")
            If Not IsNull(rs.Fields("name")) Then Label13 = rs.Fields("name")
            
            End If
            
    Else
    If cn.State = adStateOpen Then
            cn.Close
   End If
            Provider = "MAZIWA"
           cn.Open Provider, "atm", "atm"
            
            sql = ""
            sql = "select * from glsetup where glaccname='" & Z & "'"
            Set rs = New ADODB.Recordset
            rs.Open sql, cn, adOpenKeyset, adLockOptimistic
            If Not rs.EOF Then
            If Not IsNull(rs.Fields("glaccname")) Then Label16 = rs.Fields("glaccname")
            If Not IsNull((rs.Fields("ACCNO"))) Then ACCNO1 = rs.Fields("ACCNO")
            lblInterestAccno = rs.Fields("ACCNO")
            End If
    End If
    
   End If
   
End Sub

Private Sub Picture16_Click()
Dim Z
Dim rs As New ADODB.Recordset
frmsearchrecords.Show vbModal
 Z = strName
    If Z <> "" Then
     lblContraAccount = Z
        
        End If
        
        Set cn = CreateObject("adodb.connection")
    'If accdr = "" Then Exit Sub
  cn.Open Provider, "atm", "atm"
   
   ''//check where the gls are coming from
   
   rs.Open "select * from param", cn
   
   If Not rs.EOF Then
    If rs!GeneralLegerOpt = "FOSA" Then
        If cn.State = adStateOpen Then
            cn.Close
        End If
            Provider = "dsn_FOSA"
           cn.Open Provider, "atm", "atm"
            sql = ""
            sql = "select * from cub where accno='" & Z & "'"
            Set rs = New ADODB.Recordset
            rs.Open sql, cn, adOpenKeyset, adLockOptimistic
            If Not rs.EOF Then
            If Not IsNull(rs.Fields("name")) Then Label13 = rs.Fields("name")
            If Not IsNull((rs.Fields("ACCNO"))) Then ACCNO1 = rs.Fields("ACCNO")
            If Not IsNull(rs.Fields("name")) Then Label13 = rs.Fields("name")
            
            End If
            
    Else
    If cn.State = adStateOpen Then
            cn.Close
   End If
            Provider = "MAZIWA"
           cn.Open Provider, "atm", "atm"
            
            sql = ""
            sql = "select * from glsetup where glaccname='" & Z & "'"
            Set rs = New ADODB.Recordset
            rs.Open sql, cn, adOpenKeyset, adLockOptimistic
            If Not rs.EOF Then
            If Not IsNull(rs.Fields("glaccname")) Then Label16 = rs.Fields("glaccname")
            If Not IsNull((rs.Fields("ACCNO"))) Then ACCNO1 = rs.Fields("ACCNO")
            lblLoanContraAccno = rs.Fields("ACCNO")
            End If
    End If
    
   End If
   
End Sub

Private Sub Picture2_Click()
Dim Z
Dim rs As New ADODB.Recordset
frmsearchrecords.Show vbModal
 Z = strName
    If Z <> "" Then
     lblPettyCash = Z
        
        End If
        
        Set cn = CreateObject("adodb.connection")
    'If accdr = "" Then Exit Sub
  cn.Open Provider, "atm", "atm"
   
   ''//check where the gls are coming from
   
   rs.Open "select * from param", cn
   
   If Not rs.EOF Then
    If rs!GeneralLegerOpt = "FOSA" Then
        If cn.State = adStateOpen Then
            cn.Close
        End If
            Provider = "dsn_FOSA"
           cn.Open Provider, "atm", "atm"
            sql = ""
            sql = "select * from cub where accno='" & Z & "'"
            Set rs = New ADODB.Recordset
            rs.Open sql, cn, adOpenKeyset, adLockOptimistic
            If Not rs.EOF Then
            If Not IsNull(rs.Fields("name")) Then lblPettyCash = rs.Fields("name")
            If Not IsNull((rs.Fields("ACCNO"))) Then ACCNO1 = rs.Fields("ACCNO")
            If Not IsNull(rs.Fields("name")) Then lblPettyCash = rs.Fields("name")
            
            End If
            
    Else
    If cn.State = adStateOpen Then
            cn.Close
   End If
            Provider = "MAZIWA"
           cn.Open Provider, "atm", "atm"
            
            sql = ""
            sql = "select * from glsetup where glaccname='" & Z & "'"
            Set rs = New ADODB.Recordset
            rs.Open sql, cn, adOpenKeyset, adLockOptimistic
            If Not rs.EOF Then
            If Not IsNull(rs.Fields("glaccname")) Then lblSharesContra = rs.Fields("glaccname")
            If Not IsNull((rs.Fields("ACCNO"))) Then ACCNO1 = rs.Fields("ACCNO")
            lblContraAccNo = rs.Fields("ACCNO")
            End If
    End If
    
   End If
   
End Sub

Private Sub Picture3_Click()
Dim Z
Dim rs As New ADODB.Recordset
frmsearchrecords.Show vbModal
 Z = strName
    If Z <> "" Then
     Label11 = Z
        
        End If
        
        Set cn = CreateObject("adodb.connection")
    'If accdr = "" Then Exit Sub
  cn.Open Provider, "atm", "atm"
   
   ''//check where the gls are coming from
   
   rs.Open "select * from param", cn
   
   If Not rs.EOF Then
    If rs!GeneralLegerOpt = "FOSA" Then
        If cn.State = adStateOpen Then
            cn.Close
        End If
            Provider = "dsn_FOSA"
           cn.Open Provider, "atm", "atm"
            sql = ""
            sql = "select * from cub where accno='" & Z & "'"
            Set rs = New ADODB.Recordset
            rs.Open sql, cn, adOpenKeyset, adLockOptimistic
            If Not rs.EOF Then
            If Not IsNull(rs.Fields("name")) Then Label11 = rs.Fields("name")
            If Not IsNull((rs.Fields("ACCNO"))) Then ACCNO1 = rs.Fields("ACCNO")
            If Not IsNull(rs.Fields("name")) Then Label11 = rs.Fields("name")
            
            End If
            
    Else
    If cn.State = adStateOpen Then
            cn.Close
   End If
            Provider = "MAZIWA"
           cn.Open Provider, "atm", "atm"
            
            sql = ""
            sql = "select * from glsetup where glaccname='" & Z & "'"
            Set rs = New ADODB.Recordset
            rs.Open sql, cn, adOpenKeyset, adLockOptimistic
            If Not rs.EOF Then
            If Not IsNull(rs.Fields("glaccname")) Then Label11 = rs.Fields("glaccname")
            If Not IsNull((rs.Fields("ACCNO"))) Then ACCNO1 = rs.Fields("ACCNO")
            End If
    End If
    
   End If
   
End Sub

Private Sub Picture4_Click()
Dim Z
Dim rs As New ADODB.Recordset
frmsearchrecords.Show vbModal
 Z = strName
    If Z <> "" Then
     Label13 = Z
        
        End If
        
        Set cn = CreateObject("adodb.connection")
    'If accdr = "" Then Exit Sub
  cn.Open Provider, "atm", "atm"
   
   ''//check where the gls are coming from
   
   rs.Open "select * from param", cn
   
   If Not rs.EOF Then
    If rs!GeneralLegerOpt = "FOSA" Then
        If cn.State = adStateOpen Then
            cn.Close
        End If
            Provider = "dsn_FOSA"
           cn.Open Provider, "atm", "atm"
            sql = ""
            sql = "select * from cub where accno='" & Z & "'"
            Set rs = New ADODB.Recordset
            rs.Open sql, cn, adOpenKeyset, adLockOptimistic
            If Not rs.EOF Then
            If Not IsNull(rs.Fields("name")) Then Label13 = rs.Fields("name")
            If Not IsNull((rs.Fields("ACCNO"))) Then ACCNO1 = rs.Fields("ACCNO")
            If Not IsNull(rs.Fields("name")) Then Label13 = rs.Fields("name")
            
            End If
            
    Else
    If cn.State = adStateOpen Then
            cn.Close
   End If
            Provider = "MAZIWA"
           cn.Open Provider, "atm", "atm"
            
            sql = ""
            sql = "select * from glsetup where glaccname='" & Z & "'"
            Set rs = New ADODB.Recordset
            rs.Open sql, cn, adOpenKeyset, adLockOptimistic
            If Not rs.EOF Then
            If Not IsNull(rs.Fields("glaccname")) Then Label16 = rs.Fields("glaccname")
            If Not IsNull((rs.Fields("ACCNO"))) Then ACCNO1 = rs.Fields("ACCNO")
            End If
    End If
    
   End If
   
End Sub

Private Sub Picture5_Click()
Dim Z
Dim rs As New ADODB.Recordset
frmsearchrecords.Show vbModal
 Z = strName
    If Z <> "" Then
     lblloancontrolacc = Z
        
        End If
        
        Set cn = CreateObject("adodb.connection")
    'If accdr = "" Then Exit Sub
  cn.Open Provider, "atm", "atm"
   
   ''//check where the gls are coming from
   
   rs.Open "select * from param", cn
   
   If Not rs.EOF Then
    If rs!GeneralLegerOpt = "FOSA" Then
        If cn.State = adStateOpen Then
            cn.Close
        End If
            Provider = "dsn_FOSA"
           cn.Open Provider, "atm", "atm"
            sql = ""
            sql = "select * from cub where accno='" & Z & "'"
            Set rs = New ADODB.Recordset
            rs.Open sql, cn, adOpenKeyset, adLockOptimistic
            If Not rs.EOF Then
            If Not IsNull(rs.Fields("name")) Then Label16 = rs.Fields("name")
            If Not IsNull((rs.Fields("ACCNO"))) Then ACCNO1 = rs.Fields("ACCNO")
            If Not IsNull(rs.Fields("name")) Then lblloancontrolacc = rs.Fields("name")
            
            End If
            
    Else
    If cn.State = adStateOpen Then
            cn.Close
   End If
            Provider = "MAZIWA"
           cn.Open Provider, "atm", "atm"
            
            sql = ""
            sql = "select * from glsetup where glaccname='" & Z & "'"
            Set rs = New ADODB.Recordset
            rs.Open sql, cn, adOpenKeyset, adLockOptimistic
            If Not rs.EOF Then
            If Not IsNull(rs.Fields("glaccname")) Then Label16 = rs.Fields("glaccname")
            If Not IsNull((rs.Fields("ACCNO"))) Then ACCNO1 = rs.Fields("ACCNO")
            End If
    End If
    
   End If
   

End Sub
Private Sub getde(NA As String)
Dim myclass As Object
Dim rs5 As New ADODB.Recordset

 Set myclass = New cdbase
    Set cn = CreateObject("adodb.connection")
    Provider = myclass.OpenCon
   cn.Open Provider, "atm", "atm"
    
    sql = ""
    sql = "select * from param"
    rs5.Open sql, cn
    
    If Not rs5.EOF Then
        If rs5!GeneralLegerOpt = "FOSA" Then
        
    If cn.State = adStateOpen Then
        cn.Close
    End If
    Provider = "dsn_FOSA"
   cn.Open Provider, "atm", "atm"
    
        sql = "select * from cuB where name='" & NA & "'"
        Set rs = New ADODB.Recordset
        
        rs.Open sql, cn
        If Not rs.EOF Then
        If Not IsNull(rs.Fields("accno")) Then glaccno = rs.Fields("accno")
        If Not IsNull(rs.Fields("accountname")) Then glnamE1 = rs.Fields("name")
        If Not IsNull(rs.Fields("idno")) Then glidno = rs.Fields("idno")
        If Not IsNull(rs.Fields("memberno")) Then glmemno = rs.Fields("memberno")
        If Not IsNull(rs.Fields("payno")) Then glpayno = rs.Fields("payno")
        End If
        'glPremium = Scheme_GL_Field(AccountCode, "glPremium")
        'bookba = cub_balance(glaccno)
        

    Else
        sql = "select * from cuB where name='" & NA & "'"
        Set rs = New ADODB.Recordset
        
        rs.Open sql, cn
        If Not rs.EOF Then
        If Not IsNull(rs.Fields("accno")) Then glaccno = rs.Fields("accno")
        If Not IsNull(rs.Fields("accountname")) Then glnamE1 = rs.Fields("name")
        If Not IsNull(rs.Fields("idno")) Then glidno = rs.Fields("idno")
        If Not IsNull(rs.Fields("memberno")) Then glmemno = rs.Fields("memberno")
        If Not IsNull(rs.Fields("payno")) Then glpayno = rs.Fields("payno")
        End If
        'glPremium = Scheme_GL_Field(AccountCode, "glPremium")
        bookba = cub_balance(glaccno)
    
    End If
    End If
    

End Sub
Private Sub getAccno(NA As String)
Dim myclass As Object
Dim rs5 As New ADODB.Recordset

 Set myclass = New cdbase
    Set cn = CreateObject("adodb.connection")
    Provider = myclass.OpenCon
   cn.Open Provider, "atm", "atm"
    
    sql = ""
    sql = "select * from param"
    rs5.Open sql, cn
    
    If Not rs5.EOF Then
        If rs5!GeneralLegerOpt = "FOSA" Then
        
    If cn.State = adStateOpen Then
        cn.Close
    End If
    Provider = "dsn_FOSA"
   cn.Open Provider, "atm", "atm"
    
        sql = "select * from cuB where accno='" & NA & "'"
        Set rs = New ADODB.Recordset
        
        rs.Open sql, cn
        If Not rs.EOF Then
        If Not IsNull(rs.Fields("accno")) Then glaccno = rs.Fields("accno")
        If Not IsNull(rs.Fields("accountname")) Then glnamE1 = rs.Fields("name")
        If Not IsNull(rs.Fields("idno")) Then glidno = rs.Fields("idno")
        If Not IsNull(rs.Fields("memberno")) Then glmemno = rs.Fields("memberno")
        If Not IsNull(rs.Fields("payno")) Then glpayno = rs.Fields("payno")
        End If
        'glPremium = Scheme_GL_Field(AccountCode, "glPremium")
        'bookba = cub_balance(glaccno)
        

    Else
        sql = "select * from cuB where name='" & NA & "'"
        Set rs = New ADODB.Recordset
        
        rs.Open sql, cn
        If Not rs.EOF Then
        If Not IsNull(rs.Fields("accno")) Then glaccno = rs.Fields("accno")
        If Not IsNull(rs.Fields("accountname")) Then glnamE1 = rs.Fields("name")
        If Not IsNull(rs.Fields("idno")) Then glidno = rs.Fields("idno")
        If Not IsNull(rs.Fields("memberno")) Then glmemno = rs.Fields("memberno")
        If Not IsNull(rs.Fields("payno")) Then glpayno = rs.Fields("payno")
        End If
        'glPremium = Scheme_GL_Field(AccountCode, "glPremium")
        bookba = cub_balance(glaccno)
    
    End If
    End If
    

End Sub

Private Sub Picture6_Click()
Dim Z
Dim rs As New ADODB.Recordset
frmsearchrecords.Show vbModal
 Z = strName
    If Z <> "" Then
     Label19 = Z
        
        End If
        
        Set cn = CreateObject("adodb.connection")
    'If accdr = "" Then Exit Sub
  cn.Open Provider, "atm", "atm"
   
   ''//check where the gls are coming from
   
   rs.Open "select * from param", cn
   
   If Not rs.EOF Then
    If rs!GeneralLegerOpt = "FOSA" Then
        If cn.State = adStateOpen Then
            cn.Close
        End If
            Provider = "dsn_FOSA"
           cn.Open Provider, "atm", "atm"
            sql = ""
            sql = "select * from cub where accno='" & Z & "'"
            Set rs = New ADODB.Recordset
            rs.Open sql, cn, adOpenKeyset, adLockOptimistic
            If Not rs.EOF Then
            If Not IsNull(rs.Fields("name")) Then Label19 = rs.Fields("name")
            If Not IsNull((rs.Fields("ACCNO"))) Then ACCNO1 = rs.Fields("ACCNO")
            If Not IsNull(rs.Fields("name")) Then Label19 = rs.Fields("name")
            
            End If
            
    Else
    If cn.State = adStateOpen Then
            cn.Close
   End If
            Provider = "MAZIWA"
           cn.Open Provider, "atm", "atm"
            
            sql = ""
            sql = "select * from glsetup where glaccname='" & Z & "'"
            Set rs = New ADODB.Recordset
            rs.Open sql, cn, adOpenKeyset, adLockOptimistic
            If Not rs.EOF Then
            If Not IsNull(rs.Fields("glaccname")) Then Label19 = rs.Fields("glaccname")
            If Not IsNull((rs.Fields("ACCNO"))) Then ACCNO1 = rs.Fields("ACCNO")
            End If
    End If
    
   End If
   
End Sub

Private Sub Picture7_Click()
Dim Z
Dim rs As New ADODB.Recordset
frmsearchrecords.Show vbModal
 Z = strName
    If Z <> "" Then
     Label13 = Z
        
        End If
        
        Set cn = CreateObject("adodb.connection")
    'If accdr = "" Then Exit Sub
  cn.Open Provider, "atm", "atm"
   
   ''//check where the gls are coming from
   
   rs.Open "select * from param", cn
   
   If Not rs.EOF Then
    If rs!GeneralLegerOpt = "FOSA" Then
        If cn.State = adStateOpen Then
            cn.Close
        End If
            Provider = "dsn_FOSA"
           cn.Open Provider, "atm", "atm"
            sql = ""
            sql = "select * from cub where accno='" & Z & "'"
            Set rs = New ADODB.Recordset
            rs.Open sql, cn, adOpenKeyset, adLockOptimistic
            If Not rs.EOF Then
            If Not IsNull(rs.Fields("name")) Then Label23 = rs.Fields("name")
            If Not IsNull((rs.Fields("ACCNO"))) Then ACCNO1 = rs.Fields("ACCNO")
            If Not IsNull(rs.Fields("name")) Then Label23 = rs.Fields("name")
            
            End If
            
    Else
    If cn.State = adStateOpen Then
            cn.Close
   End If
            Provider = "MAZIWA"
           cn.Open Provider, "atm", "atm"
            
            sql = ""
            sql = "select * from glsetup where glaccname='" & Z & "'"
            Set rs = New ADODB.Recordset
            rs.Open sql, cn, adOpenKeyset, adLockOptimistic
            If Not rs.EOF Then
            If Not IsNull(rs.Fields("glaccname")) Then Label23 = rs.Fields("glaccname")
            If Not IsNull((rs.Fields("ACCNO"))) Then ACCNO1 = rs.Fields("ACCNO")
            End If
    End If
    
   End If
   
End Sub

Private Sub Picture8_Click()
Dim Z
Dim rs As New ADODB.Recordset
frmsearchrecords.Show vbModal
 Z = strName
    If Z <> "" Then
     lblInterest = Z
        
        End If
        
        Set cn = CreateObject("adodb.connection")
    'If accdr = "" Then Exit Sub
  cn.Open Provider, "atm", "atm"
   
   ''//check where the gls are coming from
   
   rs.Open "select * from param", cn
   
   If Not rs.EOF Then
    If rs!GeneralLegerOpt = "FOSA" Then
        If cn.State = adStateOpen Then
            cn.Close
        End If
            Provider = "dsn_FOSA"
           cn.Open Provider, "atm", "atm"
            sql = ""
            sql = "select * from cub where accno='" & Z & "'"
            Set rs = New ADODB.Recordset
            rs.Open sql, cn, adOpenKeyset, adLockOptimistic
            If Not rs.EOF Then
            If Not IsNull(rs.Fields("name")) Then lblInterest = rs.Fields("name")
            If Not IsNull((rs.Fields("ACCNO"))) Then ACCNO1 = rs.Fields("ACCNO")
            If Not IsNull(rs.Fields("name")) Then lblInterest = rs.Fields("name")
            
            End If
            
    Else
    If cn.State = adStateOpen Then
            cn.Close
   End If
            Provider = "MAZIWA"
           cn.Open Provider, "atm", "atm"
            
            sql = ""
            sql = "select * from glsetup where glaccname='" & Z & "'"
            Set rs = New ADODB.Recordset
            rs.Open sql, cn, adOpenKeyset, adLockOptimistic
            If Not rs.EOF Then
            If Not IsNull(rs.Fields("glaccname")) Then lblInterest = rs.Fields("glaccname")
            If Not IsNull((rs.Fields("ACCNO"))) Then ACCNO1 = rs.Fields("ACCNO")
            End If
    End If
    
   End If
   

End Sub

Private Sub Picture9_Click()
Dim Z
Dim rs As New ADODB.Recordset
frmsearchrecords.Show vbModal
 Z = strName
    If Z <> "" Then
     lblcashcontrol = Z
        
        End If
        
        Set cn = CreateObject("adodb.connection")
    'If accdr = "" Then Exit Sub
  cn.Open Provider, "atm", "atm"
   
   ''//check where the gls are coming from
   
   rs.Open "select * from param", cn
   
   If Not rs.EOF Then
    If rs!GeneralLegerOpt = "FOSA" Then
        If cn.State = adStateOpen Then
            cn.Close
        End If
            Provider = "dsn_FOSA"
           cn.Open Provider, "atm", "atm"
            sql = ""
            sql = "select * from cub where accno='" & Z & "'"
            Set rs = New ADODB.Recordset
            rs.Open sql, cn, adOpenKeyset, adLockOptimistic
            If Not rs.EOF Then
            If Not IsNull(rs.Fields("name")) Then lblcashcontrol = rs.Fields("name")
            If Not IsNull((rs.Fields("ACCNO"))) Then ACCNO1 = rs.Fields("ACCNO")
            If Not IsNull(rs.Fields("name")) Then lblcashcontrol = rs.Fields("name")
            
            End If
            
    Else
    If cn.State = adStateOpen Then
            cn.Close
   End If
            Provider = "MAZIWA"
           cn.Open Provider, "atm", "atm"
            
            sql = ""
            sql = "select * from glsetup where glaccname='" & Z & "'"
            Set rs = New ADODB.Recordset
            rs.Open sql, cn, adOpenKeyset, adLockOptimistic
            If Not rs.EOF Then
            If Not IsNull(rs.Fields("glaccname")) Then lblcashcontrol = rs.Fields("glaccname")
            If Not IsNull((rs.Fields("ACCNO"))) Then ACCNO1 = rs.Fields("ACCNO")
            End If
    End If
    
   End If
   
End Sub

Private Sub txtMemberNo_Change()
Dim myclass As New cdbase
    Dim cn As New ADODB.Connection
    Dim Rs1 As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    
    Provider = "MAZIWA"
   cn.Open Provider, "atm", "atm"
    
    Rs1.Open "select * from members  where memberno ='" & txtMemberNo & "'", cn
        If Not Rs1.EOF Then
            lblfullnames = Rs1!surname & "  " & Rs1!OtherNames & ""
                load_Refund_memberno (txtMemberNo)
                
        Else
            lblfullnames = ""
        End If
End Sub

Private Sub txtmembernodd_Change()
Dim rsNames As New ADODB.Recordset
    Dim cn As New ADODB.Connection
    Dim myclass As New cdbase
    Dim rsContrib As New ADODB.Recordset
    ''// get names and how one has been contributing
    
    Set rsNames = oSaccoMaster.GetRecordset("sp_getNames '" & txtmembernodd & "'")
        If Not rsNames.EOF Then
            lblmajinas.Caption = rsNames!surname & "" & "  " & rsNames!OtherNames & ""
        Else
            lblmajinas.Caption = ""
        End If
    
    ''/// it might be necesary but just do it
        Set rsContrib = oSaccoMaster.GetRecordset("select * from contrib where memberno ='" & txtmembernodd & "' order by ContrDate asc,refno asc")
            
            If Not rsContrib.EOF Then
            
                lsvvvv.ListItems.Clear
                
                Do While Not rsContrib.EOF
                     Set li = lsvvvv.ListItems.Add(, , rsContrib!memberno)
                         li.ListSubItems.Add , , rsContrib!ContrDate
                         li.ListSubItems.Add , , rsContrib!RefNo & ""
                         li.ListSubItems.Add , , rsContrib!amount & ""
                         li.ListSubItems.Add , , rsContrib!shareBal & ""
                         li.ListSubItems.Add , , rsContrib!shareBal & ""
                         li.ListSubItems.Add , , rsContrib!transby & ""
                         li.ListSubItems.Add , , rsContrib!ReceiptNo & ""
                    
                       rsContrib.MoveNext
                Loop
            Else
                lsvvvv.ListItems.Clear
            End If
    
End Sub


Private Sub Get1(NA As String)
Dim myclass As Object
Dim rs9 As New ADODB.Recordset
Set myclass = New ADODB.Connection

Set myclass = New cdbase
    Set cn = CreateObject("adodb.connection")
    Provider = myclass.OpenCon
   cn.Open Provider, "atm", "atm"
    
    '// get to know where the gls  are coming from
    
    sql = ""
    sql = "select * from param"
    
    rs9.Open sql, cn
    
    If Not rs9.EOF Then
    
        If rs9!GeneralLegerOpt = "FOSA" Then
        
        If cn.State = adStateOpen Then
            cn.Close
        End If
        Provider = "dsn_FOSA"
       cn.Open Provider, "atm", "atm"
        sql = "select * from cuB where name='" & NA & "'"
        Set rs = New ADODB.Recordset
        
        rs.Open sql, cn
        If Not rs.EOF Then
        If Not IsNull(rs.Fields("accno")) Then glaccno = rs.Fields("accno")
        If Not IsNull(rs.Fields("accountname")) Then glnamE1 = rs.Fields("name")
        If Not IsNull(rs.Fields("accountname")) Then lblPettyCash = rs.Fields("name")
        If Not IsNull(rs.Fields("idno")) Then glidno = rs.Fields("idno")
        If Not IsNull(rs.Fields("memberno")) Then glmemno = rs.Fields("memberno")
        If Not IsNull(rs.Fields("payno")) Then glpayno = rs.Fields("payno")
        End If
        'glPremium = Scheme_GL_Field(AccountCode, "glPremium")
        'bookba = cub_balance(glaccno)
        
        
        Else
        
                sql = "select * from cuB where name='" & lblPettyCash & "'"
            Set rs = New ADODB.Recordset
            
            rs.Open sql, cn
            If Not rs.EOF Then
            If Not IsNull(rs.Fields("accno")) Then glaccno = rs.Fields("accno")
            If Not IsNull(rs.Fields("accountname")) Then glnamE1 = rs.Fields("name")
            If Not IsNull(rs.Fields("idno")) Then glidno = rs.Fields("idno")
            If Not IsNull(rs.Fields("memberno")) Then glmemno = rs.Fields("memberno")
            If Not IsNull(rs.Fields("payno")) Then glpayno = rs.Fields("payno")
            End If
            'glPremium = Scheme_GL_Field(AccountCode, "glPremium")
            'bookba = cub_balance(glaccno)

        End If
    End If
End Sub

Private Sub Get2(NA As String)
Dim myclass As Object
Dim rs9 As New ADODB.Recordset
Set myclass = New ADODB.Connection

Set myclass = New cdbase
    Set cn = CreateObject("adodb.connection")
    Provider = myclass.OpenCon
   cn.Open Provider, "atm", "atm"
    
    '// get to know where the gls  are coming from
    
    sql = ""
    sql = "select * from param"
    
    rs9.Open sql, cn
    
    If Not rs9.EOF Then
    
        If rs9!GeneralLegerOpt = "FOSA" Then
        
        If cn.State = adStateOpen Then
            cn.Close
        End If
        Provider = "dsn_FOSA"
       cn.Open Provider, "atm", "atm"
        sql = "select * from cuB where accno='" & NA & "'"
        Set rs = New ADODB.Recordset
        
        rs.Open sql, cn
        If Not rs.EOF Then
        If Not IsNull(rs.Fields("accno")) Then glaccno = rs.Fields("accno")
        If Not IsNull(rs.Fields("accountname")) Then glnamE1 = rs.Fields("name")
        If Not IsNull(rs.Fields("accountname")) Then Label16 = rs.Fields("name")
        If Not IsNull(rs.Fields("idno")) Then glidno = rs.Fields("idno")
        If Not IsNull(rs.Fields("memberno")) Then glmemno = rs.Fields("memberno")
        If Not IsNull(rs.Fields("payno")) Then glpayno = rs.Fields("payno")
        Else
        Label16 = ""
        End If
        'glPremium = Scheme_GL_Field(AccountCode, "glPremium")
        Else
        
                sql = "select * from cuB where name='" & lblPettyCash & "'"
                    Set rs = New ADODB.Recordset
            
            rs.Open sql, cn
            If Not rs.EOF Then
            If Not IsNull(rs.Fields("accno")) Then glaccno = rs.Fields("accno")
            If Not IsNull(rs.Fields("accountname")) Then glnamE1 = rs.Fields("name")
            If Not IsNull(rs.Fields("idno")) Then glidno = rs.Fields("idno")
            If Not IsNull(rs.Fields("memberno")) Then glmemno = rs.Fields("memberno")
            If Not IsNull(rs.Fields("payno")) Then glpayno = rs.Fields("payno")
            End If
            'glPremium = Scheme_GL_Field(AccountCode, "glPremium")
            'bookba = cub_balance(glaccno)
        
        End If
        
    End If
End Sub

Private Sub LoanGuranter(memberno As String, Loanno As String, LoanAmount As Currency)
Dim myclass As New cdbase
Dim rsSumGuranteed As New ADODB.Recordset
Dim RsGuranters As New ADODB.Recordset
Dim sql As String
Dim sumGuranteed As Currency
            
            
            strSQL = ""
            
            strSQL = "select * from loanguar where  LoanNo='" & Loanno & "'"
            
            Set RsGuranters = oSaccoMaster.GetRecordset(strSQL)
            
            strSQL = "select sum(balance) as amtGuranted from loanguar where LoanNo='" & Loanno & "'"
            
            Set rsSumGuranteed = oSaccoMaster.GetRecordset(strSQL)
            
            If IsNull(rsSumGuranteed!amtguranted) Then
            
            sumGuranteed = 0
            
            Else
            sumGuranteed = rsSumGuranteed!amtguranted
            
            
            End If
            
            
            If Not RsGuranters.EOF Then
                Do While Not RsGuranters.EOF
                Dim GuaBalance As Currency
                
                    ''/// reduce each loan gurantereed by the much this guy is contributing
                    ''// check if amount principal being contributed is more  then the amount guranteed
                    
                    If RsGuranters!balance > 0 Then
                    
                        strSQL = ""
                        GuaBalance = ((CCur(RsGuranters!balance)) - ((RsGuranters!balance) * LoanAmount / sumGuranteed))
                        
                        If GuaBalance < 0 Then
                        GuaBalance = 0
                        End If
                        
                        
                        strSQL = "Update Loanguar set balance =" & GuaBalance & " where memberno ='" & RsGuranters!memberno & "' and loanno ='" & Loanno & "'"
                        
                        oSaccoMaster.ExecuteThis (strSQL)
                    Else
                        strSQL = ""
                        
                        strSQL = "Update Loanguar set Balance = 0 where memberno ='" & RsGuranters!memberno & "' and loanno ='" & Loanno & "'"
                        
                        oSaccoMaster.ExecuteThis strSQL
                        
                    End If
                    
                RsGuranters.MoveNext
                
                Loop
            End If
            

End Sub

Private Sub Transact(Loanno As String)


'// things are here to dispute the idea of balance base on the formula' a - reducing bal,b for straight line method,c for amortized methods
   '============ Get Information from LoanBal and Populate TextBoxes ============'
    'On Error GoTo ErrorTrap
    Set rst = oSaccoMaster.GetRecordset("select L.LoanNo,L.Balance,L.Repayperiod,L.RepayRate," _
    & "L.AutoCalc,L.IntrAmount,C.Amount,L.RepayMethod,LT.Interest from (LOANBAL L inner join" _
    & " LOANTYPE LT on L.LoanCode=LT.LoanCode) inner join CHEQUES C on L.LoanNo" _
    & "=C.LoanNo where L.LoanNo='" & Loanno & "'")
    Dim rst7 As Recordset
    Set rst7 = oSaccoMaster.GetRecordset("select top 1* from repay where loanno='" & Loanno & "' order by paymentno desc")
   
    With rst
        If Not .EOF Then
           
            If !AutoCalc = "Yes" Then
                Select Case !repaymethod
                    Case "AMRT"
                    'txtTotal.Text = !repayrate
                    If Not rst7.EOF Then
                    txtInterest = (!interest / 12 / 100) * rst7!loanbalance
                    Else
                     txtInterest = (!interest / 12 / 100) * !balance
                    End If
                    txtPrincipal = !repayrate - CCur(txtInterest)
                    Case "STL"
                    txtPrincipal = !repayrate
                    'txtInterest.Text = (!interest / 12 / 100) * (!amount / !RepayPeriod)
                    'txtTotal.Text = CCur(txtPrincipal.Text) + CCur(txtInterest.Text)
                    txtInterest = (!interest / 12 / 100) * (!amount)
                    Case "RBAL"
                    txtPrincipal = !repayrate
                    If Not rst7.EOF Then
                    txtInterest = (!interest / 12 / 100) * rst7!loanbalance
                    Else
                      txtInterest = (!interest / 12 / 100) * !balance
                    End If
                    'txtTotal.Text = CCur(txtPrincipal.Text) + CCur(txtInterest.Text)
                End Select
                
            ElseIf !AutoCalc = "No" Then
                txtInterest = !IntrAmount
                txtPrincipal = !repayrate
            End If
            If txtPrincipal = 0 Then txtPrincipal = 0
            If CCur(txtPrincipal) > !balance Then
                txtPrincipal = !balance
            End If
        End If
    End With
    If txtInterest = 0 Then txtInterest = 0
    txtPrincipal = Format(txtPrincipal, Cfmt)
    txtInterest = Format(txtInterest, Cfmt)
    'txtTotal = Format(CCur((txtinterest) + CCur(txtPrincipal)), CfMt)
    
    End Sub

Private Sub txtmno_Change()
''// get all the over recoveries for the selected guy
    Dim mysql As String
    Dim cn As New ADODB.Connection
    Dim rs10 As New ADODB.Recordset
    Dim myclass As New cdbase
    
    mysql = "select * from overRecovery where   memberno ='" & txtmno.Text & "' and paid =0"
    
    Set rs10 = oSaccoMaster.GetRecordset(mysql)
    
    If Not rs10.EOF Then
        lsvOverRecovery.ListItems.Clear
        lsvVouchers.ListItems.Clear
            Do While Not rs10.EOF
                Set li = lsvVouchers.ListItems.Add(, , rs10!memberno)
                li.ListSubItems.Add , , rs10!ContrDate
                li.ListSubItems.Add , , Format(rs10!shareBal, "#########.00")
                li.ListSubItems.Add , , Format(rs10!interest, "#########.00")
                li.ListSubItems.Add , , rs10!Loancode & ""
                li.ListSubItems.Add , , rs10!Remarks & ""
                li.ListSubItems.Add , , rs10!RefNo & ""
                li.ListSubItems.Add , , rs10!auditid & ""
                li.ListSubItems.Add , , rs10!Loanno & ""
                
            rs10.MoveNext
            Loop
    End If


End Sub

Private Sub txtRMemberno_Change()
    ''// get all the over recoveries for the selected guy
    Dim mysql As String
    Dim cn As New ADODB.Connection
    Dim rs10 As New ADODB.Recordset
    Dim myclass As New cdbase
    
    mysql = "select * from overRecovery where   memberno ='" & txtRMemberno.Text & "' and paid =0"
    
    Set rs10 = oSaccoMaster.GetRecordset(mysql)
    
    If Not rs10.EOF Then
        lsvOverRecovery.ListItems.Clear
            Do While Not rs10.EOF
                Set li = lsvOverRecovery.ListItems.Add(, , rs10!memberno)
                li.ListSubItems.Add , , rs10!ContrDate
                li.ListSubItems.Add , , Format(rs10!shareBal, "#########.00")
                li.ListSubItems.Add , , Format(rs10!interest, "#########.00")
                li.ListSubItems.Add , , rs10!Loancode & ""
                li.ListSubItems.Add , , rs10!Remarks & ""
                li.ListSubItems.Add , , rs10!RefNo & ""
                li.ListSubItems.Add , , rs10!auditid & ""
                li.ListSubItems.Add , , rs10!Loanno & ""
                
            rs10.MoveNext
            Loop
    End If
    
    If txtRMemberno.Text = "" Then
    mysql = ""
    mysql = "select * from  OverRecovery  where paid =0 and (amount+interest)>1 order by memberno,ContrDate asc"
    Set rs10 = oSaccoMaster.GetRecordset(mysql)
    
    If Not rs10.EOF Then
        lsvOverRecovery.ListItems.Clear
            Do While Not rs10.EOF
                Set li = lsvOverRecovery.ListItems.Add(, , rs10!memberno)
                li.ListSubItems.Add , , rs10!ContrDate
                li.ListSubItems.Add , , Format(rs10!amount, "#########.00")
                li.ListSubItems.Add , , Format(rs10!interest, "#########.00")
                li.ListSubItems.Add , , rs10!Remarks & ""
                li.ListSubItems.Add , , rs10!RefNo & ""
                li.ListSubItems.Add , , rs10!auditid & ""
                
            rs10.MoveNext
            Loop
    End If
    
        
    End If
    
End Sub
Private Sub Get_Vno()
Dim rsr As New ADODB.Recordset
Dim rsg As New ADODB.Recordset
Dim I As Object
Dim Mylength As Integer
Dim Mylength1 As String
'//if this record is new then look for ag_receipts no
''//clear all textboxes

mysql = ""
mysql = "select GenerateReceiptno from param"

Set rsg = oSaccoMaster.GetRecordset(mysql)
If Not rsg.EOF Then
    ''''check check
    If rsg!GenerateReceiptno = True Then
        ''//CHECK OPTION CHECKED
        If OptCashMethod.value = True Then
        'OptCash.Value = True
            mysql = ""
            mysql = "select * from VOurchernos where Vourcherno like 'CVNO%' order by VourchernoId desc"
        
            Set rsr = oSaccoMaster.GetRecordset(mysql)
        
            If Not rsr.EOF Then
                
                Mylength = CInt(Mid(rsr!Vourcherno, 4, 7))
                Mylength = Mylength + 1
                txtvno = Padding(Mylength)
                txtvno = "CVNO" & txtvno
            Else
                Mylength = 1
                txtvno = "CVNO" & Padding(Mylength)
                
            End If
        ElseIf OptSchedule.value = True Then
        'OpCheque.Value = True
        
            mysql = ""
            mysql = "select * from VOurchernos where Vourcherno like 'SVNO%' order by VourchernoId desc"
        
            Set rsr = oSaccoMaster.GetRecordset(mysql)
        
            If Not rsr.EOF Then
                
                Mylength = CInt(Mid(rsr!Vourcherno, 5, 7))
                Mylength = Mylength + 1
                txtvno = Padding(Mylength)
                txtvno = "SVNO" & txtvno
            Else
                Mylength = 1
                txtvno = "SVNO" & Padding(Mylength)
                
            End If
        
        End If
    End If
End If
End Sub
Private Sub load_Refund_memberno(memberno As String)
    Dim rsm As New ADODB.Recordset
    
    mysql = ""
    mysql = "select loanno  from  loanbal  where  memberno  ='" & memberno & "' order  by firstdate"
    
    Set rsm = oSaccoMaster.GetRecordset(mysql)
    
    If Not rsm.EOF Then
        cboLoanno.Clear
        Do While Not rsm.EOF
            cboLoanno.AddItem (rsm!Loanno & "")
            rsm.MoveNext
        Loop
    Else
        cboLoanno.Clear
    End If
End Sub




Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Caption = "SHARES" Then
cmdPostShares.Visible = True
cmdPost.Visible = False
ElseIf SSTab1.Caption = "LOANS" Then
cmdPost.Visible = True
cmdPostShares.Visible = False
End If
End Sub

Private Sub txtAmount_Change()
'txtAmount = txtSharesTotalAmount
End Sub

Private Sub txtLoanDestMemberNo_Change()
Dim rsNames As New ADODB.Recordset
Dim sql As String

Set rsNames = oSaccoMaster.GetRecordset("select surname,othernames,memberno,rank,Initshares,staffno from members where memberno ='" & txtLoanDestMemberNo.Text & "'")

If Not rsNames.EOF Then
        lblLoanfullName.Caption = rsNames!surname & "" & "  " & rsNames!OtherNames & ""
Else
        lblLoanfullName.Caption = ""
End If

'Populate Loans
Call GetLoans(Trim(txtLoanDestMemberNo), cboLoanLoanNo)
End Sub
Private Sub GetLoans(memberno As String, myCbo As ComboBox)
Dim RsLoans As New ADODB.Recordset
'Get Loans the Member has
Set RsLoans = oSaccoMaster.GetRecordset("select LoanNo from LOANS where memberno ='" & memberno & "'")
If Not RsLoans.EOF Then
    myCbo.Clear
    
    While Not RsLoans.EOF
        myCbo.AddItem RsLoans!Loanno
        RsLoans.MoveNext
    Wend
Else
    myCbo.Clear
End If
End Sub
Private Sub txtLoanMemberno_Change()
Dim myclass  As New cdbase
Dim RsRecovery As New ADODB.Recordset
Dim rsNames As New ADODB.Recordset
Dim RsLoans As New ADODB.Recordset
Dim sql As String

Set rsNames = oSaccoMaster.GetRecordset("select surname,othernames,memberno,rank,Initshares,staffno from members where memberno ='" & txtLoanMemberno.Text & "'")

If Not rsNames.EOF Then
        lblfullnames.Caption = rsNames!surname & "" & "  " & rsNames!OtherNames & ""
Else
        lblfullnames.Caption = ""
End If

'Get Loans the Member has
Set RsLoans = oSaccoMaster.GetRecordset("select LoanNo from LOANS where memberno ='" & txtLoanMemberno.Text & "'")
If Not RsLoans.EOF Then
    cboLoanno.Clear
    
    While Not RsLoans.EOF
        cboLoanno.AddItem RsLoans!Loanno
        RsLoans.MoveNext
    Wend
Else
cboLoanno.Clear
End If

End Sub

Private Sub txtShareMemberNo_Change()
Dim myclass  As New cdbase

Dim rsNames As New ADODB.Recordset
Dim RsLoans As New ADODB.Recordset
Dim sql As String

Set rsNames = oSaccoMaster.GetRecordset("select surname,othernames,memberno,rank,Initshares,staffno from members where memberno ='" & txtShareMemberNo.Text & "'")

If Not rsNames.EOF Then
        lblNames.Caption = rsNames!surname & "" & "  " & rsNames!OtherNames & ""
Else
        lblNames.Caption = ""
End If
Call PopulateList
End Sub
Private Sub PopulateList()
Dim RsRecovery As New ADODB.Recordset
sql = ""
txtSharesTotalAmount = 0
txtamount = 0
txtInterestAmount = 0

sql = "select * from Contrib where memberno ='" & txtShareMemberNo.Text & "' and Schemecode='" & Trim(cboShareType) & "' order by contrdate"

Set RsRecovery = oSaccoMaster.GetRecordset(sql)

If Not RsRecovery.EOF Then

lvwShareContrib.ListItems.Clear

    Do While Not RsRecovery.EOF
        Set li = lvwShareContrib.ListItems.Add(, , RsRecovery!ContrDate)
            li.ListSubItems.Add , , RsRecovery!amount
            li.ListSubItems.Add , , IIf(IsNull(RsRecovery!chequeno), "", RsRecovery!chequeno)
            li.ListSubItems.Add , , IIf(IsNull(RsRecovery!ReceiptNo), "", RsRecovery!ReceiptNo)
            li.ListSubItems.Add , , RsRecovery!transby
            li.ListSubItems.Add , , RsRecovery!shareBal
            li.ListSubItems.Add , , RsRecovery!auditid
            RsRecovery.MoveNext
    Loop
Else
lvwShareContrib.ListItems.Clear
End If
End Sub
Private Sub txtSharesDestMemberNo_Change()
Dim RsRecovery As New ADODB.Recordset
Dim rsNames As New ADODB.Recordset
Dim RsLoans As New ADODB.Recordset
Dim sql As String

Set rsNames = oSaccoMaster.GetRecordset("select surname,othernames,memberno,rank,Initshares,staffno from members where memberno ='" & txtSharesDestMemberNo.Text & "'")

If Not rsNames.EOF Then
        lblSharesFullName.Caption = rsNames!surname & "" & "  " & rsNames!OtherNames & ""
Else
        lblSharesFullName.Caption = ""
End If
'Populate Loans
Call GetLoans(Trim(txtSharesDestMemberNo), cboSharesLoanNo)
End Sub

