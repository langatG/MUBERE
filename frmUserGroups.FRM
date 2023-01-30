VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmUserGroups 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " System User Groups"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8940
   ForeColor       =   &H80000018&
   Icon            =   "frmUserGroups.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   8940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLast 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   4
      ToolTipText     =   "Move to Last record"
      Top             =   6000
      Width           =   495
   End
   Begin VB.CommandButton cmdCancel 
      Height          =   495
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Cancel Process"
      Top             =   6000
      Width           =   495
   End
   Begin VB.CommandButton cmdUpdate 
      Height          =   495
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Save Record"
      Top             =   6000
      Width           =   495
   End
   Begin VB.CommandButton cmdDelete 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Delete Record"
      Top             =   6000
      Width           =   495
   End
   Begin VB.CommandButton cmdEdit 
      Height          =   495
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Edit Record"
      Top             =   6000
      Width           =   495
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   3
      ToolTipText     =   "Move to the Next"
      Top             =   6000
      Width           =   495
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   2
      ToolTipText     =   "Move to the Previous record"
      Top             =   6000
      Width           =   495
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Move to the Last record"
      Top             =   6000
      Width           =   495
   End
   Begin VB.CommandButton cmdAdd 
      Height          =   495
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Add New record"
      Top             =   6000
      Width           =   495
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "CLOSE"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7560
      TabIndex        =   51
      Top             =   6000
      Width           =   1230
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3120
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserGroups.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserGroups.frx":006A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   64
      Top             =   0
      Width           =   8940
      _ExtentX        =   15769
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
   End
   Begin VB.Frame fraAllContents 
      Height          =   5085
      Left            =   120
      TabIndex        =   0
      Top             =   735
      Width           =   8715
      Begin VB.CheckBox chkOtherSchemes 
         Caption         =   "Other Schemes"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2040
         TabIndex        =   52
         Top             =   960
         Width           =   1815
      End
      Begin VB.CheckBox chkShares 
         Caption         =   "Shares"
         DataField       =   "Shares"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   24
         Top             =   4080
         Width           =   1215
      End
      Begin VB.CheckBox chkMembership 
         Caption         =   "Membership"
         DataField       =   "Membership"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   210
         TabIndex        =   12
         Top             =   960
         Width           =   1605
      End
      Begin VB.CheckBox chkUtilities 
         Caption         =   "Utilities"
         DataField       =   "Utilities"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4200
         TabIndex        =   27
         Top             =   960
         Width           =   1335
      End
      Begin VB.CheckBox chkLoans 
         Caption         =   "Loans"
         DataField       =   "Loans"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   210
         TabIndex        =   16
         Top             =   2520
         Width           =   1245
      End
      Begin VB.Frame fraLoans 
         Caption         =   "dfdg"
         Height          =   1500
         Left            =   120
         TabIndex        =   60
         Top             =   2520
         Width           =   3795
         Begin VB.CheckBox chkEffectRepay 
            Caption         =   "Effect Repayment"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1860
            TabIndex        =   22
            Top             =   600
            Width           =   1830
         End
         Begin VB.CheckBox chkLoanBal 
            Caption         =   "Loan Balance"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   90
            TabIndex        =   20
            Top             =   1110
            Width           =   1455
         End
         Begin VB.CheckBox chkTransactions 
            Caption         =   "Transactions"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1860
            TabIndex        =   23
            Top             =   960
            Width           =   1425
         End
         Begin VB.CheckBox chkChequeEntry 
            Caption         =   "Cheque Entry"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1860
            TabIndex        =   21
            Top             =   300
            Width           =   1470
         End
         Begin VB.CheckBox chkLoanEndorsement 
            Caption         =   "Endorsement"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   90
            TabIndex        =   19
            Top             =   825
            Width           =   1440
         End
         Begin VB.CheckBox chkLoanGuar 
            Caption         =   "Guarantors"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   105
            TabIndex        =   18
            Top             =   555
            Width           =   1275
         End
         Begin VB.CheckBox chkLoanApp 
            Caption         =   "Application"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   90
            TabIndex        =   17
            Top             =   255
            Width           =   1230
         End
      End
      Begin VB.Frame fraOtherSchemes 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1440
         Left            =   1905
         TabIndex        =   63
         Top             =   960
         Width           =   2010
         Begin VB.CheckBox chkBenevolentFund 
            Caption         =   "Benevolent Fund"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   54
            Top             =   705
            Width           =   1725
         End
         Begin VB.CheckBox chkSavings 
            Caption         =   "Savings"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   53
            Top             =   360
            Width           =   1005
         End
      End
      Begin VB.CheckBox chkSetup 
         Caption         =   "Setup"
         DataField       =   "Setup"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4200
         TabIndex        =   39
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Frame fraUtilities 
         Height          =   1815
         Left            =   4125
         TabIndex        =   62
         Top             =   960
         Width           =   4545
         Begin VB.CheckBox chkBackup 
            Caption         =   "Backup"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   2160
            TabIndex        =   67
            Top             =   1590
            Width           =   1095
         End
         Begin VB.CheckBox chkArchived 
            Caption         =   "Archived"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2145
            TabIndex        =   37
            Top             =   1140
            Width           =   1065
         End
         Begin VB.CheckBox chkWithdrawn 
            Caption         =   "Withdrawn"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   32
            Top             =   1500
            Width           =   1200
         End
         Begin VB.CheckBox chkDormant 
            Caption         =   "Dormant"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2160
            TabIndex        =   38
            Top             =   1380
            Width           =   1095
         End
         Begin VB.CheckBox chkUtilGuar 
            Caption         =   "Guarantor"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2130
            TabIndex        =   36
            Top             =   885
            Width           =   1170
         End
         Begin VB.CheckBox chkExportToGL 
            Caption         =   "Export To GL"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   105
            TabIndex        =   28
            Top             =   300
            Width           =   1425
         End
         Begin VB.CheckBox chkUtilStatements 
            Caption         =   "Statements"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2115
            TabIndex        =   33
            Top             =   165
            Width           =   1335
         End
         Begin VB.CheckBox chkMonthlyDeductions 
            Caption         =   "Monthly Deductions"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2130
            TabIndex        =   35
            Top             =   660
            Width           =   1995
         End
         Begin VB.CheckBox chkPeriodicTran 
            Caption         =   "Period Transactions"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2130
            TabIndex        =   34
            Top             =   420
            Width           =   2010
         End
         Begin VB.CheckBox chkDeductions 
            Caption         =   "Deductions"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   105
            TabIndex        =   31
            Top             =   1170
            Width           =   1260
         End
         Begin VB.CheckBox chkDividends 
            Caption         =   "Dividends"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   105
            TabIndex        =   30
            Top             =   870
            Width           =   1170
         End
         Begin VB.CheckBox chkCalc 
            Caption         =   "Calculator"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   105
            TabIndex        =   29
            Top             =   585
            Width           =   1200
         End
      End
      Begin VB.Frame fraSetup 
         Height          =   2010
         Left            =   4095
         TabIndex        =   61
         Top             =   2880
         Width           =   4455
         Begin VB.CheckBox chkActivate 
            Caption         =   "Activate"
            Height          =   225
            Left            =   2190
            TabIndex        =   66
            Top             =   1620
            Width           =   1995
         End
         Begin VB.CheckBox chkChangeMembershipNo 
            Caption         =   "Change Membersip No"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2190
            TabIndex        =   50
            Top             =   1350
            Width           =   2235
         End
         Begin VB.CheckBox chkClearMemRecords 
            Caption         =   "Clear Mem Records"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   44
            Top             =   1320
            Width           =   2040
         End
         Begin VB.CheckBox chkClearLoanRecords 
            Caption         =   "Clear Loan Records"
            Height          =   225
            Left            =   120
            TabIndex        =   45
            Top             =   1605
            Width           =   1995
         End
         Begin VB.CheckBox chkDatabaseSetup 
            Caption         =   "Database Setup"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2160
            TabIndex        =   46
            Top             =   270
            Width           =   1665
         End
         Begin VB.CheckBox chkSysUsers 
            Caption         =   "System Users"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2175
            TabIndex        =   49
            Top             =   1065
            Width           =   1515
         End
         Begin VB.CheckBox chkUserGrps 
            Caption         =   "User Groups"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   2175
            TabIndex        =   48
            Top             =   795
            Width           =   1395
         End
         Begin VB.CheckBox chkRejReasons 
            Caption         =   "Rejection Reasons"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2160
            TabIndex        =   47
            Top             =   540
            Width           =   1920
         End
         Begin VB.CheckBox chkBankSetup 
            Caption         =   "Bank Setup"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   43
            Top             =   1080
            Width           =   1335
         End
         Begin VB.CheckBox chkCompanySetup 
            Caption         =   "Company Setup"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   42
            Top             =   840
            Width           =   1785
         End
         Begin VB.CheckBox chkLoanType 
            Caption         =   "Loan Types"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   41
            Top             =   585
            Width           =   1650
         End
         Begin VB.CheckBox chkParametization 
            Caption         =   "Parameterization"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   40
            Top             =   285
            Width           =   1905
         End
      End
      Begin VB.Frame fraShares 
         Height          =   855
         Left            =   120
         TabIndex        =   59
         Top             =   4035
         Width           =   3795
         Begin VB.CheckBox chkShareVar 
            Caption         =   "Share Variation"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1725
            TabIndex        =   26
            Top             =   360
            Width           =   1635
         End
         Begin VB.CheckBox chkContributions 
            Caption         =   "Contribution"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   360
            Width           =   1365
         End
      End
      Begin VB.Frame fraMembers 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1440
         Left            =   120
         TabIndex        =   58
         Top             =   960
         Width           =   1755
         Begin VB.CheckBox chkMemberReg 
            Caption         =   "Registration"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   90
            TabIndex        =   13
            Top             =   375
            Width           =   1320
         End
         Begin VB.CheckBox chkMemStatement 
            Caption         =   "Statement"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   75
            TabIndex        =   15
            Top             =   1050
            Width           =   1215
         End
         Begin VB.CheckBox chkNextOfKin 
            Caption         =   "Next Of Kin"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   75
            TabIndex        =   14
            Top             =   720
            Width           =   1290
         End
      End
      Begin VB.Frame fraUserGrps 
         Caption         =   "User Groups"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   120
         TabIndex        =   55
         Top             =   120
         Width           =   8475
         Begin VB.TextBox txtGroupID 
            DataField       =   "GroupID"
            DataSource      =   "Adodc1"
            Height          =   285
            Left            =   975
            TabIndex        =   10
            Top             =   255
            Width           =   1695
         End
         Begin VB.TextBox txtDescription 
            DataField       =   "GroupName"
            DataSource      =   "Adodc1"
            Height          =   285
            Left            =   3840
            TabIndex        =   11
            Top             =   255
            Width           =   4380
         End
         Begin VB.Label Label2 
            Caption         =   "&Group ID"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   57
            Top             =   240
            Width           =   1140
         End
         Begin VB.Label Label3 
            Caption         =   "&Description"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2760
            TabIndex        =   56
            Top             =   240
            Width           =   1080
         End
      End
   End
   Begin MSComctlLib.ListView lvwSummary 
      Height          =   4965
      Left            =   120
      TabIndex        =   65
      Top             =   750
      Width           =   8730
      _ExtentX        =   15399
      _ExtentY        =   8758
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   16711680
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
End
Attribute VB_Name = "FrmUserGroups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim disablemodifying As Boolean
Private Sub chkArchived_Click()
On Error GoTo errFix
If chkArchived.Value = vbChecked And manuallychecked = True Then
    getPermision "archived"
    showPermisions chkArchived
    archivedpermision = permision
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub
Private Sub chkActivate_Click()
On Error GoTo errFix
If chkActivate.Value = vbChecked And manuallychecked = True Then
    getPermision "activate"
    showPermisions chkActivate
    activatepermision = permision
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub
Private Sub chkActivate_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    chkActivate_Click
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub
Private Sub chkArchived_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    chkArchived_Click
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkBankSetup_Click()
On Error GoTo errFix
If chkBankSetup.Value = vbChecked And manuallychecked = True Then
    getPermision "banksetup"
    showPermisions chkBankSetup
    banksetuppermision = permision
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkBankSetup_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    chkBankSetup_Click
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkBenevolentFund_Click()
On Error GoTo errFix
If chkBenevolentFund.Value = vbChecked Then
    getPermision "benevolentFund"
    showPermisions chkBenevolentFund
    benfundpermision = permision
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkCalculator_Click()
On Error GoTo errFix
If chkCalc.Value = vbChecked Then
    showPermisions chkLoanBal
    calculatorpermision = permision
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkCalc_Click()
On Error GoTo errFix
If chkCalc.Value = vbChecked And manuallychecked = True Then
    getPermision "calc"
    showPermisions chkCalc
    calculatorpermision = permision
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkCalc_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    chkCalc_Click
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkChangeMembershipNo_Click()
On Error GoTo errFix
If chkChangeMembershipNo.Value = vbChecked And manuallychecked = True Then
    getPermision "changememno"
    showPermisions chkChangeMembershipNo
    changenopermision = permision
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkChangeMembershipNo_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    chkChangeMembershipNo_Click
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkChequeEntry_Click()
On Error GoTo errFix
If chkChequeEntry.Value = vbChecked And manuallychecked = True Then
    getPermision "chequeentry"
    showPermisions chkChequeEntry
    chequeentrypermision = permision
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkChequeEntry_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    chkChequeEntry_Click
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkClearLoanRecords_Click()
On Error GoTo errFix
If chkClearLoanRecords.Value = vbChecked And manuallychecked = True Then
    showPermisions chkClearLoanRecords
    clearloanpermision = permision
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkClearLoanRecords_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    chkClearLoanRecords_Click
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkClearMemRecords_Click()
On Error GoTo errFix
If chkClearMemRecords.Value = vbChecked And manuallychecked = True Then
    getPermision "clearmemrecs"
    showPermisions chkClearMemRecords
    clearmempermision = permision
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkClearMemRecords_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    chkClearMemRecords_Click
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkCompanySetup_Click()
On Error GoTo errFix
If chkCompanySetup.Value = vbChecked And manuallychecked = True Then
    getPermision "companysetup"
    showPermisions chkCompanySetup
    companysetpermision = permision
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkCompanySetup_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    chkCompanySetup_Click
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkContributions_Click()
On Error GoTo errFix
If chkContributions.Value = vbChecked And manuallychecked = True Then
    getPermision "contributions"
    showPermisions chkContributions
    contributionspermision = permision
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkContributions_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    chkContributions_Click
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkDatabaseSetup_Click()
On Error GoTo errFix
If chkDatabaseSetup.Value = vbChecked And manuallychecked = True Then
    getPermision "databasesetup"
    showPermisions chkDatabaseSetup
    databasesetpermision = permision
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkDatabaseSetup_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    chkDatabaseSetup_Click
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkDeductions_Click()
On Error GoTo errFix
If chkDeductions.Value = vbChecked And manuallychecked = True Then
    getPermision "deductions"
    showPermisions chkDeductions
    deductionspermision = permision
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkDeductions_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    chkDeductions_Click
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkDividends_Click()
On Error GoTo errFix
If chkDividends.Value = vbChecked And manuallychecked = True Then
    getPermision "dividends"
    showPermisions chkDividends
    dividendspermision = permision
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkDividends_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    chkDividends_Click
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkDormant_Click()
On Error GoTo errFix
If chkDormant.Value = vbChecked And manuallychecked = True Then
    getPermision "dormant"
    showPermisions chkDormant
    dormantpermision = permision
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub
Private Sub chkBackup_Click()
On Error GoTo errFix
If chkBackup.Value = vbChecked And manuallychecked = True Then
    getPermision "backup"
    showPermisions chkBackup
    backuppermision = permision
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub
Private Sub chkDormant_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    chkDormant_Click
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub
Private Sub chkBackup_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    chkBackup_Click
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub
Private Sub chkEffectRepay_Click()
On Error GoTo errFix
If chkEffectRepay.Value = vbChecked And manuallychecked = True Then
    getPermision "effectrepayment"
    showPermisions chkEffectRepay
    effectrepaypermision = permision
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkEffectRepay_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    chkEffectRepay_Click
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkExportToGL_Click()
On Error GoTo errFix
If chkExportToGL.Value = vbChecked And manuallychecked = True Then
    getPermision "exporttogl"
    showPermisions chkExportToGL
    exporttoglpermision = permision
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkExportToGL_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    chkExportToGL_Click
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkLoanApp_Click()
On Error GoTo errFix
If chkLoanApp.Value = vbChecked And manuallychecked = True Then
    getPermision "loanapplic"
    showPermisions chkLoanApp
    applicationpermision = permision
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub getPermision(theField As String) 'TO SET VALUE USED BY frmPermisions TO CHECK THE RIGHT OPTION AT FORM LOAD
On Error GoTo errFix
Set Rst5 = oSaccoMaster.GetRecordSet("select " & theField & " from usergrps where groupid= '" & txtGroupID.Text & "'")
    If Not Rst5.EOF Then
        'for display purposes to know which permision is set for thisgroup for a particular form
       valToEncrOrDecr = Rst5.Fields(theField) & ""
       Call EncryptPassword
       Select Case EncryptPass
       Case "Mod"
            Set permisionShown = frmPermisions.optMod
       Case "View"
            Set permisionShown = frmPermisions.optView
       Case "Non"
            Set permisionShown = frmPermisions.optNon
       End Select
    Else
        'if this is a new group by default show Mod button checked. Then it can be changedat will.
         Set permisionShown = frmPermisions.optMod
         
    End If
        Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub showPermisions(ctrl As Control)
On Error GoTo errFix
If ctrl.Value = vbChecked And manuallychecked = True Then
    Set contrl = ctrl 'contrl is the check box causing the loading of the frmPermisions(example: chkMemReg)
    frmPermisions.Show vbModal
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkLoanApp_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    chkLoanApp_Click
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkLoanBal_Click()
On Error GoTo errFix
If chkLoanBal.Value = vbChecked And manuallychecked = True Then
    getPermision "loanbal"
    showPermisions chkLoanBal
    loanbalpermision = permision
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkLoanBal_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    chkLoanBal_Click
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkLoanEndorsement_Click()
On Error GoTo errFix
If chkLoanEndorsement.Value = vbChecked And manuallychecked = True Then
    getPermision "loanendorsement"
    showPermisions chkLoanEndorsement
    loanendorsementpermision = permision
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkLoanEndorsement_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    chkLoanEndorsement_Click
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkLoanGuar_Click()
On Error GoTo errFix
If chkLoanGuar.Value = vbChecked And manuallychecked = True Then
    getPermision "loanGuarantors"
    showPermisions chkLoanGuar
    loanguarantorspermision = permision
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkLoanGuar_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    chkLoanGuar_Click
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkLoans_Click()
On Error GoTo errFix
If chkLoans.Value = vbUnchecked Then
    chkLoanApp.Value = vbUnchecked
    chkLoanGuar.Value = vbUnchecked
    chkLoanEndorsement.Value = vbUnchecked
    chkLoanBal.Value = vbUnchecked
    chkChequeEntry.Value = vbUnchecked
    chkEffectRepay.Value = vbUnchecked
    chktransactions.Value = vbUnchecked
    fraLoans.Enabled = False
ElseIf chkLoans.Value = vbChecked Then
    fraLoans.Enabled = True
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkLoans_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    chkLoans_Click
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkLoanType_Click()
On Error GoTo errFix
If chkLoanType.Value = vbChecked And manuallychecked = True Then
    getPermision "loantypes"
    showPermisions chkLoanType
    loantypespermision = permision
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkLoanType_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    chkLoanType_Click
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkMemberReg_Click()
On Error GoTo errFix
If chkMemberReg.Value = vbChecked And manuallychecked = True Then
    getPermision "memreg"   'TO SET WHICH OPTION IS AUTOMATICALLY SELECTED WHEN THE POP UP FOR PERMISIONS SHOWS
    showPermisions chkMemberReg
    MemberRegPermision = permision
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkMemberReg_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    chkMemberReg_Click
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkMembership_Click()
On Error GoTo errFix
 If chkMembership.Value = vbUnchecked Then
    chkMemberReg.Value = vbUnchecked
    chkNextOfKin.Value = vbUnchecked
    chkMemStatement.Value = vbUnchecked
    fraMembers.Enabled = False
Else
    fraMembers.Enabled = True
 End If
     Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkMembership_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    chkMembership_Click
End If
 Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkMemStatement_Click()
On Error GoTo errFix
If chkMemStatement.Value = vbChecked And manuallychecked = True Then
    getPermision "memstatement"
    showPermisions chkMemStatement
    memstatementspermision = permision
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkMemStatement_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    chkMemStatement_Click
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkMonthlyDeductions_Click()
On Error GoTo errFix
If chkMonthlyDeductions.Value = vbChecked And manuallychecked = True Then
    getPermision "monthlydeductions"
    showPermisions chkMonthlyDeductions
    monthlydedpermision = permision
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkMonthlyDeductions_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    chkMonthlyDeductions_Click
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkNextOfKin_Click()
On Error GoTo errFix
If chkNextOfKin.Value = vbChecked And manuallychecked = True Then
    getPermision "nok"
    showPermisions chkNextOfKin
    NextOfKinPermision = permision
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkNextOfKin_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    chkNextOfKin_Click
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkOtherSchemes_Click()
On Error GoTo errFix
If chkOtherSchemes.Value = vbUnchecked Then
    chkBenevolentFund.Value = vbUnchecked
    chkSavings.Value = vbUnchecked
    fraOtherSchemes.Enabled = False
Else
    fraOtherSchemes.Enabled = True
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub



Private Sub chkOtherSchemes_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    chkOtherSchemes_Click
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkParametization_Click()
On Error GoTo errFix
If chkParametization.Value = vbChecked And manuallychecked = True Then
    getPermision "parametization"
    showPermisions chkParametization
    parampermision = permision
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkParametization_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    chkParametization_Click
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkPeriodicTran_Click()
On Error GoTo errFix
If chkPeriodicTran.Value = vbChecked And manuallychecked = True Then
    getPermision "periodictran"
    showPermisions chkPeriodicTran
    periodtranpermision = permision
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkPeriodicTran_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    chkPeriodicTran_Click
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkRejReasons_Click()
On Error GoTo errFix
If chkRejReasons.Value = vbChecked And manuallychecked = True Then
    getPermision "rejreasons"
    showPermisions chkRejReasons
    rejreasonpermision = permision
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkRejReasons_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    chkRejReasons_Click
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkSavings_Click()
If chkSavings.Value = vbChecked Then
    getPermision "savings"
    showPermisions chkSavings
    savingspermision = permision
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkSetup_Click()
On Error GoTo errFix
If chkSetup.Value = vbUnchecked Then
    chkParametization.Value = vbUnchecked
    chkLoanType.Value = vbUnchecked
    chkCompanySetup.Value = vbUnchecked
    chkBankSetup.Value = vbUnchecked
    chkClearMemRecords.Value = vbUnchecked
    chkClearLoanRecords.Value = vbUnchecked
    chkDatabaseSetup.Value = vbUnchecked
    chkUserGrps.Value = vbUnchecked
    chkSysUsers.Value = vbUnchecked
    chkChangeMembershipNo.Value = vbUnchecked
    chkRejReasons.Value = vbUnchecked
    fraSetup.Enabled = False
ElseIf chkSetup.Value = vbChecked Then
    fraSetup.Enabled = True
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkSetup_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    chkSetup_Click
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkShares_Click()
On Error GoTo errFix
If chkShares.Value = vbUnchecked Then
    chkContributions.Value = vbUnchecked
    chkShareVar.Value = vbUnchecked
    fraShares.Enabled = False
ElseIf chkShares.Value = vbChecked Then
    fraShares.Enabled = True
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkShares_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    chkShares_Click
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkShareVar_Click()
On Error GoTo errFix
If chkShareVar.Value = vbChecked And manuallychecked = True Then
    getPermision "sharevar"
    showPermisions chkShareVar
    sharevarpermision = permision
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub
Private Sub chkShareVar_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    chkShareVar_Click
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkSysUsers_Click()
On Error GoTo errFix
If chkSysUsers.Value = vbChecked And manuallychecked = True Then
    getPermision "sysusers"
    showPermisions chkSysUsers
    sysuserspermision = permision
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkSysUsers_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    chkSysUsers_Click
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkTransactions_Click()
On Error GoTo errFix
If chktransactions.Value = vbChecked And manuallychecked = True Then
    getPermision "loantrans"
    showPermisions chktransactions
    transactionspermision = permision
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkTransactions_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    chkTransactions_Click
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkUserGrps_Click()
On Error GoTo errFix
If chkUserGrps.Value = vbChecked And manuallychecked = True Then
    getPermision "usergrps"
    showPermisions chkUserGrps
    usergroupspermision = permision
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkUserGrps_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    chkUserGrps_Click
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkUtilGuar_Click()
On Error GoTo errFix
If chkUtilGuar.Value = vbChecked And manuallychecked = True Then
    getPermision "utilguarantor"
    showPermisions chkUtilGuar
    utilguarantorspermision = permision
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkUtilGuar_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    chkUtilGuar_Click
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkUtilities_Click()
On Error GoTo errFix
If chkUtilities.Value = vbUnchecked Then
    chkExportToGL.Value = vbUnchecked
    chkCalc.Value = vbUnchecked
    chkDividends.Value = vbUnchecked
    chkDeductions.Value = vbUnchecked
    chkWithdrawn.Value = vbUnchecked
    chkUtilStatements.Value = vbUnchecked
    chkPeriodicTran.Value = vbUnchecked
    chkMonthlyDeductions.Value = vbUnchecked
    chkUtilGuar.Value = vbUnchecked
    chkArchived.Value = vbUnchecked
    chkActivate.Value = vbUnchecked
    chkDormant.Value = vbUnchecked
    chkBackup.Value = vbUnchecked
    fraUtilities.Enabled = False
ElseIf chkUtilities.Value = vbChecked Then
    fraUtilities.Enabled = True
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkUtilities_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    chkUtilities_Click
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkUtilStatements_Click()
On Error GoTo errFix
If chkUtilStatements.Value = vbChecked And manuallychecked = True Then
    getPermision "utilstatements"
    showPermisions chkUtilStatements
    utilstatementpermision = permision
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkUtilStatements_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    chkUtilStatements_Click
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkWithdrawn_Click()
On Error GoTo errFix
If chkWithdrawn.Value = vbChecked And manuallychecked = True Then
    getPermision "withdrawn"
    showPermisions chkWithdrawn
    withdrawnpermision = permision
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub chkWithdrawn_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    chkWithdrawn_Click
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub cmdadd_Click()
On Error GoTo errFix
    action = "addingRecords"
    manuallychecked = True
    resetPermision
    lvwSummary.Visible = False
    fraAllContents.Visible = True
    Set Rst3 = oSaccoMaster.GetRecordSet("select * from usergrps where groupid= '" & txtGroupID.Text & "'")
    If Rst3.RecordCount > 0 Then
        MyBookMark = Rst3.Bookmark
    End If
        Toolbar1.Buttons("bSearch").Enabled = True
        Toolbar1.Buttons("bView").Enabled = True
        Dim ctrl As Control
        For Each ctrl In FrmUserGroups
            If (TypeOf ctrl Is TextBox) Then
                ctrl.Locked = False
            ElseIf (TypeOf ctrl Is CheckBox) Then
                If Not ctrl.name = "chkActivate" Then
                    ctrl.Value = vbUnchecked
                 End If
            End If
        Next ctrl
        fraAllContents.Enabled = True
        cmdEdit.Enabled = False
        cmdAdd.Enabled = False
        cmdupdate.Enabled = True
        cmdCancel.Enabled = True
        cmddelete.Enabled = False
        txtGroupID.Text = ""
        txtdescription = ""
        txtGroupID.Enabled = True
        txtdescription.Enabled = True
'        txtNumUsers = ""
'        For Each ctrl In frmSysUserGroups
'        If TypeOf ctrl Is OptionButton Then
'            ctrl.Value = False
'        End If
'        Next ctrl
        Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub cmdAdd_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    cmdadd_Click
    txtGroupID.SetFocus
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub cmdcancel_Click()
On Error GoTo errFix
    resetPermision
    action = ""
    Toolbar1.Buttons("bSearch").Enabled = True
    Toolbar1.Buttons("bView").Enabled = True
    cmdCancel.Enabled = False
    cmddelete.Enabled = True
    cmdAdd.Enabled = True
    cmdEdit.Enabled = True
    cmdupdate.Enabled = False
'    Dim ctrl As Control
'    For Each ctrl In FrmUserGroups
'        If (TypeOf ctrl Is TextBox) Then
'            ctrl.Locked = True
'        ElseIf (TypeOf ctrl Is CheckBox) Then
'            ctrl.Enabled = False
'        End If
'    Next ctrl
    fraAllContents.Enabled = False
    Set rst = oSaccoMaster.GetRecordSet("select * from usergrps order by groupid")
    If rst.RecordCount > 0 Then
        rst.Bookmark = MyBookMark
        txtGroupID.Text = rst!groupid & ""
        load_records
    End If
        Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub cmdCancel_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    cmdcancel_Click
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub cmdclose_Click()
    Unload Me
    action = ""
End Sub

Private Sub cmdClose_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdclose_Click
End If
End Sub

Private Sub cmdDelete_Click()
On Error GoTo errFix
        If lvwSummary.Visible = True Then
            If Not lvwSummary.ListItems.Count = 0 Then
                Sel = lvwSummary.SelectedItem & ""
            Else
                MsgBox "No item selected", vbInformation, "No Items"
                Exit Sub
            End If
        Else
            Sel = txtGroupID.Text
        End If
        Set Rst3 = oSaccoMaster.GetRecordSet("select activate from usergrps where groupid= '" & Sel & "'")
        If Not Rst3.EOF Then
            valToEncrOrDecr = Rst3!Activate & ""
            EncryptPassword
            If EncryptPass = "Mod" Then
                    Set Rs = oSaccoMaster.GetRecordSet("select groupid from users where userid= '" & User & "'")
                    If Not Rs.EOF Then
                        Set Rst4 = oSaccoMaster.GetRecordSet("select activate from usergrps where groupid= '" & Rs!groupid & "'")
                        If Not Rst4.EOF Then
                            valToEncrOrDecr = Rst4!Activate & ""
                            EncryptPassword
                            If Not EncryptPass = "Mod" Then
                                MsgBox "You have no authority to delete this this group.", vbInformation, "Only Vendor Can!"
                                Exit Sub
                            End If
                        End If
                         Rst4.Close
                    End If
                    Rs.Close
            End If
        End If
        Rst3.Close


    If lvwSummary.Visible Then
        If lvwSummary.ListItems.Count = 0 Then
            Exit Sub
        End If
        Sel = lvwSummary.SelectedItem
        Set rst = oSaccoMaster.GetRecordSet("select * from usergrps where groupid= '" & Sel & "'")
        If rst.RecordCount > 0 Then
            If MsgBox("Are you sure you want to delete " & rst!groupid & "" & " ? ", vbYesNo, "User Group deletion") = vbYes Then
                Set Rst5 = oSaccoMaster.GetRecordSet("select * from usergrps where groupid= '" & Sel & "'")
                If Not Rst5.EOF Then
                    Rst5.Delete
                    Rst5.Update
                End If
                load_Summary
            End If
        End If
    Else
        Set rst = oSaccoMaster.GetRecordSet("select * from usergrps where groupid= '" & txtGroupID.Text & "'")
        If rst.RecordCount > 0 Then
            If MsgBox("Are you sure you want to delete " & rst!groupid & "" & " ? ", vbYesNo, "User Group deletion") = vbYes Then
                Set rst = oSaccoMaster.GetRecordSet("select * from usergrps where groupid= '" & txtGroupID.Text & "'")
                If Not rst.EOF Then
                    rst.Delete
                    rst.Update
                End If
                If Not rst.BOF Then
                    rst.MovePrevious
                End If
                Set Rst5 = oSaccoMaster.GetRecordSet("select * from usergrps ")
                    txtGroupID.Text = Rst5!groupid & ""
                load_records
            End If
    End If
    End If
        Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub cmdDelete_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdDelete_Click
End If
End Sub

Private Sub cmdEdit_Click()
        On Error GoTo errFix
    'Refuse user from editing vendor permisions
        If lvwSummary.Visible = True Then
            If Not lvwSummary.ListItems.Count = 0 Then
                Sel = lvwSummary.SelectedItem
            Else
                MsgBox "No item select", vbInformation, "No Items"
                Exit Sub
            End If
        Else
            Sel = txtGroupID.Text
        End If
        Set Rst3 = oSaccoMaster.GetRecordSet("select activate from usergrps where groupid= '" & Sel & "'")
        If Not Rst3.EOF Then
            valToEncrOrDecr = Rst3!Activate & ""
            EncryptPassword
            If EncryptPass = "Mod" Then
                    Set Rs = oSaccoMaster.GetRecordSet("select groupid from users where userid= '" & User & "'")
                    If Not Rs.EOF Then
                        Set Rst4 = oSaccoMaster.GetRecordSet("select activate from usergrps where groupid= '" & Rs!groupid & "'")
                        If Not Rst4.EOF Then
                            valToEncrOrDecr = Rst4!Activate & ""
                            EncryptPassword
                            If Not EncryptPass = "Mod" Then
                                MsgBox "You have no authority to edit this this group.", vbInformation, "Only Vendor Can!"
                                Exit Sub
                            End If
                        End If
                         Rst4.Close
                    End If
                    Rs.Close
            End If
        End If
        Rst3.Close
    
        
        
        
        
        
        
    resetPermision
    action = "editingRecords"
    txtGroupID.Enabled = False
    txtdescription.Enabled = False
    If lvwSummary.Visible Then
        If lvwSummary.ListItems.Count = 0 Then
            Exit Sub
        End If
        Sel = lvwSummary.SelectedItem
        Set Rst3 = oSaccoMaster.GetRecordSet("select * from usergrps order by groupid")
        If Rst3.RecordCount > 0 Then
            Rst3.Find ("groupid= '" & Sel & "'")
            MyBookMark = Rst3.Bookmark
            txtGroupID.Text = Rst3!groupid & ""
            cmdNext.Enabled = True
            cmdPrevious.Enabled = True
            cmdFirst.Enabled = True
            cmdLast.Enabled = True
            load_records
        End If
    
    Else
        
        Set Rst3 = oSaccoMaster.GetRecordSet("select * from usergrps order by groupid")
        If Rst3.RecordCount > 0 Then
            Rst3.Find ("groupid= '" & txtGroupID.Text & "'")
            MyBookMark = Rst3.Bookmark
        Else
            Exit Sub
        End If
    End If
    
        Toolbar1.Buttons("bSearch").Enabled = True
        Toolbar1.Buttons("bView").Enabled = False
        cmdEdit.Enabled = False
        cmdAdd.Enabled = False
        cmdCancel.Enabled = True
        cmddelete.Enabled = False
        cmdupdate.Enabled = True
        fraAllContents.Visible = True
        lvwSummary.Visible = False
        fraAllContents.Enabled = True
        txtGroupID.Locked = True
        manuallychecked = True
        If Not txtGroupID.Enabled = False Then
        txtGroupID.SetFocus
        End If
            Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub cmdEdit_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    cmdEdit_Click
    txtdescription.SetFocus
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub cmdFirst_Click()
On Error GoTo errFix
    Set Rst1 = oSaccoMaster.GetRecordSet("select groupid from usergrps order by groupid")
    With Rst1
        If .RecordCount > 0 Then
            .MoveFirst
            txtGroupID.Text = Rst1!groupid & ""
            load_records
            cmdFirst.Enabled = False
            cmdPrevious.Enabled = False
            cmdNext.Enabled = True
            cmdLast.Enabled = True
        End If
    End With
    Rst1.Close
    action = ""
        Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub
Private Sub cmdFirst_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    cmdFirst_Click
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub cmdHelp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    'cmdHelp_Click
End If
End Sub

Private Sub cmdLast_Click()
On Error GoTo errFix
    Set Rst1 = oSaccoMaster.GetRecordSet("select groupid from usergrps order by groupid")
    With Rst1
        If .RecordCount > 0 Then
            .MoveLast
            txtGroupID.Text = !groupid & ""
            cmdFirst.Enabled = True
            cmdPrevious.Enabled = True
            cmdNext.Enabled = False
            cmdLast.Enabled = False
            load_records
        End If
    End With
    Rst1.Close
    action = ""
        Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub cmdLast_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdLast_Click
End If
End Sub

Private Sub cmdNext_Click()
On Error GoTo errFix
    Set Rst1 = oSaccoMaster.GetRecordSet("select groupid from usergrps order by groupid")
    If cmdupdate.Enabled = True Then
        Toolbar1.Buttons("bSearch").Enabled = True
        Toolbar1.Buttons("bView").Enabled = True
        Rst1.Bookmark = MyBookMark
        txtGroupID.Text = Rst1!groupid & ""
    End If
    With Rst1
        If .RecordCount > 0 Then
            .MoveFirst
            .Find ("groupid= '" & txtGroupID.Text & "'")
            If Not .EOF Then
                .MoveNext
                If .EOF Then
                    .MoveLast
                    cmdFirst.Enabled = True
                    cmdPrevious.Enabled = True
                    cmdNext.Enabled = False
                    cmdLast.Enabled = False
                Else
                    cmdFirst.Enabled = True
                    cmdPrevious.Enabled = True
                    cmdNext.Enabled = True
                    cmdLast.Enabled = True
                End If
                txtGroupID.Text = !groupid & ""
                load_records
            End If
        End If
    End With
    Rst1.Close
    action = ""
        Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub cmdNext_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdNext_Click
End If
End Sub

Private Sub cmdPrevious_Click()
On Error GoTo errFix
    Set Rst1 = oSaccoMaster.GetRecordSet("select groupid from usergrps order by groupid")
    If cmdupdate.Enabled = True Then
        Toolbar1.Buttons("bSearch").Enabled = True
        Toolbar1.Buttons("bView").Enabled = True
        Rst1.Bookmark = MyBookMark
        txtGroupID.Text = Rst1!groupid & ""
    End If
    With Rst1
        If .RecordCount > 0 Then
            .MoveFirst
            .Find ("groupid= '" & txtGroupID.Text & "'")
            If Not .EOF Then
                .MovePrevious
                If .BOF Then
                    .MoveFirst
                    cmdFirst.Enabled = False
                    cmdPrevious.Enabled = False
                    cmdNext.Enabled = True
                    cmdLast.Enabled = True
                Else
                    cmdFirst.Enabled = True
                    cmdPrevious.Enabled = True
                    cmdNext.Enabled = True
                    cmdLast.Enabled = True
                End If
              txtGroupID.Text = !groupid & ""
              load_records
            End If
        End If
    End With
    Rst1.Close
    action = ""
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub cmdPrevious_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdPrevious_Click
End If
End Sub

Private Sub cmdPrev_Click()

End Sub

Private Sub cmdupdate_Click()
On Error GoTo errFix
        While txtGroupID.Text = ""
            MsgBox "You must input a group ID", vbInformation, "User ID"
            Exit Sub
        Wend
        While txtdescription.Text = ""
            MsgBox "You must input a group name", vbOKOnly, "User Name"
            Exit Sub
        Wend
    findNonChecked
    If action = "addingRecords" Then
        Set Rst5 = oSaccoMaster.GetRecordSet("select groupid from usergrps where groupid= '" & txtGroupID.Text & "'")
        If Not Rst5.EOF Then
            MsgBox "User group with same ID exists", vbInformation, "User Groups"
            txtGroupID.Text = ""
            txtGroupID.SetFocus
            Exit Sub
        End If
        fraAllContents.Enabled = False
        Set rst = oSaccoMaster.GetRecordSet("select * from usergrps")
        With rst
            .AddNew
            rst!groupid = txtGroupID.Text & ""
            rst!description = txtdescription.Text & ""
            
            If chkMembership.Value = vbChecked Then
                valToEncrOrDecr = "Yes"
                Call EncryptPassword
                rst!rwmembers = EncryptPass
            Else
                valToEncrOrDecr = "No"
                Call EncryptPassword
                rst!rwmembers = EncryptPass
            End If
            
            If chkLoans.Value = vbChecked Then
                valToEncrOrDecr = "Yes"
                Call EncryptPassword
                rst!RWLoans = EncryptPass
            Else
                valToEncrOrDecr = "No"
                Call EncryptPassword
                rst!RWLoans = EncryptPass
            End If
            
            If chkShares.Value = vbChecked Then
                 valToEncrOrDecr = "Yes"
                Call EncryptPassword
                rst!RWShares = EncryptPass
            Else
                valToEncrOrDecr = "No"
                Call EncryptPassword
                rst!RWShares = EncryptPass
            End If
            
            If chkUtilities.Value = vbChecked Then
                valToEncrOrDecr = "Yes"
                Call EncryptPassword
                rst!RWUtilities = EncryptPass
            Else
                valToEncrOrDecr = "No"
                Call EncryptPassword
                rst!RWUtilities = EncryptPass
            End If
            
            If chkSetup.Value = vbChecked Then
                valToEncrOrDecr = "Yes"
                Call EncryptPassword
                rst!RWSetup = EncryptPass
            Else
                valToEncrOrDecr = "No"
                Call EncryptPassword
                rst!RWSetup = EncryptPass
            End If
            
            If chkOtherSchemes.Value = vbChecked Then
                valToEncrOrDecr = "Yes"
                Call EncryptPassword
                rst!RWOtherSchemes = EncryptPass
            Else
                valToEncrOrDecr = "No"
                Call EncryptPassword
                rst!RWOtherSchemes = EncryptPass
            End If
             
            If MemberRegPermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!memreg = EncryptPass
            ElseIf MemberRegPermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!memreg = EncryptPass
            ElseIf MemberRegPermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!memreg = EncryptPass
            End If
            
            If NextOfKinPermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!nok = EncryptPass
            ElseIf MemberRegPermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!nok = EncryptPass
            ElseIf MemberRegPermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!nok = EncryptPass
            End If
            If memstatementspermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!memstatement = EncryptPass
            ElseIf memstatementspermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!memstatement = EncryptPass
            ElseIf memstatementspermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!memstatement = EncryptPass
            End If
            If savingspermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!savings = EncryptPass
            ElseIf savingspermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!savings = EncryptPass
            ElseIf savingspermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!savings = EncryptPass
            End If
            If benfundpermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!benevolentfund = EncryptPass
            ElseIf benfundpermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!benevolentfund = EncryptPass
            ElseIf benfundpermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!benevolentfund = EncryptPass
            End If
            If applicationpermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!loanapplic = EncryptPass
            ElseIf applicationpermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!loanapplic = EncryptPass
            ElseIf applicationpermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!loanapplic = EncryptPass
            End If
            If loanguarantorspermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!loanguarantors = EncryptPass
            ElseIf loanguarantorspermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!loanguarantors = EncryptPass
            ElseIf loanguarantorspermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!loanguarantors = EncryptPass
            End If
            
            If loanbalpermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!loanbal = EncryptPass
            ElseIf loanbalpermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!loanbal = EncryptPass
            ElseIf loanbalpermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!loanbal = EncryptPass
            End If
            
            If activatepermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!Activate = EncryptPass
            ElseIf activatepermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!Activate = EncryptPass
            ElseIf activatepermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!Activate = EncryptPass
            End If
            If loanendorsementpermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!loanendorsement = EncryptPass
            ElseIf loanguarantorspermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!loanendorsement = EncryptPass
            ElseIf loanguarantorspermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!loanendorsement = EncryptPass
            End If
            If loanbalpermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!loanbal = EncryptPass
            ElseIf loanbalpermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!loanbal = EncryptPass
            ElseIf loanbalpermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!loanbal = EncryptPass
            End If
            
            If chequeentrypermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!chequeentry = EncryptPass
            ElseIf chequeentrypermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!chequeentry = EncryptPass
            ElseIf chequeentrypermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!chequeentry = EncryptPass
            End If
            If transactionspermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!loantrans = EncryptPass
            ElseIf transactionspermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!loantrans = EncryptPass
            ElseIf transactionspermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!loantrans = EncryptPass
            End If
            If contributionspermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!Contributions = EncryptPass
            ElseIf contributionspermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!Contributions = EncryptPass
            ElseIf contributionspermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!Contributions = EncryptPass
            End If
            If sharevarpermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!sharevar = EncryptPass
            ElseIf sharevarpermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!sharevar = EncryptPass
            ElseIf sharevarpermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!sharevar = EncryptPass
            End If
            
            If exporttoglpermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!exporttogl = EncryptPass
            ElseIf exporttoglpermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!exporttogl = EncryptPass
            ElseIf exporttoglpermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!exporttogl = EncryptPass
            End If
            If calculatorpermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!calc = EncryptPass
            ElseIf calculatorpermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!calc = EncryptPass
            ElseIf calculatorpermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!calc = EncryptPass
            End If
            If effectrepaypermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!effectrepayment = EncryptPass
            ElseIf effectrepaypermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!effectrepayment = EncryptPass
            ElseIf effectrepaypermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!effectrepayment = EncryptPass
            End If
            If dividendspermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!dividends = EncryptPass
            ElseIf dividendspermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!dividends = EncryptPass
            ElseIf dividendspermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!dividends = EncryptPass
            End If
            If withdrawnpermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!withdrawn = EncryptPass
            ElseIf withdrawnpermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!withdrawn = EncryptPass
            ElseIf withdrawnpermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!withdrawn = EncryptPass
            End If
            If deductionspermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!deductions = EncryptPass
            ElseIf deductionspermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!deductions = EncryptPass
            ElseIf deductionspermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!deductions = EncryptPass
            End If
            If utilstatementpermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                 rst!utilstatements = EncryptPass
            ElseIf utilstatementpermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!utilstatements = EncryptPass
            ElseIf utilstatementpermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!utilstatements = EncryptPass
            End If
            If periodtranpermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!periodictran = EncryptPass
            ElseIf periodtranpermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!periodictran = EncryptPass
            ElseIf periodtranpermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!periodictran = EncryptPass
            End If
            If monthlydedpermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!monthlydeductions = EncryptPass
            ElseIf monthlydedpermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!monthlydeductions = EncryptPass
            ElseIf monthlydedpermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!monthlydeductions = EncryptPass
            End If
             If utilguarantorspermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!utilguarantor = EncryptPass
            ElseIf utilguarantorspermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!utilguarantor = EncryptPass
            ElseIf utilguarantorspermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!utilguarantor = EncryptPass
            End If
            If dormantpermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!dormant = EncryptPass
            ElseIf dormantpermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!dormant = EncryptPass
            ElseIf dormantpermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!dormant = EncryptPass
            End If
            If backuppermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!Backup = EncryptPass
            ElseIf backuppermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!Backup = EncryptPass
            ElseIf backuppermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!Backup = EncryptPass
            End If
            If parampermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!parametization = EncryptPass
            ElseIf parampermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!parametization = EncryptPass
            ElseIf parampermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!parametization = EncryptPass
            End If
            If loantypespermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!loantypes = EncryptPass
            ElseIf loantypespermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!loantypes = EncryptPass
            ElseIf loantypespermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!loantypes = EncryptPass
            End If
            If banksetuppermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!banksetup = EncryptPass
            ElseIf banksetuppermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!banksetup = EncryptPass
            ElseIf banksetuppermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!banksetup = EncryptPass
            End If
            If clearmempermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!clearMemRecs = EncryptPass
            ElseIf clearmempermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!clearMemRecs = EncryptPass
            ElseIf clearmempermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!clearMemRecs = EncryptPass
            End If
            If clearloanpermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!clearLoanRecs = EncryptPass
            ElseIf clearloanpermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!clearLoanRecs = EncryptPass
            ElseIf clearloanpermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!clearLoanRecs = EncryptPass
            End If
            If databasesetpermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!databasesetup = EncryptPass
            ElseIf databasesetpermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!databasesetup = EncryptPass
            ElseIf databasesetpermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!databasesetup = EncryptPass
            End If
            If archivedpermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!Archived = EncryptPass
            ElseIf databasesetpermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!Archived = EncryptPass
            ElseIf databasesetpermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!Archived = EncryptPass
            End If
            If rejreasonpermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!rejreasons = EncryptPass
            ElseIf rejreasonpermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!rejreasons = EncryptPass
            ElseIf rejreasonpermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!rejreasons = EncryptPass
            End If
            If usergroupspermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!usergrps = EncryptPass
            ElseIf usergroupspermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!usergrps = EncryptPass
            ElseIf usergroupspermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!usergrps = EncryptPass
            End If
            If sysuserspermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!sysusers = EncryptPass
            ElseIf sysuserspermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!sysusers = EncryptPass
            ElseIf sysuserspermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!sysusers = EncryptPass
            End If
            If companysetpermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!companysetup = EncryptPass
            ElseIf companysetpermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!companysetup = EncryptPass
            ElseIf companysetpermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!companysetup = EncryptPass
            End If
            If changenopermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!changememno = EncryptPass
            ElseIf changenopermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!changememno = EncryptPass
            ElseIf changenopermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!changememno = EncryptPass
            End If
            rst!auditid = User
            rst!audittime = Format(Now, "DD/MM/YYYY hh:mm:ss")
            cmdAdd.Enabled = True
            cmdupdate.Enabled = False
            cmdEdit.Enabled = True
            cmdCancel.Enabled = False
            cmdEdit.Enabled = True
            cmddelete.Enabled = True
            Toolbar1.Buttons("bSearch").Enabled = True
            Toolbar1.Buttons("bView").Enabled = True
            rst.Update
        End With
    ElseIf action = "editingRecords" Then
        fraAllContents.Enabled = False
        Set rst = oSaccoMaster.GetRecordSet("select * from usergrps where groupid= '" & txtGroupID.Text & "'")
        If rst.RecordCount > 0 Then
        With rst
            rst!description = txtdescription.Text & ""
            
            If chkMembership.Value = vbChecked Then
                valToEncrOrDecr = "Yes"
                Call EncryptPassword
                rst!rwmembers = EncryptPass
            Else
                valToEncrOrDecr = "No"
                Call EncryptPassword
                rst!rwmembers = EncryptPass
            End If
            
            If chkLoans.Value = vbChecked Then
                valToEncrOrDecr = "Yes"
                Call EncryptPassword
                rst!RWLoans = EncryptPass
            Else
                valToEncrOrDecr = "No"
                Call EncryptPassword
                rst!RWLoans = EncryptPass
            End If
            
            If chkShares.Value = vbChecked Then
                 valToEncrOrDecr = "Yes"
                Call EncryptPassword
                rst!RWShares = EncryptPass
            Else
                valToEncrOrDecr = "No"
                Call EncryptPassword
                rst!RWShares = EncryptPass
            End If
            
            If chkUtilities.Value = vbChecked Then
                valToEncrOrDecr = "Yes"
                Call EncryptPassword
                rst!RWUtilities = EncryptPass
            Else
                valToEncrOrDecr = "No"
                Call EncryptPassword
                rst!RWUtilities = EncryptPass
            End If
            
            If chkSetup.Value = vbChecked Then
                valToEncrOrDecr = "Yes"
                Call EncryptPassword
                rst!RWSetup = EncryptPass
            Else
                valToEncrOrDecr = "No"
                Call EncryptPassword
                rst!RWSetup = EncryptPass
            End If
            
            If chkOtherSchemes.Value = vbChecked Then
                valToEncrOrDecr = "Yes"
                Call EncryptPassword
                rst!RWOtherSchemes = EncryptPass
            Else
                valToEncrOrDecr = "No"
                Call EncryptPassword
                rst!RWOtherSchemes = EncryptPass
            End If
             
            If MemberRegPermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!memreg = EncryptPass
            ElseIf MemberRegPermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!memreg = EncryptPass
            ElseIf MemberRegPermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!memreg = EncryptPass
            End If
            
            If NextOfKinPermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!nok = EncryptPass
            ElseIf MemberRegPermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!nok = EncryptPass
            ElseIf MemberRegPermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!nok = EncryptPass
            End If
            If memstatementspermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!memstatement = EncryptPass
            ElseIf memstatementspermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!memstatement = EncryptPass
            ElseIf memstatementspermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!memstatement = EncryptPass
            End If
            If savingspermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!savings = EncryptPass
            ElseIf savingspermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!savings = EncryptPass
            ElseIf savingspermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!savings = EncryptPass
            End If
            If benfundpermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!benevolentfund = EncryptPass
            ElseIf benfundpermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!benevolentfund = EncryptPass
            ElseIf benfundpermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!benevolentfund = EncryptPass
            End If
            If applicationpermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!loanapplic = EncryptPass
            ElseIf applicationpermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!loanapplic = EncryptPass
            ElseIf applicationpermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!loanapplic = EncryptPass
            End If
            If activatepermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!Activate = EncryptPass
            ElseIf activatepermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!Activate = EncryptPass
            ElseIf activatepermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!Activate = EncryptPass
            End If
            If loanguarantorspermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!loanguarantors = EncryptPass
            ElseIf loanguarantorspermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!loanguarantors = EncryptPass
            ElseIf loanguarantorspermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!loanguarantors = EncryptPass
            End If
            
            If loanbalpermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!loanbal = EncryptPass
            ElseIf loanbalpermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!loanbal = EncryptPass
            ElseIf loanbalpermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!loanbal = EncryptPass
            End If
            If loanendorsementpermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!loanendorsement = EncryptPass
            ElseIf loanguarantorspermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!loanendorsement = EncryptPass
            ElseIf loanguarantorspermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!loanendorsement = EncryptPass
            End If
            If loanbalpermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!loanbal = EncryptPass
            ElseIf loanbalpermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!loanbal = EncryptPass
            ElseIf loanbalpermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!loanbal = EncryptPass
            End If
            
            If chequeentrypermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!chequeentry = EncryptPass
            ElseIf chequeentrypermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!chequeentry = EncryptPass
            ElseIf chequeentrypermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!chequeentry = EncryptPass
            End If
            If transactionspermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!loantrans = EncryptPass
            ElseIf transactionspermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!loantrans = EncryptPass
            ElseIf transactionspermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!loantrans = EncryptPass
            End If
            If contributionspermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!Contributions = EncryptPass
            ElseIf contributionspermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!Contributions = EncryptPass
            ElseIf contributionspermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!Contributions = EncryptPass
            End If
            If sharevarpermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!sharevar = EncryptPass
            ElseIf sharevarpermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!sharevar = EncryptPass
            ElseIf sharevarpermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!sharevar = EncryptPass
            End If
            If backuppermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!Backup = EncryptPass
            ElseIf backuppermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!Backup = EncryptPass
            ElseIf backuppermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!Backup = EncryptPass
            End If
            If exporttoglpermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!exporttogl = EncryptPass
            ElseIf exporttoglpermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!exporttogl = EncryptPass
            ElseIf exporttoglpermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!exporttogl = EncryptPass
            End If
            If calculatorpermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!calc = EncryptPass
            ElseIf calculatorpermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!calc = EncryptPass
            ElseIf calculatorpermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!calc = EncryptPass
            End If
            If effectrepaypermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!effectrepayment = EncryptPass
            ElseIf effectrepaypermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!effectrepayment = EncryptPass
            ElseIf effectrepaypermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!effectrepayment = EncryptPass
            End If
            If dividendspermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!dividends = EncryptPass
            ElseIf dividendspermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!dividends = EncryptPass
            ElseIf dividendspermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!dividends = EncryptPass
            End If
            If withdrawnpermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!withdrawn = EncryptPass
            ElseIf withdrawnpermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!withdrawn = EncryptPass
            ElseIf withdrawnpermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!withdrawn = EncryptPass
            End If
            If deductionspermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!deductions = EncryptPass
            ElseIf deductionspermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!deductions = EncryptPass
            ElseIf deductionspermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!deductions = EncryptPass
            End If
            If utilstatementpermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                 rst!utilstatements = EncryptPass
            ElseIf utilstatementpermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!utilstatements = EncryptPass
            ElseIf utilstatementpermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!utilstatements = EncryptPass
            End If
            If periodtranpermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!periodictran = EncryptPass
            ElseIf periodtranpermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!periodictran = EncryptPass
            ElseIf periodtranpermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!periodictran = EncryptPass
            End If
            If monthlydedpermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!monthlydeductions = EncryptPass
            ElseIf monthlydedpermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!monthlydeductions = EncryptPass
            ElseIf monthlydedpermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!monthlydeductions = EncryptPass
            End If
             If utilguarantorspermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!utilguarantor = EncryptPass
            ElseIf utilguarantorspermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!utilguarantor = EncryptPass
            ElseIf utilguarantorspermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!utilguarantor = EncryptPass
            End If
            If dormantpermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!dormant = EncryptPass
            ElseIf dormantpermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!dormant = EncryptPass
            ElseIf dormantpermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!dormant = EncryptPass
            End If
            If parampermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!parametization = EncryptPass
            ElseIf parampermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!parametization = EncryptPass
            ElseIf parampermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!parametization = EncryptPass
            End If
            If loantypespermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!loantypes = EncryptPass
            ElseIf loantypespermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!loantypes = EncryptPass
            ElseIf loantypespermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!loantypes = EncryptPass
            End If
            If banksetuppermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!banksetup = EncryptPass
            ElseIf banksetuppermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!banksetup = EncryptPass
            ElseIf banksetuppermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!banksetup = EncryptPass
            End If
            If clearmempermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!clearMemRecs = EncryptPass
            ElseIf clearmempermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!clearMemRecs = EncryptPass
            ElseIf clearmempermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!clearMemRecs = EncryptPass
            End If
            If clearloanpermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!clearLoanRecs = EncryptPass
            ElseIf clearloanpermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!clearLoanRecs = EncryptPass
            ElseIf clearloanpermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!clearLoanRecs = EncryptPass
            End If
            If databasesetpermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!databasesetup = EncryptPass
            ElseIf databasesetpermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!databasesetup = EncryptPass
            ElseIf databasesetpermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!databasesetup = EncryptPass
            End If
            If archivedpermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!Archived = EncryptPass
            ElseIf databasesetpermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!Archived = EncryptPass
            ElseIf databasesetpermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!Archived = EncryptPass
            End If
            If rejreasonpermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!rejreasons = EncryptPass
            ElseIf rejreasonpermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!rejreasons = EncryptPass
            ElseIf rejreasonpermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!rejreasons = EncryptPass
            End If
            If usergroupspermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!usergrps = EncryptPass
            ElseIf usergroupspermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!usergrps = EncryptPass
            ElseIf usergroupspermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!usergrps = EncryptPass
            End If
            If sysuserspermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!sysusers = EncryptPass
            ElseIf sysuserspermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!sysusers = EncryptPass
            ElseIf sysuserspermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!sysusers = EncryptPass
            End If
            If companysetpermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!companysetup = EncryptPass
            ElseIf companysetpermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!companysetup = EncryptPass
            ElseIf companysetpermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!companysetup = EncryptPass
            End If
            If changenopermision = "Mod" Then
                valToEncrOrDecr = "Mod"
                Call EncryptPassword
                rst!changememno = EncryptPass
            ElseIf changenopermision = "View" Then
                valToEncrOrDecr = "View"
                Call EncryptPassword
                rst!changememno = EncryptPass
            ElseIf changenopermision = "Non" Then
                valToEncrOrDecr = "Non"
                Call EncryptPassword
                rst!changememno = EncryptPass
            End If
            rst!auditid = User
            rst!audittime = Format(Now, "DD/MM/YYYY hh:mm:ss")
            cmdAdd.Enabled = True
            cmdupdate.Enabled = False
            cmdEdit.Enabled = True
            cmdCancel.Enabled = False
            cmdEdit.Enabled = True
            cmddelete.Enabled = True
            Toolbar1.Buttons("bSearch").Enabled = True
            Toolbar1.Buttons("bView").Enabled = True
            rst.Update
        End With
        
        End If
    End If
    action = ""
    resetPermision
        Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub cmdUpdate_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    cmdupdate_Click
    txtGroupID.SetFocus
End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub

Private Sub Form_Activate()
    PositionForm Me
'    Set rst2 = oSaccomaster.GetRecordSet("select groupid from users where groupid= '" & txtGroupID.Text & "'")
'    If rst2.RecordCount > 0 Then
'        txtNumUsers.Text = rst2.RecordCount
'    End If
End Sub
Private Sub Form_Load()
On Error GoTo errFix
    Toolbar1.Buttons.Add 1, "bSearch", "&Search", 0, 1
    Toolbar1.Buttons.Add 2, "bView", "&View", 5, 2
    Toolbar1.Buttons("bView").ButtonMenus.Add 1, "mDetails", "Details"
    Toolbar1.Buttons("bView").ButtonMenus.Add 2, "mSummary", "Summary"
    fraAllContents.Enabled = False
    fraMembers.Enabled = False
    fraShares.Enabled = False
    fraLoans.Enabled = False
    fraOtherSchemes.Enabled = False
    fraUtilities.Enabled = False
    fraSetup.Enabled = False
    Set rst = oSaccoMaster.GetRecordSet("select * from usergrps order by groupid")
    If rst.RecordCount > 0 Then
        rst.MoveFirst
        txtGroupID.Text = rst!groupid & ""
        load_records
    Else
        txtGroupID.Enabled = False
        txtdescription.Enabled = False
        If txtGroupID.Text = "" Then
            chkMembership.Enabled = False
            chkShares.Enabled = False
            chkLoans.Enabled = False
            chkSetup.Enabled = False
            chkOtherSchemes.Enabled = False
            chkUtilities.Enabled = False
        Else
            chkMembership.Enabled = True
            chkShares.Enabled = True
            chkLoans.Enabled = True
            chkSetup.Enabled = True
            chkOtherSchemes.Enabled = True
            chkUtilities.Enabled = True
        End If
        manuallychecked = True
    End If
    cmdCancel.Enabled = False
    cmdupdate.Enabled = False
   Set Rst4 = oSaccoMaster.GetRecordSet("select groupid from users where userid= '" & User & "'")
    If Not Rst4.EOF Then
        Set Rst5 = oSaccoMaster.GetRecordSet("select * from usergrps where groupid= '" & Rst4!groupid & "'")
        If Not Rst5.EOF Then
            valToEncrOrDecr = Rst5!usergrps
            EncryptPassword
            If EncryptPass = "View" Then
                disablemodifying = True
                cmdAdd.Enabled = False
                cmdupdate.Enabled = False
                cmdCancel.Enabled = False
                cmddelete.Enabled = False
                cmdEdit.Enabled = False
            Else
                disablemodifying = False
            End If
             valToEncrOrDecr = Rst5!Activate
            EncryptPassword
            If EncryptPass = "View" Or EncryptPass = "Mod" Then
                chkActivate.Enabled = True
            Else
                chkActivate.Enabled = False
            End If
        
        End If
    Else
            chkActivate.Enabled = False '  VERY VERY IMPORTANT. IF ONE DELETES ALL USER GROUPS STILL HE CANT BE ABLE TO ACTIVATE.
    End If
        Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub
Private Sub load_Summary()
On Error GoTo errFix
    cmdNext.Enabled = False
    cmdPrevious.Enabled = False
    cmdLast.Enabled = False
    cmdFirst.Enabled = False
    With lvwSummary
            .ListItems.Clear
            .Columnheaders.Clear
            .Columnheaders.Add 1, , "Group ID", 2000
            .Columnheaders.Add 2, , "Description", 2000
            .Columnheaders.Add 3, , "Membership", 3000
            .Columnheaders.Add 4, , "Shares", 3000
            .Columnheaders.Add 5, , "Loans", 2000
            .Columnheaders.Add 6, , "Utilities", 2000
            .Columnheaders.Add 7, , "Setup", 2000
            .Columnheaders.Add 8, , "OtherSchemes", 3000
            .Columnheaders.Add 9, , "Registration", 3000
            .Columnheaders.Add 10, , "NextOfKin", 3000
            .Columnheaders.Add 11, , "MemberStatement", 3000
            .Columnheaders.Add 12, , "LoanApp", 2000
            .Columnheaders.Add 13, , "LoanEndordement", 3000
            .Columnheaders.Add 14, , "LoanBal", 3000
            .Columnheaders.Add 15, , "LoanTransaction", 3000
            .Columnheaders.Add 16, , "EffectRepay", 3000
            .Columnheaders.Add 17, , "LoanGuar", 3000
            .Columnheaders.Add 18, , "ChequeEntry", 3000
            .Columnheaders.Add 19, , "Calc", 1000
            .Columnheaders.Add 20, , "Dividends", 2000
            .Columnheaders.Add 21, , "Deductions", 2000
            .Columnheaders.Add 22, , "PeriodicTran", 3000
            .Columnheaders.Add 23, , "UtilStatement", 2000
            .Columnheaders.Add 24, , "MonthlyDeduct.", 3000
            .Columnheaders.Add 25, , "ExportToGL", 2000
            .Columnheaders.Add 26, , "UtilGuar", 3000
            .Columnheaders.Add 27, , "Dormant", 2000
            .Columnheaders.Add 28, , "Archived", 3000
            .Columnheaders.Add 29, , "Withdrawn", 2000
            .Columnheaders.Add 30, , "Backup", 2000
            .Columnheaders.Add 31, , "Parametization", 3000
            .Columnheaders.Add 32, , "CompanySet.", 2000
            .Columnheaders.Add 33, , "BankSet", 2000
            .Columnheaders.Add 34, , "DatabaseSetup", 2000
            .Columnheaders.Add 35, , "Activation", 2000
            .Columnheaders.Add 36, , "RejReasons", 2000
            .Columnheaders.Add 37, , "UserGrps", 2000
            .Columnheaders.Add 38, , "SysUsers", 2000
            .Columnheaders.Add 39, , "ClearLoanRecs", 2000
            .Columnheaders.Add 40, , "ClearMemRecs", 2000
            .Columnheaders.Add 41, , "ChangeMemNo", 2000
            .Columnheaders.Add 42, , "Contrib", 2000
            .Columnheaders.Add 43, , "ShareVar", 2000
            .Columnheaders.Add 44, , "Savings", 2000
            .Columnheaders.Add 45, , "LoanTypes", 2000
            .Columnheaders.Add 46, , "BenFund", 2000
            .View = lvwReport
            .Gridlines = True
    End With
    lvwSummary.ListItems.Clear
        strSQL = "Select * from usergrps order by groupid"
        Set rst = oSaccoMaster.GetRecordSet(strSQL)
        With rst
            If .RecordCount > 0 Then
                .MoveFirst
                Do While Not .EOF
                    Set li = lvwSummary.ListItems.Add(, , !groupid & "")
                    li.ListSubItems.Add , , !description & ""
                    li.ListSubItems.Add , , !rwmembers & ""
                    li.ListSubItems.Add , , !RWShares & ""
                    li.ListSubItems.Add , , !RWLoans & ""
                    li.ListSubItems.Add , , !RWUtilities & ""
                    li.ListSubItems.Add , , !RWSetup & ""
                    li.ListSubItems.Add , , !RWOtherSchemes & ""
                    li.ListSubItems.Add , , !memreg & ""
                    li.ListSubItems.Add , , !nok & ""
                    li.ListSubItems.Add , , !memstatement & ""
                    li.ListSubItems.Add , , !loanapplic & ""
                    li.ListSubItems.Add , , !loanendorsement & ""
                    li.ListSubItems.Add , , !loanbal & ""
                    li.ListSubItems.Add , , !loantrans & ""
                    li.ListSubItems.Add , , !effectrepayment & ""
                    li.ListSubItems.Add , , !loanguarantors & ""
                    li.ListSubItems.Add , , !chequeentry & ""
                    li.ListSubItems.Add , , !calc & ""
                    li.ListSubItems.Add , , !dividends & ""
                    li.ListSubItems.Add , , !deductions & ""
                    li.ListSubItems.Add , , !periodictran & ""
                    li.ListSubItems.Add , , !utilstatements & ""
                    li.ListSubItems.Add , , !monthlydeductions & ""
                    li.ListSubItems.Add , , !exporttogl & ""
                    li.ListSubItems.Add , , !utilguarantor & ""
                    li.ListSubItems.Add , , !dormant & ""
                    li.ListSubItems.Add , , !Archived & ""
                    li.ListSubItems.Add , , !withdrawn & ""
                    li.ListSubItems.Add , , !Backup & ""
                    li.ListSubItems.Add , , !parametization & ""
                    li.ListSubItems.Add , , rst!companysetup & ""
                    li.ListSubItems.Add , , !banksetup & ""
                    li.ListSubItems.Add , , !databasesetup & ""
                    li.ListSubItems.Add , , !Activate & ""
                    li.ListSubItems.Add , , !rejreasons & ""
                    li.ListSubItems.Add , , !usergrps & ""
                    li.ListSubItems.Add , , !sysusers & ""
                    li.ListSubItems.Add , , !clearLoanRecs & ""
                    li.ListSubItems.Add , , !clearMemRecs & ""
                    li.ListSubItems.Add , , !changememno & ""
                    li.ListSubItems.Add , , !Contributions & ""
                    li.ListSubItems.Add , , !sharevar & ""
                    li.ListSubItems.Add , , !savings & ""
                    li.ListSubItems.Add , , !loantypes & ""
                    li.ListSubItems.Add , , !benevolentfund & ""
                    .MoveNext
                Loop
            End If
            .Close
        End With
        Set rst = Nothing
Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Sub
Private Sub findNonChecked()
On Error GoTo errFix
            If chkMemberReg.Value = vbUnchecked Then
                MemberRegPermision = "Non"
            End If
            If chkShareVar.Value = vbUnchecked Then
                sharevarpermision = "Non"
            End If
            If chkSysUsers.Value = vbUnchecked Then
               sysuserspermision = "Non"
            End If
            If chkNextOfKin.Value = vbUnchecked Then
                NextOfKinPermision = "Non"
            End If
            If chkBackup.Value = vbUnchecked Then
                backuppermision = "Non"
            End If
            If chkCompanySetup.Value = vbUnchecked Then
                companysetpermision = "Non"
            End If
            If chkMemStatement.Value = vbUnchecked Then
                memstatementspermision = "Non"
            End If
            If chkSavings.Value = vbUnchecked Then
               savingspermision = "Non"
            End If
            
            If chkBenevolentFund.Value = vbUnchecked Then
               benfundpermision = "Non"
            End If
            If chkEffectRepay.Value = vbUnchecked Then
              effectrepaypermision = "Non"
            End If
            If chkLoanApp.Value = vbUnchecked Then
              applicationpermision = "Non"
            End If
             If chkLoanGuar.Value = vbUnchecked Then
              loanguarantorspermision = "Non"
             End If
            
            If chkLoanEndorsement.Value = vbUnchecked Then
              loanendorsementpermision = "Non"
             End If
            If chkLoanBal.Value = vbUnchecked Then
              loanbalpermision = "Non"
             End If
            If chkChequeEntry.Value = vbUnchecked Then
              chequeentrypermision = "Non"
             End If
            If chktransactions.Value = vbUnchecked Then
              transactionspermision = "Non"
             End If
            If chkContributions.Value = vbUnchecked Then
              contributionspermision = "Non"
             End If
            If chkShareVar.Value = vbUnchecked Then
              contributionspermision = "Non"
             End If
             
            If chkExportToGL.Value = vbUnchecked Then
              exporttoglpermision = "Non"
             End If
            If chkCalc.Value = vbUnchecked Then
              calculatorpermision = "Non"
             End If
            If chkDividends.Value = vbUnchecked Then
              dividendspermision = "Non"
             End If
            If chkWithdrawn.Value = vbUnchecked Then
              withdrawnpermision = "Non"
             End If
            If chkDeductions.Value = vbUnchecked Then
              deductionspermision = "Non"
             End If
            If chkUtilStatements.Value = vbUnchecked Then
                utilstatementpermision = "Non"
             End If
            If chkPeriodicTran.Value = vbUnchecked Then
                periodtranpermision = "Non"
             End If
            If chkMonthlyDeductions.Value = vbUnchecked Then
                monthlydedpermision = "Non"
             End If
             If chkUtilGuar.Value = vbUnchecked Then
                utilguarantorspermision = "Non"
             End If
            If chkDormant.Value = vbUnchecked Then
                dormantpermision = "Non"
             End If
             If chkBackup.Value = vbUnchecked Then
                backuppermision = "Non"
             End If
            If chkParametization.Value = vbUnchecked Then
                parampermision = "Non"
             End If
            If chkLoanType.Value = vbUnchecked Then
                loantypespermision = "Non"
             End If
            If chkBankSetup.Value = vbUnchecked Then
                banksetuppermision = "Non"
             End If
            If chkClearMemRecords.Value = vbUnchecked Then
                clearmempermision = "Non"
             End If
            If chkClearLoanRecords.Value = vbUnchecked Then
                clearloanpermision = "Non"
             End If
            If chkDatabaseSetup.Value = vbUnchecked Then
                databasesetpermision = "Non"
             End If
             If chkActivate.Value = vbUnchecked Then
                activatepermision = "Non"
             End If
            If chkRejReasons.Value = vbUnchecked Then
                rejreasonpermision = "Non"
             End If
            If chkUserGrps.Value = vbUnchecked Then
                usergroupspermision = "Non"
             End If
            If chkChangeMembershipNo.Value = vbUnchecked Then
                changenopermision = "Non"
             End If
             Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"

End Sub


Private Sub resetPermision()
On Error GoTo errFix
    MemberRegPermision = ""
    NextOfKinPermision = ""
    statementpermision = ""
    applicationpermision = ""
    loanguarantorspermision = ""
    loanbalpermision = ""
    loanendorsementpermision = ""
    chequeentrypermision = ""
    effectrepaypermision = ""
    loantypespermision = ""
    transactionspermision = ""
    contributionspermision = ""
    sharevarpermision = ""
    backuppermision = ""
    savingspermision = ""
    benfundpermision = ""
    exporttoglpermision = ""
    dividendspermision = ""
    calculatorpermision = ""
    withdrawnpermision = ""
    statementpermision = ""
    activatepermision = ""
    periodtranpermision = ""
    monthlydedpermision = ""
    guarantorspermision = ""
    archivedpermision = ""
    dormantpermision = ""
    parampermision = ""
    loantypespermision = ""
    companysetpermision = ""
    banksetuppermision = ""
    clearloanpermision = ""
    clearloanpermision = ""
    databasesetpermision = ""
    rejreasonpermision = ""
    usergroupspermision = ""
    sysuserspermision = ""
    changenopermision = ""
Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
    
End Sub

Function EncryptPassword()
On Error GoTo errFix
    Dim Pwd As Variant
    Dim Temp As String, PwdChr As Long
    Dim EncryptKey As Long
    Pwd = valToEncrOrDecr
    EncryptKey = Int(Sqr(Len(Pwd) * 81)) + 23
    For PwdChr = 1 To Len(Pwd)
        Temp = Temp + Chr(Asc(Mid(Pwd, PwdChr, 1)) Xor EncryptKey)
    Next PwdChr
    
    EncryptPass = Temp
    Exit Function
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
End Function

Public Sub load_records()
On Error GoTo errFix
    If disablemodifying = False Then
    cmdAdd.Enabled = True
    cmdEdit.Enabled = True
    cmddelete.Enabled = True
    End If
    cmdCancel.Enabled = False
    cmdupdate.Enabled = False
    lvwSummary.Visible = False
    fraSetup.Visible = True 'Just in case listview was the one visible previously
    fraMembers.Visible = True
    fraLoans.Visible = True
    fraUserGrps.Visible = True
    fraUtilities.Visible = True
    fraShares.Visible = True
    fraOtherSchemes.Visible = True
    fraLoans.Visible = True
    txtGroupID.Enabled = False
    txtdescription.Enabled = False
    
    Toolbar1.Buttons("bSearch").Enabled = True
    Toolbar1.Buttons("bView").Enabled = True
    fraAllContents.Enabled = False
    Set rst = oSaccoMaster.GetRecordSet("select * from usergrps where groupid= '" & txtGroupID.Text & "'")
    If rst.RecordCount > 0 Then
        txtdescription.Text = rst!description & ""
        
        valToEncrOrDecr = rst!rwmembers & ""
        Call EncryptPassword
        rst!rwmembers = EncryptPass
        If rst!rwmembers = "Yes" Then
            manuallychecked = False
            chkMembership.Value = vbChecked
        ElseIf rst!rwmembers = "No" Then
            chkMembership.Value = vbUnchecked
        End If
        
        valToEncrOrDecr = rst!RWShares & ""
        Call EncryptPassword
        rst!RWShares = EncryptPass
        If rst!RWShares = "Yes" Then
            chkShares.Value = vbChecked
            manuallychecked = False
        ElseIf rst!RWShares = "No" Then
            chkShares.Value = vbUnchecked
        End If
        
'        If Rst!rwbanking = "Yes" Then
'             manuallychecked = False
'            chkBanking.Value = vbChecked
'        ElseIf Rst!rwbanking = "View" Then
'             chkBanking.Value = vbUnchecked
'        End If
        
        valToEncrOrDecr = rst!RWLoans & ""
        Call EncryptPassword
        rst!RWLoans = EncryptPass
        If rst!RWLoans = "Yes" Then
             manuallychecked = False
            chkLoans.Value = vbChecked
        ElseIf rst!RWLoans = "No" Then
            chkLoans.Value = vbUnchecked
        End If
        
        valToEncrOrDecr = rst!RWUtilities
        Call EncryptPassword
        rst!RWUtilities = EncryptPass
        If rst!RWUtilities = "Yes" Then
            manuallychecked = False
            chkUtilities.Value = vbChecked
        ElseIf rst!RWUtilities = "No" Then
            chkUtilities.Value = vbUnchecked
        End If
        
        valToEncrOrDecr = rst!RWSetup & ""
        Call EncryptPassword
        rst!RWSetup = EncryptPass
        If rst!RWSetup = "Yes" Then
            manuallychecked = False
            chkSetup.Value = vbChecked
        ElseIf rst!RWSetup = "No" Then
            chkSetup.Value = vbUnchecked
        End If
        valToEncrOrDecr = rst!RWOtherSchemes
        Call EncryptPassword
        rst!RWOtherSchemes = EncryptPass
        If rst!RWOtherSchemes = "Yes" Then
            manuallychecked = False
            chkOtherSchemes.Value = vbChecked
        ElseIf rst!RWOtherSchemes = "No" Then
            chkOtherSchemes.Value = vbUnchecked
        End If
        'UNDER MEMBERSHIP
        valToEncrOrDecr = rst!memreg & ""
        Call EncryptPassword
        If EncryptPass = "Non" Then
            manuallychecked = False
            chkMemberReg.Value = vbUnchecked
        ElseIf EncryptPass = "Mod" Or EncryptPass = "View" Then
            chkMemberReg.Value = vbChecked
        End If
        valToEncrOrDecr = rst!nok & ""
        Call EncryptPassword
        If EncryptPass = "Non" Then
            manuallychecked = False
            chkNextOfKin.Value = vbUnchecked
        ElseIf EncryptPass = "Mod" Or EncryptPass = "View" Then
            chkNextOfKin.Value = vbChecked
        End If
        valToEncrOrDecr = rst!memstatement & ""
        Call EncryptPassword
        If EncryptPass = "Non" Then
            manuallychecked = False
            chkMemStatement.Value = vbUnchecked
        ElseIf EncryptPass = "Mod" Or EncryptPass = "View" Then
            chkMemStatement.Value = vbChecked
        End If
        
        'UNDER LOANS
        valToEncrOrDecr = rst!loanapplic & ""
        Call EncryptPassword
        If EncryptPass = "Non" Then
            manuallychecked = False
            chkLoanApp.Value = vbUnchecked
        ElseIf EncryptPass = "Mod" Or EncryptPass = "View" Then
            chkLoanApp.Value = vbChecked
        End If
        valToEncrOrDecr = rst!loanendorsement & ""
        Call EncryptPassword
        If EncryptPass = "Non" Then
            manuallychecked = False
            chkLoanEndorsement.Value = vbUnchecked
        ElseIf EncryptPass = "Mod" Or EncryptPass = "View" Then
             chkLoanEndorsement.Value = vbChecked
        End If
        valToEncrOrDecr = rst!loanbal & ""
        Call EncryptPassword
        If EncryptPass = "Non" Then
            manuallychecked = False
            chkLoanBal.Value = vbUnchecked
        ElseIf EncryptPass = "Mod" Or EncryptPass = "View" Then
            chkLoanBal.Value = vbChecked
        End If
        valToEncrOrDecr = rst!loantrans & ""
        Call EncryptPassword
        If EncryptPass = "Non" Then
            manuallychecked = False
            chktransactions.Value = vbUnchecked
        ElseIf EncryptPass = "Mod" Or EncryptPass = "View" Then
            chktransactions.Value = vbChecked
        End If
        valToEncrOrDecr = rst!effectrepayment & ""
        Call EncryptPassword
        If EncryptPass = "Non" Then
            manuallychecked = False
            chkEffectRepay.Value = vbUnchecked
        ElseIf EncryptPass = "Mod" Or EncryptPass = "View" Then
            chkEffectRepay.Value = vbChecked
        End If
        valToEncrOrDecr = rst!loanguarantors & ""
        Call EncryptPassword
        If EncryptPass = "Non" Then
            manuallychecked = False
            chkLoanGuar.Value = vbUnchecked
        ElseIf EncryptPass = "Mod" Or EncryptPass = "View" Then
            chkLoanGuar.Value = vbChecked
        End If
        valToEncrOrDecr = rst!chequeentry & ""
        Call EncryptPassword
        If EncryptPass = "Non" Then
            manuallychecked = False
            chkChequeEntry.Value = vbUnchecked
        ElseIf EncryptPass = "Mod" Or EncryptPass = "View" Then
            chkChequeEntry.Value = vbChecked
        End If
        
        'UNDER UTILITIES
        valToEncrOrDecr = rst!calc & ""
        Call EncryptPassword
        If EncryptPass = "Non" Then
            manuallychecked = False
            chkCalc.Value = vbUnchecked
        ElseIf EncryptPass = "Mod" Or EncryptPass = "View" Then
            chkCalc.Value = vbChecked
        End If
        valToEncrOrDecr = rst!dividends & ""
        Call EncryptPassword
        If EncryptPass = "Non" Then
            manuallychecked = False
            chkDividends.Value = vbUnchecked
        ElseIf EncryptPass = "Mod" Or EncryptPass = "View" Then
             chkDividends.Value = vbChecked
        End If
        valToEncrOrDecr = rst!deductions & ""
        Call EncryptPassword
        If EncryptPass = "Non" Then
            manuallychecked = False
            chkDeductions.Value = vbUnchecked
        ElseIf EncryptPass = "Mod" Or EncryptPass = "View" Then
            chkDeductions.Value = vbChecked
        End If
        valToEncrOrDecr = rst!periodictran & ""
        Call EncryptPassword
        If EncryptPass = "Non" Then
            manuallychecked = False
            chkPeriodicTran.Value = vbUnchecked
        ElseIf EncryptPass = "Mod" Or EncryptPass = "View" Then
            chkPeriodicTran.Value = vbChecked
        End If
        valToEncrOrDecr = rst!utilstatements & ""
        Call EncryptPassword
        If EncryptPass = "Non" Then
            manuallychecked = False
            chkUtilStatements.Value = vbUnchecked
        ElseIf EncryptPass = "Mod" Or EncryptPass = "View" Then
            chkUtilStatements.Value = vbChecked
        End If
        valToEncrOrDecr = rst!monthlydeductions & ""
        Call EncryptPassword
        If EncryptPass = "Non" Then
            manuallychecked = False
            chkMonthlyDeductions.Value = vbUnchecked
        ElseIf EncryptPass = "Mod" Or EncryptPass = "View" Then
            chkMonthlyDeductions.Value = vbChecked
        End If
        valToEncrOrDecr = rst!exporttogl
        Call EncryptPassword
        If EncryptPass = "Non" Then
            manuallychecked = False
            chkExportToGL.Value = vbUnchecked
        ElseIf EncryptPass = "Mod" Or EncryptPass = "View" Then
            chkExportToGL.Value = vbChecked
        End If
        valToEncrOrDecr = rst!utilguarantor & ""
        Call EncryptPassword
        If EncryptPass = "Non" Then
            manuallychecked = False
            chkUtilGuar.Value = vbUnchecked
        ElseIf EncryptPass = "Mod" Or EncryptPass = "View" Then
            chkUtilGuar.Value = vbChecked
        End If
        valToEncrOrDecr = rst!dormant & ""
        Call EncryptPassword
        If EncryptPass = "Non" Then
            manuallychecked = False
            chkDormant.Value = vbUnchecked
        ElseIf EncryptPass = "Mod" Or EncryptPass = "View" Then
            chkDormant.Value = vbChecked
        End If
        valToEncrOrDecr = rst!Backup & ""
        Call EncryptPassword
        If EncryptPass = "Non" Then
            manuallychecked = False
            chkBackup.Value = vbUnchecked
        ElseIf EncryptPass = "Mod" Or EncryptPass = "View" Then
            chkBackup.Value = vbChecked
        End If
        valToEncrOrDecr = rst!Archived & ""
        Call EncryptPassword
        If EncryptPass = "Non" Then
            manuallychecked = False
            chkArchived.Value = vbUnchecked
        ElseIf EncryptPass = "Mod" Or EncryptPass = "View" Then
            chkArchived.Value = vbChecked
        End If
        valToEncrOrDecr = rst!Activate & ""
        Call EncryptPassword
        If EncryptPass = "Non" Then
            manuallychecked = False
            chkActivate.Value = vbUnchecked
        ElseIf EncryptPass = "Mod" Or EncryptPass = "View" Then
            chkActivate.Value = vbChecked
        End If
        valToEncrOrDecr = rst!withdrawn & ""
        Call EncryptPassword
        If EncryptPass = "Non" Then
            manuallychecked = False
            chkWithdrawn.Value = vbUnchecked
        ElseIf EncryptPass = "Mod" Or EncryptPass = "View" Then
            chkWithdrawn.Value = vbChecked
        End If
        
        'UNDER SETUP
        valToEncrOrDecr = rst!parametization & ""
        Call EncryptPassword
        If EncryptPass = "Non" Then
            manuallychecked = False
            chkParametization.Value = vbUnchecked
        ElseIf EncryptPass = "Mod" Or EncryptPass = "View" Then
            chkParametization.Value = vbChecked
        End If
        valToEncrOrDecr = rst!loantypes & ""
        Call EncryptPassword
        If EncryptPass = "Non" Then
            manuallychecked = False
            chkLoanType.Value = vbUnchecked
        ElseIf EncryptPass = "Mod" Or EncryptPass = "View" Then
             chkLoanType.Value = vbChecked
        End If
        valToEncrOrDecr = rst!Activate & ""
        Call EncryptPassword
        If EncryptPass = "Non" Then
            manuallychecked = False
            chkActivate.Value = vbUnchecked
        ElseIf EncryptPass = "Mod" Or EncryptPass = "View" Then
             chkActivate.Value = vbChecked
        End If
        valToEncrOrDecr = rst!companysetup & ""
        Call EncryptPassword
        If EncryptPass = "Non" Then
           manuallychecked = False
           chkCompanySetup.Value = vbUnchecked
        ElseIf EncryptPass = "Mod" Or EncryptPass = "View" Then
            chkCompanySetup.Value = vbChecked
        End If
        valToEncrOrDecr = rst!banksetup & ""
        Call EncryptPassword
        If EncryptPass = "Non" Then
            manuallychecked = False
            chkBankSetup.Value = vbUnchecked
        ElseIf EncryptPass = "Mod" Or EncryptPass = "View" Then
            chkBankSetup.Value = vbChecked
        End If
        valToEncrOrDecr = rst!databasesetup & ""
        Call EncryptPassword
        If EncryptPass = "Non" Then
            manuallychecked = False
            chkDatabaseSetup.Value = vbUnchecked
        ElseIf EncryptPass = "Mod" Or EncryptPass = "View" Then
            chkDatabaseSetup.Value = vbChecked
        End If
        valToEncrOrDecr = rst!rejreasons & ""
        Call EncryptPassword
        If EncryptPass = "Non" Then
            manuallychecked = False
            chkRejReasons.Value = vbUnchecked
        ElseIf EncryptPass = "Mod" Or EncryptPass = "View" Then
            chkRejReasons.Value = vbChecked
        End If
        valToEncrOrDecr = rst!usergrps & ""
        Call EncryptPassword
        If EncryptPass = "Non" Then
            manuallychecked = False
            chkUserGrps.Value = vbUnchecked
        ElseIf EncryptPass = "Mod" Or EncryptPass = "View" Then
            chkUserGrps.Value = vbChecked
        End If
        valToEncrOrDecr = rst!sysusers & ""
        Call EncryptPassword
        If EncryptPass = "Non" Then
            manuallychecked = False
            chkSysUsers.Value = vbUnchecked
        ElseIf EncryptPass = "Mod" Or EncryptPass = "View" Then
            chkSysUsers.Value = vbChecked
        End If
        valToEncrOrDecr = rst!clearLoanRecs & ""
        Call EncryptPassword
        If EncryptPass = "Non" Then
            manuallychecked = False
            chkClearLoanRecords.Value = vbUnchecked
        ElseIf EncryptPass = "Mod" Or EncryptPass = "View" Then
            chkClearLoanRecords.Value = vbChecked
        End If
        valToEncrOrDecr = rst!clearMemRecs & ""
        Call EncryptPassword
        If EncryptPass = "Non" Then
            manuallychecked = False
            chkClearMemRecords.Value = vbUnchecked
        ElseIf EncryptPass = "Mod" Or EncryptPass = "View" Then
            chkClearMemRecords.Value = vbChecked
        End If
        valToEncrOrDecr = rst!changememno & ""
        Call EncryptPassword
        If EncryptPass = "Non" Then
            manuallychecked = False
            chkChangeMembershipNo.Value = vbUnchecked
        ElseIf EncryptPass = "Mod" Or EncryptPass = "View" Then
            chkChangeMembershipNo.Value = vbChecked
        End If
        
        'UNDER SHARES
        valToEncrOrDecr = rst!Contributions & ""
        Call EncryptPassword
        If EncryptPass = "Non" Then
            manuallychecked = False
            chkContributions.Value = vbUnchecked
        ElseIf EncryptPass = "Mod" Or EncryptPass = "View" Then
            chkContributions.Value = vbChecked
        End If
        valToEncrOrDecr = rst!sharevar & ""
        Call EncryptPassword
        If EncryptPass = "Non" Then
            manuallychecked = False
            chkShareVar.Value = vbUnchecked
        ElseIf EncryptPass = "Mod" Or EncryptPass = "View" Then
            chkShareVar.Value = vbChecked
        End If
        
        'UNDER OTHER SCHEMES
        valToEncrOrDecr = rst!savings & ""
        Call EncryptPassword
        If EncryptPass = "Non" Then
            manuallychecked = False
            chkSavings.Value = vbUnchecked
        ElseIf EncryptPass = "Mod" Or EncryptPass = "View" Then
            chkSavings.Value = vbChecked
        End If
        valToEncrOrDecr = rst!benevolentfund
        Call EncryptPassword
        If EncryptPass = "Non" Then
            manuallychecked = False
            chkBenevolentFund.Value = vbUnchecked
        ElseIf EncryptPass = "Mod" Or EncryptPass = "View" Then
            chkBenevolentFund.Value = vbChecked
        End If
        
       
        
'        Set rst2 = oSaccomaster.GetRecordSet("select userid from users where groupid= '" & Rst!groupid & "'")
'        If rst2.RecordCount > 0 Then
'            txtNumUsers.Text = rst2.RecordCount
'        Else
'            txtNumUsers.Text = 0
'        End If
    
    End If
Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"
        
End Sub
Public Sub onRefreshOfSearchFrm()
On Error GoTo errFix
    frmSearch.lstSearch.ListItems.Clear
    strSQL = "Select * from usergrps order by groupid"
    Set rst = oSaccoMaster.GetRecordSet(strSQL)
    With rst
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                Set li = frmSearch.lstSearch.ListItems.Add(, , !groupid & "")
                li.ListSubItems.Add , , !description & ""
                li.ListSubItems.Add , , !rwmembers & ""
                li.ListSubItems.Add , , !RWShares & ""
                li.ListSubItems.Add , , !RWLoans & ""
                li.ListSubItems.Add , , !rwbanking & ""
                li.ListSubItems.Add , , !RWUtilities & ""
                li.ListSubItems.Add , , !RWSetup & ""
                .MoveNext
            Loop
        End If
        .Close
    End With
    Set rst = Nothing
    frmSearch.Txtfrom.Text = ""
    frmSearch.Txtto.Text = ""
    frmSearch.cboCrieria.Text = "="
    frmSearch.cboField.Text = "Group ID"
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"

End Sub
Public Sub onFormLoadOfSearchFrm()
On Error GoTo errFix
    Set formCallingSearch = FrmUserGroups
    frmSearch.Caption = "Search User Groups"
    PositionForm frmSearch
    
    With frmSearch.lstSearch
        .ListItems.Clear
        .Columnheaders.Clear
        .Columnheaders.Add 1, , "Group ID", 2000
        .Columnheaders.Add 2, , "Description", 2000
        .Columnheaders.Add 3, , "Membership", 3000
        .Columnheaders.Add 4, , "Shares", 3000
        .Columnheaders.Add 5, , "Loans", 2000
        .Columnheaders.Add 6, , "Utilities", 2000
        .Columnheaders.Add 7, , "Setup", 2000
        .Columnheaders.Add 8, , "OtherSchemes", 3000
        .Columnheaders.Add 9, , "Registration", 3000
        .Columnheaders.Add 10, , "NextOfKin", 3000
        .Columnheaders.Add 11, , "MemberStatement", 3000
        .Columnheaders.Add 12, , "LoanApp", 2000
        .Columnheaders.Add 13, , "LoanEndordement", 3000
        .Columnheaders.Add 14, , "LoanBal", 3000
        .Columnheaders.Add 15, , "LoanTransaction", 3000
        .Columnheaders.Add 16, , "EffectRepay", 3000
        .Columnheaders.Add 17, , "LoanGuar", 3000
        .Columnheaders.Add 18, , "ChequeEntry", 3000
        .Columnheaders.Add 19, , "Calc", 1000
        .Columnheaders.Add 20, , "Dividends", 2000
        .Columnheaders.Add 21, , "Deductions", 2000
        .Columnheaders.Add 22, , "PeriodicTran", 3000
        .Columnheaders.Add 23, , "UtilStatement", 2000
        .Columnheaders.Add 24, , "MonthlyDeduct.", 3000
        .Columnheaders.Add 25, , "ExportToGL", 2000
        .Columnheaders.Add 26, , "UtilGuar", 3000
        .Columnheaders.Add 27, , "Dormant", 2000
        .Columnheaders.Add 28, , "Archived", 3000
        .Columnheaders.Add 29, , "Withdrawn", 2000
        .Columnheaders.Add 30, , "Parametization", 3000
        .Columnheaders.Add 31, , "CompanySet.", 2000
        .Columnheaders.Add 32, , "BankSet", 2000
        .Columnheaders.Add 33, , "DatabaseSetup", 2000
        .Columnheaders.Add 33, , "Activate", 2000
        .Columnheaders.Add 34, , "RejReasons", 2000
        .Columnheaders.Add 35, , "UserGrps", 2000
        .Columnheaders.Add 36, , "SysUsers", 2000
        .Columnheaders.Add 37, , "ClearLoanRecs", 2000
        .Columnheaders.Add 38, , "ClearMemRecs", 2000
        .Columnheaders.Add 39, , "ChangeMemNo", 2000
        .Columnheaders.Add 40, , "Contrib", 2000
        .Columnheaders.Add 41, , "ShareVar", 2000
        .Columnheaders.Add 42, , "Savings", 2000
        .Columnheaders.Add 43, , "LoanTypes", 2000
        .Columnheaders.Add 44, , "BenFund", 2000
        .View = lvwReport
        .Gridlines = True
    End With
    With frmSearch.cboField
        .AddItem ("Group ID")
        .AddItem ("Description")
    End With
    strSQL = "Select * from usergrps order by groupid"
    Set rst = oSaccoMaster.GetRecordSet(strSQL)
    With rst
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                Set li = frmSearch.lstSearch.ListItems.Add(, , !groupid & "")
                li.ListSubItems.Add , , !description & ""
                li.ListSubItems.Add , , !rwmembers & ""
                li.ListSubItems.Add , , !RWShares & ""
                li.ListSubItems.Add , , !RWLoans & ""
                li.ListSubItems.Add , , !RWUtilities & ""
                li.ListSubItems.Add , , !RWSetup & ""
                li.ListSubItems.Add , , !RWOtherSchemes & ""
                li.ListSubItems.Add , , !memreg & ""
                li.ListSubItems.Add , , !nok & ""
                li.ListSubItems.Add , , !memstatement & ""
                li.ListSubItems.Add , , !loanapplic & ""
                li.ListSubItems.Add , , !loanendorsement & ""
                li.ListSubItems.Add , , !loanbal & ""
                li.ListSubItems.Add , , !loantrans & ""
                li.ListSubItems.Add , , !effectrepayment & ""
                li.ListSubItems.Add , , !loanguarantors & ""
                li.ListSubItems.Add , , !chequeentry & ""
                li.ListSubItems.Add , , !calc & ""
                li.ListSubItems.Add , , !dividends & ""
                li.ListSubItems.Add , , !deductions & ""
                li.ListSubItems.Add , , !periodictran & ""
                li.ListSubItems.Add , , !utilstatements & ""
                li.ListSubItems.Add , , !monthlydeductions & ""
                li.ListSubItems.Add , , !exporttogl & ""
                li.ListSubItems.Add , , !utilguarantor & ""
                li.ListSubItems.Add , , !dormant & ""
                li.ListSubItems.Add , , !Archived & ""
                li.ListSubItems.Add , , !withdrawn & ""
                li.ListSubItems.Add , , !parametization & ""
                li.ListSubItems.Add , , rst!companysetup & ""
                li.ListSubItems.Add , , !banksetup & ""
                li.ListSubItems.Add , , !databasesetup & ""
                li.ListSubItems.Add , , !Activate & ""
                li.ListSubItems.Add , , !rejreasons & ""
                li.ListSubItems.Add , , !usergrps & ""
                li.ListSubItems.Add , , !sysusers & ""
                li.ListSubItems.Add , , !clearLoanRecs & ""
                li.ListSubItems.Add , , !clearMemRecs & ""
                li.ListSubItems.Add , , !changememno & ""
                li.ListSubItems.Add , , !Contributions & ""
                li.ListSubItems.Add , , !sharevar & ""
                li.ListSubItems.Add , , !savings & ""
                li.ListSubItems.Add , , !loantypes & ""
                li.ListSubItems.Add , , !benevolentfund & ""
                .MoveNext
            Loop
        End If
        .Close
    End With
    Set rst = Nothing
    frmSearch.cboCrieria.Text = frmSearch.cboCrieria.List(0)
    frmSearch.cboField.Text = frmSearch.cboField.List(0)
    frmSearch.Show vbModal
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"

End Sub
Public Sub searchSelect()
    On Error GoTo errFix
    frmSearch.Visible = False
    If lvwSummary.Visible Then
            strSQL = "Select * from usergrps order by groupid"
            Set rst = oSaccoMaster.GetRecordSet(strSQL)
            If rst.RecordCount > 0 Then
                    li = lvwSummary.ListItems(1)
                    For Each li In lvwSummary.ListItems
                        If li.Text = Sel Then

                            lvwSummary.ListItems.Remove li.Index
                            With rst
                                Set li = lvwSummary.ListItems.Add(1, , Sel & "")
                                lvwSummary.ListItems(1).selected = True
                                li.ListSubItems.Add , , !description & ""
                                li.ListSubItems.Add , , !rwmembers & ""
                                li.ListSubItems.Add , , !RWShares & ""
                                li.ListSubItems.Add , , !RWLoans & ""
                                li.ListSubItems.Add , , !RWUtilities & ""
                                li.ListSubItems.Add , , !RWSetup & ""
                                li.ListSubItems.Add , , !RWOtherSchemes & ""
                                li.ListSubItems.Add , , !memreg & ""
                                li.ListSubItems.Add , , !nok & ""
                                li.ListSubItems.Add , , !memstatement & ""
                                li.ListSubItems.Add , , !loanapplic & ""
                                li.ListSubItems.Add , , !loanendorsement & ""
                                li.ListSubItems.Add , , !loanbal & ""
                                li.ListSubItems.Add , , !loantrans & ""
                                li.ListSubItems.Add , , !effectrepayment & ""
                                li.ListSubItems.Add , , !loanguarantors & ""
                                li.ListSubItems.Add , , !chequeentry & ""
                                li.ListSubItems.Add , , !calc & ""
                                li.ListSubItems.Add , , !dividends & ""
                                li.ListSubItems.Add , , !deductions & ""
                                li.ListSubItems.Add , , !periodictran & ""
                                li.ListSubItems.Add , , !utilstatements & ""
                                li.ListSubItems.Add , , !monthlydeductions & ""
                                li.ListSubItems.Add , , !exporttogl & ""
                                li.ListSubItems.Add , , !utilguarantor & ""
                                li.ListSubItems.Add , , !dormant & ""
                                li.ListSubItems.Add , , !Archived & ""
                                li.ListSubItems.Add , , !withdrawn & ""
                                li.ListSubItems.Add , , !parametization & ""
                                li.ListSubItems.Add , , rst!companysetup & ""
                                li.ListSubItems.Add , , !banksetup & ""
                                li.ListSubItems.Add , , !databasesetup & ""
                                li.ListSubItems.Add , , !Activate & ""
                                li.ListSubItems.Add , , !rejreasons & ""
                                li.ListSubItems.Add , , !usergrps & ""
                                li.ListSubItems.Add , , !sysusers & ""
                                li.ListSubItems.Add , , !clearLoanRecs & ""
                                li.ListSubItems.Add , , !clearMemRecs & ""
                                li.ListSubItems.Add , , !changememno & ""
                                li.ListSubItems.Add , , !Contributions & ""
                                li.ListSubItems.Add , , !sharevar & ""
                                li.ListSubItems.Add , , !savings & ""
                                li.ListSubItems.Add , , !loantypes & ""
                                li.ListSubItems.Add , , !benevolentfund & ""
                                Exit For
                            End With
                    End If
                   Next li
                Else
                    MsgBox "Group does not exist", vbInformation, "Membership"
                End If
    Else
        txtGroupID.Text = Sel
        formCallingSearch.load_records
    End If
    frmSearch.Visible = False
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"

End Sub
Public Sub onFindOfSearchFrmClick()
On Error GoTo errFix
    Dim Find As Integer
    Select Case frmSearch.cboField.Text
        Case "Group ID"
            searchField = "groupid"
        Case "Description"
            searchField = "description"
    End Select
    frmSearch.lstSearch.ListItems.Clear
    If Not frmSearch.cboField.Text = "" Then
        If Not frmSearch.cboCrieria.Text = "" Then
            If Not frmSearch.cboCrieria.Text = "Between" And Not frmSearch.cboCrieria.Text = "Like" Then
                strSQL = "Select * from usergrps where " & searchField & " " & frmSearch.cboCrieria.Text & " '" & frmSearch.Txtfrom.Text & "'"
                Set rst = oSaccoMaster.GetRecordSet(strSQL)
                With rst
                    If .RecordCount > 0 Then
                        .MoveFirst
                        Do While Not .EOF
                            Set li = frmSearch.lstSearch.ListItems.Add(, , !groupid & "")
                            li.ListSubItems.Add , , !description & ""
                            li.ListSubItems.Add , , !rwmembers & ""
                            li.ListSubItems.Add , , !RWShares & ""
                            li.ListSubItems.Add , , !RWLoans & ""
                            li.ListSubItems.Add , , !RWUtilities & ""
                            li.ListSubItems.Add , , !RWSetup & ""
                            li.ListSubItems.Add , , !RWOtherSchemes & ""
                            li.ListSubItems.Add , , !memreg & ""
                            li.ListSubItems.Add , , !nok & ""
                            li.ListSubItems.Add , , !memstatement & ""
                            li.ListSubItems.Add , , !loanapplic & ""
                            li.ListSubItems.Add , , !loanendorsement & ""
                            li.ListSubItems.Add , , !loanbal & ""
                            li.ListSubItems.Add , , !loantrans & ""
                            li.ListSubItems.Add , , !effectrepayment & ""
                            li.ListSubItems.Add , , !loanguarantors & ""
                            li.ListSubItems.Add , , !chequeentry & ""
                            li.ListSubItems.Add , , !calc & ""
                            li.ListSubItems.Add , , !dividends & ""
                            li.ListSubItems.Add , , !deductions & ""
                            li.ListSubItems.Add , , !periodictran & ""
                            li.ListSubItems.Add , , !utilstatements & ""
                            li.ListSubItems.Add , , !monthlydeductions & ""
                            li.ListSubItems.Add , , !exporttogl & ""
                            li.ListSubItems.Add , , !utilguarantor & ""
                            li.ListSubItems.Add , , !dormant & ""
                            li.ListSubItems.Add , , !Archived & ""
                            li.ListSubItems.Add , , !withdrawn & ""
                            li.ListSubItems.Add , , !parametization & ""
                            li.ListSubItems.Add , , rst!companysetup & ""
                            li.ListSubItems.Add , , !banksetup & ""
                            li.ListSubItems.Add , , !databasesetup & ""
                            li.ListSubItems.Add , , !Activate & ""
                            li.ListSubItems.Add , , !rejreasons & ""
                            li.ListSubItems.Add , , !usergrps & ""
                            li.ListSubItems.Add , , !sysusers & ""
                            li.ListSubItems.Add , , !clearLoanRecs & ""
                            li.ListSubItems.Add , , !clearMemRecs & ""
                            li.ListSubItems.Add , , !changememno & ""
                            li.ListSubItems.Add , , !Contributions & ""
                            li.ListSubItems.Add , , !sharevar & ""
                            li.ListSubItems.Add , , !savings & ""
                            li.ListSubItems.Add , , !loantypes & ""
                            li.ListSubItems.Add , , !benevolentfund & ""
                            .MoveNext
                        Loop
                    End If
                End With
                Set rst = Nothing
            ElseIf frmSearch.cboCrieria.Text = "Like" Then
                sql = "Select * from usergrps order by groupid"
                Set rst = oSaccoMaster.GetRecordSet(strSQL)
                With rst
                    If .RecordCount > 0 Then
                        .MoveFirst
                        Do While Not .EOF
                            .Find "" & searchField & " " & frmSearch.cboCrieria.Text & " '" & frmSearch.Txtfrom.Text & "%'", , adSearchForward
                            If Not .EOF Then
                                Set li = frmSearch.lstSearch.ListItems.Add(, , !groupid & "")
                                li.ListSubItems.Add , , !description & ""
                                li.ListSubItems.Add , , !rwmembers & ""
                                li.ListSubItems.Add , , !RWShares & ""
                                li.ListSubItems.Add , , !RWLoans & ""
                                li.ListSubItems.Add , , !RWUtilities & ""
                                li.ListSubItems.Add , , !RWSetup & ""
                                li.ListSubItems.Add , , !RWOtherSchemes & ""
                                li.ListSubItems.Add , , !memreg & ""
                                li.ListSubItems.Add , , !nok & ""
                                li.ListSubItems.Add , , !memstatement & ""
                                li.ListSubItems.Add , , !loanapplic & ""
                                li.ListSubItems.Add , , !loanendorsement & ""
                                li.ListSubItems.Add , , !loanbal & ""
                                li.ListSubItems.Add , , !loantrans & ""
                                li.ListSubItems.Add , , !effectrepayment & ""
                                li.ListSubItems.Add , , !loanguarantors & ""
                                li.ListSubItems.Add , , !chequeentry & ""
                                li.ListSubItems.Add , , !calc & ""
                                li.ListSubItems.Add , , !dividends & ""
                                li.ListSubItems.Add , , !deductions & ""
                                li.ListSubItems.Add , , !periodictran & ""
                                li.ListSubItems.Add , , !utilstatements & ""
                                li.ListSubItems.Add , , !monthlydeductions & ""
                                li.ListSubItems.Add , , !exporttogl & ""
                                li.ListSubItems.Add , , !utilguarantor & ""
                                li.ListSubItems.Add , , !dormant & ""
                                li.ListSubItems.Add , , !Archived & ""
                                li.ListSubItems.Add , , !withdrawn & ""
                                li.ListSubItems.Add , , !parametization & ""
                                li.ListSubItems.Add , , rst!companysetup & ""
                                li.ListSubItems.Add , , !banksetup & ""
                                li.ListSubItems.Add , , !databasesetup & ""
                                li.ListSubItems.Add , , !Activate & ""
                                li.ListSubItems.Add , , !rejreasons & ""
                                li.ListSubItems.Add , , !usergrps & ""
                                li.ListSubItems.Add , , !sysusers & ""
                                li.ListSubItems.Add , , !clearLoanRecs & ""
                                li.ListSubItems.Add , , !clearMemRecs & ""
                                li.ListSubItems.Add , , !changememno & ""
                                li.ListSubItems.Add , , !Contributions & ""
                                li.ListSubItems.Add , , !sharevar & ""
                                li.ListSubItems.Add , , !savings & ""
                                li.ListSubItems.Add , , !loantypes & ""
                                li.ListSubItems.Add , , !benevolentfund & ""
                                .MoveNext
                            End If
                        Loop
                    End If
                End With
                Set rst = Nothing
            Else
                    strSQL = "select * from usergrps where " & searchField & " between'" & frmSearch.Txtfrom.Text & "' And '" & frmSearch.Txtto.Text & " '"
                    Set rst = oSaccoMaster.GetRecordSet(strSQL)
                With rst
                    If .RecordCount > 0 Then
                        .MoveFirst
                        Do While Not .EOF
                            Set li = frmSearch.lstSearch.ListItems.Add(, , !groupid & "")
                             li.ListSubItems.Add , , !description & ""
                            li.ListSubItems.Add , , !rwmembers & ""
                            li.ListSubItems.Add , , !RWShares & ""
                            li.ListSubItems.Add , , !RWLoans & ""
                            li.ListSubItems.Add , , !RWUtilities & ""
                            li.ListSubItems.Add , , !RWSetup & ""
                            li.ListSubItems.Add , , !RWOtherSchemes & ""
                            li.ListSubItems.Add , , !memreg & ""
                            li.ListSubItems.Add , , !nok & ""
                            li.ListSubItems.Add , , !memstatement & ""
                            li.ListSubItems.Add , , !loanapplic & ""
                            li.ListSubItems.Add , , !loanendorsement & ""
                            li.ListSubItems.Add , , !loanbal & ""
                            li.ListSubItems.Add , , !loantrans & ""
                            li.ListSubItems.Add , , !effectrepayment & ""
                            li.ListSubItems.Add , , !loanguarantors & ""
                            li.ListSubItems.Add , , !chequeentry & ""
                            li.ListSubItems.Add , , !calc & ""
                            li.ListSubItems.Add , , !dividends & ""
                            li.ListSubItems.Add , , !deductions & ""
                            li.ListSubItems.Add , , !periodictran & ""
                            li.ListSubItems.Add , , !utilstatements & ""
                            li.ListSubItems.Add , , !monthlydeductions & ""
                            li.ListSubItems.Add , , !exporttogl & ""
                            li.ListSubItems.Add , , !utilguarantor & ""
                            li.ListSubItems.Add , , !dormant & ""
                            li.ListSubItems.Add , , !Archived & ""
                            li.ListSubItems.Add , , !withdrawn & ""
                            li.ListSubItems.Add , , !parametization & ""
                            li.ListSubItems.Add , , rst!companysetup & ""
                            li.ListSubItems.Add , , !banksetup & ""
                            li.ListSubItems.Add , , !databasesetup & ""
                            li.ListSubItems.Add , , !Activate & ""
                            li.ListSubItems.Add , , !rejreasons & ""
                            li.ListSubItems.Add , , !usergrps & ""
                            li.ListSubItems.Add , , !sysusers & ""
                            li.ListSubItems.Add , , !clearLoanRecs & ""
                            li.ListSubItems.Add , , !clearMemRecs & ""
                            li.ListSubItems.Add , , !changememno & ""
                            li.ListSubItems.Add , , !Contributions & ""
                            li.ListSubItems.Add , , !sharevar & ""
                            li.ListSubItems.Add , , !savings & ""
                            li.ListSubItems.Add , , !loantypes & ""
                            li.ListSubItems.Add , , !benevolentfund & ""
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
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"

End Sub

Private Sub Form_Unload(Cancel As Integer)
action = ""
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo errFix
    Select Case Button.Key
        Case "bSearch"
        Set formCallingSearch = FrmUserGroups
            onFormLoadOfSearchFrm
    End Select
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"

End Sub
Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
On Error GoTo errFix
    Select Case ButtonMenu.Key
        Case "mDetails"
        If lvwSummary.Visible Then
            If Not lvwSummary.ListItems.Count = 0 Then
             Sel = lvwSummary.SelectedItem
            txtGroupID.Text = Sel
            End If
            
        End If
        cmdNext.Enabled = True
        cmdPrevious.Enabled = True
        cmdFirst.Enabled = True
        cmdLast.Enabled = True
        load_records
        lvwSummary.Visible = False
        fraAllContents.Visible = True
        Case "mSummary"
        fraAllContents.Visible = False
        load_Summary
        lvwSummary.Visible = True
        
    End Select
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"

End Sub



Private Sub txtDescription_Change()
On Error GoTo errFix
txtdescription.Text = UCase(txtdescription.Text)
txtdescription.SelStart = Len(txtdescription.Text)
Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"

End Sub

Private Sub txtdescription_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    chkMembership.SetFocus
End If
End Sub

Private Sub txtGroupID_Change()
On Error GoTo errFix
    txtGroupID.Text = UCase(txtGroupID.Text)
    txtGroupID.SelStart = Len(txtGroupID.Text)
    If txtGroupID.Text = "" Then
        chkMembership.Enabled = False
        chkShares.Enabled = False
        chkLoans.Enabled = False
        chkSetup.Enabled = False
        chkOtherSchemes.Enabled = False
        chkUtilities.Enabled = False
        If Not action = "addingRecords" Then
        fraAllContents.Enabled = False
        End If
    Else
        chkMembership.Enabled = True
        chkShares.Enabled = True
        chkLoans.Enabled = True
        chkSetup.Enabled = True
        chkOtherSchemes.Enabled = True
        chkUtilities.Enabled = True
        fraAllContents.Enabled = True
    End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"

End Sub

Private Sub txtGroupID_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
Set Rs = oSaccoMaster.GetRecordSet("select groupid from users where userid= '" & User & "'")
If Not Rs.EOF Then
    Set Rst5 = oSaccoMaster.GetRecordSet("select activate from usergrps where groupid= '" & Rs!groupid & "'")
    If Not Rst5.EOF Then
        valToEncrOrDecr = Rst5!Activate & ""
        EncryptPassword
        If Not EncryptPass = "Mod" Then
            MsgBox "You have no authority to create a user of this group.", vbInformation, "Only Vendor Can!"
            Exit Sub
        End If
    End If
End If
    If action = "addingRecords" Then
        If Not txtGroupID.Text = "" Then
            Set Rst5 = oSaccoMaster.GetRecordSet("select groupid from usergrps where groupid= '" & txtGroupID.Text & "'")
            If Not Rst5.EOF Then
                MsgBox "User group with same ID exists", vbInformation, "User Groups"
                txtGroupID.Text = ""
                txtGroupID.SetFocus
                Exit Sub
            End If
        End If
    End If
    txtdescription.SetFocus
End If
Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"

End Sub

Private Sub txtGroupID_LostFocus()
On Error GoTo errFix
'    txtNumUsers.Text = "0"
    If action = "addingRecords" Then
    
        If Not txtGroupID.Text = "" Then
            Set Rst5 = oSaccoMaster.GetRecordSet("select groupid from usergrps where groupid= '" & txtGroupID.Text & "'")
            If Not Rst5.EOF Then
                MsgBox "User group with same ID exists", vbInformation, "User Groups"
                txtGroupID.Text = ""
                txtGroupID.SetFocus
                Exit Sub
            End If
        End If
    End If
    Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "User Groups"

End Sub

Private Sub txtID_Change()

End Sub
