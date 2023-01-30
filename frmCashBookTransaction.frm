VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCashBookTransaction 
   Caption         =   "Cash Book Transaction"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10980
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   10980
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Refresh"
      Height          =   345
      Left            =   1575
      TabIndex        =   92
      Top             =   6765
      Width           =   1035
   End
   Begin VB.TextBox txtOpBalance 
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
      Height          =   300
      Left            =   5160
      TabIndex        =   81
      Top             =   300
      Width           =   1500
   End
   Begin VB.TextBox Text1 
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
      Height          =   300
      Left            =   6690
      TabIndex        =   80
      Top             =   300
      Width           =   1575
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   405
      Left            =   9420
      TabIndex        =   79
      Top             =   6690
      Width           =   1335
   End
   Begin VB.CommandButton cmdRefresh_data 
      Caption         =   "Load"
      Height          =   345
      Left            =   2700
      TabIndex        =   78
      Top             =   6765
      Width           =   1395
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "&Export"
      Height          =   345
      Left            =   45
      TabIndex        =   6
      Top             =   6765
      Width           =   1395
   End
   Begin VB.ComboBox cboBank 
      Height          =   315
      ItemData        =   "frmCashBookTransaction.frx":0000
      Left            =   0
      List            =   "frmCashBookTransaction.frx":0007
      TabIndex        =   5
      Top             =   285
      Width           =   1125
   End
   Begin VB.TextBox txtBankName 
      Height          =   300
      Left            =   1125
      TabIndex        =   4
      Top             =   300
      Width           =   2700
   End
   Begin VB.CommandButton cmdTransferFunds 
      Caption         =   "Transfer Funds"
      Height          =   345
      Left            =   4095
      TabIndex        =   2
      Top             =   6765
      Width           =   1395
   End
   Begin VB.CommandButton cmdOffset 
      Caption         =   "Change To Offset"
      Height          =   345
      Left            =   5610
      TabIndex        =   1
      Top             =   6795
      Width           =   1770
   End
   Begin VB.CommandButton cmdAmount 
      Caption         =   "Update Amount"
      Height          =   360
      Left            =   7530
      TabIndex        =   0
      Top             =   6765
      Width           =   1395
   End
   Begin MSComctlLib.ListView lvwAccounts 
      Height          =   1560
      Left            =   1140
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   2752
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "AccNo"
         Object.Width           =   18
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "AccName"
         Object.Width           =   10583
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3075
      Top             =   6690
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5880
      Left            =   45
      TabIndex        =   7
      Top             =   750
      Width           =   10755
      _ExtentX        =   18971
      _ExtentY        =   10372
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Cash Book Transactions"
      TabPicture(0)   =   "frmCashBookTransaction.frx":0011
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ListView1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraTransfer"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Bank Debits"
      TabPicture(1)   =   "frmCashBookTransaction.frx":002D
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label15"
      Tab(1).Control(1)=   "Label16"
      Tab(1).Control(2)=   "Label17"
      Tab(1).Control(3)=   "Label18"
      Tab(1).Control(4)=   "Label24"
      Tab(1).Control(5)=   "Label25"
      Tab(1).Control(6)=   "lvwDebits"
      Tab(1).Control(7)=   "txtCrAccNo"
      Tab(1).Control(8)=   "txtCrAccName"
      Tab(1).Control(9)=   "txtDrNarration"
      Tab(1).Control(10)=   "txtDrAmount"
      Tab(1).Control(11)=   "txtDrDocumentNo"
      Tab(1).Control(12)=   "dtpDrTransDate"
      Tab(1).Control(13)=   "cmdAddDebits"
      Tab(1).Control(14)=   "lvwCrAccounts"
      Tab(1).Control(15)=   "cmdDrPost"
      Tab(1).ControlCount=   16
      TabCaption(2)   =   "Bank Credits"
      TabPicture(2)   =   "frmCashBookTransaction.frx":0049
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label19"
      Tab(2).Control(1)=   "Label20"
      Tab(2).Control(2)=   "Label21"
      Tab(2).Control(3)=   "Label22"
      Tab(2).Control(4)=   "Label23"
      Tab(2).Control(5)=   "Label26"
      Tab(2).Control(6)=   "dtpCrTransDate"
      Tab(2).Control(7)=   "lvwCredits"
      Tab(2).Control(8)=   "txtDrAccNo"
      Tab(2).Control(9)=   "txtDrAccName"
      Tab(2).Control(10)=   "txtCrNarration"
      Tab(2).Control(11)=   "txtCrAmount"
      Tab(2).Control(12)=   "lvwDrAccounts"
      Tab(2).Control(13)=   "txtCrDocumentNo"
      Tab(2).Control(14)=   "cmdAddCredit"
      Tab(2).Control(15)=   "cmdCrPost"
      Tab(2).ControlCount=   16
      TabCaption(3)   =   "Reconciliation Report"
      TabPicture(3)   =   "frmCashBookTransaction.frx":0065
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label27"
      Tab(3).Control(1)=   "Label6"
      Tab(3).Control(2)=   "Label14"
      Tab(3).Control(3)=   "Label13"
      Tab(3).Control(4)=   "Label12"
      Tab(3).Control(5)=   "Label11"
      Tab(3).Control(6)=   "Label10"
      Tab(3).Control(7)=   "Label9"
      Tab(3).Control(8)=   "Label8"
      Tab(3).Control(9)=   "Label7"
      Tab(3).Control(10)=   "cmdPostRecon"
      Tab(3).Control(11)=   "txtDifference"
      Tab(3).Control(12)=   "txtOpeningBal"
      Tab(3).Control(13)=   "txtBankBal"
      Tab(3).Control(14)=   "txtBankDebits"
      Tab(3).Control(15)=   "txtBankCredits"
      Tab(3).Control(16)=   "txtDeposits"
      Tab(3).Control(17)=   "txtUnpresentedChq"
      Tab(3).Control(18)=   "txtCBBalance"
      Tab(3).Control(19)=   "txtPayments"
      Tab(3).Control(20)=   "txtReceipts"
      Tab(3).ControlCount=   21
      Begin VB.TextBox txtReceipts 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -71055
         Locked          =   -1  'True
         TabIndex        =   51
         Top             =   1057
         Width           =   1950
      End
      Begin VB.TextBox txtPayments 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -71055
         Locked          =   -1  'True
         TabIndex        =   50
         Top             =   1417
         Width           =   1950
      End
      Begin VB.TextBox txtCBBalance 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -71055
         Locked          =   -1  'True
         TabIndex        =   49
         Top             =   1792
         Width           =   1950
      End
      Begin VB.TextBox txtUnpresentedChq 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -71055
         Locked          =   -1  'True
         TabIndex        =   48
         Top             =   2572
         Width           =   1950
      End
      Begin VB.TextBox txtDeposits 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -71055
         Locked          =   -1  'True
         TabIndex        =   47
         Top             =   2962
         Width           =   1950
      End
      Begin VB.TextBox txtBankCredits 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -71055
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   3367
         Width           =   1950
      End
      Begin VB.TextBox txtBankDebits 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -71055
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   3757
         Width           =   1950
      End
      Begin VB.TextBox txtBankBal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -71055
         TabIndex        =   44
         Top             =   4162
         Width           =   1950
      End
      Begin VB.TextBox txtOpeningBal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -71055
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   682
         Width           =   1950
      End
      Begin VB.TextBox txtCrAccNo 
         Height          =   300
         Left            =   -74700
         TabIndex        =   42
         Top             =   4170
         Width           =   1350
      End
      Begin VB.TextBox txtCrAccName 
         Height          =   300
         Left            =   -73350
         TabIndex        =   41
         Top             =   4170
         Width           =   2760
      End
      Begin VB.TextBox txtDrNarration 
         Height          =   300
         Left            =   -70605
         TabIndex        =   40
         Top             =   4170
         Width           =   3045
      End
      Begin VB.TextBox txtDrAmount 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   -67575
         TabIndex        =   39
         Top             =   4170
         Width           =   1605
      End
      Begin VB.TextBox txtDrAccNo 
         Height          =   300
         Left            =   -74700
         TabIndex        =   38
         Top             =   4170
         Width           =   1350
      End
      Begin VB.TextBox txtDrAccName 
         Height          =   300
         Left            =   -73350
         TabIndex        =   37
         Top             =   4170
         Width           =   2760
      End
      Begin VB.TextBox txtCrNarration 
         Height          =   300
         Left            =   -70605
         TabIndex        =   36
         Top             =   4170
         Width           =   3045
      End
      Begin VB.TextBox txtCrAmount 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   -67575
         TabIndex        =   35
         Top             =   4170
         Width           =   1605
      End
      Begin VB.TextBox txtCrDocumentNo 
         Height          =   300
         Left            =   -65955
         TabIndex        =   33
         Top             =   4170
         Width           =   1590
      End
      Begin VB.TextBox txtDrDocumentNo 
         Height          =   300
         Left            =   -65955
         TabIndex        =   32
         Top             =   4170
         Width           =   1590
      End
      Begin VB.CommandButton cmdAddCredit 
         Caption         =   "&Add"
         Height          =   375
         Left            =   -70380
         TabIndex        =   30
         Top             =   4785
         Width           =   1380
      End
      Begin VB.CommandButton cmdAddDebits 
         Caption         =   "&Add"
         Height          =   375
         Left            =   -70380
         TabIndex        =   29
         Top             =   4785
         Width           =   1380
      End
      Begin VB.TextBox txtDifference 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -71055
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   4560
         Width           =   1950
      End
      Begin VB.CommandButton cmdDrPost 
         Caption         =   "&Post"
         Height          =   375
         Left            =   -68835
         TabIndex        =   26
         Top             =   4785
         Width           =   1380
      End
      Begin VB.CommandButton cmdCrPost 
         Caption         =   "&Post"
         Height          =   375
         Left            =   -68835
         TabIndex        =   25
         Top             =   4785
         Width           =   1380
      End
      Begin VB.CommandButton cmdPostRecon 
         Caption         =   "Post Reconciliation"
         Height          =   435
         Left            =   -66375
         TabIndex        =   24
         Top             =   5310
         Width           =   2010
      End
      Begin VB.Frame fraTransfer 
         Caption         =   "Transfer Funds"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2130
         Left            =   1770
         TabIndex        =   8
         Top             =   2520
         Visible         =   0   'False
         Width           =   6210
         Begin VB.CommandButton cmdCancel 
            Caption         =   "&Cancel"
            Height          =   390
            Left            =   3075
            TabIndex        =   17
            Top             =   1650
            Width           =   1425
         End
         Begin VB.CommandButton cmdTransfer 
            Caption         =   "Transfer"
            Height          =   390
            Left            =   1635
            TabIndex        =   16
            Top             =   1650
            Width           =   1425
         End
         Begin VB.TextBox txtAccName 
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
            Left            =   1530
            TabIndex        =   15
            Top             =   480
            Width           =   4305
         End
         Begin VB.TextBox txtAccNo 
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
            Left            =   45
            TabIndex        =   14
            Top             =   480
            Width           =   1335
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
            Left            =   1380
            TabIndex        =   12
            Top             =   1080
            Width           =   1335
         End
         Begin VB.TextBox txtMemberNo 
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
            Left            =   2745
            TabIndex        =   11
            Top             =   1080
            Width           =   1800
         End
         Begin VB.TextBox txtDocumentNo 
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
            Left            =   4575
            TabIndex        =   10
            Top             =   1080
            Width           =   1590
         End
         Begin MSComCtl2.DTPicker dtpTransDate 
            Height          =   315
            Left            =   60
            TabIndex        =   9
            Top             =   1080
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   " dd-MM-yyyy"
            Format          =   187695107
            CurrentDate     =   39549
         End
         Begin MSComctlLib.ListView ListView2 
            Height          =   1245
            Left            =   1545
            TabIndex        =   13
            Top             =   765
            Visible         =   0   'False
            Width           =   4290
            _ExtentX        =   7567
            _ExtentY        =   2196
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
               Text            =   "AccNo"
               Object.Width           =   18
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "AccName"
               Object.Width           =   10583
            EndProperty
         End
         Begin VB.Label lblAccName 
            AutoSize        =   -1  'True
            Caption         =   "Account Name"
            Height          =   195
            Left            =   1545
            TabIndex        =   23
            Top             =   240
            Width           =   1035
         End
         Begin VB.Label lblAccNo 
            AutoSize        =   -1  'True
            Caption         =   "Account No"
            Height          =   210
            Left            =   60
            TabIndex        =   22
            Top             =   240
            Width           =   975
         End
         Begin VB.Label lblAmount 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "New Amount"
            Height          =   195
            Left            =   1725
            TabIndex        =   21
            Top             =   840
            Width           =   915
         End
         Begin VB.Label lblMemberNo 
            AutoSize        =   -1  'True
            Caption         =   "MemberNo/Organisation"
            Height          =   195
            Left            =   2775
            TabIndex        =   20
            Top             =   855
            Width           =   1740
         End
         Begin VB.Label lblDocumentNo 
            AutoSize        =   -1  'True
            Caption         =   "Document No"
            Height          =   195
            Left            =   4590
            TabIndex        =   19
            Top             =   840
            Width           =   960
         End
         Begin VB.Label lblTransDate 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Transaction Date"
            Height          =   195
            Left            =   90
            TabIndex        =   18
            Top             =   840
            Width           =   1230
         End
      End
      Begin MSComctlLib.ListView lvwCrAccounts 
         Height          =   1320
         Left            =   -73350
         TabIndex        =   28
         Top             =   4470
         Visible         =   0   'False
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   2328
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "AccountName"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "AccNo"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComCtl2.DTPicker dtpDrTransDate 
         Height          =   330
         Left            =   -74700
         TabIndex        =   31
         Top             =   4845
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   " dd-MM-yyyy"
         Format          =   187695107
         CurrentDate     =   39407
      End
      Begin MSComctlLib.ListView lvwDrAccounts 
         Height          =   1320
         Left            =   -73365
         TabIndex        =   34
         Top             =   4470
         Visible         =   0   'False
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   2328
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "AccountName"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "AccNo"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lvwCredits 
         Height          =   3060
         Left            =   -74715
         TabIndex        =   52
         Top             =   765
         Width           =   10110
         _ExtentX        =   17833
         _ExtentY        =   5398
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "TransDate"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Narration"
            Object.Width           =   6703
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "AccNo to Debit"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Amount"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Document No"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   5445
         Left            =   0
         TabIndex        =   53
         Top             =   390
         Width           =   10740
         _ExtentX        =   18944
         _ExtentY        =   9604
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
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "TransDate"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "TransDescription"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Receipts"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Payments"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Balance"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Document No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "TransNo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Cheque No"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lvwDebits 
         Height          =   3060
         Left            =   -74715
         TabIndex        =   54
         Top             =   765
         Width           =   10110
         _ExtentX        =   17833
         _ExtentY        =   5398
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "TransDate"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Narration"
            Object.Width           =   6703
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "AccNo to Credit"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Amount"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Document No"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComCtl2.DTPicker dtpCrTransDate 
         Height          =   330
         Left            =   -74700
         TabIndex        =   55
         Top             =   4845
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   " dd-MM-yyyy"
         Format          =   187695107
         CurrentDate     =   39407
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Receipts ........................................................."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74385
         TabIndex        =   77
         Top             =   1110
         Width           =   3345
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Payments ......................................................"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74385
         TabIndex        =   76
         Top             =   1470
         Width           =   3330
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Cash Book Balance ......................................"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74385
         TabIndex        =   75
         Top             =   1845
         Width           =   3330
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "ADD: Unprecented Cheques ......................"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74385
         TabIndex        =   74
         Top             =   2625
         Width           =   3345
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "LESS: Deposits not yet credited ..............."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74385
         TabIndex        =   73
         Top             =   3015
         Width           =   3345
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "ADD: Bank Credits not in Cash Book ........"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74385
         TabIndex        =   72
         Top             =   3420
         Width           =   3360
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "LESS: Bank Debits not in Cash Book ........"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74385
         TabIndex        =   71
         Top             =   3810
         Width           =   3330
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Balance as per Bank Statement .............."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74385
         TabIndex        =   70
         Top             =   4215
         Width           =   3330
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Opening Balance .........................................."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74385
         TabIndex        =   69
         Top             =   735
         Width           =   3330
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Acc No to Credit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74685
         TabIndex        =   68
         Top             =   3945
         Width           =   1335
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Acc Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -72750
         TabIndex        =   67
         Top             =   3945
         Width           =   825
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Narration"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -70545
         TabIndex        =   66
         Top             =   3945
         Width           =   795
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -66675
         TabIndex        =   65
         Top             =   3945
         Width           =   675
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Acc No to Credit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74685
         TabIndex        =   64
         Top             =   3945
         Width           =   1335
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Acc Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -72750
         TabIndex        =   63
         Top             =   3945
         Width           =   825
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Narration"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -70545
         TabIndex        =   62
         Top             =   3945
         Width           =   795
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -66675
         TabIndex        =   61
         Top             =   3945
         Width           =   675
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Document No"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -65940
         TabIndex        =   60
         Top             =   3945
         Width           =   1125
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Document No"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -65940
         TabIndex        =   59
         Top             =   3945
         Width           =   1125
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Trans Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74580
         TabIndex        =   58
         Top             =   4620
         Width           =   930
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Trans Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74580
         TabIndex        =   57
         Top             =   4620
         Width           =   930
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Difference ......................................................"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74385
         TabIndex        =   56
         Top             =   4620
         Width           =   3345
      End
   End
   Begin MSComCtl2.DTPicker dtpStatDate 
      Height          =   300
      Left            =   3885
      TabIndex        =   82
      Top             =   300
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   " dd-MM-yyyy"
      Format          =   276168707
      CurrentDate     =   39406
   End
   Begin MSComCtl2.DTPicker dtpStartDate 
      Height          =   300
      Left            =   8295
      TabIndex        =   83
      Top             =   300
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   " dd-MM-yyyy"
      Format          =   276168707
      CurrentDate     =   39406
   End
   Begin MSComCtl2.DTPicker dtpFinishDate 
      Height          =   300
      Left            =   9600
      TabIndex        =   84
      Top             =   300
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   " dd-MM-yyyy"
      Format          =   276168707
      CurrentDate     =   39406
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Statement Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3780
      TabIndex        =   91
      Top             =   90
      Width           =   1365
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Opening Balance"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5235
      TabIndex        =   90
      Top             =   90
      Width           =   1395
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Closing Balance"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6915
      TabIndex        =   89
      Top             =   90
      Width           =   1305
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Start  Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8355
      TabIndex        =   88
      Top             =   90
      Width           =   930
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Finish Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   9600
      TabIndex        =   87
      Top             =   75
      Width           =   930
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      Caption         =   "Bank Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1215
      TabIndex        =   86
      Top             =   90
      Width           =   945
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      Caption         =   "Bank Code"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   45
      TabIndex        =   85
      Top             =   75
      Width           =   885
   End
End
Attribute VB_Name = "frmCashBookTransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim li As ListItem

Private Sub cmdAddCredit_Click()
    On Error GoTo SysError
    If Trim$(txtDrAccNo) = "" Then
        MsgBox "Please select an account to Debit", vbInformation, Me.Caption
        txtDrAccName.SetFocus
        Exit Sub
    End If
    If Trim$(txtDrAccName) = "" Then
        MsgBox "Please select an account to Debit", vbInformation, Me.Caption
        txtDrAccName.SetFocus
        Exit Sub
    End If
    If Trim$(txtCrNarration) = "" Then
        MsgBox "Please enter a description for the Transaction", vbInformation, Me.Caption
        txtCrNarration.SetFocus
        Exit Sub
    End If
    If Trim$(txtCrAmount) = "" Then
        MsgBox "Please enter an Amount.", vbInformation, Me.Caption
        txtCrAmount.SetFocus
        Exit Sub
    End If
    If Trim$(txtCrDocumentNo) = "" Then
        MsgBox "Please enter a Document No", vbInformation, Me.Caption
        txtCrDocumentNo.SetFocus
        Exit Sub
    End If
    Set li = lvwCredits.ListItems.Add(, , dtpCrTransDate)
    li.SubItems(1) = txtCrNarration
    li.SubItems(2) = txtDrAccNo
    li.SubItems(3) = Format(txtCrAmount, Cfmt)
    li.SubItems(4) = txtCrDocumentNo
    txtCrNarration = ""
    txtCrAmount = ""
    txtCrDocumentNo = ""
    txtDrAccName = ""
    txtDrAccNo = ""
    txtDrAccNo.SetFocus
    SendKeys "{Home}+{End}"
    Load_Statement
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub cmdAddDebits_Click()
    On Error GoTo SysError
    On Error GoTo SysError
    If Trim$(txtCrAccNo) = "" Then
        MsgBox "Please select an account to Credit", vbInformation, Me.Caption
        txtCrAccName.SetFocus
        Exit Sub
    End If
    If Trim$(txtCrAccName) = "" Then
        MsgBox "Please select an account to Credit", vbInformation, Me.Caption
        txtDrAccName.SetFocus
        Exit Sub
    End If
    If Trim$(txtDrNarration) = "" Then
        MsgBox "Please enter a description for the Transaction", vbInformation, Me.Caption
        txtDrNarration.SetFocus
        Exit Sub
    End If
    If Trim$(txtDrAmount) = "" Then
        MsgBox "Please enter an Amount.", vbInformation, Me.Caption
        txtDrAmount.SetFocus
        Exit Sub
    End If
    If Trim$(txtDrDocumentNo) = "" Then
        MsgBox "Please enter a Document No", vbInformation, Me.Caption
        txtDrDocumentNo.SetFocus
        Exit Sub
    End If
    Set li = lvwDebits.ListItems.Add(, , dtpDrTransDate)
    li.SubItems(1) = txtDrNarration
    li.SubItems(2) = txtCrAccNo
    li.SubItems(3) = Format(txtDrAmount, Cfmt)
    li.SubItems(4) = txtDrDocumentNo
    txtDrNarration = ""
    txtDrAmount = ""
    txtDrDocumentNo = ""
    txtCrAccName = ""
    txtCrAccNo = ""
    txtCrAccNo.SetFocus
    SendKeys "{Home}+{End}"
    Load_Statement
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub cmdAmount_Click()
    On Error GoTo SysError
    Dim DocNo As String, memberno As String, TransTyp As String, mCredit As String
    SearchForm = "Update Amount"
    If ListView1.ListItems.Count > 0 Then
        Set li = ListView1.SelectedItem
        Select Case CDbl(li.SubItems(3))
            Case 0
            txtamount = li.SubItems(2)
            Case Else
            txtamount = li.SubItems(3)
        End Select
        DocNo = li.SubItems(5)
        TransTyp = li.SubItems(6)
        txtDocumentNo = DocNo
        txtMemberNo = TransTyp
        txtamount.Visible = True
        txtMemberNo.Visible = True
        txtDocumentNo.Visible = True
        lblAmount.Visible = True
        lblDocumentNo.Visible = True
        lblMemberNo.Visible = True
        txtAccName.Visible = False
        Txtaccno.Visible = False
        lblaccname.Visible = False
        lblAccNo.Visible = False
        fraTransfer.Visible = True
    End If
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub cmdcancel_Click()
    fraTransfer.Visible = False
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdCrPost_Click()
    On Error GoTo SysError
    Dim dAmount As Double, DRaccno As String, DocNo As String, sNarration As String
    For I = 1 To lvwCredits.ListItems.Count
        dAmount = CDbl(lvwCredits.ListItems(I).SubItems(3))
        DRaccno = lvwCredits.ListItems(I).SubItems(2)
        DocNo = lvwCredits.ListItems(I).SubItems(4)
        sNarration = lvwCredits.ListItems(I).SubItems(1)
        mTransDate = lvwCredits.ListItems(I)
        If Not Save_To_GL(DRaccno, "A001", dAmount, DocNo, DocNo, mTransDate, DocNo, _
        sNarration, ErrorMessage, "Bank Rec") Then
            If ErrorMessage <> "" Then
                MsgBox ErrorMessage, vbInformation, Me.Caption
                ErrorMessage = ""
            End If
        End If
    Next
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub cmdDrPost_Click()
    On Error GoTo SysError
    Dim dAmount As Double, DRaccno As String, DocNo As String, sNarration As String
    For I = 1 To lvwDebits.ListItems.Count
        dAmount = CDbl(lvwDebits.ListItems(I).SubItems(3))
        DRaccno = lvwDebits.ListItems(I).SubItems(2)
        DocNo = lvwDebits.ListItems(I).SubItems(4)
        sNarration = lvwDebits.ListItems(I).SubItems(1)
        mTransDate = lvwDebits.ListItems(I)
        If Not Save_To_GL("A001", DRaccno, dAmount, DocNo, DocNo, mTransDate, DocNo, _
        sNarration, ErrorMessage, "Bank Rec") Then
            If ErrorMessage <> "" Then
                MsgBox ErrorMessage, vbInformation, Me.Caption
                ErrorMessage = ""
            End If
        End If
    Next
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub cmdExport_Click()
    On Error GoTo SsyError
    Dim MyFso As New FileSystemObject, strData As String, MFile As TextStream, _
    FileName As String, I As Long, li As ListItem
    If ListView1.ListItems.Count > 0 Then
        With CommonDialog1
            .Filter = "Comma Seperated Values|*.csv"
            .ShowSave
            If .FileName <> "" Then
                FileName = .FileName
            End If
            .FileName = ""
        End With
        Set MFile = MyFso.OpenTextFile(FileName, ForWriting, True)
        strData = "TransDate,MemberNo,Names,ag_receipts,Payments,Document No"
        MFile.WriteLine strData
        strData = ""
        For I = 1 To ListView1.ListItems.Count
            Set li = ListView1.ListItems(I)
            strData = li & "," & li.SubItems(6) & "," & li.SubItems(1) & "," & CDbl(li.SubItems(2)) _
            & "," & CDbl(li.SubItems(3)) & "," & li.SubItems(5)
            MFile.WriteLine strData
            strData = ""
        Next I
    Else
        MsgBox "There are no records to be exported", vbInformation, Me.Caption
    End If
    Exit Sub
SsyError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub cmdLoad_Click()
    'On Error GoTo SysError
    Dim rsRecon As New Recordset, BankBal As Double, bCredits As Double, bDebits As Double, _
    RsDesc As New Recordset
    If Trim(cboBank) = "" Then
        MsgBox "Please indicate the Bank Account to Reconcile", vbInformation, Me.Caption
        txtBankName.SetFocus
        SendKeys "{Home}+{End}"
        Exit Sub
    End If
    If dtpStartDate > dtpFinishDate Then
        MsgBox "The StartDate should be Earlier than the FinishDate", vbInformation, Me.Caption
        Exit Sub
    End If
    Get_GL_AccDetails cboBank
    If dtpStartDate < EarliestTransDate Then
        MsgBox "The StartDate should  not be earlier than " & Format(EarliestTransDate, _
        "dd-MM-yyyy"), vbInformation, Me.Caption
        dtpStartDate.SetFocus
        'Exit Sub
    End If
    
    BankBal = GlAccBalance
    'Set rsRecon = oSaccoMaster.GetRecordSet("Set DateFormat DMY Select Sum(Amount) as Amount " _
    & "From CUSTOMERBALANCE where AccNo='A001' and TransType='DR' and TransDate<'" & dtpStartDate _
    & "' and TransDate>'31/10/2007'")
    Set rsRecon = oSaccoMaster.GetRecordset("Set DateFormat DMY Select Sum(Amount) as Amount " _
    & "From CashbookTransaction where AccNo='" & glaccno & "' and TransType='DR' and TransDate<'" & dtpStartDate _
    & "'")
    With rsRecon
        If .State = adStateOpen Then
            If Not .EOF Then
                BankBal = BankBal + IIf(IsNull(!amount), 0, !amount)
            End If
        End If
    End With
    'Set rsRecon = oSaccoMaster.GetRecordSet("Set DateFormat DMY Select Sum(Amount) as Amount " _
    & "From CUSTOMERBALANCE where AccNo='A001' and TransType='CR' and TransDate<'" & _
    dtpStartDate & "' and TransDate>'31/10/2007'")
    Set rsRecon = oSaccoMaster.GetRecordset("Set DateFormat DMY Select Sum(Amount) as Amount " _
    & "From CashbookTransaction where AccNo='" & glaccno & "' and TransType='CR' and TransDate<'" & _
    dtpStartDate & "'")
    With rsRecon
        If .State = adStateOpen Then
            If Not .EOF Then
                BankBal = BankBal - IIf(IsNull(!amount), 0, !amount)
            End If
        End If
    End With
    txtOpeningBal = Format(BankBal, Cfmt)
    Set rsRecon = oSaccoMaster.GetRecordset("Set DateFormat DMY Select TransDate,Sum(Amount) " _
    & "as Amount,TransType,ChequeNo,VNo,transby From CashbookTransaction where AccNo='" & glaccno _
    & "' and TransDate>='" & dtpStartDate & "' and TransDate<='" & dtpFinishDate & "' and VNo<>'' " _
    & "and IDNo<>'2' and Reconciled=0 group by TransDate,ChequeNo,VNo,TransType,transby order by TransDate,chequeno")
    'set rsRecon=oSaccoMaster.GetRecordSet(
    ListView1.ListItems.Clear
    With rsRecon
        While Not .EOF
            DoEvents
            Set li = ListView1.ListItems.Add(, , !transdate)
            li.Checked = True
            Set RsDesc = oSaccoMaster.GetRecordset("Select Name From d_COMPANY")
           
            With RsDesc
                If .State = adStateOpen Then
                    If Not .EOF Then
                        li.SubItems(1) = IIf(IsNull(!name), "", !name)
                    Else
                        .Close
                        Set RsDesc = oSaccoMaster.GetRecordset("Select names," _
                        & " sno From d_suppliers where sno='" & rsRecon!vno & "'")
                        With RsDesc
                            If .State = adStateOpen Then
                                If Not .EOF Then
                                    li.SubItems(1) = IIf(IsNull(!NAMES), "", !NAMES)
                                Else
                                    li.SubItems(1) = IIf(IsNull(rsRecon!sno), "", rsRecon!sno)
                                End If
                            End If
                        End With
                    End If
                End If
            End With
            li.SubItems(2) = Format(IIf(!transtype <> "DR", 0, !amount), Cfmt)
            bDebits = bDebits + CDbl(li.SubItems(2))
            li.SubItems(3) = Format(IIf(!transtype <> "CR", 0, !amount), Cfmt)
            bCredits = bCredits + CDbl(li.SubItems(3))
            BankBal = BankBal + CDbl(li.SubItems(2)) - CDbl(li.SubItems(3))
            li.SubItems(4) = Format(BankBal, Cfmt)
            li.SubItems(5) = IIf(IsNull(!chequeno), "", !chequeno)
            li.SubItems(6) = IIf(IsNull(!vno), "", !vno)
            li.SubItems(7) = IIf(IsNull(!transby), "", !transby)
            .MoveNext
        Wend
    End With
    txtReceipts = Format(bDebits, Cfmt)
    txtPayments = Format(bCredits, Cfmt)
    txtCBBalance = Format(BankBal, Cfmt)
    Load_Statement
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub Load_Statement()
    On Error GoTo SysEror
    Dim UnprecCheques As Double, BankDebits As Double, BankCredits As Double, _
    UnDebs As Double
    For I = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(I).Checked = False Then
            If CDbl(ListView1.ListItems(I).SubItems(3)) <> 0 Then
                UnprecCheques = UnprecCheques + CDbl(ListView1.ListItems(I).SubItems(3))
            End If
            If CDbl(ListView1.ListItems(I).SubItems(2)) <> 0 Then
                UnDebs = UnDebs + CDbl(ListView1.ListItems(I).SubItems(2))
            End If
        Else
        End If
    Next I
    For I = 1 To lvwDebits.ListItems.Count
        If CDbl(lvwDebits.ListItems(I).SubItems(3)) <> 0 Then
            BankDebits = BankDebits + CDbl(lvwDebits.ListItems(I).SubItems(3))
        End If
    Next
    For I = 1 To lvwCredits.ListItems.Count
        If CDbl(lvwCredits.ListItems(I).SubItems(3)) <> 0 Then
            BankCredits = BankCredits + CDbl(lvwCredits.ListItems(I).SubItems(3))
        End If
    Next
    txtUnpresentedChq = Format(UnprecCheques, Cfmt)
    txtBankCredits = Format(BankCredits, Cfmt)
    txtBankDebits = Format(BankDebits, Cfmt)
    txtDeposits = Format(UnDebs, Cfmt)
    Exit Sub
SysEror:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub cmdrefresh_Click()
    On Error GoTo SysError
    Dim UnprecCheques As Double
    For I = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(I).Checked = False Then
            If CDbl(ListView1.ListItems(I).SubItems(2)) <> 0 Then
                UnprecCheques = UnprecCheques + CDbl(ListView1.ListItems(I).SubItems(2))
            End If
        Else
        End If
    Next I
    txtUnpresentedChq = Format(UnprecCheques, Cfmt)
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdOffset_Click()
    On Error GoTo SysError
    Dim FromAcc As String, ToAcc As String, DocumentNo As String, TransT As String, _
    transdate As Date, mCredit As String
    Dim rsTransfer As New Recordset
    If MsgBox("Do you want to Change the selected Transaction to an Offset?", _
    vbQuestion + vbYesNo, Me.Caption) = vbNo Then
        Exit Sub
    End If
    If MsgBox("This process is not reversible. Do you want to continue?", _
    vbExclamation + vbYesNo, Me.Caption) = vbNo Then
        Exit Sub
    End If
    If ListView1.ListItems.Count > 0 Then
        Set li = ListView1.SelectedItem
        transdate = li
        DocumentNo = CStr(li.SubItems(5))
        TransT = CStr(li.SubItems(6))
        Select Case CDbl(li.SubItems(3)) 'PAYMENTS
            Case 0 'XXXXXXXXXXX ag_receipts XXXXXXXXX
            mCredit = "DR"
            Case Else 'XXXXXXXX PAYMENTS XXXXXXXXX
            mCredit = "CR"
        End Select
        Set rsTransfer = oSaccoMaster.GetRecordset("Set DateFormat DMY Update CUSTOMERBALANCE" _
        & " Set IDNo='2' where AccNo='" & cboBank & "' and VNo='" & TransT & _
        "' and ChequeNo='" & DocumentNo & "' and TransDate='" & transdate & "' and TransType='" _
        & mCredit & "'")
        cmdLoad_Click
    End If
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub cmdPostRecon_Click()
    Dim rsRecon As New Recordset
    On Error GoTo SysError
    'XXXXXXXXXXX Update CashBook Entries XXXXXXXXXXXXXXXXXXXX'
    For I = 1 To ListView1.ListItems.Count
        mTransDate = ListView1.ListItems(I)
'        If ListView1.ListItems(I).Checked = True Then
'            Set rsRecon = oSaccoMaster.GetRecordset("Set DateFormat DMY Update " _
'            & "CUSTOMERBALANCE Set Reconcile=1 where AccNo='A001' and TransDate='" _
'            & mTransDate & "' and VNo='" & ListView1.ListItems(I).SubItems(1) _
'            & "' and ChequeNo='" & ListView1.ListItems(I).SubItems(5) & "'")
'        End If
    Next
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub cmdRefresh_data_Click()
On Error GoTo hell

mysql = "delete  from CashbookTransaction"
oSaccoMaster.ExecuteThis (mysql)
mysql = "Cash_Book_Non_member_Transaction '" & dtpStartDate & "','" & dtpFinishDate & "'"
oSaccoMaster.ExecuteThis (mysql)
cmdLoad_Click
Exit Sub
hell:
MsgBox err.description
End Sub

Private Sub cmdTransfer_Click()
    On Error GoTo SysError
    Dim FromAcc As String, ToAcc As String, DocumentNo As String, TransT As String, _
    transdate As Date, mCredit As String, mAmount As Double
    Dim rsTransfer As New Recordset
    Select Case SearchForm
        Case "Transfer"
        If ListView1.ListItems.Count > 0 Then
            Set li = ListView1.SelectedItem
            transdate = li
            DocumentNo = CStr(li.SubItems(5))
            TransT = CStr(li.SubItems(6))
            Select Case CDbl(li.SubItems(3)) 'PAYMENTS
                Case 0 'ag_receipts
                mCredit = "DR"
                mAmount = CDbl(li.SubItems(2))
                Case Else 'PAYMENTS
                mCredit = "CR"
                mAmount = CDbl(li.SubItems(3))
            End Select
            Set rsTransfer = oSaccoMaster.GetRecordset("Set DateFormat DMY Update CUSTOMERBALANCE" _
            & " Set AccNo='" & Txtaccno & "' where AccNo='" & cboBank & "' and VNo='" & TransT & _
            "' and ChequeNo='" & DocumentNo & "' and TransDate='" & transdate & "' and TransType='" _
            & mCredit & "'")
            fraTransfer.Visible = False
            cmdLoad_Click
        End If
        Case "Update Amount"
        If ListView1.ListItems.Count > 0 Then
            Set li = ListView1.SelectedItem
            DTPTransdate = Format(li, " dd-MM-yyyy")
            transdate = Format(li, "dd-MM-yyyy")
            DocumentNo = CStr(li.SubItems(5))
            TransT = CStr(li.SubItems(6))
            Select Case CDbl(li.SubItems(3)) 'PAYMENTS
                Case 0 'ag_receipts
                mCredit = "DR"
                mAmount = CDbl(li.SubItems(2))
                'txtAmount = li.SubItems(4)
                Case Else 'PAYMENTS
                mCredit = "CR"
                mAmount = CDbl(li.SubItems(3))
                'txtAmount = li.SubItems(3)
            End Select
            Set rsTransfer = oSaccoMaster.GetRecordset("Set DateFormat DMY Update CUSTOMERBALANCE" _
            & " Set Amount='" & txtamount & "' where VNo='" & TransT & "' and ChequeNo='" & DocumentNo _
            & "' and TransDate='" & DTPTransdate & "' and Amount=" & mAmount)
            fraTransfer.Visible = False
            cmdLoad_Click
        End If
    End Select
    lblAmount.Visible = False
    txtamount.Visible = False
    txtDocumentNo.Visible = False
    txtMemberNo.Visible = False
    lblMemberNo.Visible = False
    lblDocumentNo.Visible = False
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub cmdTransferFunds_Click()
    On Error GoTo SysError
    'frmTransferFunds.Show , Me
    SearchForm = "Transfer"
    txtAccName.Visible = True
    Txtaccno.Visible = True
    lblaccname.Visible = True
    lblAccNo.Visible = True
    txtamount.Visible = False
    txtDocumentNo.Visible = False
    txtMemberNo.Visible = False
    lblMemberNo.Visible = False
    lblDocumentNo.Visible = False
    lblAmount.Visible = False
    fraTransfer.Visible = True
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub Form_Load()
    On Error GoTo SysError
    dtpStatDate = Format(Get_Server_Date, " dd-MM-yyyy")
    dtpStartDate = Format(Get_Server_Date, " dd-MM-yyyy")
    dtpFinishDate = Format(Get_Server_Date, " dd-MM-yyyy")
    dtpDrTransDate = Format(Get_Server_Date, " dd-MM-yyyy")
    dtpCrTransDate = Format(Get_Server_Date, " dd-MM-yyyy")
    
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub ListView1_Click()
    Load_Statement
End Sub

Private Sub listview1_DblClick()
    If ListView1.ListItems.Count > 0 Then
        Set li = ListView1.SelectedItem
        mTransDate = CDate(li)
        TransNo = li.SubItems(6)
        mDocNo = li.SubItems(5)
    End If
    frmLedgers.Show vbModal, Me
End Sub

Private Sub ListView2_Click()
    On Error GoTo SysError
    If ListView2.ListItems.Count > 0 Then
        Set li = ListView2.SelectedItem
        Txtaccno = li
        txtAccName = li.SubItems(1)
        lvwAccounts.ListItems.Clear
        lvwAccounts.Visible = False
    End If
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub lvwAccounts_Click()
    On Error GoTo SysError
    If lvwAccounts.ListItems.Count > 0 Then
        Set li = lvwAccounts.SelectedItem
        cboBank = li
        txtBankName = li.SubItems(1)
        lvwAccounts.ListItems.Clear
        lvwAccounts.Visible = False
    End If
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub lvwCrAccounts_Click()
    On Error GoTo SysError
    If lvwCrAccounts.ListItems.Count > 0 Then
        txtCrAccName = lvwCrAccounts.SelectedItem
        txtCrAccNo = lvwCrAccounts.SelectedItem.SubItems(1)
    End If
    lvwCrAccounts.Visible = False
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub lvwDrAccounts_Click()
    On Error GoTo SysError
    If lvwDrAccounts.ListItems.Count > 0 Then
        txtDrAccName = lvwDrAccounts.SelectedItem
        txtDrAccNo = lvwDrAccounts.SelectedItem.SubItems(1)
    End If
    lvwDrAccounts.Visible = False
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub txtAccName_Change()
'    On Error GoTo SysError
'    Dim rsAccount As New Recordset
'    ListView2.ListItems.Clear
'    If Trim$(txtAccName) <> "" Then
'        If Not Editing Then
'            Set rsAccount = oSaccoMaster.GetRecordset("Select * From GLSETUP where " _
'            & "GLAccName like '%" & txtAccName & "%'")
'            With rsAccount
'                If .State = adStateOpen Then
'                    If Not .EOF Then
'                        ListView2.Visible = True
'                        While Not .EOF
'                            Set li = ListView2.ListItems.Add(, , IIf(IsNull(!accno), "", !accno))
'                            li.SubItems(1) = IIf(IsNull(!GlAccName), "", !GlAccName)
'                            .MoveNext
'                        Wend
'                        If ListView2.ListItems.Count = 1 Then
'                            txtAccNo = li
'                            txtAccName = li.SubItems(1)
'                            ListView2.ListItems.Clear
'                            ListView2.Visible = False
'                        End If
'                    Else
'                        ListView2.Visible = False
'                    End If
'                End If
'            End With
'        End If
'    Else
'        ListView2.Visible = False
'    End If
'    Exit Sub
'SysError:
'    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub txtAmount_KeyPress(KeyAscii As Integer)
    On Error GoTo SysError
    Select Case KeyAscii
        Case 48 To 57
        Case Is = 46
        Case Is = 8
        Case Else
        KeyAscii = 0
    End Select
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub txtBankName_Change()
    'On Error GoTo SysError
    Dim rsAccount As New Recordset
    lvwAccounts.ListItems.Clear
    If Trim$(txtBankName) <> "" Then
        If Not Editing Then
            Set rsAccount = oSaccoMaster.GetRecordset("Select * From GLSETUP where " _
            & "GLAccName like '%" & txtBankName & "%'")
            With rsAccount
                If .State = adStateOpen Then
                    If Not .EOF Then
                        lvwAccounts.Visible = True
                        While Not .EOF
                            Set li = lvwAccounts.ListItems.Add(, , IIf(IsNull(!ACCNO), "", !ACCNO))
                            li.SubItems(1) = IIf(IsNull(!GlAccName), "", !GlAccName)
                            .MoveNext
                        Wend
                        If lvwAccounts.ListItems.Count = 1 Then
                            cboBank = li
                            txtBankName = li.SubItems(1)
                            lvwAccounts.ListItems.Clear
                            lvwAccounts.Visible = False
                        End If
                    Else
                        lvwAccounts.Visible = False
                    End If
                End If
            End With
        End If
    Else
        lvwAccounts.Visible = False
    End If
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub txtCrAccName_Change()
    On Error GoTo SysError
    Dim rsAcc As New Recordset
    lvwCrAccounts.ListItems.Clear
    If Trim$(txtCrAccName) <> "" Then
        Set rsAcc = oSaccoMaster.GetRecordset("Select AccNo,GLAccName From GLSETUP " _
        & "where GLACCName like '%" & txtCrAccName & "%'")
        With rsAcc
            If Not .EOF Then
                lvwCrAccounts.Visible = True
                While Not .EOF
                    Set li = lvwCrAccounts.ListItems.Add(, , !GlAccName)
                    li.SubItems(1) = IIf(IsNull(!ACCNO), "", !ACCNO)
                    .MoveNext
                Wend
            Else
                lvwCrAccounts.Visible = False
            End If
        End With
    Else
        lvwCrAccounts.Visible = False
    End If
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub txtCrAmount_KeyPress(KeyAscii As Integer)
    On Error GoTo SysError
    Select Case KeyAscii
        Case 48 To 57
        Case Is = 46
        Case Is = 8
        Case 13
        If Trim$(txtCrAmount) Then
            Set li = lvwCredits.ListItems.Add(, , dtpCrTransDate)
            li.SubItems(1) = txtCrNarration
            li.SubItems(2) = txtDrAccNo
            li.SubItems(3) = Format(txtCrAmount, Cfmt)
            li.SubItems(4) = txtCrDocumentNo
        End If
        Case Else
        KeyAscii = 0
    End Select
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub txtDrAccName_Change()
    On Error GoTo SysError
    Dim rsAcc As New Recordset
    lvwDrAccounts.ListItems.Clear
    If Trim$(txtDrAccName) <> "" Then
        Set rsAcc = oSaccoMaster.GetRecordset("Select AccNo,GLAccName From GLSETUP " _
        & "where GLACCName like '%" & txtDrAccName & "%'")
        With rsAcc
            If Not .EOF Then
                lvwDrAccounts.Visible = True
                While Not .EOF
                    Set li = lvwDrAccounts.ListItems.Add(, , !GlAccName)
                    li.SubItems(1) = IIf(IsNull(!ACCNO), "", !ACCNO)
                    .MoveNext
                Wend
            Else
                lvwDrAccounts.Visible = False
            End If
        End With
    Else
        lvwDrAccounts.Visible = False
    End If
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub txtDrAmount_KeyPress(KeyAscii As Integer)
    On Error GoTo SysError
    Select Case KeyAscii
        Case 48 To 57
        Case Is = 46
        Case Is = 8
        Case 13
'        If Trim(txtDrAmount) <> "" Then
'            Set li = lvwCredits.ListItems.Add(, , dtpDrTransDate)
'            li.SubItems(1) = txtDrNarration
'            li.SubItems(2) = txtDrAccNo
'            li.SubItems(3) = Format(txtDrAmount, CfMt)
'            li.SubItems(4) = txtDrDocumentNo
'        End If
        Case Else
        KeyAscii = 0
    End Select
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub Load_Cash_Transaction(Startdate As Date, Enddate As Date)

mysql = ""
mysql = "delete from CashbookTransaction"
oSaccoMaster.ExecuteThis (mysql)
mysql = "Get_CashBook_transaction '" & Startdate & "','" & Enddate & "'"

oSaccoMaster.ExecuteThis (mysql)

End Sub


