VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmcashreciept 
   Caption         =   "Cash Reciept"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12210
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   12210
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   3465
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   11895
      Begin VB.TextBox txttchpbalance 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8880
         TabIndex        =   51
         Top             =   3000
         Width           =   2175
      End
      Begin VB.TextBox txtTCHPMonthlyPremium 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6480
         TabIndex        =   45
         Top             =   3000
         Width           =   2175
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3840
         TabIndex        =   44
         Top             =   3000
         Width           =   2535
      End
      Begin VB.TextBox txtSNo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1920
         TabIndex        =   43
         Top             =   3000
         Width           =   1335
      End
      Begin VB.CommandButton cmdfind 
         Caption         =   "F"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3240
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   3000
         Width           =   375
      End
      Begin VB.ComboBox cboreceiptpurpose 
         Height          =   315
         ItemData        =   "frmcashreciept.frx":0000
         Left            =   1920
         List            =   "frmcashreciept.frx":000D
         TabIndex        =   41
         Text            =   "General"
         Top             =   2520
         Width           =   2295
      End
      Begin VB.ComboBox cbobrnch 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmcashreciept.frx":0028
         Left            =   7560
         List            =   "frmcashreciept.frx":002A
         TabIndex        =   38
         Top             =   360
         Width           =   3015
      End
      Begin VB.Frame Frame 
         ClipControls    =   0   'False
         Height          =   1695
         Left            =   225
         TabIndex        =   17
         Top             =   690
         Width           =   11415
         Begin VB.CommandButton cmdnew 
            Caption         =   "New"
            Height          =   330
            Left            =   8085
            TabIndex        =   29
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txtcontra 
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
            Left            =   2130
            TabIndex        =   28
            Top             =   780
            Width           =   1410
         End
         Begin VB.PictureBox Picture4 
            Height          =   285
            Left            =   3525
            Picture         =   "frmcashreciept.frx":002C
            ScaleHeight     =   225
            ScaleWidth      =   240
            TabIndex        =   27
            Top             =   780
            Width           =   300
         End
         Begin VB.TextBox TxtDRAmount 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   2130
            Locked          =   -1  'True
            MaxLength       =   9
            TabIndex        =   26
            Text            =   "0"
            Top             =   1215
            Width           =   1410
         End
         Begin VB.TextBox txtAmountDue 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   8100
            Locked          =   -1  'True
            MaxLength       =   9
            TabIndex        =   25
            Top             =   780
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.Frame Frame7 
            Height          =   570
            Left            =   2130
            TabIndex        =   20
            Top             =   120
            Width           =   5895
            Begin VB.TextBox txtChequeno 
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
               Left            =   4065
               TabIndex        =   22
               Top             =   195
               Visible         =   0   'False
               Width           =   1695
            End
            Begin VB.TextBox txtReceiptsno 
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
               Left            =   1200
               TabIndex        =   21
               Top             =   195
               Width           =   1935
            End
            Begin VB.Label lblVoucher 
               AutoSize        =   -1  'True
               Caption         =   "Voucher No"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   210
               TabIndex        =   24
               Top             =   225
               Width           =   870
            End
            Begin VB.Label lblCheque 
               AutoSize        =   -1  'True
               Caption         =   "Cheque No"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   3240
               TabIndex        =   23
               Top             =   240
               Visible         =   0   'False
               Width           =   795
            End
         End
         Begin VB.OptionButton Optcash 
            Caption         =   "Cash"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   19
            Top             =   300
            Width           =   855
         End
         Begin VB.OptionButton Optcheque 
            Caption         =   "Cheque"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   1080
            TabIndex        =   18
            Top             =   300
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            Caption         =   "Account Status"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   3735
            TabIndex        =   34
            Top             =   1245
            Width           =   1125
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Cash Receipts Account"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   330
            TabIndex        =   33
            Top             =   825
            Width           =   1710
         End
         Begin VB.Label lblcontra 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   3825
            TabIndex        =   32
            Top             =   780
            Width           =   4170
         End
         Begin VB.Label Labal 
            Caption         =   "Avaliable Amount"
            Height          =   255
            Left            =   750
            TabIndex        =   31
            Top             =   1260
            Width           =   1335
         End
         Begin VB.Label LblStatus 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   300
            Left            =   4905
            TabIndex        =   30
            Top             =   1200
            Width           =   1320
         End
      End
      Begin MSComCtl2.DTPicker dtptransdate 
         Height          =   315
         Left            =   4875
         TabIndex        =   35
         Top             =   390
         Width           =   1365
         _ExtentX        =   2408
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
         CurrentDate     =   39954
      End
      Begin VB.Label Label11 
         Caption         =   "TCHP Balance"
         Height          =   255
         Left            =   9000
         TabIndex        =   52
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "Supplier No."
         Height          =   255
         Left            =   240
         TabIndex        =   48
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "TCHP Monthly Premium"
         Height          =   255
         Left            =   6360
         TabIndex        =   47
         Top             =   2520
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "Name:"
         Height          =   255
         Left            =   4320
         TabIndex        =   46
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Receipt Purpose"
         Height          =   255
         Left            =   240
         TabIndex        =   40
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Branch"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6480
         TabIndex        =   39
         Top             =   360
         Width           =   840
      End
      Begin VB.Label Label1 
         Caption         =   "CASH RECEIPTS"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   495
         Left            =   240
         TabIndex        =   37
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Transaction Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3585
         TabIndex        =   36
         Top             =   450
         Width           =   1230
      End
   End
   Begin VB.Frame FraOtherpayment 
      Caption         =   "Transaction Details"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   3525
      Width           =   11895
      Begin VB.CommandButton cmdprintreceipts1 
         Caption         =   "Print Receipts"
         Height          =   375
         Left            =   7680
         TabIndex        =   53
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txttchpmonths 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   10320
         TabIndex        =   49
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdpost 
         Caption         =   "Post"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3420
         TabIndex        =   10
         Top             =   1320
         Width           =   1140
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6210
         TabIndex        =   9
         Top             =   1305
         Width           =   1215
      End
      Begin VB.CommandButton cmdsave 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2160
         TabIndex        =   8
         Top             =   1320
         Width           =   1110
      End
      Begin VB.TextBox TxtOtherPayment 
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
         Left            =   7560
         MaxLength       =   9
         TabIndex        =   7
         Text            =   "0"
         Top             =   240
         Width           =   1215
      End
      Begin VB.PictureBox Picture21 
         Height          =   285
         Left            =   2685
         Picture         =   "frmcashreciept.frx":02EE
         ScaleHeight     =   225
         ScaleWidth      =   240
         TabIndex        =   6
         Top             =   225
         Width           =   300
      End
      Begin VB.TextBox TxtOtherPAcc 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "##-##-####"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
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
         Left            =   1320
         TabIndex        =   5
         Top             =   210
         Width           =   1305
      End
      Begin VB.TextBox txtNarration 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "##-##-####"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
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
         Left            =   3000
         TabIndex        =   4
         ToolTipText     =   "The person who is taking the money"
         Top             =   600
         Width           =   3510
      End
      Begin VB.CommandButton cmdprint 
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4920
         TabIndex        =   3
         Top             =   1320
         Width           =   1140
      End
      Begin VB.TextBox txtnarations 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "##-##-####"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
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
         Left            =   2970
         TabIndex        =   2
         Top             =   960
         Width           =   8310
      End
      Begin VB.CheckBox chkperiodicreceipts 
         Caption         =   "Print Period Vouchers"
         Height          =   255
         Left            =   9240
         TabIndex        =   1
         Top             =   1320
         Width           =   2055
      End
      Begin MSComctlLib.ListView lvwTrans 
         Height          =   2535
         Left            =   120
         TabIndex        =   11
         Top             =   1680
         Width           =   11685
         _ExtentX        =   20611
         _ExtentY        =   4471
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
         NumItems        =   13
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
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Sno"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Type"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Balance"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label10 
         Caption         =   "TCHP Months"
         Height          =   255
         Left            =   9120
         TabIndex        =   50
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   4
         Left            =   6915
         TabIndex        =   15
         Top             =   270
         Width           =   555
      End
      Begin VB.Label lblOtherPaymentAcc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   3000
         TabIndex        =   14
         Top             =   210
         Width           =   3495
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Receipts From:"
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
         Left            =   1725
         TabIndex        =   13
         Top             =   645
         Width           =   1245
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Description:"
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
         Left            =   1950
         TabIndex        =   12
         Top             =   1005
         Width           =   990
      End
   End
End
Attribute VB_Name = "frmcashreciept"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DRaccno As String
Dim Craccno As String
Dim IntAccNo As String
Dim glmemno As String
Dim glnamE1 As String
Dim amount As Currency
Dim ReceiptNo As String
Dim mysql As String
Dim loanno1 As String
Dim loanno2 As String
Dim loanno3 As String
Dim loanno4 As String
Dim loanno5 As String
Dim loanno6 As String
Dim loanno7 As String



Private Sub cboreceiptpurpose_Change()
If cboreceiptpurpose = "Shares" Then
Label10.Visible = False
txttchpmonths.Visible = False
TxtOtherPayment.Locked = False
txtSNo.Locked = False
cmdfind.Enabled = True
ElseIf cboreceiptpurpose = "General" Then
Label10.Visible = False
txttchpmonths.Visible = False
txtSNo.Locked = True
cmdfind.Enabled = False
Else
Label10.Visible = True
txttchpmonths.Visible = True
TxtOtherPayment.Locked = True
txtSNo.Locked = False
cmdfind.Enabled = True
End If
End Sub

Private Sub cboreceiptpurpose_Click()
cboreceiptpurpose_Change
End Sub

Private Sub cmdFind_Click()
        Me.MousePointer = vbHourglass
        frmSearchSupplier.Show vbModal
        txtSNo = sel
        txtSNo_Validate True
        Me.MousePointer = 0
End Sub

Private Sub cmdNew_Click()
Dim rsr As New ADODB.Recordset
Dim rsg As New ADODB.Recordset
Dim I As Object
Dim Mylength As Integer
'//if this record is new then look for receipts no

''//clear all textboxes





mysql = ""
mysql = "select GenerateReceiptno from param"

Set rsg = oSaccoMaster.GetRecordset(mysql)
If Not rsg.EOF Then
    ''''check check
    If rsg!GenerateReceiptno = True Then
    
        mysql = ""
        mysql = "select * from Receiptno where receiptno like 'PC-%' order by Receipthnoid desc"
        
        Set rsr = oSaccoMaster.GetRecordset(mysql)
        
        If Not rsr.EOF Then
            Mylength = CInt(Mid(rsr!ReceiptNo, 5, 10))
            Mylength = Mylength + 1
            txtReceiptsno = Padding(Mylength)
            txtReceiptsno = "PC-" & txtReceiptsno
        Else
            Mylength = 1
            txtReceiptsno = "PC-" & Padding(Mylength)
            
        End If
Else
    ''//receiptno  will be keyed in
End If
End If

End Sub

Private Sub cmdPost_Click()
    On Error GoTo SysError
    If Check_Period_If_Closed(DTPtransdate) = True Then
         Exit Sub
     End If
    Dim Cubaccount As Cub_Acc_Details
    Dim balance As Double, ctype As String, sno As String
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
    ' If lvwTrans.ListItems.Item(I).Checked = True Then
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
        If txtSNo <> "NA" Then
        balance = 0 ' lvwTrans.ListItems(I).SubItems(12)
        ctype = lvwTrans.ListItems(I).SubItems(11)
        sno = lvwTrans.ListItems(I).SubItems(10)
        Else
        balance = 0
        ctype = "General"
        sno = "NA"
        End If
'IDENTI = lvwTrans.ListItems(i).SubItems(9)
        CashBook = 1
        If DocumentNo = "" Then DocumentNo = "CB"

                If chknonmemberpostings = vbChecked Then
                doc_posted = 1
                Else
                doc_posted = 0
                End If
        
        Set rs = oSaccoMaster.GetRecordset("sp_chequeno_used '" & Trim(chequeno) & "','" & Trim(TransSource) & "'")
        If Not rs.EOF Then

         End If

       
        
        If Not Save_GLTRANSACTION(transdate, amount, DRaccno, Craccno, DocumentNo, _
        TransSource, User, ErrorMessage, transDescription, CashBook, doc_posted, chequeno, TransNo) Then
            If ErrorMessage <> "" Then
                MsgBox ErrorMessage, vbInformation, Me.Caption
                ErrorMessage = ""
            End If
        End If
        
        
         sql = " INSERT INTO PettyCash"
         sql = sql & "             (transdate, AccName, Pvcno, Amount, Naration, auditid,ctype,sno,balance)"
         sql = sql & "  VALUES     ('" & transdate & "','" & TransSource & "','" & DocumentNo & "'," & amount & ",'" & transDescription & "','" & User & "','" & ctype & "','" & sno & "'," & balance & ")"
         oSaccoMaster.ExecuteThis (sql)
        
       ' End If
    Next I
    
    '//clear listview
    mysql = ""
    mysql = "set dateformat dmy Insert into Receiptno(Receiptno,Auditdate,auditid)values('" & txtReceiptsno & "','" & Format(Get_Server_Date, "dd/MM/yyyy") & "','" & User & "')"
    oSaccoMaster.ExecuteThis (mysql)
    
    '//INSERT INTO TCHP IF IT IS A TCHP
Dim txtTCHPBalances As Double
If cboreceiptpurpose = "TCHP" Then
sql = "SELECT     balance   FROM         tchp_trxs  WHERE     sno ='" & txtSNo & "'  ORDER BY transdate DESC, id DESC "
Dim rr As New ADODB.Recordset
Set rr = oSaccoMaster.GetRecordset(sql)
If Not rr.EOF Then
txtTCHPBalances = rr.Fields(0)
End If
'balance = txtTCHPBalances

sql = ""
sql = "set dateformat dmy INSERT INTO tchp_trxs"
sql = sql & "     (sno,transdate, description, Debits, CreditsD, CreditsC, Balance, auditid)"
sql = sql & " VALUES     ('" & txtSNo & "','" & DTPtransdate & "','Cash Receipt',0,0," & amount & "," & balance & ",'" & User & "')"
oSaccoMaster.ExecuteThis (sql)
End If

    lvwTrans.ListItems.Clear
    If cboreceiptpurpose = "General" Then
    txtSNo = ""
    txtTCHPMonthlyPremium = ""
    txtName = ""
    GoTo mwisho
    Else
    End If
    Me.MousePointer = vbDefault
    
    '********************shares contribution********************************
    
  
    If UCase(cboreceiptpurpose) = UCase("Shares") Then
    'get the idno for supplier
    
    Dim rr2 As New ADODB.Recordset, idno As String
    sql = "SELECT     idno   FROM         d_suppliers  WHERE     sno ='" & txtSNo & "'"
    Set rr2 = oSaccoMaster.GetRecordset(sql)
    If Not rr2.EOF Then
    idno = IIf(IsNull(rr2.Fields(0)), 0, rr2.Fields(0))
    End If
    '//get the balance
    Set rst = New ADODB.Recordset
    sql = "select bal from d_shares where sno= '" & txtSNo & "'"
    Set rst = oSaccoMaster.GetRecordset(sql)
    If Not rst.EOF Then
    txtTCHPBalances = rst.Fields(0)
    
'     '//get the balance
'
'    sql = "SELECT     bal   FROM         d_sconribution  WHERE     sno ='" & txtSNo & "'  ORDER BY transdate DESC, id DESC "
'
'    Set rr = oSaccoMaster.GetRecordset(sql)
'    If Not rr.EOF Then
    txtTCHPBalances = txtTCHPBalances + CCur(amount)
    ',[sno],[transdate],[amount],[bal],[transdescription],[auditid],[auditdate],[mno]
      'From [EASYTEA].[dbo].[d_sconribution]
      sql = ""
      sql = "set dateformat dmy insert into d_sconribution([sno],[transdate],[amount],[bal],[transdescription],[auditid])"
      sql = sql & " values ('" & txtSNo & "','" & DTPtransdate & "'," & amount & "," & txtTCHPBalances & ",'Shares-cash','" & User & "') "
      oSaccoMaster.ExecuteThis (sql)
      
      'UPDATE SHARES BALANCE
      sql = ""
      sql = "update d_shares set bal=" & txtTCHPBalances & " where sno='" & txtSNo & "' "
      oSaccoMaster.ExecuteThis (sql)
    'txtTCHPBALANCE = rr.Fields(0)
    End If
    Else
    '//add new one
    txtTCHPBalances = txtTCHPBalances + CCur(amount)
    sql = "insert into d_Shares(sno, Cash,bal,idno,auditid)"
    sql = sql & " values('" & txtSNo & "',1,'" & idno & "', " & amount & ",'" & User & "')"
    oSaccoMaster.ExecuteThis (sql)
    sql = ""
    sql = "set dateformat dmy insert into d_sconribution([sno],[transdate],[amount],[bal],[transdescription],[auditid])"
    sql = sql & " values ('" & txtSNo & "','" & DTPtransdate & "'," & amount & "," & amount & ",'Shares-Cash','" & User & "') "
    oSaccoMaster.ExecuteThis (sql)
    
    'UPDATE SHARES BALANCE
     sql = ""
      sql = "update d_shares set bal=" & txtTCHPBalances & " where sno='" & txtSNo & "' "
      oSaccoMaster.ExecuteThis (sql)
    End If
    
'End If
    '**************************end*******************************************************
mwisho:
    MsgBox "Posting Successfull", vbInformation, Me.Caption
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
    Me.MousePointer = vbDefault

End Sub

Private Sub cmdPrint_Click()
'//pettycashvoucher
If chkperiodicreceipts = vbChecked Then
 'STRFORMULA = "{PettyCash.Pvcno}='" & txtReceiptsno & "'"
    reportname = "receiptsvouchers.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, title
'//periodics
Else
    STRFORMULA = "{PettyCash.Pvcno}='" & txtReceiptsno & "'"
    reportname = "pettycashvoucher99.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, title
'//periodics
'pettycashvoucherperiodic
End If
End Sub

Private Sub cmdprintreceipts1_Click()
'cashreceiptstchp
'STRFORMULA = "{PettyCash.Pvcno}='" & txtReceiptsno & "'"
    reportname = "cashreceiptstchp.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, title
End Sub

Private Sub cmdsave_Click()
    On Error GoTo SysError
    If Trim$(CCur(TxtOtherPayment)) > CCur(TxtDRAmount) Then
       ' MsgBox "You do not have sufficient Amount in Petty Cash Account", vbInformation, Me.Caption
        'Exit Sub
    End If
    
    If Trim(txtChequeno) = "" Then
       ' MsgBox "Please Enter The chequne No", vbInformation, Me.Caption
       ' Exit Sub
    End If
    '// PLEASE TOP UP YOUR IMPREST
    If TxtDRAmount < 0 Then
    MsgBox "Please top up your imprest amount", vbCritical
    Exit Sub
    End If
    
    If Val(TxtOtherPayment) = 0 Then
        MsgBox "Please enter a valid Amount", vbInformation, Me.Caption
        TxtOtherPayment.SetFocus
        Exit Sub
    End If
    If Trim$(cboreceiptpurpose) = "" Then
        MsgBox "Please select  the receipt purpose to continue.", vbInformation, Me.Caption
        cboreceiptpurpose.SetFocus
        Exit Sub
    End If
    
    If Trim$(txtcontra) = "" Then
        MsgBox "Please enter the Account to Debit.", vbInformation, Me.Caption
        txtDrAccNo.SetFocus
        SendKeys "{Home}+{End}"
        Exit Sub
    End If
    If txtnarations = "" Then
    MsgBox "Please enter the naration", vbCritical
    Exit Sub
    End If
    If Trim$(TxtOtherPAcc) = "" Then
        MsgBox "Please enter the Account to Credit.", vbInformation, Me.Caption
        txtCrAccNo.SetFocus
        SendKeys "{Home}+{End}"
        Exit Sub
    End If
    
    If Trim$(txtReceiptsno) = "" Then
        MsgBox "Please enter the Transaction Source", vbInformation, Me.Caption
        txtSource.SetFocus
        SendKeys "{Home}+{End}"
        Exit Sub
    End If
    If Trim$(txtnarration) = "" Then
        MsgBox "Please enter the Transaction Description", vbInformation, Me.Caption
        txtnarration.SetFocus
        SendKeys "{Home}+{End}"
        Exit Sub
    End If
    If txtSNo = "" Then txtSNo = "NA"
    If txtSNo <> "NA" Then
'    txttchpbalance = CDbl(txttchpbalance) - CDbl(TxtOtherPayment)
    End If
   TxtDRAmount = CCur(TxtDRAmount) - CCur(TxtOtherPayment)
    Set li = lvwTrans.ListItems.Add(, , DTPtransdate)
    li.SubItems(1) = Format(CDbl(TxtOtherPayment), "#,##0.00")
    li.SubItems(3) = TxtOtherPAcc
    li.SubItems(2) = txtcontra
    li.SubItems(4) = txtReceiptsno
    li.SubItems(5) = txtnarations & "-" & (lblOtherPaymentAcc) & "-" & TxtOtherPAcc
    li.SubItems(6) = txtnarration
    li.SubItems(7) = 1
    li.SubItems(8) = txtChequeno
    li.SubItems(10) = txtSNo
    li.SubItems(11) = cboreceiptpurpose
    li.SubItems(12) = txttchpbalance
    TxtOtherPayment = "0"
    TxtOtherPAcc = ""
    
    'txtReceiptsno = ""
    'txtNarration = ""
    txtnarations = ""
    
    
    Exit Sub
    'lblDrAc = ""
    lblOtherPaymentAcc = ""
    'txtchequeno = ""
SysError:
    MsgBox err.description, vbInformation, Me.Caption

End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim RsLoancode As New ADODB.Recordset
    Dim RsScheme As New ADODB.Recordset
    Dim rscompany As New ADODB.Recordset
    'optCash_Click
    DTPtransdate.value = Format(Get_Server_Date, "dd/MM/yyyy")
    'get_availbalance
    'Headers
    'fraLoanRepayment.Visible = True
    'Frashares.Visible = True
    FraOtherpayment.Visible = True
    'cboreceiptpurpose = ""
    'optCash_Click
    'Load_Data
    
    cboreceiptpurpose_Change
    
    txtChequeno.Visible = False
    lblVoucher.Visible = True
    Set rs = CreateObject("adodb.recordset")
    
    rs.Open "SELECT BName FROM d_Branch", cn
    
    If rs.EOF Then Exit Sub
    
    With rs
        
        While Not .EOF
         
         cbobrnch.AddItem rs.Fields(0)
         
         
         .MoveNext
        
        Wend
    
    End With
End Sub



Private Sub Optcash_Click()
If Optcash = True Then
        txtcontra = ""
        lblcontra = GetLedgerDesc(txtcontra)
        txtChequeno.Visible = False
        lblCheque.Visible = False
        lblVoucher.Visible = True
    End If
End Sub

Private Sub Picture21_Click()
frmSearchGLAccounts.Show vbModal, Me
    If Continue Then
        If SearchValue <> "" Then
            TxtOtherPAcc = SearchValue
            SearchValue = ""
        End If
    End If
End Sub

Private Sub Picture4_Click()
frmSearchGLAccounts.Show vbModal, Me
    If Continue Then
        If SearchValue <> "" Then
            txtcontra = SearchValue
            SearchValue = ""
        End If
    End If
End Sub

Private Sub txtcontra_Change()
    On Error GoTo SysError
    Dim Account As Acc_Details
    Editing = True
    Account = Get_Acc_Details(txtcontra, ErrorMessage)
    If Account.AccNo <> "" Then
        lblcontra = Account.AccName
    Else
        If ErrorMessage <> "" Then
            MsgBox ErrorMessage, vbInformation, Me.Caption
            ErrorMessage = ""
        End If
        lblcontra = ""
    End If
    '//GET THE BALANCE AMOUNT
mysql = "delete  from GLTRANSACTIONS2"
oSaccoMaster.ExecuteThis (mysql)
 MousePointer = vbHourglass
 
 'TxtDRAmount = getGlCurrentBalance(txtcontra)
 
 
 '// Get Opening Balances
'mysql = ""
'mysql = "Get_OpeningBalances '30/12/2009'"
'oSaccoMaster.ExecuteThis (mysql)
'
''//Get Non-Member Transactions
'Dim bal As Currency, CR As Currency, DR As Currency
'bal = 0
'mysql = ""
'mysql = "Get_Non_member_Transaction '01/01/2013','" & Format(Get_Server_Date, "dd/MM/yyyy") & "'"
'oSaccoMaster.ExecuteThis (mysql)
'sql = "SELECT     SUM(Amount) AS a, Transtype   FROM         GLTRANSACTIONS2   WHERE     (Accno ='" & txtcontra & "')   GROUP BY Transtype order by transtype desc"
'Set rs = oSaccoMaster.GetRecordset(sql)
'While Not rs.EOF
'If rs.Fields(1) = "DR" Then DR = rs.Fields(0)
'If rs.Fields(1) = "CR" Then CR = rs.Fields(0)
'bal = rs.Fields(0) - bal
'
'rs.MoveNext
'Wend
'If DR > CR Then
'TxtDRAmount = Abs(bal)
'Else
'TxtDRAmount = (bal * -1)
'End If

MousePointer = vbNormal
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption

End Sub

Private Sub TxtOtherPAcc_Change()
 On Error GoTo SysError
    Dim Account As Acc_Details
    Editing = True
    Account = Get_Acc_Details(TxtOtherPAcc, ErrorMessage)
    If Account.AccNo <> "" Then
        lblOtherPaymentAcc = Account.AccName
    Else
        If ErrorMessage <> "" Then
            MsgBox ErrorMessage, vbInformation, Me.Caption
            ErrorMessage = ""
        End If
        lblOtherPaymentAcc = ""
    End If
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption

End Sub

Private Sub txtSNo_Validate(Cancel As Boolean)
On Error GoTo ErrorHandler
Dim tchpa As Integer
Set rs = New ADODB.Recordset
sql = "d_sp_SelectSuppliers '" & txtSNo & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
If Not IsNull(rs.Fields(2)) Then txtName.Text = rs.Fields(2)
Else
txtName.Text = ""
End If
''tchp_tchpmember
''SELECT     sno, aarno, mpremium, premium, tchpactive,balance    FROM         tchp_members where sno=@sno
Set rst = New ADODB.Recordset
sql = "tchp_tchpmember '" & txtSNo & "'"
Set rst = oSaccoMaster.GetRecordset(sql)
If Not rst.EOF Then
txtTCHPMonthlyPremium = rst.Fields(2)
End If
 '//get the balance
 Dim txtTCHPBalances As Double
sql = "SELECT     balance   FROM         tchp_trxs  WHERE     sno ='" & txtSNo & "'  ORDER BY transdate DESC, id DESC "
Dim rr As New ADODB.Recordset
Set rr = oSaccoMaster.GetRecordset(sql)
If Not rr.EOF Then
txtTCHPBalances = rr.Fields(0)
txttchpbalance = txtTCHPBalances
txtTCHPMonthlyPremium.Locked = True
txtTCHPMonthlyPremium.Enabled = False
txttchpbalance.Locked = True
txttchpbalance.Enabled = False
If cboreceiptpurpose = "TCHP" Then
TxtOtherPayment.Locked = True
TxtOtherPayment.Enabled = False
Else
TxtOtherPayment.Locked = False
TxtOtherPayment.Enabled = True
End If

'txtTCHPBALANCE = rr.Fields(0)
End If

Exit Sub
ErrorHandler:
MsgBox err.description


End Sub

Private Sub txttchpmonths_Change()
On Error GoTo ErrorHandler
If txtTCHPMonthlyPremium = "" Then txtTCHPMonthlyPremium = 0
If txttchpmonths = "" Then txttchpmonths = 0
TxtOtherPayment = CDbl(txtTCHPMonthlyPremium) * CDbl(txttchpmonths)
ErrorHandler:
End Sub
