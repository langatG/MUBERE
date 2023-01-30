VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmmembertransactions 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Transactions"
   ClientHeight    =   7755
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   10890
   Icon            =   "frmmembertransactions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   10890
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame frasalk 
      BackColor       =   &H80000001&
      Height          =   5820
      Left            =   1185
      TabIndex        =   20
      Top             =   -5715
      Width           =   9465
      Begin VB.TextBox TXTPASSWORD 
         Appearance      =   0  'Flat
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2880
         PasswordChar    =   " "
         TabIndex        =   23
         Top             =   2520
         Width           =   5295
      End
      Begin VB.CommandButton cmdcancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   3840
         TabIndex        =   22
         Top             =   4680
         Width           =   855
      End
      Begin VB.CommandButton cmdok1 
         Caption         =   "GO"
         Default         =   -1  'True
         Height          =   375
         Left            =   6000
         TabIndex        =   21
         Top             =   4680
         Width           =   855
      End
      Begin VB.Label Label5 
         BackColor       =   &H0000C000&
         Caption         =   " ENTER YOUR PASSWORD HERE"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   24
         Top             =   1800
         Width           =   4335
      End
   End
   Begin VB.CommandButton cmdclose 
      Appearance      =   0  'Flat
      Caption         =   "&Close"
      Height          =   345
      Left            =   9240
      TabIndex        =   19
      Top             =   7365
      Width           =   1245
   End
   Begin VB.TextBox TXTTGNO 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2400
      TabIndex        =   4
      Top             =   105
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Txtaccno 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2400
      MaxLength       =   15
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox txtpayno 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   2400
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox txtidno 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   2400
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.ComboBox Cbodetail 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmmembertransactions.frx":0442
      Left            =   120
      List            =   "frmmembertransactions.frx":0452
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5655
      Left            =   0
      TabIndex        =   11
      Top             =   1680
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   9975
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Journal Entry -Others"
      TabPicture(0)   =   "frmmembertransactions.frx":0481
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdreverse"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdsave"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Cmdjv"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label16"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label9"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Actual Transactions"
      TabPicture(1)   =   "frmmembertransactions.frx":049D
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdTransSetup"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdrefresh"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmdadj"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmddelete"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdprintstatement"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label17"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "Transactions Details"
      TabPicture(2)   =   "frmmembertransactions.frx":04B9
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "lvememtrans"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "fraTrans"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      Begin VB.CommandButton cmdTransSetup 
         Caption         =   "Trans Setup"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   -67005
         TabIndex        =   105
         Top             =   5070
         Width           =   1440
      End
      Begin VB.Frame fraTrans 
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Caption         =   "Transaction Details"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   1275
         TabIndex        =   100
         Top             =   2265
         Visible         =   0   'False
         Width           =   8190
         Begin VB.CommandButton cmdTrans 
            Caption         =   "&Close"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   7020
            TabIndex        =   104
            Top             =   2445
            Width           =   1080
         End
         Begin MSComctlLib.ListView lvwTrans 
            Height          =   2325
            Left            =   45
            TabIndex        =   101
            Top             =   45
            Width           =   8100
            _ExtentX        =   14288
            _ExtentY        =   4101
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483624
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Account Name"
               Object.Width           =   7056
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   1
               Text            =   "Debit Amount"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Text            =   "Credit Amount"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Description"
               Object.Width           =   3528
            EndProperty
         End
         Begin VB.Label lblCredit 
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
            Height          =   255
            Left            =   5445
            TabIndex        =   103
            Top             =   2430
            Width           =   1440
         End
         Begin VB.Label lblDebit 
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
            Height          =   255
            Left            =   3975
            TabIndex        =   102
            Top             =   2430
            Width           =   1440
         End
      End
      Begin VB.CommandButton cmdreverse 
         Appearance      =   0  'Flat
         Caption         =   "&Reversals"
         Height          =   350
         Left            =   -71280
         TabIndex        =   90
         Top             =   4290
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton cmdrefresh 
         Caption         =   "Refresh"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   -70725
         TabIndex        =   79
         Top             =   5070
         Width           =   1020
      End
      Begin VB.Frame Frame1 
         Height          =   4455
         Left            =   -74760
         TabIndex        =   45
         Top             =   450
         Width           =   10335
         Begin VB.TextBox lblglstamp 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2520
            TabIndex        =   96
            Top             =   4005
            Width           =   1440
         End
         Begin VB.TextBox LBLGLCOMMISSION 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2520
            TabIndex        =   95
            Top             =   3555
            Width           =   1440
         End
         Begin VB.TextBox LBLGLCONTRA 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2520
            TabIndex        =   94
            Top             =   3120
            Width           =   1440
         End
         Begin MSComCtl2.DTPicker DTP 
            Height          =   300
            Left            =   2520
            TabIndex        =   92
            Top             =   2227
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   529
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   " dd-MM-yyyy"
            Format          =   121831427
            CurrentDate     =   38943
         End
         Begin VB.Frame Frame4 
            Caption         =   "Other Charges"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2175
            Left            =   6360
            TabIndex        =   70
            Top             =   810
            Width           =   3975
            Begin VB.PictureBox Picture5 
               Height          =   255
               Left            =   2160
               Picture         =   "frmmembertransactions.frx":04D5
               ScaleHeight     =   195
               ScaleWidth      =   195
               TabIndex        =   83
               Top             =   360
               Width           =   255
            End
            Begin VB.PictureBox Picture6 
               Height          =   255
               Left            =   2160
               Picture         =   "frmmembertransactions.frx":0797
               ScaleHeight     =   195
               ScaleWidth      =   195
               TabIndex        =   82
               Top             =   840
               Width           =   255
            End
            Begin VB.PictureBox Picture7 
               Height          =   255
               Left            =   2160
               Picture         =   "frmmembertransactions.frx":0A59
               ScaleHeight     =   195
               ScaleWidth      =   195
               TabIndex        =   81
               Top             =   1320
               Width           =   255
            End
            Begin VB.PictureBox Picture8 
               Height          =   255
               Left            =   2160
               Picture         =   "frmmembertransactions.frx":0D1B
               ScaleHeight     =   195
               ScaleWidth      =   195
               TabIndex        =   80
               Top             =   1800
               Width           =   255
            End
            Begin VB.ComboBox cbocharge4 
               Height          =   315
               Left            =   1080
               TabIndex        =   74
               Top             =   1800
               Width           =   1095
            End
            Begin VB.ComboBox cbocharge3 
               Height          =   315
               Left            =   1080
               TabIndex        =   73
               Top             =   1320
               Width           =   1095
            End
            Begin VB.ComboBox cbocharge2 
               Height          =   315
               Left            =   1080
               TabIndex        =   72
               Top             =   840
               Width           =   1095
            End
            Begin VB.ComboBox cbocharge1 
               Height          =   315
               Left            =   1080
               TabIndex        =   71
               Top             =   360
               Width           =   1095
            End
            Begin VB.Label cbocharge1glaccno 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   285
               Left            =   2400
               TabIndex        =   87
               Top             =   360
               Width           =   1455
            End
            Begin VB.Label cbocharge2accno 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   285
               Left            =   2400
               TabIndex        =   86
               Top             =   840
               Width           =   1455
            End
            Begin VB.Label cbocharge3accno 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   285
               Left            =   2400
               TabIndex        =   85
               Top             =   1320
               Width           =   1455
            End
            Begin VB.Label cbocharge4accno 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   285
               Left            =   2400
               TabIndex        =   84
               Top             =   1800
               Width           =   1455
            End
            Begin VB.Label Label25 
               Caption         =   "B/Chq"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   78
               Top             =   1800
               Width           =   975
            End
            Begin VB.Label Label24 
               Caption         =   "Counter"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   77
               Top             =   1320
               Width           =   975
            End
            Begin VB.Label Label23 
               Caption         =   "IOUCHQS/L"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   76
               Top             =   840
               Width           =   975
            End
            Begin VB.Label Label20 
               Caption         =   "IOA/U"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   75
               Top             =   360
               Width           =   975
            End
         End
         Begin VB.TextBox txtstamp 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   7560
            TabIndex        =   69
            Text            =   "2"
            Top             =   4005
            Width           =   495
         End
         Begin VB.PictureBox Picture3 
            Height          =   255
            Left            =   2040
            Picture         =   "frmmembertransactions.frx":0FDD
            ScaleHeight     =   195
            ScaleWidth      =   315
            TabIndex        =   64
            Top             =   4020
            Width           =   375
         End
         Begin VB.PictureBox Picture2 
            Height          =   255
            Left            =   2040
            Picture         =   "frmmembertransactions.frx":129F
            ScaleHeight     =   195
            ScaleWidth      =   315
            TabIndex        =   63
            Top             =   3570
            Width           =   375
         End
         Begin VB.PictureBox Picture1 
            Height          =   255
            Left            =   2040
            Picture         =   "frmmembertransactions.frx":1561
            ScaleHeight     =   195
            ScaleWidth      =   315
            TabIndex        =   58
            Top             =   3135
            Width           =   375
         End
         Begin VB.TextBox txtvoucherno 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2520
            MaxLength       =   50
            ScrollBars      =   3  'Both
            TabIndex        =   51
            Top             =   1342
            Width           =   3135
         End
         Begin VB.Frame Frame3 
            Caption         =   "Transaction Type"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   48
            Top             =   240
            Width           =   6855
            Begin VB.OptionButton Optdebit 
               Caption         =   "Debit"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   4440
               TabIndex        =   50
               Top             =   180
               Width           =   855
            End
            Begin VB.OptionButton optcredit 
               Caption         =   "Credit"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   2400
               TabIndex        =   49
               Top             =   180
               Width           =   1575
            End
         End
         Begin VB.TextBox txtadjamnt 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2520
            MaxLength       =   10
            TabIndex        =   47
            Top             =   1777
            Width           =   3135
         End
         Begin VB.ComboBox cbomemtrans 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmmembertransactions.frx":1823
            Left            =   2535
            List            =   "frmmembertransactions.frx":1825
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   46
            Top             =   885
            Width           =   3135
         End
         Begin MSComCtl2.DTPicker DTPadj 
            Height          =   300
            Left            =   2520
            TabIndex        =   93
            Top             =   2662
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   529
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   " dd-MM-yyyy"
            Format          =   121831427
            CurrentDate     =   38943
         End
         Begin VB.Label glnamestamp1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   4035
            TabIndex        =   68
            Top             =   4005
            Width           =   3525
         End
         Begin VB.Label glnamecom1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   4035
            TabIndex        =   67
            Top             =   3555
            Width           =   3525
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "GL Com"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   480
            TabIndex        =   66
            Top             =   3600
            Width           =   540
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "GL Stamp"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   480
            TabIndex        =   65
            Top             =   4050
            Width           =   675
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "GL Contra Account"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   480
            TabIndex        =   60
            Top             =   3165
            Width           =   1350
         End
         Begin VB.Label glnamE1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   4035
            TabIndex        =   59
            Top             =   3120
            Width           =   3525
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Date Posted"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   480
            TabIndex        =   56
            Top             =   2715
            Width           =   885
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Voucher Number/RefNo"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   480
            TabIndex        =   55
            Top             =   1395
            Width           =   1695
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Transaction Date"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   480
            TabIndex        =   54
            Top             =   2280
            Width           =   1230
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Amount "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   480
            TabIndex        =   53
            Top             =   1830
            Width           =   600
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Narration/ Description"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   480
            TabIndex        =   52
            Top             =   945
            Width           =   1575
         End
      End
      Begin VB.Frame Frame2 
         Height          =   3495
         Left            =   -74760
         TabIndex        =   25
         Top             =   570
         Width           =   8535
         Begin VB.PictureBox Picture10 
            Height          =   255
            Left            =   1335
            Picture         =   "frmmembertransactions.frx":1827
            ScaleHeight     =   195
            ScaleWidth      =   210
            TabIndex        =   107
            Top             =   1605
            Width           =   270
         End
         Begin VB.PictureBox Picture9 
            Height          =   255
            Left            =   1335
            Picture         =   "frmmembertransactions.frx":1AE9
            ScaleHeight     =   195
            ScaleWidth      =   210
            TabIndex        =   106
            Top             =   690
            Width           =   270
         End
         Begin VB.TextBox txtnarationcr 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1605
            MaxLength       =   50
            TabIndex        =   61
            Top             =   2040
            Width           =   4935
         End
         Begin VB.TextBox txtcomm 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   6840
            TabIndex        =   33
            Text            =   "0"
            Top             =   2955
            Width           =   975
         End
         Begin VB.OptionButton Optdr 
            Caption         =   "Debit"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1800
            TabIndex        =   32
            Top             =   270
            Width           =   1560
         End
         Begin VB.OptionButton OpCR 
            Caption         =   "Credits"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   31
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox Txtamount 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1605
            MaxLength       =   8
            TabIndex        =   30
            Top             =   2955
            Width           =   2895
         End
         Begin VB.TextBox Txtvou 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1605
            MaxLength       =   20
            TabIndex        =   29
            Top             =   2505
            Width           =   4935
         End
         Begin VB.TextBox txtreason 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1605
            MaxLength       =   50
            TabIndex        =   28
            Top             =   1125
            Width           =   4935
         End
         Begin VB.TextBox acccr 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1605
            MaxLength       =   14
            TabIndex        =   27
            Top             =   1590
            Width           =   1635
         End
         Begin VB.TextBox accdr 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1605
            MaxLength       =   14
            TabIndex        =   26
            Top             =   675
            Width           =   1635
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Naration Cr"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   62
            Top             =   2085
            Width           =   825
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Commission/Charges"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   5175
            TabIndex        =   43
            Top             =   2985
            Width           =   1485
         End
         Begin VB.Label lblamount2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   6480
            TabIndex        =   42
            Top             =   1590
            Width           =   1935
         End
         Begin VB.Label lblamount1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   6480
            TabIndex        =   41
            Top             =   675
            Width           =   1935
         End
         Begin VB.Label Lblname2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   3330
            TabIndex        =   40
            Top             =   1590
            Width           =   3030
         End
         Begin VB.Label lblname1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   3330
            TabIndex        =   39
            Top             =   675
            Width           =   3030
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Amount."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   38
            Top             =   3000
            Width           =   615
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Voucher Number."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   37
            Top             =   2550
            Width           =   1245
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Naration Dr"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   36
            Top             =   1170
            Width           =   825
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Account Cr."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   35
            Top             =   1635
            Width           =   855
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Account Dr."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   34
            Top             =   720
            Width           =   855
         End
      End
      Begin VB.CommandButton cmdadj 
         Appearance      =   0  'Flat
         Caption         =   "&Commit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   -72075
         TabIndex        =   16
         Top             =   5070
         Width           =   1020
      End
      Begin VB.CommandButton cmddelete 
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   -73410
         TabIndex        =   15
         Top             =   5070
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.CommandButton cmdsave 
         Appearance      =   0  'Flat
         Caption         =   "&Save"
         Height          =   375
         Left            =   -69120
         TabIndex        =   14
         Top             =   4230
         Width           =   975
      End
      Begin VB.CommandButton cmdprintstatement 
         Caption         =   "Edit Member Transactions"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   -69390
         TabIndex        =   13
         Top             =   5070
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CommandButton Cmdjv 
         Caption         =   "Jv"
         Height          =   375
         Left            =   -74040
         TabIndex        =   12
         Top             =   4290
         Width           =   975
      End
      Begin MSComctlLib.ListView lvememtrans 
         Height          =   4935
         Left            =   120
         TabIndex        =   17
         Top             =   510
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   8705
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         OLEDragMode     =   1
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
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
         OLEDragMode     =   1
         NumItems        =   0
      End
      Begin VB.Label Label17 
         BackColor       =   &H80000015&
         Height          =   4215
         Left            =   -74640
         TabIndex        =   57
         Top             =   810
         Width           =   10335
      End
      Begin VB.Label Label16 
         BackColor       =   &H80000015&
         Height          =   3495
         Left            =   -74880
         TabIndex        =   44
         Top             =   570
         Width           =   8655
      End
      Begin VB.Label Label9 
         Caption         =   "Label9"
         Height          =   255
         Left            =   -68160
         TabIndex        =   18
         Top             =   630
         Width           =   975
      End
   End
   Begin VB.PictureBox Picture4 
      Height          =   285
      Left            =   4440
      Picture         =   "frmmembertransactions.frx":1DAB
      ScaleHeight     =   225
      ScaleWidth      =   195
      TabIndex        =   91
      Top             =   120
      Width           =   255
   End
   Begin VB.CheckBox chkbackofficetransactions 
      Caption         =   "Back Office Transactions"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   97
      Top             =   1320
      Width           =   2535
   End
   Begin VB.TextBox txtmemberno 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2400
      TabIndex        =   98
      Top             =   1275
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblmemberno 
      AutoSize        =   -1  'True
      Caption         =   "Member No"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   105
      TabIndex        =   99
      Top             =   1320
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Label lbluncleared 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   6585
      TabIndex        =   89
      Top             =   945
      Width           =   1830
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      Caption         =   "Uncleared Effects"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5280
      TabIndex        =   88
      Top             =   1005
      Width           =   1275
   End
   Begin VB.Label lblname 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   5280
      TabIndex        =   10
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label lblaccname 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   5280
      TabIndex        =   9
      Top             =   525
      Width           =   3135
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Current Balance"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   105
      TabIndex        =   8
      Top             =   570
      Width           =   1155
   End
   Begin VB.Label lblavail 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2400
      TabIndex        =   7
      Top             =   525
      Width           =   2055
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Available Balance"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   105
      TabIndex        =   6
      Top             =   945
      Width           =   1245
   End
   Begin VB.Label lblbookbalance 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   2400
      TabIndex        =   5
      Top             =   900
      Width           =   2055
   End
End
Attribute VB_Name = "frmmembertransactions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myclass As cdbase
Dim ulmt As Integer
Dim number As Integer
Dim principal As Currency
Dim ref As Integer
Dim interest1 As Currency
Dim currentbalance As Currency
Dim remi As Double
Dim expectedamount1 As Currency, expected2 As Currency, expected3 As Currency, expected4 As Currency, expected5 As Currency
Dim remainder1 As Double, remainder2 As Double, remainder3 As Double, remainder4 As Double, remainder5 As Double
Dim remainder6 As Double, remainder7 As Double

Public Event CloseControl(bExit As Boolean)
Dim rsd As Object
Dim AccName As String
Dim custno As String
Dim desc As String
Dim lblacname As String
Dim withcharges As Currency
Dim totalcharges As Currency
Dim charge1 As Currency
Dim charge2 As Currency
Dim charge3 As Currency
Dim charge4 As Currency
Dim minBal As Currency
Dim AVAIL1 As Currency
Dim custno1 As String
Dim idno1 As String
Dim payno1 As String
Dim name1 As String
Dim teller As String
Dim accname1 As String
Dim AVAIL2 As Currency
Dim custno2 As String
Dim idno2 As String
Dim payno2 As String
Dim name2 As String
Dim accname2 As String
Dim glnamE As String 'FOR CONTRA
Dim glidno As String 'FOR CONTRA
Dim glmemno As String 'FOR CONTRA
Dim glpayno As String 'FOR CONTRA
Dim bookba As Currency
Dim bookba1 As Currency
Dim bookba2 As Currency
Dim bookba3 As Currency
Dim glcomm As String 'FOR CONTRA
Dim glaccno As String
Dim authorisecomm As Currency
Dim glnamecom As String 'FOR COMMISSION
Dim glcommemno As String 'FOR COMMISSION
Dim glcomidno As String 'FOR COMMISSION
Dim glcompayno As String 'FOR COMMISSION
Dim glcommission As String
Dim glnamestamp As String
Dim glidnostamp As String
Dim glpaynostamp As String
Dim glmemnostamp As String
Dim glnameteller As String
Dim glcombal As Currency
Dim gltellerbal As Currency
Dim glstampbal As Currency
Dim glcbocharge1accno As String
Dim glcbocharge1idno As String
Dim glcbocharge1memberno As String
Dim glcbocharge1payno As String
Dim glcbocharge1boobal As Currency
Dim glcbocharge1name As String
Dim glcbocharge2accno As String
Dim glcbocharge2idno As String
Dim glcbocharge2memberno As String
Dim glcbocharge2payno As String
Dim glcbocharge2boobal As Currency
Dim glcbocharge2name As String
Dim glcbocharge3accno As String
Dim glcbocharge3idno As String
Dim glcbocharge3memberno As String
Dim glcbocharge3payno As String
Dim glcbocharge3boobal As Currency
Dim glcbocharge3name As String
Dim glcbocharge4accno As String
Dim glcbocharge4idno As String
Dim glcbocharge4memberno As String
Dim glcbocharge4payno As String
Dim glcbocharge4boobal As Currency
Dim glcbocharge4name As String
Dim loan
Public maxRec As Long
Public myLevel As Integer

Private Type accoInfo
    ACCNO As String
    custName As String
    custBal As Currency
    AccName As String
    custno As String
    pic As String
    sign As String
End Type
Private Type faInfo
    minCall As Currency
    minFixed As Currency
End Type
Private Type shareinfo
    memberno As String
    totalshares As Currency
    
End Type
Private Type loansinfo
    MemNo As String
    Loanno As String
    LoanAmount As Currency
    repayperiod As Integer
End Type
Private Type sinfo
   meberno As String
   transdate As Date
   totalshares As Currency
End Type
Private Type tellerInfo
    tellerName As String
    tellerCubicle  As String
    tellerCurrBal As Currency
    tellerMaxBal As Currency
    tellerReplenish As Currency
    amtPayManager As Currency
    amtPaySuper As Currency
    amtPayTeller As Currency
    ttype As Byte '0 for not known, 1 for manager, 2 for Super, 3 for Teller
End Type



Private Type saInfo
    withLmt As Currency
    withInt As Integer
    FOSATarriffGuide As Currency
    lessThanWithIntCharge As Currency
    minBal As Currency
    lessThanMinBalCharge As Currency
    withCharge As Currency
    bankerscheque As Currency
    group3 As Currency
    individual3 As Currency
    over30 As String
    amoutover As Currency
    mobile As Currency
    intonauthorisedod As Integer
    intonunauthorisedod As Integer
    intonclearedchqs As Integer
    intonloans As Integer
    staffcode As String
    STAMPDUTY As Currency
End Type

Private Type transInfo
    accType As Byte   'the Type of account
                            '            1 is fixed call account,
                            '            2 is fixed term account,
                            '            3 is normal account,
    transAmt As Currency
    month As Byte
    desc As String
    tdate As Date
    openFee As Currency
    ACCNO As String
    idno As String
    PAYNO As String
    custno As String
    availbal As Currency
    AccName As String
    dateValid As Boolean
    fTrans As Boolean
    lastWithDate As Date
    vno As String
End Type
Private Type loanbalinfo
 loan As String
 memberno As String
 balance As String
 repayrate As Integer ' principal
 lastdate As Date
 interest As Double '% percenatage
 repaymethod As String ' either stl,rbal ,amrt
 repayperiod   As Integer 'period in months
End Type
Private loanbal As loanbalinfo
Private Loans As loansinfo
Private tos As sinfo
Private share As shareinfo
Private sa As saInfo
Private fa As faInfo

Private tInfo As transInfo
Private accData() As accoInfo

Private transtype As Byte '1 for deposit, withdrawal for 2,3 for loan repayment ,4 for share contribution

Private Sub cbotgNo_Click()
    cbotgNo_Change
End Sub

Private Sub cbotgNo_Change()
'On Error Resume Next
'Dim myrec As Object
'Dim rss As Object
'Dim amt As Long
'Set Myclass = New cdbase
'Set cn = CreateObject("adodb.connection")
'Provider = Myclass.OpenCon
'cn.Open Provider
'Set Myclass = New cdbase
'    Set cn = CreateObject("adodb.connection")
'    Provider = Myclass.OpenCon
'    Set rs = CreateObject("adodb.recordset")
'   cn.Open Provider, "atm","atm"
'    sql = "select distinct customerno,accno from customerbalance"
'    rs.Open sql, cn
'
'
'
'    Do
'    If Not rs.EOF Then
'        If Not IsNull(rs!customerno) Then cbotgno.AddItem rs!customerno & ""
'
'        rs.movenext
'        Else
'
'        Exit Sub
'        End If
'
'    Loop Until rs.EOF = True
'
'    rs.Close
' Set rs = CreateObject("adodb.recordset")
'    rs.Open "SELECT   *  FROM CustomerBalance ORDER BY CustomerBalanceid DESC", cn
'    rs.Close
'      cbotgno = cbotgno.list(0)
'Set rss = CreateObject("adodb.recordset")
'sql = "select minimumbalance from savingsaccountsparameters "
'rss.Open sql, cn
''get the name
''++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'Set myrec = CreateObject("adodb.recordset")
'            sql = "SELECT "
'            sql = sql & " CustomerAccount.AccountNumber as acc,"
'            sql = sql & " Customers.CustomerNo,"
'            sql = sql & " CustomerBalance.AvailableBalance as bal,"
'            sql = sql & " CustomerBalance.AccName as an,"
'            sql = sql & " Customers.Surname + ',  ' + Customers.OtherNames AS name"
'            sql = sql & " FROM Customers LEFT OUTER JOIN CustomerAccount"
'            sql = sql & " LEFT OUTER JOIN CustomerBalance"
'            sql = sql & " ON cast(CustomerAccount.CustomerNo as varchar(50))=CustomerBalance.CustomerNo"
'            sql = sql & " ON cast(Customers.CustomerNo as varchar(50))=CustomerAccount.CustomerNo   where CustomerAccount.CustomerNo='" & cbotgno & "'"
'            sql = sql & " ORDER BY CustomerBalance.CustomerBalanceid DESC"
'     myrec.Open sql, cn
'     If myrec.EOF Then
'     lblName = ""
'     lblaccname = ""
'     lblaccno = ""
'     lblavail = ""
'     lvememtrans.Visible = False
'     'MsgBox "Check if Member  Exist OR Check if the account is valid?? ", vbInformation, "Transactional details"
'     Exit Sub
'     Else
'      lvememtrans.Visible = True
'     lblName = myrec!Name
'     lblaccname = myrec!an
'     lblaccno = myrec!acc
'     lblavail = myrec!bal
'     End If
'
''++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'Dim actual As Currency
'If Not IsNull(rsd!amnt) Then
'amt = rsd!amnt
'Else:
'amt = 0
'End If
'    Dim lv As ListItem
'
'    lvememtrans.ListItems.Clear
'    rs.Open
'
'    rs.Filter = "customerno like '" & cbotgno.Text & "'"
'
'
'
'    If Not rsd.EOF Then
'    actual = CLng(rs!AvailableBalance) + amt
'
'    Else
'     actual = CLng(rs!AvailableBalance)
'
'
'     End If
'    Do While Not rs.EOF
'
'    With lvememtrans
'            If rs!transDate <> "" Then
'                    Set lv = .ListItems.add(, , rs!transDate)
'
'                    If rs!Transdescription <> "" Then lv.ListSubItems.add 1, , rs!Transdescription
'                    If rs!amount <> "" Then lv.ListSubItems.add 2, , rs!amount Else rs!amount = 0
'
'                    lv.ListSubItems.add 3, , rs!AvailableBalance
'                    lv.ListSubItems.add 4, , rs!commission
'                    lv.ListSubItems.item(3).Bold = True
'
'            End If
'    End With
'
'
'    rs.movenext
'    Loop
'
'    rs.Filter = 0
'    rs.Close
'
    
End Sub



'
'
Private Sub cbotgNo_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 39
KeyAscii = 0
End Select
If ValidChar(KeyAscii) = False Then KeyAscii = 0

End Sub

Private Sub acccr_Change()
    Set myclass = New cdbase
    Set cn = CreateObject("adodb.connection")
    Provider = myclass.OpenCon
    Set rs = CreateObject("ADODB.RECORDSET")
   cn.Open Provider, "atm", "atm"
    sql = ""
    sql = "Select * From CUB where AccNo='" & acccr & "'"
    rs.Open sql, cn
    If rs.EOF Then
        AVAIL1 = 0
        idno1 = "** not in file **"
        payno1 = "** not in file"
        name1 = "** not in file"
        accname1 = "** not in file"
    Else
        AVAIL1 = Format(IIf(IsNull(rs!availablebalance), 0, rs!availablebalance), Cfmt)
        If Not IsNull(rs!PAYNO) Then payno1 = rs!PAYNO
        If Not IsNull(rs!idno) Then idno1 = rs!idno
        If Not IsNull(rs!name) Then name1 = rs!name
        If Not IsNull(rs!name) Then accname1 = rs!name
        Lblname2 = accname1
        lblamount2 = Format(AVAIL1, Cfmt)
    End If
End Sub

Private Sub acccr_Validate(Cancel As Boolean)
Dim myrec1 As Object
Dim rss As Object
Dim amt As Currency
Set myclass = New cdbase
Set cn = CreateObject("adodb.connection")
Provider = myclass.OpenCon
cn.Open Provider, "atm", "atm"
Dim MYRE As Recordset
    Set MYRE = CreateObject("adodb.recordset")
    sql = "SELECT top 1 * from cub where accno='" & acccr & "' "
     MYRE.Open sql, cn
     If acccr <> "" Then
     If MYRE.EOF Then
      MsgBox "The account does not exist Please Seek assistance from the customer services", vbInformation, "Transactions"
     Exit Sub
     End If
     End If
End Sub

Private Sub accdr_Change()
    Set myclass = New cdbase
    Set cn = CreateObject("adodb.connection")
    Provider = myclass.OpenCon
    Set rs = CreateObject("adodb.recordset")
   cn.Open Provider, "atm", "atm"
    sql = ""
    sql = "select * from cub where accno='" & accdr & "'"
    rs.Open sql, cn
    If rs.EOF Then
         AVAIL1 = 0
         idno1 = "** not in file **"
         payno1 = "** not in file"
         name1 = "** not in file"
         accname1 = "** not in file"
    Else
        AVAIL2 = Format(IIf(IsNull(rs!availablebalance), 0, rs!availablebalance), Cfmt)
        payno2 = IIf(IsNull(rs!PAYNO), "", rs!PAYNO)
        idno2 = IIf(IsNull(rs!idno), "", rs!idno)
        name2 = IIf(IsNull(rs!name), "", rs!name)
        accname2 = IIf(IsNull(rs!name), "", rs!name)
        lblname1 = accname2
        lblamount1 = Format(AVAIL2, Cfmt)
    End If
End Sub

Private Sub accdr_Validate(Cancel As Boolean)
Dim myrec1 As Object
Dim rss As Object
Dim amt As Currency
Set myclass = New cdbase
Set cn = CreateObject("adodb.connection")
Provider = myclass.OpenCon
cn.Open Provider, "atm", "atm"
Dim MYRE As Recordset
    Set MYRE = CreateObject("adodb.recordset")
    sql = "SELECT top 1 * from cub where accno='" & accdr & "' "
     MYRE.Open sql, cn
     If accdr <> "" Then
     If MYRE.EOF Then
      MsgBox "The account does not exist Please Seek assistance from the customer services", vbInformation, "Transactions"
     Exit Sub
     End If
     End If
End Sub

Private Sub cbodetail_Change()
    cbodetail_Change
End Sub

Private Sub Cbodetail_Click()
    On Error Resume Next
    Select Case Cbodetail
        Case "Account Number"
        Txtaccno.Enabled = True
        Txtaccno.Visible = True
        Txtaccno.SetFocus
        txtidno.Enabled = False
        txtidno.Visible = False
        TXTTGNO.Enabled = False
        txtPayNo.Enabled = False
        txtPayNo.Visible = False
        txtidno = ""
        txtPayNo = ""
        TXTTGNO = ""
        Txtaccno.SetFocus
        Case "IDNo"
        txtidno.Enabled = True
        txtidno.Visible = True
        txtidno.SetFocus
        TXTTGNO.Enabled = False
        TXTTGNO.Visible = False
        Txtaccno.Enabled = False
        Txtaccno.Visible = False
        txtPayNo.Enabled = False
        txtPayNo.Visible = False
        txtPayNo = ""
        Txtaccno = ""
        TXTTGNO = ""
        Case "MemberNo"
        TXTTGNO.Enabled = True
        TXTTGNO.Visible = True
        TXTTGNO.SetFocus
        Txtaccno.Enabled = False
        Txtaccno.Visible = False
        txtidno.Enabled = False
        txtidno.Visible = False
        txtPayNo.Enabled = False
        txtPayNo.Visible = False
        txtidno = ""
        txtPayNo = ""
        Txtaccno = ""
        Case "PayrollNo"
        txtPayNo.Enabled = True
        txtPayNo.Visible = True
        txtPayNo.SetFocus
        Txtaccno.Enabled = False
        Txtaccno.Visible = False
        txtidno.Enabled = False
        txtidno.Visible = False
        txtidno = ""
        txtPayNo = ""
        Txtaccno = ""
    End Select
End Sub
Private Sub update_commissionacc()
    sql = ""
    sql = "set dateformat dmy insert into customerbalance (customerno,accname,amount,availablebalance,accno,transdescription,transdate,commission,period,transType,posted,locked,status,vno,auditid,auditdate,moduleid,accd) "
    sql = sql & " values ('" & glcommemno & "','" & glnamecom1 & "'," & withcharges & "," & bookba + withcharges & ",'" & LBLGLCOMMISSION & "','Comm/charges','" & DTP & "',0,'" & month(Date) & "','CR',0,0,0,'C/W Com','" & User & "','" & Now & "','3','" & Txtaccno & "' )"
    myclass.save sql
    
    sql = ""
    sql = "set dateformat dmy update cub set amount=" & totalcharges + withcharges & ",Active=1,transdescription='" & desc & "',availablebalance=" & bookba + withcharges + totalcharges & ",transdate='" & Date & "',vno='" & txtvoucherno & "',period='" & month(Date) & "',auditid='" & teller & "',auditdate='" & Now & "',moduleid=2,active=1 where accno='" & LBLGLCOMMISSION & "'"
    myclass.save sql
End Sub
Private Sub update_stampacc()
    sql = ""
    sql = "set dateformat dmy insert into customerbalance (customerno,accname,amount,availablebalance,accno,transdescription,transdate,commission,period,transType,posted,locked,status,vno,auditid,auditdate,moduleid,accd) "
    sql = sql & " values ('" & glmemnostamp & "','" & glnamestamp1 & "'," & txtstamp & "," & bookba2 + txtstamp & ",'" & lblglstamp & "','S/Duty','" & DTP & "',0,'" & month(Date) & "','CR',0,0,0,'" & txtvoucherno & "','" & User & "','" & Now & "','3','" & Txtaccno & "' )"
    myclass.save sql
    
    sql = ""
    sql = "set dateformat dmy update cub set amount=" & txtstamp & ",Active=1,transdescription='" & desc & "',availablebalance=" & bookba2 + txtstamp & ",transdate='" & Date & "',vno='" & txtvoucherno & "',period='" & month(Date) & "',auditid='" & teller & "',auditdate='" & Now & "',moduleid=2,active=1 where accno='" & lblglstamp & "'"
    myclass.save sql
End Sub

Private Sub getulmt()
Dim r As Object
Dim nav As Integer
setRsr r, "select  top 1 paymentno  from repay where loanno='" & loan & "' order by paymentno desc "
If r.EOF Then Exit Sub
 If Not IsNull(r!paymentno) Then ulmt = r!paymentno
 ulmt = ulmt + 1
r.Close
Set r = Nothing
End Sub
Private Sub getloaninfo()

On Error GoTo ErrorHandler
Dim rd As Object

Dim nav As Integer
number = 0
setRsr rd, "select * from loans where memberno= '" & txtMemberNo & "' order by loanno asc"
If rd.EOF Then
 getloanbal
Exit Sub
Else
GoTo UkoHaha
End If




UkoHaha:

           While Not rd.EOF
              getloanbal
                  number = number + 1
                       If Not IsNull(rd!Loanno) Then Loans.Loanno = rd!Loanno
                       If Not IsNull(rd!memberno) Then Loans.MemNo = rd!memberno
                       If Not IsNull(rd!loanAmt) Then Loans.LoanAmount = rd!loanAmt
                            
                            principal = (Loans.LoanAmount / loanbal.repayperiod)
                            
                            interest1 = (Loans.LoanAmount * (15 / 100) * (1 / 12))
                            getulmt
                            expectedamount1 = principal + interest1
                            saveloans
                           remi = tInfo.transAmt - principal - interest1
                 rd.MoveNext
            If Not rd.EOF Then GoTo UkoHaha
            
          Wend
           
rd.Close
Set rd = Nothing
Exit Sub
ErrorHandler:
MsgBox err.description
End Sub
Private Sub getloanbal()
Dim p As Integer
Dim r As Object
setRsr r, "select  loanno,memberno,balance,repayrate,lastdate,interest,repaymethod,repayperiod from loanbal where memberno='" & txtMemberNo & "' and balance>1"
If r.EOF Then Exit Sub
With loanbal
  
        If Not IsNull(r!Loanno) Then .loan = r!Loanno
        If Not IsNull(r!Loanno) Then loan = r!Loanno
        If Not IsNull(r!memberno) Then .memberno = r!memberno
        If Not IsNull(r!balance) Then .balance = r!balance
        If Not IsNull(r!repayrate) Then .repayrate = r!repayrate
        If Not IsNull(r!lastdate) Then .lastdate = r!lastdate
        If Not IsNull(r!interest) Then .interest = r!interest
        If Not IsNull(r!repaymethod) Then .repaymethod = r!repaymethod
        If Not IsNull(r!repayperiod) Then .repayperiod = r!repayperiod
        '// get the interest calculation
        'p = month(.lastdate)
        'If month(Date) = p Then
       ' Else
       ' End If
        'interest1 = 1 / 12 * .balance * 15 / 100
        If cbomemtrans = "Loan Repayment" Then
        getulmt
        saveloans
        End If
End With
r.Close
Set r = Nothing

End Sub
Private Sub setRsr(thisRs As Object, strSQL As String)
Set myclass = New cdbase
Dim S As String
S = SelectedDsn
    Set thisRs = CreateObject("adodb.recordset")
    Set cn = CreateObject("adodb.connection")
    cn.Open S
    thisRs.Open strSQL, cn
End Sub
Private Sub setRs(thisRs As Object, strSQL As String)
Set myclass = New cdbase
    Set thisRs = CreateObject("adodb.recordset")
    Set cn = CreateObject("adodb.connection")
    Provider = myclass.OpenCon
   cn.Open Provider, "atm", "atm"
    thisRs.Open strSQL, cn
End Sub
Private Sub saveloans()
Dim cnnn As ADODB.Connection
Set myclass = New cdbase
Dim St As String
St = SelectedDsn
Set cnnn = New ADODB.Connection
cnnn.Open St

If number <= 1 And tInfo.transAmt < expectedamount1 And interest1 < tInfo.transAmt Then
      'interest1 = tInfo.transAmt
      
      principal = tInfo.transAmt - interest1
        'principal = tInfo.transAmt - expectedamount1
        ' If principal < 1 Then principal = 0
    sql = ""
    sql = "INSERT INTO REPAY"
    sql = sql & "(LoanNo, DateReceived, PaymentNo, Amount, Principal, "
    sql = sql & "Interest, IntrCharged, IntrOwed, LoanBalance, Locked, Posted, Accrued, Remarks, AuditID,"
    sql = sql & " AuditTime)"
    sql = sql & " VALUES('" & Loans.Loanno & "', '" & Date & "', '" & ulmt & "', '" & interest1 + principal & "', "
    sql = sql & "'" & principal & "', '" & interest1 & "', '" & interest1 & "', '" & interest1 - interest1 & "', "
    sql = sql & " '" & loanbal.balance - principal & "', 'No', 'No', 'No', 'Received from teller', 'Teller', '" & Now & "')"
    cn.Execute sql
    sql = ""
  sql = " update  loanbal set  balance= '" & loanbal.balance - principal & "'where loanno='" & Loans.Loanno & "'"
  cn.Execute sql
    expectedamount1 = expectedamount1
    
ElseIf number <= 1 And tInfo.transAmt < expectedamount1 And interest1 > tInfo.transAmt Then
    interest1 = tInfo.transAmt
      
      
        principal = tInfo.transAmt - expectedamount1
         If principal < 1 Then principal = 0

    sql = ""
    sql = "INSERT INTO REPAY"
    sql = sql & "(LoanNo, DateReceived, PaymentNo, Amount, Principal, "
    sql = sql & "Interest, IntrCharged, IntrOwed, LoanBalance, Locked, Posted, Accrued, Remarks, AuditID,"
    sql = sql & " AuditTime)"
    sql = sql & " VALUES('" & Loans.Loanno & "', '" & Date & "', '" & ulmt & "', '" & interest1 + principal & "', "
    sql = sql & "'" & principal & "', '" & interest1 & "', '" & interest1 & "', '" & interest1 - interest1 & "', "
    sql = sql & " '" & loanbal.balance - principal & "', 'No', 'No', 'No', 'Received from teller', 'Teller', '" & Now & "')"
    cn.Execute sql
    sql = ""
  sql = " update  loanbal set  balance= '" & loanbal.balance - principal & "'where loanno='" & Loans.Loanno & "'"
  cn.Execute sql
  expectedamount1 = expectedamount1
    
ElseIf number <= 1 And tInfo.transAmt >= expectedamount1 Then

    principal = txtadjamnt
    If principal > 0 Then
    sql = ""
    sql = "SET DATEFORMAT DMY INSERT INTO REPAY"
    sql = sql & "(LoanNo, DateReceived, PaymentNo, Amount, Principal, "
    sql = sql & "Interest, IntrCharged, IntrOwed, LoanBalance, Locked, Posted, Accrued, Remarks, AuditID,"
    sql = sql & " AuditTime)"
    sql = sql & " VALUES('" & loan & "', '" & Date & "', " & ulmt & ", " & txtadjamnt & ", "
    sql = sql & "" & txtadjamnt & ",0, 0, 0, "
    sql = sql & " " & loanbal.balance - txtadjamnt & ", 'No', 'No', 'No', 'Received from teller', '" & User & "', '" & Now & "')"
   cnnn.Execute sql
    
    sql = ""
   sql = "set dateformat dmy  update  loanbal set  balance= " & loanbal.balance - principal & " ,lastdate='" & Date & "' where loanno='" & loan & "'"
  cn.Execute sql
  ' End If
   expectedamount1 = expectedamount1
   Else
   sql = ""
    sql = "SET DATEFORMAT DMY INSERT INTO REPAY"
    sql = sql & "(LoanNo, DateReceived, PaymentNo, Amount, Principal, "
    sql = sql & "Interest, IntrCharged, IntrOwed, LoanBalance, Locked, Posted, Accrued, Remarks, AuditID,"
    sql = sql & " AuditTime)"
    sql = sql & " VALUES('" & Loans.Loanno & "', '" & Date & "', " & ulmt & ", " & txtadjamnt & ", "
    sql = sql & " " & txtadjamnt & " , 0, 0, 0, "
    sql = sql & " " & loanbal.balance & ", 'No', 'No', 'No', 'Received from teller', '" & User & "', '" & Now & "')"
   cn.Execute sql
   
   sql = ""
   sql = "set dateformat dmy  update  loanbal set  balance= " & loanbal.balance & " ,lastdate='" & Date & "' where loanno='" & loan & "'"
  cn.Execute sql
   End If
   
ElseIf number = 2 Then
sql = ""
sql = "INSERT INTO REPAY"
sql = sql & "(LoanNo, DateReceived, PaymentNo, Amount, Principal, "
sql = sql & "Interest, IntrCharged, IntrOwed, LoanBalance, Locked, Posted, Accrued, Remarks, AuditID,"
sql = sql & " AuditTime)"
sql = sql & " VALUES('" & Loans.Loanno & "', '" & Date & "', '" & ulmt & "', '" & interest1 + principal & "', "
sql = sql & "'" & principal & "', '" & interest1 & "', '" & interest1 & "', '" & interest1 - interest1 & "', "
sql = sql & " '" & Loans.LoanAmount - principal & "', 'No', 'No', 'No', 'Received from teller', 'Teller', '" & Now & "')"
cn.Execute sql
ElseIf number = 3 Then
sql = ""
sql = "INSERT INTO REPAY"
sql = sql & "(LoanNo, DateReceived, PaymentNo, Amount, Principal, "
sql = sql & "Interest, IntrCharged, IntrOwed, LoanBalance, Locked, Posted, Accrued, Remarks, AuditID,"
sql = sql & " AuditTime)"
sql = sql & " VALUES('" & Loans.Loanno & "', '" & Date & "', '" & ulmt & "', '" & interest1 + principal & "', "
sql = sql & "'" & principal & "', '" & interest1 & "', '" & interest1 & "', '" & interest1 - interest1 & "', "
sql = sql & " '" & Loans.LoanAmount - principal & "', 'No', 'No', 'No', 'Received from teller', 'Teller', '" & Now & "')"
cn.Execute sql
ElseIf number = 4 Then
 sql = ""
sql = "INSERT INTO REPAY"
sql = sql & "(LoanNo, DateReceived, PaymentNo, Amount, Principal, "
sql = sql & "Interest, IntrCharged, IntrOwed, LoanBalance, Locked, Posted, Accrued, Remarks, AuditID,"
sql = sql & " AuditTime)"
sql = sql & " VALUES('" & Loans.Loanno & "', '" & Date & "', '" & ulmt & "', '" & interest1 + principal & "', "
sql = sql & "'" & principal & "', '" & interest1 & "', '" & interest1 & "', '" & interest1 - interest1 & "', "
sql = sql & " '" & Loans.LoanAmount - principal & "', 'No', 'No', 'No', 'Received from teller', 'Teller', '" & Now & "')"
cn.Execute sql
ElseIf number = 5 Then
sql = ""
sql = "INSERT INTO REPAY"
sql = sql & "(LoanNo, DateReceived, PaymentNo, Amount, Principal, "
sql = sql & "Interest, IntrCharged, IntrOwed, LoanBalance, Locked, Posted, Accrued, Remarks, AuditID,"
sql = sql & " AuditTime)"
sql = sql & " VALUES('" & Loans.Loanno & "', '" & Date & "', '" & ulmt & "', '" & interest1 + principal & "', "
sql = sql & "'" & principal & "', '" & interest1 & "', '" & interest1 & "', '" & interest1 - interest1 & "', "
sql = sql & " '" & Loans.LoanAmount - principal & "', 'No', 'No', 'No', 'Received from teller', 'Teller', '" & Now & "')"
cn.Execute sql
End If

'    lblBalance = ""
'        txtamount = ""
'        lblCustomerNumber = ""
'        lblname = ""
'         Lblidno = ""
'          lblmemno = ""
'          Lblpayno = ""
'          lblaccno = ""
'          txtvou = ""
          

End Sub

Private Sub chkbackofficetransactions_Click()
If chkbackofficetransactions = vbChecked Then
lblMemberNo.Visible = True
txtMemberNo.Visible = True
Else
lblMemberNo.Visible = False
txtMemberNo.Visible = False
End If
End Sub

Private Sub cmdadj_Click()
    Dim avail As Currency, id
    Dim ACCNO As String
    Dim cubrs As Recordset
    Dim adjamnt As Currency
    Dim co As Currency
    Dim rcom As Object
    Dim rss As Object
    Dim DESCR As String
    Dim cub As String
    teller = teller
    cub = GetSetting("FOSAdll", "Teller", "Cubie Number")
    Set myclass = New cdbase
    Set cn = CreateObject("adodb.connection")
    Provider = myclass.OpenCon
    Set rs = CreateObject("adodb.recordset")
   cn.Open Provider, "atm", "atm"
    'validation
    If LBLGLCONTRA = "" Then
        MsgBox "Enter the account to for the contra", vbInformation, "Member Transactions"
        Exit Sub
    End If
    If txtadjamnt <= 0 Then
        MsgBox "You are trying to post zero shillings? Please Enter the amount", vbInformation
        Exit Sub
    End If
    'total_charges
    desc = cbomemtrans
    If Not IsNumeric(txtadjamnt) Then
        MsgBox "Please Enter Values", vbInformation, "Adjustment transactions"
        Exit Sub
    End If
    adjamnt = txtadjamnt
    If optcredit.value = True Then
        If Not Save_To_GL(LBLGLCONTRA, Txtaccno, txtadjamnt, txtvoucherno, txtvoucherno, _
        DTP, txtvoucherno, cbomemtrans, ErrorMessage) Then
            If ErrorMessage <> "" Then
                MsgBox ErrorMessage, vbInformation, Me.Caption
                ErrorMessage = ""
            End If
        End If
    ElseIf Optdebit.value = True Then
        If Not Save_To_GL(Txtaccno, LBLGLCONTRA, txtadjamnt, txtvoucherno, txtvoucherno, _
        DTP, txtvoucherno, cbomemtrans, ErrorMessage) Then
            If ErrorMessage <> "" Then
                MsgBox ErrorMessage, vbInformation, Me.Caption
                ErrorMessage = ""
            End If
        End If
    End If
    MsgBox "Transaction Completed Successfully", vbInformation, Me.Caption
    Txtaccno = ""
    LBLGLCONTRA = ""
    txtadjamnt = "0.00"
    lblavail = "0.00"
    lbluncleared = "0.00"
    lblbookbalance = "0.00"
    Exit Sub
    'teller = GetSetting("FOSAdll", "Teller", "Name")
    Select Case Cbodetail
        Case "Account Number"
        Set rss = CreateObject("adodb.recordset")
        rss.Open "SELECT * FROM CustomerBalance where accno='" & Txtaccno & "' " _
        & "ORDER BY CustomerBalanceid DESC", cn
        If rss.EOF Then
            If optcredit = True Then
                Set rss = CreateObject("adodb.recordset")
                ' RSS.Open ""
                sql = "SELECT   *  FROM Cub where accno='" & Txtaccno & "'"
                Set cubrs = New ADODB.Recordset
                cubrs.Open sql, cn
                If Not cubrs.EOF Then
                    custno = cubrs.Fields("memberno")
                    AccName = cubrs.Fields("accountname")
                    ' custno = cubrs.Fields("")
                Else
                    custno = "NA"
                    AccName = ""
                End If
                sql = ""
                sql = "set dateformat dmy insert into customerbalance (customerno,accname,amount,availablebalance,accno,transdescription,transdate,commission,period,transType,posted,locked,status,vno,auditid,auditdate,moduleid,accd) "
                sql = sql & " values ('" & custno & "','" & AccName & "'," & adjamnt & "," & avail + adjamnt & ",'" & Txtaccno & "','" & desc & "','" & DTP & "',0,'" & month(Date) & "','CR',0,0,0,'" & txtvoucherno & "','" & User & "','" & Now & "','3','" & glaccno & "' )"
                cn.Execute sql
                sql = ""
                sql = "set dateformat dmy update cub set amount=" & adjamnt & ",Active=1,transdescription='" & desc & "',availablebalance=" & avail + adjamnt & ",transdate='" & Date & "',vno='" & txtvoucherno & "',period='" & month(Date) & "',auditid='" & teller & "',auditdate='" & Now & "',moduleid=2,active=1 where accno='" & Txtaccno & "'"
                cn.Execute sql
                DESCR = "From - " & AccName & ""
                If glnamE1 <> "" Then
                    sql = ""
                    sql = "set dateformat dmy insert into customerbalance (customerno,accname,amount,availablebalance,accno,transdescription,transdate,commission,period,transType,posted,locked,status,vno,auditid,auditdate,moduleid,accd) "
                    sql = sql & " values ('" & glmemno & "','" & glnamE1 & "'," & adjamnt & "," & bookba + adjamnt & ",'" & glaccno & "','" & DESCR & "','" & DTP & "',0,'" & month(Date) & "','DR',0,0,0,'" & txtvoucherno & "','" & User & "','" & Now & "','3','" & Txtaccno & "' )"
                    cn.Execute sql
    
                    sql = ""
                    sql = "set dateformat dmy update cub set amount=" & adjamnt & ",Active=1,transdescription='" & DESCR & "',availablebalance=" & bookba + adjamnt & ",transdate='" & Date & "',vno='" & txtvoucherno & "',period='" & month(Date) & "',auditid='" & teller & "',auditdate='" & Now & "',moduleid=2,active=1 where accno='" & glaccno & "'"
                    cn.Execute sql
                End If
                '//let also affect the teller for transfer to kericho and other branches
                sql = "set dateformat dmy insert into [teller transactions] (tellername,Cubiclenumber,Deposits,Withdrawals,Accountnumber,tgno,TransactionDate,vno,Posted,Locked,Auditid,audittime,transdescription,balance,name,accname,commission,printed,cash) "
                sql = sql & "values('" & User & "','" & cub & "'," & adjamnt & ",0,'" & Txtaccno & "','" & custno & "','" & Date & "','" & txtvoucherno & "',0,0,'" & User & "','" & Now & "','" & desc & "',0,'" & AccName & "','" & AccName & "',0,0,1)"
                cn.Execute sql
                
                sql = "set dateformat dmy INSERT INTO CashTransactions"
                sql = sql & "   (Customerno, AccNo, AccName, Amount, Transdescription, Transdate, Commission, Chequeno, Transtype, vno, Posted, Locked, Auditid, AuditTime,userName)"
                sql = sql & " VALUES ('" & custno & "', '" & Txtaccno & "', '" & AccName & "', " & adjamnt & ", 'adj', '" & Date & "', 0, 'NON', 'CR', '" & txtvoucherno & "', 0, 0, '" & User & "', '" & Now & "', '" & User & "')"
                cn.Execute sql
                
            ElseIf Optdebit = True Then
                Set rss = CreateObject("adodb.recordset")
                'RSS.Open ""
                sql = "SELECT   *  FROM Cub where accno='" & Txtaccno & "'"
                Set cubrs = New ADODB.Recordset
                cubrs.Open sql, cn
                If Not cubrs.EOF Then
                    custno = cubrs.Fields("memberno")
                    AccName = cubrs.Fields("accountname")
                Else
                    custno = "NA"
                    AccName = ""
                End If
                avail = cubrs.Fields("availablebalance")
                sql = ""
                sql = ""
                sql = "set dateformat dmy insert into customerbalance (customerno,accname,amount,availablebalance,accno,transdescription,transdate,commission,period,transType,posted,locked,status,vno,auditid,auditdate,moduleid) "
                sql = sql & " values ('" & custno & "','" & AccName & "'," & adjamnt & "," & avail - adjamnt & ",'" & Txtaccno & "','" & desc & "','" & DTP & "',0,'" & month(Date) & "','DR',0,0,0,'" & txtvoucherno & "','" & User & "','" & Now & "','3' )"
                cn.Execute sql
                sql = ""
                sql = "set dateformat dmy update cub set amount=" & adjamnt & ",Active=1,transdescription='" & desc & "',availablebalance=" & avail - adjamnt & ",transdate='" & Date & "',vno='" & txtvoucherno & "',period='" & month(Date) & "',auditid='" & teller & "',auditdate='" & Now & "',moduleid=2,active=1 where accno='" & ACCNO & "'"
                cn.Execute sql
                DESCR = "To - " & Txtaccno & ""
                If glnamE1 <> "" Then
                    sql = ""
                    sql = "set dateformat dmy insert into customerbalance (customerno,accname,amount,availablebalance,accno,transdescription,transdate,commission,period,transType,posted,locked,status,vno,auditid,auditdate,moduleid,accd) "
                    sql = sql & " values ('" & glmemno & "','" & glnamE & "'," & adjamnt & "," & bookba - adjamnt & ",'" & glaccno & "','" & desc & "','" & DTP & "',0,'" & month(Date) & "','CR',0,0,0,'" & txtvoucherno & "','" & User & "','" & Now & "','3','" & Txtaccno & "' )"
                    cn.Execute sql
                    
                    sql = ""
                    sql = "set dateformat dmy update cub set amount=" & adjamnt & ",Active=1,transdescription='" & desc & "',availablebalance=" & bookba + adjamnt & ",transdate='" & Date & "',vno='" & txtvoucherno & "',period='" & month(Date) & "',auditid='" & teller & "',auditdate='" & Now & "',moduleid=2,active=1 where accno='" & glaccno & "'"
                    cn.Execute sql
                End If
                sql = "set dateformat dmy insert into [teller transactions] (tellername,Cubiclenumber,Deposits,Withdrawals,Accountnumber,tgno,TransactionDate,vno,Posted,Locked,Auditid,audittime,transdescription,balance,name,accname,commission,printed,cash) "
                sql = sql & "values('" & User & "','" & cub & "',0," & adjamnt & ",'" & Txtaccno & "','" & custno & "','" & Date & "','" & txtvoucherno & "',0,0,'" & User & "','" & Now & "','" & desc & "',0,'" & AccName & "','" & AccName & "',0,0,1)"
                cn.Execute sql
                
                sql = "set dateformat dmy INSERT INTO CashTransactions"
                sql = sql & "   (Customerno, AccNo, AccName, Amount, Transdescription, Transdate, Commission, Chequeno, Transtype, vno, Posted, Locked, Auditid, AuditTime,userName)"
                sql = sql & " VALUES ('" & custno & "', '" & Txtaccno & "', '" & AccName & "', " & adjamnt & ", 'adj', '" & Date & "', 0, 'NON', 'DR', '" & txtvoucherno & "', 0, 0, '" & User & "', '" & Now & "', '" & User & "')"
                cn.Execute sql
            End If
        Else
            If Not IsNull(rss!ACCNO) Then ACCNO = rss!ACCNO
            If Not IsNull(rss!availablebalance) Then avail = rss!availablebalance
            If Not IsNull(rss!customerbalanceid) Then id = rss!customerbalanceid
            If Not IsNull(rss!AccName) Then AccName = rss!AccName
            If Not IsNull(rss!CustomerNo) Then custno = rss!CustomerNo
            If optcredit = True Then
                sql = ""
                sql = "set dateformat dmy insert into customerbalance (customerno,accname,amount,availablebalance,accno,transdescription,transdate,commission,period,transType,posted,locked,status,vno,auditid,auditdate,moduleid,accd) "
                sql = sql & " values ('" & custno & "','" & AccName & "'," & adjamnt & "," & avail + adjamnt & ",'" & Txtaccno & "','" & desc & "','" & DTP & "',0,'" & month(Date) & "','CR',0,0,0,'" & txtvoucherno & "','" & User & "','" & Now & "','3','" & glaccno & "' )"
                cn.Execute sql
                sql = "set dateformat dmy update cub set amount=" & adjamnt & ",Active=1,transdescription='" _
                & desc & "',availablebalance=" & avail + adjamnt & ",transdate='" & Date & "',vno='" & txtvoucherno _
                & "',period='" & month(Date) & "',auditid='" & teller & "',auditdate='" & Now & "',moduleid=2 " _
                & "where accno='" & ACCNO & "'"
                cn.Execute sql
                DESCR = "FROM - " & Txtaccno & ""
                If glnamE1 <> "" Then
                    sql = ""
                    sql = "set dateformat dmy insert into customerbalance (customerno,accname,amount,availablebalance,accno,transdescription,transdate,commission,period,transType,posted,locked,status,vno,auditid,auditdate,moduleid,accd) "
                    sql = sql & " values ('" & glmemno & "','" & glnamE & "'," & adjamnt & "," & bookba + adjamnt & ",'" & glaccno & "','" & DESCR & "','" & DTP & "',0,'" & month(Date) & "','DR',0,0,0,'" & txtvoucherno & "','" & User & "','" & Now & "','3','" & Txtaccno & "' )"
                    cn.Execute sql
                    
                    sql = ""
                    sql = "set dateformat dmy update cub set amount=" & adjamnt & ",Active=1,transdescription='" _
                    & desc & "',availablebalance=" & bookba + adjamnt & ",transdate='" & Date & "',vno='" & _
                    txtvoucherno & "',period='" & month(Date) & "',auditid='" & teller & "',auditdate='" & Now & _
                    "',moduleid=2 where accno='" & glaccno & "'"
                    cn.Execute sql
                End If
                '//let also affect the teller for transfer to kericho and other branches
                sql = "set dateformat dmy insert into [teller transactions] (tellername,Cubiclenumber,Deposits,Withdrawals,Accountnumber,tgno,TransactionDate,vno,Posted,Locked,Auditid,audittime,transdescription,balance,name,accname,commission,printed,cash) "
                sql = sql & "values('" & User & "','" & cub & "'," & adjamnt & ",0,'" & Txtaccno & "','" & custno & "','" & Date & "','" & txtvoucherno & "',0,0,'" & User & "','" & Now & "','" & desc & "',0,'" & AccName & "','" & AccName & "',0,0,1)"
                cn.Execute sql
                
                sql = "set dateformat dmy INSERT INTO CashTransactions"
                sql = sql & "   (Customerno, AccNo, AccName, Amount, Transdescription, Transdate, Commission, Chequeno, Transtype, vno, Posted, Locked, Auditid, AuditTime,userName)"
                sql = sql & " VALUES ('" & custno & "', '" & Txtaccno & "', '" & AccName & "', " & adjamnt & ", 'adj', '" & Date & "', 0, 'NON', 'CR', '" & txtvoucherno & "', 0, 0, '" & User & "', '" & Now & "', '" & User & "')"
                cn.Execute sql
            ElseIf Optdebit = True Then
                sql = ""
                If desc = "Cash Withdrawal" Then
                    sql = "set dateformat dmy insert into customerbalance (customerno,accname,amount,availablebalance,accno,transdescription,transdate,commission,period,transType,posted,locked,status,vno,auditid,auditdate,moduleid,accd) "
                    sql = sql & " values ('" & custno & "','" & AccName & "'," & adjamnt & "," & avail - adjamnt & ",'" & Txtaccno & "','" & desc & "','" & DTP & "',0,'" & month(Date) & "','DR',0,0,0,'" & txtvoucherno & "','" & User & "','" & Now & "','3','" & glcomm & "' )"
                    cn.Execute sql
                    
                    sql = "set dateformat dmy insert into customerbalance (customerno,accname,amount,availablebalance,accno,transdescription,transdate,commission,period,transType,posted,locked,status,vno,auditid,auditdate,moduleid,accd) "
                    sql = sql & " values ('" & custno & "','" & AccName & "'," & withcharges & "," & avail - adjamnt - withcharges & ",'" & Txtaccno & "','CW/Charges','" & DTP & "',0,'" & month(Date) & "','DR',0,0,0,'" & txtvoucherno & "','" & User & "','" & Now & "','3','" & glcomm & "' )"
                    cn.Execute sql
                If lblglstamp <> "" Then
                    sql = "set dateformat dmy insert into customerbalance (customerno,accname,amount,availablebalance,accno,transdescription,transdate,commission,period,transType,posted,locked,status,vno,auditid,auditdate,moduleid,accd) "
                    sql = sql & " values ('" & custno & "','" & AccName & "'," & txtstamp & "," & avail - adjamnt - withcharges - txtstamp & ",'" & Txtaccno & "','Stamp D/C','" & DTP & "',0,'" & month(Date) & "','DR',0,0,0,'" & txtvoucherno & "','" & User & "','" & Now & "','3','" & lblglstamp & "' )"
                    cn.Execute sql
                    update_stampacc
                End If
                If LBLGLCOMMISSION <> "" Then
                    update_commissionacc
                End If

                sql = "update cub set amount=" & adjamnt & ",Active=1,transdescription='" & desc & "',availablebalance=" & avail - adjamnt - withcharges - txtstamp & ",transdate='" & Date & "',vno='" & txtvoucherno & "',period='" & month(Date) & "',auditid='" & teller & "',auditdate='" & Now & "',moduleid=2,active=1 where accno='" & ACCNO & "'"
                myclass.save sql
                DESCR = "TO - " & Txtaccno & ""
                If glnamE1 <> "" Then
                    sql = ""
                    sql = "set dateformat dmy insert into customerbalance (customerno,accname,amount,availablebalance,accno,transdescription,transdate,commission,period,transType,posted,locked,status,vno,auditid,auditdate,moduleid,accd) "
                    sql = sql & " values ('" & glmemno & "','" & glnamE & "'," & adjamnt & "," & bookba - adjamnt & ",'" & glaccno & "','" & DESCR & "','" & DTP & "',0,'" & month(Date) & "','CR',0,0,0,'" & txtvoucherno & "','" & User & "','" & Now & "','3','" & Txtaccno & "' )"
                    myclass.save sql
                    
                    sql = ""
                    sql = "set dateformat dmy update cub set amount=" & adjamnt & ",Active=1,transdescription='" & DESCR & "',availablebalance=" & bookba - adjamnt & ",transdate='" & Date & "',vno='" & txtvoucherno & "',period='" & month(Date) & "',auditid='" & teller & "',auditdate='" & Now & "',moduleid=2,active=1 where accno='" & glaccno & "'"
                    myclass.save sql
                End If
            Else
                sql = "set dateformat dmy insert into customerbalance (customerno,accname,amount,availablebalance,accno,transdescription,transdate,commission,period,transType,posted,locked,status,vno,auditid,auditdate,moduleid,accd) "
                sql = sql & " values ('" & custno & "','" & AccName & "'," & adjamnt & "," & avail - adjamnt & ",'" & Txtaccno & "','" & desc & "','" & DTP & "',0,'" & month(Date) & "','DR',0,0,0,'" & txtvoucherno & "','" & User & "','" & Now & "','3','" & glcomm & "' )"
                myclass.save sql
                sql = "set dateformat dmy update cub set amount=" & adjamnt & ",Active=1,transdescription='" _
                & desc & "',availablebalance=" & avail - adjamnt - withcharges & ",transdate='" & Date & _
                "',vno='" & txtvoucherno & "',period='" & month(Date) & "',auditid='" & teller & "',auditdate='" _
                & Now & "',moduleid=2 where accno='" & ACCNO & "'"
                myclass.save sql
                If Optdebit = True Then
                    DESCR = "FROM - " & Txtaccno & ""
                    If glnamE1 <> "" Then
                        sql = ""
                        sql = "set dateformat dmy insert into customerbalance (customerno,accname,amount,availablebalance,accno,transdescription,transdate,commission,period,transType,posted,locked,status,vno,auditid,auditdate,moduleid,accd) "
                        sql = sql & " values ('" & glmemno & "','" & glnamE & "'," & adjamnt & "," & bookba - adjamnt & ",'" & glaccno & "','" & DESCR & "','" & DTP & "',0,'" & month(Date) & "','CR',0,0,0,'" & txtvoucherno & "','" & User & "','" & Now & "','3','" & Txtaccno & "' )"
                        myclass.save sql
                        
                        sql = ""
                        sql = "set dateformat dmy update cub set amount=" & adjamnt & ",Active=1,transdescription='" _
                        & DESCR & "',availablebalance=" & bookba + adjamnt & ",transdate='" & Date & "',vno='" & _
                        txtvoucherno & "',period='" & month(Date) & "',auditid='" & teller & "',auditdate='" & Now & _
                        "',moduleid=2 where accno='" & glaccno & "'"
                        myclass.save sql
                    End If
                End If
            End If
            '//let also affect the teller for transfer to kericho and other branches
            sql = "insert into [teller transactions] (tellername,Cubiclenumber,Deposits,Withdrawals,Accountnumber,tgno,TransactionDate,vno,Posted,Locked,Auditid,audittime,transdescription,balance,name,accname,commission,printed,cash) "
            sql = sql & "values('" & User & "','" & cub & "',0," & adjamnt & ",'" & Txtaccno & "','" & custno & "','" & Date & "','" & txtvoucherno & "',0,0,'" & User & "','" & Now & "','" & desc & "',0,'" & AccName & "','" & AccName & "',0,0,1)"
            'myclass.save sql
            sql = "INSERT INTO CashTransactions"
            sql = sql & "   (Customerno, AccNo, AccName, Amount, Transdescription, Transdate, Commission, Chequeno, Transtype, vno, Posted, Locked, Auditid, AuditTime,userName)"
            sql = sql & " VALUES ('" & custno & "', '" & Txtaccno & "', '" & AccName & "', " & adjamnt & ", 'adj', '" & Date & "', 0, 'NON', 'DR', '" & txtvoucherno & "', 0, 0, '" & User & "', '" & Now & "', '" & User & "')"
            ' cn.Execute sql
        End If
    End If
    Case "IDNo"
        
Set rss = CreateObject("adodb.recordset")
rss.Open "SELECT   *  FROM CustomerBalance where IDno='" & txtidno & "' ORDER BY CustomerBalanceid DESC", cn

If rss.EOF Then Exit Sub
If Not IsNull(rss!ACCNO) Then ACCNO = rss!ACCNO
If Not IsNull(rss!availablebalance) Then avail = rss!availablebalance
If Not IsNull(rss!customerbalanceid) Then id = rss!customerbalanceid
If Not IsNull(rss!AccName) Then AccName = rss!AccName
If Not IsNull(rss!CustomerNo) Then custno = rss!CustomerNo

If optcredit = True Then

sql = ""

sql = "set dateformat dmy insert into customerbalance (customerno,accname,amount,availablebalance,accno,transdescription,transdate,commission,period,transType,posted,locked,status,vno,auditid,auditdate,moduleid) "
sql = sql & " values ('" & custno & "','" & AccName & "'," & adjamnt & "," & avail + adjamnt & ",'" & Txtaccno & "','" & desc & "','" & DTP & "',0,'" & month(Date) & "','CR',0,0,0,'" & txtvoucherno & "','" & User & "','" & Now & "','3' )"
myclass.save sql

sql = "set dateformat dmy update cub set amount=" & adjamnt & ",Active=1,transdescription='" & desc & "',availablebalance=" & avail + adjamnt & ",transdate='" & Date & "',vno='" & txtvoucherno & "',period='" & month(Date) & "',auditid='" & teller & "',auditdate='" & Date & "',moduleid=2,active=1 where accno='" & ACCNO & "'"
myclass.save sql
'//let also affect the teller for transfer to kericho and other branches
sql = "set dateformat dmy insert into [teller transactions] (tellername,Cubiclenumber,Deposits,Withdrawals,Accountnumber,tgno,TransactionDate,vno,Posted,Locked,Auditid,audittime,transdescription,balance,name,accname,commission,printed,cash) "
sql = sql & "values('" & User & "','" & cub & "'," & adjamnt & ",0,'" & Txtaccno & "','" & custno & "','" & Date & "','" & txtvoucherno & "',0,0,'" & User & "','" & Now & "','" & desc & "',0,'" & AccName & "','" & AccName & "',0,0,1)"
myclass.save sql

ElseIf Optdebit = True Then

sql = ""

sql = "set dateformat dmy insert into customerbalance (customerno,accname,amount,availablebalance,accno,transdescription,transdate,commission,period,transType,posted,locked,status,vno,auditid,auditdate,moduleid) "
sql = sql & " values ('" & custno & "','" & AccName & "'," & adjamnt & "," & avail - adjamnt & ",'" & Txtaccno & "','" & desc & "','" & DTP & "',0,'" & month(Date) & "','DR',0,0,0,'" & txtvoucherno & "','" & User & "','" & Now & "','3' )"
myclass.save sql

sql = "update cub set amount=" & adjamnt & ",Active=1,transdescription='" & desc & "',availablebalance=" & avail - adjamnt & ",transdate='" & Date & "',vno='" & txtvoucherno & "',period='" & month(Date) & "',auditid='" & teller & "',auditdate='" & Date & "',moduleid=2,active=1 where accno='" & ACCNO & "'"
myclass.save sql

sql = "insert into [teller transactions] (tellername,Cubiclenumber,Deposits,Withdrawals,Accountnumber,tgno,TransactionDate,vno,Posted,Locked,Auditid,audittime,transdescription,balance,name,accname,commission,printed,cash) "
sql = sql & "values('" & User & "','" & cub & "',0," & adjamnt & ",'" & Txtaccno & "','" & custno & "','" & Date & "','" & txtvoucherno & "',0,0,'" & User & "','" & Now & "','" & desc & "',0,'" & AccName & "','" & AccName & "',0,0,1)"
myclass.save sql

End If
Case "MemberNo"
Set rss = CreateObject("adodb.recordset")
rss.Open "SELECT   *  FROM CustomerBalance where custno ='" & TXTTGNO & "' ORDER BY CustomerBalanceid DESC", cn

If rss.EOF Then Exit Sub
If Not IsNull(rss!ACCNO) Then ACCNO = rss!ACCNO
If Not IsNull(rss!availablebalance) Then avail = rss!availablebalance
If Not IsNull(rss!customerbalanceid) Then id = rss!customerbalanceid
If Not IsNull(rss!AccName) Then AccName = rss!AccName
If Not IsNull(rss!CustomerNo) Then custno = rss!CustomerNo
If optcredit = True Then
sql = ""
sql = "set dateformat dmy insert into customerbalance (customerno,accname,amount,availablebalance,accno,transdescription,transdate,commission,period,transType,posted,locked,status,vno,auditid,auditdate,moduleid) "
sql = sql & " values ('" & custno & "','" & AccName & "'," & adjamnt & "," & avail + adjamnt & ",'" & Txtaccno & "','" & desc & "','" & DTP & "',0,'" & month(Date) & "','CR',0,0,0,'" & txtvoucherno & "','" & User & "','" & Now & "','3' )"
myclass.save sql
sql = "set dateformat dmy update cub set amount=" & adjamnt & ",Active=1,transdescription='" & desc & "',availablebalance=" & avail + adjamnt & ",transdate='" & Date & "',vno='" & txtvoucherno & "',period='" & month(Date) & "',auditid='" & teller & "',auditdate='" & Date & "',moduleid=2,active=1 where accno='" & ACCNO & "'"
myclass.save sql
'//let also affect the teller for transfer to kericho and other branches
sql = "set dateformat dmy insert into [teller transactions] (tellername,Cubiclenumber,Deposits,Withdrawals,Accountnumber,tgno,TransactionDate,vno,Posted,Locked,Auditid,audittime,transdescription,balance,name,accname,commission,printed,cash) "
sql = sql & "values('" & User & "','" & cub & "'," & adjamnt & ",0,'" & Txtaccno & "','" & custno & "','" & Date & "','" & txtvoucherno & "',0,0,'" & User & "','" & Now & "','" & desc & "',0,'" & AccName & "','" & AccName & "',0,0,1)"
myclass.save sql

ElseIf Optdebit = True Then

sql = ""

sql = "set dateformat dmy insert into customerbalance (customerno,accname,amount,availablebalance,accno,transdescription,transdate,commission,period,transType,posted,locked,status,vno,auditid,auditdate,moduleid) "
sql = sql & " values ('" & custno & "','" & AccName & "'," & adjamnt & "," & avail - adjamnt & ",'" & Txtaccno & "','" & desc & "','" & DTP & "',0,'" & month(Date) & "','DR',0,0,0,'" & txtvoucherno & "','" & User & "','" & Now & "','3' )"
myclass.save sql

sql = "set dateformat dmy update cub set amount=" & adjamnt & ",Active=1,transdescription='" & desc & "',availablebalance=" & avail - adjamnt & ",transdate='" & Date & "',vno='" & txtvoucherno & "',period='" & month(Date) & "',auditid='" & teller & "',auditdate='" & Date & "',moduleid=2,active=1 where accno='" & ACCNO & "'"
myclass.save sql

sql = "set dateformat dmy insert into [teller transactions] (tellername,Cubiclenumber,Deposits,Withdrawals,Accountnumber,tgno,TransactionDate,vno,Posted,Locked,Auditid,audittime,transdescription,balance,name,accname,commission,printed,cash) "
sql = sql & "values('" & User & "','" & cub & "',0," & adjamnt & ",'" & Txtaccno & "','" & custno & "','" & Date & "','" & txtvoucherno & "',0,0,'" & User & "','" & Now & "','" & desc & "',0,'" & AccName & "','" & AccName & "',0,0,1)"
myclass.save sql
End If
Case "PayrollNo"
Set rss = CreateObject("adodb.recordset")
rss.Open "SELECT   *  FROM CustomerBalance where PAYROLLNO='" & txtPayNo & "' ORDER BY CustomerBalanceid DESC", cn

If rss.EOF Then Exit Sub
If Not IsNull(rss!ACCNO) Then ACCNO = rss!ACCNO
If Not IsNull(rss!availablebalance) Then avail = rss!availablebalance
If Not IsNull(rss!customerbalanceid) Then id = rss!customerbalanceid
If Not IsNull(rss!AccName) Then AccName = rss!AccName
If Not IsNull(rss!CustomerNo) Then custno = rss!CustomerNo
If optcredit = True Then

sql = ""

sql = "set dateformat dmy insert into customerbalance (customerno,accname,amount,availablebalance,accno,transdescription,transdate,commission,period,transType,posted,locked,status,vno,auditid,auditdate,moduleid) "
sql = sql & " values ('" & custno & "','" & AccName & "'," & adjamnt & "," & avail + adjamnt & ",'" & Txtaccno & "','" & desc & "','" & DTP & "',0,'" & month(Date) & "','CR',0,0,0,'" & txtvoucherno & "','" & User & "','" & Now & "','3' )"
myclass.save sql

sql = "set dateformat dmy update cub set amount=" & adjamnt & ",Active=1,transdescription='" & desc & "',availablebalance=" & avail + adjamnt & ",transdate='" & Date & "',vno='" & txtvoucherno & "',period='" & month(Date) & "',auditid='" & teller & "',auditdate='" & Date & "',moduleid=2,active=1 where accno='" & ACCNO & "'"
myclass.save sql

'//let also affect the teller for transfer to kericho and other branches
sql = "set dateformat dmy insert into [teller transactions] (tellername,Cubiclenumber,Deposits,Withdrawals,Accountnumber,tgno,TransactionDate,vno,Posted,Locked,Auditid,audittime,transdescription,balance,name,accname,commission,printed,cash) "
sql = sql & "values('" & User & "','" & cub & "'," & adjamnt & ",0,'" & Txtaccno & "','" & custno & "','" & Date & "','" & txtvoucherno & "',0,0,'" & User & "','" & Now & "','" & desc & "',0,'" & AccName & "','" & AccName & "',0,0,1)"
myclass.save sql

ElseIf Optdebit = True Then

sql = ""

sql = "set dateformat dmy insert into customerbalance (customerno,accname,amount,availablebalance,accno,transdescription,transdate,commission,period,transType,posted,locked,status,vno,auditid,auditdate,moduleid) "
sql = sql & " values ('" & custno & "','" & AccName & "'," & adjamnt & "," & avail - adjamnt & ",'" & Txtaccno & "','" & desc & "','" & DTP & "',0,'" & month(Date) & "','DR',0,0,0,'" & txtvoucherno & "','" & User & "','" & Now & "','3' )"
myclass.save sql

sql = "set dateformat dmy update cub set amount=" & adjamnt & ",Active=1,transdescription='" & desc & "',availablebalance=" & avail - adjamnt & ",transdate='" & Date & "',vno='" & txtvoucherno & "',period='" & month(Date) & "',auditid='" & teller & "',auditdate='" & Date & "',moduleid=2,active=1 where accno='" & ACCNO & "'"
myclass.save sql

sql = "set dateformat dmy insert into [teller transactions] (tellername,Cubiclenumber,Deposits,Withdrawals,Accountnumber,tgno,TransactionDate,vno,Posted,Locked,Auditid,audittime,transdescription,balance,name,accname,commission,printed,cash) "
sql = sql & "values('" & User & "','" & cub & "',0," & adjamnt & ",'" & Txtaccno & "','" & custno & "','" & Date & "','" & txtvoucherno & "',0,0,'" & User & "','" & Now & "','" & desc & "',0,'" & AccName & "','" & AccName & "',0,0,1)"
myclass.save sql
End If
End Select
If chkbackofficetransactions = vbChecked Then
'//get loan
If cbomemtrans = "Loan Repayment" Then
getloaninfo
End If
'get interest

If cbomemtrans = "Int on loan repayment" Then
getloanbal
getulmt
Dim cnnn As ADODB.Connection
Set cnnn = New ADODB.Connection
Dim korir As String
korir = SelectedDsn
cnnn.Open korir

sql = ""
sql = "SET DATEFORMAT DMY INSERT INTO REPAY"
sql = sql & "(LoanNo, DateReceived, PaymentNo, Amount, Principal, "
sql = sql & "Interest, IntrCharged, IntrOwed, LoanBalance, Locked, Posted, Accrued, Remarks, AuditID,"
sql = sql & " AuditTime)"
sql = sql & " VALUES('" & loan & "', '" & Date & "', " & ulmt & ", " & txtadjamnt & ", 0,"
sql = sql & " " & txtadjamnt & " , 0, 0, "
sql = sql & " " & loanbal.balance & ", 'No', 'No', 'No', 'Received from teller', '" & User & "', '" & Now & "')"
cnnn.Execute sql
' Int on loan repayment
End If
' get shares
If cbomemtrans = "Shares" Then
'//
SAVESHARES
End If
End If
chkbackofficetransactions = vbUnchecked
txtadjamnt = 0
optcredit = False
Optdebit = False
lblname = ""
lblaccname = ""
lblavail = ""
lblname = ""
TXTTGNO = ""
txtadjamnt = 0
txtamount = 0
lbluncleared = ""
'txtcomm = 0
Txtvou = ""
txtMemberNo = ""
MsgBox "Transactions Complete", vbInformation, "Accounts Updated"
End Sub
Private Sub gettotalshare()

Dim r As Object
    setRsr r, "select  totalshares,memberno,transdate from shares where memberno = '" & txtMemberNo & "'"
    If r.EOF Then Exit Sub
   With tos
      If Not IsNull(r!totalshares) Then .totalshares = r!totalshares
      If Not IsNull(r!memberno) Then .meberno = r!memberno
      If Not IsNull(r!transdate) Then .transdate = r!transdate
   End With
   getrefno
 r.Close
Set r = Nothing
End Sub
Private Sub getrefno()
ref = 11
Dim r As Object
    setRsr r, "select  refno from contrib where memberno = '" & share.memberno & "'order by refno desc"
    If r.EOF Then Exit Sub
   With tos
      If Not IsNull(r!RefNo) Then ref = r!RefNo
   End With
   ref = ref + 1
 r.Close
Set r = Nothing
End Sub

Private Sub SAVESHARES()
Set myclass = New cdbase


Dim cnnn As ADODB.Connection
        Set cnnn = New ADODB.Connection
        Dim korir As String
        korir = SelectedDsn
        cnnn.Open korir
Dim re As Integer
gettotalshare
'getrefno

sql = ""
sql = "set dateformat dmy INSERT INTO CONTRIB"
sql = sql & "(MemberNo, ContrDate, RefNo, Amount, ShareBal, TransBy, ChequeNo, "
sql = sql & " Locked, Posted, Remarks, AuditID, AuditTime)"
sql = sql & "VALUES     ('" & txtMemberNo & "', '" & Date & "', '" & ref & "',"
sql = sql & "" & txtadjamnt & ", " & tos.totalshares + txtadjamnt & " , 'from " & Txtaccno & "', '', 'No', 'No', 'no', '" & User & "', '" & Now & "') "
cnnn.Execute sql
sql = ""
sql = "  Update shares"
sql = sql & " Set totalshares = " & tos.totalshares + txtadjamnt & " where memberno='" & txtMemberNo & "' "
cnnn.Execute sql
  
    'Teller trans tabl
          
          
End Sub

Private Sub cmdcancel_Click()
Unload Me
End Sub

Private Sub cmdclose_Click()

    Unload Me

End Sub



Private Sub cmddelete_Click()
Dim myrec As Object
Dim rss As Object
Dim ans, an
Dim amt As Long
Set myclass = New cdbase
Set cn = CreateObject("adodb.connection")
Provider = myclass.OpenCon
cn.Open Provider, "atm", "atm"
Set myclass = New cdbase
    Set cn = CreateObject("adodb.connection")
    Provider = myclass.OpenCon
   
   cn.Open Provider, "atm", "atm"
    'If Opttgno = True Then
                  Set rs = CreateObject("adodb.recordset")
                 rs.Open "SELECT   *  FROM CustomerBalance where customerno='" & TXTTGNO & "'and vno='" & txtvoucherno & "' ORDER BY CustomerBalanceid ", cn
                 
                 If rs.EOF Then Exit Sub
                 If txtvoucherno = "" Then
                 MsgBox "Please Enter the corresponding Voucher Number used  Before you Proceed", vbInformation, "Member Transactions"
                 Exit Sub
                 End If
                 ans = MsgBox("Are You sure you want Delete ,Voucher No " & txtvoucherno & "?", vbYesNo, "Deleting Bank Code")
                If ans = vbYes Then
                    an = MsgBox("Are You sure want ", vbYesNo, "MemberTransactions")
                 If an = vbYes Then
                    sql = ""
                      sql = "delete from customerbalance where vno ='" & txtvoucherno & "' "
                       myclass.Delete sql
                       sql = ""
                       sql = "delete from [Teller Transactions] where vno= '" & txtvoucherno & "'"
                       myclass.Delete sql
                       End If
                 End If
'ElseIf Optaccountno = True Then
  rs.Open "SELECT   *  FROM CustomerBalance where customerno='" & TXTTGNO & "'and vno='" & txtvoucherno & "' ORDER BY CustomerBalanceid ", cn
                 
                 If rs.EOF Then Exit Sub
                 If txtvoucherno = "" Then
                 MsgBox "Please Enter the corresponding Voucher Number used  Before you Proceed", vbInformation, "Member Transactions"
                 Exit Sub
                 End If
                 ans = MsgBox("Are You sure you want Delete ,Voucher No " & txtvoucherno & "?", vbYesNo, "Deleting Bank Code")
                If ans = vbYes Then
                    an = MsgBox("Are You sure want ", vbYesNo, "MemberTransactions")
                 If an = vbYes Then
                    sql = ""
                      sql = "delete from customerbalance where vno ='" & txtvoucherno & "' "
                       myclass.Delete sql
                       sql = ""
                       sql = "delete from [Teller Transactions] where vno= '" & txtvoucherno & "'"
                       myclass.Delete sql
                       End If
                 End If


On Error Resume Next
    rs.Close
    Form_Load
End Sub

Private Sub cmdok1_Click()
Dim myclass As cdbase
Set myclass = New cdbase
Set cn = CreateObject("adodb.connection")
Provider = myclass.OpenCon
cn.Open Provider, "atm", "atm"

Dim ans As VbMsgBoxResult
Dim ats
Dim an As VbMsgBoxResult
Dim sup As Object
Dim bals As VbMsgBoxResult
Dim over As VbMsgBoxResult
Dim mynull As Boolean
Dim totCharge As Currency
Dim CanTransact As Boolean
Dim validacc As Boolean
teller = GetSetting("FOSAdll", "Teller", "Name")
txtpassword = modsecurity.Decript_String(txtpassword)

Set rs = CreateObject("adodb.recordset")
sql = ""
sql = "select * from useraccounts where password='" & txtpassword & "' and UserLoginIDs='" & User & "'"
rs.Open sql, cn

If rs.EOF Then

MsgBox "You don't have Access to Member Transactions Module", vbExclamation, "Cash Transaction"
'If Not IsNull(rs!username) Then teller = rs!username
teller = GetSetting("FOSAdll", "Teller", "Name")
sql = ""
sql = "SET DATEFORMAT dmy INSERT INTO AuditTable"
sql = sql & "(UserName, LoginDate, LoginTime, UserTransaction, LogoffTime, moduleid)"
sql = sql & "VALUES     ('" & User & "', '" & Date & "', '" & Time & "', 'Adj-membertransaction(login fail)', '" & Time & "', '2')"
cn.Execute sql
On Error Resume Next

txtpassword = ""
txtpassword.SetFocus
Exit Sub
Else
If Not IsNull(rs!username) Then teller = rs!username
frasalk.Visible = False
sql = "SET DATEFORMAT DMY INSERT INTO AuditTable"
sql = sql & "  (UserName, LoginDate, LoginTime, UserTransaction, LogoffTime, moduleid)"
sql = sql & "VALUES ('" & User & "', '" & Date & "', '" & Time & "', 'Adj-membertransaction(login successfull)', '" & Time & "', '2')"
cn.Execute sql
End If
teller = teller
End Sub

Private Sub cmdrefresh_Click()
    Exit Sub
    LBLGLCONTRA = ""
    LBLGLCOMMISSION = ""
    lblglstamp = ""
    glnamE1 = ""
    glnamecom1 = ""
    glnamestamp1 = ""
    cbocharge1 = ""
    cbocharge2 = ""
    cbocharge3 = ""
    cbocharge4 = ""
    lbluncleared = ""
    Dim myclass As cdbase
    Set myclass = New cdbase
    Set cn = CreateObject("adodb.connection")
    Provider = myclass.OpenCon
   cn.Open Provider, "atm", "atm"
    Dim rsproc As Recordset
    Set rsproc = New ADODB.Recordset
    Dim sssql As String
    sssql = "proc_rebuild '" & Txtaccno & "'"
    'ssql = "sp_inquiry '" & txtaccno & "'"
    Set rsproc = cn.Execute(sssql)
    'rebuild_accno Txtaccno
    txtAccno_Change
End Sub

Private Sub cmdreverse_Click()

'If optcredit = True Then
'    sql = ""
'
'    sql = "insert into customerbalance (customerno,accname,amount,availablebalance,accno,transdescription,transdate,commission,period,transType,posted,locked,status,vno,auditid,auditdate,moduleid,accd) "
'    sql = sql & " values ('" & custno & "','" & accname & "'," & adjamnt & "," & avail + adjamnt & ",'" & Txtaccno & "','" & DESC & "','" & DTP & "',0,'" & month(Date) & "','CR',0,0,0,'" & txtvoucherno & "','" & user & "','" & Now & "','3','" & glaccno & "' )"
'    myclass.Save sql
'
'    sql = ""
'    sql = "update cub set amount=" & adjamnt & ",transdescription='" & DESC & "',availablebalance=" & avail + adjamnt & ",transdate='" & Date & "',vno='" & txtvoucherno & "',period='" & month(Date) & "',auditid='" & teller & "',auditdate='" & Now & "',moduleid=2 where accno='" & Txtaccno & "'"
'    myclass.Save sql
'
'    DESCR = "From - " & accname & ""
'    If glnamE1 <> "" Then
'    sql = ""
'    sql = "insert into customerbalance (customerno,accname,amount,availablebalance,accno,transdescription,transdate,commission,period,transType,posted,locked,status,vno,auditid,auditdate,moduleid,accd) "
'    sql = sql & " values ('" & glmemno & "','" & glnamE1 & "'," & adjamnt & "," & bookba + adjamnt & ",'" & glaccno & "','" & DESCR & "','" & DTP & "',0,'" & month(Date) & "','DR',0,0,0,'" & txtvoucherno & "','" & user & "','" & Now & "','3','" & Txtaccno & "' )"
'    myclass.Save sql
'
'    sql = ""
'    sql = "update cub set amount=" & adjamnt & ",transdescription='" & DESCR & "',availablebalance=" & bookba + adjamnt & ",transdate='" & Date & "',vno='" & txtvoucherno & "',period='" & month(Date) & "',auditid='" & teller & "',auditdate='" & Now & "',moduleid=2 where accno='" & glaccno & "'"
'    myclass.Save sql
'    End If
'    If Optdebit = True Then
'    sql = ""
'
'    sql = "insert into customerbalance (customerno,accname,amount,availablebalance,accno,transdescription,transdate,commission,period,transType,posted,locked,status,vno,auditid,auditdate,moduleid,accd) "
'    sql = sql & " values ('" & custno & "','" & accname & "'," & adjamnt & "," & avail + adjamnt & ",'" & Txtaccno & "','" & DESC & "','" & DTP & "',0,'" & month(Date) & "','CR',0,0,0,'" & txtvoucherno & "','" & user & "','" & Now & "','3','" & glaccno & "' )"
'    myclass.Save sql
'
'    sql = ""
'    sql = "update cub set amount=" & adjamnt & ",transdescription='" & DESC & "',availablebalance=" & avail + adjamnt & ",transdate='" & Date & "',vno='" & txtvoucherno & "',period='" & month(Date) & "',auditid='" & teller & "',auditdate='" & Now & "',moduleid=2 where accno='" & Txtaccno & "'"
'    myclass.Save sql
'
'    DESCR = "From - " & accname & ""
'    If glnamE1 <> "" Then
'    sql = ""
'    sql = "insert into customerbalance (customerno,accname,amount,availablebalance,accno,transdescription,transdate,commission,period,transType,posted,locked,status,vno,auditid,auditdate,moduleid,accd) "
'    sql = sql & " values ('" & glmemno & "','" & glnamE1 & "'," & adjamnt & "," & bookba + adjamnt & ",'" & glaccno & "','" & DESCR & "','" & DTP & "',0,'" & month(Date) & "','DR',0,0,0,'" & txtvoucherno & "','" & user & "','" & Now & "','3','" & Txtaccno & "' )"
'    myclass.Save sql
'
'    sql = ""
'    sql = "update cub set amount=" & adjamnt & ",transdescription='" & DESCR & "',availablebalance=" & bookba + adjamnt & ",transdate='" & Date & "',vno='" & txtvoucherno & "',period='" & month(Date) & "',auditid='" & teller & "',auditdate='" & Now & "',moduleid=2 where accno='" & glaccno & "'"
'    myclass.Save sql
'    End If
End Sub

Private Sub cmdsave_Click()
    Set myclass = New cdbase
    Set cn = CreateObject("adodb.connection")
    Provider = myclass.OpenCon
   cn.Open Provider, "atm", "atm"
    Dim bal As Currency
    Dim I As Integer
  
    Dim tel As Object
    Dim cub As String
    
   ' teller = GetSetting("FOSAdll", "Teller", "Name")
    cub = GetSetting("FOSAdll", "Teller", "Cubie Number")
    Set myclass = New cdbase
    Set cn = CreateObject("adodb.connection")
    Provider = myclass.OpenCon
   cn.Open Provider, "atm", "atm"
    Set rs = CreateObject("adodb.recordset")
           If acccr = "" Then
           MsgBox "Please enter the account To be credited", vbInformation, "Member Transactions"
           Exit Sub
           End If
           
           If accdr = "" Then
           MsgBox "Please enter the account To be Debited", vbInformation, "Member Transactions"
           Exit Sub
           End If
'
           If Not IsNumeric(txtamount) Then
           MsgBox "Please Enter Values ", vbInformation, "Member transactions"
           Exit Sub
           End If
sql = "select balance from [teller transactions] where tellername='" & teller & "' and transactiondate='" & Date & "'order by tellerid desc"
Set tel = CreateObject("adodb.recordset")
tel.Open sql, cn
If tel.EOF Then
 If Not IsNull(tel!balance) Then bal = 0
Else
If Not IsNull(tel!balance) Then bal = tel!balance
End If

If acccr <> "" Then

AVAIL1 = AVAIL1 + CCur(txtamount)
'bal = bal + CCur(txtamount)
sql = ""
sql = "set dateformat dmy insert into customerbalance (customerno,accname,amount,availablebalance,accno,transdescription,transdate,commission,period,transType,posted,locked,status,vno,auditid,auditdate,moduleid) "
sql = sql & " values ('" & custno1 & "','" & accname1 & "'," & CCur(txtamount) & "," & AVAIL1 & ",'" & acccr & "','" & txtnarationcr & " ','" & Date & "',0,'" & month(Date) & "','CR',0,0,0,'" & Txtvou & "','" & User & "','" & Now & "','3' )"
myclass.save sql

        sql = "set dateformat dmy INSERT INTO CashTransactions"
        sql = sql & "   (Customerno, AccNo, AccName, Amount, Transdescription, Transdate, Commission, Chequeno, Transtype, vno, Posted, Locked, Auditid, AuditTime,userName)"
        sql = sql & " VALUES ('" & custno1 & "', '" & acccr & "', '" & accname1 & "', " & CCur(txtamount) & ", '" & txtreason & "', '" & Date & "', 0, 'NON', 'CR', '" & Txtvou & "', 0, 0, '" & User & "', '" & Now & "', '" & User & "')"
        cn.Execute sql



sql = "set dateformat dmy update cub set amount=" & CCur(txtamount) & ",Active=1,transdescription='" & txtreason & "',availablebalance=" & AVAIL1 & ",transdate='" & Date & "',vno='" & Txtvou & "',period='" & month(Date) & "',auditid='" & teller & "',auditdate='" & Date & "',moduleid=2,active=1 where accno='" & acccr & "'"
myclass.save sql
'// affect the teller accounts

 If OpCR = True Then '// Means we are correcting a mis-posted credit
   For I = 1 To 2 '// We want to insert only twice into the teller transactions table
    If I = 1 Then '//if its the first time,increase the balance
      bal = bal + CCur(txtamount)
      sql = "set dateformat dmy insert into [teller transactions] (tellername,Cubiclenumber,Deposits,Withdrawals,Accountnumber,tgno,TransactionDate,vno,Posted,Locked,Auditid,audittime,transdescription,balance,payno,idno,name,accname,commission,printed,cash) "
      sql = sql & "values('" & User & "','" & cub & "'," & CCur(txtamount) & ",0,'" & accdr & "','" & custno2 & "','" & Date & "','" & Txtvou & "',0,0,'" & User & "','" & Now & "','Correction'," & bal & ",'" & payno2 & "','" & idno2 & "','" & accname2 & "','" & accname2 & "',0,0,1)"
     ' myclass.Save sql
    Else '//this is for the sake of neutralising i.e. to make it zero
      sql = "set dateformat dmy insert into [teller transactions] (tellername,Cubiclenumber,Deposits,Withdrawals,Accountnumber,tgno,TransactionDate,vno,Posted,Locked,Auditid,audittime,transdescription,balance,payno,idno,name,accname,commission,printed,cash) "
      sql = sql & "values('" & User & "','" & cub & "'," & CCur(txtamount) & ",0,'" & acccr & "','" & custno1 & "','" & Date & "','" & Txtvou & "',0,0,'" & User & "','" & Now & "','Correction'," & bal & ",'" & payno1 & "','" & idno1 & "','" & accname1 & "','" & accname1 & "',0,0,1)"
     ' myclass.Save sql
    End If
  Next
Else '//Correcting mis-posted debit
   For I = 1 To 2 '// We want to insert only twice into the teller transactions table
    If I = 1 Then '//if its the first time,increase the balance
      bal = bal - CCur(txtamount)
      sql = "set dateformat dmy insert into [teller transactions] (tellername,Cubiclenumber,Deposits,Withdrawals,Accountnumber,tgno,TransactionDate,vno,Posted,Locked,Auditid,audittime,transdescription,balance,payno,idno,name,accname,commission,printed,cash) "
      sql = sql & "values('" & User & "','" & cub & "',0," & CCur(txtamount) & ",'" & acccr & "','" & custno1 & "','" & Date & "','" & Txtvou & "',0,0,'" & User & "','" & Now & "','Correction'," & bal & ",'" & payno1 & "','" & idno1 & "','" & accname1 & "','" & accname1 & "',0,0,1)"
      'myclass.Save sql
    Else '//this is for the sake of neutralising i.e. to make it zero
      sql = "set dateformat dmy insert into [teller transactions] (tellername,Cubiclenumber,Deposits,Withdrawals,Accountnumber,tgno,TransactionDate,vno,Posted,Locked,Auditid,audittime,transdescription,balance,payno,idno,name,accname,commission,printed,cash) "
      sql = sql & "values('" & User & "','" & cub & "',0," & CCur(txtamount) & ",'" & accdr & "','" & custno2 & "','" & Date & "','" & Txtvou & "',0,0,'" & User & "','" & Now & "','Correction'," & bal & ",'" & payno2 & "','" & idno2 & "','" & accname2 & "','" & accname2 & "',0,0,1)"
      'myclass.Save sql
    End If
  Next
End If
 'End If

End If

If accdr <> "" Then
bal = bal - CCur(txtamount)

AVAIL2 = AVAIL2 - CCur(txtamount)
sql = ""
sql = "set dateformat dmy insert into customerbalance (customerno,accname,amount,availablebalance,accno,transdescription,transdate,commission,period,transType,posted,locked,status,vno,auditid,auditdate,moduleid) "
sql = sql & " values ('" & custno2 & "','" & accname2 & "'," & CCur(txtamount) & "," & AVAIL2 & ",'" & accdr & "','" & txtreason & " ','" & Date & "',0,'" & month(Date) & "','DR',0,0,0,'" & Txtvou & "','" & User & "','" & Now & "','3' )"
myclass.save sql

        sql = "set dateformat dmy INSERT INTO CashTransactions"
        sql = sql & "   (Customerno, AccNo, AccName, Amount, Transdescription, Transdate, Commission, Chequeno, Transtype, vno, Posted, Locked, Auditid, AuditTime,userName)"
        sql = sql & " VALUES ('" & custno2 & "', '" & accdr & "', '" & accname2 & "', " & txtamount & ", '" & txtreason & "', '" & Date & "', 0, 'NON', 'DR', '" & Txtvou & "', 0, 0, '" & User & "', '" & Now & "', '" & User & "')"
        cn.Execute sql


sql = "set dateformat dmy update cub set amount=" & CCur(txtamount) & ",Active=1,transdescription='" & txtreason & "',availablebalance=" & AVAIL2 & ",transdate='" & Date & "',vno='" & Txtvou & "',period='" & month(Date) & "',auditid='" & teller & "',auditdate='" & Date & "',moduleid=2,active=1 where accno='" & accdr & "'"
myclass.save sql
'// affects the tellers accounts
'
'    sql = "insert into [teller transactions] (tellername,Cubiclenumber,Deposits,Withdrawals,Accountnumber,tgno,TransactionDate,vno,Posted,Locked,Auditid,audittime,transdescription,balance,payno,idno,name,accname,commission,printed) "
'    sql = sql & "values('" & user & "','" & cub & "',0," & CCur(txtamount) & ",'" & accdr & "','" & custno2 & "','" & Date & "','" & txtvou & "',0,0,'" & user & "','" & Now & "','Correction'," & bal & ",'" & payno2 & "','" & idno2 & "','" & accname2 & "','" & accname2 & "',0,0)"
'    Myclass.Save sql
'    sql = ""

End If

acccr = ""
accdr = ""
txtreason = ""
Txtvou = ""
txtamount = ""
lblname1 = ""
Lblname2 = ""
lblamount1 = ""
lblamount2 = ""
 txtnarationcr = ""
 lbluncleared = ""
    'Form_Load
End Sub

Private Sub cmdTrans_Click()
    fraTrans.Visible = False
End Sub

Private Sub cmdTransSetup_Click()
    frmtransactioncode.Show vbModal, Me
End Sub

Private Sub Form_Initialize()
    On Error GoTo SysError
    lvememtrans.Width = Me.ScaleWidth
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub


Private Sub LBLGLCOMMISSION_Change()
Dim Z, S, U
    Dim rs As Recordset
    'frmsearchaccounts.Show vbModal
    Z = LBLGLCOMMISSION
    If Z <> "" Then
        LBLGLCOMMISSION = Z
        glcommission = Z
        End If
          
    Set myclass = New cdbase
    Set cn = CreateObject("adodb.connection")
    Provider = myclass.OpenCon
   cn.Open Provider, "atm", "atm"
sql = "select * from cuB where ACCno='" & Z & "'"
Set rs = New ADODB.Recordset
rs.Open sql, cn
'Dim glnamecom As String 'FOR COMMISSION
'Dim glcommemno As String 'FOR COMMISSION
'Dim glcomidno As String 'FOR COMMISSION
'Dim glcompayno As String 'FOR COMMISSION
'Dim GLCOMMISSION As String
If Not rs.EOF Then
If Not IsNull(rs.Fields("availablebalance")) Then glcommission = rs.Fields("availablebalance")
If Not IsNull(rs.Fields("accountname")) Then glnamecom1 = rs.Fields("name")
If Not IsNull(rs.Fields("idno")) Then glcomidno = rs.Fields("idno")
If Not IsNull(rs.Fields("memberno")) Then glcommemno = rs.Fields("memberno")
If Not IsNull(rs.Fields("payno")) Then glcompayno = rs.Fields("payno")
bookba1 = cub_balance(glcommission)
Else
glnamecom1 = ""
End If
'bookba1 = cub_balance(glcommission)

End Sub

Private Sub LBLGLCONTRA_Change()
    Dim Z, S, U
    Dim rs As Recordset
    'frmsearchaccounts.Show vbModal
    Z = LBLGLCONTRA
    If Z <> "" Then
    LBLGLCONTRA = Z
    glcomm = Z
    End If
    'frmsearchnewacc.Show vbModal
    ' Z = strName
    Set myclass = New cdbase
    Set cn = CreateObject("adodb.connection")
    Provider = myclass.OpenCon
   cn.Open Provider, "atm", "atm"
    sql = "select * from cuB where ACCNO='" & glcomm & "'"
    Set rs = New ADODB.Recordset
    rs.Open sql, cn
    If Not rs.EOF Then
    If Not IsNull(rs.Fields("accno")) Then glaccno = rs.Fields("accno")
    If Not IsNull(rs.Fields("accountname")) Then glnamE1 = rs.Fields("name")
    If Not IsNull(rs.Fields("idno")) Then glidno = rs.Fields("idno")
    If Not IsNull(rs.Fields("memberno")) Then glmemno = rs.Fields("memberno")
    If Not IsNull(rs.Fields("payno")) Then glpayno = rs.Fields("payno")
    bookba = cub_balance(glaccno)
    Else
    glnamE1 = ""
    End If
    'glPremium = Scheme_GL_Field(AccountCode, "glPremium")
    'bookba = cub_balance(glaccno)

End Sub

Private Sub lblglstamp_Change()
Dim Z, S, U
    Dim rs As Recordset



    'frmsearchaccounts.Show vbModal
    Z = lblglstamp
    If Z <> "" Then
        lblglstamp = Z
        
        End If
        
         Set myclass = New cdbase
    Set cn = CreateObject("adodb.connection")
    Provider = myclass.OpenCon
   cn.Open Provider, "atm", "atm"
sql = "select * from cuB where ACCno='" & lblglstamp & "'"
Set rs = New ADODB.Recordset
rs.Open sql, cn
If Not rs.EOF Then
If Not IsNull(rs.Fields("availablebalance")) Then glstampbal = rs.Fields("availablebalance")
If Not IsNull(rs.Fields("accountname")) Then glnamestamp1 = rs.Fields("name")
If Not IsNull(rs.Fields("idno")) Then glidnostamp = rs.Fields("idno")
If Not IsNull(rs.Fields("memberno")) Then glmemnostamp = rs.Fields("memberno")
If Not IsNull(rs.Fields("payno")) Then glpaynostamp = rs.Fields("payno")
bookba2 = cub_balance(lblglstamp)
Else
glnamestamp1 = ""
End If


End Sub

Private Sub lvememtrans_DblClick()
    On Error GoTo SysError
    If lvememtrans.ListItems.Count > 0 Then
        transdate = lvememtrans.SelectedItem
        vno = lvememtrans.SelectedItem.SubItems(6)
'        frmGLTrans.txtTransDate = TransDate
'        frmGLTrans.txtVoucherNo = VNo
'        frmGLTrans.Show vbModal, Me
        lvwTrans.ListItems.Clear
        Dim rsTrans As New Recordset, DRTotal As Double, CRTotal As Double
        Set rsTrans = oSaccoMaster.GetRecordset("Set Dateformat dmy Select * From " _
        & "CustomerBalance Where TransDate='" & transdate & "' and VNo='" & vno & "'")
        DRTotal = 0
        CRTotal = 0
        With rsTrans
            If .State = adStateOpen Then
                While Not .EOF
                    Set li = lvwTrans.ListItems.Add(, , IIf(IsNull(!AccName), "", !AccName))
                    Select Case !transtype
                        Case "DR"
                        li.SubItems(1) = Format(IIf(IsNull(!amount), 0, !amount), "###,###,###,###,##0.00")
                        li.SubItems(2) = "0.00"
                        DRTotal = DRTotal + !amount
                        Case "CR"
                        li.SubItems(2) = Format(IIf(IsNull(!amount), 0, !amount), "###,###,###,###,##0.00")
                        li.SubItems(1) = "0.00"
                        CRTotal = CRTotal + !amount
                    End Select
                    li.SubItems(3) = IIf(IsNull(!transDescription), "", !transDescription)
                    .MoveNext
                Wend
            End If
        End With
        lblDebit = Format(DRTotal, "###,###,###,##0.00")
        lblCredit = Format(CRTotal, "###,###,###,##0.00")
        fraTrans.Visible = True
    End If
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub lvememtrans_KeyPress(KeyAscii As Integer)
If KeyAscii = 32 Then
txtidno = ""
Txtaccno = ""

TXTTGNO = ""
lvememtrans.Visible = False
End If
On Error Resume Next
Txtaccno.SetFocus
End Sub

Private Sub Picture1_Click()
    frmsearchacc.Show vbModal
    If Continue Then
        If sel <> "" Then
            LBLGLCONTRA = sel
        End If
    End If
    Exit Sub
    Dim Z, S, U
    Dim rs As Recordset
    frmsearchaccounts.Show vbModal
    Z = strName
    If Z <> "" Then
        LBLGLCONTRA = Z
        glcomm = Z
        
        End If


    'frmsearchnewacc.Show vbModal
   ' Z = strName
    Set myclass = New cdbase
    Set cn = CreateObject("adodb.connection")
    Provider = myclass.OpenCon
   cn.Open Provider, "atm", "atm"
sql = "select * from cuB where ACCNO='" & glcomm & "'"
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

End Sub

Private Sub Picture10_Click()
    On Error GoTo SysError
    frmsearchacc.Show vbModal, Me
    If Continue Then
        If sel <> "" Then
            acccr = sel
        End If
    End If
    Continue = False
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub Picture2_Click()
    frmsearchacc.Show vbModal
    If Continue Then
        If sel <> "" Then
            LBLGLCOMMISSION = sel
        End If
        sel = ""
    End If
    Continue = False
    Exit Sub
    Dim Z, S, U
    Dim rs As Recordset
    frmsearchaccounts.Show vbModal
    Z = strName
    If Z <> "" Then
        LBLGLCOMMISSION = Z
        glcommission = Z
        End If
          
    Set myclass = New cdbase
    Set cn = CreateObject("adodb.connection")
    Provider = myclass.OpenCon
   cn.Open Provider, "atm", "atm"
sql = "select * from cuB where ACCno='" & Z & "'"
Set rs = New ADODB.Recordset
rs.Open sql, cn
'Dim glnamecom As String 'FOR COMMISSION
'Dim glcommemno As String 'FOR COMMISSION
'Dim glcomidno As String 'FOR COMMISSION
'Dim glcompayno As String 'FOR COMMISSION
'Dim GLCOMMISSION As String
If Not rs.EOF Then
If Not IsNull(rs.Fields("availablebalance")) Then glcommission = rs.Fields("availablebalance")
If Not IsNull(rs.Fields("accountname")) Then glnamecom1 = rs.Fields("name")
If Not IsNull(rs.Fields("idno")) Then glcomidno = rs.Fields("idno")
If Not IsNull(rs.Fields("memberno")) Then glcommemno = rs.Fields("memberno")
If Not IsNull(rs.Fields("payno")) Then glcompayno = rs.Fields("payno")
End If
bookba1 = cub_balance(glcommission)

End Sub

Private Sub Picture3_Click()
    frmsearchacc.Show vbModal
    If Continue Then
        If sel <> "" Then
            lblglstamp = sel
        End If
        sel = ""
    End If
    Continue = False
    Exit Sub
    
    Dim Z, S, U
    Dim rs As Recordset
    frmsearchaccounts.Show vbModal
    Z = strName
    If Z <> "" Then
        lblglstamp = Z
    End If
    Set myclass = New cdbase
    Set cn = CreateObject("adodb.connection")
    Provider = myclass.OpenCon
   cn.Open Provider, "atm", "atm"
    sql = "select * from cuB where ACCno='" & lblglstamp & "'"
    Set rs = New ADODB.Recordset
    rs.Open sql, cn
    If Not rs.EOF Then
        If Not IsNull(rs.Fields("availablebalance")) Then glstampbal = rs.Fields("availablebalance")
        If Not IsNull(rs.Fields("accountname")) Then glnamestamp1 = rs.Fields("name")
        If Not IsNull(rs.Fields("idno")) Then glidnostamp = rs.Fields("idno")
        If Not IsNull(rs.Fields("memberno")) Then glmemnostamp = rs.Fields("memberno")
        If Not IsNull(rs.Fields("payno")) Then glpaynostamp = rs.Fields("payno")
    End If
    bookba2 = cub_balance(lblglstamp)
End Sub

Private Sub Picture4_Click()
    Me.MousePointer = vbHourglass
    frmsearchacc.Show vbModal
    Txtaccno = sel
    txtAccNo_Validate True
    Me.MousePointer = 0
End Sub

Private Sub Picture5_Click()
'// for charge one
Dim Z, S, U

    Dim rs As Recordset
frmsearchaccounts.Show vbModal
    Z = strName
    If Z <> "" Then
        cbocharge1glaccno = Z
        End If


    'frmsearchnewacc.Show vbModal
   ' Z = strName
   
          
    Set myclass = New cdbase
    Set cn = CreateObject("adodb.connection")
    Provider = myclass.OpenCon
   cn.Open Provider, "atm", "atm"
sql = "select * from cuB where ACCNO='" & cbocharge1glaccno & "'"
Set rs = New ADODB.Recordset

rs.Open sql, cn
If Not rs.EOF Then
If Not IsNull(rs.Fields("accno")) Then glcbocharge1accno = rs.Fields("accno")
If Not IsNull(rs.Fields("accountname")) Then glcbocharge1name = rs.Fields("name")
If Not IsNull(rs.Fields("idno")) Then glcbocharge1idno = rs.Fields("idno")
If Not IsNull(rs.Fields("memberno")) Then glcbocharge1memberno = rs.Fields("memberno")
If Not IsNull(rs.Fields("payno")) Then glcbocharge1payno = rs.Fields("payno")
End If
'glPremium = Scheme_GL_Field(AccountCode, "glPremium")
glcbocharge1boobal = cub_balance(glcbocharge1accno)


End Sub

Private Sub Picture6_Click()
Dim Z, S, U

    Dim rs As Recordset
frmsearchaccounts.Show vbModal
    Z = strName
    If Z <> "" Then
        cbocharge2accno = Z
        End If


    'frmsearchnewacc.Show vbModal
   ' Z = strName
   
          
    Set myclass = New cdbase
    Set cn = CreateObject("adodb.connection")
    Provider = myclass.OpenCon
   cn.Open Provider, "atm", "atm"
sql = "select * from cuB where ACCNO='" & cbocharge2accno & "'"
Set rs = New ADODB.Recordset

rs.Open sql, cn
If Not rs.EOF Then
If Not IsNull(rs.Fields("accno")) Then glcbocharge2accno = rs.Fields("accno")
If Not IsNull(rs.Fields("accountname")) Then glcbocharge2name = rs.Fields("name")
If Not IsNull(rs.Fields("idno")) Then glcbocharge2idno = rs.Fields("idno")
If Not IsNull(rs.Fields("memberno")) Then glcbocharge2memberno = rs.Fields("memberno")
If Not IsNull(rs.Fields("payno")) Then glcbocharge2payno = rs.Fields("payno")
End If
'glPremium = Scheme_GL_Field(AccountCode, "glPremium")
glcbocharge2boobal = cub_balance(glcbocharge2accno)



End Sub

Private Sub Picture7_Click()
Dim Z, S, U

    Dim rs As Recordset
frmsearchaccounts.Show vbModal
    Z = strName
    If Z <> "" Then
        cbocharge3accno = Z
        End If


    'frmsearchnewacc.Show vbModal
   ' Z = strName
   
          
    Set myclass = New cdbase
    Set cn = CreateObject("adodb.connection")
    Provider = myclass.OpenCon
   cn.Open Provider, "atm", "atm"
sql = "select * from cuB where ACCNO='" & cbocharge3accno & "'"
Set rs = New ADODB.Recordset

rs.Open sql, cn
If Not rs.EOF Then
If Not IsNull(rs.Fields("accno")) Then glcbocharge3accno = rs.Fields("accno")
If Not IsNull(rs.Fields("accountname")) Then glcbocharge3name = rs.Fields("name")
If Not IsNull(rs.Fields("idno")) Then glcbocharge3idno = rs.Fields("idno")
If Not IsNull(rs.Fields("memberno")) Then glcbocharge3memberno = rs.Fields("memberno")
If Not IsNull(rs.Fields("payno")) Then glcbocharge3payno = rs.Fields("payno")
End If
'glPremium = Scheme_GL_Field(AccountCode, "glPremium")
glcbocharge3boobal = cub_balance(glcbocharge3accno)



End Sub

Private Sub Picture8_Click()
Dim Z, S, U

    Dim rs As Recordset
frmsearchaccounts.Show vbModal
    Z = strName
    If Z <> "" Then
        cbocharge4accno = Z
        End If


    'frmsearchnewacc.Show vbModal
   ' Z = strName
   
          
    Set myclass = New cdbase
    Set cn = CreateObject("adodb.connection")
    Provider = myclass.OpenCon
   cn.Open Provider, "atm", "atm"
sql = "select * from cuB where ACCNO='" & cbocharge4accno & "'"
Set rs = New ADODB.Recordset

rs.Open sql, cn
If Not rs.EOF Then
If Not IsNull(rs.Fields("accno")) Then glcbocharge4accno = rs.Fields("accno")
If Not IsNull(rs.Fields("accountname")) Then glcbocharge4name = rs.Fields("name")
If Not IsNull(rs.Fields("idno")) Then glcbocharge4idno = rs.Fields("idno")
If Not IsNull(rs.Fields("memberno")) Then glcbocharge4memberno = rs.Fields("memberno")
If Not IsNull(rs.Fields("payno")) Then glcbocharge4payno = rs.Fields("payno")
End If
'glPremium = Scheme_GL_Field(AccountCode, "glPremium")
glcbocharge4boobal = cub_balance(glcbocharge4accno)



End Sub

Private Sub Picture9_Click()
    On Error GoTo SysError
    frmsearchacc.Show vbModal, Me
    If Continue Then
        If sel <> "" Then
            accdr = sel
        End If
    End If
    Continue = False
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub txtAccno_Change()
    Dim myrec1 As Object
    Dim rss As Object
    Dim amt As Long
    Dim rsCODE As Recordset
    Set myclass = New cdbase
    Set cn = CreateObject("adodb.connection")
    Provider = myclass.OpenCon
   cn.Open Provider, "atm", "atm"
    Dim rsun As Recordset
    Dim uncleared As Currency
    '// check if there exist uncleared cheques
    sql = "SELECT SUM(Amount) AS unclearedamnt FROM CustomerBalance WHERE " _
    & "AccNo = '" & Txtaccno & "'"
    Set rsun = New ADODB.Recordset
    rsun.Open sql, cn
    With rsun
        If Not .EOF Then
            lblavail = Format(IIf(IsNull(!UnClearedAmnt), 0, !UnClearedAmnt), Cfmt)
            lblbookbalance = lblavail
            lbluncleared = "0.00"
        Else
            lblavail = "0.00"
            lblbookbalance = "0.00"
            lbluncleared = "0.00"
        End If
    End With
    Set rs = oSaccoMaster.GetRecordset("Select * From GLSETUP where AccNo='" & Txtaccno & "'")
    With rs
        If Not .EOF Then
            lblaccname = IIf(IsNull(!GlAccName), "", !GlAccName)
            lblname = lblaccname
        End If
    End With
    Exit Sub
    Set rs = CreateObject("adodb.recordset")
    sql = "SELECT *  FROM CustomerBalance where accno='" & Txtaccno & _
    "' ORDER BY TRANSDATE,customerbalanceid ASC"
    rs.Open sql, cn
    rs.Close
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    Set myrec1 = CreateObject("adodb.recordset")
    sql = "SELECT top 1 * from cub where accno='" & Txtaccno & "' "
    myrec1.Open sql, cn
    If myrec1.EOF Then
    lblname = ""
    lblaccname = ""
    
    lblavail = 0#
    lvememtrans.Visible = False
    'MsgBox "Check if Member  Exist OR Check if the account is valid?? ", vbInformation, "Transactional details"
     Exit Sub
    Else
    lvememtrans.Visible = True
    If Not IsNull(myrec1!name) Then lblname = myrec1!name
    If Not IsNull(myrec1!AccountName) Then lblaccname = myrec1!AccountName
    Set rsCODE = CreateObject("ADODB.Recordset")
    rsCODE.Open "SELECT * from AccountCodes WHERE AccountName='" & lblaccname & "'", cn
    If rsCODE.EOF Then
    ''MsgBox "Try eiditing the accounttype. The account type you have does not exist in our records ", vbCritical, "Transactions"
    'Exit Sub
    Else
    lblacname = rsCODE!AccountName
    minBal = rsCODE!Minimumbal
    End If
    'rebuild_accno Txtaccno
    Dim rsproc As Recordset
    Set rsproc = New ADODB.Recordset
    Dim sssql As String
    sssql = "proc_rebuild '" & Txtaccno & "'"
    'If Not IsNull(myrec1!accNo) Then lblaccno = myrec1!accNo
    If Not IsNull(myrec1!availablebalance) Then lblavail = Format(myrec1!availablebalance, "#,###,###.00") Else lblavail = 0#
    If Not IsNull(myrec1!availablebalance) Then lblbookbalance = Format(myrec1!availablebalance - minBal, "#,###,###.00") Else lblbookbalance = 0#
    
    If lbluncleared = "" Then lbluncleared = 0
    lblavail = CCur(lblavail) + CCur(lbluncleared)
    lblavail = Format(lblavail, "###,###,###.00")
    'If Not IsNull(myrec1!memberno) Then lblgno = myrec1!memberno
    End If
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    Dim LV As ListItem
    lvememtrans.ListItems.Clear
    rs.Open
    Do While Not rs.EOF
    With lvememtrans
    If rs!transdate <> "" Then
    Set LV = .ListItems.Add(, , rs!transdate)
    If rs!transDescription <> "" Then
    LV.ListSubItems.Add , , rs!transDescription
    Else
    LV.ListSubItems.Add , , "No Desc"
    End If
    If rs!amount <> "" Then
    If UCase(Trim(rs!transtype)) = "DR" Then
    LV.ListSubItems.Add , , Format(rs!amount, "###,###,###.00")
    LV.ListSubItems.Add , , Format(0, "0.00")
    ' lvememtrans.ListItems.Add , , RS!amount = lvwColumnRight
    Else
    ' lvememtrans.ListItems.item(3).Left
    LV.ListSubItems.Add , , Format(0, "0.00")
    LV.ListSubItems.Add , , Format(rs!amount, "###,###,###.00")
    ' lvememtrans.ListItems.Add , , RS!amount = lvwColumnRight
    End If
    Else
    rs!amount = 0
    End If
    If Not IsNull(rs!availablebalance) Then
    LV.ListSubItems.Add , , Format(rs!availablebalance, "###,###,###.00")
    Else
    LV.ListSubItems.Add , , "0.00"
    End If
    If Not IsNull(rs!Commission) Then
    LV.ListSubItems.Add , , rs!Commission
    Else
    LV.ListSubItems.Add , , "0.00"
    End If
    If Not IsNull(rs!vno) Then
    LV.ListSubItems.Add , , rs!vno
    Else
    LV.ListSubItems.Add , , "DNN"
    End If
    LV.ListSubItems.Item(3).Bold = True
    
    End If
    End With
    rs.MoveNext
    Loop
    rs.Filter = 0
    rs.Close
End Sub

Private Sub txtAccNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 32 Then

Txtaccno = ""

lvememtrans.Visible = False
End If
End Sub

Private Sub txtAccNo_Validate(Cancel As Boolean)
Dim myrec1 As Object
Dim rss As Object
Dim amt As Currency
Set myclass = New cdbase
Set cn = CreateObject("adodb.connection")
Provider = myclass.OpenCon
cn.Open Provider, "atm", "atm"
Dim MYRE As Recordset
    Set MYRE = CreateObject("adodb.recordset")
    sql = "SELECT top 1 * from cub where accno='" & Txtaccno & "' "
     MYRE.Open sql, cn
     If Txtaccno <> "" Then
     If MYRE.EOF Then
      MsgBox "The account does not exist Please Seek assistance from the customer services", vbInformation, "Transactions"
     Exit Sub
     End If
     End If
End Sub

Private Sub txtIDNo_Change()
Dim myrec1 As Object
Dim rss As Object
Dim amt As Currency
Set myclass = New cdbase
Set cn = CreateObject("adodb.connection")
Provider = myclass.OpenCon
cn.Open Provider, "atm", "atm"

 Set rs = CreateObject("adodb.recordset")
    rs.Open "SELECT *  FROM CustomerBalance where idno='" & txtidno & _
    "' ORDER BY TRANSDATE , customerbalanceid ASC", cn
    rs.Close

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Set myrec1 = CreateObject("adodb.recordset")
    sql = "SELECT top 1 * from cub where idno='" & txtidno & "' "
     myrec1.Open sql, cn
     If myrec1.EOF Then
     lblname = ""
     lblaccname = ""
   

     lblavail = 0#
     lvememtrans.Visible = False
     'MsgBox "Check if Member  Exist OR Check if the account is valid?? ", vbInformation, "Transactional details"
     Exit Sub
     Else
        lvememtrans.Visible = True
        If Not IsNull(myrec1!name) Then lblname = myrec1!name
        If Not IsNull(myrec1!AccountName) Then lblaccname = myrec1!AccountName
        'If Not IsNull(myrec1!accNo) Then lblaccno = myrec1!accNo
        If Not IsNull(myrec1!availablebalance) Then lblavail = Format(myrec1!availablebalance + 500, "#,###,###.0#") Else lblavail = 0#
        If Not IsNull(myrec1!availablebalance) Then lblbookbalance = Format(myrec1!availablebalance, "#,###,###.##") Else lblbookbalance = 0#
  
       ' If Not IsNull(myrec1!memberno) Then lblgno = myrec1!memberno
     End If
     
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++


    Dim LV As ListItem

    lvememtrans.ListItems.Clear
    rs.Open

    Do While Not rs.EOF

   With lvememtrans
      If rs!transdate <> "" Then
      Set LV = .ListItems.Add(, , rs!transdate)
                 If Not IsNull(rs!vno) Then
                  LV.ListSubItems.Add , , rs!vno
                Else
                  LV.ListSubItems.Add , , "DNN"
                End If
        If rs!amount <> "" Then
          If UCase(rs!transtype) = "DR" Then
            LV.ListSubItems.Add , , Format(rs!amount, "###,###,###.##")
            LV.ListSubItems.Add , , Format(0, "0.00")
          Else
            LV.ListSubItems.Add , , Format(0, "0.00")
            LV.ListSubItems.Add , , Format(rs!amount, "###,###,###.##")
          End If
        Else
           rs!amount = 0
        End If
        
        If Not IsNull(rs!availablebalance) Then
             LV.ListSubItems.Add , , rs!availablebalance
        Else
             LV.ListSubItems.Add , , "0.00"
        End If
        If Not IsNull(rs!Commission) Then
            LV.ListSubItems.Add , , rs!Commission
        Else
           LV.ListSubItems.Add , , "0.00"
        End If
         If rs!transDescription <> "" Then
              LV.ListSubItems.Add , , rs!transDescription
         Else
              LV.ListSubItems.Add , , "No Desc"
         End If
      LV.ListSubItems.Item(3).Bold = True
      
      End If
    End With



    rs.MoveNext
    Loop

    rs.Filter = 0
    rs.Close
    


End Sub

Private Sub txtidno_KeyPress(KeyAscii As Integer)
If KeyAscii = 32 Then
txtidno = ""
lvememtrans.Visible = False
End If
End Sub

Private Sub txtpayno_Change()
Dim myrec1 As Object
Dim rss As Object
Dim amt As Long
Set myclass = New cdbase
Set cn = CreateObject("adodb.connection")
Provider = myclass.OpenCon
cn.Open Provider, "atm", "atm"

 Set rs = CreateObject("adodb.recordset")
    rs.Open "SELECT *  FROM CustomerBalance where payrollno='" & txtPayNo & _
    "' ORDER BY TRANSDATE , customerbalanceid ASC", cn
    rs.Close

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Set myrec1 = CreateObject("adodb.recordset")
    sql = "SELECT top 1 * from cub where payno='" & txtPayNo & "' "
     myrec1.Open sql, cn
     If myrec1.EOF Then
     lblname = ""
     lblaccname = ""
   

     lblavail = 0#
     lvememtrans.Visible = False
     'MsgBox "Check if Member  Exist OR Check if the account is valid?? ", vbInformation, "Transactional details"
     Exit Sub
     Else
        lvememtrans.Visible = True
        If Not IsNull(myrec1!name) Then lblname = myrec1!name
        If Not IsNull(myrec1!AccountName) Then lblaccname = myrec1!AccountName
        'If Not IsNull(myrec1!accNo) Then lblaccno = myrec1!accNo
        If Not IsNull(myrec1!availablebalance) Then lblavail = Format(myrec1!availablebalance + 500, "#,###,###.0#") Else lblavail = 0#
        If Not IsNull(myrec1!availablebalance) Then lblbookbalance = Format(myrec1!availablebalance, "#,###,###.##")
  
       ' If Not IsNull(myrec1!memberno) Then lblgno = myrec1!memberno
     End If
     
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++


    Dim LV As ListItem

    lvememtrans.ListItems.Clear
    rs.Open

    Do While Not rs.EOF

    With lvememtrans
      If rs!transdate <> "" Then
      Set LV = .ListItems.Add(, , rs!transdate)
                 If Not IsNull(rs!vno) Then
                  LV.ListSubItems.Add , , rs!vno
                Else
                  LV.ListSubItems.Add , , "DNN"
                End If
        If rs!amount <> "" Then
          If UCase(rs!transtype) = "DR" Then
            LV.ListSubItems.Add , , Format(rs!amount, "###,###,###.##")
            LV.ListSubItems.Add , , Format(0, "0.00")
          Else
            LV.ListSubItems.Add , , Format(0, "0.00")
            LV.ListSubItems.Add , , Format(rs!amount, "###,###,###.##")
          End If
        Else
           rs!amount = 0
        End If
        
        If Not IsNull(rs!availablebalance) Then
             LV.ListSubItems.Add , , rs!availablebalance
        Else
             LV.ListSubItems.Add , , "0.00"
        End If
        If Not IsNull(rs!Commission) Then
            LV.ListSubItems.Add , , rs!Commission
        Else
           LV.ListSubItems.Add , , "0.00"
        End If
         If rs!transDescription <> "" Then
              LV.ListSubItems.Add , , rs!transDescription
         Else
              LV.ListSubItems.Add , , "No Desc"
         End If
      LV.ListSubItems.Item(3).Bold = True
      
      End If
    End With


    rs.MoveNext
    Loop

    rs.Filter = 0
    rs.Close
    

End Sub

Private Sub txtpayno_KeyPress(KeyAscii As Integer)
If KeyAscii = 32 Then
 txtPayNo = ""
lvememtrans.Visible = False
End If
End Sub

Private Sub txttgno_Change()
    Dim myrec2 As Object
    Dim rss As Object
    Dim amt As Long
    Set myclass = New cdbase
    Set cn = CreateObject("adodb.connection")
    Provider = myclass.OpenCon
   cn.Open Provider, "atm", "atm"
    Dim rsun As Recordset
    Dim uncleared As Currency
    '// check if there exist uncleared cheques
    sql = "Select SUM(Amount) AS UnclearedAmnt FROM CustomerBalance where AccNo= '" _
    & Txtaccno & "'"
    Set rsun = New ADODB.Recordset
    rsun.Open sql, cn
    With rsun
        If Not .EOF Then
            lbluncleared = IIf(IsNull(!UnClearedAmnt), 0, !UnClearedAmnt)
        Else
            lbluncleared = 0
        End If
    End With
    sql = ""
    Set rs = CreateObject("adodb.recordset")
    sql = "Select * From CUSTOMERBALANCE where AccNo='" & TXTTGNO & "' " _
    & "Order BY TransDate,customerbalanceid ASC"
    rs.Open sql, cn
    rs.Close
    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXX              XXXXXXXXXXXXXXXXXXXXXXXXXXXXXX'
    Set myrec2 = CreateObject("adodb.recordset")
    sql = "SELECT * from cub where memberno='" & TXTTGNO & "' "
    myrec2.Open sql, cn
    If myrec2.EOF Then
        lblname = ""
        lblaccname = ""
        lblavail = 0#
        lvememtrans.Visible = False
    Else
        lvememtrans.Visible = True
        If Not IsNull(myrec2!name) Then lblname = myrec2!name
        If Not IsNull(myrec2!AccountName) Then lblaccname = myrec2!AccountName
        ' If Not IsNull(myrec2!accNo) Then lblaccno = myrec2!accNo
        If Not IsNull(myrec2!availablebalance) Then lblavail = myrec2!availablebalance + 500 Else lblavail = 0#
        If Not IsNull(myrec2!availablebalance) Then lblbookbalance = Format(myrec2!availablebalance, "#,###,###.##") Else lblbookbalance = 0#
    End If
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    Dim LV As ListItem
    lvememtrans.ListItems.Clear
    rs.Open
    Do While Not rs.EOF
    With lvememtrans
    If rs!transdate <> "" Then
    Set LV = .ListItems.Add(, , rs!transdate)
    If Not IsNull(rs!vno) Then
    LV.ListSubItems.Add , , rs!vno
    Else
    LV.ListSubItems.Add , , "DNN"
    End If
    If rs!amount <> "" Then
    If UCase(rs!transtype) = "DR" Then
    LV.ListSubItems.Add , , Format(rs!amount, "###,###,###.##")
    LV.ListSubItems.Add , , Format(0, "0.00")
    Else
    LV.ListSubItems.Add , , Format(0, "0.00")
    LV.ListSubItems.Add , , Format(rs!amount, "###,###,###.##")
    End If
    Else
    rs!amount = 0
    End If
    
    If Not IsNull(rs!availablebalance) Then
    LV.ListSubItems.Add , , rs!availablebalance
    Else
    LV.ListSubItems.Add , , "0.00"
    End If
    If Not IsNull(rs!Commission) Then
    LV.ListSubItems.Add , , rs!Commission
    Else
    LV.ListSubItems.Add , , "0.00"
    End If
    If rs!transDescription <> "" Then
    LV.ListSubItems.Add , , rs!transDescription
    Else
    LV.ListSubItems.Add , , "No Desc"
    End If
    LV.ListSubItems.Item(3).Bold = True
    End If
    End With
    rs.MoveNext
    Loop
    rs.Filter = 0
    rs.Close
End Sub

Private Sub txttgno_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
        TXTTGNO = ""
        lvememtrans.Visible = False
    End If
End Sub

Private Sub Form_Load()
    'Lbldate = Date
    On Error Resume Next
    DTP.value = Date
    DTPadj = Date
    txtamount = 0#
    frasalk.Move -360, -90, 11415, 7455
    With Cbodetail
        If .ListCount > 0 Then
            .ListIndex = 0
        End If
    End With
    With lvememtrans
        .ColumnHeaders.Clear
        .ListItems.Clear
    End With
    With lvememtrans
        .ColumnHeaders.Clear
        .ColumnHeaders.Add 1, , "Transaction Date"
        .ColumnHeaders.Add 2, , "Description", 2000
        .ColumnHeaders.Add 3, , "Dr", , lvwColumnRight
        .ColumnHeaders.Add 4, , "Cr", , lvwColumnRight
        .ColumnHeaders.Add 5, , "Available Balance", , lvwColumnRight
        .ColumnHeaders.Add 6, , "Commissions"
        .ColumnHeaders.Add 7, , "Voucher Number"
    End With
    optcredit = False
    Optdebit = False
    lblname = ""
    lblaccname = ""
    lblavail = ""
    lblname = ""
    TXTTGNO = ""
    txtadjamnt = 0
    txtamount = 0
    Txtvou = ""
    txtvoucherno = ""
    Set myclass = New cdbase
    Provider = myclass.OpenCon
    Set cn = CreateObject("adodb.connection")
   cn.Open Provider, "atm", "atm"
    Dim rstcode As Recordset
    Set rstcode = CreateObject("ADODB.Recordset")
    sql = ""
    sql = "SELECT * from TransCode"
    rstcode.Open sql, cn
    While Not rstcode.EOF
        cbomemtrans.AddItem rstcode.Fields("description")
        rstcode.MoveNext
    Wend
    myclass.CloseCon
    Set myclass = Nothing
    txtvoucherno.SetFocus
    OpCR = True
    frasalk.Visible = True
    txtpassword = ""
     txtpassword.SetFocus
End Sub

