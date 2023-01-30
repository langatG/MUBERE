VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmNominals 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NOMINAL RECEIPTS/PAYMENTS"
   ClientHeight    =   7710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9975
   Icon            =   "frmNominals.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   9975
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   6735
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   11880
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "RECEIPTS"
      TabPicture(0)   =   "frmNominals.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "PAYMENTS"
      TabPicture(1)   =   "frmNominals.frx":045E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame2(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame2 
         Height          =   6240
         Index           =   1
         Left            =   120
         TabIndex        =   31
         Top             =   480
         Width           =   9510
         Begin VB.ComboBox CboParticulars 
            Height          =   315
            Index           =   1
            Left            =   1680
            TabIndex        =   66
            Top             =   1800
            Width           =   3375
         End
         Begin VB.ComboBox Cbo1 
            Height          =   315
            Index           =   1
            Left            =   7440
            TabIndex        =   64
            Top             =   1560
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.CommandButton cmdupdatereceipt 
            Caption         =   "&Post"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   8160
            TabIndex        =   60
            Top             =   3240
            Width           =   1185
         End
         Begin VB.CommandButton cmdClear 
            Caption         =   "<< All"
            Height          =   345
            Index           =   1
            Left            =   6360
            TabIndex        =   57
            Top             =   3240
            Width           =   900
         End
         Begin VB.CommandButton cmdRemove 
            Caption         =   "<<Remove"
            Height          =   345
            Index           =   1
            Left            =   5520
            TabIndex        =   56
            Top             =   3240
            Width           =   900
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Add>>"
            Height          =   345
            Index           =   1
            Left            =   4680
            TabIndex        =   55
            Top             =   3240
            Width           =   900
         End
         Begin VB.CommandButton cmdAcctsSearch 
            Height          =   300
            Index           =   1
            Left            =   1485
            Picture         =   "frmNominals.frx":047A
            Style           =   1  'Graphical
            TabIndex        =   45
            Top             =   2730
            Width           =   330
         End
         Begin VB.TextBox txtAccNames 
            Height          =   315
            Index           =   1
            Left            =   1815
            TabIndex        =   44
            Top             =   2730
            Width           =   3225
         End
         Begin VB.ComboBox cboAccno 
            Height          =   315
            Index           =   1
            Left            =   315
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Top             =   2730
            Width           =   1200
         End
         Begin VB.ComboBox cboBanks 
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
            Index           =   1
            ItemData        =   "frmNominals.frx":057C
            Left            =   2100
            List            =   "frmNominals.frx":057E
            TabIndex        =   42
            Top             =   720
            Width           =   1350
         End
         Begin VB.CommandButton cmdBank 
            Caption         =   "<>"
            Height          =   300
            Index           =   1
            Left            =   3465
            TabIndex        =   41
            Top             =   765
            Width           =   345
         End
         Begin VB.TextBox txtmode 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   1
            Left            =   3180
            TabIndex        =   40
            Top             =   1410
            Width           =   1380
         End
         Begin VB.CommandButton cmdVoucher 
            Caption         =   "<>"
            Height          =   300
            Index           =   1
            Left            =   8835
            TabIndex        =   39
            Top             =   1200
            Width           =   345
         End
         Begin VB.ComboBox cboMode 
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
            Index           =   1
            ItemData        =   "frmNominals.frx":0580
            Left            =   1680
            List            =   "frmNominals.frx":0593
            TabIndex        =   38
            Text            =   "Cash"
            Top             =   1395
            Width           =   1425
         End
         Begin VB.TextBox txtAmountPaid 
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
            Index           =   1
            Left            =   5160
            TabIndex        =   37
            Text            =   "0"
            Top             =   2400
            Width           =   1740
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
            Height          =   285
            Index           =   1
            Left            =   7440
            TabIndex        =   36
            Top             =   1200
            Width           =   1380
         End
         Begin VB.TextBox txtBalance 
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
            Index           =   1
            Left            =   7440
            TabIndex        =   35
            Text            =   "0"
            Top             =   2775
            Width           =   1860
         End
         Begin VB.TextBox txtDistributed 
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
            Index           =   1
            Left            =   7200
            TabIndex        =   34
            Text            =   "0"
            Top             =   2280
            Width           =   2100
         End
         Begin VB.TextBox txtParticulars 
            Appearance      =   0  'Flat
            Height          =   375
            Index           =   1
            Left            =   1680
            TabIndex        =   33
            Top             =   1920
            Visible         =   0   'False
            Width           =   3225
         End
         Begin VB.TextBox txtPayee 
            Appearance      =   0  'Flat
            Height          =   375
            Index           =   1
            Left            =   1680
            TabIndex        =   32
            Top             =   2160
            Width           =   3225
         End
         Begin MSComctlLib.ListView lvwNtrans 
            Height          =   2340
            Index           =   1
            Left            =   270
            TabIndex        =   46
            Top             =   3795
            Width           =   8940
            _ExtentX        =   15769
            _ExtentY        =   4128
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
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Code"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Trans Description"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Text            =   "Amount"
               Object.Width           =   3175
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Accno"
               Object.Width           =   2540
            EndProperty
         End
         Begin MSComCtl2.DTPicker DTPDatedeposited 
            Height          =   300
            Index           =   1
            Left            =   6210
            TabIndex        =   58
            Top             =   240
            Width           =   1470
            _ExtentX        =   2593
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
            Format          =   131268611
            CurrentDate     =   40463
         End
         Begin VB.Label Label3 
            Caption         =   "Ref Doc"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6360
            TabIndex        =   63
            Top             =   1560
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Date Deposited"
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
            Index           =   1
            Left            =   4920
            TabIndex        =   59
            Top             =   285
            Width           =   1245
         End
         Begin VB.Label lblbankname 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Index           =   1
            Left            =   3810
            TabIndex        =   54
            Top             =   765
            Width           =   4095
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Source(CR)"
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
            Left            =   450
            TabIndex        =   53
            Top             =   810
            Width           =   930
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Payment Mode"
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
            Index           =   1
            Left            =   420
            TabIndex        =   52
            Top             =   1440
            Width           =   1230
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Received Amount"
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
            Index           =   1
            Left            =   5400
            TabIndex        =   51
            Top             =   2085
            Width           =   1455
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Doc  No"
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
            Index           =   1
            Left            =   6405
            TabIndex        =   50
            Top             =   1200
            Width           =   600
         End
         Begin VB.Label Label7 
            Caption         =   "Distributed"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   7920
            TabIndex        =   49
            Top             =   2085
            Width           =   1005
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Particulars"
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
            Left            =   375
            TabIndex        =   48
            Top             =   1875
            Width           =   885
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Payee"
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
            Index           =   1
            Left            =   465
            TabIndex        =   47
            Top             =   2400
            Width           =   495
         End
      End
      Begin VB.Frame Frame2 
         Height          =   6240
         Index           =   0
         Left            =   -74880
         TabIndex        =   1
         Top             =   360
         Width           =   9510
         Begin VB.ComboBox CboParticulars 
            Height          =   315
            Index           =   0
            Left            =   1680
            TabIndex        =   65
            Top             =   1800
            Width           =   3135
         End
         Begin VB.ComboBox Cbo1 
            Height          =   315
            Index           =   0
            Left            =   6960
            TabIndex        =   62
            Top             =   1560
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.TextBox txtPayee 
            Appearance      =   0  'Flat
            Height          =   375
            Index           =   0
            Left            =   1680
            TabIndex        =   19
            Top             =   2160
            Width           =   3225
         End
         Begin VB.TextBox txtParticulars 
            Appearance      =   0  'Flat
            Height          =   375
            Index           =   0
            Left            =   4320
            TabIndex        =   18
            Top             =   2040
            Visible         =   0   'False
            Width           =   3225
         End
         Begin VB.TextBox txtDistributed 
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
            Index           =   0
            Left            =   7080
            Locked          =   -1  'True
            MaxLength       =   9
            TabIndex        =   17
            Text            =   "0"
            Top             =   2400
            Width           =   2100
         End
         Begin VB.TextBox txtBalance 
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
            Index           =   0
            Left            =   7200
            Locked          =   -1  'True
            TabIndex        =   16
            Text            =   "0"
            Top             =   2775
            Width           =   2100
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
            Height          =   285
            Index           =   0
            Left            =   6960
            TabIndex        =   15
            Top             =   1200
            Width           =   1500
         End
         Begin VB.TextBox txtAmountPaid 
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
            Index           =   0
            Left            =   5160
            MaxLength       =   9
            TabIndex        =   14
            Text            =   "0"
            Top             =   2400
            Width           =   1860
         End
         Begin VB.ComboBox cboMode 
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
            Index           =   0
            ItemData        =   "frmNominals.frx":05B6
            Left            =   1680
            List            =   "frmNominals.frx":05C9
            TabIndex        =   13
            Text            =   "Cash"
            Top             =   1395
            Width           =   1425
         End
         Begin VB.CommandButton cmdReceipt 
            Caption         =   "<>"
            Height          =   300
            Index           =   0
            Left            =   8880
            TabIndex        =   12
            Top             =   1200
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.TextBox txtmode 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   0
            Left            =   3180
            TabIndex        =   11
            Top             =   1410
            Width           =   1380
         End
         Begin VB.CommandButton cmdupdatereceipt 
            Caption         =   "&Post!"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   7920
            TabIndex        =   10
            Top             =   3240
            Width           =   1185
         End
         Begin VB.CommandButton cmdBank 
            Caption         =   "<>"
            Height          =   300
            Index           =   0
            Left            =   3465
            TabIndex        =   9
            Top             =   765
            Width           =   345
         End
         Begin VB.ComboBox cboBanks 
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
            Index           =   0
            ItemData        =   "frmNominals.frx":05EC
            Left            =   2100
            List            =   "frmNominals.frx":05EE
            TabIndex        =   8
            Top             =   750
            Width           =   1350
         End
         Begin VB.ComboBox cboAccno 
            Height          =   315
            Index           =   0
            Left            =   315
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   2730
            Width           =   1200
         End
         Begin VB.TextBox txtAccNames 
            Height          =   315
            Index           =   0
            Left            =   1815
            TabIndex        =   6
            Top             =   2730
            Width           =   3225
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Add>>"
            Height          =   345
            Index           =   0
            Left            =   4680
            TabIndex        =   5
            Top             =   3240
            Width           =   900
         End
         Begin VB.CommandButton cmdRemove 
            Caption         =   "<<Remove"
            Height          =   345
            Index           =   0
            Left            =   5520
            TabIndex        =   4
            Top             =   3240
            Width           =   900
         End
         Begin VB.CommandButton cmdAcctsSearch 
            Height          =   300
            Index           =   0
            Left            =   1485
            Picture         =   "frmNominals.frx":05F0
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   2730
            Width           =   330
         End
         Begin VB.CommandButton cmdClear 
            Caption         =   "<< All"
            Height          =   345
            Index           =   0
            Left            =   6360
            TabIndex        =   2
            Top             =   3240
            Width           =   900
         End
         Begin MSComctlLib.ListView lvwNtrans 
            Height          =   2340
            Index           =   0
            Left            =   270
            TabIndex        =   20
            Top             =   3795
            Width           =   8940
            _ExtentX        =   15769
            _ExtentY        =   4128
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
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Code"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Trans Description"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Text            =   "Amount"
               Object.Width           =   3175
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Accno"
               Object.Width           =   2540
            EndProperty
         End
         Begin MSComCtl2.DTPicker DTPDatedeposited 
            Height          =   300
            Index           =   0
            Left            =   6210
            TabIndex        =   21
            Top             =   240
            Width           =   1470
            _ExtentX        =   2593
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
            Format          =   131203075
            CurrentDate     =   40463
         End
         Begin VB.Label Label1 
            Caption         =   "Ref Doc"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6120
            TabIndex        =   61
            Top             =   1560
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Payee"
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
            Index           =   0
            Left            =   465
            TabIndex        =   30
            Top             =   2280
            Width           =   495
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Particulars"
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
            Index           =   0
            Left            =   375
            TabIndex        =   29
            Top             =   1875
            Width           =   885
         End
         Begin VB.Label Label11 
            Caption         =   "Distributed"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   0
            Left            =   8085
            TabIndex        =   28
            Top             =   2085
            Width           =   1005
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Doc No"
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
            Index           =   0
            Left            =   6120
            TabIndex        =   27
            Top             =   1200
            Width           =   555
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Received Amount"
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
            Left            =   5280
            TabIndex        =   26
            Top             =   2085
            Width           =   1455
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Payment Mode"
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
            Left            =   420
            TabIndex        =   25
            Top             =   1440
            Width           =   1230
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Source(DR)"
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
            Left            =   450
            TabIndex        =   24
            Top             =   810
            Width           =   915
         End
         Begin VB.Label lblbankname 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Index           =   0
            Left            =   3810
            TabIndex        =   23
            Top             =   765
            Width           =   4095
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Date Deposited"
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
            Index           =   0
            Left            =   4920
            TabIndex        =   22
            Top             =   285
            Width           =   1245
         End
      End
   End
   Begin MSComDlg.CommonDialog dlg9 
      Left            =   4680
      Top             =   7080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmNominals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim totalamount As Currency
Dim pushed As Currency
Dim objLabelEdit As LabelEdit
Dim objLabelEdit2 As LabelEdit
Dim objLabelEdit3 As LabelEdit
Dim interestAcc As String, LoanAcc As String
Dim IsBatch As Boolean, IsGroup As Boolean, isfixed As Boolean
Dim k As Integer
Dim shareBal As Double, Balance As Double
Dim daysIntoTheMonth As Integer
Dim Posted As Boolean
Dim ref As String, RefNo As String
Dim penalise As Boolean
Dim commit As Boolean
Dim rsDefaults As ADODB.Recordset
Dim fSeed As Integer
Dim newMember As Boolean
Private Sub cboAccno_Change(index As Integer)
    Dim ACCNO As String
    ACCNO = cboAccno(index).Text
    sql = "select GLACCNAME from glsetup where accno='" & ACCNO & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
        txtAccNames(index).Text = rs(0)
    Else
        txtAccNames(index).Text = ""
    End If
End Sub

Private Sub cboAccno_Click(index As Integer)
    cboAccno_Change (index)
End Sub



Private Sub cboBanks_Change(index As Integer)
    On Error GoTo SysError
      Dim rsGlaccname As New ADODB.Recordset
'
'        Set rs = oSaccoMaster.GetRecordSet(" select   AssignGl from  UserAccounts where UserLoginID='" & User & "'  and AssignGl<>''  ")
'              If rs.EOF Then
'                MsgBox "Only teller can use this", vbInformation
'              Exit Sub
'              Else
'              End If

  Dim ACCNO As String
    ACCNO = cboBanks(index).Text
    sql = "select Glaccname from glsetup where accno='" & ACCNO & "'"
    Set rsGlaccname = oSaccoMaster.GetRecordset(sql)
    If Not rsGlaccname.EOF Then
        lblbankname(index).Caption = rsGlaccname("GLACCNAME")
    Else
       lblbankname(index).Caption = ""
    End If
    Exit Sub

'           ' End If
'    sql = "select b.bankname from banks b inner join glsetup g on b.Accno=g.Accno where b.Accno='" & cboBanks(Index).Text & "'AND  b.Accno <>'' "
'    Set rst = oSaccoMaster.GetRecordSet(sql)
'    If Not rst.EOF Then
'        lblbankname(Index).Caption = rst(0)
'    Else
'        lblbankname(Index).Caption = ""
'    End If
'    Exit Sub
'
  
       
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub cboBanks_Click(index As Integer)
    cboBanks_Change (index)
End Sub




Private Sub cboMode_Change(index As Integer)
'On Error Resume Next
    Select Case cboMode(index).Text
        Case "Cash"
            txtmode(index).Visible = False
            txtReceiptsno(index).SetFocus
        Case "Cheque"
            txtmode(index).Visible = True
            txtmode(index).SetFocus
        Case "Direct Deposit"
            txtmode(index).Visible = True
            txtmode(index).SetFocus
        Case Else

    End Select
End Sub

Private Sub cboMode_Click(index As Integer)
    cboMode_Change (index)
End Sub

Private Sub CboParticulars_Change(index As Integer)
   Dim rs333  As New ADODB.Recordset
   Set rs333 = oSaccoMaster.GetRecordset("select GL from TransCodebosa where  GL  in ( select accno from Glsetup) and Description='" & CboParticulars(index) & "'")
       If Not rs333.EOF Then
       cboAccno(index) = IIf(IsNull(rs333("GL")), "", rs333("GL"))
      Else
       'cboAccno(Index) = ""
       End If
End Sub

Private Sub CboParticulars_Click(index As Integer)
CboParticulars_Change (index)
End Sub

Private Sub cmdAcctsSearch_Click(index As Integer)
    On Error Resume Next
    frmAcctsSearch.Show vbModal, Me
    If Continue Then
        If SearchValue <> "" Then
            cboAccno(index) = SearchValue
            SearchValue = ""
            Continue = False
        End If
    End If
End Sub

Private Sub cmdadd_Click(index As Integer)
    On Error GoTo SysError
    If cboAccno(index).Text = "" Then
        Exit Sub
    End If

    Set li = lvwNtrans(index).ListItems.Add(, , cboAccno(index))
    li.SubItems(1) = txtAccNames(index)
    li.SubItems(2) = "0.00"
   
    Recalculate (index)
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub cmdBank_Click(index As Integer)

 On Error Resume Next
    frmAcctsSearch.Show vbModal, Me
    If Continue Then
        If SearchValue <> "" Then
            cboBanks(index) = SearchValue
            SearchValue = ""
            Continue = False
        End If
    End If
    
' On Error Resume Next
'    frmAcctsSearch.Show vbModal, Me
'    If Continue Then
'        If SearchValue <> "" Then
'            cboAccno(index) = SearchValue
'            SearchValue = ""
'            Continue = False
'        End If
'    End If
'
'    On Error GoTo SysError
'    frmsearchBanks.Show vbModal, Me
'    If Continue Then
'        If SearchValue <> "" Then
'            cboBanks(index) = SearchValue
'            SearchValue = ""
'        End If
'    End If
'    Exit Sub
'SysError:
'    MsgBox Err.Description, vbInformation, Me.Caption
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub



Private Sub cmdRemove_Click(index As Integer)
    On Error GoTo SysError
    With lvwNtrans(index)
        If .ListItems.Count > 0 Then
            If MsgBox("Do you want to remove " & lvwNtrans(index).SelectedItem.SubItems(1) & _
            " From the list?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
                Exit Sub
            End If
            pushed = pushed - .SelectedItem.ListSubItems(2)
            .ListItems.Remove (.SelectedItem.index)
        End If
    End With
    Recalculate (index)
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub



Private Sub cmdupdatereceipt_Click(index As Integer)
    On Error GoTo SysError
    Dim j As Integer
    Dim amt As Double, interest As Double
    Dim code As String, rsLoan As New Recordset
    Dim k As Long, repaymethod As String
    Dim sharebals As Double, amount As Double
    Dim ptype As String, principal As Double, chequeno As String
    Dim intWaive As Boolean
    Dim GlAccMainGroup As String
     Dim Budget As Double
     
     oSaccoMaster.Execute " TRUNCATE TABLE  PrintTemp "
    
    If lvwNtrans(index).ListItems.Count = 0 Then
        Exit Sub
    End If
    
    If txtReceiptsno(index) = "" Then
        MsgBox "Please enter the receiptno", vbCritical
        Exit Sub
    End If
    If txtAmountPaid(index) <= 0 Then
        MsgBox "Amount should be greater than zero", vbCritical
        Exit Sub
    End If
    If txtAmountPaid(index) = "" Then
        MsgBox "Amount should be have a figure on it", vbCritical
        Exit Sub
    End If
    If cboBanks(index).Text = "" Then
        MsgBox ("You Must Select the Bank Control Account"), vbCritical
        Exit Sub
    End If
    If cboMode(index) = "Cheque" Then
        If txtmode(index) = "" Then
            MsgBox ("Cheque Receipt No Required"), vbInformation
            Exit Sub
        End If
    End If
    If cboMode(index) = "Cash" Then
        If txtReceiptsno(index) = "" Then
            MsgBox ("Cash Receipt No Required"), vbInformation
            Exit Sub
        End If
    End If
    If cboMode(index) = "EFT" Then
        If txtmode(index) = "" Then
            MsgBox ("EFT Receipt No Required"), vbInformation
            Exit Sub
        End If
    End If
    If cboMode(index) = "Mpesa" Then
        If txtmode(index) = "" Then
            MsgBox ("Mpesa Receipt No Required"), vbInformation
            Exit Sub
        End If
    End If
    
     
     
    If cboMode(index) = "Zap" Then
        If txtmode(index) = "" Then
            MsgBox ("Zap Receipt No Required"), vbInformation
            Exit Sub
        End If
    End If
    If CDbl(txtBalance(index)) <> 0 Then
        MsgBox "The Amount Received should be equal to the Amount Distributed", vbInformation, Me.Caption
        Exit Sub
    End If
    
    If lblbankname(index).Caption = "" Then
      MsgBox "Please select or put  Correct account  ", vbInformation
      Exit Sub
    End If
    If MsgBox("Do you want to post this document: " & txtReceiptsno(index) & "?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
        Exit Sub
    End If
'    GlAccMainGroup = getGlAccMainGroup(cboAccno(index))
'       If GlAccMainGroup = "EXPENSES" Then
'          Budget = getbudget(cboAccno(index), month(DTPDatedeposited(index)), Year(DTPDatedeposited(index)))
'            If CDbl(txtAmountPaid(index).Text) > Budget Then
'             MsgBox " These Amount to be post is more than the budget or budgetted ", vbInformation
'              Exit Sub
'             Else
'              End If
'
'             If Budget = 0 Then
'               MsgBox " The Accno  have not been  budgetted! please do it first ", vbInformation
'              Exit Sub
'            End If
'
'       Else
'       End If
    
    
    sql = "select receiptno,chequeno from receiptbooking where receiptno in ('" & txtReceiptsno(index) & "')"
    Set rst = oSaccoMaster.GetRecordset(sql)
    If Not rst.EOF Then
        MsgBox "Either receiptno or Chequeno is already used, Get another One!"
        Exit Sub
    End If
    Select Case txtmode(index).Text
        Case "Cheque"
            chequeno = txtmode(index)
        Case Else
            chequeno = txtmode(index).Text
    End Select
    
    
    '***********************************BEGINNING OF POSTING**************************************
    ref = "Nominal"
    RefNo = ""

    With lvwNtrans(index)
        If .ListItems.Count < 0 Then
            Exit Sub
        End If
        I = 0
        j = 0
        I = .ListItems.Count
'        saveTransno (User)
'        TransNo = getTransno
    
        oSaccoMaster.goConn.BeginTrans
        On Error GoTo TransError
        'save TransactionNo
        transactionTotal = CDbl(txtAmountPaid(index).Text)
        NewTransaction transactionTotal, DTPDatedeposited(index), "Nominal Receipting - Document No " & txtReceiptsno(index).Text
        
            
            Dim fperiod As Integer, intRate As Double
        
            For j = 1 To I
                    mMemberNo = ""
                If Save_GLTRANSACTION(DTPDatedeposited(index), lvwNtrans(index).ListItems(j).SubItems(2), IIf(index = 0, cboBanks(index), lvwNtrans(index).ListItems(j)), IIf(index = 0, lvwNtrans(index).ListItems(j), cboBanks(index)), _
                txtReceiptsno(index), mMemberNo, User, "", CboParticulars(index), 1, 1, txtmode(index), transactionNo, TransNo) = False Then
                    GoTo TransError
                End If
             oSaccoMaster.Execute (" insert into PrintTemp(MemberNo,Amount,Description) values('" & mMemberNo & "','" & lvwNtrans(index).ListItems(j).SubItems(2) & "','" & CboParticulars(index) & "')")

               '// SAVE TO MASTERS TABLE
               Dim prodtype As Double
               Dim sharecode As Double
               Dim TransCode As Double
                 prodtype = "NOMINAL"
                 sharecode = ""
                 TransCode = Cbo1(index)
               
     oSaccoMaster.ExecuteThis (" set dateformat dmy Insert into Masters(Transdate,source,Productcode,ProdType,Amount,Refno,Transcode,Users, TransactionNo,Machine) " _
    & " Values('" & DTPDatedeposited(index) & "','" & mMemberNo & "','" & sharecode & "','" & prodtype & "'," & lvwNtrans(index).ListItems(j).SubItems(2) & ",'" & txtReceiptsno(index) & "','" & TransCode & "','" & User & "')")

            Next j
            
            If Not saveReceipt(txtReceiptsno(index), ref, RefNo, cboBanks(index).Text, cboBanks(index).Text, "Receipt", DTPDatedeposited(index).value, CDbl(txtAmountPaid(index).Text), txtmode(index), cboMode(index).Text, "Deposit") Then
                GoTo TransError
            Else
                For I = 1 To lvwNtrans(index).ListItems.Count
                    oSaccoMaster.ExecuteThis ("if exists (select * from receiptbooking where receiptno='" & txtReceiptsno(index).Text & "') update receiptbooking set posted=1 where receiptno='" & txtReceiptsno(index).Text & "' Insert into ReceiptDetails (ReceiptNo,Description,Amount) values('" & txtReceiptsno(index) & "','" & lvwNtrans(index).ListItems(I).ListSubItems(1) & "'," & lvwNtrans(index).ListItems(I).ListSubItems(2) & ")")
                    If success = False Then
                        GoTo TransError
                    End If
                    

                Next I
                
            End If
        End With
    oSaccoMaster.goConn.CommitTrans
    
    lvwNtrans(index).ListItems.clear
    MsgBox ("Receipt updated successfully!"), vbInformation
    If SSTab1.Caption = "RECEIPTS" Then
      If MsgBox("Print Receipt?", vbQuestion + vbYesNo) = vbYes Then
          
        PrintReceipt
        PrintReceipt
    End If
    End If
    
   'Call clear
  'Form_Load
'   txtReceiptsno(0) = Generate_ReceiptNo("Receipt")
'    txtReceiptsno(1) = Generate_ReceiptNo("Voucher")
'txtReceiptsno(Index) = Generate_ReceiptNo("Receipt")
    
    Exit Sub
SysError:
    MsgBox IIf(ErrorMessage = "", err.description, ErrorMessage), vbCritical, Me.Caption
    GoTo TransError
    Exit Sub
TransError:
    If err.number = 35600 Then
        Resume Next
    End If
    If ErrorMessage = "" And err.description = "" Then
        oSaccoMaster.goConn.RollbackTrans
    Else
        MsgBox IIf(ErrorMessage = "", err.description, ErrorMessage) & vbNewLine & "Action Therefore Aborted. ", vbCritical, Me.Caption
        oSaccoMaster.goConn.RollbackTrans
    End If
End Sub
Private Sub cmdClear_Click(index As Integer)
'    On Error Resume Next
'    Dim vote As String
'    Dim msgb As Response
'    Dim rsponse As String
'
'
'    With lvwNtrans(index)
'        If .ListItems.Count > 0 Then
'            If MsgBox("Are you sure you want to clear the entire list?", vbQuestion + vbYesNo) = vbYes Then
'                .ListItems.clear
'            End If
'        End If
'    End With
'
'    Recalculate (index)
        
End Sub
Private Sub Form_Load()
    On Error Resume Next
    DTPDatedeposited(0) = Format(Get_Server_Date, "dd/mm/yyyy")
    DTPDatedeposited(1) = Format(Get_Server_Date, "dd/mm/yyyy")
    
    Dim rscompany As New Recordset
    Dim rsTrans As New Recordset

    'Load Gl's
    sql = "Select accno from glsetup order by accno asc"
    Set rst = oSaccoMaster.GetRecordset(sql)
    While Not rst.EOF
        cboAccno(0).AddItem (rst(0))
        cboAccno(1).AddItem (rst(0))
        rst.MoveNext
    Wend
    'load the banks
    cboBanks(0).clear
    cboBanks(1).clear
    sql = "select b.Accno from banks b inner join glsetup g  on b.Accno = g.Accno   where b.accno not in (select assigngl from useraccounts where userloginid not in('" & User & "') and assigngl<>'' and assigngl is not null )"
    Set rst = oSaccoMaster.GetRecordset(sql)
      
    While Not rst.EOF
        cboBanks(0).AddItem rst(0)
        cboBanks(1).AddItem rst(0)
        rst.MoveNext
    Wend
    
    
    
     CboParticulars(0).clear
    CboParticulars(1).clear
    sql = "select  Description   from TransCodebosa  order by Description  asc "
    Set rsTrans = oSaccoMaster.GetRecordset(sql)
      
    While Not rsTrans.EOF
        CboParticulars(0).AddItem rsTrans(0)
        CboParticulars(1).AddItem rsTrans(0)
        rsTrans.MoveNext
    Wend
    
    
    cboBanks(0).List(0) = currentUser.tellerGlAcc
    cboBanks(0).Text = cboBanks(0).List(0)
    cboBanks(1).List(0) = currentUser.tellerGlAcc
    cboBanks(1).Text = cboBanks(1).List(0)
    
   
    totalamount = 0
    pushed = 0
    
'     Cbo2.clear
'
'     Set rst1 = Nothing
'     sql = "Select TranscationCode from TransactionCode"
'    Set rst1 = oSaccoMaster.GetRecordSet(sql)
'    While Not rst1.EOF
'        Cbo2.AddItem (rst1(0))
'        rst1.MoveNext
'    Wend
    Cbo1(0).clear
    Cbo1(1).clear
    'Combo2.Text = "OR"
     Set rst1 = Nothing
     sql = "Select TranscationCode from TransactionCode"
    Set rst1 = oSaccoMaster.GetRecordset(sql)
    While Not rst1.EOF
        Cbo1(0).AddItem (rst1(0))
         Cbo1(1).AddItem (rst1(0))
        rst1.MoveNext
    Wend
    
'    txtReceiptsno(0) = Generate_ReceiptNo("Receipt")
'    txtReceiptsno(1) = Generate_ReceiptNo("Voucher")
'

    IsBatch = False
    commit = False 'will be used to determine whether a loaded interest is saved as a transaction
End Sub
Private Sub Form_Unload(Cancel As Integer)
    'Stop subclassing
    CloseSubClass
    'Clean up by setting the classes to Nothing
    Set objLabelEdit = Nothing
    'Set objLabelEdit2 = Nothing
End Sub

Private Sub lvwNTrans_DblClick(index As Integer)
    Dim total As Double, amt As Double
    Dim ccount As Integer
    On Error Resume Next
    total = 0
    With lvwNtrans(index)
        If .ListItems.Count > 0 Then
            amt = InputBox("Enter the amount", "AMOUNT", .SelectedItem.ListSubItems(2))
            .SelectedItem.ListSubItems(2) = amt
        End If
    End With
    Recalculate (index)
End Sub
Private Sub Recalculate(index As Integer)
    On Error Resume Next
    Dim Balance As Double
    Dim thislistview As ListView
    
    If lvwNtrans(index).ListItems.Count > 0 Then
        For I = 1 To lvwNtrans(index).ListItems.Count
            Balance = Balance + CDbl(lvwNtrans(index).ListItems(I).SubItems(2))
        Next I
    End If

    txtDistributed(index) = Balance
    txtBalance(index) = (txtAmountPaid(index)) - (txtDistributed(index))
End Sub


Private Sub lvwNtrans_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
       lvwNTrans_DblClick (index)
    End If
End Sub

Private Sub txtAmountPaid_Change(index As Integer)
On Error Resume Next
    If txtAmountPaid(index).Text = "" Then txtAmountPaid(index).Text = 0
    totalamount = CDbl(txtAmountPaid(index).Text)
    pushed = 0
    txtBalance(index).Text = totalamount - CDbl(txtDistributed(index).Text)
End Sub
Private Sub txtAmountPaid_KeyPress(index As Integer, KeyAscii As Integer)
    If keyIsValid(KeyAscii, 1) = False Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub cmdPrint_Click()
  Dim ssss As String
    On Error GoTo SysError
    If Trim$(txtReceiptsno(1)) = "" Then
        MsgBox "Please enter the receipt number", vbInformation, Me.Caption
        Exit Sub
    End If
    ssss = txtReceiptsno(1).Text
       If ssss = txtReceiptsno(0).Text Then
       GoTo www
      End If
    reportname = "Voucherothers.rpt"
    STRFORMULA = "{ReceiptBooking.ReceiptNo}='" & ssss & "'  "
    Show_Sales_Crystal_Report STRFORMULA, reportname, CompanyName
    
www:

    Exit Sub
SysError:
    MsgBox err.description, vbInformation
'txtothers = ""
End Sub

Private Sub PrintReceipt()
    On Error GoTo SysError
    
    Dim pay, tot, disc As Currency
    Dim Z, X As Integer
    'for number of copies
    Dim a As Integer
    Dim b As Integer
    Dim c As Integer
    Dim rsprint As New ADODB.Recordset
     Dim printamount As Double
     Dim DESC As String
    dlg9.CancelError = True
    dlg9.FontName = "Garamond"
    Dim j As Printer
    
    a = dlg9.Copies
    Printer.CurrentY = 500
    Printer.CurrentX = 9000
    Printer.FontSize = 10
    Printer.CurrentY = 600
    Printer.CurrentX = 1000
    Printer.Print Tab(0); "TRANSACTION VOUCHER"
     Printer.Print "DATE :  "; DTPDatedeposited(0).value
    Printer.Print Tab(0); "-----------------------------"
    Printer.Print Tab(0); CompanyName
    Printer.Print Tab(0); CompanyPhone
    Printer.Print Tab(0); CompanyTown
    Printer.Print Tab(0); "-----------------------------"
    Printer.Print Tab(0); txtReceiptsno(0).Text
    'Printer.Print Tab(0); txtMemberNo.Text
   ' Printer.Print Tab(0); lblfullnames
    Printer.Print
    Printer.CurrentX = 500#
    Printer.FontSize = 10
    Printer.CurrentX = 500
    Printer.FontSize = 8
    Printer.CurrentX = 500
    Printer.FontSize = 8
    Printer.CurrentX = 500
    Printer.FontSize = 8
    Printer.Print Tab(0); "ITEM DESCRIPTION"; Tab(18); "AMOUNT"
    Printer.Print Tab(0); "---------------------------------"
      Set rsprint = oSaccoMaster.GetRecordset("select  Amount, Description   from PrintTemp  ")
         While Not rsprint.EOF
          printamount = IIf(IsNull(rsprint(0)), 0, rsprint(0))
          DESC = IIf(IsNull(rsprint(1)), "", rsprint(1))
    
    Printer.Print Tab(0); DESC; Tab(18); printamount
    rsprint.MoveNext
    Wend
    Printer.Print Tab(0); "------------------------------"
    Printer.Print "Total Amount :"; Tab(18); txtDistributed(0)
    'Printer.Print "Your Balance is :"; Tab(20); asa
    Printer.Print Tab(0); "------------------------------"
    Printer.Print Tab(0); "You were Served by: " & User
    Printer.Print Tab(0); "Customer Signature   /Thumb Print"
    Printer.Print
     Printer.Print
    Printer.Print Tab(0); "_____________________________"
    Printer.Print Tab(0); " Teller Signature  /Thumb Print"
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print Tab(0); "______________________________"
    Printer.Print Tab(0); "Remarks"
    Printer.Print
    Printer.Print Tab(0); "Thank you and come again****"
   
    Printer.Print
    Printer.Print Tab(0); "POWERED BY EASYSACCO"
    Printer.Print
    Printer.EndDoc
    
    Exit Sub
SysError:
    MsgBox err.description, vbInformation
End Sub

 Sub clear()
    If SSTab1.Caption = "RECEIPTS" Then
     CboParticulars(0).Text = ""
        txtPayee(0).Text = ""
        txtAmountPaid(0).Text = ""
        txtDistributed(0).Text = ""
        txtBalance(0).Text = ""
        cboAccno(0).clear
        cboBanks(0).clear
        lblbankname(0).Caption = ""
        txtAccNames(0).Text = ""
        
    Else
         CboParticulars(1).Text = ""
        txtPayee(1).Text = ""
        txtAmountPaid(1).Text = ""
        txtDistributed(1).Text = ""
        txtBalance(1).Text = ""
        cboAccno(1).clear
        txtAccNames(1).Text = ""
        cboBanks(1).clear
        lblbankname(1).Caption = ""
        
    End If
 End Sub

Private Function getGlAccMainGroup(ACCNO As String)
    Dim rstacc As New ADODB.Recordset
   Set rstacc = oSaccoMaster.GetRecordset(" select GlAccMainGroup from glsetup  where accno ='" & ACCNO & "' ")
       If Not rstacc.EOF Then
       getGlAccMainGroup = IIf(IsNull(rstacc("GlAccMainGroup")), "", rstacc("GlAccMainGroup"))
       Else
       getGlAccMainGroup = ""
       End If

End Function

Private Function getbudget(ACCNO, mm As Integer, yy As Integer)
   Dim rsBudgetted  As New ADODB.Recordset
    Set rsBudgetted = oSaccoMaster.GetRecordset(" select Budgetted  from   budgets where Accno='" & ACCNO & "'  and mmonth= '" & mm & "' and  yyear= '" & yy & "'  ")
       If Not rsBudgetted.EOF Then
         getbudget = IIf(IsNull(rsBudgetted(0)), 0, rsBudgetted(0))
         Else
          getbudget = 0
       End If
End Function


