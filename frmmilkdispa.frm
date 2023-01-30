VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   Caption         =   "MILK DISPATCH FORM"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "frmmilkdispa"
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Height          =   8400
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12495
      Begin VB.CommandButton cmdstatement 
         Caption         =   "Debtors Statement"
         Height          =   375
         Left            =   2880
         TabIndex        =   69
         Top             =   6840
         Width           =   2415
      End
      Begin VB.TextBox txtIntake 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1680
         TabIndex        =   66
         Top             =   3360
         Width           =   2295
      End
      Begin VB.TextBox txtVariance 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1680
         TabIndex        =   65
         Top             =   3720
         Width           =   2295
      End
      Begin VB.CheckBox chkPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Caption         =   "Print Receipt"
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   0
         TabIndex        =   64
         Top             =   7320
         Width           =   1695
      End
      Begin VB.ComboBox ports 
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         ItemData        =   "frmmilkdispa.frx":0000
         Left            =   3120
         List            =   "frmmilkdispa.frx":0010
         TabIndex        =   62
         Text            =   "\\127.0.0.1\E-PoS 80mm Thermal Printer"
         Top             =   7320
         Width           =   2175
      End
      Begin VB.CheckBox chprint 
         Caption         =   "Use LPT1 Printer"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   0
         TabIndex        =   61
         Top             =   6840
         Width           =   3255
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   375
         Left            =   2640
         TabIndex        =   60
         Top             =   7800
         Width           =   1335
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "New"
         Height          =   375
         Left            =   0
         TabIndex        =   59
         Top             =   7800
         Width           =   1335
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit"
         Height          =   375
         Left            =   1320
         TabIndex        =   58
         Top             =   7800
         Width           =   1335
      End
      Begin VB.CommandButton cmdreprint 
         Caption         =   "Reprint"
         Height          =   375
         Left            =   3960
         TabIndex        =   57
         Top             =   7800
         Width           =   1335
      End
      Begin VB.TextBox txtdcode 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1680
         TabIndex        =   50
         Top             =   2280
         Width           =   2055
      End
      Begin VB.PictureBox Picture3 
         Height          =   285
         Left            =   3720
         Picture         =   "frmmilkdispa.frx":002C
         ScaleHeight     =   225
         ScaleWidth      =   195
         TabIndex        =   49
         Top             =   2280
         Width           =   255
      End
      Begin VB.TextBox txtDispatch 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   1680
         TabIndex        =   46
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox txtamountp 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   1680
         TabIndex        =   45
         Top             =   1200
         Width           =   2295
      End
      Begin VB.TextBox txtRefNo 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   1680
         TabIndex        =   43
         Top             =   240
         Width           =   2295
      End
      Begin VB.PictureBox Picture1 
         Height          =   285
         Left            =   3960
         Picture         =   "frmmilkdispa.frx":02EE
         ScaleHeight     =   225
         ScaleWidth      =   195
         TabIndex        =   42
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton cmdnewsearch 
         Caption         =   "New "
         Height          =   285
         Left            =   4200
         TabIndex        =   41
         Top             =   240
         Width           =   615
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2055
         Left            =   120
         TabIndex        =   40
         Top             =   4800
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   3625
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   65280
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Dcode"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "DQuantity"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.TextBox txtmode 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   10140
         TabIndex        =   39
         Top             =   1170
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.TextBox txtPayee 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   8640
         TabIndex        =   23
         Top             =   2040
         Width           =   2265
      End
      Begin VB.TextBox txtParticulars 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   0
         Left            =   8640
         TabIndex        =   22
         Top             =   1605
         Width           =   2745
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
         Index           =   0
         Left            =   5040
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   21
         Text            =   "0"
         Top             =   2880
         Width           =   900
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
         Index           =   0
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "0"
         Top             =   3240
         Width           =   900
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
         Left            =   8640
         TabIndex        =   19
         Top             =   2400
         Width           =   1380
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
         Left            =   5040
         MaxLength       =   9
         TabIndex        =   18
         Text            =   "0"
         Top             =   2520
         Width           =   900
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
         ItemData        =   "frmmilkdispa.frx":05B0
         Left            =   8640
         List            =   "frmmilkdispa.frx":05B7
         TabIndex        =   17
         Text            =   "Cash"
         Top             =   1155
         Width           =   1425
      End
      Begin VB.CommandButton cmdReceipt 
         Caption         =   "<>"
         Height          =   300
         Index           =   0
         Left            =   10035
         TabIndex        =   16
         Top             =   2400
         Width           =   345
      End
      Begin VB.CommandButton cmdupdatereceipt 
         Caption         =   "&Post"
         Height          =   375
         Index           =   0
         Left            =   7080
         TabIndex        =   15
         Top             =   6480
         Width           =   1425
      End
      Begin VB.CommandButton cmdBank 
         Caption         =   "<>"
         Height          =   300
         Index           =   0
         Left            =   8505
         TabIndex        =   14
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
         ItemData        =   "frmmilkdispa.frx":05C1
         Left            =   7200
         List            =   "frmmilkdispa.frx":05C3
         TabIndex        =   13
         Top             =   750
         Width           =   1350
      End
      Begin VB.ComboBox cboAccno 
         Height          =   315
         Index           =   0
         Left            =   7515
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   3210
         Width           =   1200
      End
      Begin VB.TextBox txtAccNames 
         Height          =   315
         Index           =   0
         Left            =   9015
         TabIndex        =   11
         Top             =   3210
         Width           =   2985
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add>>"
         Height          =   345
         Index           =   0
         Left            =   5040
         TabIndex        =   10
         Top             =   3600
         Width           =   930
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "<<Remove"
         Height          =   345
         Index           =   0
         Left            =   5040
         TabIndex        =   9
         Top             =   3960
         Width           =   930
      End
      Begin VB.CommandButton cmdAcctsSearch 
         Height          =   300
         Index           =   0
         Left            =   8685
         Picture         =   "frmmilkdispa.frx":05C5
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   3210
         Width           =   330
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear<<<"
         Height          =   345
         Index           =   0
         Left            =   5040
         TabIndex        =   7
         Top             =   4320
         Width           =   930
      End
      Begin VB.CheckBox chkCreditors 
         Caption         =   "Debtors"
         Height          =   255
         Index           =   0
         Left            =   7560
         TabIndex        =   6
         Top             =   2880
         Width           =   975
      End
      Begin VB.CommandButton cmddebitcreditnote 
         Caption         =   "Debit/Credit "
         Height          =   315
         Left            =   9825
         TabIndex        =   5
         Top             =   6480
         Width           =   1215
      End
      Begin VB.CheckBox chkDebtno 
         Caption         =   "Paid by Debtor"
         Height          =   195
         Left            =   5880
         TabIndex        =   4
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox txtTCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   4920
         TabIndex        =   3
         Top             =   1440
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtDNames 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   6000
         TabIndex        =   2
         Top             =   1440
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.PictureBox Picture5 
         Height          =   255
         Left            =   5760
         Picture         =   "frmmilkdispa.frx":06C7
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   1
         Top             =   1440
         Visible         =   0   'False
         Width           =   255
      End
      Begin MSComctlLib.ListView lvwNtrans 
         Height          =   2580
         Index           =   0
         Left            =   6480
         TabIndex        =   24
         Top             =   3840
         Width           =   4620
         _ExtentX        =   8149
         _ExtentY        =   4551
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
      Begin MSComCtl2.DTPicker dtpTransDate 
         Height          =   375
         Index           =   0
         Left            =   4995
         TabIndex        =   25
         Top             =   480
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   122748929
         CurrentDate     =   40421
      End
      Begin MSComDlg.CommonDialog cdgPrint 
         Left            =   5400
         Top             =   7680
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         FileName        =   "c:\receipt.txt"
      End
      Begin MSComDlg.CommonDialog dlg9 
         Left            =   8880
         Top             =   6600
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "Intake :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   120
         TabIndex        =   68
         Top             =   3360
         Width           =   1335
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "Variance :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   120
         TabIndex        =   67
         Top             =   3720
         Width           =   1260
      End
      Begin VB.Line Line1 
         X1              =   4920
         X2              =   4920
         Y1              =   120
         Y2              =   6840
      End
      Begin VB.Label Label18 
         Caption         =   "Printer Port"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   63
         Top             =   7320
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Acc Dr"
         Height          =   255
         Left            =   120
         TabIndex        =   56
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Acc Cr"
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label Label10 
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   54
         Top             =   2640
         Width           =   2295
      End
      Begin VB.Label Label11 
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   53
         Top             =   3000
         Width           =   2295
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "Dispatch : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   480
         TabIndex        =   52
         Top             =   720
         Width           =   930
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "Debtors Code :"
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   51
         Top             =   2280
         Width           =   1410
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "Dispatch : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   -1800
         TabIndex        =   48
         Top             =   1320
         Width           =   930
      End
      Begin VB.Label Label13 
         Caption         =   "Amounts payable"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   47
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "Reference No. :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   240
         TabIndex        =   44
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label Label8 
         Caption         =   "Transaction Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   5160
         TabIndex        =   38
         Top             =   120
         Width           =   1590
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
         Left            =   7560
         TabIndex        =   37
         Top             =   2040
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
         Left            =   7440
         TabIndex        =   36
         Top             =   1680
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
         Left            =   6000
         TabIndex        =   35
         Top             =   2880
         Width           =   1005
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Receipt No"
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
         Left            =   7605
         TabIndex        =   34
         Top             =   2400
         Width           =   870
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
         Index           =   0
         Left            =   6000
         TabIndex        =   33
         Top             =   2520
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
         Left            =   7380
         TabIndex        =   32
         Top             =   1200
         Width           =   1230
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Bank A/C (Source)"
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
         Left            =   8280
         TabIndex        =   31
         Top             =   240
         Width           =   1485
      End
      Begin VB.Label lblbankname 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   0
         Left            =   8850
         TabIndex        =   30
         Top             =   765
         Width           =   3255
      End
      Begin VB.Label lblSupplier 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   1680
         TabIndex        =   29
         Top             =   2760
         Width           =   3165
      End
      Begin VB.Label Label1 
         Caption         =   "Code"
         Height          =   255
         Left            =   5040
         TabIndex        =   28
         Top             =   1200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "CR A/C"
         Height          =   255
         Index           =   0
         Left            =   7320
         TabIndex        =   27
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "DR A/C"
         Height          =   255
         Index           =   0
         Left            =   8640
         TabIndex        =   26
         Top             =   2880
         Width           =   975
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Price As Currency
Dim capp As Integer
Dim crate As Double
Dim rsq As New Recordset
Dim milksup As Double
Dim amtpayable As Double
Dim receipno As Double
Dim dispatchby As Double
Dim qty As Double
Dim Totalamount As Currency
Dim pushed As Currency
'Dim objLabelEdit As LabelEdit
'Dim objLabelEdit2 As LabelEdit
'Dim objLabelEdit3 As LabelEdit
Dim interestAcc As String, LoanAcc As String
Dim IsBatch As Boolean, IsGroup As Boolean, isfixed As Boolean
Dim k As Integer
Dim shareBal As Double, balance As Double
Dim daysIntoTheMonth As Integer
Dim Posted As Boolean
Dim ref As String, RefNo As String
Dim penalise As Boolean
Dim commit As Boolean
Dim rsDefaults As ADODB.Recordset
Dim fSeed As Integer
Dim newMember As Boolean



Private Sub cboAccno_Change(Index As Integer)
    Dim ACCNO As String
    ACCNO = cboAccno(Index).Text
    sql = "select GLACCNAME from glsetup where accno='" & ACCNO & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
        txtAccNames(Index).Text = rs(0)
    Else
        txtAccNames(Index).Text = ""
    End If
End Sub

Private Sub cboAccno_Click(Index As Integer)
    cboAccno_Change (Index)
End Sub



Private Sub cboBanks_Change(Index As Integer)
    On Error GoTo SysError
    Dim rsGlaccname As New ADODB.Recordset
'    sql = "select bankname from banks where Accno='" & cboBanks(index).Text & "'"
'    Set rst = oSaccoMaster.GetRecordset(sql)
'    If Not rst.EOF Then
'        lblbankname(index).Caption = rst(0)
'    Else
'        lblbankname(index).Caption = ""
'    End If
'    Exit Sub
    Dim ACCNO As String
    ACCNO = cboBanks(Index).Text
    sql = "select Glaccname from glsetup where accno='" & ACCNO & "'"
    Set rsGlaccname = oSaccoMaster.GetRecordset(sql)
    If Not rsGlaccname.EOF Then
        lblbankname(Index).Caption = rsGlaccname("GLACCNAME")
    Else
       lblbankname(Index).Caption = ""
    End If
    Exit Sub

SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub cboBanks_Click(Index As Integer)
    cboBanks_Change (Index)
End Sub




Private Sub cboMode_Change(Index As Integer)
'On Error Resume Next
    Select Case cboMode(Index).Text
        Case "Cash"
            txtmode(Index).Visible = False
            txtReceiptsno(Index).SetFocus
        Case "Cheque"
            txtmode(Index).Visible = True
            txtmode(Index).SetFocus
        Case "Direct Deposit"
            txtmode(Index).Visible = True
            txtmode(Index).SetFocus
        Case Else

    End Select
End Sub

Private Sub cboMode_Click(Index As Integer)
    cboMode_Change (Index)
End Sub





Private Sub chkCreditors_Click(Index As Integer)
    With chkCreditors(Index)
        If Index = 0 Then
            If .value = 1 Then
                frmSearchDebtors.Show vbModal
                sql = "Select dname companyName from d_debtors where dcode='" & sel & "'"
            End If
        Else
            If .value = 1 Then
                frmSearchVendor.Show vbModal
                sql = "Select companyName from ag_supplier1 where supplierId='" & sel & "'"
            End If
        End If
        
        
        If .value = 1 Then
            Set rst = oSaccoMaster.GetRecordset(sql)
            If Not rst.EOF Then
                lblSupplier(Index).Caption = rst!CompanyName
            Else
                lblSupplier(Index).Caption = ""
            End If
        Else
            InvoiceNo = ""
            InvoiceBal = 0#
            lblSupplier(Index).Caption = ""
        End If
    End With
End Sub

Private Sub chkDebtno_Click()
If chkDebtno.value = vbChecked Then
    txtTCode.Visible = True
    Picture5.Visible = True
    txtDNames.Visible = True
    Label1.Visible = True
    txtDNames = "<select Debtors Code here>"
    'txtNames.Visible = True
    txtTCode.SetFocus
    
    
Else
    txtTCode.Visible = True
    Picture5.Visible = True
    txtDNames.Visible = True
    Label1.Visible = False
'    txtDebt.Visible = False

End If
    If txtTCode = "" Then
        MsgBox "Please enter the Debtors Code", vbCritical
        Exit Sub
    End If
End Sub

Private Sub cmdAcctsSearch_Click(Index As Integer)
    On Error Resume Next
    frmAcctsSearch.Show vbModal, Me
    If Continue Then
        If SearchValue <> "" Then
            cboAccno(Index) = SearchValue
            SearchValue = ""
            Continue = False
        End If
    End If
End Sub

Private Sub cmdAdd_Click(Index As Integer)
    On Error GoTo SysError
    If cboAccno(Index).Text = "" Then
        Exit Sub
    End If

    Set li = lvwNtrans(Index).ListItems.Add(, , cboAccno(Index))
    li.SubItems(1) = txtAccNames(Index)
    li.SubItems(2) = InvoiceBal
   
    Recalculate (Index)
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub cmdBank_Click(Index As Integer)
    On Error Resume Next
    frmAcctsSearch.Show vbModal, Me
    If Continue Then
        If SearchValue <> "" Then
            cboBanks(Index) = SearchValue
            SearchValue = ""
            Continue = False
        End If
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub PrintReceipt(Index As Integer)
    On Error GoTo SysError
    
    
    If Index = 0 Then
        reportname = "Receipt.rpt"
        STRFORMULA = "{VwReceipt.ReceiptNo}='" & txtReceiptsno(Index) & "'"
    Else
        reportname = "Voucher.rpt"
        STRFORMULA = "{paymentBooking.VoucherNo}='" & txtReceiptsno(Index) & "'"
    End If
    Show_Sales_Crystal_Report STRFORMULA, reportname, CompanyName
    
    Exit Sub

    Dim pay, tot, disc As Currency
    Dim Z, X As Integer
    'for number of copies
    Dim a As Integer
    Dim b As Integer
    Dim C As Integer
    dlg9.CancelError = True
    dlg9.FontName = "Garamond"
    Dim j As Printer

    a = dlg9.Copies
    Printer.CurrentY = 500
    Printer.CurrentX = 9000
    Printer.FontSize = 11
    Printer.CurrentY = 600
    Printer.CurrentX = 1000
    Printer.Print Tab(4); "TRANSACTION VOUCHER"
    Printer.Print Tab(0); "---------------------------------------"
    Printer.Print Tab(0); CompanyName
    'Printer.Print Tab(0); CompanyPhone
    'Printer.Print Tab(0); CompanyTown
    Printer.Print Tab(0); "---------------------------------------"
    Printer.Print Tab(0); "Document No"; Tab(20); txtReceiptsno(Index).Text
    Printer.Print Tab(0); "Date :  "; dtpTransDate(0).value
    Printer.Print Tab(0); "Source"; Tab(10); lblbankname(Index)

    Printer.CurrentX = 500#
    Printer.FontSize = 10
    Printer.CurrentX = 500
    Printer.FontSize = 8
    
    Printer.CurrentX = 500
    Printer.FontSize = 8
    Printer.Print
    Printer.CurrentX = 500
    Printer.FontSize = 8
    Printer.Print
    Printer.Print Tab(2); "Payment Type"; Tab(30); cboMode(0).Text
    Printer.Print Tab(0); "-------------------------------------------------"
    If Index = 0 Then
        Printer.Print "Total Received :"; Tab(20); txtDistributed(Index)
    Else
        Printer.Print "Total Paid :"; Tab(20); txtDistributed(Index)
    End If
    Printer.Print "Particulars :"; Tab(10); txtParticulars(Index)
    Printer.Print Tab(0); "------------------------------------------------"
    Printer.Print Tab(2); "ITEM"; Tab(30); "AMOUNT"
    Printer.Print Tab(0); "------------------------------------------------"
    Printer.FontSize = 8
    For I = 1 To lvwNtrans(Index).ListItems.Count
        Printer.Print lvwNtrans(Index).ListItems(I).ListSubItems(1); Tab(30); lvwNtrans(Index).ListItems(I).ListSubItems(2)
    Next I

    'Printer.Print "Your Balance is :"; Tab(20); asa
    Printer.Print Tab(0); "------------------------------------------------"
    Printer.Print Tab(0); "Document No: "; txtReceiptsno(0).Text
    Printer.Print Tab(2); "You were Served by: " & User
    Printer.Print Tab(2); "Receipient Signature   /Thumb Print"
    Printer.Print
    Printer.Print Tab(0); "Remarks"
    Printer.Print
    Printer.Print Tab(0); "---------------------------------------"
    'Printer.Print Tab(0); CompanyTagLine
    Printer.Print
    Printer.Print
    Printer.Print Tab(0); "POWERED BY EASYMA"
    Printer.Print
    Printer.EndDoc

'//-------------------------
'mysql = ""
'mysql = "set dateformat dmy Insert into Receiptno(Receiptno,Auditdate,auditid,memberno)values('" & txtReceiptsno & "','" & Format(Get_Server_Date, "dd/MM/yyyy") & "','" & User & "','" & txtMemberNo & "')"
'oSaccoMaster.ExecuteThis (mysql)
    Exit Sub
SysError:
    MsgBox err.description, vbInformation
End Sub
Private Sub cmddebitcreditnote_Click()
frmdebitcreditnote.Show vbModal, Me
End Sub

Private Sub cmdnewsearch_Click()
Dim rsr As New ADODB.Recordset
Dim rsg As New ADODB.Recordset
Dim I As Object
Dim Mylength As Integer
'//if this record is new then look for receipts no

''//clear all textboxes





'mysql = ""
'mysql = "set dateformat dmy select GenerateReceiptno from param"
sql = ""
sql = "set dateformat dmy select GenerateReceiptno from param"
Set rsg = oSaccoMaster.GetRecordset(sql)
If Not rsg.EOF Then
    ''''check check
    If rsg!GenerateReceiptno = True Then
    
        sql = ""
        sql = "select * from Receiptno where receiptno like 'RF-%' order by Receipthnoid desc"
        
        Set rsr = oSaccoMaster.GetRecordset(sql)
        
        If Not rsr.EOF Then
            Mylength = CInt(Mid(rsr!ReceiptNo, 5, 10))
            Mylength = Mylength + 1
            txtRefNo = Padding(Mylength)
            txtRefNo = "RF-" & txtRefNo
        Else
            Mylength = 1
            txtRefNo = "RF-" & Padding(Mylength)
            
        End If
Else
    ''//receiptno  will be keyed in
End If
End If
End Sub

Private Sub cmdRemove_Click(Index As Integer)
    On Error GoTo SysError
    With lvwNtrans(Index)
        If .ListItems.Count > 0 Then
            If MsgBox("Do you want to remove " & lvwNtrans(Index).SelectedItem.SubItems(1) & _
            " From the list?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
                Exit Sub
            End If
            pushed = pushed - .SelectedItem.ListSubItems(2)
            .ListItems.Remove (.SelectedItem.Index)
        End If
    End With
    Recalculate (Index)
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub cmdreprint_Click()
STRFORMULA = "{d_MilkControl.RefNo}='" & txtRefNo & "'"
    reportname = "milkinvoice.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, title
End Sub

Private Sub cmdSave_Click()
If txtdcode = "" Then
MsgBox "Debtors code cannot be blank; input an existing one", vbCritical
Exit Sub
End If
If txtDispatch = "" Then
    MsgBox "Please enter the dispatch quantity."
        txtDispatch.SetFocus
    Exit Sub
End If

If txtDipping = "" Then
    MsgBox "Please enter the dipping quantity."
        txtDipping.SetFocus
    Exit Sub
End If

If txtIntake = "" Then
    MsgBox "Please enter the intake quantity."
        txtIntake.SetFocus
    Exit Sub
End If

If txtVariance = "" Then
    MsgBox "Please enter the variance quantity."
        txtVariance.SetFocus
    Exit Sub
End If



If txtRefNo = "" Then
    MsgBox "Please enter the reference number."
        txtRefNo.SetFocus
    Exit Sub
End If
'//check if the dispatch is greater than the dipping
If CDbl(txtDipping) < CDbl(txtDispatch) Then 'raiise an alarm
MsgBox "You cannot take more you have in the tank", vbCritical
Exit Sub
End If
Dim Debit As String
Dim Credit As String

sql = ""
    sql = "SET      dateformat dmy     SELECT     *     FROM         d_MilkControl    WHERE     DispDate = '" & dtpTransDate & "' and DispQnty = '" & txtDispatch & "'and dcode = '" & txtdcode & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
    MsgBox "You have already dispatch for that day", vbInformation
    Exit Sub
    End If


'Dim Price As Currency

'Set rs = oSaccoMaster.GetRecordset("d_sp_getAccName '" & lblDebtors & "'")
'If IsNull(rs.Fields(0)) Then
'    MsgBox "The debtors account not set. " & vbNewLine & "Please contact the accountant to set GL for " & lblDebtors
'        Exit Sub
'End If
'
Debit = Label10
'
'Set rs = oSaccoMaster.GetRecordset("d_sp_getAccName 'Milk sale'")
'If IsNull(rs.Fields(0)) Then
'    MsgBox "The Creditors account not set. " & vbNewLine & "Please contact the accountant to set GL for milk sales"
'        Exit Sub
'End If
'
Credit = Label11

    

    If Not Save_GLTRANSACTION(Format(dtpTransDate, "dd/mm/yyyy"), (CCur(Price) * CCur(txtDispatch)), Debit, Credit, "Milk Sales ", txtRefNo, User, ErrorMessage, "Milk Sales", 1, 1, txtRefNo, transactionNo, "", "", 0) Then
            If ErrorMessage <> "" Then
                MsgBox ErrorMessage, vbInformation, Me.Caption
                ErrorMessage = ""
            End If
    End If
    
    If capp = 1 Then
    
    If Not Save_GLTRANSACTION(Format(dtpTransDate, "dd/mm/yyyy"), (CCur(crate) * CCur(txtDispatch)), cessdr, cesscr, "Cess Deductions ", txtRefNo, User, ErrorMessage, "Cess Deductions", 1, 1, txtRefNo, transactionNo, "", "", 0) Then
            If ErrorMessage <> "" Then
                MsgBox ErrorMessage, vbInformation, Me.Caption
                ErrorMessage = ""
            End If
    End If
    
    End If
        
'd_sp_MilkControl @DispDate char(10), @DipsQnty float,@DipQnty float,@InQnty float,@VarQnty float,@Price char(10),@RefNo varchar(35),@CreditAcc varchar(35),@DebitAcc varchar(35),@AuditID varchar (50)
Set rs = New ADODB.Recordset
sql = "d_sp_MilkControl  '" & dtpTransDate & "'," & txtDispatch & "," & txtDipping & "," & txtIntake & "," & txtVariance & "," & Price & ",'" & txtRefNo & "','" & Credit & "','" & Debit & "','" & User & "','" & txtdcode & "','" & txtvehicleno & "'"
oSaccoMaster.ExecuteThis (sql)

'//subtract from the dispatch table

    sql = ""
    sql = "SET      dateformat dmy     SELECT     ID, Intake,transdate     FROM         d_dispatch    WHERE     transdate = '" & dtpTransDate & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If rs.EOF Then
sql = ""
sql = "set dateformat dmy INSERT INTO d_dispatch (Transdate, descrip, Intake, dipping, dispatch, auditid, auditdate)values ('" & dtpTransDate.value & "','Dispatch'," & txtDipping & "," & CDbl(txtDipping) - CDbl(txtDispatch) & "," & CDbl(txtDispatch) & ",'" & User & "','" & Get_Server_Date & "')"
oSaccoMaster.ExecuteThis (sql)
'sql = ""
'sql = "set dateformat dmy INSERT INTO d_DetailDispatch (DCode,Name, Transdate, dispatch, auditid,auditdate)values ('" & txtdcode & "','','" & dtpTransDate.value & "','" & txtDispatch & "','" & User & "','" & Get_Server_Date & "')"
'oSaccoMaster.ExecuteThis (sql)

'    sql = ""
'    sql = "SET      dateformat dmy     SELECT DCode, DName FROM d_Debtors    WHERE     DCode = " & txtdcode & ""
'    Set rsx = oSaccoMaster.GetRecordset(sql)
'    If rsx.EOF Then
'    sql = ""
'    sql = "set dateformat dmy INSERT INTO d_DetailDispatch (DName)values ('" & dtpTransDate.value & "','Dispatch'," & txtDipping & "," & CDbl(txtDipping) - CDbl(txtDispatch) & "," & CDbl(txtDispatch) & ",'" & User & "','" & Get_Server_Date & "')"
'    oSaccoMaster.ExecuteThis (sql)
'
'    End If

Else
sql = ""
sql = "set dateformat dmy UPDATE    d_dispatch  SET   dipping =" & CDbl(txtDipping) - CDbl(txtDispatch) & ",dispatch=" & txtDispatch & "  WHERE     (Transdate = '" & dtpTransDate & "')"
oSaccoMaster.ExecuteThis (sql)
'sql = ""
'sql = "set dateformat dmy UPDATE    d_DetailDispatch  SET   dispatch =" & txtDispatch & "  WHERE     (Transdate = '" & dtpTransDate & "') and DCode='" & txtdcode & "'"
'oSaccoMaster.ExecuteThis (sql)
End If
'Dim rsd As Recordset
'sql = ""
'sql = "select dispatch1, dispatch2, dispatch3, dispatch4, dispatch5  from d_DetailDispatch where Transdate='" & dtpTransDate & "' "
'
'Set rsd = oSaccoMaster.GetRecordset(sql)
'
'Dim DName As Double, two As Double, three As Double, four As Double, five As Double
'one = rsd.Fields(0)
'two = rsd!dispatch2
'three = rsd!dispatch3
'four = rsd!dispatch4
'five = rsd!dispatch5
'If one = "" Then
         Dim DName As String
          Set rs = New ADODB.Recordset
          sql = "SELECT DName from d_Debtors where DCode='" & txtdcode & "'"
          Set rs = oSaccoMaster.GetRecordset(sql)
          If Not rs.EOF Then
          DName = rs!DName
          End If

'sql = ""
'sql = "set dateformat dmy INSERT INTO d_DetailDispatch (DCode, Transdate,Name, dispatch, auditid, auditdate)values ('" & txtdcode & "','" & dtpTransDate.value & "','" & DName & "','" & txtDispatch & "','" & User & "','" & Get_Server_Date & "')"
'oSaccoMaster.ExecuteThis (sql)
''Else
'..............INSERT DAILY INTAKE FOR DEBTORS ONLY.........................
'Dim rsd As Recordset
'  sql = ""
'  sql = "set dateformat DMY select  DCode, Transdate,Name, dispatch, auditid, auditdate from d_DetailDispatch where DCode= 'Intake' and dispatch='" & txtIntake & "' and Transdate='" & dtpTransDate & "' "
'  Set rsd = New ADODB.Recordset
'  rsd.Open sql, cn
'  If rsd.EOF Then
'    sql = ""
'    sql = "set dateformat dmy INSERT INTO d_DetailDispatch (DCode, Transdate,Name, dispatch, auditid, auditdate)values ('Intake','" & dtpTransDate.value & "','Intake','" & txtIntake & "','" & User & "','" & Get_Server_Date & "')"
'    oSaccoMaster.ExecuteThis (sql)
'  Else
'  sql = ""
'  sql = "set dateformat dmy UPDATE d_DetailDispatch SET dispatch='" & txtIntake & "'  WHERE DCode='Intake' and   (Transdate = '" & dtpTransDate & "')"
'  oSaccoMaster.ExecuteThis (sql)
'  End If
'..............END OF  DAILY INTAKE INSERT FOR DEBTORS ONLY.........................
mysql = "set dateformat dmy Insert into Receiptno(Receiptno,Auditdate,auditid)values('" & txtRefNo & "','" & Format(Get_Server_Date, "dd/MM/yyyy") & "','" & User & "')"
oSaccoMaster.ExecuteThis (mysql)
If chkPrint = vbChecked Then
    
If chkPrint = vbChecked Then
    
'/*Print out
 Dim fso, chkPrinter, txtfile
    Dim ttt
     Dim escFeedAndCut As String
     Dim escNewLine As String
     Dim escUnerLineON As String
     Dim escUnerLineOnX2 As String
     Dim escUnerLineOff As String
     Dim escBoldOn As String
     Dim escBoldOff As String
     Dim escNegativeOn As String
     Dim escNegativeOff As String
     Dim esc8CpiOn As String
     Dim esc8CPiOff As String
     Dim esc16Cpi As String
     Dim esc20Cpi As String
     Dim escAlignLeft As String
     Dim escAlignCenter As String
     Dim escAlignRight As String
    
     
        escNewLine = Chr(10) '//New Line (LF Line Feed)
        escUnerLineON = Chr(27) + Chr(45) + Chr(1) '//Unerline on
        escUnerLineOnX2 = Chr(27) + Chr(45) + Chr(1) '//Unerline on X2
        escUnerLineOff = Chr(27) + Chr(45) + Chr(0) '//unerline off
        escBoldOn = Chr(27) + Chr(69) + Chr(1) '//Bold on
        escBoldOff = Chr(27) + Chr(69) + Chr(0) '//Bold off
        escNegativeOn = Chr(29) + Chr(66) + Chr(1) '//White on Black on
        escNegativeOff = Chr(29) + Chr(66) + Chr(0) '//white on
        esc8CpiOn = Chr(29) + Chr(33) + Chr(16) '//Font Size X2 on
        esc8CPiOff = Chr(29) + Chr(33) + Chr(0) '//Font size X2 off
        esc16Cpi = Chr(27) + Chr(77) + Chr(48) '//Font A - Normal Size
        esc20Cpi = Chr(27) + Chr(77) + Chr(49) '//Font B - Small Font
        escAlignLeft = Chr(27) + Chr(97) + Chr(48) '//Align text to the left
        escAlignCenter = Chr(27) + Chr(97) + Chr(49) '//Align text to the center
        escAlignRight = Chr(27) + Chr(97) + Chr(50) '//Align text to the right
        escFeedAndCut = Chr(29) + Chr(86) + Chr(65) '//Partial cut and feed
       'cdgPrint.Filter = "*.csv|*.txt"
        'cdgPrint.ShowSave
        Dim PORT As String
   '     PORT = ports
        'ttt = "LPT1" 'LPT1
        ttt = ports
        'ttt = "D:\PROJECTS\FOSA\DAILY" & Date & ""
        Set fso = CreateObject("Scripting.FileSystemObject")
        'Set chkPrinter = fso.GetFile(ttt)
        
    Set txtfile = fso.CreateTextFile(ttt, True)
    txtfile.WriteLine "         " & cname & ""
    txtfile.WriteLine "         Address :" & paddress & ""
    txtfile.WriteLine "         Phone :" & Phone & ""
    txtfile.WriteLine "         Email :" & Email & ""
    'txtfile.WriteLine " " & txtSNo
    
    txtfile.WriteLine "          Delivery Note"
    txtfile.WriteLine "**********************************************"
        
    Set rs2 = New ADODB.Recordset
    sql = "d_sp_ReceiptNumber"
    Set rs2 = oSaccoMaster.GetRecordset(sql)
    
    Dim RNumber As String
    'RNumber = rs2.Fields(0)
    If Not IsNull(rs2.Fields(0)) Then RNumber = rs2.Fields(0)
    'Else
    'RNumber = "0"
    'End If
    
    txtfile.WriteLine "CsNO :" & txtRefNo
    txtfile.WriteLine "To :" & lblDebtors
   txtfile.WriteLine " *********************************************************************"
    txtfile.WriteLine "DESCRIPTION " & vbTab & "" & vbTab & "value"
    sql = "SELECT     d.DCode, d.DName, SUM(m.DispQnty) AS quantity FROM         d_Debtors AS d INNER JOIN d_MilkControl AS m ON d.DCode = m.dcode WHERE     (DispDate = '" & dtpTransDate & "') GROUP BY d.DCode, d.DName"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
    'txtamountp = rs!quantity*
   ' Dim milksup As Double
'    Dim amtpayable As Double
'    Dim receipno As Double
'    Dim dispatchby As Double
   ' Exit Sub
   ' End If
'    Set rs = New ADODB.Recordset
'    sql = "d_sp_TotalMonth " & txtSNo & ",'" & Startdate & "','" & DTPMilkDate & "'"
'    Set rs = oSaccoMaster.GetRecordset(sql)
'    If Not rs.EOF Then
'    If Not IsNull(rs.Fields(0)) Then CummulKgs = rs.Fields(0)
'    Else
'    CummulKgs = "0.00"
'    End If
    txtfile.WriteLine "Milk supplied :" & vbTab & "" & vbTab & " " & rs!Quantity & ""
    txtfile.WriteLine "Amount Payable :" & vbTab & "  " & txtamountp
    txtfile.WriteLine "Receipt Number :" & vbTab & "  " & txtRefNo
    txtfile.WriteLine "Dispatched by :" & vbTab & " " & username & ""
    
    txtfile.WriteLine "---------------------------------------"
    End If
'    txtFile.WriteLine "Receipt Number :" & RNumber
'    txtFile.WriteLine "TRANSPORTER :" & TRANSPORTER
    txtfile.WriteLine "Vehicle No :" & txtvehicleno
    txtfile.WriteLine "Received by :" & txtreceiveby
    txtfile.WriteLine "---------------------------------------"
    txtfile.WriteLine "     Date :" & Format(dtpTransDate, "dd/mm/yyyy") & " ,Time : " & Format(Time, "hh:mm:ss AM/PM")
    txtfile.WriteLine "" & motto & ""
    txtfile.WriteLine "---------------------------------------"
    'If chkComment.value = vbChecked Then dtpTransDate
        'txtFile.WriteLine txtComment
        txtfile.WriteLine "---------------------------------------"
        txtfile.WriteLine "********POWERED BY EASYMA***************"
    'End If
    txtfile.WriteLine escFeedAndCut
    
 txtfile.Close
 Reset
End If
txtdcode = ""
txtDispatch = ""
txtIntake = ""
txtDipping = ""
txtRefNo = ""


'* writing to notepad

'If chkNotepad.value = vbChecked Then

'    Dim fso, txtfile
'    Dim ttt
'     Dim escFeedAndCut As String
'     escFeedAndCut = Chr(29) + Chr(86) + Chr(65) '//Partial cut and feed
'       cdgPrint.Filter = "*.csv|*.txt"
'        cdgPrint.ShowSave
'        ttt = cdgPrint.FileName
'        If ttt = "" Then
'        MsgBox "File should not be blank", vbCritical, "Data transfer"
'        Exit Sub
'        End If
'        'ttt = "D:\PROJECTS\FOSA\DAILY" & Date & ""
'        Set fso = CreateObject("Scripting.FileSystemObject")
'        Set txtFile = fso.CreateTextFile(ttt, True)
'        txtFile.WriteLine
'
'    txtFile.WriteLine "---------------------------------------"
'    txtFile.WriteLine "" & cname & ""
'    txtFile.WriteLine " " & paddress & ""
'    txtFile.WriteLine " " & Phone & ""
'   ' Printer.Print Tab(0); "Kimathi House Branch"
'    txtFile.WriteLine " " & paddress & " "
'    txtFile.WriteLine "" & town & ""
'    txtFile.WriteLine "Milk Receipt"
'    txtFile.WriteLine "---------------------------------------"
''    If cbomemtrans = "Shares" Then
''    DESC = bosanames & " -Member No " & memberno
'    txtFile.WriteLine "SNo :" & txtSNo
'    txtFile.WriteLine "Name :" & lblNames
''    Else
'    txtFile.WriteLine "Quantity Supplied :" & txtQnty & " Kgs"
'    Startdate = DateSerial(year(DTPMilkDate), month(DTPMilkDate) - 1, 1)
'    'sql = "d_sp_TotalMonth " & txtSNo & ",'" & StartDate & "','" & DTPMilkDate & "'"
'    Set rs = New ADODB.Recordset
'    sql = "d_sp_TotalMonth " & txtSNo & ",'" & Startdate & "','" & DTPMilkDate & "'"
'    Set rs = oSaccoMaster.GetRecordset(sql)
'    If Not rs.EOF Then
'    If Not IsNull(rs.Fields(0)) Then CummulKgs = rs.Fields(0)
'    Else
'    CummulKgs = "0.00"
'    End If
'    txtFile.WriteLine "Cummulative This Month " & Format(CummulKgs, "#,##0.00" & " Kgs")
''    End If
'    Set rs = New ADODB.Recordset
'    sql = "d_sp_TransName '" & txtSNo & "'"
'    Set rs = oSaccoMaster.GetRecordset(sql)
'    If Not rs.EOF Then
'    If Not IsNull(rs.Fields(0)) Then TRANSPORTER = rs.Fields(0)
'    Else
'    TRANSPORTER = "Self"
'    End If
'    txtFile.WriteLine "---------------------------------------"
'    txtFile.WriteLine "Transporter :" & TRANSPORTER
'    txtFile.WriteLine "Received by :" & username
'    txtFile.WriteLine "---------------------------------------"
'    txtFile.WriteLine "Date :" & Format(Get_Server_Date, "dd/mm/yyyy HH:MM:SS AM/PM")
'    txtFile.WriteLine "     " & motto & " "
'    txtFile.WriteLine "---------------------------------------"
'    txtFile.WriteLine escFeedAndCut
'
'txtFile.Close








End If

MsgBox "Records saved successifully."
'Exit Sub





'//PRINT THE REPORT HERE
'milkinvoice

'd_MilkControl."RefNo"
'    STRFORMULA = "{d_MilkControl.RefNo}='" & txtRefNo & "'"
'    reportname = "milkinvoice.rpt"
'    Show_Sales_Crystal_Report STRFORMULA, reportname, title
    Form_Load
    Exit Sub
ErrorHandler:
        
        MsgBox err.description
End Sub

Private Sub cmdstatement_Click()
    reportname = "d_DebtorsInvoice.rpt"
    Show_Sales_Crystal_Report "", reportname, ""
End Sub

Private Sub cmdupdatereceipt_Click(Index As Integer)
     On Error GoTo SysError
    Dim j As Integer
    Dim amt As Double, interest As Double
    Dim code As String, rsLoan As New Recordset
    Dim k As Long, repaymethod As String
    Dim sharebals As Double, amount As Double
    Dim ptype As String, principal As Double, chequeno As String
    Dim intWaive As Boolean
    'Recalculate
    If lvwNtrans(Index).ListItems.Count = 0 Then
        Exit Sub
    End If
    
    If txtReceiptsno(Index) = "" Then
        MsgBox "Please enter the receiptno", vbCritical
        Exit Sub
    End If
    If txtAmountPaid(Index) <= 0 Then
        MsgBox "Amount should be greater than zero", vbCritical
        Exit Sub
    End If
    If txtAmountPaid(Index) = "" Then
        MsgBox "Amount should be have a figure on it", vbCritical
        Exit Sub
    End If
    If cboBanks(Index).Text = "" Then
        MsgBox ("You Must Select the Bank Control Account"), vbCritical
        Exit Sub
    End If
    If cboMode(Index) = "Cheque" Then
        If txtmode(Index) = "" Then
            MsgBox ("Cheque Receipt No Required"), vbInformation
            Exit Sub
        End If
    End If
    If cboMode(Index) = "Cash" Then
        If txtReceiptsno(Index) = "" Then
            MsgBox ("Cash Receipt No Required"), vbInformation
            Exit Sub
        End If
    End If
    If cboMode(Index) = "EFT" Then
        If txtmode(Index) = "" Then
            MsgBox ("EFT Receipt No Required"), vbInformation
            Exit Sub
        End If
    End If
    If cboMode(Index) = "Mpesa" Then
        If txtmode(Index) = "" Then
            MsgBox ("Mpesa Receipt No Required"), vbInformation
            Exit Sub
        End If
    End If
    If cboMode(Index) = "Zap" Then
        If txtmode(Index) = "" Then
            MsgBox ("Zap Receipt No Required"), vbInformation
            Exit Sub
        End If
    End If
    If CDbl(txtBalance(Index)) <> 0 Then
        MsgBox "The Amount Received should be equal to the Amount Distributed", vbInformation, Me.Caption
        Exit Sub
    End If
    

    
    If Index = 0 Then
        sql = "select receiptno,chequeno from receiptbooking where receiptno in ('" & txtReceiptsno(Index) & "')"
        Set rst = oSaccoMaster.GetRecordset(sql)
        If Not rst.EOF Then
            MsgBox "Either the Receiptno or Chequeno is already used!", vbCritical
            Exit Sub
        End If
        
        If MsgBox("Do you want to post this Receipt: " & txtReceiptsno(0) & "?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
            Exit Sub
        End If
    
    Else
        sql = "select voucherno,chequeno from paymentbooking where voucherno in ('" & txtReceiptsno(Index) & "')"
        Set rst = oSaccoMaster.GetRecordset(sql)
        If Not rst.EOF Then
            MsgBox "Either the VoucherNo or Chequeno is already used!", vbCritical
            Exit Sub
        End If
        
        If MsgBox("Do you want to post this Voucher: " & txtReceiptsno(1) & "?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
            Exit Sub
        End If
        
    End If
    ''...................insert the amount to debtor if available................................
      If chkDebtno.value = vbChecked Then
       Dim Amount1 As Integer
       Set rs = New ADODB.Recordset
       sql = ""
       sql = "SET dateformat dmy Select Amount  from d_MilkControl  where DCode ='" & txtTCode & "' and DispDate='" & dtpTransDate(Index) & "'"
       Set rs = oSaccoMaster.GetRecordset(sql)
       
       If Not rs.EOF Then
'        sql = ""
'        sql = "set dateformat dmy insert into  d_MilkControl(Amount) values('" & CDbl(txtDistributed(Index)) & "') where DCode ='" & txtTCode & "' and DispDate='" & DTPTransdate(Index) & "'"
'        oSaccoMaster.ExecuteThis (sql)
'
'       Else
         sql = ""
         sql = "set dateformat DMY update d_MilkControl set Amount=" & rs.Fields("amount") + CDbl(txtDistributed(Index)) & " where DCode ='" & txtTCode & "' and DispDate='" & dtpTransDate(Index) & "' "
         oSaccoMaster.ExecuteThis (sql)
       End If
     Else
     End If
    
    '''..................end of debtor...........................................................
    
    
    Select Case txtmode(Index).Text
        Case "Cheque"
            chequeno = txtmode(Index)
        Case Else
            chequeno = txtmode(Index).Text
    End Select
    
    
    
    '***********************************BEGINNING OF POSTING**************************************

    ref = "Nominal"
    RefNo = ""

    
    With lvwNtrans(Index)
        If .ListItems.Count < 0 Then
            Exit Sub
        End If
        I = 0
        j = 0
        I = .ListItems.Count
        
        oSaccoMaster.goConn.BeginTrans
        On Error GoTo TransError
        'save TransactionNo
        transactionTotal = CDbl(txtAmountPaid(Index).Text)
        NewTransaction transactionTotal, dtpTransDate(Index), "ReceiptPosting -Receiptno " & txtReceiptsno(Index).Text
            
            Dim fperiod As Integer, intRate As Double

            If Index = 0 Then
                If Not saveReceipt(txtReceiptsno(Index), ref, RefNo, cboBanks(Index).Text, cboBanks(Index).Text, txtParticulars(Index), dtpTransDate(Index).value, CDbl(txtAmountPaid(Index).Text), txtmode(Index), cboMode(Index).Text, "Deposit") Then
                    GoTo TransError
                Else
                    For I = 1 To lvwNtrans(Index).ListItems.Count
                    
                        If Save_GLTRANSACTION(dtpTransDate(Index), lvwNtrans(Index).ListItems(I).SubItems(2), cboBanks(Index), lvwNtrans(Index).ListItems(I), _
                        txtReceiptsno(Index), mMemberNo, User, ErrorMessage, txtParticulars(Index) + "-" + Replace(lblSupplier(Index), ",", ""), 0, 1, txtmode(Index), transactionNo) = False Then
                            GoTo TransError
                        End If

                    Next I
                    
                    'Invoice
                    
                    If InvoiceNo <> "" And InvoiceBal > 0 And chkCreditors(0).value = 1 Then
                    
                        sql = "update SalesOrder set balance=balance-" & txtAmountPaid(1) & " where orderno='" & InvoiceNo & "'"
                        If Not oSaccoMaster.Execute(sql) Then
                            GoTo TransError
                        End If
                    
                        'Statement
                        
                        sql = "insert into CustomerStmt (invId,TransDate,Refno,Amount,TransType,Balance,Auditid,Remarks,Transactionno) " _
                        & " Values ('" & InvoiceNo & "','" & dtpTransDate(Index) & "','" & txtReceiptsno(Index) & "'," & txtAmountPaid(Index) & ",'CR',0,'" & User & "','" & txtParticulars(Index) & "','" & transactionNo & "')"
                        
                        If Not oSaccoMaster.Execute(sql) Then
                            GoTo TransError
                        End If
                        
                    End If
                    
                    
                End If
            Else
            
            
                    For I = 1 To lvwNtrans(Index).ListItems.Count
                    
                        If Save_GLTRANSACTION(dtpTransDate(Index), lvwNtrans(Index).ListItems(I).SubItems(2), lvwNtrans(Index).ListItems(I), cboBanks(Index), txtReceiptsno(Index), mMemberNo, User, ErrorMessage, txtParticulars(Index), 1, 1, txtmode(Index), TransNo) = False Then

                            GoTo TransError
                        End If
'           If Save_GLTRANSACTION(dtpTransDate(index), lvwNtrans(index).ListItems(I).SubItems(2), cboBanks(index), lvwNtrans(index).ListItems(I), _
'                        txtReceiptsno(index), mMemberNo, User, ErrorMessage, txtParticulars(index) + "-" + Replace(lblSupplier(index), ",", ""), 0, 1, txtmode(index), transactionNo) = False Then
'                            GoTo TransError
'                        End If
                    Next I
                    'Save the Voucher

                    sql = "set dateformat dmy INSERT INTO PaymentBooking (VoucherNo,Memberno,PayeeDesc,Ccode,Name,Transdate," _
                    & "Amount, Chequeno, Ptype, auditid,Transactionno) VALUES ('" & txtReceiptsno(Index) & "','" & _
                    cboBanks(Index) & "','" & lblbankname(Index) & "','" & cboBanks(Index) & "','" & lblbankname(Index).Caption & "','" & dtpTransDate(Index) & "'," & CDbl(txtDistributed(Index)) & ",'" & _
                    txtmode(Index).Text & "','" & cboMode(Index).Text & "','" & User & "','" & transactionNo & "')"
                    
                    If Not oSaccoMaster.Execute(sql) Then
                        GoTo TransError
                    End If
                    
                    
                    'Invoice
                    
                    If InvoiceNo <> "" And InvoiceBal > 0 And chkCreditors(Index).value = 1 Then
                    
                        sql = "update d_Invoice set balance=balance-" & txtAmountPaid(1) & " where invid='" & InvoiceNo & "'"
                        If Not oSaccoMaster.Execute(sql) Then
                            GoTo TransError
                        End If
                    
                        'Statement
                        
                        sql = "insert into supplierStmt (invId,TransDate,Refno,Amount,TransType,Balance,Auditid,Remarks,Transactionno) " _
                        & " Values ('" & InvoiceNo & "','" & dtpTransDate(Index) & "','" & txtReceiptsno(Index) & "'," & txtAmountPaid(Index) & ",'DR',0,'" & User & "','" & txtParticulars(Index) & "','" & transactionNo & "')"
                        
                        If Not oSaccoMaster.Execute(sql) Then
                            GoTo TransError
                        End If
                        
                    End If
                    
                        
                    
                    
            End If
        End With
    oSaccoMaster.goConn.CommitTrans
    
    If MsgBox("Receipt updated successfully, Print Receipt?", vbQuestion + vbYesNo) = vbYes Then
        PrintReceipt Index
    End If
    
    chkDebtno.value = vbUnchecked
    
    lvwNtrans(Index).ListItems.Clear
    
    txtReceiptsno(0) = NewRefno(0) 'Generate_ReceiptNo("Receipt")
    txtReceiptsno(1) = NewRefno(1) 'Generate_ReceiptNo("Voucher")
    
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



Private Sub cmdClear_Click(Index As Integer)
    On Error Resume Next
    Dim vote As String
    'Dim msgb As Response
    Dim rsponse As String

        
    With lvwNtrans(Index)
        If .ListItems.Count > 0 Then
            If MsgBox("Are you sure you want to clear the entire list?", vbQuestion + vbYesNo) = vbYes Then
                .ListItems.Clear
            End If
        End If
    End With

    Recalculate (Index)
        
End Sub

Private Sub dtpTransDate_CallbackKeyDown(Index As Integer, ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
    Set rs = New ADODB.Recordset
    sql = ""
    Set rs = New ADODB.Recordset
    sql = "d_sp_DailyTotal '" & dtpTransDate & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not IsNull(rs.Fields(0)) Then
    txtIntake = Format(rs.Fields(0), "#0.00")
    'txtDipping = txtIntake
End Sub

Private Sub Form_Load()
    On Error Resume Next
    dtpTransDate(0) = Format(Get_Server_Date, "dd/mm/yyyy")
    dtpTransDate(1) = Format(Get_Server_Date, "dd/mm/yyyy")
    
    Dim rscompany As New ADODB.Recordset
    InvoiceBal = 0

    'Load Gl's
    sql = "Select accno from glsetup order by accno asc"
    Set rst = oSaccoMaster.GetRecordset(sql)
    While Not rst.EOF
        cboAccno(0).AddItem (rst(0))
        cboAccno(1).AddItem (rst(0))
        rst.MoveNext
    Wend
    'load the banks
    cboBanks(0).Clear
    cboBanks(1).Clear
    sql = "select Accno from banks where accno not in (select assigngl from useraccounts where UserLoginIDs not in('" & User & "') and assigngl<>'' and assigngl is not null )"
    Set rst = oSaccoMaster.GetRecordset(sql)
    While Not rst.EOF
        cboBanks(0).AddItem rst(0)
        cboBanks(1).AddItem rst(0)
        rst.MoveNext
    Wend
    
    cboBanks(0).List(0) = currentUser.tellerGlAcc
    cboBanks(0).Text = cboBanks(0).List(0)
    cboBanks(1).List(0) = currentUser.tellerGlAcc
    cboBanks(1).Text = cboBanks(1).List(0)
    
   
    Totalamount = 0
    pushed = 0
    

    
    txtReceiptsno(0) = NewRefno(0) 'Generate_ReceiptNo("Receipt")
    txtReceiptsno(1) = NewRefno(1) 'Generate_ReceiptNo("Voucher")
 
'    lvwJuniors.Visible = False
    IsBatch = False
    commit = False 'will be used to determine whether a loaded interest is saved as a transaction
End Sub

Public Function NewRefno(Index As Integer) As String
    
    Dim Rno As String
    Dim prefix As String
    Dim rcount As Integer
    If Index = 0 Then
        sql = "select count(distinct receiptno)ccount from receiptbooking"
        prefix = "MCR-"
    Else
        sql = "select count(distinct voucherno)ccount from paymentbooking"
        prefix = "MCV-"
    End If
        
    
    Set rst = oSaccoMaster.GetRecordset(sql)
    If Not rst.EOF Then
        rcount = rst(0) + 1
    Else
        rcount = 1
    End If
    
    'thisday = Get_Server_Date
    
    Rno = Format(CStr(rcount), "000000")
    
    
    NewRefno = prefix & "-" & CStr(Rno)
    
End Function
Private Sub Form_Unload(Cancel As Integer)
    'Stop subclassing
    CloseSubClass
    'Clean up by setting the classes to Nothing
    'Set objLabelEdit = Nothing
    'Set objLabelEdit2 = Nothing
End Sub






Private Sub Label2_Click()

End Sub

Private Sub listview1_DblClick()
frmmilkdidprev.txtdcode = ListView1.SelectedItem
frmmilkdidprev.txtdesc = ListView1.SelectedItem.ListSubItems(1)
frmmilkdidprev.txtquantity = ListView1.SelectedItem.ListSubItems(2)
Dim q As Double
frmmilkdidprev.Show vbModal
End Sub

Private Sub lvwNTrans_DblClick(Index As Integer)
    Dim total As Double, amt As Double
    Dim ccount As Integer
    On Error Resume Next
    total = 0
    With lvwNtrans(Index)
        If .ListItems.Count > 0 Then
            amt = InputBox("Enter the amount", "AMOUNT", .SelectedItem.ListSubItems(2))
            .SelectedItem.ListSubItems(2) = amt
        End If
    End With
    Recalculate (Index)
End Sub









Private Sub Recalculate(Index As Integer)
    On Error Resume Next
    Dim balance As Double
    Dim thislistview As ListView
    
    If lvwNtrans(Index).ListItems.Count > 0 Then
        For I = 1 To lvwNtrans(Index).ListItems.Count
            balance = balance + CDbl(lvwNtrans(Index).ListItems(I).SubItems(2))
        Next I
    End If

    txtDistributed(Index) = Format(balance, Cfmt)
    txtBalance(Index) = Format(CDbl(txtAmountPaid(Index)) - CDbl(txtDistributed(Index)), Cfmt)
End Sub


Private Sub lvwNtrans_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
      lvwNTrans_DblClick Index
    End If
End Sub

Private Sub mnuCustStatements_Click()
    reportname = "CustomerStatement.rpt"
    Show_Sales_Crystal_Report "", reportname, CompanyName
End Sub

Private Sub mnuPurDayBook_Click()
    reportname = "PeriodicPurchases.rpt"
    Show_Sales_Crystal_Report "", reportname, CompanyName
End Sub

Private Sub mnureceiptslist_Click()
  reportname = "cashandbanklists.rpt"
    Show_Sales_Crystal_Report "", reportname, CompanyName
End Sub

Private Sub mnuSalesDayBook_Click()
    reportname = "PeriodicSales.rpt"
    Show_Sales_Crystal_Report "", reportname, CompanyName
End Sub

Private Sub mnuSupStatement_Click()
    reportname = "SupplierStatement.rpt"
    Show_Sales_Crystal_Report "", reportname, CompanyName
End Sub

Private Sub Picture1_Click()
Me.MousePointer = vbHourglass
         frmSearchMilkControl.Show vbModal
        txtRefNo = sel
        txtRefNo_Validate True
        Me.MousePointer = 0
End Sub

Private Sub Picture3_Click()
Me.MousePointer = vbHourglass
         frmSearchDebtors.Show vbModal
        txtdcode = sel
        txtdcode_Validate True
        Me.MousePointer = 0
End Sub

Private Sub Picture5_Click()
Me.MousePointer = vbHourglass
        frmSearchDebtors.Show vbModal
        txtTCode = sel
        txtTCode_Validate True
        Me.MousePointer = 0
End Sub

Private Sub txtamountp_Change()

End Sub

Private Sub txtAmountPaid_Change(Index As Integer)
On Error Resume Next
    If txtAmountPaid(Index).Text = "" Then txtAmountPaid(Index).Text = 0
    Totalamount = CDbl(txtAmountPaid(Index).Text)
    pushed = 0
    txtBalance(Index).Text = Totalamount - CDbl(txtDistributed(Index).Text)
End Sub




Private Sub txtAmountPaid_KeyPress(Index As Integer, KeyAscii As Integer)
    If keyIsValid(KeyAscii, 1) = False Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub txtdcode_Validate(Cancel As Boolean)
Set rs = oSaccoMaster.GetRecordset("SELECT dname,Price,accdr,acccr,drcess,crcess,capp,crate FROM d_Debtors WHERE DCode = '" & txtdcode & "'")
If rs.RecordCount > 0 Then
If Not IsNull(rs.Fields(1)) Then Price = rs.Fields(1)
If Not IsNull(rs.Fields(0)) Then lblDebtors = rs.Fields(0)
If Not IsNull(rs.Fields(2)) Then Label10 = rs.Fields(2)
If Not IsNull(rs.Fields(3)) Then Label11 = rs.Fields(3)
If Not IsNull(rs.Fields(4)) Then cessdr = rs.Fields(4)
If Not IsNull(rs.Fields(5)) Then cesscr = rs.Fields(5)
If Not IsNull(rs.Fields(6)) Then capp = Abs(rs.Fields(6))
If Not IsNull(rs.Fields(7)) Then crate = rs.Fields(7)
txtamountp = txtDispatch * rs.Fields(1)
If capp = 1 Then
chkapp = vbChecked
Else
chkapp = vbUnchecked
End If
Else
lblDebtors = ""
End If
End Sub

Private Sub txtDispatch_Change()
'txtDipping = txtDispatch
If txtDispatch = "" Then
txtDispatch = "0"
End If
'If txtDipping = "" Then
'txtDipping = "0"
'End If

'**************PRICE***************'
Set rs = oSaccoMaster.GetRecordset("SELECT dname,Price,accdr,acccr,drcess,crcess,capp,crate FROM d_Debtors WHERE DCode = '" & txtdcode & "'")
If rs.RecordCount > 0 Then
If Not IsNull(rs.Fields(1)) Then Price = rs.Fields(1)
If Not IsNull(rs.Fields(0)) Then lblDebtors = rs.Fields(0)
If Not IsNull(rs.Fields(2)) Then Label10 = rs.Fields(2)
If Not IsNull(rs.Fields(3)) Then Label11 = rs.Fields(3)
If Not IsNull(rs.Fields(4)) Then cessdr = rs.Fields(4)
If Not IsNull(rs.Fields(5)) Then cesscr = rs.Fields(5)
If Not IsNull(rs.Fields(6)) Then capp = Abs(rs.Fields(6))
If Not IsNull(rs.Fields(7)) Then crate = rs.Fields(7)
txtamountp = txtDispatch * rs.Fields(1)
If capp = 1 Then
chkapp = vbChecked
Else
chkapp = vbUnchecked
End If
Else
lblDebtors = ""
End If


'****************END********************'





txtVariance = Format(txtIntake - txtDispatch, "#0.00")

End Sub

Private Sub txtIntake_Change()
txtDispatch_Change
End Sub

Private Sub txtRefNo_Change()
On Error GoTo ErrorHandler
'SELECT TOP 300 DispDate,RefNo,dcode,DispQnty,Price,InQnty,Variance FROM dbo.d_MilkControl"
If Trim(txtRefNo) = "" Then
Exit Sub
End If
 Set rs = oSaccoMaster.GetRecordset("SELECT DispDate,dcode,DispQnty,Price,InQnty,Variance FROM d_MilkControl WHERE RefNo = '" & txtRefNo & "'")
 
 If rs.RecordCount > 0 Then
    dtpTransDate = rs.Fields(0)
    txtDispatch = rs.Fields(2)
    txtDipping = txtDispatch
    txtIntake = rs.Fields(4)
    txtVariance = rs.Fields(5)
    txtdcode = rs.Fields(1)
    
    cmdEdit.Enabled = True
Else
    cmdEdit.Enabled = False
    
End If
txtdcode_Validate True
Exit Sub
ErrorHandler:
MsgBox err.description
End Sub

Private Sub txtTCode_Validate(Cancel As Boolean)
Dim a As Boolean, b As Integer
Set rs = New ADODB.Recordset
sql = "d_sp_Selectdebtors '" & txtTCode & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
If Not IsNull(rs.Fields(0)) Then txtDNames = rs.Fields(0)
'If Not IsNull(rs.Fields(1)) Then txtId = rs.Fields(1)
'If Not IsNull(rs.Fields(2)) Then cboLocation = rs.Fields(2)
'If Not IsNull(rs.Fields(3)) Then DTPRegDate = rs.Fields(3)
'If Not IsNull(rs.Fields(4)) Then txtEMail = rs.Fields(4)
'If Not IsNull(rs.Fields(5)) Then txtPhone = rs.Fields(5)
'If Not IsNull(rs.Fields(6)) Then txtTown = rs.Fields(6)
'If Not IsNull(rs.Fields(7)) Then txtPAddress = rs.Fields(7)
'If Not IsNull(rs.Fields(8)) Then txtsubsidy = Format(rs.Fields(8), "#0.00")
'If Not IsNull(rs.Fields(9)) Then Txtaccno = rs.Fields(9)
'If Not IsNull(rs.Fields(10)) Then cboBName = rs.Fields(10)
'If Not IsNull(rs.Fields(11)) Then cboBBranch = rs.Fields(11)
'If Not IsNull(rs.Fields(12)) Then a = rs.Fields(12)
'If Not IsNull(rs.Fields(13)) Then cbobranch = rs.Fields(13)
'If Not IsNull(rs.Fields(14)) Then txtPrice = Format(rs.Fields(14), "#0.00")
'If Not IsNull(rs.Fields(15)) Then txtDrAccNo = rs.Fields(15)
'If Not IsNull(rs.Fields(16)) Then txtCrAccNo = rs.Fields(16)
'If Not IsNull(rs.Fields(17)) Then txtcessrate = rs.Fields(17)
'If Not IsNull(rs.Fields(18)) Then txtcessdebit = rs.Fields(18)
'If Not IsNull(rs.Fields(19)) Then txtcesscredit = rs.Fields(19)
'If Not IsNull(rs.Fields(20)) Then b = rs.Fields(20)
'If b = 1 Then
'chkcessapp = vbChecked
'Else
'chkcessapp = vbUnchecked
'
'End If
'If a = True Then
'chkActive = vbChecked
'Else
'chkActive = vbUnchecked
'End If
'cmdedit.Enabled = True
'cmdSave.Enabled = True
End If
End Sub




















Private Sub txtVariance_Change()

End Sub
