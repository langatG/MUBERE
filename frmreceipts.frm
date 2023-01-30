VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmreceipts 
   Caption         =   "RECEIPTS ENTRY"
   ClientHeight    =   10185
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11310
   Icon            =   "frmreceipts.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10185
   ScaleWidth      =   11310
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.CommandButton cmdchangepro 
      Caption         =   "Change price or Quantity"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      TabIndex        =   99
      Top             =   4320
      Width           =   2535
   End
   Begin VB.CommandButton cmdrepo67 
      Caption         =   "Feeds Balance  Process"
      Height          =   495
      Left            =   7680
      TabIndex        =   98
      Top             =   9480
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CheckBox chkhalf 
      Caption         =   "Half Product?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   97
      Top             =   2040
      Width           =   1815
   End
   Begin VB.ComboBox Cmbstation 
      Height          =   315
      ItemData        =   "frmreceipts.frx":0442
      Left            =   9120
      List            =   "frmreceipts.frx":0444
      TabIndex        =   96
      Top             =   3480
      Width           =   2055
   End
   Begin VB.CommandButton cmdstorece 
      Caption         =   "Receive Stock"
      Height          =   495
      Left            =   6360
      TabIndex        =   95
      Top             =   9480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdproprocess 
      Caption         =   "Products Process"
      Height          =   495
      Left            =   3120
      TabIndex        =   93
      Top             =   9480
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Stock To Branch"
      Height          =   495
      Left            =   4920
      TabIndex        =   90
      Top             =   9480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Reports"
      Height          =   495
      Left            =   9720
      TabIndex        =   89
      Top             =   9480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Repackage"
      Height          =   495
      Left            =   3120
      TabIndex        =   88
      Top             =   9480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdenq 
      Caption         =   "Supplier Enqury"
      Height          =   495
      Left            =   240
      TabIndex        =   87
      Top             =   9480
      Width           =   1335
   End
   Begin VB.TextBox txtserai 
      Height          =   285
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   86
      Top             =   1560
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Expired products rpt"
      Height          =   555
      Left            =   10200
      TabIndex        =   85
      Top             =   9360
      Width           =   1815
   End
   Begin VB.OptionButton Optmpesa 
      Caption         =   "M-pesa"
      Height          =   195
      Left            =   840
      TabIndex        =   83
      Top             =   3360
      Width           =   855
   End
   Begin VB.TextBox txtname 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   11520
      TabIndex        =   81
      Top             =   4680
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.OptionButton Optothers 
      Caption         =   "Institutions"
      Height          =   375
      Left            =   11160
      TabIndex        =   80
      Top             =   3960
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txtsellingprice 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8160
      Locked          =   -1  'True
      TabIndex        =   78
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox txtbuyingprice 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   8160
      Locked          =   -1  'True
      TabIndex        =   76
      Top             =   120
      Width           =   1455
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
      Left            =   8400
      TabIndex        =   74
      Top             =   9000
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox txtmobile 
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   8400
      TabIndex        =   73
      Top             =   5400
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox txtidno 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6240
      TabIndex        =   71
      Top             =   5520
      Width           =   1935
   End
   Begin VB.TextBox txttransby 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3600
      TabIndex        =   68
      Top             =   5520
      Width           =   2055
   End
   Begin VB.Frame fra1 
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   6840
      TabIndex        =   59
      Top             =   1320
      Width           =   4455
      Begin VB.PictureBox Picture4 
         Height          =   255
         Left            =   1320
         Picture         =   "frmreceipts.frx":0446
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   63
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox Picture3 
         Height          =   255
         Left            =   1320
         Picture         =   "frmreceipts.frx":0D10
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   62
         Top             =   1440
         Width           =   255
      End
      Begin VB.TextBox txtdracc 
         Height          =   375
         Left            =   1680
         TabIndex        =   61
         Top             =   480
         Width           =   2655
      End
      Begin VB.TextBox txtcracc 
         Height          =   375
         Left            =   1680
         TabIndex        =   60
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Label lbldracc 
         BackColor       =   &H8000000E&
         Height          =   375
         Left            =   120
         TabIndex        =   67
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label lblcracc 
         BackColor       =   &H8000000E&
         Height          =   375
         Left            =   120
         TabIndex        =   66
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label20 
         Caption         =   "DrAccNo"
         Height          =   255
         Left            =   120
         TabIndex        =   65
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label19 
         Caption         =   "Craccno"
         Height          =   255
         Left            =   120
         TabIndex        =   64
         Top             =   960
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdproductaging 
      Caption         =   "Aging Products"
      Height          =   375
      Left            =   12120
      TabIndex        =   58
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton cmdsalesreturn 
      Caption         =   "Sales Return"
      Height          =   495
      Left            =   1680
      TabIndex        =   57
      Top             =   9480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox TXTTOTAL 
      Enabled         =   0   'False
      Height          =   495
      Left            =   8640
      TabIndex        =   56
      Text            =   "0"
      Top             =   6360
      Width           =   2535
   End
   Begin VB.TextBox TXTCHANGE 
      Height          =   495
      Left            =   8640
      TabIndex        =   54
      Text            =   "0"
      Top             =   8520
      Width           =   2535
   End
   Begin VB.TextBox txtamtreceived 
      Height          =   495
      Left            =   8640
      TabIndex        =   52
      Text            =   "0"
      Top             =   7320
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Remove"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   50
      Top             =   4920
      Width           =   2175
   End
   Begin VB.TextBox txtstaffno 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   12120
      TabIndex        =   48
      Top             =   4920
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton optstaff 
      Caption         =   "Staff"
      Height          =   255
      Left            =   10680
      TabIndex        =   46
      Top             =   4320
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.OptionButton Optbranch 
      Caption         =   "Station"
      Height          =   255
      Left            =   8400
      TabIndex        =   45
      Top             =   3120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ComboBox ports 
      Height          =   315
      ItemData        =   "frmreceipts.frx":15DA
      Left            =   960
      List            =   "frmreceipts.frx":15EA
      TabIndex        =   42
      Text            =   "\\127.0.0.1\E-PoS 80mm Thermal Printer1"
      Top             =   9000
      Width           =   4095
   End
   Begin MSComCtl2.DTPicker DTPto 
      Height          =   255
      Left            =   9360
      TabIndex        =   41
      Top             =   960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      Format          =   143327233
      CurrentDate     =   40588
   End
   Begin VB.TextBox txttranscode 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3960
      TabIndex        =   38
      Top             =   4800
      Width           =   1335
   End
   Begin VB.OptionButton opttransport 
      Caption         =   "Transporters"
      Height          =   255
      Left            =   9960
      TabIndex        =   35
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CheckBox chkPrint 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Print Receipt"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   8400
      TabIndex        =   34
      Top             =   9000
      Value           =   2  'Grayed
      Width           =   2175
   End
   Begin VB.ComboBox cboproductname 
      Height          =   315
      Left            =   1680
      TabIndex        =   33
      Top             =   1200
      Width           =   4215
   End
   Begin VB.TextBox txtSNo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3720
      TabIndex        =   23
      Top             =   3360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.OptionButton lblCheckOff 
      Caption         =   "Check Off"
      Height          =   255
      Left            =   1800
      TabIndex        =   21
      Top             =   3360
      Width           =   1215
   End
   Begin VB.OptionButton optCash 
      Caption         =   "Cash"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   3360
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      DrawStyle       =   2  'Dot
      DrawWidth       =   17015
      Height          =   360
      Left            =   4200
      Picture         =   "frmreceipts.frx":1606
      ScaleHeight     =   360
      ScaleWidth      =   240
      TabIndex        =   16
      Top             =   720
      Width           =   240
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      DrawStyle       =   2  'Dot
      DrawWidth       =   17015
      Height          =   360
      Left            =   4200
      Picture         =   "frmreceipts.frx":1788
      ScaleHeight     =   360
      ScaleWidth      =   240
      TabIndex        =   15
      Top             =   240
      Width           =   240
   End
   Begin VB.CommandButton cmdnextitem 
      Caption         =   "Next item"
      Default         =   -1  'True
      Height          =   495
      Left            =   2040
      TabIndex        =   14
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
      Height          =   495
      Left            =   4080
      TabIndex        =   4
      Top             =   2640
      Width           =   855
   End
   Begin VB.TextBox txtpcode 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   720
      Width           =   2535
   End
   Begin VB.TextBox txtquantity 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Text            =   "1"
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox txtamount 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   2040
      Width           =   1935
   End
   Begin MSComCtl2.DTPicker txtransdate 
      Height          =   375
      Left            =   4920
      TabIndex        =   8
      Top             =   600
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   143327233
      CurrentDate     =   40265
   End
   Begin VB.TextBox txtrno 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
   Begin MSComctlLib.ListView Lvwitems 
      Height          =   1455
      Left            =   240
      TabIndex        =   19
      Top             =   5880
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   2566
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      MousePointer    =   4
      MouseIcon       =   "frmreceipts.frx":190A
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ITEM"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "QNTY"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "PRICE"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "AMOUNT"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Cash"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTPfrom 
      Height          =   255
      Left            =   7440
      TabIndex        =   91
      Top             =   960
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      _Version        =   393216
      Format          =   143327233
      CurrentDate     =   40588
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1575
      Left            =   240
      TabIndex        =   92
      Top             =   7320
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   2778
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
      NumItems        =   0
   End
   Begin VB.CommandButton cmdsagroded 
      Caption         =   "Staff Agrovet Deductions"
      Height          =   375
      Left            =   120
      TabIndex        =   49
      Top             =   7320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdnew 
      Caption         =   "New"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   7080
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   6720
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSComctlLib.ProgressBar prgStatus 
      Height          =   255
      Left            =   5040
      TabIndex        =   100
      Top             =   9000
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   8160
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Label Label27 
      BackColor       =   &H0080FFFF&
      Height          =   255
      Left            =   5280
      TabIndex        =   94
      Top             =   9000
      Width           =   2775
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Left            =   9960
      TabIndex        =   84
      Top             =   360
      Width           =   855
   End
   Begin VB.Label lblinsname 
      AutoSize        =   -1  'True
      Caption         =   "Ins Name"
      Height          =   195
      Left            =   10680
      TabIndex        =   82
      Top             =   5040
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Label lblstnames 
      BackColor       =   &H00FF00FF&
      Height          =   255
      Left            =   10800
      TabIndex        =   79
      Top             =   4560
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label25 
      Caption         =   "Selling Price"
      Height          =   255
      Left            =   6960
      TabIndex        =   77
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label24 
      Caption         =   "Buying Price"
      Height          =   255
      Left            =   6960
      TabIndex        =   75
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label23 
      Caption         =   "Mobile no"
      Height          =   255
      Left            =   8520
      TabIndex        =   72
      Top             =   5040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label22 
      Caption         =   "Id no"
      Height          =   255
      Left            =   5760
      TabIndex        =   70
      Top             =   5520
      Width           =   495
   End
   Begin VB.Label Label21 
      Caption         =   "Transby"
      Height          =   255
      Left            =   2160
      TabIndex        =   69
      Top             =   5520
      Width           =   975
   End
   Begin VB.Label Label18 
      Caption         =   "TOTAL"
      Height          =   255
      Left            =   8520
      TabIndex        =   55
      Top             =   6000
      Width           =   1575
   End
   Begin VB.Label Label17 
      Caption         =   "CHANGE"
      Height          =   255
      Left            =   8520
      TabIndex        =   53
      Top             =   8160
      Width           =   975
   End
   Begin VB.Label Label15 
      Caption         =   "AMOUNT RECEIVED"
      Height          =   255
      Left            =   8520
      TabIndex        =   51
      Top             =   6960
      Width           =   1695
   End
   Begin VB.Label Label14 
      Caption         =   "Staff No"
      Height          =   255
      Left            =   10560
      TabIndex        =   47
      Top             =   4920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblstation 
      Caption         =   "Agrovet Branch"
      Height          =   255
      Left            =   7920
      TabIndex        =   44
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label12 
      Caption         =   "Printer Port"
      Height          =   375
      Left            =   120
      TabIndex        =   43
      Top             =   9000
      Width           =   855
   End
   Begin VB.Label lbltransnetpay 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3960
      TabIndex        =   40
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Label Label10 
      Caption         =   "Net Pay:"
      Height          =   255
      Left            =   2760
      TabIndex        =   39
      Top             =   5160
      Width           =   855
   End
   Begin VB.Label lbltransportername 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF00FF&
      Height          =   195
      Left            =   5520
      TabIndex        =   37
      Top             =   4800
      Width           =   165
   End
   Begin VB.Label Label5 
      Caption         =   "Transport Code"
      Height          =   255
      Left            =   2760
      TabIndex        =   36
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label lblSNames 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF00FF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   4800
      TabIndex        =   32
      Top             =   3360
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Label Label13 
      Caption         =   "Total Kgs :"
      Height          =   255
      Left            =   120
      TabIndex        =   31
      Top             =   3840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblTKgs 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1080
      TabIndex        =   30
      Top             =   3840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label11 
      Caption         =   "Gross Pay:"
      Height          =   255
      Left            =   2520
      TabIndex        =   29
      Top             =   3840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblGPay 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3480
      TabIndex        =   28
      Top             =   3840
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label8 
      Caption         =   "Deductions :"
      Height          =   255
      Left            =   5400
      TabIndex        =   27
      Top             =   3840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblDed 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   6600
      TabIndex        =   26
      Top             =   3840
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblNPay 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4200
      TabIndex        =   25
      Top             =   4320
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblNetPay 
      BackColor       =   &H0000FF00&
      Caption         =   "NetPay:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   24
      Top             =   4320
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblSNo 
      Caption         =   "SNo :"
      Height          =   255
      Left            =   3120
      TabIndex        =   22
      Top             =   3480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblbalance 
      BackColor       =   &H00FF00FF&
      Height          =   255
      Left            =   3840
      TabIndex        =   17
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "Product Code"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Quantity"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Amount"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Trans_Date"
      Height          =   255
      Left            =   4920
      TabIndex        =   7
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Product Name"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Receipt No."
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label9 
      Caption         =   "Balance"
      Height          =   255
      Left            =   3960
      TabIndex        =   18
      Top             =   1920
      Width           =   615
   End
End
Attribute VB_Name = "frmreceipts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Provider As String
Dim SelectedDsn As String
Dim DIA
Dim amount As Double

Private Sub cboproductname_Change()
If Cmbstation = "" Then
MsgBox "Please select the branch", vbInformation
Exit Sub
End If

Set rst = oSaccoMaster.GetRecordset("select p_code,pprice,sprice,AI from ag_products where p_name ='" & cboproductname & "' and branch='" & Cmbstation & "'")
If Not rst.EOF Then
txtpcode = rst.Fields("p_code")
txtbuyingprice = rst.Fields("pprice")
txtsellingprice = rst.Fields("sprice")
txtserai = rst.Fields("AI")
'txtsel
End If
'sql = "select P_CODE,P_NAME,S_NO,QOUT,supplierid,pprice,sprice from ag_products where p_code='" & txtpcode & "'"
'Set rs = New ADODB.Recordset
'rs.Open sql, cn
'If Not rs.EOF Then
'
'If Not IsNull(rs.Fields(1)) Then cboproductname = (rs.Fields(1))
'If Not IsNull(rs.Fields(5)) Then txtbuyingprice = (rs.Fields(5))
'If Not IsNull(rs.Fields(6)) Then txtsellingprice = (rs.Fields(6))


End Sub

Private Sub cboproductname_Click()
If Cmbstation = "" Then
MsgBox "Please select the branch", vbInformation
Exit Sub
End If
'sql = "SELECT     UserLoginIDs, UserGroup, Branch,SUPERUSER From UserAccounts where UserLoginIDs='" & User & "'"
'    Set rs = oSaccoMaster.GetRecordset(sql)
'    If Not rs.EOF Then
'    'Cmbstation = rs!Branch
'    If rs!Branch <> Cmbstation Then
'    MsgBox "You are not allowed to sell in this branch", vbInformation
'    Exit Sub
'    End If
'    End If
Set rst = oSaccoMaster.GetRecordset("select p_code from ag_products where p_name ='" & cboproductname & "' and branch='" & Cmbstation & "'")
If Not rst.EOF Then
txtpcode = rst.Fields("p_code")
End If
If Cmbstation = "" Then
MsgBox "Please select the branch", vbInformation
Exit Sub
End If

Set rst = oSaccoMaster.GetRecordset("select p_code,pprice,sprice,AI from ag_products where p_name ='" & cboproductname & "' and branch='" & Cmbstation & "'")
If Not rst.EOF Then
txtpcode = rst.Fields("p_code")
txtbuyingprice = rst.Fields("pprice")
txtsellingprice = rst.Fields("sprice")
txtserai = rst.Fields("AI")
'txtsel
End If
'sql = "select P_CODE,P_NAME,S_NO,QOUT,supplierid,pprice,sprice from ag_products where p_code='" & txtpcode & "'"
'Set rs = New ADODB.Recordset
'rs.Open sql, cn
'If Not rs.EOF Then
'
'If Not IsNull(rs.Fields(1)) Then cboproductname = (rs.Fields(1))
'If Not IsNull(rs.Fields(5)) Then txtbuyingprice = (rs.Fields(5))
'If Not IsNull(rs.Fields(6)) Then txtsellingprice = (rs.Fields(6))
End Sub

Private Sub cboproductname_KeyPress(KeyAscii As Integer)
If Cmbstation = "" Then
MsgBox "Please select the branch", vbInformation
Exit Sub
End If
'KeyAscii = 0
'cboproductname_Validate (True)
Set rst = oSaccoMaster.GetRecordset("select p_code from ag_products where p_name ='" & cboproductname & "' and branch='" & Cmbstation & "'")
If Not rst.EOF Then
txtpcode = rst.Fields("p_code")
End If

End Sub

Private Sub cboproductname_Validate(Cancel As Boolean)
cmdNew_Click

Provider = cn
Set cn = New ADODB.Connection
Dim p As Integer
'cn.Open Provider, "atm"
'If reportpath = "" Then reportpath = GetSetting("payroll", "AppName", "rptPath", rptPath)
'Dim rst As New ADODB.Recordset
sql = ""
'SELECT p_code, p_name, S_No, Qout, sprice FROM   ag_Products
sql = "select p_code, S_No,Qin ,Qout, sprice,AI from ag_products where p_name='" & cboproductname & "' and p_code='" & txtpcode & "' and branch='" & Cmbstation & "'"
Set rs = New ADODB.Recordset
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
txtpcode = rs.Fields(0)
lblbalance = rs.Fields(3)
'txtserialno = rs.Fields(1)
txtamount = rs.Fields(4)
txtserai = rs.Fields(5)

End If

End Sub

Private Sub chkhalf_Click()
If chkhalf.value = vbChecked Then
chkhalf = 1
Else
chkhalf = 0
End If

End Sub

Private Sub chprint_Click()
ports.Clear
ports = ""
'//If the drivers are installed it won't matter whether the Port is indicated
' or not it will just work.

If chprint.value = vbChecked Then
ports.AddItem "LPT1"
ports = "LPT1"
ports.AddItem "LPT2"
ports.AddItem "LPT3"
ports.AddItem "LPT4"
ports.AddItem "LPT5"
Else
'Share the printer first the use of 127.0.0.1 which is
'standard IP address for a loopback network connection
'instead of getting the computer name or IP Address
'
Dim prnPrinter As Printer
Dim pr As String
ports.Clear

For Each prnPrinter In Printers
   If InStr(prnPrinter.DeviceName, "\\") Then
    ports.AddItem prnPrinter.DeviceName
    If InStr(prnPrinter.DeviceName, "G") Then
    ports.Text = prnPrinter.DeviceName
    End If
    Else
    ports.AddItem "\\127.0.0.1\" & prnPrinter.DeviceName
    If InStr(prnPrinter.DeviceName, "G") Then
    ports.Text = "\\127.0.0.1\" & prnPrinter.DeviceName
    End If
    End If
   
   
Next
End If
'This code will work only if there is a connection e.g LAN or modem.
'It is not a must that it is an internet connection because
'computer's network interface card has to be functional
End Sub
Private Sub Cmbstation_Click()
lblCheckOff.Visible = True
lblSNo.Visible = False
txtSNo.Visible = False
Label13.Visible = False
lblTKgs.Visible = False
Label11.Visible = False
lblGPay.Visible = False
Label8.Visible = False
lblDed.Visible = False
lblNetPay.Visible = False
lblNPay.Visible = False
Label5.Visible = False
txttranscode.Visible = False
Label10.Visible = False
lbltransnetpay.Visible = False
    cboproductname.Clear
    sql = ""
    sql = "select DISTINCT(P_NAME) from ag_products where Branch='" & Cmbstation & "'"
    Set rst = oSaccoMaster.GetRecordset(sql)
    While Not rst.EOF
    cboproductname.AddItem rst.Fields(0)
    rst.MoveNext
    Wend
End Sub
Private Sub Cmbstation_Change()
lblCheckOff.Visible = True
lblSNo.Visible = False
txtSNo.Visible = False
Label13.Visible = False
lblTKgs.Visible = False
Label11.Visible = False
lblGPay.Visible = False
Label8.Visible = False
lblDed.Visible = False
lblNetPay.Visible = False
lblNPay.Visible = False
Label5.Visible = False
txttranscode.Visible = False
Label10.Visible = False
lbltransnetpay.Visible = False
lblSNames.Visible = False
    cboproductname.Clear
    sql = ""
    sql = "select DISTINCT(P_NAME) from ag_products where Branch='" & Cmbstation & "'"
    Set rst = oSaccoMaster.GetRecordset(sql)
    While Not rst.EOF
    cboproductname.AddItem rst.Fields(0)
    rst.MoveNext
    Wend
End Sub

Private Sub cmdchangepro_Click()
FRMCHANGE.Show vbModal
End Sub

Private Sub cmdclose_Click()
Unload Me
End Sub
Private Sub cmdenq_Click()
frmEnquery.Show vbModal
End Sub
Private Sub cmdNew_Click()
'Set rs = oSaccoMaster.GetRecordset("d_sp_NextReceipt")
Set rs = oSaccoMaster.GetRecordset("select rcpno from rcpno")
If Not (rs.EOF) Then
txtrno = rs.Fields(0) + 1
Else
txtrno = 1
End If

' txtpcode = ""
 'txtserialno = ""
 txtquantity = 1
 txtamount = 0
 txtamtreceived = 0
 TXTCHANGE = 0
 TXTTOTAL = 0
End Sub

Private Sub cmdnextitem_Click()
Dim cash As Integer
Dim total As Double
On Error GoTo ErrorHandler
'check the user
If Cmbstation = "" Then
MsgBox "Please select your branch", vbInformation
Exit Sub
End If

If Trim(txtquantity) = "" Then
MsgBox "Quantity cannot be Zero", vbInformation
Exit Sub
End If

'''*****************************check if supliers branch is correct ****
'If lblCheckOff.value = 1 Then
    sql = ""
    sql = "select SNo,Location from d_Suppliers where SNo='" & txtSNo & "'"
    Set rss = oSaccoMaster.GetRecordset(sql)
    If Not rss.EOF Then
    If rss!Location <> Cmbstation Then
     MsgBox "This supplier is not from this branch: " & Cmbstation & "", vbInformation
     Exit Sub
    End If
    End If
'End If
'''*****************************end **************

    If Optbranch = True Then
    If Trim(Cmbstation.Text) = "" Then
        MsgBox "Please enter the Agrovet Station."
            Cmbstation.SetFocus
    Exit Sub
    End If
    
    Set rst = oSaccoMaster.GetRecordset("select pprice from ag_products where p_code='" & txtpcode & "'")
    If Not rst.EOF Then
    'txtAmount = Rst.Fields("pprice")
    End If
    End If
    
    If opttransport = True Then
    If Trim(txttranscode) = "" Then
        MsgBox "Please enter the Transporter"
    
    Exit Sub
    End If
    End If
    
    If txtpcode = "" Then
        MsgBox "Please Enter the Product CODE before You Proceed!", vbCritical
        Exit Sub
    End If
    If txtrno = "" Then
        MsgBox "Please Enter Receipt Number before you Proceed!", vbCritical
        Exit Sub
    End If
    
 
    
If txtamount = "" Then
txtamount = 0
End If
Provider = "maziwa"
Set cn = New ADODB.Connection
cn.Open Provider, "atm", "atm"
'// check if they are in stock.
Dim rsinstock As Recordset
sql = ""
sql = "select P_CODE,qin,qout,AI from ag_products where p_code='" & txtpcode & "' and branch='" & Cmbstation & "'"
Set rsinstock = New ADODB.Recordset
rsinstock.Open sql, cn
'// check the stock if it is less than zero
If rsinstock.Fields(2) <= 0 Then
MsgBox "Sorry Stock is Zero please re-stock before your proceed", vbInformation
Exit Sub
End If
'// check the quanttity being sold versus the balance
Dim piu As Double

If chkhalf = 0 Then
piu = rsinstock.Fields(2) - CInt(txtquantity)
Else
piu = rsinstock.Fields(2) - CInt(txtquantity * 0.5)
End If

If piu < 0 Then
MsgBox "Stock will be negative please re-stock before you proceed", vbInformation
Exit Sub
End If

If optCash.value = True Then
cash = 1
Else
cash = 0
End If

Dim j, Coun As Integer
j = 1

If chkhalf = 1 Then
txtquantity = txtquantity * 0.5
End If

'Check if same item is in the list
   Do While Not j > (Coun)
         Lvwitems.ListItems.Item(j).selected = True
            
    If Lvwitems.SelectedItem = txtpcode Then
        txtquantity = (CCur(txtquantity) + CCur(Lvwitems.SelectedItem.ListSubItems(2)))
        Lvwitems.ListItems.Remove (Lvwitems.SelectedItem.Index)
                        
        Set li = Lvwitems.ListItems.Add(, , txtpcode)
                        li.SubItems(1) = cboproductname & ""
                        li.SubItems(2) = txtquantity & ""
                        li.SubItems(3) = txtamount & ""
                        li.SubItems(4) = CCur(txtamount) * CCur(txtquantity) & ""
                        li.SubItems(5) = cash
                        'Total = CCur(Total + li.SubItems(4))
                        TXTTOTAL = total
                                                
        j = Coun + 1
        
        lblbalance = CCur(lblbalance) - CCur(txtquantity)

        txtpcode = ""
        txtquantity = ""
       ' txtserialno = ""
        txtpcode.SetFocus
        Exit Sub
         
    
   
'   lvwItems.ListItems.Item(J).selected = True
   End If
   j = j + 1
    Loop
    
     If j > 1 Then
   
    Set li = Lvwitems.ListItems.Add(, , txtpcode)
                        li.SubItems(1) = cboproductname & ""
                        li.SubItems(2) = txtquantity & ""
                        li.SubItems(3) = txtamount & ""
                        li.SubItems(4) = CCur(txtamount) * (CCur(txtquantity)) & ""
                        li.SubItems(5) = cash
                        'Total = CCur(Total + li.SubItems(4))
                        TXTTOTAL = total
                        
        lblbalance = CCur(lblbalance) - CCur(txtquantity)
        txtpcode = ""
        txtquantity = ""
        'txtserialno = ""
        txtpcode.SetFocus
        Exit Sub
    End If
     If Coun = 0 Then
     Set li = Lvwitems.ListItems.Add(, , txtpcode)
                        li.SubItems(1) = cboproductname & ""
                        li.SubItems(2) = txtquantity & ""
                        li.SubItems(3) = txtamount & ""
                        li.SubItems(4) = CCur(txtamount) * (CCur(txtquantity)) & ""
                        li.SubItems(5) = cash
                        'Total = CCur(Total + li.SubItems(4))
                        TXTTOTAL = total
    End If

lblbalance = CCur(lblbalance) - CCur(txtquantity)
TXTTOTAL = 0
Do While Not j > (Lvwitems.ListItems.Count)
'For j = 1 To Lvwitems.ListItems.Count
 Lvwitems.ListItems.Item(j).selected = True
 total = total + CCur(Lvwitems.SelectedItem.SubItems(4))
 TXTTOTAL = total
j = j + 1
Loop
chkhalf.value = vbUnchecked
txtpcode = ""
txtquantity = ""
'txtserialno = ""
txtpcode.SetFocus
Exit Sub
ErrorHandler:
MsgBox err.description
End Sub
Private Sub cmdproprocess_Click()
On Error GoTo ErrorHandler
Startdate = DateSerial(Year(txtransdate), Month(txtransdate), 1)
Enddate = DateSerial(Year(txtransdate), Month(txtransdate) + 1, 1 - 1)
sql = ""
sql = "set dateformat dmy delete from ag_ReceiptsProcess where Date>= '" & Startdate & "' And Date<='" & Enddate & "' and Name<>'FEEDS CARRY FORWARD' and Name<>'FEEDS RECEIVE'"
cn.Execute sql
Label27.Visible = True
Label27.Enabled = True
Dim U, C As Integer
sql = ""
sql = "set dateformat dmy Select count(distinct(SNO)) as u from ag_Receipts where SNO<>''and T_Date>= '" & Startdate & "' And T_Date<='" & Enddate & "'  "
Set rs = cn.Execute(sql)
Dim a As Double
a = rs.Fields(0)

  U = rs.Fields(0)
 prgStatus.Visible = True
'prgStatus.max = rs.Fields(0)
'I = prgStatus
'MsgBox "Please wait " & U & ""
prgStatus.max = 100
prgStatus.Min = 0
I = 0
  sql = ""
  sql = "set dateformat dmy Select distinct(SNO) as y   from   ag_Receipts where SNO<>''and T_Date>= '" & Startdate & "' And T_Date<='" & Enddate & "' order by SNO asc  "
  Set rsg = cn.Execute(sql)
  While Not rs.EOF
  Do While Not U <= 0
  I = I + 1
prgStatus = Round((I / a) * 100, 0)
      'I = I - 1
  'prgStatus.value = (rs.AbsolutePosition / I)
  
  If Not rsg.EOF Then
    C = rsg.Fields(0)
    Label27.FontBold = True
    Label27 = "Please wait as we process"
       sql = ""
       sql = "set dateformat dmy Select distinct(Remarks)from ag_Receipts where S_No='" & C & "'and T_Date >= '" & Startdate & "' And T_Date<='" & Enddate & "' "
       Set rsb = cn.Execute(sql)
       sql = ""
       sql = "set dateformat dmy Select  count(distinct(Remarks))from ag_Receipts where S_No='" & C & "' and T_Date >= '" & Startdate & "' And T_Date<='" & Enddate & "' "
       Set rsh = cn.Execute(sql)
       Do While Not rsh.EOF
      If Not rsb.EOF Then
       If Not rsh.EOF Then
           sql = ""
           sql = "set dateformat dmy Select  P_code, T_Date, Amount, S_No,Remarks, Branch,Qua from ag_Receipts where S_No='" & C & "'and Remarks='" & rsb.Fields(0) & "' and T_Date >= '" & Startdate & "' And T_Date<='" & Enddate & "'order by S_No"
           Set rst = cn.Execute(sql)
          If Not rst.EOF Then
            sql = ""
            sql = "set dateformat dmy Select sum(Qua) from ag_Receipts where S_No='" & C & "' and Remarks like'%" & rst.Fields(4) & "%'and T_Date >= '" & Startdate & "' And T_Date<='" & Enddate & "' "
            Set rsz = cn.Execute(sql)
                      
           sql = ""
           sql = "set dateformat dmy select * from ag_ReceiptsProcess where SNo='" & rsg.Fields(0) & "' and Date >= '" & Startdate & "' And Date<='" & Enddate & "'"
           Set rss = oSaccoMaster.GetRecordset(sql)
           If rss.EOF Then
             sql = ""
             sql = "set dateformat dmy insert into  ag_ReceiptsProcess(SNo, Name, Date, Branch)"
             sql = sql & "  values('" & rst.Fields(3) & "','','" & txtransdate & "','" & rst.Fields(5) & "')"
             cn.Execute sql
            Else
            End If
           sql = ""
           sql = "set dateformat dmy select Pro1, Pro2, Pro3, Pro4, Pro5, Pro6, Pro7,Pro8,Pro9 from ag_ReceiptsProcess where SNo='" & rsg.Fields(0) & "' and Date >= '" & Startdate & "' And Date<='" & Enddate & "'"
           Set rsl = oSaccoMaster.GetRecordset(sql)
            
           sql = ""
           sql = "select p_code,p_name from ag_Products where p_name='" & rst.Fields(4) & "' ORDER BY p_code asc"
           Set rsm = oSaccoMaster.GetRecordset(sql)
           If Not rsm.EOF Then
           Dim strong As Integer
           strong = rsm.Fields(0)
                    'rsm.Fields(0) = " & rst.Fields(4) & "
            Select Case strong
             Case "1"
              sql = ""
              sql = "set dateformat dmy Update ag_ReceiptsProcess SET Pro1 ='" & rsl.Fields(0) + rsz.Fields(0) & "' WHERE SNo='" & rst.Fields(3) & "' and Date >= '" & Startdate & "'And Date<='" & Enddate & "'"
              cn.Execute sql
             Case "2"
              sql = ""
              sql = "set dateformat dmy Update ag_ReceiptsProcess SET Pro2 ='" & rsl.Fields(1) + rsz.Fields(0) & "' WHERE SNo='" & rst.Fields(3) & "'and Date >= '" & Startdate & "'And Date<='" & Enddate & "'"
              cn.Execute sql
             Case "3"
              sql = ""
              sql = "set dateformat dmy Update ag_ReceiptsProcess SET Pro3 ='" & rsl.Fields(2) + rsz.Fields(0) & "' WHERE SNo='" & rst.Fields(3) & "'and Date >= '" & Startdate & "'And Date<='" & Enddate & "'"
              cn.Execute sql
             Case "4"
              sql = ""
              sql = "set dateformat dmy Update ag_ReceiptsProcess SET Pro4 ='" & rsl.Fields(3) + rsz.Fields(0) & "' WHERE SNo='" & rst.Fields(3) & "'and Date >= '" & Startdate & "'And Date<='" & Enddate & "'"
              cn.Execute sql
             Case "5"
              sql = ""
              sql = "set dateformat dmy Update ag_ReceiptsProcess SET Pro5 ='" & rsl.Fields(4) + rsz.Fields(0) & "' WHERE SNo='" & rst.Fields(3) & "'and Date >= '" & Startdate & "'And Date<='" & Enddate & "'"
              cn.Execute sql
             Case "6"
              sql = ""
              sql = "set dateformat dmy Update ag_ReceiptsProcess SET Pro6 ='" & rsl.Fields(5) + rsz.Fields(0) & "' WHERE SNo='" & rst.Fields(3) & "'and Date >= '" & Startdate & "'And Date<='" & Enddate & "'"
              cn.Execute sql
             Case "7"
              sql = ""
              sql = "set dateformat dmy Update ag_ReceiptsProcess SET Pro7 ='" & rsl.Fields(6) + rsz.Fields(0) & "' WHERE SNo='" & rst.Fields(3) & "'and Date >= '" & Startdate & "'And Date<='" & Enddate & "'"
              cn.Execute sql
             Case "8"
              sql = ""
              sql = "set dateformat dmy Update ag_ReceiptsProcess SET Pro8 ='" & rsl.Fields(7) + rsz.Fields(0) & "' WHERE SNo='" & rst.Fields(3) & "'and Date >= '" & Startdate & "'And Date<='" & Enddate & "'"
              cn.Execute sql
             Case "9"
              sql = ""
              sql = "set dateformat dmy Update ag_ReceiptsProcess SET Pro9 ='" & rsl.Fields(8) + rsz.Fields(0) & "' WHERE SNo='" & rst.Fields(3) & "'and Date >= '" & Startdate & "'And Date<='" & Enddate & "'"
              cn.Execute sql
'             Case "10"
'              sql = ""
'              sql = "set dateformat dmy Update ag_ReceiptsProcess SET Pro10 ='" & rsl.Fields(9) + rsz.Fields(0) & "' WHERE SNo='" & rst.Fields(3) & "'and Date >= '" & Startdate & "'And Date<='" & Enddate & "'"
'              cn.Execute sql
             Case Else
            End Select
          
           End If
           End If
           rsb.MoveNext
           End If
         Else
           rsh.MoveNext
          End If
         Loop
        
        sql = ""
        sql = "set dateformat dmy Select SNo, Names from d_Suppliers where SNo='" & rsg.Fields(0) & "' "
        Set rst = cn.Execute(sql)
        sql = ""
        sql = "set dateformat dmy Update ag_ReceiptsProcess SET Name ='" & rst.Fields(1) & "' WHERE SNo='" & rsg.Fields(0) & "'and Date >= '" & Startdate & "'And Date<='" & Enddate & "'"
        cn.Execute sql
        
      Else
      Label27.Visible = False
     MsgBox "Completed succesfully ", vbInformation
    End If
   U = U - 1
 rsg.MoveNext
Loop
siche
Label27.Visible = False
MsgBox "Completed succesfully ", vbInformation
Exit Sub
Wend

Exit Sub
ErrorHandler:
MsgBox err.description
End Sub
Private Sub siche()
Startdate = DateSerial(Year(txtransdate), Month(txtransdate), 1)
Enddate = DateSerial(Year(txtransdate), Month(txtransdate) + 1, 1 - 1)

sql = ""
sql = "set dateformat dmy delete from ag_ReceiptsProcess where Date>= '" & Startdate & "' And Date<='" & Enddate & "' and Name='FEEDS RECEIVE'"
cn.Execute sql

Dim U As Integer
Dim C As String
sql = ""
sql = "set dateformat dmy Select count(distinct(Branch)) as u from DRAWNSTOCK WHERE DATE>='" & Startdate & "' AND DATE<= '" & Enddate & "'"
Set rs = cn.Execute(sql)
  U = rs.Fields(0)
  sql = ""
  sql = "set dateformat dmy Select distinct(Branch) as y from DRAWNSTOCK WHERE DATE>='" & Startdate & "' AND DATE<= '" & Enddate & "' order by Branch asc  "
  Set rsg = cn.Execute(sql)
  While Not rs.EOF
  Do While Not U <= 0
  If Not rsg.EOF Then
    C = rsg.Fields(0)
       sql = ""
       sql = "set dateformat dmy Select distinct(DESCRIPTION)from DRAWNSTOCK where Branch='" & C & "' AND DATE>='" & Startdate & "' AND DATE<= '" & Enddate & "' "
       Set rsb = cn.Execute(sql)
       sql = ""
       sql = "set dateformat dmy Select  count(distinct(DESCRIPTION))from DRAWNSTOCK where Branch='" & C & "' AND DATE>='" & Startdate & "' AND DATE<= '" & Enddate & "'"
       Set rsh = cn.Execute(sql)
       Do While Not rsh.EOF
      If Not rsb.EOF Then
       If Not rsh.EOF Then
           sql = ""
           sql = "set dateformat dmy Select PRODUCTID,DATE, QUANTITY,DESCRIPTION,Branch from DRAWNSTOCK where Branch='" & C & "' and DESCRIPTION='" & rsb.Fields(0) & "' and DATE >= '" & Startdate & "' And DATE<='" & Enddate & "' order by PRODUCTID"
           Set rst = cn.Execute(sql)
          If Not rst.EOF Then
           sql = ""
           sql = "set dateformat dmy select * from ag_ReceiptsProcess where Name='FEEDS RECEIVE' and Branch='" & C & "' and Date >= '" & Startdate & "' And Date<='" & Enddate & "'"
           Set rss = oSaccoMaster.GetRecordset(sql)
           If rss.EOF Then
             sql = ""
             sql = "set dateformat dmy insert into  ag_ReceiptsProcess(SNo, Name, Date, Branch)"
             sql = sql & "  values('','FEEDS RECEIVE','" & rst.Fields(1) & "','" & rst.Fields(4) & "')"
             cn.Execute sql
            Else
            End If
           sql = ""
           sql = "set dateformat dmy select sPro1, sPro2, sPro3, sPro4, spro5, sPro6, sPro7,sPro8,sPro9,Name from ag_ReceiptsProcess where Name='FEEDS RECEIVE' and Branch='" & C & "' and Date >= '" & Startdate & "' And Date<='" & Enddate & "'"
           Set rsl = oSaccoMaster.GetRecordset(sql)
            
           sql = ""
           sql = "select PRODUCTID, PRODUCTNAME from DRAWNSTOCK where Branch='" & C & "' and PRODUCTNAME='" & rst.Fields(3) & "' and DATE >= '" & Startdate & "' And DATE<='" & Enddate & "' ORDER BY PRODUCTID asc"
           Set rsm = oSaccoMaster.GetRecordset(sql)
           If Not rsm.EOF Then
           Dim strong As Integer
           strong = rsm.Fields(0)
           If strong = "8" Then
           MsgBox "hi"
           Else
           End If
            Select Case strong
             Case "1"
              sql = ""
              sql = "set dateformat dmy Update ag_ReceiptsProcess SET sPro1 ='" & rsl.Fields(0) + rst.Fields(2) & "' WHERE Branch='" & rst.Fields(4) & "' and Name='" & rsl.Fields(9) & "' and Date >= '" & Startdate & "'And Date<='" & Enddate & "'"
              cn.Execute sql
             Case "2"
              sql = ""
              sql = "set dateformat dmy Update ag_ReceiptsProcess SET sPro2 ='" & rsl.Fields(1) + rst.Fields(2) & "' WHERE Branch='" & rst.Fields(4) & "' and Name='" & rsl.Fields(9) & "' and Date >= '" & Startdate & "'And Date<='" & Enddate & "'"
              cn.Execute sql
             Case "3"
              sql = ""
              sql = "set dateformat dmy Update ag_ReceiptsProcess SET sPro3 ='" & rsl.Fields(2) + rst.Fields(2) & "' WHERE Branch='" & rst.Fields(4) & "'and Name='" & rsl.Fields(9) & "' and Date >= '" & Startdate & "'And Date<='" & Enddate & "'"
              cn.Execute sql
             Case "4"
              sql = ""
              sql = "set dateformat dmy Update ag_ReceiptsProcess SET sPro4 ='" & rsl.Fields(3) + rst.Fields(2) & "' WHERE Branch='" & rst.Fields(4) & "' and Name='" & rsl.Fields(9) & "' and Date >= '" & Startdate & "'And Date<='" & Enddate & "'"
              cn.Execute sql
             Case "5"
              sql = ""
              sql = "set dateformat dmy Update ag_ReceiptsProcess SET sPro5 ='" & rsl.Fields(4) + rst.Fields(2) & "' WHERE Branch='" & rst.Fields(4) & "' and Name='" & rsl.Fields(9) & "' and Date >= '" & Startdate & "'And Date<='" & Enddate & "'"
              cn.Execute sql
             Case "6"
              sql = ""
              sql = "set dateformat dmy Update ag_ReceiptsProcess SET sPro6 ='" & rsl.Fields(5) + rst.Fields(2) & "' WHERE Branch='" & rst.Fields(4) & "' and Name='" & rsl.Fields(9) & "' and Date >= '" & Startdate & "'And Date<='" & Enddate & "'"
              cn.Execute sql
             Case "7"
              sql = ""
              sql = "set dateformat dmy Update ag_ReceiptsProcess SET sPro7 ='" & rsl.Fields(6) + rst.Fields(2) & "' WHERE Branch='" & rst.Fields(4) & "' and Name='" & rsl.Fields(9) & "' and Date >= '" & Startdate & "'And Date<='" & Enddate & "'"
              cn.Execute sql
              Case "8"
              sql = ""
              sql = "set dateformat dmy Update ag_ReceiptsProcess SET sPro8 ='" & rsl.Fields(7) + rst.Fields(2) & "' WHERE Branch='" & rst.Fields(4) & "' and Name='" & rsl.Fields(9) & "' and Date >= '" & Startdate & "'And Date<='" & Enddate & "'"
              cn.Execute sql
              Case "9"
              sql = ""
              sql = "set dateformat dmy Update ag_ReceiptsProcess SET sPro9 ='" & rsl.Fields(8) + rst.Fields(2) & "' WHERE Branch='" & rst.Fields(4) & "' and Name='" & rsl.Fields(9) & "' and Date >= '" & Startdate & "'And Date<='" & Enddate & "'"
              cn.Execute sql
             Case Else
            End Select
          
           End If
           End If
           rsb.MoveNext
           End If
         Else
           rsh.MoveNext
          End If
         Loop
      Else
        'MsgBox "Completed succesfully ", vbInformation
    End If
   U = U - 1
 rsg.MoveNext
Loop
 'MsgBox "Completed succesfully ", vbInformation
Exit Sub
Wend
Exit Sub
ErrorHandler:
MsgBox err.description

End Sub
Private Sub cmdrepo67_Click()
Startdate = DateSerial(Year(txtransdate), Month(txtransdate), 1)
Enddate = DateSerial(Year(txtransdate), Month(txtransdate) + 1, 1 - 1)

sql = ""
sql = "set dateformat dmy delete from ag_ReceiptsProcess where Date>= '" & Startdate & "' And Date<='" & Enddate & "' and Name='FEEDS CARRY FORWARD'"
cn.Execute sql

Dim U As Integer
Dim C As String
 prgStatus.Visible = True
sql = ""
sql = "set dateformat dmy Select count(distinct(Branch)) as u from ag_Products "
Set rs = cn.Execute(sql)
Dim a As Double
a = rs.Fields(0)
U = rs.Fields(0)
'prgStatus = rs.Fields(0)
'I = prgStatus + 3
prgStatus.max = 100
prgStatus.Min = 0
I = 0
'MsgBox "Please wait " & U & ""
  sql = ""
  sql = "set dateformat dmy Select distinct(Branch) as y   from   ag_Products order by Branch asc  "
  Set rsg = cn.Execute(sql)
  While Not rs.EOF
'        I = I - 1
'  prgStatus.value = (rs.AbsolutePosition + I)
  Do While Not U <= 0
  I = I + 1
prgStatus = Round((I / a) * 100, 0)
  If Not rsg.EOF Then
    C = rsg.Fields(0)

       sql = ""
       sql = "set dateformat dmy Select distinct(p_name)from ag_Products where Branch='" & C & "' "
       Set rsb = cn.Execute(sql)
       sql = ""
       sql = "set dateformat dmy Select  count(distinct(p_name))from ag_Products where Branch='" & C & "'"
       Set rsh = cn.Execute(sql)
       Do While Not rsh.EOF
      If Not rsb.EOF Then
       If Not rsh.EOF Then
           sql = ""
           sql = "set dateformat dmy Select  p_code,p_name,Qout, Date_Entered,Branch from ag_Products where Branch='" & C & "' and p_name='" & rsb.Fields(0) & "' order by p_code"
           Set rst = cn.Execute(sql)
          If Not rst.EOF Then
           sql = ""
           sql = "set dateformat dmy select * from ag_ReceiptsProcess where Name='FEEDS CARRY FORWARD' and Branch='" & C & "' and Date >= '" & Startdate & "' And Date<='" & Enddate & "'"
           Set rss = oSaccoMaster.GetRecordset(sql)
           If rss.EOF Then
             sql = ""
             sql = "set dateformat dmy insert into  ag_ReceiptsProcess(SNo, Name, Date, Branch)"
             sql = sql & "  values('','FEEDS CARRY FORWARD','" & txtransdate & "','" & rst.Fields(4) & "')"
             cn.Execute sql
            Else
            End If
           sql = ""
           sql = "set dateformat dmy select bPro1, bPro2, bPro3, bPro4, bPro5, bPro6, bPro7,bPro8,bPro9,Name from ag_ReceiptsProcess where Branch='" & C & "' and Name='FEEDS CARRY FORWARD'and Date >= '" & Startdate & "' And Date<='" & Enddate & "'"
           Set rsl = oSaccoMaster.GetRecordset(sql)
            
           sql = ""
           sql = "select p_code,p_name from ag_Products where Branch='" & C & "' and p_name='" & rst.Fields(1) & "' ORDER BY p_code asc"
           Set rsm = oSaccoMaster.GetRecordset(sql)
           If Not rsm.EOF Then
           Dim strong As Integer
           strong = rsm.Fields(0)
            Select Case strong
             Case "1"
              sql = ""
              sql = "set dateformat dmy Update ag_ReceiptsProcess SET bPro1 ='" & rsl.Fields(0) + rst.Fields(2) & "' WHERE Branch='" & rst.Fields(4) & "' and Name ='" & rsl.Fields(9) & "' and Date >= '" & Startdate & "'And Date<='" & Enddate & "'"
              cn.Execute sql
             Case "2"
              sql = ""
              sql = "set dateformat dmy Update ag_ReceiptsProcess SET bPro2 ='" & rsl.Fields(1) + rst.Fields(2) & "' WHERE Branch='" & rst.Fields(4) & "' and Name ='" & rsl.Fields(9) & "' and Date >= '" & Startdate & "'And Date<='" & Enddate & "'"
              cn.Execute sql
             Case "3"
              sql = ""
              sql = "set dateformat dmy Update ag_ReceiptsProcess SET bPro3 ='" & rsl.Fields(2) + rst.Fields(2) & "' WHERE Branch='" & rst.Fields(4) & "'and Name ='" & rsl.Fields(9) & "' and Date >= '" & Startdate & "'And Date<='" & Enddate & "'"
              cn.Execute sql
             Case "4"
              sql = ""
              sql = "set dateformat dmy Update ag_ReceiptsProcess SET bPro4 ='" & rsl.Fields(3) + rst.Fields(2) & "' WHERE Branch='" & rst.Fields(4) & "' and Name ='" & rsl.Fields(9) & "' and Date >= '" & Startdate & "'And Date<='" & Enddate & "'"
              cn.Execute sql
             Case "5"
              sql = ""
              sql = "set dateformat dmy Update ag_ReceiptsProcess SET bPro5 ='" & rsl.Fields(4) + rst.Fields(2) & "' WHERE Branch='" & rst.Fields(4) & "' and Name ='" & rsl.Fields(9) & "' and Date >= '" & Startdate & "'And Date<='" & Enddate & "'"
              cn.Execute sql
             Case "6"
              sql = ""
              sql = "set dateformat dmy Update ag_ReceiptsProcess SET bPro6 ='" & rsl.Fields(5) + rst.Fields(2) & "' WHERE Branch='" & rst.Fields(4) & "' and Name ='" & rsl.Fields(9) & "' and Date >= '" & Startdate & "'And Date<='" & Enddate & "'"
              cn.Execute sql
             Case "7"
              sql = ""
              sql = "set dateformat dmy Update ag_ReceiptsProcess SET bPro7 ='" & rsl.Fields(6) + rst.Fields(2) & "' WHERE Branch='" & rst.Fields(4) & "' and Name ='" & rsl.Fields(9) & "' and Date >= '" & Startdate & "'And Date<='" & Enddate & "'"
              cn.Execute sql
             Case "8"
              sql = ""
              sql = "set dateformat dmy Update ag_ReceiptsProcess SET bPro8 ='" & rsl.Fields(7) + rst.Fields(2) & "' WHERE Branch='" & rst.Fields(4) & "' and Name ='" & rsl.Fields(9) & "' and Date >= '" & Startdate & "'And Date<='" & Enddate & "'"
              cn.Execute sql
             Case "9"
              sql = ""
              sql = "set dateformat dmy Update ag_ReceiptsProcess SET bPro9 ='" & rsl.Fields(8) + rst.Fields(2) & "' WHERE Branch='" & rst.Fields(4) & "' and Name ='" & rsl.Fields(9) & "' and Date >= '" & Startdate & "'And Date<='" & Enddate & "'"
              cn.Execute sql
             Case Else
            End Select
          
           End If
           End If
           rsb.MoveNext
           End If
         Else
           rsh.MoveNext
          End If
         Loop
      Else
        MsgBox "Completed succesfully ", vbInformation
    End If
   U = U - 1
 rsg.MoveNext
Loop
 MsgBox "Completed succesfully ", vbInformation
Exit Sub
Wend
Exit Sub
ErrorHandler:
MsgBox err.description
End Sub

Private Sub cmdsagroded_Click()
'//staffagrovetdeductions
'//d_payroll\
'//call the companyname

 reportname = "staffagrovetdeductions.rpt"

 
 Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub cmdsalesreturn_Click()
'Unload Me
'check the user
sql = "SELECT     UserLoginIDs,levels, UserGroup, SUPERUSER From UserAccounts where UserLoginIDs='" & User & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
If rs!SuperUser <> "1" Then
MsgBox "You are not allowed to sell", vbInformation
Exit Sub

End If
End If
frmsalesreturn.Show vbModal

End Sub

Private Sub cmdsave_Click()
On Error GoTo HEREEE
Dim rss As ADODB.Recordset
If opttransport = True Then
savetransporters
loadBranchesTypes
Exit Sub
End If

If Optbranch = True Then
savestation
loadBranchesTypes
Exit Sub
End If
If lblCheckOff = True Then
   If txtSNo = "" Then
    MsgBox "Enter the SupplierNo  ", vbInformation, "CheckOff"
     Exit Sub
     End If
If txtSNo <> 0 Then
 sql = ""
sql = "select* from d_Suppliers where sno=" & txtSNo & " and freezed=1"
Set rss = oSaccoMaster.GetRecordset(sql)
If Not rss.EOF Then
MsgBox "The supplier is Freezed please Check Loan Balance first", vbInformation
Exit Sub
End If
End If
savesno
loadBranchesTypes
Exit Sub
End If
If optCash = True Then
    If TXTCHANGE < 0 Then
        If MsgBox("Insufficient Amount Received,do you want to transfer balance to check off ", vbYesNo) = vbYes Then
            lblCheckOff_Click
            lblCheckOff.value = True
            optCash.value = False
           Exit Sub
        Else
           Exit Sub
         End If
    End If
    savecash
    loadBranchesTypes
   Exit Sub
End If
'mpesa
If Optmpesa = True Then
    savempesa
    loadBranchesTypes
   Exit Sub
End If
If optstaff = True Then
savestaff
loadBranchesTypes
Exit Sub
End If
If Optothers = True Then
saveothers
Exit Sub
End If
HEREEE:
MsgBox err.description & " error occured."

End Sub

Private Sub savesno()
On Error GoTo ebraim
'Startdate & "','" & Enddate

Startdate = DateSerial(Year(txtransdate), Month(txtransdate), 1)
Enddate = DateSerial(Year(txtransdate), Month(txtransdate) + 1, 1 - 1)

If lblCheckOff = True Then

Dim a, b, X
DIA = 0
Dim U As Double, S As Double
Dim cn As Connection
Dim j As Integer
txtserialno_LostFocus
If DIA = 1 Then
Exit Sub
End If
If Lvwitems.ListItems.Count = 0 Then
MsgBox "There are no items sold."
Exit Sub
End If
j = 1

Dim total As Currency
total = 0
Do While Not j > (Lvwitems.ListItems.Count)
'For j = 1 To Lvwitems.ListItems.Count
 Lvwitems.ListItems.Item(j).selected = True
 total = total + CCur(Lvwitems.SelectedItem.SubItems(4))
j = j + 1
Loop

If optCash.value = False Then

'startdate = DateSerial(year(txtransdate), month(txtransdate), 1)
'Enddate = startdate + 19
'If lblNPay < 0 And txtransdate = Enddate Then
'MsgBox "Farmer Account Is negative"
'End If
'End Sub
Set rs = oSaccoMaster.GetRecordset("d_sp_IsClosed '" & Enddate & "'")
If Not rs.EOF Then
    MsgBox "The period " & Enddate & " has been closed by " & rs.Fields(0)
    Exit Sub
End If
'End If
Dim tar As Date
End If
'Startdate = DateSerial(year(txtransdate), month(txtransdate), 1)
'tar = Startdate + 19
'If lblNPay < 0 And txtransdate >= tar Then
'MsgBox "Farmer Account Is negative"
'Exit Sub
'End If
j = 1
For j = 1 To Lvwitems.ListItems.Count
'Do While Not j > (Lvwitems.ListItems.Count)
 Lvwitems.ListItems.Item(j).selected = True
If Trim(txtSNo) = "" Then
txtSNo = "0"
End If


'// sum total debt.
Dim xk As Double
Dim rstt As Recordset
sql = ""
sql = "set dateformat dmy select isnull(sum(amount),0)as amount from d_supplier_deduc where description='Agrovet' and startdate>='" & Startdate & "' and  enddate<='" & Enddate & "' and sno=" & txtSNo & ""

Set rstt = oSaccoMaster.GetRecordset(sql)
xk = rstt.Fields(0)
'// check if they are in stock.

Dim rsinstock As Recordset
sql = ""
sql = "select P_CODE,qin,qout,pprice, sprice,AI  from ag_products where p_code='" & Lvwitems.SelectedItem & "' AND  branch='" & Cmbstation & "'"

Set rsinstock = oSaccoMaster.GetRecordset(sql)

Dim Remain As Double, Bpprice As Double, spprice, r As Double, AI As Integer
r = Lvwitems.SelectedItem.SubItems(2)
Remain = (rsinstock.Fields(2) - (r))
Bpprice = rsinstock!pprice * CInt(Lvwitems.SelectedItem.SubItems(2))
spprice = rsinstock!sprice * CInt(Lvwitems.SelectedItem.SubItems(2))
AI = rsinstock!AI
If AI = 0 Then

sql = ""
sql = sql & "SET dateformat DMY INSERT INTO ag_Receipts(R_No, P_code, T_Date, Amount,S_no, Qua, S_Bal, user_id, Cash, SNo,Transby, Idno, mobile,branch,remarks,SPRICE,BPRICE,AI) VALUES ("
sql = sql & txtrno & ",'" & Lvwitems.SelectedItem & "','" & txtransdate & "'," & Lvwitems.SelectedItem.SubItems(4)
sql = sql & "," & txtSNo & "," & Lvwitems.SelectedItem.SubItems(2) & "," & Remain & ",'" & User & "'," & Lvwitems.SelectedItem.SubItems(5) & "," & txtSNo & ",'" & txttransby & "','" & txtidno & "','" & txtmobile & "','" & Cmbstation & "','" & Lvwitems.SelectedItem.SubItems(1) & "','" & spprice & "','" & Bpprice & "'," & txtserai & ")"

oSaccoMaster.ExecuteThis (sql)
sql = ""
sql = sql & "SET dateformat DMY INSERT INTO ag_Receipts3(R_No, P_code, T_Date, Amount,S_no, Qua, S_Bal, user_id, Cash, SNo,Transby, Idno, mobile,branch,AI) VALUES ("
sql = sql & txtrno & ",'" & Lvwitems.SelectedItem & "','" & txtransdate & "'," & Lvwitems.SelectedItem.SubItems(4)
sql = sql & "," & txtSNo & "," & Lvwitems.SelectedItem.SubItems(2) & "," & Remain & ",'" & User & "'," & Lvwitems.SelectedItem.SubItems(5) & "," & txtSNo & ",'" & txttransby & "','" & txtidno & "','" & txtmobile & "','" & Cmbstation & "'," & txtserai & ")"

oSaccoMaster.ExecuteThis (sql)
oSaccoMaster.ExecuteThis ("Update ag_Products SET Qout =" & CCur(Remain) & " WHERE p_code= '" & Lvwitems.SelectedItem & "' and branch='" & Cmbstation & "'")
oSaccoMaster.ExecuteThis ("Update Rcpno SET rcpno =" & txtrno & "")
'//XXXXXXXXXXXXXXX
    sql = ""
    sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & txtransdate & "'," & Lvwitems.SelectedItem.SubItems(4) & ",'" & lbldracc & "','" & lblcracc & "','" & Lvwitems.SelectedItem & "','" & cboproductname & "' ,'CHECK OFF SALES- " & "" & cboproductname & "','" & User & "',0,0)"
    oSaccoMaster.ExecuteThis (sql)

''sql = ""
'''amount = Bpprice
''    sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & txtransdate & "'," & Bpprice & ",'33-104','33-501','" & Lvwitems.SelectedItem & "','" & cboproductname & "' ,'CHECK OFF SALES- " & "" & cboproductname & "','" & User & "',0,0)"
''    oSaccoMaster.ExecuteThis (sql)

Else

sql = ""
sql = sql & "SET dateformat DMY INSERT INTO ag_Receipts(R_No, P_code, T_Date, Amount,S_no, Qua, S_Bal, user_id, Cash, SNo,Transby, Idno, mobile,branch,remarks,SPRICE,BPRICE,AI) VALUES ("
sql = sql & txtrno & ",'" & Lvwitems.SelectedItem & "','" & txtransdate & "'," & Lvwitems.SelectedItem.SubItems(4)
sql = sql & "," & txtSNo & "," & Lvwitems.SelectedItem.SubItems(2) & "," & Remain & ",'" & User & "'," & Lvwitems.SelectedItem.SubItems(5) & "," & txtSNo & ",'" & txttransby & "','" & txtidno & "','" & txtmobile & "','" & Cmbstation & "','" & Lvwitems.SelectedItem.SubItems(1) & "','" & spprice & "','" & Bpprice & "'," & txtserai & ")"

oSaccoMaster.ExecuteThis (sql)
sql = ""
sql = sql & "SET dateformat DMY INSERT INTO ag_Receipts3(R_No, P_code, T_Date, Amount,S_no, Qua, S_Bal, user_id, Cash, SNo,Transby, Idno, mobile,branch,AI) VALUES ("
sql = sql & txtrno & ",'" & Lvwitems.SelectedItem & "','" & txtransdate & "'," & Lvwitems.SelectedItem.SubItems(4)
sql = sql & "," & txtSNo & "," & Lvwitems.SelectedItem.SubItems(2) & "," & Remain & ",'" & User & "'," & Lvwitems.SelectedItem.SubItems(5) & "," & txtSNo & ",'" & txttransby & "','" & txtidno & "','" & txtmobile & "','" & Cmbstation & "'," & txtserai & ")"

oSaccoMaster.ExecuteThis (sql)
oSaccoMaster.ExecuteThis ("Update ag_Products SET Qout =" & CCur(Remain) & " WHERE p_code= '" & Lvwitems.SelectedItem & "' and branch='" & Cmbstation & "'")
oSaccoMaster.ExecuteThis ("Update Rcpno SET rcpno =" & txtrno & "")
'//XXXXXXXXXXXXXXX
    sql = ""
    sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & txtransdate & "'," & Lvwitems.SelectedItem.SubItems(4) & ",'" & lbldracc & "','" & lblcracc & "','" & Lvwitems.SelectedItem & "','" & cboproductname & "' ,'CHECK OFF SALES- " & "" & cboproductname & "','" & User & "',0,0)"
    oSaccoMaster.ExecuteThis (sql)

''sql = ""
'''amount = Bpprice
''    sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & txtransdate & "'," & Bpprice & ",'22-04','22-03','" & Lvwitems.SelectedItem & "','" & cboproductname & "' ,'CHECK OFF SALES- " & "" & cboproductname & "','" & User & "',0,0)"
''    oSaccoMaster.ExecuteThis (sql)
End If

'XXXXXXXXXXXXXXXXXXXXXX
'Next j
''j = j + 1
''Loop
'
'If optCash.value = False Then
'Set cn = New ADODB.Connection
'sql = "d_sp_SupplierDeduct " & txtSNo & ",'" & txtransdate & "','Agrovet'," & total & ",'" & Startdate & "','" & Enddate & "'," & year(txtransdate) & ",'" & User & "','" & Lvwitems.SelectedItem.SubItems(1) & "','" & Cmbstation & "'"
'oSaccoMaster.ExecuteThis (sql)
'End If
'
'If CDbl(txtamtreceived) > 0 Then
'    '******Deduct Amount paid in cash
'
'    amount = 0
'    amount = CDbl(txtamtreceived)
'    sql = "d_sp_SupplierDeduct " & txtSNo & ",'" & txtransdate & "','Agrovet'," & -1 * amount & ",'" & Startdate & "','" & Enddate & "'," & year(txtransdate) & ",'" & User & "','Cash','" & Cmbstation & "'"
'oSaccoMaster.ExecuteThis (sql)
Next j
'j = j + 1
'Loop

If optCash.value = False Then
Set cn = New ADODB.Connection
'sql = "d_sp_SupplierDeduct " & txtSNo & ",'" & txtransdate & "','Agrovet'," & total & ",'" & Startdate & "','" & Enddate & "'," & year(txtransdate) & ",'" & User & "','" & Lvwitems.SelectedItem.SubItems(1) & "','" & Cmbstation & "'"
'oSaccoMaster.ExecuteThis (sql)
End If
total = 0
j = 1
Do While Not j > (Lvwitems.ListItems.Count)
'For j = 1 To Lvwitems.ListItems.Count
 Lvwitems.ListItems.Item(j).selected = True
 total = total + CCur(Lvwitems.SelectedItem.SubItems(4))
 
 '*****differenciate AI and Agrovet********
'   Dim rsi As Recordset
'   sql = ""
'   sql = "select P_CODE,qin,qout,pprice, sprice,AI  from ag_products where p_code='" & Lvwitems.SelectedItem & "' AND  branch='" & Cmbstation & "'"
'   Set rsi = oSaccoMaster.GetRecordset(sql)
'   Dim AI As Integer
'   AI = rsinstock!AI
'  If AI = 0 Then
  sql = "d_sp_SupplierDeduct " & txtSNo & ",'" & txtransdate & "','Agrovet'," & CCur(Lvwitems.SelectedItem.SubItems(4)) & ",'" & Startdate & "','" & Enddate & "'," & Year(txtransdate) & ",'" & User & "','" & Lvwitems.SelectedItem.SubItems(1) & "','" & Cmbstation & "'"
  oSaccoMaster.ExecuteThis (sql)
'  Else
'  sql = "d_sp_SupplierDeduct " & txtSNo & ",'" & txtransdate & "'," & CCur(Lvwitems.SelectedItem.SubItems(4)) & ",'" & Startdate & "','" & Enddate & "'," & year(txtransdate) & ",'" & User & "','" & Lvwitems.SelectedItem.SubItems(1) & "','" & Cmbstation & "'," & txtserai & ""
'  oSaccoMaster.ExecuteThis (sql)
'  End If
j = j + 1
Loop

If CDbl(txtamtreceived) > 0 Then
    '******Deduct Amount paid in cash
   
    amount = 0
    amount = CDbl(txtamtreceived)
    AI = rsinstock!AI
    If AI = 0 Then
    sql = "d_sp_SupplierDeduct " & txtSNo & ",'" & txtransdate & "','Agrovet'," & -1 * amount & ",'" & Startdate & "','" & Enddate & "'," & Year(txtransdate) & ",'" & User & "','Cash','" & Cmbstation & "'"
    oSaccoMaster.ExecuteThis (sql)
    Else
    sql = "d_sp_SupplierDeduct " & txtSNo & ",'" & txtransdate & "','AI'," & -1 * amount & ",'" & Startdate & "','" & Enddate & "'," & Year(txtransdate) & ",'" & User & "','Cash','" & Cmbstation & "'," & txtserai & ""
    oSaccoMaster.ExecuteThis (sql)
    End If
'j =
End If

'//Update deductions
If chkPrint.value = vbChecked Then
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
       Dim fso, chkPrinter, txtFile
        'ttt = "LPT1" 'LPT1
         Dim PORT As String
        PORT = ports
        'ttt = "LPT1" 'LPT1
        ttt = PORT
        Set fso = CreateObject("Scripting.FileSystemObject")
        Dim strReceipts As String
        j = 1
        strReceipts = ""
        Do While Not j > (Lvwitems.ListItems.Count)
            Lvwitems.ListItems.Item(j).selected = True
            strReceipts = strReceipts & Lvwitems.SelectedItem.SubItems(1) & vbNewLine & Lvwitems.SelectedItem.SubItems(2) & vbTab & vbTab
            strReceipts = strReceipts & Format(Lvwitems.SelectedItem.SubItems(3), "#,##0.00") & vbTab & vbTab & Format(Lvwitems.SelectedItem.SubItems(4), "#,##0.00") & vbNewLine
            j = j + 1
        Loop

        'MsgBox strReceipts
        strReceipts = strReceipts & vbNewLine & "---------------------------------------" & vbNewLine
        strReceipts = strReceipts & "TOTAL" & vbTab & vbTab & vbTab & vbTab & Format(total, "#,##0.00") & vbNewLine
        strReceipts = strReceipts & "======================================="
        Set txtFile = fso.CreateTextFile(ttt, True)
    txtFile.WriteLine "     " & cname & ""
    txtFile.WriteLine "      " & paddress & ""
    txtFile.WriteLine "---------------------------------------"
    txtFile.WriteLine "    AGROVET RECEIPT"
    txtFile.WriteLine "     Check-off"
    txtFile.WriteLine "......................................."
    If lblCheckOff = True Then
    txtFile.WriteLine "SNo:" & txtSNo
    txtFile.WriteLine "Name:" & lblSNames
    End If
    txtFile.WriteLine "---------------------------------------"
'nAME QNTY PRICE AMOUNT
    txtFile.WriteLine "QNTY" & vbTab & vbTab & "PRICE" & vbTab & vbTab & "AMOUNT"
    txtFile.WriteLine "......................................."
    txtFile.WriteLine strReceipts
    txtFile.WriteLine
    txtFile.WriteLine "AMOUNT RECEVED" & vbTab & vbTab & txtamtreceived
    txtFile.WriteLine
    txtFile.WriteLine "CHANGE" & vbTab & vbTab & IIf(CDbl(TXTCHANGE) < 0, 0, CDbl(TXTCHANGE))
    txtFile.WriteLine
    txtFile.WriteLine "Current Credit:" & xk
    txtFile.WriteLine "Trans By" & vbTab & txttransby
    txtFile.WriteLine "Id No" & vbTab & txtidno
    txtFile.WriteLine
    txtFile.WriteLine "YOU WERE SERVED By " & UCase(username)
    txtFile.WriteLine "    Date :" & Format(Get_Server_Date, "dd/mm/yyyy HH:MM:SS AM/PM")
    txtFile.WriteLine "     THANK YOU AND WELCOME "
    txtFile.WriteLine "****************************************"
    txtFile.WriteLine escFeedAndCut
    txtFile.Close
End If
End If
'sms
 '************message***************'
              '//check settings  for sms alert
        'If SendSmsOnSalary = True Then
        'Dim phone As String
        Dim RsPhoneNumbers As New ADODB.Recordset
        Dim MsgContent As String
        Set RsPhoneNumbers = oSaccoMaster.GetRecordset("select  PhoneNo  from d_Suppliers  where SNo ='" & txtSNo & "' and PhoneNo  is not null and PhoneNo <>'' and LEN(PhoneNo) =10 ")

        If Not RsPhoneNumbers.EOF Then
            Phone = Trim(RsPhoneNumbers.Fields!PhoneNo & "")
        Else
            Phone = ""
        End If


            '//insert into sms alert
'        Set rs = New ADODB.Recordset
'        sql = "d_sp_TotalMonth " & txtSNo & ",'" & Startdate & "','" & Enddate & "'"
'        Set rs = oSaccoMaster.GetRecordset(sql)
'        If Not rs.EOF Then
'        If Not IsNull(rs.Fields(0)) Then CummulKgs = rs.Fields(0)
'        Else
'        CummulKgs = "0.00"
'        End If

            If Trim(Phone) <> "" Then
            MsgContent = ""
            MsgContent = "Dear supplier, You have been issued with  items worth ," & TXTTOTAL & " :From  olenguruone Dairy Agrovet"
            strSQL = ""
            strSQL = "INSERT INTO Swift_Messages(SaccoCode,PhoneNumber,Message, msgstatus,Auditdate)"
            strSQL = strSQL & "Values (14,'" & Phone & "','" & MsgContent & "',0,'" & Date & "')"

            oSaccoMaster.ExecuteThis (strSQL)
            End If

Lvwitems.ListItems.Clear
txtpcode.Text = ""
txtquantity = ""
txtamount = ""
cboproductname = ""
txtrno = ""
txtSNo = ""
lblTKgs = ""
lblGPay = ""
lblDed = ""
lblNPay = ""
lblSNames = ""
txttransby = ""
txtidno = ""
txtmobile = ""
'xk = ""
cmdNew_Click
MsgBox "Records saved"
Exit Sub
ebraim:
MsgBox err.description & " error occured."

End Sub
Private Sub savetransporters()
On Error GoTo kiparu2

Dim Startdate As Date
Dim Enddate As Date

Startdate = DateSerial(Year(txtransdate), Month(txtransdate), 1)
'Enddate = DateSerial(year(DTPDDeduction), month(DTPDDeduction) + 1, 1 - 1)
If opttransport = True Then
If txttranscode = "" Then
MsgBox "Please enter the transporter"
Exit Sub
End If
'txttranscode
'If txttranscode <> 0 Then
Dim rss As New Recordset
 sql = ""
sql = "select* from d_transporters where transcode='" & txttranscode & "' and freezed=1"
Set rss = oSaccoMaster.GetRecordset(sql)
If Not rss.EOF Then
MsgBox "The Transporter is Freezed please Check Loan Balance first", vbInformation
Exit Sub
End If
'End
'dairy sno
'// sum total debt.
Dim xkk As Double
Dim rstt As Recordset
sql = ""
sql = "set dateformat dmy select certno as id from d_transporters where transcode ='" & txttranscode & "'"

Set rstt = oSaccoMaster.GetRecordset(sql)
xkk = IIf(IsNull(rstt.Fields(0)), 0, rstt.Fields(0))


'checksno
Dim kk As Double
Dim yy As Double
kk = xkk
'yy = ""
'// sum total debt.
Dim yk As Double
Dim rsk As Recordset
'sql = ""
'sql = "set dateformat dmy select idno as idno,sno from d_suppliers where idno ='" & xkk & "'"
'
'Set rsk = oSaccoMaster.GetRecordset(sql)
'yk = IIf(IsNull(rsk.Fields(0)), 0, rsk.Fields(0))
'Label26.Caption = IIf(IsNull(rsk.Fields(1)), 0, rsk.Fields(1))
'yy = yk
'If kk = yy Then
'If MsgBox("Transporter number " & txttranscode & " has a another supply numbet " & Label26 & " do you wish to proceed?", vbYesNo) = vbYes Then
'Else
''MsgBox "Transporter is also asupplier"
'Exit Sub
'End If
'End If
Set rst = New Recordset
Dim a, b, X
DIA = 0
Dim U As Double, S As Double
Dim cn As Connection
Dim j As Integer
txtserialno_LostFocus
If DIA = 1 Then
Exit Sub
End If
If Lvwitems.ListItems.Count = 0 Then
MsgBox "There are no items sold."
Exit Sub
End If
j = 1

Dim total As Currency
total = 0
Do While Not j > (Lvwitems.ListItems.Count)
 Lvwitems.ListItems.Item(j).selected = True
 total = total + CCur(Lvwitems.SelectedItem.SubItems(4))
j = j + 1
Loop

If optCash.value = False Then
If total > CCur(lbltransnetpay) Then
If MsgBox("Transporter number " & txttranscode & " has a netpay of " & lblNPay & " do you wish to proceed?", vbYesNo) = vbYes Then
Else
Exit Sub
End If
End If


Startdate = DateSerial(Year(txtransdate), Month(txtransdate), 1)
Enddate = DateSerial(Year(txtransdate), Month(txtransdate) + 1, 1 - 1)


Set rs = oSaccoMaster.GetRecordset("d_sp_IsClosed '" & Enddate & "'")
If Not rs.EOF Then
    MsgBox "The period " & Enddate & " has been closed by " & rs.Fields(0)
    Exit Sub
End If
Startdate = DateSerial(Year(txtransdate), Month(txtransdate), 1)
Enddate = Startdate + 19
If lbltransnetpay < 0 And txtransdate = Enddate Then
MsgBox "Transpoter Account Is negative"
Exit Sub
End If


'End If
End If
j = 1
For j = 1 To Lvwitems.ListItems.Count
'Do While Not j > (Lvwitems.ListItems.Count)
 Lvwitems.ListItems.Item(j).selected = True
If Trim(txttranscode) = "" Then
txttranscode = "0"
End If
'// check if they are in stock.

Dim rsinstock As Recordset
sql = ""
'sql = "select P_CODE,qin,qout from ag_products where p_code='" & Lvwitems.SelectedItem & "'"
sql = "select P_CODE,qin,qout,pprice, sprice,AI  from ag_products where p_code='" & Lvwitems.SelectedItem & "' AND  branch='" & Cmbstation & "'"
'//Set rsinstock = New ADODB.Recordset
Set rsinstock = oSaccoMaster.GetRecordset(sql)
Dim Remain As Double, Bpprice As Double, spprice As Double, AI As Integer
Remain = rsinstock.Fields(2) - CInt(Lvwitems.SelectedItem.SubItems(2))
Bpprice = rsinstock!pprice * CInt(Lvwitems.SelectedItem.SubItems(2))
spprice = rsinstock!sprice * CInt(Lvwitems.SelectedItem.SubItems(2))
AI = rsinstock!AI

'// Where it differentiate AI and Agrovet products
If AI = 0 Then

'// check the stock if it is less than zero
If rsinstock.Fields(2) <= 0 Then
MsgBox "Sorry Stock is Zero for item " & Lvwitems.SelectedItem.SubItems(1) & " please re-stock before your proceed", vbInformation
Exit Sub
End If
'// check the quanttity being sold versus the balance
'Dim Remain As Double
Remain = rsinstock.Fields(2) - CInt(Lvwitems.SelectedItem.SubItems(2))
If Remain < 0 Then
MsgBox "Stock will be negative " & Remain & " please re-stock before you proceed", vbInformation
Exit Sub
End If

sql = ""
sql = sql & "SET dateformat DMY INSERT INTO ag_Receipts(R_No, P_code, T_Date, Amount, Qua, S_Bal, user_id, Cash, SNo,Transby, Idno,mobile,branch,remarks,SPRICE,BPRICE,AI) VALUES ("
sql = sql & txtrno & ",'" & Lvwitems.SelectedItem & "','" & txtransdate & "'," & Lvwitems.SelectedItem.SubItems(4)
sql = sql & "," & Lvwitems.SelectedItem.SubItems(2) & "," & Remain & ",'" & User & "'," & Lvwitems.SelectedItem.SubItems(5) & ",'" & txttranscode & "','" & txttransby & "','" & txtidno & "','" & txtmobile & "','" & Cmbstation & "','" & Lvwitems.SelectedItem.SubItems(1) & "','" & spprice & "','" & Bpprice & "'," & txtserai & ")"

oSaccoMaster.ExecuteThis (sql)
sql = ""
sql = sql & "SET dateformat DMY INSERT INTO ag_Receipts3(R_No, P_code, T_Date, Amount, Qua, S_Bal, user_id, Cash, SNo,Transby, Idno,mobile,branch,AI) VALUES ("
sql = sql & txtrno & ",'" & Lvwitems.SelectedItem & "','" & txtransdate & "'," & Lvwitems.SelectedItem.SubItems(4)
sql = sql & "," & Lvwitems.SelectedItem.SubItems(2) & "," & Remain & ",'" & User & "'," & Lvwitems.SelectedItem.SubItems(5) & ",'" & txttranscode & "','" & txttransby & "','" & txtidno & "','" & txtmobile & "','" & Cmbstation & "'," & txtserai & ")"
'
oSaccoMaster.ExecuteThis (sql)
oSaccoMaster.ExecuteThis ("Update ag_Products SET Qout =" & CCur(Remain) & " WHERE p_code= '" & Lvwitems.SelectedItem & "' and branch='" & Cmbstation & "'")
'j = j + 1
'Loop
'    sql = ""
'    sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & txtransdate & "'," & Lvwitems.SelectedItem.SubItems(4) & ",'A007','I005','" & Lvwitems.SelectedItem & "','" & cboproductname & "' ,'TRANSPORTERS SALES','" & User & "',0,0)"
'    oSaccoMaster.ExecuteThis (sql)
     sql = ""
    sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & txtransdate & "'," & Lvwitems.SelectedItem.SubItems(4) & ",'" & lbldracc & "','" & lblcracc & "','" & Lvwitems.SelectedItem & "','" & cboproductname & "' ,'TRANSPORTERS SALES- " & "" & cboproductname & "','" & User & "',1,0)"
    oSaccoMaster.ExecuteThis (sql)
''    sql = ""
''    sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & txtransdate & "','" & Bpprice & "','33-104','33-501','" & Lvwitems.SelectedItem & "','" & cboproductname & "' ,'TRANSPORTERS SALES- " & "" & cboproductname & "','" & User & "',1,0)"
''    oSaccoMaster.ExecuteThis (sql)
    
Else

'// check the stock if it is less than zero
If rsinstock.Fields(2) <= 0 Then
MsgBox "Sorry Stock is Zero for item " & Lvwitems.SelectedItem.SubItems(1) & " please re-stock before your proceed", vbInformation
Exit Sub
End If
'// check the quanttity being sold versus the balance
'Dim Remain As Double
Remain = rsinstock.Fields(2) - CInt(Lvwitems.SelectedItem.SubItems(2))
If Remain < 0 Then
MsgBox "Stock will be negative " & Remain & " please re-stock before you proceed", vbInformation
Exit Sub
End If

sql = ""
sql = sql & "SET dateformat DMY INSERT INTO ag_Receipts(R_No, P_code, T_Date, Amount, Qua, S_Bal, user_id, Cash, SNo,Transby, Idno,mobile,branch,remarks,SPRICE,BPRICE,AI) VALUES ("
sql = sql & txtrno & ",'" & Lvwitems.SelectedItem & "','" & txtransdate & "'," & Lvwitems.SelectedItem.SubItems(4)
sql = sql & "," & Lvwitems.SelectedItem.SubItems(2) & "," & Remain & ",'" & User & "'," & Lvwitems.SelectedItem.SubItems(5) & ",'" & txttranscode & "','" & txttransby & "','" & txtidno & "','" & txtmobile & "','" & Cmbstation & "','" & Lvwitems.SelectedItem.SubItems(1) & "','" & spprice & "','" & Bpprice & "'," & txtserai & ")"

oSaccoMaster.ExecuteThis (sql)
sql = ""
sql = sql & "SET dateformat DMY INSERT INTO ag_Receipts3(R_No, P_code, T_Date, Amount, Qua, S_Bal, user_id, Cash, SNo,Transby, Idno,mobile,branch,AI) VALUES ("
sql = sql & txtrno & ",'" & Lvwitems.SelectedItem & "','" & txtransdate & "'," & Lvwitems.SelectedItem.SubItems(4)
sql = sql & "," & Lvwitems.SelectedItem.SubItems(2) & "," & Remain & ",'" & User & "'," & Lvwitems.SelectedItem.SubItems(5) & ",'" & txttranscode & "','" & txttransby & "','" & txtidno & "','" & txtmobile & "','" & Cmbstation & "'," & txtserai & ")"
'
oSaccoMaster.ExecuteThis (sql)
oSaccoMaster.ExecuteThis ("Update ag_Products SET Qout =" & CCur(Remain) & " WHERE p_code= '" & Lvwitems.SelectedItem & "' and branch='" & Cmbstation & "'")
'j = j + 1
'Loop
'    sql = ""
'    sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & txtransdate & "'," & Lvwitems.SelectedItem.SubItems(4) & ",'A007','I005','" & Lvwitems.SelectedItem & "','" & cboproductname & "' ,'TRANSPORTERS SALES','" & User & "',0,0)"
'    oSaccoMaster.ExecuteThis (sql)
     sql = ""
    sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & txtransdate & "'," & Lvwitems.SelectedItem.SubItems(4) & ",'" & lbldracc & "','" & lblcracc & "','" & Lvwitems.SelectedItem & "','" & cboproductname & "' ,'TRANSPORTERS SALES- " & "" & cboproductname & "','" & User & "',1,0)"
    oSaccoMaster.ExecuteThis (sql)
'    sql = ""
'    sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & txtransdate & "','" & Bpprice & "','22-04','22-03','" & Lvwitems.SelectedItem & "','" & cboproductname & "' ,'TRANSPORTERS SALES- " & "" & cboproductname & "','" & User & "',1,0)"
'    oSaccoMaster.ExecuteThis (sql)

End If

Next j
'//Update deductions
If optCash.value = False Then
Set cn = New ADODB.Connection
'sql = "d_sp_TransDeduct '" & txttranscode & "','" & txtransdate & "','Agrovet'," & total & ",'" & Startdate & "','" & Enddate & "','" & User & "','" & Lvwitems.SelectedItem.SubItems(1) & "'"
'oSaccoMaster.ExecuteThis (sql)
End If

If CDbl(txtamtreceived) > 0 Then
amount = 0
amount = CDbl(txtamtreceived) * 1
Set cn = New ADODB.Connection
AI = rsinstock!AI
'// where it differentiate AI and Agrovet
If AI = 0 Then
sql = "d_sp_TransDeduct '" & txttranscode & "','" & txtransdate & "','Agrovet'," & -1 * amount & ",'" & Startdate & "','" & Enddate & "','" & User & "','Cash'," & txtserai & ""
oSaccoMaster.ExecuteThis (sql)
Else
sql = "d_sp_TransDeduct '" & txttranscode & "','" & txtransdate & "','AI'," & -1 * amount & ",'" & Startdate & "','" & Enddate & "','" & User & "','Cash'," & txtserai & ""
oSaccoMaster.ExecuteThis (sql)
End If
End If
j = 1
For j = 1 To Lvwitems.ListItems.Count
'Do While Not j > (Lvwitems.ListItems.Count)
 Lvwitems.ListItems.Item(j).selected = True
If Trim(txttranscode) = "" Then
txttranscode = "0"
Else
AI = rsinstock!AI
'// where it differentiate AI and Agrovet
If AI = 0 Then
sql = "d_sp_TransDeduct '" & txttranscode & "','" & txtransdate & "','Agrovet'," & Lvwitems.SelectedItem.SubItems(4) & ",'" & Startdate & "','" & Enddate & "','" & User & "','" & Lvwitems.SelectedItem.SubItems(1) & "'," & txtserai & ""
oSaccoMaster.ExecuteThis (sql)
Else
sql = "d_sp_TransDeduct '" & txttranscode & "','" & txtransdate & "','AI'," & Lvwitems.SelectedItem.SubItems(4) & ",'" & Startdate & "','" & Enddate & "','" & User & "','" & Lvwitems.SelectedItem.SubItems(1) & "'," & txtserai & ""
oSaccoMaster.ExecuteThis (sql)
End If
End If
Next j
If chkPrint.value = vbChecked Then
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
       Dim fso, chkPrinter, txtFile
        Dim PORT As String
        PORT = ports
        ttt = PORT 'LPT1
        
        Set fso = CreateObject("Scripting.FileSystemObject")
        Dim strReceipts As String
        j = 1
        strReceipts = ""
        Do While Not j > (Lvwitems.ListItems.Count)
            Lvwitems.ListItems.Item(j).selected = True
            strReceipts = strReceipts & Lvwitems.SelectedItem.SubItems(1) & vbNewLine & Lvwitems.SelectedItem.SubItems(2) & vbTab & vbTab
            strReceipts = strReceipts & Format(Lvwitems.SelectedItem.SubItems(3), "#,##0.00") & vbTab & vbTab & Format(Lvwitems.SelectedItem.SubItems(4), "#,##0.00") & vbNewLine
            j = j + 1
        Loop

        'MsgBox strReceipts
        strReceipts = strReceipts & vbNewLine & "---------------------------------------" & vbNewLine
        strReceipts = strReceipts & "TOTAL" & vbTab & vbTab & vbTab & vbTab & Format(total, "#,##0.00") & vbNewLine
        strReceipts = strReceipts & "======================================="
                 
        Set txtFile = fso.CreateTextFile(ttt, True)
    txtFile.WriteLine "      " & cname & ""
    txtFile.WriteLine "       " & paddress & ""
    txtFile.WriteLine "---------------------------------------"
    txtFile.WriteLine "      AGROVET RECEIPT"
    txtFile.WriteLine "          Check-off"
    txtFile.WriteLine "---------------------------------------"
    txtFile.WriteLine "TCode:" & txttranscode
    txtFile.WriteLine "Name:" & lbltransportername
    
'NAME QNTY PRICE AMOUNT
    txtFile.WriteLine "QNTY" & vbTab & vbTab & "PRICE" & vbTab & vbTab & "AMOUNT"
    txtFile.WriteLine "......................................."
    txtFile.WriteLine strReceipts
        txtFile.WriteLine
    txtFile.WriteLine "TOTAL" & TXTTOTAL
    txtFile.WriteLine
    txtFile.WriteLine "AMOUNT RECEVED" & txtamtreceived
    txtFile.WriteLine
    txtFile.WriteLine "CHANGE" & TXTCHANGE
    txtFile.WriteLine
    txtFile.WriteLine "YOU WERE SERVED By " & UCase(username)
    txtFile.WriteLine "    Date :" & Format(Get_Server_Date, "dd/mm/yyyy HH:MM:SS AM/PM")
    txtFile.WriteLine "     THANK YOU AND WELCOME "
    txtFile.WriteLine "****************************************"
    txtFile.WriteLine escFeedAndCut
    txtFile.Close
End If



Lvwitems.ListItems.Clear
txttranscode = ""
txtrno.Text = ""
txtpcode.Text = ""
'txtserialno = ""
lbltransnetpay = ""
txtquantity = 1
txtamount = ""
 
MsgBox "Records saved"
Exit Sub
kiparu2:
MsgBox err.description & " error occured."
End If

End Sub

Private Sub savestation()
On Error GoTo kiparu

If Optbranch = True Then
Dim centre As String
centre = Cmbstation.Text
If Trim(Cmbstation.Text) = "" Then
 MsgBox "Please enter the Agrovet Station."
   Cmbstation.SetFocus
Exit Sub
End If


Dim j As Integer
If Lvwitems.ListItems.Count = 0 Then
MsgBox "There are no items sold."
Exit Sub
End If
j = 1

Dim total As Currency, pprice As Currency
total = 0
Do While Not j > (Lvwitems.ListItems.Count)
 Lvwitems.ListItems.Item(j).selected = True
 total = total + CCur(Lvwitems.SelectedItem.SubItems(4))
j = j + 1
Loop


Startdate = DateSerial(Year(txtransdate), Month(txtransdate), 1)
Enddate = DateSerial(Year(txtransdate), Month(txtransdate) + 1, 1 - 1)


Set rs = oSaccoMaster.GetRecordset("d_sp_IsClosed '" & Enddate & "'")
If Not rs.EOF Then
    MsgBox "The period " & Enddate & " has been closed by " & rs.Fields(0)
    Exit Sub
End If
'End If
'End If
For j = 1 To Lvwitems.ListItems.Count
'Do While Not j > (Lvwitems.ListItems.Count)
 Lvwitems.ListItems.Item(j).selected = True

'// check if they are in stock.

Dim rsinstock As Recordset
sql = ""
'sql = "select P_CODE,qin,qout,PPrice from ag_products where p_code='" & Lvwitems.SelectedItem & "'"
sql = "select P_CODE,qin,qout,pprice, sprice,AI  from ag_products where p_code='" & Lvwitems.SelectedItem & "' AND  branch='" & Cmbstation & "'"
'//Set rsinstock = New ADODB.Recordset
Set rsinstock = oSaccoMaster.GetRecordset(sql)
Dim Remain As Double, Bpprice As Double, spprice As Double, AI As Integer
Remain = rsinstock.Fields(2) - CInt(Lvwitems.SelectedItem.SubItems(2))
Bpprice = rsinstock!pprice * CInt(Lvwitems.SelectedItem.SubItems(2))
spprice = rsinstock!sprice * CInt(Lvwitems.SelectedItem.SubItems(2))
AI = rsinstock!AI
'//where it differentiate AI and Agrovet
If AI = 0 Then

'// check the stock if it is less than zero
If rsinstock.Fields(2) <= 0 Then
MsgBox "Sorry Stock is Zero for item " & Lvwitems.SelectedItem.SubItems(1) & " please re-stock before your proceed", vbInformation
Exit Sub
End If
'// check the quanttity being sold versus the balance
'Dim Remain As Double
Remain = rsinstock.Fields(2) - CInt(Lvwitems.SelectedItem.SubItems(2))
If Remain < 0 Then
MsgBox "Stock will be negative " & Remain & " please re-stock before you proceed", vbInformation
Exit Sub
End If
amount = 0
amount = Lvwitems.SelectedItem.SubItems(3) * Lvwitems.SelectedItem.SubItems(2)
sql = ""
sql = sql & "SET dateformat DMY INSERT INTO ag_Receipts(R_No, P_code, T_Date, Amount, Qua, S_Bal, user_id, Cash, SNo,branch) VALUES ("
sql = sql & txtrno & ",'" & Lvwitems.SelectedItem & "','" & txtransdate & "'," & amount
sql = sql & "," & Lvwitems.SelectedItem.SubItems(2) & "," & Remain & ",'" & User & "'," & Lvwitems.SelectedItem.SubItems(5) & ",'" & centre & "','" & Cmbstation & "')"

oSaccoMaster.ExecuteThis (sql)
sql = ""
sql = sql & "SET dateformat DMY INSERT INTO ag_Receipts3(R_No, P_code, T_Date, Amount, Qua, S_Bal, user_id, Cash, SNo,branch) VALUES ("
sql = sql & txtrno & ",'" & Lvwitems.SelectedItem & "','" & txtransdate & "'," & amount
sql = sql & "," & Lvwitems.SelectedItem.SubItems(2) & "," & Remain & ",'" & User & "'," & Lvwitems.SelectedItem.SubItems(5) & ",'" & centre & "','" & Cmbstation & "')"

oSaccoMaster.ExecuteThis (sql)
oSaccoMaster.ExecuteThis ("Update ag_Products SET Qout =" & CCur(Remain) & " WHERE p_code= '" & Lvwitems.SelectedItem & "' and branch='" & Cmbstation & "'")

Dim DRaccno As String
Dim Craccno As String
If centre = "SANGALO" Then
    DRaccno = "A008"
    Craccno = "I004"
ElseIf centre = "OLMAROROI" Then
    DRaccno = "A010"
    Craccno = "I005"
ElseIf centre = "KABISAGA" Then
    DRaccno = "A012"
    Craccno = "I006"
ElseIf centre = "KOISOLIK" Then
    DRaccno = "A009"
    Craccno = "I007"
ElseIf centre = "CHEMUSWO" Then
    DRaccno = "A011"
    Craccno = "I008"
ElseIf centre = "BELEKENYA" Then
    DRaccno = "A013"
End If
'XXXXXXXXXXX SAVE TO GL
'    sql = ""
'    sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & txtransdate & "'," & amount & ",'" & DRaccno & "','" & Craccno & "','" & Lvwitems.SelectedItem & "','" & cboproductname & "' ,' CHECK OFF SALES ','" & User & "',0,0)"
'    oSaccoMaster.ExecuteThis (sql)
'
    
     sql = ""
    sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & txtransdate & "'," & Lvwitems.SelectedItem.SubItems(4) & ",'" & lbldracc & "','" & lblcracc & "','" & Lvwitems.SelectedItem & "','" & cboproductname & "' ,'Staff Sales- " & "" & cboproductname & "','" & User & "',1,0)"
    oSaccoMaster.ExecuteThis (sql)
'    sql = ""
'    sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & txtransdate & "','" & Bpprice & "','33-104','33-501','" & Lvwitems.SelectedItem & "','" & cboproductname & "' ,'Staff Sales- " & "" & cboproductname & "','" & User & "',1,0)"
'    oSaccoMaster.ExecuteThis (sql)


'XXXXXXXXXXXXXXXXXXXXXX

Else

'// check the stock if it is less than zero
If rsinstock.Fields(2) <= 0 Then
MsgBox "Sorry Stock is Zero for item " & Lvwitems.SelectedItem.SubItems(1) & " please re-stock before your proceed", vbInformation
Exit Sub
End If
'// check the quanttity being sold versus the balance
'Dim Remain As Double
Remain = rsinstock.Fields(2) - CInt(Lvwitems.SelectedItem.SubItems(2))
If Remain < 0 Then
MsgBox "Stock will be negative " & Remain & " please re-stock before you proceed", vbInformation
Exit Sub
End If
amount = 0
amount = Lvwitems.SelectedItem.SubItems(3) * Lvwitems.SelectedItem.SubItems(2)
sql = ""
sql = sql & "SET dateformat DMY INSERT INTO ag_Receipts(R_No, P_code, T_Date, Amount, Qua, S_Bal, user_id, Cash, SNo,branch) VALUES ("
sql = sql & txtrno & ",'" & Lvwitems.SelectedItem & "','" & txtransdate & "'," & amount
sql = sql & "," & Lvwitems.SelectedItem.SubItems(2) & "," & Remain & ",'" & User & "'," & Lvwitems.SelectedItem.SubItems(5) & ",'" & centre & "','" & Cmbstation & "')"

oSaccoMaster.ExecuteThis (sql)
sql = ""
sql = sql & "SET dateformat DMY INSERT INTO ag_Receipts3(R_No, P_code, T_Date, Amount, Qua, S_Bal, user_id, Cash, SNo,branch) VALUES ("
sql = sql & txtrno & ",'" & Lvwitems.SelectedItem & "','" & txtransdate & "'," & amount
sql = sql & "," & Lvwitems.SelectedItem.SubItems(2) & "," & Remain & ",'" & User & "'," & Lvwitems.SelectedItem.SubItems(5) & ",'" & centre & "','" & Cmbstation & "')"

oSaccoMaster.ExecuteThis (sql)
oSaccoMaster.ExecuteThis ("Update ag_Products SET Qout =" & CCur(Remain) & " WHERE p_code= '" & Lvwitems.SelectedItem & "' and branch='" & Cmbstation & "'")

'Dim DRaccno As String
'Dim Craccno As String
If centre = "SANGALO" Then
    DRaccno = "A008"
    Craccno = "I004"
ElseIf centre = "OLMAROROI" Then
    DRaccno = "A010"
    Craccno = "I005"
ElseIf centre = "KABISAGA" Then
    DRaccno = "A012"
    Craccno = "I006"
ElseIf centre = "KOISOLIK" Then
    DRaccno = "A009"
    Craccno = "I007"
ElseIf centre = "CHEMUSWO" Then
    DRaccno = "A011"
    Craccno = "I008"
ElseIf centre = "BELEKENYA" Then
    DRaccno = "A013"
End If
'XXXXXXXXXXX SAVE TO GL
'    sql = ""
'    sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & txtransdate & "'," & amount & ",'" & DRaccno & "','" & Craccno & "','" & Lvwitems.SelectedItem & "','" & cboproductname & "' ,' CHECK OFF SALES ','" & User & "',0,0)"
'    oSaccoMaster.ExecuteThis (sql)
'
    
     sql = ""
    sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & txtransdate & "'," & Lvwitems.SelectedItem.SubItems(4) & ",'" & lbldracc & "','" & lblcracc & "','" & Lvwitems.SelectedItem & "','" & cboproductname & "' ,'STATION Sales- " & "" & cboproductname & "','" & User & "',1,0)"
    oSaccoMaster.ExecuteThis (sql)
    sql = ""
    sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & txtransdate & "','" & Bpprice & "','22-04','22-03','" & Lvwitems.SelectedItem & "','" & cboproductname & "' ,'STATION Sales- " & "" & cboproductname & "','" & User & "',1,0)"
    oSaccoMaster.ExecuteThis (sql)


'XXXXXXXXXXXXXXXXXXXXXX
End If

Next j



If chkPrint.value = vbChecked Then
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
        
        
        
        
        Dim fso, chkPrinter, txtFile
        'ttt = "LPT1" 'LPT1
         Dim PORT As String
        PORT = ports
        'ttt = "LPT1" 'LPT1
        ttt = PORT
        Set fso = CreateObject("Scripting.FileSystemObject")
        Dim strReceipts As String
        j = 1
        
        
        strReceipts = ""
        Do While Not j > (Lvwitems.ListItems.Count)
            Lvwitems.ListItems.Item(j).selected = True
            strReceipts = strReceipts & Lvwitems.SelectedItem.SubItems(1) & vbNewLine & Lvwitems.SelectedItem.SubItems(2) & vbTab & vbTab
            strReceipts = strReceipts & Format(Lvwitems.SelectedItem.SubItems(3), "#,##0.00") & vbTab & vbTab & Format(Lvwitems.SelectedItem.SubItems(4), "#,##0.00") & vbNewLine
            j = j + 1
        Loop

        'MsgBox strReceipts
        strReceipts = strReceipts & vbNewLine & "---------------------------------------" & vbNewLine
        strReceipts = strReceipts & "TOTAL" & vbTab & vbTab & vbTab & vbTab & Format(total, "#,##0.00") & vbNewLine
        strReceipts = strReceipts & "======================================="
        Set txtFile = fso.CreateTextFile(ttt, True)
        
        txtFile.WriteLine "  " & cname & ""
        txtFile.WriteLine "     " & paddress & ""
        txtFile.WriteLine "     " & town & ""
        txtFile.WriteLine "     " & Phone & ""
        'txtfile.WriteLine "     " & Email & ""
        txtFile.WriteLine "---------------------------------------"
        txtFile.WriteLine "    AGROVET RECEIPT"
        txtFile.WriteLine "  STOCK DISPATCHED TO " & centre & ""
        txtFile.WriteLine "---------------------------------------" '
        txtFile.WriteLine "QNTY" & vbTab & vbTab & "PRICE" & vbTab & vbTab & "AMOUNT"
        txtFile.WriteLine "......................................."
        txtFile.WriteLine strReceipts
        txtFile.WriteLine
        txtFile.WriteLine "TOTAL" & vbTab & vbTab & vbTab & vbTab & Format(total, "#,##0.00") & vbNewLine
        txtFile.WriteLine
        txtFile.WriteLine "TOTAL" & vbTab & TXTTOTAL
        txtFile.WriteLine
        txtFile.WriteLine "YOU WERE SERVED By " & UCase(username)
        txtFile.WriteLine "    Date :" & Format(Get_Server_Date, "dd/mm/yyyy HH:MM:SS AM/PM")
        txtFile.WriteLine " Stock Dispatched to " & centre & "at the selling price"
        txtFile.WriteLine "     THANK YOU AND WELCOME "
        txtFile.WriteLine "****************************************"
        txtFile.WriteLine escFeedAndCut
        txtFile.Close
    End If
End If
'//Update deductions
'If optCash.value = False Then
'Set cn = New ADODB.Connection
'sql = "d_sp_TransDeduct '" & txttranscode & "','" & txtransdate & "','Agrovet'," & Total & ",'" & Startdate & "','" & Enddate & "','" & User & "'"
'oSaccoMaster.ExecuteThis (sql)

'Set cn = New ADODB.Connection
'sql = "d_sp_TransDeduct '" & txtTCode & "','" & DTPDDate & "','" & cboDeductionType & "'," & txtamount & ",'" & DTPStartDate & "','" & DTPEndDate & "','" & User & "'"
'oSaccoMaster.ExecuteThis (sql)


''XXXXXXXXXXXXXXXXXXXXXXXxx
''\\ save to gl
'
'
'    sql = ""
'    sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & txtransdate & "'," & txtquantity & " *" & txtPPrice & ",'" & lbldracc & "','" & lblcracc & "','stock intake','" & cbosupplier & "' ,'stock intake','" & User & "',0,0)"
'    oSaccoMaster.ExecuteThis (sql)
''
'
''XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Lvwitems.ListItems.Clear
txtrno = ""
txtpcode.Text = ""
txtquantity = 1
txtamount = ""
Cmbstation.Text = ""

MsgBox "Record saved Successfully"
Exit Sub
kiparu:
MsgBox err.description & " error occured."
End Sub
Private Sub savestaff()
On Error GoTo olkalou

If optstaff = True Then
Dim C As String
Dim D As String
C = "Staff" & txtstaffno
D = lblstnames
Dim j As Integer
If Lvwitems.ListItems.Count = 0 Then
MsgBox "There are no items sold."
Exit Sub
End If
If txtstaffno = "" Then
MsgBox "Enter Staff Number before you continue", vbCritical, "Maziwa"

Exit Sub
End If
j = 1

Dim total As Currency
total = 0
Do While Not j > (Lvwitems.ListItems.Count)
 Lvwitems.ListItems.Item(j).selected = True
 total = total + CCur(Lvwitems.SelectedItem.SubItems(4))
j = j + 1
Loop



Startdate = DateSerial(Year(txtransdate), Month(txtransdate), 1)
Enddate = DateSerial(Year(txtransdate), Month(txtransdate) + 1, 1 - 1)

Set rs = oSaccoMaster.GetRecordset("d_sp_IsClosed '" & Enddate & "'")
If Not rs.EOF Then
    MsgBox "The period " & Enddate & " has been closed by " & rs.Fields(0)
    Exit Sub
End If


'// check if they are in stock.
For j = 1 To Lvwitems.ListItems.Count
 Lvwitems.ListItems.Item(j).selected = True


Dim rsinstock As Recordset
sql = ""
sql = "select P_CODE,qin,qout,pprice, sprice,AI from ag_products where p_code='" & Lvwitems.SelectedItem & "'AND  branch='" & Cmbstation & "'"
'//Set rsinstock = New ADODB.Recordset
Set rsinstock = oSaccoMaster.GetRecordset(sql)
Dim Remain As Double, Bpprice As Double, spprice As Double, AI As Integer
Remain = rsinstock.Fields(2) - CInt(Lvwitems.SelectedItem.SubItems(2))
Bpprice = rsinstock!pprice * CInt(Lvwitems.SelectedItem.SubItems(2))
spprice = rsinstock!sprice * CInt(Lvwitems.SelectedItem.SubItems(2))
AI = rsinstock!AI
'// where it differentiate AI and Agrovet
If AI = 0 Then
'// check the stock if it is less than zero
If rsinstock.Fields(2) <= 0 Then
MsgBox "Sorry Stock is Zero for item " & Lvwitems.SelectedItem.SubItems(1) & " please re-stock before your proceed", vbInformation
Exit Sub
End If
'// check the quanttity being sold versus the balance
'Dim Remain As Double
Remain = rsinstock.Fields(2) - CInt(Lvwitems.SelectedItem.SubItems(2))
If Remain < 0 Then
MsgBox "Stock will be negative " & Remain & " please re-stock before you proceed", vbInformation
Exit Sub
End If

sql = ""
sql = sql & "SET dateformat DMY INSERT INTO ag_Receipts(R_No, P_code, T_Date, Amount, Qua, S_Bal, user_id, Cash, SNo,Transby,branch,remarks,SPRICE,BPRICE,AI) VALUES ("
sql = sql & txtrno & ",'" & Lvwitems.SelectedItem & "','" & txtransdate & "'," & Lvwitems.SelectedItem.SubItems(4)
sql = sql & "," & Lvwitems.SelectedItem.SubItems(2) & "," & Remain & ",'" & User & "'," & Lvwitems.SelectedItem.SubItems(5) & ",'" & C & "','" & D & "','" & Cmbstation & "','" & Lvwitems.SelectedItem.SubItems(1) & "','" & spprice & "','" & Bpprice & "'," & txtserai & ")"

oSaccoMaster.ExecuteThis (sql)
sql = ""
sql = sql & "SET dateformat DMY INSERT INTO ag_Receipts3(R_No, P_code, T_Date, Amount, Qua, S_Bal, user_id, Cash, SNo,Transby,branch,AI) VALUES ("
sql = sql & txtrno & ",'" & Lvwitems.SelectedItem & "','" & txtransdate & "'," & Lvwitems.SelectedItem.SubItems(4)
sql = sql & "," & Lvwitems.SelectedItem.SubItems(2) & "," & Remain & ",'" & User & "'," & Lvwitems.SelectedItem.SubItems(5) & ",'" & C & "','" & D & "','" & Cmbstation & "'," & txtserai & ")"

oSaccoMaster.ExecuteThis (sql)
oSaccoMaster.ExecuteThis ("Update ag_Products SET Qout =" & CCur(Remain) & " WHERE p_code= '" & Lvwitems.SelectedItem & "' and branch='" & Cmbstation & "'")
If txtamount <> 0 Then

sql = ""
sql = sql & "SET dateformat DMY INSERT INTO ag_Receipts(R_No, P_code, T_Date, Amount, Qua, S_Bal, user_id, Cash, SNo,Transby,branch,AI) VALUES ("
sql = sql & txtrno & ",'" & Lvwitems.SelectedItem & "','" & txtransdate & "'," & Lvwitems.SelectedItem.SubItems(4)
sql = sql & "," & Lvwitems.SelectedItem.SubItems(2) & "," & Remain & ",'" & User & "'," & Lvwitems.SelectedItem.SubItems(5) & ",'" & C & "','" & D & "','" & Cmbstation & "'," & txtserai & ")"
End If

'    sql = ""
'    sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & txtransdate & "'," & Lvwitems.SelectedItem.SubItems(4) & ",'33-104','33-501','" & Lvwitems.SelectedItem & "','" & cboproductname & "' ,'" & C & "','" & User & "',0,0)"
'    oSaccoMaster.ExecuteThis (sql)
    
    '//XXXXXXXXXXXXXXX
    sql = ""
    sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & txtransdate & "'," & Lvwitems.SelectedItem.SubItems(4) & ",'" & lbldracc & "','" & lblcracc & "','" & Lvwitems.SelectedItem & "','" & cboproductname & "' ,'CHECK OFF SALES STAFF- " & "" & cboproductname & "','" & User & "',0,0)"
    oSaccoMaster.ExecuteThis (sql)

''sql = ""
'''amount = Bpprice
''    sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & txtransdate & "'," & Bpprice & ",'33-104','33-501','" & Lvwitems.SelectedItem & "','" & cboproductname & "' ,'CHECK OFF SALES STAFF- " & "" & cboproductname & "','" & User & "',0,0)"
''    oSaccoMaster.ExecuteThis (sql)


'XXXXXXXXXXXXXXXXXXXXXX

Else

'// check the stock if it is less than zero
If rsinstock.Fields(2) <= 0 Then
MsgBox "Sorry Stock is Zero for item " & Lvwitems.SelectedItem.SubItems(1) & " please re-stock before your proceed", vbInformation
Exit Sub
End If
'// check the quanttity being sold versus the balance
'Dim Remain As Double
Remain = rsinstock.Fields(2) - CInt(Lvwitems.SelectedItem.SubItems(2))
If Remain < 0 Then
MsgBox "Stock will be negative " & Remain & " please re-stock before you proceed", vbInformation
Exit Sub
End If

sql = ""
sql = sql & "SET dateformat DMY INSERT INTO ag_Receipts(R_No, P_code, T_Date, Amount, Qua, S_Bal, user_id, Cash, SNo,Transby,branch,remarks,SPRICE,BPRICE,AI) VALUES ("
sql = sql & txtrno & ",'" & Lvwitems.SelectedItem & "','" & txtransdate & "'," & Lvwitems.SelectedItem.SubItems(4)
sql = sql & "," & Lvwitems.SelectedItem.SubItems(2) & "," & Remain & ",'" & User & "'," & Lvwitems.SelectedItem.SubItems(5) & ",'" & C & "','" & D & "','" & Cmbstation & "','" & Lvwitems.SelectedItem.SubItems(1) & "','" & spprice & "','" & Bpprice & "'," & txtserai & ")"

oSaccoMaster.ExecuteThis (sql)
sql = ""
sql = sql & "SET dateformat DMY INSERT INTO ag_Receipts3(R_No, P_code, T_Date, Amount, Qua, S_Bal, user_id, Cash, SNo,Transby,branch,AI) VALUES ("
sql = sql & txtrno & ",'" & Lvwitems.SelectedItem & "','" & txtransdate & "'," & Lvwitems.SelectedItem.SubItems(4)
sql = sql & "," & Lvwitems.SelectedItem.SubItems(2) & "," & Remain & ",'" & User & "'," & Lvwitems.SelectedItem.SubItems(5) & ",'" & C & "','" & D & "','" & Cmbstation & "'," & txtserai & ")"

oSaccoMaster.ExecuteThis (sql)
oSaccoMaster.ExecuteThis ("Update ag_Products SET Qout =" & CCur(Remain) & " WHERE p_code= '" & Lvwitems.SelectedItem & "' and branch='" & Cmbstation & "'")
If txtamount <> 0 Then

sql = ""
sql = sql & "SET dateformat DMY INSERT INTO ag_Receipts(R_No, P_code, T_Date, Amount, Qua, S_Bal, user_id, Cash, SNo,Transby,branch,AI) VALUES ("
sql = sql & txtrno & ",'" & Lvwitems.SelectedItem & "','" & txtransdate & "'," & Lvwitems.SelectedItem.SubItems(4)
sql = sql & "," & Lvwitems.SelectedItem.SubItems(2) & "," & Remain & ",'" & User & "'," & Lvwitems.SelectedItem.SubItems(5) & ",'" & C & "','" & D & "','" & Cmbstation & "'," & txtserai & ")"
End If

'    sql = ""
'    sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & txtransdate & "'," & Lvwitems.SelectedItem.SubItems(4) & ",'33-104','33-501','" & Lvwitems.SelectedItem & "','" & cboproductname & "' ,'" & C & "','" & User & "',0,0)"
'    oSaccoMaster.ExecuteThis (sql)
    
    '//XXXXXXXXXXXXXXX
    sql = ""
    sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & txtransdate & "'," & Lvwitems.SelectedItem.SubItems(4) & ",'" & lbldracc & "','" & lblcracc & "','" & Lvwitems.SelectedItem & "','" & cboproductname & "' ,'CHECK OFF SALES STAFF- " & "" & cboproductname & "','" & User & "',0,0)"
    oSaccoMaster.ExecuteThis (sql)

'sql = ""
''amount = Bpprice
'    sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & txtransdate & "'," & Bpprice & ",'22-04','22-03','" & Lvwitems.SelectedItem & "','" & cboproductname & "' ,'CHECK OFF SALES STAFF- " & "" & cboproductname & "','" & User & "',0,0)"
'    oSaccoMaster.ExecuteThis (sql)


'XXXXXXXXXXXXXXXXXXXXXX
End If
    
    
    
Next j

If chkPrint.value = vbChecked Then
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
       Dim fso, chkPrinter, txtFile
        'ttt = "LPT1" 'LPT1
         Dim PORT As String
        PORT = ports
        'ttt = "LPT1" 'LPT1
        ttt = PORT
        Set fso = CreateObject("Scripting.FileSystemObject")
        Dim strReceipts As String
        j = 1
        strReceipts = ""
        Do While Not j > (Lvwitems.ListItems.Count)
            Lvwitems.ListItems.Item(j).selected = True
            strReceipts = strReceipts & Lvwitems.SelectedItem.SubItems(1) & vbNewLine & Lvwitems.SelectedItem.SubItems(2) & vbTab & vbTab
            strReceipts = strReceipts & Format(Lvwitems.SelectedItem.SubItems(3), "#,##0.00") & vbTab & vbTab & Format(Lvwitems.SelectedItem.SubItems(4), "#,##0.00") & vbNewLine
            j = j + 1
        Loop

        'MsgBox strReceipts
        strReceipts = strReceipts & vbNewLine & "---------------------------------------" & vbNewLine
        strReceipts = strReceipts & "TOTAL" & vbTab & vbTab & vbTab & vbTab & Format(total, "#,##0.00") & vbNewLine
        strReceipts = strReceipts & "======================================="
        Set txtFile = fso.CreateTextFile(ttt, True)
        
        If optCash = True Then
        Set rs = New ADODB.Recordset
        Dim a As String
        sql = "select Adress from d_company"
        Set rs = oSaccoMaster.GetRecordset(sql)
        If Not rs.EOF Then
        If Not IsNull(rs.Fields(0)) Then a = rs.Fields(0)
        End If
    txtFile.WriteLine "  " & cname & ""
    txtFile.WriteLine "     " & a & ""
    txtFile.WriteLine "---------------------------------------"
    txtFile.WriteLine "  AGROVET RECEIPT"
    txtFile.WriteLine "     STAFF SALES"
    txtFile.WriteLine "---------------------------------------"
'nAME QNTY PRICE AMOUNT
    txtFile.WriteLine "QNTY" & vbTab & vbTab & "PRICE" & vbTab & vbTab & "AMOUNT"
    txtFile.WriteLine "........................................"
    txtFile.WriteLine strReceipts
    txtFile.WriteLine
    txtFile.WriteLine "TOTAL" & TXTTOTAL
    txtFile.WriteLine
    txtFile.WriteLine "AMOUNT RECEVED" & txtamtreceived
    txtFile.WriteLine
    txtFile.WriteLine "CHANGE" & TXTCHANGE
    txtFile.WriteLine
    txtFile.WriteLine "YOU WERE SERVED By " & UCase(username)
    txtFile.WriteLine "    Date :" & Format(Get_Server_Date, "dd/mm/yyyy HH:MM:SS AM/PM")
    txtFile.WriteLine "     THANK YOU AND WELCOME "
    txtFile.WriteLine "****************************************"
    txtFile.WriteLine escFeedAndCut
    txtFile.Close
    End If
End If
'//Update deductions
'If optCash.value = False Then
'Set cn = New ADODB.Connection
'sql = "d_sp_TransDeduct '" & txttranscode & "','" & txtransdate & "','Agrovet'," & Total & ",'" & Startdate & "','" & Enddate & "','" & User & "'"
'oSaccoMaster.ExecuteThis (sql)

'Set cn = New ADODB.Connection
'sql = "d_sp_TransDeduct '" & txtTCode & "','" & DTPDDate & "','" & cboDeductionType & "'," & txtamount & ",'" & DTPStartDate & "','" & DTPEndDate & "','" & User & "'"
'oSaccoMaster.ExecuteThis (sql)





Lvwitems.ListItems.Clear
txtrno = ""
txtpcode.Text = ""
txtquantity = 1
txtamount = ""

MsgBox "Record saved Successfully"
Exit Sub
End If
olkalou:
MsgBox err.description & " error occured."

End Sub
Private Sub saveothers()
On Error GoTo olkalou

If Optothers = True Then
Dim C As String
Dim D As String
If txtname = "" Then
MsgBox "Enter Institution Name before you continue", vbCritical, "Maziwa"

Exit Sub
End If
C = "Other"
D = txtname
Dim j As Integer
If Lvwitems.ListItems.Count = 0 Then
MsgBox "There are no items sold."
Exit Sub
End If
'If txtstaffno = "" Then
'MsgBox "Enter Staff Number before you continue", vbCritical, "Maziwa"
'
'Exit Sub
'End If
j = 1

Dim total As Currency
total = 0
Do While Not j > (Lvwitems.ListItems.Count)
 Lvwitems.ListItems.Item(j).selected = True
 total = total + CCur(Lvwitems.SelectedItem.SubItems(4))
j = j + 1
Loop



Startdate = DateSerial(Year(txtransdate), Month(txtransdate), 1)
Enddate = DateSerial(Year(txtransdate), Month(txtransdate) + 1, 1 - 1)

Set rs = oSaccoMaster.GetRecordset("d_sp_IsClosed '" & Enddate & "'")
If Not rs.EOF Then
    MsgBox "The period " & Enddate & " has been closed by " & rs.Fields(0)
    Exit Sub
End If


'// check if they are in stock.
For j = 1 To Lvwitems.ListItems.Count
 Lvwitems.ListItems.Item(j).selected = True


Dim rsinstock As Recordset
sql = ""
sql = "select P_CODE,qin,qout,pprice, sprice,AI from ag_products where p_code='" & Lvwitems.SelectedItem & "'"
'//Set rsinstock = New ADODB.Recordset
Set rsinstock = oSaccoMaster.GetRecordset(sql)
Dim Remain As Double, Bpprice As Double, spprice As Double, AI As Integer
Remain = rsinstock.Fields(2) - CInt(Lvwitems.SelectedItem.SubItems(2))
Bpprice = rsinstock!pprice * CInt(Lvwitems.SelectedItem.SubItems(2))
spprice = rsinstock!sprice * CInt(Lvwitems.SelectedItem.SubItems(2))
AI = rsinstock!AI
'//where it differentiate AI and Agrovet
If AI = 0 Then
'// check the stock if it is less than zero
If rsinstock.Fields(2) <= 0 Then
MsgBox "Sorry Stock is Zero for item " & Lvwitems.SelectedItem.SubItems(1) & " please re-stock before your proceed", vbInformation
Exit Sub
End If
'// check the quanttity being sold versus the balance
'Dim Remain As Double
Remain = rsinstock.Fields(2) - CInt(Lvwitems.SelectedItem.SubItems(2))
If Remain < 0 Then
MsgBox "Stock will be negative " & Remain & " please re-stock before you proceed", vbInformation
Exit Sub
End If

sql = ""
sql = sql & "SET dateformat DMY INSERT INTO ag_Receipts(R_No, P_code, T_Date, Amount, Qua, S_Bal, user_id, Cash, SNo,Transby,branch,remarks,SPRICE,BPRICE,AI) VALUES ("
sql = sql & txtrno & ",'" & Lvwitems.SelectedItem & "','" & txtransdate & "'," & Lvwitems.SelectedItem.SubItems(4)
sql = sql & "," & Lvwitems.SelectedItem.SubItems(2) & "," & Remain & ",'" & User & "'," & Lvwitems.SelectedItem.SubItems(5) & ",'" & C & "','" & D & "','" & Cmbstation & "','" & Lvwitems.SelectedItem.SubItems(1) & "','" & spprice & "','" & Bpprice & "'," & txtserai & ")"

oSaccoMaster.ExecuteThis (sql)
sql = ""
sql = sql & "SET dateformat DMY INSERT INTO ag_Receipts3(R_No, P_code, T_Date, Amount, Qua, S_Bal, user_id, Cash, SNo,Transby,branch,AI) VALUES ("
sql = sql & txtrno & ",'" & Lvwitems.SelectedItem & "','" & txtransdate & "'," & Lvwitems.SelectedItem.SubItems(4)
sql = sql & "," & Lvwitems.SelectedItem.SubItems(2) & "," & Remain & ",'" & User & "'," & Lvwitems.SelectedItem.SubItems(5) & ",'" & C & "','" & D & "','" & Cmbstation & "'," & txtserai & ")"

oSaccoMaster.ExecuteThis (sql)
oSaccoMaster.ExecuteThis ("Update ag_Products SET Qout =" & CCur(Remain) & " WHERE p_code= '" & Lvwitems.SelectedItem & "' and branch='" & Cmbstation & "'")


'    sql = ""
'    sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & txtransdate & "'," & Lvwitems.SelectedItem & ",'A006','I004','" & Lvwitems.SelectedItem & "','" & cboproductname & "' ,'" & C & "','" & User & "',0,0)"
'    oSaccoMaster.ExecuteThis (sql)
'//XXXXXXXXXXXXXXX
    sql = ""
    sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & txtransdate & "'," & Lvwitems.SelectedItem.SubItems(4) & ",'" & lbldracc & "','" & lblcracc & "','" & Lvwitems.SelectedItem & "','" & cboproductname & "' ,'OTHERS CHECK OFF SALES- " & "" & cboproductname & "','" & User & "',0,0)"
    oSaccoMaster.ExecuteThis (sql)

'sql = ""
''amount = Bpprice
'    sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & txtransdate & "'," & Bpprice & ",'33-104','33-501','" & Lvwitems.SelectedItem & "','" & cboproductname & "' ,'OTHERS CHECK OFF SALES- " & "" & cboproductname & "','" & User & "',0,0)"
'    oSaccoMaster.ExecuteThis (sql)
    
Else

'// check the stock if it is less than zero
If rsinstock.Fields(2) <= 0 Then
MsgBox "Sorry Stock is Zero for item " & Lvwitems.SelectedItem.SubItems(1) & " please re-stock before your proceed", vbInformation
Exit Sub
End If
'// check the quanttity being sold versus the balance
'Dim Remain As Double
Remain = rsinstock.Fields(2) - CInt(Lvwitems.SelectedItem.SubItems(2))
If Remain < 0 Then
MsgBox "Stock will be negative " & Remain & " please re-stock before you proceed", vbInformation
Exit Sub
End If

sql = ""
sql = sql & "SET dateformat DMY INSERT INTO ag_Receipts(R_No, P_code, T_Date, Amount, Qua, S_Bal, user_id, Cash, SNo,Transby,branch,remarks,SPRICE,BPRICE,AI) VALUES ("
sql = sql & txtrno & ",'" & Lvwitems.SelectedItem & "','" & txtransdate & "'," & Lvwitems.SelectedItem.SubItems(4)
sql = sql & "," & Lvwitems.SelectedItem.SubItems(2) & "," & Remain & ",'" & User & "'," & Lvwitems.SelectedItem.SubItems(5) & ",'" & C & "','" & D & "','" & Cmbstation & "','" & Lvwitems.SelectedItem.SubItems(1) & "','" & spprice & "','" & Bpprice & "'," & txtserai & ")"

oSaccoMaster.ExecuteThis (sql)
sql = ""
sql = sql & "SET dateformat DMY INSERT INTO ag_Receipts3(R_No, P_code, T_Date, Amount, Qua, S_Bal, user_id, Cash, SNo,Transby,branch,AI) VALUES ("
sql = sql & txtrno & ",'" & Lvwitems.SelectedItem & "','" & txtransdate & "'," & Lvwitems.SelectedItem.SubItems(4)
sql = sql & "," & Lvwitems.SelectedItem.SubItems(2) & "," & Remain & ",'" & User & "'," & Lvwitems.SelectedItem.SubItems(5) & ",'" & C & "','" & D & "','" & Cmbstation & "'," & txtserai & ")"

oSaccoMaster.ExecuteThis (sql)
oSaccoMaster.ExecuteThis ("Update ag_Products SET Qout =" & CCur(Remain) & " WHERE p_code= '" & Lvwitems.SelectedItem & "' and branch='" & Cmbstation & "'")


'    sql = ""
'    sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & txtransdate & "'," & Lvwitems.SelectedItem & ",'A006','I004','" & Lvwitems.SelectedItem & "','" & cboproductname & "' ,'" & C & "','" & User & "',0,0)"
'    oSaccoMaster.ExecuteThis (sql)
'//XXXXXXXXXXXXXXX
    sql = ""
    sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & txtransdate & "'," & Lvwitems.SelectedItem.SubItems(4) & ",'" & lbldracc & "','" & lblcracc & "','" & Lvwitems.SelectedItem & "','" & cboproductname & "' ,'OTHERS CHECK OFF SALES- " & "" & cboproductname & "','" & User & "',0,0)"
    oSaccoMaster.ExecuteThis (sql)

'sql = ""
''amount = Bpprice
'    sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & txtransdate & "'," & Bpprice & ",'22-04','22-03','" & Lvwitems.SelectedItem & "','" & cboproductname & "' ,'OTHERS CHECK OFF SALES- " & "" & cboproductname & "','" & User & "',0,0)"
'    oSaccoMaster.ExecuteThis (sql)
End If
Next j

If chkPrint.value = vbChecked Then
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
       Dim fso, chkPrinter, txtFile
        'ttt = "LPT1" 'LPT1
         Dim PORT As String
        PORT = ports
        'ttt = "LPT1" 'LPT1
        ttt = PORT
        Set fso = CreateObject("Scripting.FileSystemObject")
        Dim strReceipts As String
        j = 1
        strReceipts = ""
        Do While Not j > (Lvwitems.ListItems.Count)
            Lvwitems.ListItems.Item(j).selected = True
            strReceipts = strReceipts & Lvwitems.SelectedItem.SubItems(1) & vbNewLine & Lvwitems.SelectedItem.SubItems(2) & vbTab & vbTab
            strReceipts = strReceipts & Format(Lvwitems.SelectedItem.SubItems(3), "#,##0.00") & vbTab & vbTab & Format(Lvwitems.SelectedItem.SubItems(4), "#,##0.00") & vbNewLine
            j = j + 1
        Loop

        'MsgBox strReceipts
        strReceipts = strReceipts & vbNewLine & "---------------------------------------" & vbNewLine
        strReceipts = strReceipts & "TOTAL" & vbTab & vbTab & vbTab & vbTab & Format(total, "#,##0.00") & vbNewLine
        strReceipts = strReceipts & "======================================="
        Set txtFile = fso.CreateTextFile(ttt, True)
        
        If optCash = True Then
        Set rs = New ADODB.Recordset
        Dim a As String
        sql = "select Adress from d_company"
        Set rs = oSaccoMaster.GetRecordset(sql)
        If Not rs.EOF Then
        If Not IsNull(rs.Fields(0)) Then a = rs.Fields(0)
        End If
    txtFile.WriteLine "  " & cname & ""
    txtFile.WriteLine "     " & a & ""
    txtFile.WriteLine "---------------------------------------"
    txtFile.WriteLine "  AGROVET RECEIPT"
    txtFile.WriteLine "     NON SUPPLIERS SALES"
    txtFile.WriteLine "---------------------------------------"
'nAME QNTY PRICE AMOUNT
    txtFile.WriteLine "QNTY" & vbTab & vbTab & "PRICE" & vbTab & vbTab & "AMOUNT"
    txtFile.WriteLine "........................................"
    txtFile.WriteLine strReceipts
    txtFile.WriteLine
    txtFile.WriteLine "TOTAL" & TXTTOTAL
    txtFile.WriteLine
    txtFile.WriteLine "AMOUNT RECEVED" & txtamtreceived
    txtFile.WriteLine
    txtFile.WriteLine "CHANGE" & TXTCHANGE
    txtFile.WriteLine
    txtFile.WriteLine "YOU WERE SERVED By " & UCase(username)
    txtFile.WriteLine "    Date :" & Format(Get_Server_Date, "dd/mm/yyyy HH:MM:SS AM/PM")
    txtFile.WriteLine "     THANK YOU AND WELCOME "
    txtFile.WriteLine "****************************************"
    txtFile.WriteLine escFeedAndCut
    txtFile.Close
    End If
End If
'//Update deductions
'If optCash.value = False Then
'Set cn = New ADODB.Connection
'sql = "d_sp_TransDeduct '" & txttranscode & "','" & txtransdate & "','Agrovet'," & Total & ",'" & Startdate & "','" & Enddate & "','" & User & "'"
'oSaccoMaster.ExecuteThis (sql)

'Set cn = New ADODB.Connection
'sql = "d_sp_TransDeduct '" & txtTCode & "','" & DTPDDate & "','" & cboDeductionType & "'," & txtamount & ",'" & DTPStartDate & "','" & DTPEndDate & "','" & User & "'"
'oSaccoMaster.ExecuteThis (sql)





Lvwitems.ListItems.Clear
txtrno = ""
txtpcode.Text = ""
txtquantity = 1
txtamount = ""

MsgBox "Record saved Successfully"
Exit Sub
End If
olkalou:
MsgBox err.description & " error occured."

End Sub
Private Sub savecash()
On Error GoTo olkalou

If optCash = True Then
Dim C As String
C = "Cash"

Dim j As Integer
If Lvwitems.ListItems.Count = 0 Then
MsgBox "There are no items sold."
Exit Sub
End If
j = 1

Dim total As Currency
total = 0
Do While Not j > (Lvwitems.ListItems.Count)
 Lvwitems.ListItems.Item(j).selected = True
 total = total + CCur(Lvwitems.SelectedItem.SubItems(4))
j = j + 1
Loop



Startdate = DateSerial(Year(txtransdate), Month(txtransdate), 1)
Enddate = DateSerial(Year(txtransdate), Month(txtransdate) + 1, 1 - 1)

Set rs = oSaccoMaster.GetRecordset("d_sp_IsClosed '" & Enddate & "'")
If Not rs.EOF Then
    MsgBox "The period " & Enddate & " has been closed by " & rs.Fields(0)
    Exit Sub
End If


'// check if they are in stock.
For j = 1 To Lvwitems.ListItems.Count
 Lvwitems.ListItems.Item(j).selected = True


Dim rsinstock As Recordset
sql = ""
'sql = "select P_CODE,qin,qout,pprice, sprice  from ag_products where p_code='" & Lvwitems.SelectedItem & "'"
sql = "select P_CODE,qin,qout,pprice, sprice,AI from ag_products where p_code='" & Lvwitems.SelectedItem & "' AND  branch='" & Cmbstation & "'"

Set rsinstock = oSaccoMaster.GetRecordset(sql)

Dim Remain As Double, Bpprice As Double, spprice As Double, AI As Integer
Remain = rsinstock.Fields(2) - CInt(Lvwitems.SelectedItem.SubItems(2))
Bpprice = rsinstock!pprice * CInt(Lvwitems.SelectedItem.SubItems(2))
spprice = rsinstock!sprice * CInt(Lvwitems.SelectedItem.SubItems(2))
AI = rsinstock!AI
'//where it differentiate AI and Agrovet
If AI = 0 Then
'// check the stock if it is less than zero
If rsinstock.Fields(2) <= 0 Then
MsgBox "Sorry Stock is Zero for item " & Lvwitems.SelectedItem.SubItems(1) & " please re-stock before your proceed", vbInformation
Exit Sub
End If
'// check the quanttity being sold versus the balance

Remain = rsinstock.Fields(2) - CInt(Lvwitems.SelectedItem.SubItems(2))
If Remain < 0 Then
MsgBox "Stock will be negative " & Remain & " please re-stock before you proceed", vbInformation
Exit Sub
End If

sql = ""
sql = sql & "SET dateformat DMY INSERT INTO ag_Receipts(R_No, P_code, T_Date, Amount, Qua, S_Bal, user_id, Cash, SNo,Transby, Idno, mobile,branch,remarks,SPRICE,BPRICE,AI) VALUES ("
sql = sql & txtrno & ",'" & Lvwitems.SelectedItem & "','" & txtransdate & "'," & Lvwitems.SelectedItem.SubItems(4)
sql = sql & "," & Lvwitems.SelectedItem.SubItems(2) & "," & Remain & ",'" & User & "'," & Lvwitems.SelectedItem.SubItems(5) & ",'" & C & "','" & txttransby & "','" & txtidno & "','" & txtmobile & "','" & Cmbstation & "','" & Lvwitems.SelectedItem.SubItems(1) & "','" & spprice & "','" & Bpprice & "'," & txtserai & ")"

oSaccoMaster.ExecuteThis (sql)
sql = ""
sql = sql & "SET dateformat DMY INSERT INTO ag_Receipts3(R_No, P_code, T_Date, Amount, Qua, S_Bal, user_id, Cash, SNo,Transby, Idno, mobile,branch,AI) VALUES ("
sql = sql & txtrno & ",'" & Lvwitems.SelectedItem & "','" & txtransdate & "'," & Lvwitems.SelectedItem.SubItems(4)
sql = sql & "," & Lvwitems.SelectedItem.SubItems(2) & "," & Remain & ",'" & User & "'," & Lvwitems.SelectedItem.SubItems(5) & ",'" & C & "','" & txttransby & "','" & txtidno & "','" & txtmobile & "','" & Cmbstation & "'," & txtserai & ")"

oSaccoMaster.ExecuteThis (sql)
oSaccoMaster.ExecuteThis ("Update ag_Products SET Qout =" & CCur(Remain) & " WHERE p_code= '" & Lvwitems.SelectedItem & "' and branch='" & Cmbstation & "'")
oSaccoMaster.ExecuteThis ("Update Rcpno SET rcpno =" & txtrno & "")

'\\ save to gl


    sql = ""
    sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & txtransdate & "'," & Lvwitems.SelectedItem.SubItems(4) & ",'" & lbldracc & "','" & lblcracc & "','" & Lvwitems.SelectedItem & "','" & cboproductname & "' ,'Cash Sales- " & "" & cboproductname & "','" & User & "','1','0')"
    oSaccoMaster.ExecuteThis (sql)
''    sql = ""
''    sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & txtransdate & "'," & Bpprice & ",'33-104','33-501','" & Lvwitems.SelectedItem & "','" & cboproductname & "' ,'Cash Sales- " & "" & cboproductname & "','" & User & "','1','0')"
''    oSaccoMaster.ExecuteThis (sql)

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

Else

'// check the stock if it is less than zero
If rsinstock.Fields(2) <= 0 Then
MsgBox "Sorry Stock is Zero for item " & Lvwitems.SelectedItem.SubItems(1) & " please re-stock before your proceed", vbInformation
Exit Sub
End If
'// check the quanttity being sold versus the balance

Remain = rsinstock.Fields(2) - CInt(Lvwitems.SelectedItem.SubItems(2))
If Remain < 0 Then
MsgBox "Stock will be negative " & Remain & " please re-stock before you proceed", vbInformation
Exit Sub
End If

sql = ""
sql = sql & "SET dateformat DMY INSERT INTO ag_Receipts(R_No, P_code, T_Date, Amount, Qua, S_Bal, user_id, Cash, SNo,Transby, Idno, mobile,branch,remarks,SPRICE,BPRICE,AI) VALUES ("
sql = sql & txtrno & ",'" & Lvwitems.SelectedItem & "','" & txtransdate & "'," & Lvwitems.SelectedItem.SubItems(4)
sql = sql & "," & Lvwitems.SelectedItem.SubItems(2) & "," & Remain & ",'" & User & "'," & Lvwitems.SelectedItem.SubItems(5) & ",'" & C & "','" & txttransby & "','" & txtidno & "','" & txtmobile & "','" & Cmbstation & "','" & Lvwitems.SelectedItem.SubItems(1) & "','" & spprice & "','" & Bpprice & "'," & txtserai & ")"

oSaccoMaster.ExecuteThis (sql)
sql = ""
sql = sql & "SET dateformat DMY INSERT INTO ag_Receipts3(R_No, P_code, T_Date, Amount, Qua, S_Bal, user_id, Cash, SNo,Transby, Idno, mobile,branch,AI) VALUES ("
sql = sql & txtrno & ",'" & Lvwitems.SelectedItem & "','" & txtransdate & "'," & Lvwitems.SelectedItem.SubItems(4)
sql = sql & "," & Lvwitems.SelectedItem.SubItems(2) & "," & Remain & ",'" & User & "'," & Lvwitems.SelectedItem.SubItems(5) & ",'" & C & "','" & txttransby & "','" & txtidno & "','" & txtmobile & "','" & Cmbstation & "'," & txtserai & ")"

oSaccoMaster.ExecuteThis (sql)
oSaccoMaster.ExecuteThis ("Update ag_Products SET Qout =" & CCur(Remain) & " WHERE p_code= '" & Lvwitems.SelectedItem & "' and branch='" & Cmbstation & "'")
oSaccoMaster.ExecuteThis ("Update Rcpno SET rcpno =" & txtrno & "")

'\\ save to gl


    sql = ""
    sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & txtransdate & "'," & Lvwitems.SelectedItem.SubItems(4) & ",'" & lbldracc & "','" & lblcracc & "','" & Lvwitems.SelectedItem & "','" & cboproductname & "' ,'Cash Sales- " & "" & cboproductname & "','" & User & "','1','0')"
    oSaccoMaster.ExecuteThis (sql)
'    sql = ""
'    sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & txtransdate & "'," & Bpprice & ",'22-04','22-03','" & Lvwitems.SelectedItem & "','" & cboproductname & "' ,'Cash Sales- " & "" & cboproductname & "','" & User & "','1','0')"
'    oSaccoMaster.ExecuteThis (sql)

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
End If
Next j


If chkPrint.value = vbChecked Then
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
        Dim fso, chkPrinter, txtFile
        'ttt = "LPT1" 'LPT1
         Dim PORT As String
        PORT = ports
        'ttt = "LPT1" 'LPT1
        ttt = PORT
        Set fso = CreateObject("Scripting.FileSystemObject")
        Dim strReceipts As String
        j = 1
        strReceipts = ""
        Do While Not j > (Lvwitems.ListItems.Count)
            Lvwitems.ListItems.Item(j).selected = True
            strReceipts = strReceipts & Lvwitems.SelectedItem.SubItems(1) & vbNewLine & Lvwitems.SelectedItem.SubItems(2) & vbTab & vbTab
            strReceipts = strReceipts & Format(Lvwitems.SelectedItem.SubItems(3), "#,##0.00") & vbTab & vbTab & Format(Lvwitems.SelectedItem.SubItems(4), "#,##0.00") & vbNewLine
            j = j + 1
        Loop

        'MsgBox strReceipts
        strReceipts = strReceipts & vbNewLine & "---------------------------------------" & vbNewLine
        strReceipts = strReceipts & "TOTAL" & vbTab & vbTab & vbTab & vbTab & Format(total, "#,##0.00") & vbNewLine
        strReceipts = strReceipts & "======================================="
        Set txtFile = fso.CreateTextFile(ttt, True)
        
        txtFile.WriteLine "      " & cname & ""
        txtFile.WriteLine "      AGROVET"
        txtFile.WriteLine "      " & paddress & ""
        txtFile.WriteLine "      " & town & ""
        txtFile.WriteLine "      " & Phone & ""
        'txtfile.WriteLine "      " & Email & ""
        
        txtFile.WriteLine "---------------------------------------"
        txtFile.WriteLine "  AGROVET RECEIPT"
        txtFile.WriteLine "     CASH SALES"
        txtFile.WriteLine "---------------------------------------"
        txtFile.WriteLine "QNTY" & vbTab & vbTab & "PRICE" & vbTab & vbTab & "AMOUNT"
        txtFile.WriteLine "........................................"
        txtFile.WriteLine "---------------------------------------"

        txtFile.WriteLine strReceipts
        txtFile.WriteLine
        
        txtFile.WriteLine "TOTAL" & vbTab & vbTab & vbTab & vbTab & Format(total, "#,##0.00") & vbNewLine
        txtFile.WriteLine
        txtFile.WriteLine "TOTAL" & vbTab & TXTTOTAL
        txtFile.WriteLine
        txtFile.WriteLine "AMOUNT RECEVED" & vbTab & txtamtreceived
        txtFile.WriteLine
        txtFile.WriteLine "CHANGE" & vbTab & TXTCHANGE
        txtFile.WriteLine
        txtFile.WriteLine "Trans By" & vbTab & txttransby
        txtFile.WriteLine "Id No" & vbTab & txtidno
        txtFile.WriteLine
        txtFile.WriteLine "YOU WERE SERVED By " & UCase(username)
        txtFile.WriteLine "    Date :" & Format(Get_Server_Date, "dd/mm/yyyy HH:MM:SS AM/PM")
        txtFile.WriteLine "     THANK YOU AND WELCOME "
        
        txtFile.WriteLine " GOODS ONCE SOLD WILL NOT BE RE-ACCEPTED"
        txtFile.WriteLine "****************************************"
        txtFile.WriteLine escFeedAndCut
        txtFile.Close
    End If
End If


Lvwitems.ListItems.Clear
txtrno = ""
txtpcode.Text = ""
txtquantity = 1
txtamount = ""
txttransby = ""
txtidno = ""
txtmobile = ""
MsgBox "Record saved Successfully"
Exit Sub
olkalou:
MsgBox err.description & " error occured."

End Sub
Private Sub savempesa()

'End Sub
On Error GoTo olkalou

'If Optmpesa = True Then
'Dim C As String
'C = "MPESA"

'If Cmbstation <> "OLENGURUONE" Then
'
'Exit Sub
'End If

'If optCash = True Then
'If Cmbstation <> "OLENGURUONE" Then
'MsgBox "Please yours branch is only allowed to sell using Mpesa", vbInformation
'Exit Sub
'End If
'End If

If Optmpesa = True Then
Dim C As String
C = "MPESA"

Dim j As Integer
If Lvwitems.ListItems.Count = 0 Then
MsgBox "There are no items sold."
Exit Sub
End If
j = 1

Dim total As Currency
total = 0
Do While Not j > (Lvwitems.ListItems.Count)
 Lvwitems.ListItems.Item(j).selected = True
 total = total + CCur(Lvwitems.SelectedItem.SubItems(4))
j = j + 1
Loop



Startdate = DateSerial(Year(txtransdate), Month(txtransdate), 1)
Enddate = DateSerial(Year(txtransdate), Month(txtransdate) + 1, 1 - 1)

Set rs = oSaccoMaster.GetRecordset("d_sp_IsClosed '" & Enddate & "'")
If Not rs.EOF Then
    MsgBox "The period " & Enddate & " has been closed by " & rs.Fields(0)
    Exit Sub
End If


'// check if they are in stock.
For j = 1 To Lvwitems.ListItems.Count
 Lvwitems.ListItems.Item(j).selected = True


Dim rsinstock As Recordset
sql = ""
'sql = "select P_CODE,qin,qout from ag_products where p_code='" & Lvwitems.SelectedItem & "'"
sql = "select P_CODE,qin,qout,pprice, sprice,AI  from ag_products where p_code='" & Lvwitems.SelectedItem & "' AND  branch='" & Cmbstation & "'"
'//Set rsinstock = New ADODB.Recordset
Set rsinstock = oSaccoMaster.GetRecordset(sql)
Dim Remain As Double, Bpprice As Double, spprice As Double, AI As Integer
Remain = rsinstock.Fields(2) - CInt(Lvwitems.SelectedItem.SubItems(2))
Bpprice = rsinstock!pprice * CInt(Lvwitems.SelectedItem.SubItems(2))
spprice = rsinstock!sprice * CInt(Lvwitems.SelectedItem.SubItems(2))
AI = rsinstock!AI
'//where it differentiate AI and Agrovet
If AI = 0 Then
'// check the stock if it is less than zero
If rsinstock.Fields(2) <= 0 Then
MsgBox "Sorry Stock is Zero for item " & Lvwitems.SelectedItem.SubItems(1) & " please re-stock before your proceed", vbInformation
Exit Sub
End If
'// check the quanttity being sold versus the balance
'Dim Remain As Double
Remain = rsinstock.Fields(2) - CInt(Lvwitems.SelectedItem.SubItems(2))
If Remain < 0 Then
MsgBox "Stock will be negative " & Remain & " please re-stock before you proceed", vbInformation
Exit Sub
End If

sql = ""
sql = sql & "SET dateformat DMY INSERT INTO ag_Receipts(R_No, P_code, T_Date, Amount, Qua, S_Bal, user_id, Cash, SNo,Transby, Idno, mobile,branch,remarks,SPRICE,BPRICE,AI) VALUES ("
sql = sql & txtrno & ",'" & Lvwitems.SelectedItem & "','" & txtransdate & "'," & Lvwitems.SelectedItem.SubItems(4)
sql = sql & "," & Lvwitems.SelectedItem.SubItems(2) & "," & Remain & ",'" & User & "'," & Lvwitems.SelectedItem.SubItems(5) & ",'" & C & "','" & txttransby & "','" & txtidno & "','" & txtmobile & "','" & Cmbstation & "','" & Lvwitems.SelectedItem.SubItems(1) & "','" & spprice & "','" & Bpprice & "'," & txtserai & ")"

oSaccoMaster.ExecuteThis (sql)
sql = ""
sql = sql & "SET dateformat DMY INSERT INTO ag_Receipts3(R_No, P_code, T_Date, Amount, Qua, S_Bal, user_id, Cash, SNo,Transby, Idno, mobile,branch,AI) VALUES ("
sql = sql & txtrno & ",'" & Lvwitems.SelectedItem & "','" & txtransdate & "'," & Lvwitems.SelectedItem.SubItems(4)
sql = sql & "," & Lvwitems.SelectedItem.SubItems(2) & "," & Remain & ",'" & User & "'," & Lvwitems.SelectedItem.SubItems(5) & ",'" & C & "','" & txttransby & "','" & txtidno & "','" & txtmobile & "','" & Cmbstation & "'," & txtserai & ")"

oSaccoMaster.ExecuteThis (sql)
oSaccoMaster.ExecuteThis ("Update ag_Products SET Qout =" & CCur(Remain) & " WHERE p_code= '" & Lvwitems.SelectedItem & "' and branch='" & Cmbstation & "'")
oSaccoMaster.ExecuteThis ("Update Rcpno SET rcpno =" & txtrno & "")

'\\ save to gl


    sql = ""
    sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & txtransdate & "'," & Lvwitems.SelectedItem.SubItems(4) & ",'" & lbldracc & "','" & lblcracc & "','" & Lvwitems.SelectedItem & "','" & cboproductname & "' ,'MPESA Sales- " & "" & cboproductname & "','" & User & "',1,0)"
    oSaccoMaster.ExecuteThis (sql)
''    sql = ""
''    sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & txtransdate & "','" & Bpprice & "','33-104','33-501','" & Lvwitems.SelectedItem & "','" & cboproductname & "' ,'MPESA Sales- " & "" & cboproductname & "','" & User & "',1,0)"
''    oSaccoMaster.ExecuteThis (sql)
'

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

Else

'// check the stock if it is less than zero
If rsinstock.Fields(2) <= 0 Then
MsgBox "Sorry Stock is Zero for item " & Lvwitems.SelectedItem.SubItems(1) & " please re-stock before your proceed", vbInformation
Exit Sub
End If
'// check the quanttity being sold versus the balance
'Dim Remain As Double
Remain = rsinstock.Fields(2) - CInt(Lvwitems.SelectedItem.SubItems(2))
If Remain < 0 Then
MsgBox "Stock will be negative " & Remain & " please re-stock before you proceed", vbInformation
Exit Sub
End If

sql = ""
sql = sql & "SET dateformat DMY INSERT INTO ag_Receipts(R_No, P_code, T_Date, Amount, Qua, S_Bal, user_id, Cash, SNo,Transby, Idno, mobile,branch,remarks,SPRICE,BPRICE,AI) VALUES ("
sql = sql & txtrno & ",'" & Lvwitems.SelectedItem & "','" & txtransdate & "'," & Lvwitems.SelectedItem.SubItems(4)
sql = sql & "," & Lvwitems.SelectedItem.SubItems(2) & "," & Remain & ",'" & User & "'," & Lvwitems.SelectedItem.SubItems(5) & ",'" & C & "','" & txttransby & "','" & txtidno & "','" & txtmobile & "','" & Cmbstation & "','" & Lvwitems.SelectedItem.SubItems(1) & "','" & spprice & "','" & Bpprice & "'," & txtserai & ")"

oSaccoMaster.ExecuteThis (sql)
sql = ""
sql = sql & "SET dateformat DMY INSERT INTO ag_Receipts3(R_No, P_code, T_Date, Amount, Qua, S_Bal, user_id, Cash, SNo,Transby, Idno, mobile,branch,AI) VALUES ("
sql = sql & txtrno & ",'" & Lvwitems.SelectedItem & "','" & txtransdate & "'," & Lvwitems.SelectedItem.SubItems(4)
sql = sql & "," & Lvwitems.SelectedItem.SubItems(2) & "," & Remain & ",'" & User & "'," & Lvwitems.SelectedItem.SubItems(5) & ",'" & C & "','" & txttransby & "','" & txtidno & "','" & txtmobile & "','" & Cmbstation & "'," & txtserai & ")"

oSaccoMaster.ExecuteThis (sql)
oSaccoMaster.ExecuteThis ("Update ag_Products SET Qout =" & CCur(Remain) & " WHERE p_code= '" & Lvwitems.SelectedItem & "' and branch='" & Cmbstation & "'")
oSaccoMaster.ExecuteThis ("Update Rcpno SET rcpno =" & txtrno & "")

'\\ save to gl


    sql = ""
    sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & txtransdate & "'," & Lvwitems.SelectedItem.SubItems(4) & ",'" & lbldracc & "','" & lblcracc & "','" & Lvwitems.SelectedItem & "','" & cboproductname & "' ,'MPESA Sales- " & "" & cboproductname & "','" & User & "',1,0)"
    oSaccoMaster.ExecuteThis (sql)
'    sql = ""
'    sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & txtransdate & "','" & Bpprice & "','22-04','22-03','" & Lvwitems.SelectedItem & "','" & cboproductname & "' ,'MPESA Sales- " & "" & cboproductname & "','" & User & "',1,0)"
'    oSaccoMaster.ExecuteThis (sql)
'

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
End If
Next j


If chkPrint.value = vbChecked Then
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
        Dim fso, chkPrinter, txtFile
        'ttt = "LPT1" 'LPT1
         Dim PORT As String
        PORT = ports
        'ttt = "LPT1" 'LPT1
        ttt = PORT
        Set fso = CreateObject("Scripting.FileSystemObject")
        Dim strReceipts As String
        j = 1
        strReceipts = ""
        Do While Not j > (Lvwitems.ListItems.Count)
            Lvwitems.ListItems.Item(j).selected = True
            strReceipts = strReceipts & Lvwitems.SelectedItem.SubItems(1) & vbNewLine & Lvwitems.SelectedItem.SubItems(2) & vbTab & vbTab
            strReceipts = strReceipts & Format(Lvwitems.SelectedItem.SubItems(3), "#,##0.00") & vbTab & vbTab & Format(Lvwitems.SelectedItem.SubItems(4), "#,##0.00") & vbNewLine
            j = j + 1
        Loop

        'MsgBox strReceipts
        strReceipts = strReceipts & vbNewLine & "---------------------------------------" & vbNewLine
        strReceipts = strReceipts & "TOTAL" & vbTab & vbTab & vbTab & vbTab & Format(total, "#,##0.00") & vbNewLine
        strReceipts = strReceipts & "======================================="
        Set txtFile = fso.CreateTextFile(ttt, True)
        
        txtFile.WriteLine "      " & cname & ""
        txtFile.WriteLine "      AGROVET"
        txtFile.WriteLine "      " & paddress & ""
        txtFile.WriteLine "      " & town & ""
        txtFile.WriteLine "      " & Phone & ""
        'txtfile.WriteLine "      " & Email & ""
        
        txtFile.WriteLine "---------------------------------------"
        txtFile.WriteLine "  AGROVET RECEIPT"
        txtFile.WriteLine "     M-PESA SALES"
        txtFile.WriteLine "---------------------------------------"
        txtFile.WriteLine "QNTY" & vbTab & vbTab & "PRICE" & vbTab & vbTab & "AMOUNT"
        txtFile.WriteLine "........................................"
        txtFile.WriteLine "---------------------------------------"

        txtFile.WriteLine strReceipts
        txtFile.WriteLine
        
        txtFile.WriteLine "TOTAL" & vbTab & vbTab & vbTab & vbTab & Format(total, "#,##0.00") & vbNewLine
        txtFile.WriteLine
        txtFile.WriteLine "TOTAL" & vbTab & TXTTOTAL
        txtFile.WriteLine
        txtFile.WriteLine "AMOUNT RECEVED" & vbTab & txtamtreceived
        txtFile.WriteLine
        txtFile.WriteLine "CHANGE" & vbTab & TXTCHANGE
        txtFile.WriteLine
        txtFile.WriteLine "Trans By" & vbTab & txttransby
        txtFile.WriteLine "Id No" & vbTab & txtidno
        txtFile.WriteLine
        txtFile.WriteLine "YOU WERE SERVED By " & UCase(username)
        txtFile.WriteLine "    Date :" & Format(Get_Server_Date, "dd/mm/yyyy HH:MM:SS AM/PM")
        txtFile.WriteLine "     THANK YOU AND WELCOME "
        
        txtFile.WriteLine " GOODS ONCE SOLD WILL NOT BE RE-ACCEPTED"
        txtFile.WriteLine "****************************************"
        txtFile.WriteLine escFeedAndCut
        txtFile.Close
    End If
End If


Lvwitems.ListItems.Clear
txtrno = ""
txtpcode.Text = ""
txtquantity = 1
txtamount = ""
txttransby = ""
txtidno = ""
txtmobile = ""
MsgBox "Record saved Successfully"
Exit Sub
olkalou:
MsgBox err.description & " error occured."

End Sub

Private Sub cmdsigei_Click()
On Error GoTo ErrorHandler
'Startdate = DateSerial(year(txtransdate), month(txtransdate), 1)
'Enddate = DateSerial(year(txtransdate), month(txtransdate) + 1, 1 - 1)
'sql = ""
'sql = "set dateformat dmy delete from ag_ReceiptsProcess where Date>= '" & Startdate & "' And Date<='" & Enddate & "'"
'cn.Execute sql
'Label27.Visible = True
'Label27.Enabled = True
Dim U, C As Integer
sql = ""
sql = "set dateformat dmy Select count(distinct(SNO)) as u from ag_Receipts where SNO<>''and T_Date>= '" & Startdate & "' And T_Date<='" & Enddate & "'  "
Set rs = cn.Execute(sql)

  U = rs.Fields(0)
'MsgBox "Please wait " & U & ""
  sql = ""
  sql = "set dateformat dmy Select distinct(SNO) as y   from   ag_Receipts where SNO<>''and T_Date>= '" & Startdate & "' And T_Date<='" & Enddate & "' order by SNO asc  "
  Set rsg = cn.Execute(sql)
  While Not rs.EOF
  Do While Not U <= 0
  If Not rsg.EOF Then
    C = rsg.Fields(0)
    Label27.FontBold = True
    Label27 = "Please wait as we process"
       sql = ""
       sql = "set dateformat dmy Select distinct(Remarks)from ag_Receipts where S_No='" & C & "'and T_Date >= '" & Startdate & "' And T_Date<='" & Enddate & "' "
       Set rsb = cn.Execute(sql)
       sql = ""
       sql = "set dateformat dmy Select  count(distinct(Remarks))from ag_Receipts where S_No='" & C & "' and T_Date >= '" & Startdate & "' And T_Date<='" & Enddate & "' "
       Set rsh = cn.Execute(sql)
       Do While Not rsh.EOF
      If Not rsb.EOF Then
       If Not rsh.EOF Then
           sql = ""
           sql = "set dateformat dmy Select  P_code, T_Date, Amount, S_No,Remarks, Branch,Qua from ag_Receipts where S_No='" & C & "'and Remarks='" & rsb.Fields(0) & "' and T_Date >= '" & Startdate & "' And T_Date<='" & Enddate & "'order by S_No"
           Set rst = cn.Execute(sql)
          If Not rst.EOF Then
            sql = ""
            sql = "set dateformat dmy Select sum(Qua) from ag_Receipts where S_No='" & C & "' and Remarks='" & rst.Fields(4) & "'and T_Date >= '" & Startdate & "' And T_Date<='" & Enddate & "' "
            Set rsz = cn.Execute(sql)
                      
           sql = ""
           sql = "set dateformat dmy select * from ag_ReceiptsProcess where SNo='" & rsg.Fields(0) & "' and Date >= '" & Startdate & "' And Date<='" & Enddate & "'"
           Set rss = oSaccoMaster.GetRecordset(sql)
           If rss.EOF Then
             sql = ""
             sql = "set dateformat dmy insert into  ag_ReceiptsProcess(SNo, Name, Date, Branch)"
             sql = sql & "  values('" & rst.Fields(3) & "','','" & txtransdate & "','" & rst.Fields(5) & "')"
             cn.Execute sql
            Else
            End If
           sql = ""
           sql = "set dateformat dmy select Pro1, Pro2, Pro3, Pro4, Pro5, Pro6, Pro7, Pro8, Pro9, Pro10 from ag_ReceiptsProcess where SNo='" & rsg.Fields(0) & "' and Date >= '" & Startdate & "' And Date<='" & Enddate & "'"
           Set rsl = oSaccoMaster.GetRecordset(sql)
            
           sql = ""
           sql = "select p_code,p_name from ag_Products where p_name='" & rst.Fields(4) & "' ORDER BY p_code asc"
           Set rsm = oSaccoMaster.GetRecordset(sql)
           If Not rsm.EOF Then
           Dim strong As Integer
           strong = rsm.Fields(0)
                    'rsm.Fields(0) = " & rst.Fields(4) & "
            Select Case strong
             Case "1"
              sql = ""
              sql = "set dateformat dmy Update ag_ReceiptsProcess SET Pro1 ='" & rsl.Fields(0) + rsz.Fields(0) & "' WHERE SNo='" & rst.Fields(3) & "' and Date >= '" & Startdate & "'And Date<='" & Enddate & "'"
              cn.Execute sql
             Case "2"
              sql = ""
              sql = "set dateformat dmy Update ag_ReceiptsProcess SET Pro2 ='" & rsl.Fields(1) + rsz.Fields(0) & "' WHERE SNo='" & rst.Fields(3) & "'and Date >= '" & Startdate & "'And Date<='" & Enddate & "'"
              cn.Execute sql
             Case "3"
              sql = ""
              sql = "set dateformat dmy Update ag_ReceiptsProcess SET Pro3 ='" & rsl.Fields(2) + rsz.Fields(0) & "' WHERE SNo='" & rst.Fields(3) & "'and Date >= '" & Startdate & "'And Date<='" & Enddate & "'"
              cn.Execute sql
             Case "4"
              sql = ""
              sql = "set dateformat dmy Update ag_ReceiptsProcess SET Pro4 ='" & rsl.Fields(3) + rsz.Fields(0) & "' WHERE SNo='" & rst.Fields(3) & "'and Date >= '" & Startdate & "'And Date<='" & Enddate & "'"
              cn.Execute sql
             Case "5"
              sql = ""
              sql = "set dateformat dmy Update ag_ReceiptsProcess SET Pro5 ='" & rsl.Fields(4) + rsz.Fields(0) & "' WHERE SNo='" & rst.Fields(3) & "'and Date >= '" & Startdate & "'And Date<='" & Enddate & "'"
              cn.Execute sql
             Case "6"
              sql = ""
              sql = "set dateformat dmy Update ag_ReceiptsProcess SET Pro6 ='" & rsl.Fields(5) + rsz.Fields(0) & "' WHERE SNo='" & rst.Fields(3) & "'and Date >= '" & Startdate & "'And Date<='" & Enddate & "'"
              cn.Execute sql
             Case "7"
              sql = ""
              sql = "set dateformat dmy Update ag_ReceiptsProcess SET Pro7 ='" & rsl.Fields(6) + rsz.Fields(0) & "' WHERE SNo='" & rst.Fields(3) & "'and Date >= '" & Startdate & "'And Date<='" & Enddate & "'"
              cn.Execute sql
             Case "8"
              sql = ""
              sql = "set dateformat dmy Update ag_ReceiptsProcess SET Pro8 ='" & rsl.Fields(7) + rsz.Fields(0) & "' WHERE SNo='" & rst.Fields(3) & "'and Date >= '" & Startdate & "'And Date<='" & Enddate & "'"
              cn.Execute sql
             Case "9"
              sql = ""
              sql = "set dateformat dmy Update ag_ReceiptsProcess SET Pro9 ='" & rsl.Fields(8) + rsz.Fields(0) & "' WHERE SNo='" & rst.Fields(3) & "'and Date >= '" & Startdate & "'And Date<='" & Enddate & "'"
              cn.Execute sql
             Case "10"
              sql = ""
              sql = "set dateformat dmy Update ag_ReceiptsProcess SET Pro10 ='" & rsl.Fields(9) + rsz.Fields(0) & "' WHERE SNo='" & rst.Fields(3) & "'and Date >= '" & Startdate & "'And Date<='" & Enddate & "'"
              cn.Execute sql
             Case Else
            End Select
          
           End If
           End If
           rsb.MoveNext
           End If
         Else
           rsh.MoveNext
          End If
         Loop
        
        sql = ""
        sql = "set dateformat dmy Select SNo, Names from d_Suppliers where SNo='" & rsg.Fields(0) & "' "
        Set rst = cn.Execute(sql)
        sql = ""
        sql = "set dateformat dmy Update ag_ReceiptsProcess SET Name ='" & rst.Fields(1) & "' WHERE SNo='" & rsg.Fields(0) & "'and Date >= '" & Startdate & "'And Date<='" & Enddate & "'"
        cn.Execute sql
        
      Else
      Label27.Visible = False
     MsgBox "Completed succesfully ", vbInformation
    End If
   U = U - 1
 rsg.MoveNext
Loop
Label27.Visible = False
MsgBox "Completed succesfully ", vbInformation
Exit Sub
Wend

Exit Sub
ErrorHandler:
MsgBox err.description
End Sub

Private Sub cmdstorece_Click()
frmproduct1s.Show vbModal
End Sub

Private Sub Command1_Click()
frmproduct1s.Show vbModal
End Sub

Private Sub Command2_Click()
Dim total As Double
Dim j, Coun As Integer
j = 1
On Error GoTo ErrorHandler
TXTTOTAL = 0
'If Lvwitems.ListItems.Count > 0 Then
''Total = CCur(txttotal - li.SubItems(4))
Lvwitems.ListItems.Remove (Lvwitems.SelectedItem.Index)  '// removes the selected item

Do While Not j > (Lvwitems.ListItems.Count)
'For j = 1 To Lvwitems.ListItems.Count
 Lvwitems.ListItems.Item(j).selected = True
 total = total + CCur(Lvwitems.SelectedItem.SubItems(4))
 TXTTOTAL = total
j = j + 1
Loop

'End If
Exit Sub
ErrorHandler:
MsgBox err.description

End Sub

Private Sub Command3_Click()
'Expiredgoods
reportname = "Expiredgoods.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, ""


End Sub

Private Sub Command4_Click()
'check the user
sql = "SELECT     UserLoginIDs,levels, UserGroup, SUPERUSER From UserAccounts where UserLoginIDs='" & User & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
If rs!SuperUser <> "1" Then
MsgBox "You are not allowed to sell", vbInformation
Exit Sub

End If
End If

frmproductrepackaging.Show vbModal
End Sub

Private Sub Command5_Click()

frmAI.Show vbModal, Me
End Sub

Private Sub Command6_Click()
'check the user
sql = "SELECT     UserLoginIDs,levels, UserGroup, SUPERUSER From UserAccounts where UserLoginIDs='" & User & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
If rs!SuperUser <> "1" Then
MsgBox "You are not allowed ", vbInformation
Exit Sub

End If
End If
frmdrawnstock.Show vbModal
End Sub
Public Sub loadBranchesTypes()
    
    With ListView1
    
        .ListItems.Clear
        
        .ColumnHeaders.Clear
    
    End With

    Set rs2 = CreateObject("adodb.recordset")
    
    sql = "set dateformat dmy Select d.S_No,Qua,m.Names,d.Amount,d.Remarks, d.Branch from ag_Receipts AS d INNER JOIN d_Suppliers AS m ON d.S_No = m.SNo where T_Date='" & txtransdate & "' and user_id='" & User & "' order by audit_date desc"
'    sql = ""
'    sql = "set dateformat dmy SELECT d.RefNo,m.DName, d.DispDate, d.DispQnty,d.Amount,d.PaidAmount FROM d_MilkControl AS d INNER JOIN d_Debtors AS m ON d.DCode = m.DCode WHERE     (DispDate = '" & txtdateenterered & "') and vehicleno='" & cboVehicle & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    
    Set rs2 = CreateObject("adodb.recordset")
    
    Set clsClass = New cdbase
    
    Provider = clsClass.OpenCon
    
    Set cn = CreateObject("adodb.connection")
    
   cn.Open Provider, "atm", "atm"
    
    rs2.Open sql, cn
    
    With ListView1
        
        .ColumnHeaders.Add , , "SNo"
        .ColumnHeaders.Add , , "Name"
        .ColumnHeaders.Add , , "Quantity"
        .ColumnHeaders.Add , , "Amount"
        .ColumnHeaders.Add , , "Remarks"
        .ColumnHeaders.Add , , "Branch"
'        .ColumnHeaders.Add , , "Mpesa"
'        .ColumnHeaders.Add , , "Outlet"
        While Not rs2.EOF
        
            Set li = .ListItems.Add(, , Trim(rs2.Fields("S_No")))
            
            li.ListSubItems.Add , , Trim(rs2.Fields("Names"))
            li.ListSubItems.Add , , Trim(rs2.Fields("Qua"))
            li.ListSubItems.Add , , Trim(rs2.Fields("Amount"))
            li.ListSubItems.Add , , Trim(rs2.Fields("Remarks"))
            li.ListSubItems.Add , , Trim(rs2.Fields("Branch"))
            rs2.MoveNext
        
        Wend
        
    End With
    
    rs2.Close
    
    Set rs2 = Nothing
    
ListView1.View = lvwReport
End Sub


Private Sub Command7_Click()

End Sub

Private Sub Form_Load()
On Error GoTo ErrorHandler
Label5.Visible = False
txttranscode.Visible = False
lbltransportername.Visible = False
Label10.Visible = False
lbltransnetpay.Visible = False
chkhalf.value = vbUnchecked
txtransdate = Format(Get_Server_Date, "dd/mm/yyyy")
DTPfrom = DateSerial(Year(txtransdate), Month(txtransdate), 1)
DTPto = DateSerial(Year(DTPfrom), Month(DTPfrom) + 1, 1 - 1)


Provider = "MAZIWA"
Set cn = New ADODB.Connection
cn.Open Provider, "atm", "atm"
'sql = "SELECT     UserLoginIDs, UserGroup, Branch,SUPERUSER From UserAccounts where UserLoginIDs='" & User & "'"
'    Set rs = oSaccoMaster.GetRecordset(sql)
'    If Not rs.EOF Then
'    Cmbstation = rs!Branch
'    If rs!Branch <> Cmbstation Then
'    MsgBox "You are not allowed to sell in this branch", vbInformation
'    Exit Sub
'    End If
'    End If
'If reportpath = "" Then reportpath = GetSetting("payroll", "AppName", "rptPath", rptPath)
'sql = "select P_NAME  from ag_products where branch='" & Cmbstation & "' ORDER BY P_NAME ASC"
sql = "select P_NAME  from ag_products  ORDER BY P_NAME ASC"
Set rs = New ADODB.Recordset
rs.Open sql, cn

While Not rs.EOF
cboproductname.AddItem rs.Fields(0)
rs.MoveNext
Wend

    'Set rst = New Recordset
   ' Dim cn As Connection
    Set cn = New ADODB.Connection
    Provider = "MAZIWA"
    'Provider = cnn
    cn.Open Provider, "atm", "atm"
    'Set rst = New Recordset
    sql = ""
    sql = "Select distinct(Bname) from   d_Branch order by Bname"
    Set rst = oSaccoMaster.GetRecordset(sql)
   ' rst.Open sql, cn, adOpenKeyset, adLockOptimistic
    While Not rst.EOF
    Cmbstation.AddItem rst.Fields(0)
    rst.MoveNext
    Wend

    lbldracc = "33-103"

    lblcracc = "33-302"


cboproductname.Enabled = True
chkPrint.value = vbUnchecked
Label27.Visible = False
Label27.Enabled = True
Exit Sub
ErrorHandler:

MsgBox err.description
End Sub
Private Sub cboname()
'Provider = cn
'Set cn = New ADODB.Connection
''cn.Open Provider, "atm"
''If reportpath = "" Then reportpath = GetSetting("payroll", "AppName", "rptPath", rptPath)
'sql = "select P_NAME from ag_products where p_Name='" & cboproductname.Text & "'"
'Set rs = New ADODB.Recordset
'Set rs = oSaccoMaster.GetRecordset(sql)
'If Not rs.EOF Then
'If Not IsNull(rs.Fields(0)) Then cboproductname.Text = (rs.Fields(0))
'If Not IsNull(rs.Fields(1)) Then lblbalance = rs.Fields(1)
'End If
End Sub

Private Sub lblCheckOff_Click()
lblSNo.Visible = True
txtSNo.Visible = True
lblNetPay.Visible = True
lblNPay.Visible = True
lblDed.Visible = True
lblTKgs.Visible = True
lblGPay.Visible = True
Label11.Visible = True
Label13.Visible = True
Label8.Visible = True
txttranscode.Visible = False
Label5.Visible = False
Label10.Visible = False
lbltransnetpay.Visible = False
lbltransportername.Visible = False
lblSNames.Visible = True
End Sub

Private Sub lblcracc_Change()
 Set rst = oSaccoMaster.GetRecordset("select glaccname from glsetup where accno='" & lblcracc & "'")
    If Not rst.EOF Then
    txtcracc = rst.Fields("glaccname")
    End If
End Sub

Private Sub lbldracc_Change()
 Set rst = oSaccoMaster.GetRecordset("select glaccname from glsetup where accno='" & lbldracc & "'")
    If Not rst.EOF Then
    txtdracc = rst.Fields("glaccname")
    End If
End Sub

Private Sub Optbranch_Click()
lblSNo.Visible = False
txtSNo.Visible = False
Label13.Visible = False
lblTKgs.Visible = False
Label11.Visible = False
lblGPay.Visible = False
Label8.Visible = False
lblDed.Visible = False
lblNetPay.Visible = False
lblNPay.Visible = False
Label5.Visible = False
txttranscode.Visible = False
Label10.Visible = False
lbltransnetpay.Visible = False
lbltransportername.Visible = False
End Sub

Private Sub Optcash_Click()
lblSNo.Visible = False
txtSNo.Visible = False
lblNetPay.Visible = False
lblNPay.Visible = False

lblDed.Visible = False
lblTKgs.Visible = False
lblGPay.Visible = False
Label11.Visible = False
Label13.Visible = False
Label8.Visible = False
lblSNames.Visible = False

'If Cmbstation <> "OLENGURUONE" Then
'MsgBox "Please yours branch is only allowed to sell using Mpesa", vbInformation
'Exit Sub
'End If

End Sub

Private Sub Optmpesa_Click()
lblSNo.Visible = False
txtSNo.Visible = False
lblNetPay.Visible = False
lblNPay.Visible = False

lblDed.Visible = False
lblTKgs.Visible = False
lblGPay.Visible = False
Label11.Visible = False
Label13.Visible = False
Label8.Visible = False
lblSNames.Visible = False
End Sub

Private Sub Optothers_Click()
lblSNo.Visible = False
txtSNo.Visible = False
lblNetPay.Visible = False
lblNPay.Visible = False

lblDed.Visible = False
lblTKgs.Visible = False
lblGPay.Visible = False
Label11.Visible = False
Label13.Visible = False
Label8.Visible = False
End Sub

Private Sub optstaff_Click()
lblSNo.Visible = False
txtSNo.Visible = False
lblNetPay.Visible = False
lblNPay.Visible = False

lblDed.Visible = False
lblTKgs.Visible = False
lblGPay.Visible = False
Label11.Visible = False
Label13.Visible = False
Label8.Visible = False

End Sub

Private Sub opttransport_Click()
If opttransport = True Then
Label5.Visible = True
txttranscode.Visible = True
lbltransportername.Visible = True
Label10.Visible = True
lbltransnetpay.Visible = True
lblSNames.Visible = False

lblSNo.Visible = False
txtSNo.Visible = False
Label13.Visible = False
lblTKgs.Visible = False
Label11.Visible = False
lblGPay.Visible = False
Label8.Visible = False
lblDed.Visible = False
lblNetPay.Visible = False
lblNPay.Visible = False
lblSNames.Visible = False

Else
Label5.Visible = False
txttranscode.Visible = False
lbltransportername.Visible = False
Label10.Visible = False
lbltransnetpay.Visible = False
End If
End Sub

Private Sub opttransport_Validate(Cancel As Boolean)
opttransport_Click
End Sub

Private Sub Cmbstation_KeyPress(KeyAscii As Integer)
'KeyAscii = 0
End Sub

Private Sub Picture1_Click()
frmSearch.Show vbModal
Dim Y As String
Y = sel
Dim p As Integer
If Y <> "" Then
'Provider = cn
Set cn = New ADODB.Connection
'cn.Open Provider, "atm"
'If reportpath = "" Then reportpath = GetSetting("payroll", "AppName", "rptPath", rptPath)
If Cmbstation = "" Then
MsgBox "Please select the branch", vbInformation
Exit Sub
End If
sql = "select P_CODE,P_NAME,S_NO,QOUT,seria,s_no,pprice,sprice from ag_products where p_code='" & Y & "' and branch='" & Cmbstation & "'"
Set rs = New ADODB.Recordset
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
If Not IsNull(rs.Fields(0)) Then txtpcode.Text = (rs.Fields(0))
If Not IsNull(rs.Fields(4)) Then p = (rs.Fields(4))
If p = 1 Then
If Not IsNull(rs.Fields(5)) Then 'txtserialno = (rs.Fields(5))
'lblserialno.Visible = True
'txtserialno.Visible = True
Else
'lblserialno.Visible = False
'txtserialno.Visible = False
End If
End If

If Not IsNull(rs.Fields(1)) Then cboproductname = (rs.Fields(1))
If Not IsNull(rs.Fields(0)) Then txtpcode = (rs.Fields(0))
If Not IsNull(rs.Fields(6)) Then txtbuyingprice = (rs.Fields(6))
If Not IsNull(rs.Fields(7)) Then txtsellingprice = (rs.Fields(7))
If Not IsNull(rs.Fields(3)) Then lblbalance = (rs.Fields(3))
'// check if it has the serial numbers
'get_serialno Y
End If

'// check if the product have the serial then show the ag_receipts details
cboproductname_Validate True

End If
End Sub
Private Sub get_serialno(pcode As String)
Provider = cn
Set cn = New ADODB.Connection
cn.Open Provider, "atm", "atm"
Dim RSSE As Recordset
sql = ""
sql = "select top 1 serialno,p_code,used from serialno where p_code='" & txtpcode & "'  order by serialid desc"
Set RSSE = New ADODB.Recordset

RSSE.Open sql, cn, adOpenKeyset, adLockOptimistic
If RSSE.Fields(2) = 1 Then
MsgBox "Serial Number and receipt no used please check again before posting", vbCritical
Exit Sub
End If
End Sub
Private Sub Picture2_Click()
On Error Resume Next
frmsearchre.Show vbModal
Dim Y As String
Y = sel

If Y <> "" Then
Provider = cn
Set cn = New ADODB.Connection
cn.Open Provider, "atm", "atm"
sql = ""
'If reportpath = "" Then reportpath = GetSetting("payroll", "AppName", "rptPath", rptPath)
sql = "select r_no,P_CODE,S_NO,Qua,amount from ag_receipts where r_no=" & Y & ""
Set rs = New ADODB.Recordset
rs.Open sql, cn
If Not rs.EOF Then
If Not IsNull(rs.Fields(0)) Then txtrno = (rs.Fields(0))
If Not IsNull(rs.Fields(1)) Then txtpcode = (rs.Fields(1))
'If Not IsNull(rs.Fields(2)) Then txtserialno = (rs.Fields(2))
If Not IsNull(rs.Fields(3)) Then txtquantity = (rs.Fields(3))
If Not IsNull(rs.Fields(4)) Then txtamount = (rs.Fields(4))
If Not IsNull(rs.Fields(3)) Then lblbalance = (rs.Fields(3))
Call cboname
End If
End If
End Sub

Private Sub txtpassword_LostFocus()
'fra1.Visible = True
Dim rsp As Recordset
Set cn = CreateObject("adodb.connection")
Provider = cn
Set cn = New ADODB.Connection
cn.Open Provider, "atm", "atm"
Set rsp = CreateObject("adodb.recordset")
sql = "select *  from useraccounts where UserLoginIDs='" & User & "' and usergroup='administrator'"
rsp.Open sql, cn
Dim pass As String


txtransdate = Format(Date, "DD/MM/YYYY")
'fra1.Visible = True
'End If
End Sub
Private Sub Picture5_Click()
        Me.MousePointer = vbHourglass
        frmSearchSupplier.Show vbModal
        txtSNo = sel
        txtSNo_Validate True
        Me.MousePointer = 0
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Picture3_Click()
Me.MousePointer = vbHourglass
        frmSearchGLAccounts.Show vbModal, Me
    If Continue Then
        If SearchValue <> "" Then
            lblcracc = SearchValue
            SearchValue = ""
        End If
    End If
    Me.MousePointer = 0

End Sub

Private Sub Picture4_Click()
Me.MousePointer = vbHourglass
        frmSearchGLAccounts.Show vbModal, Me
    If Continue Then
        If SearchValue <> "" Then
            lbldracc = SearchValue
            SearchValue = ""
        End If
    End If
    Me.MousePointer = 0
End Sub

Private Sub txtamtreceived_Change()
On Error Resume Next
TXTCHANGE = txtamtreceived - TXTTOTAL
End Sub

Private Sub txtpcode_KeyPress(KeyAscii As Integer)
'//TWNG001
If KeyAscii = 13 Then
Provider = "MAZIWA"
Set cn = New ADODB.Connection
cn.Open Provider, "atm", "atm"
'If reportpath = "" Then reportpath = GetSetting("payroll", "AppName", "rptPath", rptPath)
sql = "select P_CODE,P_NAME,S_NO,QOUT,supplierid,pprice,sprice from ag_products where p_code='" & txtpcode & "'"
Set rs = New ADODB.Recordset
rs.Open sql, cn
If Not rs.EOF Then
 
If Not IsNull(rs.Fields(1)) Then cboproductname = (rs.Fields(1))
If Not IsNull(rs.Fields(5)) Then txtbuyingprice = (rs.Fields(5))
If Not IsNull(rs.Fields(6)) Then txtsellingprice = (rs.Fields(6))

End If
End If
'// check with serial no if it exist
End Sub



Private Sub txtquantity_Validate(Cancel As Boolean)
If Not IsNumeric(txtquantity) Then
MsgBox "Enter a value please", vbCritical
txtquantity = ""
txtquantity.SetFocus
Exit Sub
End If

End Sub

Private Sub txtransdate_change()
'fra1.Visible = True
''txtransdate = Format(Get_Server_Date, "dd/mm/yyyy")
''DTPfrom = DateSerial(year(txtransdate), month(txtransdate), 1)
''DTPto = DateSerial(year(DTPfrom), month(DTPfrom) + 1, 1 - 1)

DTPfrom = DateSerial(Year(txtransdate), Month(txtransdate), 1)
DTPto = DateSerial(Year(txtransdate), Month(txtransdate) + 1, 1 - 1)
loadBranchesTypes
End Sub

Private Sub txtransdate_KeyPress(KeyAscii As Integer)
'fra1.Visible = True
End Sub

Private Sub txtransdate_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'fra1.Visible = True
End Sub
Private Sub txtpassword_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
txtpassword_LostFocus
End Sub

Private Sub txtpcode_LostFocus()
Call cboname

End Sub
Private Sub txtserialno_LostFocus()
Dim rss As ADODB.Recordset
Dim rsproduct As ADODB.Recordset
sql = ""
sql = "select * from ag_products where seria=1 AND P_CODE='" & txtpcode & "'"
Set rsproduct = New ADODB.Recordset
rsproduct.Open sql, cn, adOpenKeyset, adLockOptimistic
If Not rsproduct.EOF Then
sql = ""
sql = "select serialno  from serialno "
Set rss = New ADODB.Recordset
rss.Open sql, cn, adOpenKeyset, adLockOptimistic
 If Not rss.EOF Then
'// check if gth
 While Not rss.EOF
 Dim ser As String
 ser = rss.Fields(0)

'If ser = txtserialno Then GoTo hererere

 rss.MoveNext
 Wend
 Else
  MsgBox "Serial no not in our database", vbInformation

DIA = 1
Exit Sub
End If
End If
hererere:
End Sub

Private Sub txtSNo_Validate(Cancel As Boolean)
On Error GoTo ErrorHandler

 '''*****************************check if supliers branch is correct ****
sql = ""
sql = "select SNo,Location from d_Suppliers where SNo='" & txtSNo & "'"
Set rss = oSaccoMaster.GetRecordset(sql)
If Not rss.EOF Then
If rss!Location <> Cmbstation Then
 MsgBox "This supplier is not from this branch: " & Cmbstation & "", vbInformation
 Exit Sub
End If
Else
Exit Sub
End If

'''*****************************end **************


Set rs = New ADODB.Recordset
sql = "d_sp_SelectSuppliers '" & txtSNo & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
If Not IsNull(rs.Fields(0)) Then lblSNames = rs.Fields(2)
Else
lblSNames = ""
End If

Startdate = DateSerial(Year(txtransdate), Month(txtransdate), 1)
Enddate = DateSerial(Year(txtransdate), Month(txtransdate) + 1, 1 - 1)

Set rs = oSaccoMaster.GetRecordset("d_sp_SupNet '" & txtSNo & "','" & Startdate & "','" & Enddate & "', 0")

If Not rs.EOF Then
lblTKgs = IIf(IsNull(rs.Fields(0)), 0, rs.Fields(0))
lblGPay = IIf(IsNull(rs.Fields(1)), 0, rs.Fields(1))
Else
lblTKgs = "0.00"
lblGPay = "0.00"
End If





'If Not IsNull(rs.Fields(1)) Then
'lblGPay = rs.Fields(1)
'Else
'lblGPay = "0.00"
'End If

Set rs = oSaccoMaster.GetRecordset("d_sp_SupNet '" & txtSNo & "','" & Startdate & "','" & Enddate & "', 1")
If Not IsNull(rs.Fields(0)) Then
lblDed = rs.Fields(0)
Else
lblDed = "0.00"
End If

lblNPay = Format((CCur(lblGPay) - CCur(lblDed)), "#,##0.00")

Exit Sub
ErrorHandler:
MsgBox err.description
End Sub

Private Sub txtStaffNo_Change()
On Error GoTo ErrorHandler
Set rs = New ADODB.Recordset
sql = "select staffno,staffname from staffs where staffno= '" & txtstaffno & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
If Not IsNull(rs.Fields(0)) Then lblstnames = rs.Fields(1)
Else
lblstnames = ""
End If
ErrorHandler:
'MsgBox err.description
End Sub

Private Sub txttotal_Change()
On Error Resume Next
TXTCHANGE = txtamtreceived - TXTTOTAL
End Sub

Private Sub txttranscode_Change()
Set rs = New ADODB.Recordset
Dim DTPfrom As Date
sql = "d_sp_TransEnquiry  '" & txttranscode & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
If Not IsNull(rs.Fields(0)) Then lbltransportername = rs.Fields(0)
End If
DTPfrom = DateSerial(Year(txtransdate), Month(txtransdate), 1)
DTPto = DateSerial(Year(txtransdate), Month(txtransdate) + 1, 1 - 1)
'oSaccoMaster.ExecuteThis ("d_sp_UpdateTranstmpEnquery '" & txttranscode & "','" & DTPto & "'")
'oSaccoMaster.ExecuteThis ("d_sp_UpdateTranstmpEnqueryDed '" & txttranscode & "','" & DTPfrom & "','" & DTPto & "'")
'
'sql = ""
'sql = "SELECT     TOP 1 Bal  FROM         d_tmpTransEnquery WHERE     (Code = '" & txttranscode & "') ORDER BY Bal DESC"
'Set Rst = oSaccoMaster.GetRecordset(sql)
'If Not Rst.EOF Then
'lbltransnetpay = IIf(IsNull(Rst.Fields(0)), 0, Rst.Fields(0))
'End If
' get transporter netpay
   Dim mMonth, yYear As Integer
   mMonth = Month(txtransdate)
   yYear = Year(txtransdate)
   
  sql = " Select(Select isnull(SUM(Amount + Subsidy),0) from d_TransDetailed where Trans_Code='" & txttranscode & "' and MMonth= " & mMonth & " and YYear=" & yYear & "),"
  sql = sql & " (Select isnull(SUM(Amount),0) from d_Transport_Deduc where TransCode='" & txttranscode & "' and MONTH(TDate_Deduc)=" & mMonth & " and YEAR(TDate_Deduc)= " & yYear & ")"
   Set rs2 = oSaccoMaster.GetRecordset(sql)
   If Not rs2.EOF Then
   lbltransnetpay = Format(rs2.Fields(0) - rs2.Fields(1), Cfmt)
  
   Else
   lbltransnetpay = "0.00"
   
   End If
End Sub
