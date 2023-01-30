VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmplantmilksale 
   Caption         =   "Plant Milk sales and Boiling"
   ClientHeight    =   8835
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10950
   LinkTopic       =   "Form1"
   ScaleHeight     =   8835
   ScaleWidth      =   10950
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   9480
      TabIndex        =   39
      Top             =   5880
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker txtdateenterered 
      Height          =   375
      Left            =   9000
      TabIndex        =   1
      Top             =   1080
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   144244737
      CurrentDate     =   38814
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   15478
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   49152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Elephant"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Customer Registration"
      TabPicture(0)   =   "frmplantmilksale.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label11"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label12"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label14"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "ListView2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtcust"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cboName1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Command1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Command2"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdSearch"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtDrAccNo"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lblDrAccName"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Command4"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).ControlCount=   15
      TabCaption(1)   =   "Point Of Sales"
      TabPicture(1)   =   "frmplantmilksale.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label8"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label13"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label5"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label6(0)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label7"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label6(1)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label9"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label10"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label4"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label6(2)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label6(3)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label6(4)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Lvwitems"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "ListView3"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "fra1"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "cboNamecust"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "txtCustName"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "txtprice"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "chkBoil"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "chksales"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "chklocal"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "txtamount"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "cmdsave"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "txtquantity"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "Command3"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "chkRepay"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "cmdnextitem"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "TXTTOTAL"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "TXTCHANGE"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "cmdremove"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "cmd5"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "cmddelete"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "cmdmonthly"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "cmdindi"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).ControlCount=   34
      Begin VB.CommandButton cmdindi 
         Caption         =   "Individual Report"
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
         Left            =   4080
         TabIndex        =   57
         Top             =   6000
         Width           =   1575
      End
      Begin VB.CommandButton cmdmonthly 
         Caption         =   "Monthly Sales Report"
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
         Left            =   2160
         TabIndex        =   56
         Top             =   6000
         Width           =   1815
      End
      Begin VB.CommandButton cmddelete 
         Caption         =   "Delete"
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
         Left            =   7440
         TabIndex        =   55
         Top             =   6000
         Width           =   975
      End
      Begin VB.CommandButton cmd5 
         Caption         =   "Gemija Report"
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
         Left            =   5880
         TabIndex        =   54
         Top             =   6000
         Width           =   1455
      End
      Begin VB.CommandButton cmdremove 
         Caption         =   "Remove"
         Height          =   495
         Left            =   6480
         TabIndex        =   53
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox TXTCHANGE 
         Height          =   495
         Left            =   9360
         TabIndex        =   50
         Text            =   "0"
         Top             =   5160
         Width           =   1215
      End
      Begin VB.TextBox TXTTOTAL 
         Enabled         =   0   'False
         Height          =   495
         Left            =   9360
         TabIndex        =   49
         Text            =   "0"
         Top             =   3840
         Width           =   1215
      End
      Begin VB.CommandButton cmdnextitem 
         Caption         =   "Next item"
         Default         =   -1  'True
         Height          =   495
         Left            =   8400
         TabIndex        =   47
         Top             =   2760
         Width           =   1455
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Report"
         Height          =   495
         Left            =   -70320
         TabIndex        =   46
         Top             =   7320
         Width           =   1335
      End
      Begin VB.TextBox lblDrAccName 
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
         Left            =   -71280
         TabIndex        =   42
         Top             =   2640
         Width           =   3225
      End
      Begin VB.TextBox txtDrAccNo 
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
         Left            =   -72420
         TabIndex        =   41
         Top             =   2640
         Width           =   1080
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "..."
         Height          =   285
         Left            =   -72720
         TabIndex        =   40
         Top             =   2640
         Width           =   300
      End
      Begin VB.CheckBox chkRepay 
         Caption         =   "Pay for 2nd time"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   6840
         TabIndex        =   38
         Top             =   3600
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Plan Sales Reports"
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
         Left            =   480
         TabIndex        =   35
         Top             =   6000
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Save"
         Height          =   495
         Left            =   -72240
         TabIndex        =   34
         Top             =   7320
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "New"
         Height          =   435
         Left            =   -74160
         TabIndex        =   33
         Top             =   7320
         Width           =   975
      End
      Begin VB.TextBox txtquantity 
         Height          =   375
         Left            =   2160
         TabIndex        =   27
         Top             =   2040
         Width           =   1335
      End
      Begin VB.CommandButton cmdsave 
         Caption         =   "Save"
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
         Left            =   8520
         TabIndex        =   25
         Top             =   6000
         Width           =   855
      End
      Begin VB.TextBox txtamount 
         Height          =   405
         Left            =   9360
         TabIndex        =   23
         Top             =   4560
         Width           =   1215
      End
      Begin VB.CheckBox chklocal 
         Caption         =   "Local Sales"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6480
         TabIndex        =   22
         Top             =   2760
         Width           =   1695
      End
      Begin VB.CheckBox chksales 
         Caption         =   "Sales from Siche"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4080
         TabIndex        =   21
         Top             =   2640
         Width           =   1935
      End
      Begin VB.CheckBox chkBoil 
         Caption         =   "Boiling"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   1920
         TabIndex        =   20
         Top             =   2640
         Width           =   1575
      End
      Begin VB.TextBox txtprice 
         Height          =   405
         Left            =   4680
         TabIndex        =   19
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox txtCustName 
         Height          =   405
         Left            =   2160
         TabIndex        =   15
         Top             =   840
         Width           =   1575
      End
      Begin VB.ComboBox cboNamecust 
         Height          =   315
         Left            =   2160
         TabIndex        =   14
         Top             =   1560
         Width           =   4815
      End
      Begin VB.Frame fra1 
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   1800
         TabIndex        =   4
         Top             =   3240
         Width           =   4695
         Begin VB.PictureBox Picture1 
            Height          =   255
            Left            =   1320
            Picture         =   "frmplantmilksale.frx":0038
            ScaleHeight     =   195
            ScaleWidth      =   195
            TabIndex        =   8
            Top             =   240
            Width           =   255
         End
         Begin VB.PictureBox Picture3 
            Height          =   255
            Left            =   1320
            Picture         =   "frmplantmilksale.frx":0902
            ScaleHeight     =   195
            ScaleWidth      =   195
            TabIndex        =   7
            Top             =   720
            Width           =   255
         End
         Begin VB.TextBox txtdracc 
            Height          =   375
            Left            =   1680
            TabIndex        =   6
            Top             =   120
            Width           =   2895
         End
         Begin VB.TextBox txtcracc 
            Height          =   375
            Left            =   1680
            TabIndex        =   5
            Top             =   600
            Width           =   2895
         End
         Begin VB.Label lbldracc 
            BackColor       =   &H8000000E&
            Height          =   375
            Left            =   120
            TabIndex        =   10
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label lblcracc 
            BackColor       =   &H8000000E&
            Height          =   375
            Left            =   120
            TabIndex        =   9
            Top             =   600
            Width           =   1095
         End
      End
      Begin VB.ComboBox cboName1 
         Height          =   315
         Left            =   -73200
         TabIndex        =   3
         Top             =   1560
         Width           =   3615
      End
      Begin VB.TextBox txtcust 
         Height          =   285
         Left            =   -73200
         TabIndex        =   2
         Top             =   960
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   3255
         Left            =   -74160
         TabIndex        =   13
         Top             =   3720
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   5741
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777152
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Ledger Accno"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   2055
         Left            =   480
         TabIndex        =   26
         Top             =   6600
         Width           =   10095
         _ExtentX        =   17806
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Quantity"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Price"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Amount"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Paid"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView Lvwitems 
         Height          =   1695
         Left            =   480
         TabIndex        =   48
         Top             =   4320
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   2990
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
         MouseIcon       =   "frmplantmilksale.frx":11CC
         NumItems        =   8
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
            Text            =   "Description"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Debtor"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Cr"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label6 
         Caption         =   "Bal."
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   8640
         TabIndex        =   52
         Top             =   5280
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   8520
         TabIndex        =   51
         Top             =   3960
         Width           =   1095
      End
      Begin VB.Label Label14 
         Caption         =   "Ledger Name:"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -70800
         TabIndex        =   45
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label Label12 
         Caption         =   "Accno:"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -72360
         TabIndex        =   44
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label Label11 
         Caption         =   "Ledger Name:"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74760
         TabIndex        =   43
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   6720
         TabIndex        =   37
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFC0&
         Height          =   375
         Left            =   7680
         TabIndex        =   36
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Customer Name:"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74760
         TabIndex        =   32
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Customer No:"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74760
         TabIndex        =   31
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Date:"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -68400
         TabIndex        =   30
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label10 
         Caption         =   "Date:"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8160
         TabIndex        =   29
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "Quantity:"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   28
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Receive"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   8520
         TabIndex        =   24
         Top             =   4560
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Price:"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   18
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Customer Name"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   17
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Customer Code:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label13 
         Caption         =   "Dr"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   12
         Top             =   3360
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "Cr "
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   11
         Top             =   3960
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmplantmilksale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As String

Private Sub cboNamecust_Click()
Provider = "MAZIWA"
Set cn = New ADODB.Connection
cn.Open Provider, "atm", "atm"
Set rst = New ADODB.Recordset
rst.Open sql, cn
'If rs.EOF Then
Set rst = oSaccoMaster.GetRecordset("select Code, Name,Accno from d_Outsales where Name ='" & cboNamecust & "'")
If Not rst.EOF Then
txtCustName.Text = rst.Fields("Code")
cboproductname1 = rst.Fields("Name")
lbldracc = rst.Fields("Accno")
'lblcracc = rst.Fields("Accno")
'txtsel
End If
 loadoutsale
'End If
txtquantity.SetFocus
End Sub

Private Sub chkBoil_Click()
If chkBoil = 1 Then
chksales.Visible = False
chklocal.Visible = False
If chkBoil = 1 Then

If chkRepay = vbuncheked Then
If txtquantity = "" Then
MsgBox "Please insert quantity", vbInformation
Exit Sub
End If
If txtprice = "" Then
MsgBox "Please insert the price", vbInformation
chkBoil = 0
Exit Sub
End If
Label4 = txtquantity * txtprice
Else
End If
a = "Boiling fee"
lblcracc = "P012"
'txtcracc = "DANDORA 2"
End If
Else
lblcracc = ""
txtcracc = ""
chksales.Visible = True
chkBoil.Visible = True
chklocal.Visible = True
Label4 = ""
End If
End Sub

Private Sub chklocal_Click()
If chklocal = 1 Then
chkBoil.Visible = False
chksales.Visible = False
If chklocal = 1 Then
 If chkRepay = vbuncheked Then
If txtquantity = "" Then
MsgBox "Please insert quantity", vbInformation
Exit Sub
End If
If txtprice = "" Then
MsgBox "Please insert the price", vbInformation
Exit Sub
End If
Label4 = txtquantity * txtprice
Else
End If
a = "Local Sales"
lbldracc = "P010"
'lblcracc = ""

Provider = "MAZIWA"
Set cn = New ADODB.Connection
cn.Open Provider, "atm", "atm"
Set rst = New ADODB.Recordset
rst.Open sql, cn
'If rs.EOF Then
Set rst = oSaccoMaster.GetRecordset("select Code, Name,Accno from d_Outsales where Name ='" & cboNamecust & "'")
If Not rst.EOF Then
lblcracc = rst.Fields("Accno")
'txtsel
End If
End If
Else
lbldracc = ""
txtdracc = ""
chksales.Visible = True
chkBoil.Visible = True
chklocal.Visible = True
Label4 = ""
End If
End Sub

Private Sub chkRepay_Click()
If chkRepay = vbcheked Then
fra1.Visible = True
Label13.Visible = True
Label8.Visible = True
Else
fra1.Visible = False
Label13.Visible = False
Label8.Visible = False
End If
End Sub

Private Sub chksales_Click()
If chksales = 1 Then
chkBoil.Visible = False
chklocal.Visible = False
If chksales = 1 Then

If chkRepay = vbuncheked Then
If txtquantity = "" Then
MsgBox "Please insert quantity", vbInformation
Exit Sub
End If
If txtprice = "" Then
MsgBox "Please insert the price", vbInformation
chksales = 0
Exit Sub
End If
Label4 = txtquantity * txtprice
Else
End If
a = "Sales from siche"
lblcracc = "P011"
'txtcracc = ""
End If
Else
lblcracc = ""
txtcracc = ""
chksales.Visible = True
chkBoil.Visible = True
chklocal.Visible = True
Label4 = ""
End If
End Sub

'Private Sub cmdClose_Click()
'Unload Me
'End Sub
Public Sub loadBranchesTypes()
    
    With ListView2
    
        .ListItems.Clear
        
        .ColumnHeaders.Clear
    
    End With

    Set rs2 = CreateObject("adodb.recordset")
    
    'sql = "Select * from d_Outsales"
    sql = "set dateformat dmy SELECT d.Code, d.Name, d.Accno, m.GlAccName FROM d_Outsales AS d INNER JOIN GLSETUP AS m ON d.Accno = m.AccNo"
    
    Set rs2 = CreateObject("adodb.recordset")
    
    Set clsClass = New cdbase
    
    Provider = clsClass.OpenCon
    
    Set cn = CreateObject("adodb.connection")
    
   cn.Open Provider, "atm", "atm"
    
    rs2.Open sql, cn
    
    With ListView2
        
        .ColumnHeaders.Add , , "Customer Code"
        .ColumnHeaders.Add , , "Customer Name"
        .ColumnHeaders.Add , , "Customer Ledger"
        .ColumnHeaders.Add , , "Ledger Name"
        
        While Not rs2.EOF
        
            Set li = .ListItems.Add(, , Trim(rs2.Fields("Code")))
            
            li.ListSubItems.Add , , Trim(rs2.Fields("Name"))
            li.ListSubItems.Add , , Trim(rs2.Fields("Accno"))
            li.ListSubItems.Add , , Trim(rs2.Fields("GlAccName"))
            
            rs2.MoveNext
        
        Wend
        
    End With
    
    rs2.Close
    
    Set rs2 = Nothing
    
ListView2.View = lvwReport

End Sub


Private Sub cmd5_Click()
reportname = "MonthlyDebtors Report1.rpt"
Show_Sales_Crystal_Report "", reportname, ""
End Sub

Private Sub cmdclose_Click(Index As Integer)
Unload Me
End Sub

Private Sub cmdreports_Click()

End Sub

Private Sub cmddelete_Click()
On Error GoTo ErrorHandler
       '"delete from ag_products where p_code='" & txtpcode & "'"
       sql = ""
       sql = "delete from d_Outsalesb where Code ='" & txtCustName.Text & "' and Name='" & cboNamecust.Text & "' and Date='" & txtdateenterered.value & "' and Description='" & a & "'"
       cn.Execute sql
       sql = ""
       sql = "delete from gltransactions where transdate ='" & txtdateenterered.value & "' and amount=" & txtamount & " and documentno='" & cboNamecust.Text & "'and transdescript='SALES ON-- " & a & "'"
       oSaccoMaster.ExecuteThis (sql)
       MsgBox "Record deleted succesfully", vbInformation
       loadoutsale
       Exit Sub
ErrorHandler:
MsgBox err.description
End Sub

Private Sub cmdindi_Click()
reportname = "Planindividual.rpt"
Show_Sales_Crystal_Report "", reportname, ""
End Sub

Private Sub cmdmonthly_Click()
reportname = "Plan sales Report1.rpt"
Show_Sales_Crystal_Report "", reportname, ""
End Sub

Private Sub cmdnextitem_Click()
On Error GoTo ErrorHandler
If lbldracc = "" Then
  MsgBox "Reselect the product", vbInformation
Exit Sub
End If
If txtquantity = "" Then
 MsgBox "Quantity needed", vbInformation
Exit Sub
End If
If txtprice = "" Then
 MsgBox "Price needed", vbInformation
Exit Sub
End If

If chkBoil = 0 And chksales = 0 Then
 MsgBox "Please select if its for Boil or Sales", vbInformation
Exit Sub
End If

Dim cash As Integer
Dim total As Double
Dim j, Coun As Integer
j = 1




'Check if same item is in the list
   Do While Not j > (Coun)
         Lvwitems.ListItems.Item(j).selected = True
            
    If Lvwitems.SelectedItem = txtCustName Then
        txtquantity = (CCur(txtquantity) + CCur(Lvwitems.SelectedItem.ListSubItems(2)))
        Lvwitems.ListItems.Remove (Lvwitems.SelectedItem.Index)
                        
        Set li = Lvwitems.ListItems.Add(, , txtCustName)
                        li.SubItems(1) = cboNamecust & ""
                        li.SubItems(2) = txtquantity & ""
                        li.SubItems(3) = txtprice & ""
                        li.SubItems(4) = CCur(txtprice) * CCur(txtquantity) & ""
                        li.SubItems(5) = a & ""
                        li.SubItems(6) = txtdracc & ""
                        li.SubItems(7) = lblcracc & ""
                        'Total = CCur(Total + li.SubItems(4))
                        TXTTOTAL = total
                                                
        j = Coun + 1
        
        lblbalance = CCur(lblbalance) - CCur(txtquantity)

        cboNamecust = ""
        txtquantity = ""
       ' txtserialno = ""
        cboNamecust.SetFocus
        Exit Sub
         
    
   
'   lvwItems.ListItems.Item(J).selected = True
   End If
   j = j + 1
    Loop
    
     If j > 1 Then
   
    Set li = Lvwitems.ListItems.Add(, , txtCustName)
                        li.SubItems(1) = cboNamecust & ""
                        li.SubItems(2) = txtquantity & ""
                        li.SubItems(3) = txtprice & ""
                        li.SubItems(4) = CCur(txtprice) * (CCur(txtquantity)) & ""
                        li.SubItems(5) = a & ""
                        li.SubItems(6) = txtdracc & ""
                        li.SubItems(7) = lblcracc & ""
                        'Total = CCur(Total + li.SubItems(4))
                        TXTTOTAL = total
                        
        'lblbalance = CCur(lblbalance) - CCur(txtquantity)
        cboNamecust = ""
        txtquantity = ""
        'txtserialno = ""
        cboNamecust.SetFocus
        Exit Sub
    End If
     If Coun = 0 Then
     Set li = Lvwitems.ListItems.Add(, , txtCustName)
                        li.SubItems(1) = cboNamecust & ""
                        li.SubItems(2) = txtquantity & ""
                        li.SubItems(3) = txtprice & ""
                        li.SubItems(4) = CCur(txtprice) * (CCur(txtquantity)) & ""
                        li.SubItems(5) = a & ""
                        li.SubItems(6) = txtdracc & ""
                        li.SubItems(7) = lblcracc & ""
                        'Total = CCur(Total + li.SubItems(4))
                        TXTTOTAL = total
    End If

'lblbalance = CCur(lblbalance) - CCur(txtquantity)
TXTTOTAL = 0
'Coun = Lvwitems.ListItems.Count
'For j = 1 To Lvwitems.ListItems.Count
'    Total = CCur(Total + li.SubItems(4))
'    txttotal = Total
'
'Next j
Do While Not j > (Lvwitems.ListItems.Count)
'For j = 1 To Lvwitems.ListItems.Count
 Lvwitems.ListItems.Item(j).selected = True
 total = total + CCur(Lvwitems.SelectedItem.SubItems(4))
 TXTTOTAL = total
j = j + 1
Loop
chksales = 0
chkBoil = 0
txtquantity = ""
txtprice = ""
Exit Sub
ErrorHandler:
MsgBox err.description
End Sub


Private Sub cmdsave_Click()
On Error GoTo ErrorHandler
If lbldracc = "" Then
  MsgBox "Reselect the product", vbInformation
Exit Sub
End If
'**********************if to pay for second time
If chkRepay = 0 Then

'   If Label4 = "" Then
'    MsgBox "Please fill all required data", vbInformation
'   Exit Sub
'   End If

   If txtamount = "" Then
      MsgBox "Amount paid needed", vbInformation
   Exit Sub
   End If
   
  Dim j As Integer
   If Lvwitems.ListItems.Count = 0 Then
     MsgBox "There are no items sold."
   Exit Sub
   End If
   j = 1
   
   Dim total, bam As Currency
   total = 0
   Do While Not j > (Lvwitems.ListItems.Count)
     Lvwitems.ListItems.Item(j).selected = True
     total = total + CCur(Lvwitems.SelectedItem.SubItems(4))
     j = j + 1
   Loop
      '*************************** check if amount paid is less expected
    If TXTCHANGE < 0 Then
        'If MsgBox("Insufficient Amount Received,do you want to transfer balance to check off ", vbYesNo) = vbYes Then
'            lblCheckOff_Click
'            lblCheckOff.value = True
'            optCash.value = False
'           Exit Sub
         bam = TXTCHANGE / (j - 1)
        Else
         '  Exit Sub
         'End If
          bam = TXTCHANGE / (j - 1)
    End If
   
   

'// check if they are in stock.
 For j = 1 To Lvwitems.ListItems.Count
 Lvwitems.ListItems.Item(j).selected = True

   Provider = "MAZIWA"
   Set cn = New ADODB.Connection
  cn.Open Provider, "atm", "atm"
   'If reportpath = "" Then reportpath = GetSetting("payroll", "AppName", "rptPath", rptPath)
   sql = "set dateformat dmy select Code,Name,Quantity,Date,Price,Amount,APaid,Description,Owner from d_Outsalesb where Code='" & txtCustName & "' AND Date='" & txtdateenterered.value & "' and Description='" & a & "'"
   Set rs = New ADODB.Recordset
   rs.Open sql, cn
   If Not rs.EOF Then
      If MsgBox("You have already insert the records for this Customer " & Lvwitems.SelectedItem.SubItems(1) & ", Do you want to continue receivinmg? ", vbYesNo) = vbYes Then
      Else
       Exit Sub
      End If
    End If
   '// insert into ag_products
    If TXTCHANGE < 0 Then
      If txtamount = 0 Then
        sql = ""
        sql = "set dateformat dmy insert into  d_Outsalesb(Code,Name,Date,Quantity,Price,Amount,APaid,Description,Owner)"
        sql = sql & "  values('" & Lvwitems.SelectedItem & "','" & Lvwitems.SelectedItem.SubItems(1) & "','" & txtdateenterered.value & "','" & Lvwitems.SelectedItem.SubItems(2) & "'," & Lvwitems.SelectedItem.SubItems(3) & "," & Lvwitems.SelectedItem.SubItems(4) & ",'0','" & Lvwitems.SelectedItem.SubItems(5) & "','" & Lvwitems.SelectedItem.SubItems(6) & "')"
         cn.Execute sql
'        sql = ""
'        sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & txtdateenterered.value & "'," & Lvwitems.SelectedItem.SubItems(4) + bam & ",'" & lbldracc & "','" & lblcracc & "','" & Lvwitems.SelectedItem.SubItems(1) & "','" & Lvwitems.SelectedItem.SubItems(1) & "' ,'SALES ON-- " & a & "','" & User & "','1','0')"
'        oSaccoMaster.ExecuteThis (sql)
       Else
        sql = ""
        sql = "set dateformat dmy insert into  d_Outsalesb(Code,Name,Date,Quantity,Price,Amount,APaid,Description,Owner)"
        sql = sql & "  values('" & Lvwitems.SelectedItem & "','" & Lvwitems.SelectedItem.SubItems(1) & "','" & txtdateenterered.value & "','" & Lvwitems.SelectedItem.SubItems(2) & "'," & Lvwitems.SelectedItem.SubItems(3) & "," & Lvwitems.SelectedItem.SubItems(4) & "," & Lvwitems.SelectedItem.SubItems(4) + bam & ",'" & Lvwitems.SelectedItem.SubItems(5) & "','" & Lvwitems.SelectedItem.SubItems(6) & "')"
         cn.Execute sql
        sql = ""
        sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & txtdateenterered.value & "'," & Lvwitems.SelectedItem.SubItems(4) + bam & ",'" & lbldracc & "','" & Lvwitems.SelectedItem.SubItems(7) & "','" & Lvwitems.SelectedItem.SubItems(1) & "','" & Lvwitems.SelectedItem.SubItems(1) & "' ,'SALES ON-- " & a & "','" & User & "','1','0')"
        oSaccoMaster.ExecuteThis (sql)
      End If
     Else
       sql = ""
       sql = "set dateformat dmy insert into  d_Outsalesb(Code,Name,Date,Quantity,Price,Amount,APaid,Description,Owner)"
       sql = sql & "  values('" & Lvwitems.SelectedItem & "','" & Lvwitems.SelectedItem.SubItems(1) & "','" & txtdateenterered.value & "','" & Lvwitems.SelectedItem.SubItems(2) & "'," & Lvwitems.SelectedItem.SubItems(3) & "," & Lvwitems.SelectedItem.SubItems(4) & "," & Lvwitems.SelectedItem.SubItems(4) + bam & ",'" & Lvwitems.SelectedItem.SubItems(5) & "','" & Lvwitems.SelectedItem.SubItems(6) & "')"
        cn.Execute sql
       sql = ""
       sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & txtdateenterered.value & "'," & Lvwitems.SelectedItem.SubItems(4) + bam & ",'" & lbldracc & "','" & Lvwitems.SelectedItem.SubItems(7) & "','" & Lvwitems.SelectedItem.SubItems(1) & "','" & Lvwitems.SelectedItem.SubItems(1) & "' ,'SALES ON-- " & Lvwitems.SelectedItem.SubItems(5) & "','" & User & "','1','0')"
       oSaccoMaster.ExecuteThis (sql)
     End If
       
    
   Next j
     
    

Else
          sql = ""
          sql = "set dateformat dmy select Code,Name,Quantity,Date,Price,Amount,APaid,Description,Owner from d_Outsalesb where Code='" & txtCustName & "'"
          Set rs = oSaccoMaster.GetRecordset(sql)
       sql = ""
       sql = "set dateformat dmy insert into  d_Outsalesb(Code,Name,Date,Quantity,Price,Amount,APaid,Description,Owner)"
       sql = sql & "  values('" & txtCustName.Text & "','" & cboNamecust.Text & "','" & txtdateenterered.value & "','0','0','0'," & txtamount & ",'MILK PAYMENT','" & rs.Fields(8) & "')"
        cn.Execute sql
       sql = ""
       sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & txtdateenterered.value & "'," & txtamount & ",'" & lbldracc & "','" & lblvracc & "','" & cboNamecust.Text & "','" & cboNamecust.Text & "' ,'SALES ON-- " & rs.Fields(1) & "','" & User & "','1','0')"
       oSaccoMaster.ExecuteThis (sql)
     

End If


MsgBox "Records successively updated."
chkRepay.value = vbUnchecked
chksales.value = vbUnchecked
chkBoil.value = vbUnchecked
chklocal.value = vbUnchecked
Label4 = ""
lblcracc = ""
txtcracc = ""
txtamount.Text = ""
txtprice.Text = ""
txtquantity.Text = ""
cboNamecust.Text = ""
TXTCHANGE.Text = ""
TXTTOTAL.Text = ""
txtCustName.Text = ""
Lvwitems.ListItems.Clear
loadoutsale

Exit Sub
ErrorHandler:
MsgBox err.description


'********************

End Sub

Private Sub cmdSearch_Click()
frmSearchGLAccounts.Show vbModal, Me
    If Continue Then
        If SearchValue <> "" Then
            txtDrAccNo = SearchValue
            SearchValue = ""
        End If
    End If
End Sub

Private Sub Command1_Click()
txtcust = ""
cboName1 = ""
txtDrAccNo = ""
lblDrAccName = ""
txtcust.Locked = False
cboName1.Locked = False
Command1.Enabled = False
Command2.Enabled = True
sql = ""
sql = "select count(Code) from d_Outsales"
Set rs = oSaccoMaster.GetRecordset(sql)

If Not rs.EOF Then
txtcust = rs.Fields(0) + 1
Else
txtcust = 1
End If
loadBranchesTypes
End Sub

Private Sub Command2_Click()
On Error GoTo ErrorHandler
If cboName1 = "" Then
MsgBox "Enter the Branch Code", vbInformation
Exit Sub
End If
Set cn = New ADODB.Connection
sql = "d_sp_Outsales '" & txtcust & "','" & cboName1 & "','" & txtDrAccNo & "'"
oSaccoMaster.ExecuteThis (sql)
txtcust = ""
cboName1 = ""
txtDrAccNo = ""
lblDrAccName = ""
txtcust.Locked = True
cboName1.Locked = True
Command1.Enabled = True
'cmdEdit.Enabled = False
'cmdsave.Enabled = True
loadBranchesTypes
SSTab1_DblClick
MsgBox "Records successively updated."
Exit Sub
ErrorHandler:
MsgBox err.description
End Sub

Private Sub Command3_Click()
reportname = "Plan sales Report.rpt"
Show_Sales_Crystal_Report "", reportname, ""
End Sub

Private Sub Command4_Click()
reportname = "Plan list Report.rpt"
Show_Sales_Crystal_Report "", reportname, ""
End Sub

Private Sub Form_Load()
    
'    txtdateenterered = Format(Get_Server_Date, "dd/mm/yyyy")
'    txtdateenterered.MaxDate = Format(Get_Server_Date, "dd/mm/yyyy")
    txtdateenterered = Format(Get_Server_Date, "dd/mm/yyyy")
    txtdateenterered = Format(Get_Server_Date, "dd/mm/yyyy")
    
    
    Set rst = New Recordset
    Dim cn As Connection
    Set cn = New ADODB.Connection
    Provider = "MAZIWA"
    'Provider = cnn
    cn.Open Provider, "atm", "atm"
    Set rst = New Recordset
    sql = "Select distinct(Name) from   d_Outsales"
    rst.Open sql, cn, adOpenKeyset, adLockOptimistic
    While Not rst.EOF
    cboNamecust.AddItem rst.Fields(0)
    rst.MoveNext
    Wend



loadBranchesTypes


'    '/// to list view//////////
'sql = ""
'sql = "set dateformat dmy SELECT Branch, Quantity,Actual, Variance From d_MilkBranch where Date ='" & DTPMilkDate & "'"
''sql = "set dateformat dmy SELECT d.DCode, d.DName, m.DispQnty,m.DispDate FROM  d_Debtors AS d INNER JOIN d_MilkControl AS m ON d.DCode = m.dcode WHERE     (DispDate = '" & DTPDispatchDate & "') and status=0"
'Set rs = oSaccoMaster.GetRecordset(sql)
'If rs.EOF Then
'Exit Sub
'End If
'ListView3.ListItems.Clear
'While Not rs.EOF
'If Not IsNull(rs.Fields(0)) Then
'Set li = ListView3.ListItems.Add(, , rs.Fields(0))
'End If
'                    If Not IsNull(rs.Fields(1)) Then li.SubItems(1) = rs.Fields(1) & ""
'                    If Not IsNull(rs.Fields(2)) Then li.SubItems(2) = rs.Fields(2) & ""
'                    If Not IsNull(rs.Fields(3)) Then li.SubItems(3) = rs.Fields(3) & ""
'                  '  If Not IsNull(rs.Fields(4)) Then li.SubItems(4) = rs.Fields(4) & ""
''                    If Not IsNull(rs.Fields(5)) Then li.SubItems(5) = rs.Fields(5) & ""
'rs.MoveNext
'
'Wend
'
''////// end of view
loadoutsale

End Sub

Public Sub loadoutsale()
    

    
    
    With ListView3
    
        .ListItems.Clear
        
        .ColumnHeaders.Clear
    
    End With

    Set rs2 = CreateObject("adodb.recordset")
    
    sql = "set dateformat dmy Select * from d_Outsalesb where Date='" & txtdateenterered.value & "'"
    
    Set rs2 = CreateObject("adodb.recordset")
    
    Set clsClass = New cdbase
    
    Provider = clsClass.OpenCon
    
    Set cn = CreateObject("adodb.connection")
    
   cn.Open Provider, "atm", "atm"
    
    rs2.Open sql, cn
    
    With ListView3
        
        .ColumnHeaders.Add , , "Code"
        .ColumnHeaders.Add , , "Name"
        .ColumnHeaders.Add , , "Quantity"
        .ColumnHeaders.Add , , "Price"
        .ColumnHeaders.Add , , "Amount"
        .ColumnHeaders.Add , , "Paid"
        .ColumnHeaders.Add , , "Description"
        While Not rs2.EOF
        'Code, Name, Date, Quantity, Price, Amount, APaid
            Set li = .ListItems.Add(, , Trim(rs2.Fields("Code")))
            li.ListSubItems.Add , , Trim(rs2.Fields("Name"))
            li.ListSubItems.Add , , Trim(rs2.Fields("Quantity"))
            li.ListSubItems.Add , , Trim(rs2.Fields("Price"))
            li.ListSubItems.Add , , Trim(rs2.Fields("Amount"))
            li.ListSubItems.Add , , Trim(rs2.Fields("APaid"))
            li.ListSubItems.Add , , Trim(rs2.Fields("Description"))
            rs2.MoveNext
        
        Wend
        
    End With
    
    rs2.Close
    
    Set rs2 = Nothing
    
ListView3.View = lvwReport

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

Private Sub ListView2_DblClick()
txtcust = ListView2.SelectedItem
cboName1 = ListView2.SelectedItem.SubItems(1)
End Sub

Private Sub ListView3_DblClick()
chksales = 0
chkBoil = 0
txtquantity = ListView3.SelectedItem.SubItems(2)
txtprice = ListView3.SelectedItem.SubItems(3)
txtamount = ListView3.SelectedItem.SubItems(4)
If ListView3.SelectedItem.SubItems(6) = "Boiling fee" Then
 chkBoil = 1
Else
 chksales = 1
End If
cboNamecust = ListView3.SelectedItem.SubItems(1)
cboNamecust_Click
End Sub

Private Sub Picture1_Click()
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

Private Sub SSTab1_DblClick()
    cboNamecust.Clear
    Set rst = New Recordset
    Dim cn As Connection
    Set cn = New ADODB.Connection
    Provider = "MAZIWA"
    'Provider = cnn
    cn.Open Provider, "atm", "atm"
    Set rst = New Recordset
    sql = "Select distinct(Name) from   d_Outsales"
    rst.Open sql, cn, adOpenKeyset, adLockOptimistic
    While Not rst.EOF
    cboNamecust.AddItem rst.Fields(0)
    rst.MoveNext
    Wend

End Sub
Private Sub cmdRemove_Click()
On Error GoTo ErrorHandler
Dim total As Double
Dim j, Coun As Integer
j = 1
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

Private Sub txtAmount_Change()
On Error Resume Next
TXTCHANGE = txtamount - TXTTOTAL
End Sub

Private Sub txtdateenterered_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
loadoutsale
End Sub

Private Sub txtDrAccNo_Change()
    On Error GoTo SysError
    Dim Account As Acc_Details
    Editing = True
    Account = Get_Acc_Details(txtDrAccNo, ErrorMessage)
    If Account.ACCNO <> "" Then
        lblDrAccName = Account.AccName
    Else
        If ErrorMessage <> "" Then
            MsgBox ErrorMessage, vbInformation, Me.Caption
            ErrorMessage = ""
        End If
        lblDrAccName = ""
    End If
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub txtprice_Change()
'cmdnextitem.SetFocus
End Sub

Private Sub txtquantity_Change()
'txtPrice.SetFocus
End Sub

Private Sub txttotal_Change()
On Error Resume Next
TXTCHANGE = txtamount - TXTTOTAL
End Sub
