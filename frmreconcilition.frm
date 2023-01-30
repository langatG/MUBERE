VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmreconcilition 
   Caption         =   "Reconciliation"
   ClientHeight    =   8880
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10215
   LinkTopic       =   "Form1"
   ScaleHeight     =   8880
   ScaleWidth      =   10215
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command9 
      Caption         =   "Close"
      Height          =   375
      Left            =   8640
      TabIndex        =   32
      Top             =   8400
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Reconcile"
      Height          =   375
      Left            =   7440
      TabIndex        =   31
      Top             =   8400
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Resume"
      Height          =   375
      Left            =   1680
      TabIndex        =   30
      Top             =   8400
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Pause"
      Height          =   375
      Left            =   480
      TabIndex        =   29
      Top             =   8400
      Width           =   975
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7455
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   13150
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Transaction"
      TabPicture(0)   =   "frmreconcilition.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Text5"
      Tab(0).Control(1)=   "Text4"
      Tab(0).Control(2)=   "Command5"
      Tab(0).Control(3)=   "Command4"
      Tab(0).Control(4)=   "Command3"
      Tab(0).Control(5)=   "Command2"
      Tab(0).Control(6)=   "ListView4"
      Tab(0).Control(7)=   "ListView3"
      Tab(0).Control(8)=   "ListView2"
      Tab(0).Control(9)=   "ListView1"
      Tab(0).Control(10)=   "Command1"
      Tab(0).Control(11)=   "Combo2"
      Tab(0).Control(12)=   "Text3"
      Tab(0).Control(13)=   "Text2"
      Tab(0).Control(14)=   "Text1"
      Tab(0).Control(15)=   "Combo1"
      Tab(0).Control(16)=   "DTPicker1"
      Tab(0).Control(17)=   "Label13"
      Tab(0).Control(18)=   "Label12"
      Tab(0).Control(19)=   "Label11"
      Tab(0).Control(20)=   "Label10"
      Tab(0).Control(21)=   "Label9"
      Tab(0).Control(22)=   "Label8"
      Tab(0).Control(23)=   "Label7"
      Tab(0).Control(24)=   "Label6"
      Tab(0).Control(25)=   "Label5"
      Tab(0).Control(26)=   "Label4"
      Tab(0).Control(27)=   "Label3"
      Tab(0).Control(28)=   "Label2"
      Tab(0).ControlCount=   29
      TabCaption(1)   =   "Adjustments"
      TabPicture(1)   =   "frmreconcilition.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label14"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label15"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label16"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label17"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Text6"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Text7"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Combo3"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Text8"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Command10"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Command11"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "ListView5"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).ControlCount=   11
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "frmreconcilition.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin MSComctlLib.ListView ListView5 
         Height          =   3255
         Left            =   240
         TabIndex        =   45
         Top             =   3000
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   5741
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Cheque#"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Credit"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Debit"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Remove"
         Height          =   375
         Left            =   3840
         TabIndex        =   44
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Add"
         Height          =   375
         Left            =   3840
         TabIndex        =   43
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1680
         TabIndex        =   42
         Top             =   2400
         Width           =   1935
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   1680
         TabIndex        =   40
         Text            =   "Combo3"
         Top             =   1800
         Width           =   1935
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1680
         TabIndex        =   38
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1680
         TabIndex        =   37
         Top             =   600
         Width           =   3735
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   -67800
         TabIndex        =   26
         Top             =   6360
         Width           =   1815
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   -73080
         TabIndex        =   24
         Top             =   6360
         Width           =   1575
      End
      Begin VB.CommandButton Command5 
         Caption         =   "<<"
         Height          =   375
         Left            =   -70800
         TabIndex        =   22
         Top             =   5400
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   ">>"
         Height          =   375
         Left            =   -70800
         TabIndex        =   21
         Top             =   4800
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "<<"
         Height          =   375
         Left            =   -70800
         TabIndex        =   20
         Top             =   3360
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   ">>"
         Height          =   375
         Left            =   -70800
         TabIndex        =   19
         Top             =   2760
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView4 
         Height          =   1335
         Left            =   -69240
         TabIndex        =   18
         Top             =   4800
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   2355
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Cheque#"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Amount"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   1335
         Left            =   -74640
         TabIndex        =   17
         Top             =   4800
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   2355
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Cheque#"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Amount"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   1215
         Left            =   -69240
         TabIndex        =   16
         Top             =   2760
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   2143
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Cheque#"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Amount"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1335
         Left            =   -74640
         TabIndex        =   15
         Top             =   2760
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   2355
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Cheque#"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Amount"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Go"
         Height          =   375
         Left            =   -69120
         TabIndex        =   14
         Top             =   2040
         Width           =   495
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   -70800
         TabIndex        =   13
         Text            =   "Combo2"
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   -72840
         TabIndex        =   12
         Top             =   2040
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   -68520
         TabIndex        =   9
         Top             =   1560
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   -68520
         TabIndex        =   8
         Top             =   1080
         Width           =   1935
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   -72840
         TabIndex        =   3
         Text            =   "Combo1"
         Top             =   960
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   255
         Left            =   -72840
         TabIndex        =   5
         Top             =   1560
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         _Version        =   393216
         Format          =   60030977
         CurrentDate     =   40119
      End
      Begin VB.Label Label17 
         Caption         =   "Label17"
         Height          =   375
         Left            =   120
         TabIndex        =   41
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label16 
         Caption         =   "Label16"
         Height          =   375
         Left            =   120
         TabIndex        =   39
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label15 
         Caption         =   "Label15"
         Height          =   375
         Left            =   120
         TabIndex        =   36
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label14 
         Caption         =   "Label14"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label13 
         Caption         =   "Deposits/Other Credits to be Reconcile"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -69240
         TabIndex        =   34
         Top             =   4560
         Width           =   3495
      End
      Begin VB.Label Label12 
         Caption         =   "Deposits/Other Credits"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74640
         TabIndex        =   33
         Top             =   4560
         Width           =   2055
      End
      Begin VB.Label Label11 
         Caption         =   "Cheque/Payments Reconciled"
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
         Left            =   -69120
         TabIndex        =   28
         Top             =   2520
         Width           =   2655
      End
      Begin VB.Label Label10 
         Caption         =   "Cheque/Payments"
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
         Left            =   -74640
         TabIndex        =   27
         Top             =   2520
         Width           =   2175
      End
      Begin VB.Label Label9 
         Caption         =   "Difference"
         Height          =   375
         Left            =   -69360
         TabIndex        =   25
         Top             =   6360
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Closing Balance"
         Height          =   375
         Left            =   -74640
         TabIndex        =   23
         Top             =   6360
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Search Cheque"
         Height          =   375
         Left            =   -74760
         TabIndex        =   11
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Select Account"
         Height          =   375
         Left            =   -74760
         TabIndex        =   10
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Statement Balance"
         Height          =   255
         Left            =   -70080
         TabIndex        =   7
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Opening Balance"
         Height          =   255
         Left            =   -70080
         TabIndex        =   6
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Date"
         Height          =   375
         Left            =   -74760
         TabIndex        =   4
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Cash Account"
         Height          =   375
         Left            =   -74760
         TabIndex        =   2
         Top             =   960
         Width           =   1575
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Reconcile Bank and Cash Accounts"
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
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   3255
   End
End
Attribute VB_Name = "frmreconcilition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
