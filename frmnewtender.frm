VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmnewtender 
   Caption         =   "Tenders"
   ClientHeight    =   7635
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10035
   LinkTopic       =   "Form1"
   ScaleHeight     =   7635
   ScaleWidth      =   10035
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   375
      Left            =   8280
      TabIndex        =   11
      Top             =   7200
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   375
      Left            =   6720
      TabIndex        =   10
      Top             =   7200
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   9551
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Tender Information"
      TabPicture(0)   =   "frmnewtender.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label6"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Text1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Text3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "DTPicker1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Bonds"
      TabPicture(1)   =   "frmnewtender.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label7"
      Tab(1).Control(1)=   "Label8"
      Tab(1).Control(2)=   "Label9"
      Tab(1).Control(3)=   "Label10"
      Tab(1).Control(4)=   "Label11"
      Tab(1).Control(5)=   "Label12"
      Tab(1).Control(6)=   "Check1"
      Tab(1).Control(7)=   "Text2"
      Tab(1).Control(8)=   "Combo1"
      Tab(1).Control(9)=   "Check2"
      Tab(1).Control(10)=   "Text4"
      Tab(1).Control(11)=   "Combo2"
      Tab(1).ControlCount=   12
      TabCaption(2)   =   "Tender Items"
      TabPicture(2)   =   "frmnewtender.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label13"
      Tab(2).Control(1)=   "Label14"
      Tab(2).Control(2)=   "ListView1"
      Tab(2).Control(3)=   "ListView2"
      Tab(2).Control(4)=   "Command3"
      Tab(2).Control(5)=   "Command4"
      Tab(2).ControlCount=   6
      Begin VB.CommandButton Command4 
         Caption         =   "Remove"
         Height          =   375
         Left            =   -67200
         TabIndex        =   29
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Add"
         Height          =   375
         Left            =   -68400
         TabIndex        =   28
         Top             =   2520
         Width           =   975
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   2055
         Left            =   -74880
         TabIndex        =   25
         Top             =   3000
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   3625
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
            Text            =   "ReqId"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Item"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Cost Center"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Quantity"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1695
         Left            =   -74880
         TabIndex        =   24
         Top             =   720
         Width           =   9255
         _ExtentX        =   16325
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ReqID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "iTEM"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Cost Center"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Item Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Quantity"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   -72720
         TabIndex        =   23
         Text            =   "Combo2"
         Top             =   4440
         Width           =   1575
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   -72720
         TabIndex        =   21
         Top             =   3720
         Width           =   1575
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Performance Bond Required"
         Height          =   495
         Left            =   -74760
         TabIndex        =   19
         Top             =   3120
         Width           =   2415
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   -72720
         TabIndex        =   17
         Text            =   "Combo1"
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -72720
         TabIndex        =   16
         Top             =   1560
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Bid Bond Required"
         Height          =   375
         Left            =   -73920
         TabIndex        =   13
         Top             =   960
         Width           =   2415
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2400
         TabIndex        =   8
         Top             =   1920
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   206569473
         CurrentDate     =   40110
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   1095
         Left            =   2400
         TabIndex        =   7
         Top             =   2640
         Width           =   3975
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2400
         TabIndex        =   4
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label14 
         Caption         =   "Selected Items"
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
         Left            =   -74880
         TabIndex        =   27
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label Label13 
         Caption         =   "Available Item"
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
         Left            =   -74880
         TabIndex        =   26
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label12 
         Caption         =   "Type"
         Height          =   375
         Left            =   -73560
         TabIndex        =   22
         Top             =   4440
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "Amount"
         Height          =   375
         Left            =   -73560
         TabIndex        =   20
         Top             =   3840
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Performance Bond"
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
         Left            =   -74760
         TabIndex        =   18
         Top             =   2760
         Width           =   2775
      End
      Begin VB.Label Label9 
         Caption         =   "Type"
         Height          =   255
         Left            =   -73440
         TabIndex        =   15
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "Amount"
         Height          =   375
         Left            =   -73440
         TabIndex        =   14
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Bid Bond"
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
         Left            =   -74880
         TabIndex        =   12
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "General Information"
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
         Left            =   240
         TabIndex        =   9
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label5 
         Caption         =   "Description"
         Height          =   375
         Left            =   480
         TabIndex        =   6
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Closing Date"
         Height          =   375
         Left            =   480
         TabIndex        =   5
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Tender No"
         Height          =   375
         Left            =   480
         TabIndex        =   3
         Top             =   1200
         Width           =   1695
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Add and Edit Tenders"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Tenders"
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
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "frmnewtender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
