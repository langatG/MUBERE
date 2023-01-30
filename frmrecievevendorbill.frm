VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmrecievevendorbill 
   Caption         =   "Recieve Vendor Bill"
   ClientHeight    =   7230
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7635
   LinkTopic       =   "Form1"
   ScaleHeight     =   7230
   ScaleWidth      =   7635
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   375
      Left            =   5520
      TabIndex        =   3
      Top             =   6720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   6720
      Width           =   855
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6255
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   11033
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmrecievevendorbill.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "DTPicker2"
      Tab(0).Control(1)=   "Text3"
      Tab(0).Control(2)=   "Text2"
      Tab(0).Control(3)=   "DTPicker1"
      Tab(0).Control(4)=   "Combo1"
      Tab(0).Control(5)=   "Text1"
      Tab(0).Control(6)=   "Label8"
      Tab(0).Control(7)=   "Label7"
      Tab(0).Control(8)=   "Label6"
      Tab(0).Control(9)=   "Label5"
      Tab(0).Control(10)=   "Label4"
      Tab(0).Control(11)=   "Label3"
      Tab(0).Control(12)=   "Label2"
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "Others"
      TabPicture(1)   =   "frmrecievevendorbill.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label9"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label10"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label11"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label12"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label13"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label14"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label15"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label16"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label17"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label18"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label19"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Combo2"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Text4"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Text5"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Text6"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Combo3"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Text7"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Combo4"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Text8"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Text9"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).ControlCount=   20
      TabCaption(2)   =   "Bill Amounts"
      TabPicture(2)   =   "frmrecievevendorbill.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label20"
      Tab(2).Control(1)=   "Label21"
      Tab(2).Control(2)=   "Label22"
      Tab(2).Control(3)=   "Text10"
      Tab(2).Control(4)=   "Text11"
      Tab(2).Control(5)=   "Command3"
      Tab(2).Control(6)=   "ListView1"
      Tab(2).ControlCount=   7
      Begin MSComctlLib.ListView ListView1 
         Height          =   3015
         Left            =   -74760
         TabIndex        =   43
         Top             =   2040
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   5318
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
            Text            =   "Account"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Accountname"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Amount"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Add"
         Height          =   375
         Left            =   -70320
         TabIndex        =   42
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox Text11 
         Height          =   375
         Left            =   -72960
         TabIndex        =   41
         Top             =   1560
         Width           =   2415
      End
      Begin VB.TextBox Text10 
         Height          =   375
         Left            =   -72960
         TabIndex        =   39
         Top             =   960
         Width           =   2415
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2400
         TabIndex        =   36
         Top             =   5520
         Width           =   2055
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2400
         TabIndex        =   34
         Top             =   4920
         Width           =   2055
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   2400
         TabIndex        =   32
         Text            =   "Combo4"
         Top             =   4440
         Width           =   2055
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         TabIndex        =   30
         Top             =   3840
         Width           =   2055
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   2400
         TabIndex        =   28
         Text            =   "Combo3"
         Top             =   3360
         Width           =   2055
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2400
         TabIndex        =   25
         Top             =   2520
         Width           =   1935
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2400
         TabIndex        =   23
         Top             =   1920
         Width           =   1935
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2400
         TabIndex        =   21
         Top             =   1440
         Width           =   1935
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   2400
         TabIndex        =   19
         Text            =   "Combo2"
         Top             =   960
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   -72840
         TabIndex        =   16
         Top             =   2760
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         Format          =   20054017
         CurrentDate     =   40112
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -72840
         TabIndex        =   14
         Top             =   1800
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         Height          =   2055
         Left            =   -74880
         TabIndex        =   12
         Top             =   3600
         Width           =   7335
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   -72840
         TabIndex        =   10
         Top             =   2160
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         Format          =   20054017
         CurrentDate     =   40112
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   -72840
         TabIndex        =   8
         Text            =   "Combo1"
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   -72840
         TabIndex        =   6
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label22 
         Caption         =   "Label22"
         Height          =   255
         Left            =   -74760
         TabIndex        =   40
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label21 
         Caption         =   "Label21"
         Height          =   375
         Left            =   -74760
         TabIndex        =   38
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label20 
         Caption         =   "Details"
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
         TabIndex        =   37
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label19 
         Caption         =   "Total Amount"
         Height          =   375
         Left            =   360
         TabIndex        =   35
         Top             =   5400
         Width           =   1335
      End
      Begin VB.Label Label18 
         Caption         =   "VAT"
         Height          =   375
         Left            =   360
         TabIndex        =   33
         Top             =   4920
         Width           =   975
      End
      Begin VB.Label Label17 
         Caption         =   "VAT Type"
         Height          =   495
         Left            =   360
         TabIndex        =   31
         Top             =   4440
         Width           =   975
      End
      Begin VB.Label Label16 
         Caption         =   "Invoice Amount"
         Height          =   375
         Left            =   360
         TabIndex        =   29
         Top             =   3840
         Width           =   1335
      End
      Begin VB.Label Label15 
         Caption         =   "Service"
         Height          =   255
         Left            =   360
         TabIndex        =   27
         Top             =   3480
         Width           =   1455
      End
      Begin VB.Label Label14 
         Caption         =   "Invoice Amount"
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
         TabIndex        =   26
         Top             =   3120
         Width           =   1935
      End
      Begin VB.Label Label13 
         Caption         =   "Account No"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "Back Branch"
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "Back Name"
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Payment Method"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Payment Instruction"
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
         TabIndex        =   17
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label Label8 
         Caption         =   "Invoice Date"
         Height          =   375
         Left            =   -74760
         TabIndex        =   15
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Comments"
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
         TabIndex        =   13
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Invoice Due"
         Height          =   255
         Left            =   -74760
         TabIndex        =   11
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Invoice ID"
         Height          =   375
         Left            =   -74760
         TabIndex        =   9
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Vendor"
         Height          =   375
         Left            =   -74760
         TabIndex        =   7
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Reference"
         Height          =   375
         Left            =   -74760
         TabIndex        =   5
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Invoice Header"
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
         Left            =   -74760
         TabIndex        =   4
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Recieve Vendor Bill"
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
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmrecievevendorbill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
