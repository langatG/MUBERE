VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmmilkdispatch 
   Caption         =   "Milk Dispatch"
   ClientHeight    =   9045
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7080
   LinkTopic       =   "Form1"
   ScaleHeight     =   9045
   ScaleWidth      =   7080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Remove"
      Height          =   375
      Left            =   5160
      TabIndex        =   33
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Add"
      Height          =   375
      Left            =   3960
      TabIndex        =   32
      Top             =   5520
      Width           =   975
   End
   Begin VB.TextBox Text10 
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   1440
      TabIndex        =   31
      Text            =   "Text10"
      Top             =   600
      Width           =   1815
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   4560
      TabIndex        =   29
      Top             =   2880
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   122093569
      CurrentDate     =   40114
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   4560
      TabIndex        =   27
      Top             =   2280
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   122093569
      CurrentDate     =   40114
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   375
      Left            =   5040
      TabIndex        =   25
      Top             =   8520
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   375
      Left            =   3720
      TabIndex        =   24
      Top             =   8520
      Width           =   975
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2175
      Left            =   360
      TabIndex        =   23
      Top             =   6000
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   3836
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Recid"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Def No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Quantity"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Temperature"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Lactometer"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Resazuring"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Butter Fat"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox Text9 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   5040
      TabIndex        =   22
      Top             =   4920
      Width           =   1575
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   5040
      TabIndex        =   21
      Top             =   4320
      Width           =   1575
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1440
      TabIndex        =   20
      Top             =   5520
      Width           =   1575
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1440
      TabIndex        =   19
      Top             =   4920
      Width           =   1575
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1440
      TabIndex        =   18
      Top             =   4320
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1440
      TabIndex        =   11
      Top             =   3360
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1440
      TabIndex        =   10
      Top             =   2760
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1440
      TabIndex        =   9
      Top             =   2280
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1440
      TabIndex        =   8
      Top             =   1680
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1440
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   1080
      Width           =   4215
   End
   Begin VB.Label Label16 
      Caption         =   "Dispatch Ref"
      Height          =   375
      Left            =   120
      TabIndex        =   30
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label15 
      Caption         =   "Dispatch Time"
      Height          =   375
      Left            =   3480
      TabIndex        =   28
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Date"
      Height          =   375
      Left            =   3600
      TabIndex        =   26
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label Label14 
      Caption         =   "Resazuring"
      Height          =   255
      Left            =   3720
      TabIndex        =   17
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Label13 
      Caption         =   "Temperature"
      Height          =   255
      Left            =   3720
      TabIndex        =   16
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label Label12 
      Caption         =   "Butter Fat"
      Height          =   495
      Left            =   120
      TabIndex        =   15
      Top             =   5520
      Width           =   1455
   End
   Begin VB.Label Label11 
      Caption         =   "Lactometer"
      Height          =   495
      Left            =   120
      TabIndex        =   14
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label Label10 
      Caption         =   "Qty"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Label Label9 
      Caption         =   "Dispatch Details"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "Vihicle Reg No"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "Recieved By"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Dispatch By"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Dispatch Qty"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Customer"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Dispatch Header"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Dispatch Details And Save"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmmilkdispatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
