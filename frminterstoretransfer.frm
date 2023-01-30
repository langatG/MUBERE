VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frminterstoretransfer 
   Caption         =   "Interstore Transfer"
   ClientHeight    =   8250
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   ScaleHeight     =   8250
   ScaleWidth      =   6735
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Close"
      Height          =   375
      Left            =   2880
      TabIndex        =   22
      Top             =   7560
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Save"
      Height          =   375
      Left            =   1560
      TabIndex        =   21
      Top             =   7560
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Clear"
      Height          =   375
      Left            =   240
      TabIndex        =   20
      Top             =   7560
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Remove"
      Height          =   375
      Left            =   4560
      TabIndex        =   19
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   375
      Left            =   4560
      TabIndex        =   18
      Top             =   3720
      Width           =   1095
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2535
      Left            =   120
      TabIndex        =   17
      Top             =   4920
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   4471
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Item"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Quantity"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2160
      TabIndex        =   16
      Top             =   4320
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2160
      TabIndex        =   14
      Top             =   3720
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2160
      TabIndex        =   12
      Top             =   3120
      Width           =   2175
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   2160
      TabIndex        =   8
      Text            =   "Combo3"
      Top             =   2400
      Width           =   3255
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   2160
      TabIndex        =   7
      Text            =   "Combo2"
      Top             =   1920
      Width           =   3255
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2160
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   1440
      Width           =   3255
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   960
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      _Version        =   393216
      Format          =   103809025
      CurrentDate     =   40112
   End
   Begin VB.Label Label10 
      Caption         =   "Quantity"
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "Balance"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Inventory Item"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "ITEMS"
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
      TabIndex        =   10
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "TO"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "From"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Employee"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Issue Date"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Transfer Information"
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
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Transfer Items To Other Stores"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "frminterstoretransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
