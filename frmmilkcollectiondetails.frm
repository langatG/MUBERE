VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmmilkcollectiondetails 
   Caption         =   "Milk Collection Details"
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7080
   LinkTopic       =   "Form1"
   ScaleHeight     =   7005
   ScaleWidth      =   7080
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1320
      TabIndex        =   19
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2760
      TabIndex        =   18
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   375
      Left            =   1440
      TabIndex        =   17
      Top             =   5880
      Width           =   855
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1320
      TabIndex        =   13
      Top             =   5400
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   60096513
      CurrentDate     =   40112
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1320
      TabIndex        =   12
      Top             =   4200
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1320
      TabIndex        =   11
      Top             =   3600
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   1320
      TabIndex        =   10
      Top             =   3000
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1320
      TabIndex        =   8
      Text            =   "Combo1"
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1320
      TabIndex        =   6
      Top             =   1680
      Width           =   2655
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Print"
      Height          =   255
      Left            =   5760
      TabIndex        =   4
      Top             =   960
      Width           =   1335
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Quick Mode"
      Height          =   255
      Left            =   4080
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   840
      Width           =   2655
   End
   Begin VB.Label Label9 
      Caption         =   "Date Delivered"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   5520
      Width           =   975
   End
   Begin VB.Label Label8 
      Caption         =   "Qty Delivered"
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "Transport Cost/Lir"
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Price/Lir"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Transport"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Route"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Member"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Member Code"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Collection Details and Save"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "frmmilkcollectiondetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
