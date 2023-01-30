VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRejectedmilk 
   Caption         =   "Rejected Milk"
   ClientHeight    =   6165
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7260
   LinkTopic       =   "Form1"
   ScaleHeight     =   6165
   ScaleWidth      =   7260
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1800
      TabIndex        =   19
      Text            =   "Combo2"
      Top             =   3960
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5040
      TabIndex        =   18
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   375
      Left            =   3720
      TabIndex        =   17
      Top             =   5160
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1800
      TabIndex        =   16
      Top             =   4440
      Width           =   2415
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1800
      TabIndex        =   14
      Top             =   3480
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      Format          =   60162049
      CurrentDate     =   40113
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1800
      TabIndex        =   12
      Top             =   2880
      Width           =   2175
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1800
      TabIndex        =   10
      Text            =   "Combo1"
      Top             =   2400
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1800
      TabIndex        =   8
      Top             =   960
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   1920
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   1440
      Width           =   2175
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label9 
      Caption         =   "Ref No"
      Height          =   375
      Left            =   240
      TabIndex        =   20
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "Comments"
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "Reason"
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "Date"
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Qty Rejected"
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Route"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Member"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Member Code"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Add Rejected Details Add Save"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
End
Attribute VB_Name = "frmRejectedmilk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

