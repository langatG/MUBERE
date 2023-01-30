VERSION 5.00
Begin VB.Form frmnewcashaccount 
   Caption         =   "New Cash Account"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6015
   LinkTopic       =   "Form2"
   ScaleHeight     =   6000
   ScaleWidth      =   6015
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2280
      TabIndex        =   15
      Top             =   3600
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2280
      TabIndex        =   14
      Top             =   2880
      Width           =   2295
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   2280
      TabIndex        =   11
      Text            =   "Combo2"
      Top             =   1680
      Width           =   2295
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2280
      TabIndex        =   10
      Text            =   "Combo1"
      Top             =   1080
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Close"
      Height          =   375
      Left            =   4080
      TabIndex        =   9
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save"
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clear"
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Default Vender Back Payment"
      Height          =   495
      Left            =   2280
      TabIndex        =   6
      Top             =   3960
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2280
      TabIndex        =   5
      Top             =   2280
      Width           =   2295
   End
   Begin VB.TextBox r 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2280
      TabIndex        =   4
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Label8 
      Caption         =   "General Ledger Account"
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
      Top             =   4440
      Width           =   3015
   End
   Begin VB.Label Label7 
      Caption         =   "Account Details"
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
      TabIndex        =   16
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "Account Number"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Branch Name"
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Account Name"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Currency"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Account Type"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Account #"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "frmnewcashaccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
