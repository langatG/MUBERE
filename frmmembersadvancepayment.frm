VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmmembersadvancepayment 
   Caption         =   "members Advance Payment"
   ClientHeight    =   7245
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7785
   LinkTopic       =   "Form2"
   ScaleHeight     =   7245
   ScaleWidth      =   7785
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text9 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2160
      TabIndex        =   25
      Top             =   3960
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   375
      Left            =   3840
      TabIndex        =   24
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   375
      Left            =   2520
      TabIndex        =   23
      Top             =   6240
      Width           =   1095
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2040
      TabIndex        =   22
      Top             =   5760
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2040
      TabIndex        =   21
      Top             =   5160
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2040
      TabIndex        =   20
      Top             =   4560
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   5280
      TabIndex        =   19
      Top             =   5280
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   5280
      TabIndex        =   18
      Top             =   4680
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2160
      TabIndex        =   12
      Top             =   3360
      Width           =   2055
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   2160
      TabIndex        =   10
      Top             =   2880
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      Format          =   20185089
      CurrentDate     =   40116
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   2160
      TabIndex        =   9
      Text            =   "Combo2"
      Top             =   2400
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      Top             =   1800
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2160
      TabIndex        =   7
      Text            =   "Combo1"
      Top             =   1320
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label13 
      Caption         =   "Amount"
      Height          =   375
      Left            =   120
      TabIndex        =   26
      Top             =   5760
      Width           =   1335
   End
   Begin VB.Label Label12 
      Caption         =   "Net Available"
      Height          =   375
      Left            =   3600
      TabIndex        =   17
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   "Gross Pay"
      Height          =   375
      Left            =   3600
      TabIndex        =   16
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "Deeduction"
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Label Label9 
      Caption         =   "Total Kgs"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "Route"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "Date"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Finacial Period"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Member"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Cash Account"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Advance Id"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Member Details and Save"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "frmmembersadvancepayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
