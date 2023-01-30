VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmvendorscreditdebitmemos 
   Caption         =   "Vendor Credit/Debit Momos"
   ClientHeight    =   8850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   ScaleHeight     =   8850
   ScaleWidth      =   6465
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   1800
      TabIndex        =   25
      Top             =   7080
      Width           =   2895
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   1800
      TabIndex        =   23
      Top             =   6480
      Width           =   2895
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      Left            =   1800
      TabIndex        =   21
      Text            =   "Combo4"
      Top             =   6000
      Width           =   2895
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1800
      TabIndex        =   19
      Top             =   5400
      Width           =   2895
   End
   Begin VB.CheckBox Check1 
      Caption         =   "VAT INCLUSIVE"
      Height          =   375
      Left            =   1800
      TabIndex        =   17
      Top             =   4920
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1800
      TabIndex        =   15
      Top             =   4320
      Width           =   3015
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   1800
      TabIndex        =   13
      Text            =   "Combo3"
      Top             =   3840
      Width           =   3015
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1920
      TabIndex        =   10
      Top             =   2760
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      _Version        =   393216
      Format          =   65142785
      CurrentDate     =   40112
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1920
      TabIndex        =   9
      Top             =   2280
      Width           =   3015
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1920
      TabIndex        =   8
      Text            =   "Combo2"
      Top             =   1800
      Width           =   3015
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1920
      TabIndex        =   7
      Text            =   "Combo1"
      Top             =   1320
      Width           =   3015
   End
   Begin VB.Label Label15 
      Caption         =   "Total"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Label Label14 
      Caption         =   "VAT"
      Height          =   375
      Left            =   120
      TabIndex        =   22
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Label Label13 
      Caption         =   "VAT Type"
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Label Label12 
      Caption         =   "Amount"
      Height          =   495
      Left            =   120
      TabIndex        =   18
      Top             =   5520
      Width           =   1335
   End
   Begin VB.Label Label11 
      Caption         =   "Amount Details"
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
      TabIndex        =   16
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Label Label10 
      Caption         =   "Description"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Label Label9 
      Caption         =   "Type"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label Label8 
      Caption         =   "Memo Details"
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
      TabIndex        =   11
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "Date"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Invoice Amount"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Invoice"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Vendor"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Invoice Details"
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
      TabIndex        =   2
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Create Vendor Credit/Debit Memos"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Vendor Credit/Debit Memos"
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
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "frmvendorscreditdebitmemos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
