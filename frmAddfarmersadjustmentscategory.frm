VERSION 5.00
Begin VB.Form frmAddfarmersadjustmentscategory 
   Caption         =   "Farmers Adjustment Category"
   ClientHeight    =   5010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   ScaleHeight     =   5010
   ScaleWidth      =   5295
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   375
      Left            =   2760
      TabIndex        =   13
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   375
      Left            =   1560
      TabIndex        =   12
      Top             =   4200
      Width           =   975
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   2040
      TabIndex        =   11
      Text            =   "Combo3"
      Top             =   3480
      Width           =   2295
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   2040
      TabIndex        =   9
      Text            =   "Combo2"
      Top             =   2880
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      Top             =   2280
      Width           =   2295
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2040
      TabIndex        =   7
      Text            =   "Combo1"
      Top             =   1800
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Priority"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "GL Account"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Type"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Description"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "General"
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
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Label1Farmers Adjustment Category"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmAddfarmersadjustmentscategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
