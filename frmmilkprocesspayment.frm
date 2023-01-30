VERSION 5.00
Begin VB.Form frmmilkprocesspayment 
   Caption         =   "Milk Payment Processing"
   ClientHeight    =   4065
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   ScaleHeight     =   4065
   ScaleWidth      =   5805
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Period"
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Process"
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Process Period"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Processing Details"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Milk Payment Processing"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "frmmilkprocesspayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
