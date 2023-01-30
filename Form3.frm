VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9180
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11715
   LinkTopic       =   "Form1"
   ScaleHeight     =   9180
   ScaleWidth      =   11715
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRefNo 
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   1440
      TabIndex        =   2
      Top             =   360
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      Height          =   285
      Left            =   3720
      Picture         =   "Form3.frx":0000
      ScaleHeight     =   225
      ScaleWidth      =   195
      TabIndex        =   1
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton cmdnewsearch 
      Caption         =   "New "
      Height          =   285
      Left            =   3960
      TabIndex        =   0
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "Reference No. :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   0
      TabIndex        =   3
      Top             =   360
      Width           =   1395
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
