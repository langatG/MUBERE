VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmqbmps 
   Caption         =   "Qbmps"
   ClientHeight    =   3960
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7920
   LinkTopic       =   "Form1"
   ScaleHeight     =   3960
   ScaleWidth      =   7920
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "&Close"
      Height          =   375
      Left            =   4440
      TabIndex        =   22
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Save"
      Height          =   375
      Left            =   3240
      TabIndex        =   21
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Add"
      Height          =   375
      Left            =   2160
      TabIndex        =   20
      Top             =   3120
      Width           =   855
   End
   Begin VB.ComboBox cbooffce 
      Height          =   315
      Left            =   4560
      TabIndex        =   19
      Top             =   240
      Width           =   2175
   End
   Begin VB.ComboBox cboscore 
      Height          =   315
      Left            =   5040
      TabIndex        =   17
      Top             =   1800
      Width           =   1935
   End
   Begin VB.ComboBox Cboremarks 
      Height          =   315
      Left            =   5040
      TabIndex        =   15
      Top             =   1320
      Width           =   1815
   End
   Begin VB.ComboBox cbototals 
      Height          =   315
      Left            =   5040
      TabIndex        =   13
      Top             =   2280
      Width           =   1815
   End
   Begin VB.ComboBox Cboadulteration 
      Height          =   315
      Left            =   1680
      TabIndex        =   11
      Top             =   2400
      Width           =   1815
   End
   Begin VB.ComboBox Cboantires 
      Height          =   315
      Left            =   1680
      TabIndex        =   9
      Top             =   1920
      Width           =   1815
   End
   Begin VB.ComboBox Cbotpc 
      Height          =   315
      Left            =   1680
      TabIndex        =   7
      Top             =   1440
      Width           =   1815
   End
   Begin MSComCtl2.DTPicker DTPenddate 
      Height          =   375
      Left            =   4560
      TabIndex        =   5
      Top             =   720
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      Format          =   118030337
      CurrentDate     =   42715
   End
   Begin MSComCtl2.DTPicker DTPstartdate 
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   840
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      Format          =   118030337
      CurrentDate     =   42715
   End
   Begin VB.TextBox Ttxcanno 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Office"
      Height          =   195
      Left            =   3720
      TabIndex        =   18
      Top             =   360
      Width           =   420
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "% Sore"
      Height          =   195
      Left            =   4080
      TabIndex        =   16
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Remarks"
      Height          =   195
      Left            =   3960
      TabIndex        =   14
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Totalsolids"
      Height          =   195
      Left            =   3600
      TabIndex        =   12
      Top             =   2400
      Width           =   750
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Adulteration"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   2400
      Width           =   1080
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Antibioticresidue"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   1395
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Totalpalecount"
      Height          =   195
      Left            =   0
      TabIndex        =   6
      Top             =   1440
      Width           =   1305
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Enddate"
      Height          =   195
      Left            =   3600
      TabIndex        =   4
      Top             =   840
      Width           =   600
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Startdate"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Canno"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   465
   End
End
Attribute VB_Name = "frmqbmps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
