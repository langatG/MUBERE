VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmStartup 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4890
   ClientLeft      =   2145
   ClientTop       =   915
   ClientWidth     =   7650
   ControlBox      =   0   'False
   Enabled         =   0   'False
   FillColor       =   &H00C0C0C0&
   ForeColor       =   &H00404000&
   Icon            =   "frmStartup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   7650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H0000FF00&
      Caption         =   "Frame1"
      Height          =   5415
      Left            =   0
      TabIndex        =   0
      Top             =   -600
      Width           =   7695
      Begin VB.PictureBox Picture1 
         Height          =   495
         Left            =   3480
         Picture         =   "frmStartup.frx":08CA
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   9
         Top             =   2640
         Width           =   495
      End
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   3960
         Top             =   5040
      End
      Begin VB.Timer Timer3 
         Interval        =   1
         Left            =   960
         Top             =   5040
      End
      Begin MSComctlLib.ProgressBar Bar 
         Height          =   375
         Left            =   1920
         TabIndex        =   1
         ToolTipText     =   "Loading..."
         Top             =   5040
         Visible         =   0   'False
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   661
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label Label13 
         Caption         =   "Copyright"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5280
         TabIndex        =   18
         Top             =   4800
         Width           =   1095
      End
      Begin VB.Label Label12 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4800
         TabIndex        =   16
         Top             =   4800
         Width           =   495
      End
      Begin VB.Label Label11 
         Caption         =   "@"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   15
         Top             =   4800
         Width           =   375
      End
      Begin VB.Label Label10 
         BackColor       =   &H0000FF00&
         Caption         =   "BIRGEN"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5520
         TabIndex        =   14
         Top             =   4440
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label5 
         BackColor       =   &H0000FF00&
         Caption         =   "&&"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5880
         TabIndex        =   13
         Top             =   4200
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "www.amtechafrica.com"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   120
         TabIndex        =   12
         Top             =   4680
         Width           =   2160
      End
      Begin VB.Label Label7 
         BackColor       =   &H0000FF00&
         Caption         =   "AMTECH TECHNOLOGIES LTD           "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   11
         Top             =   3960
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "                                                  AMTECH TECHNOLOGIES"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   480
         TabIndex        =   10
         Top             =   5040
         Visible         =   0   'False
         Width           =   6735
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H80000005&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   45
         Left            =   120
         Top             =   1440
         Width           =   7335
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EASYMA"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   1440
         TabIndex        =   8
         Top             =   480
         Width           =   4170
      End
      Begin VB.Label Label3 
         BackColor       =   &H00808080&
         Height          =   1215
         Left            =   0
         TabIndex        =   7
         Top             =   360
         Width           =   7695
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "info@amtechafrica.com"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   120
         TabIndex        =   6
         Top             =   4440
         Width           =   1905
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "For Comments  Please Us Email At "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   120
         TabIndex        =   5
         Top             =   4200
         Width           =   2850
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Haettenschweiler"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3450
         TabIndex        =   4
         Top             =   3240
         Width           =   555
      End
      Begin VB.Label lblWarning 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unauthorized Use of This Software Is Strictly Prohibited"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   3
         Top             =   3960
         Width           =   4560
      End
      Begin VB.Label lblCopyright 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright © 2003  SCaVeNGeR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   210
         Left            =   7800
         TabIndex        =   2
         Top             =   5340
         Width           =   2280
      End
   End
   Begin VB.Label Label8 
      Caption         =   "@"
      Height          =   255
      Left            =   9240
      TabIndex        =   17
      Top             =   3240
      Width           =   375
   End
End
Attribute VB_Name = "frmStartup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X, Y As Integer

Private Sub Timer1_Timer()
    Dim a As Long, b As Integer
    Label12 = Year(Now)
    Bar.value = Bar.value + 2
    Screen.MousePointer = vbHourglass
    If Bar.value <= 30 Then
    Label1 = "Initializing....."
    ElseIf Bar.value <= 50 Then
    Label1 = "Loading components....."
    ElseIf Bar.value <= 70 Then
    Label1 = "Integrating Database...."
    ElseIf Bar.value <= 100 Then
    Label1 = "Please wait..."
    End If
    If Bar.value = 100 Then
    If Timer1.Interval >= 1 Then
    Unload frmStartup
    Screen.MousePointer = vbDefault
    frmODBCLogon.Show
    End If
    End If
End Sub

