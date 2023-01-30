VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmRangeSelect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Data Range Selection:Member No."
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4410
   Icon            =   "frmRangeSelect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   4410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   2775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   4895
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Range Selector"
      TabPicture(0)   =   "frmRangeSelect.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "CommandButton1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "CommandButton2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtFrom"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtTo"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdOK"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdCancel"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   350
         Left            =   2040
         TabIndex        =   6
         Top             =   2280
         Width           =   1335
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   350
         Left            =   600
         TabIndex        =   5
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox txtTo 
         Height          =   350
         Left            =   240
         TabIndex        =   4
         Top             =   1680
         Width           =   3495
      End
      Begin VB.TextBox txtFrom 
         Height          =   350
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   3495
      End
      Begin VB.Frame Frame1 
         Height          =   25
         Left            =   0
         TabIndex        =   2
         Top             =   320
         Width           =   4215
      End
      Begin VB.Frame Frame2 
         Height          =   25
         Left            =   0
         TabIndex        =   1
         Top             =   2160
         Width           =   4335
      End
      Begin VB.PictureBox CommandButton2 
         Height          =   345
         Left            =   3720
         ScaleHeight     =   285
         ScaleWidth      =   315
         TabIndex        =   10
         Top             =   1680
         Width           =   375
      End
      Begin VB.PictureBox CommandButton1 
         Height          =   350
         Left            =   3720
         ScaleHeight     =   285
         ScaleWidth      =   315
         TabIndex        =   9
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "To Member No."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   1440
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "From Member No."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   1260
      End
   End
End
Attribute VB_Name = "frmRangeSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strfrom As String
Dim strto As String
Private Sub cmdcancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
'With frmUtilGenMemStatements
    strfrom = txtFrom.Text
    strto = txtTo.Text
'End With
Me.Hide
Unload Me
End Sub

Private Sub CommandButton1_Click()
isItFrom = True
frmRecSelect.Show vbModal, Me
Me.txtFrom = strRangeFrom
End Sub

Private Sub CommandButton2_Click()
isItFrom = False
frmRecSelect.Show vbModal, Me
txtTo = strRangeTo
txtFrom = strRangeFrom
End Sub

