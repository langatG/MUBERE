VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Begin VB.Form frmApprovePayReq 
   Caption         =   "PAYMENT APPROVALS"
   ClientHeight    =   6600
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9630
   LinkTopic       =   "Form1"
   ScaleHeight     =   6600
   ScaleWidth      =   9630
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboaction 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmApprovePayReq.frx":0000
      Left            =   1920
      List            =   "frmApprovePayReq.frx":000D
      TabIndex        =   5
      Top             =   1080
      Width           =   2055
   End
   Begin VB.TextBox txtestimate 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Text            =   "0"
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox txtcomments 
      Appearance      =   0  'Flat
      Height          =   615
      Left            =   1920
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   2160
      Width           =   4695
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      Height          =   375
      Left            =   5040
      TabIndex        =   0
      Top             =   5400
      Width           =   1215
   End
   Begin MSComctlLib.ListView lvwapprovals 
      Height          =   1575
      Left            =   0
      TabIndex        =   2
      Top             =   3480
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   2778
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "User"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Level"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblLPONo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1440
      TabIndex        =   14
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label lblInvId 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4080
      TabIndex        =   13
      Top             =   600
      Width           =   2655
   End
   Begin VB.Label Label7 
      Caption         =   "LPO #"
      Height          =   255
      Left            =   480
      TabIndex        =   12
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Invoice Id"
      Height          =   255
      Left            =   3120
      TabIndex        =   11
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "           PROCESS APPROVAL LEVELS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   10
      Top             =   0
      Width           =   5175
   End
   Begin VB.Label Label3 
      Caption         =   "History Option"
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Action"
      Height          =   255
      Left            =   480
      TabIndex        =   8
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Estimate Cost"
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Comments"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   2160
      Width           =   1095
   End
End
Attribute VB_Name = "frmApprovePayReq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdsave_Click()

If Trim(cboaction) = "" Then
MsgBox "Please select the action."
cboaction.SetFocus
Exit Sub
End If

oSaccoMaster.ExecuteThis ("UPDATE d_PaymentReq SET Posted = 1,Status='" & cboaction & "'")
MsgBox "Record Updated successfully."
cmdclose_Click

End Sub
