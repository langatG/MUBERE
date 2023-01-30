VERSION 5.00
Begin VB.Form Frmqualitysetup 
   Caption         =   "Quality setup"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4590
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Quality setup"
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.CommandButton Cmdclose 
         Caption         =   "&Close"
         Height          =   255
         Left            =   2880
         TabIndex        =   8
         Top             =   1680
         Width           =   855
      End
      Begin VB.CommandButton cmdsave 
         Caption         =   "&Save"
         Height          =   255
         Left            =   2040
         TabIndex        =   7
         Top             =   1680
         Width           =   615
      End
      Begin VB.CommandButton Cmdedit 
         Caption         =   "&Edit"
         Height          =   255
         Left            =   1080
         TabIndex        =   6
         Top             =   1680
         Width           =   735
      End
      Begin VB.CommandButton Cmdnew 
         Caption         =   "&New"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox txtrate 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1080
         TabIndex        =   4
         Top             =   840
         Width           =   1095
      End
      Begin VB.ComboBox Cboquality 
         Height          =   315
         ItemData        =   "Frmqualitysetup.frx":0000
         Left            =   1080
         List            =   "Frmqualitysetup.frx":000D
         TabIndex        =   2
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Rate"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Quality"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   480
      End
   End
End
Attribute VB_Name = "Frmqualitysetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdNew_Click()
Cboquality = ""
txtrate = ""
End Sub

Private Sub cmdsave_Click()
On Error GoTo errorhandler
Dim rsquality As New Recordset
If Cboquality = "" Then
    MsgBox "please select the quality", vbInformation
    Exit Sub
End If

If txtrate = "" Then
    MsgBox "please enter the rate", vbInformation
    Exit Sub
End If
Set rsquality = oSaccoMaster.GetRecordset("select * from qsetup where quality='" & Cboquality & "'")
If Not rsquality.EOF Then
oSaccoMaster.Execute "update qsetup set irate=" & txtrate & " where quality='" & Cboquality & "' "

Else
Set cn = New ADODB.Connection
sql = "insert into Qsetup(Quality,irate,auditid)values('" & Cboquality & "','" & txtrate & "','" & User & "')"
oSaccoMaster.ExecuteThis (sql)
End If
Cboquality = ""
txtrate = ""

'loadstaffs
MsgBox "Records successively updated."
Exit Sub
errorhandler:
MsgBox err.description
End Sub
