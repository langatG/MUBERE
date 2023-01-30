VERSION 5.00
Begin VB.Form frmmenuregister 
   Caption         =   "Menu Register"
   ClientHeight    =   3105
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6120
   Icon            =   "frmmenuregister.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3105
   ScaleWidth      =   6120
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdnew 
      Caption         =   "New"
      Height          =   375
      Left            =   480
      TabIndex        =   7
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox txtalias 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      ToolTipText     =   "Those starting with mnu"
      Top             =   1080
      Width           =   3735
   End
   Begin VB.TextBox txtbname 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      ToolTipText     =   "Clear name for the users"
      Top             =   480
      Width           =   3735
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CheckBox chkenable 
      Caption         =   "Enable/Disable"
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Alias"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Menu Name"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "frmmenuregister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim E As Integer

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdNew_Click()

txtalias = ""
txtbname = ""

End Sub

Private Sub cmdsave_Click()
If chkenable = True Then
E = 0
Else
E = 1
End If
sql = ""
sql = "SELECT     * FROM         tbl_menus where Alias='" & txtalias & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If rs.EOF Then
oSaccoMaster.ExecuteThis ("d_insert_tbl_menu '" & txtbname & "','" & txtalias & "'," & E & "")
Else
sql = ""
sql = "update tbl_menus set menu='" & txtbname & "' where alias='" & txtalias & "'"
oSaccoMaster.ExecuteThis (sql)
MsgBox "Record Already Available"
End If
MsgBox "Record Successfully Done"
End Sub

