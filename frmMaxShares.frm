VERSION 5.00
Begin VB.Form frmMaxShares 
   Caption         =   "Set Maximum Shares"
   ClientHeight    =   2430
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6255
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   2430
   ScaleWidth      =   6255
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   960
      TabIndex        =   4
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox txtMaxShares 
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   1920
      TabIndex        =   3
      Top             =   840
      Width           =   2895
   End
   Begin VB.TextBox txtIdNo 
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   1920
      TabIndex        =   2
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Maximum Amount"
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1680
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Id Number"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "frmMaxShares"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
If Trim(txtIdNo) = "" Then
    MsgBox "Please enter Id Number."
        txtIdNo.SetFocus
    Exit Sub
End If

If Trim(txtMaxShares) = "" Then
    MsgBox "Please enter the maximum shares."
        txtMaxShares.SetFocus
    Exit Sub
End If

oSaccoMaster.ExecuteThis ("d_sp_MaxShares '" & txtIdNo & "'," & txtMaxShares & ",'" & User & "'")

txtIdNo = ""
txtMaxShares = ""
MsgBox "Records updated successfully."

End Sub

Private Sub txtIdNo_Validate(Cancel As Boolean)
Set rs = oSaccoMaster.GetRecordset("SELECT MaxAmnt FROM d_MaxShares WHERE IdNo = '" & txtIdNo & "'")
If rs.RecordCount > 0 Then
txtMaxShares = Format(rs.Fields(0), "0.00")
Else
txtMaxShares = "0.00"
End If
End Sub
