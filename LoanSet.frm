VERSION 5.00
Begin VB.Form frmLoanSet 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Loan Setting"
   ClientHeight    =   3945
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8010
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
   ScaleHeight     =   3945
   ScaleWidth      =   8010
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Check Status"
      Height          =   285
      Left            =   6240
      TabIndex        =   13
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "Change Status"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2760
      TabIndex        =   12
      Top             =   3480
      Width           =   1575
   End
   Begin VB.TextBox txtSNo 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1800
      TabIndex        =   10
      Top             =   240
      Width           =   855
   End
   Begin VB.Label lblStatus 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Current Status : No Loan"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   630
      Left            =   360
      TabIndex        =   11
      Top             =   2640
      Width           =   5505
   End
   Begin VB.Label lblSName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2760
      TabIndex        =   9
      Top             =   240
      Width           =   5175
   End
   Begin VB.Label lblIdNo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1800
      TabIndex        =   8
      Top             =   720
      Width           =   4215
   End
   Begin VB.Label lblAccNo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1800
      TabIndex        =   7
      Top             =   1200
      Width           =   4215
   End
   Begin VB.Label lblBBranch 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1800
      TabIndex        =   6
      Top             =   1680
      Width           =   4215
   End
   Begin VB.Label lblBName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   2160
      Width           =   4215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Bank Name"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bank Branch "
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   1245
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Account Number"
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Id No"
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier Number"
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1560
   End
End
Attribute VB_Name = "frmLoanSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdChange_Click()
txtSNo_Validate True

If Trim(lblSName) = "" Then
MsgBox "Enter a valid supplier"
txtSNo.SetFocus
Exit Sub
End If

If Trim(lblBName) = "" Then
MsgBox "Supplier has no bank name."
txtSNo.SetFocus
Exit Sub
End If

If lblStatus = "Current Status : Has Loan" Then
oSaccoMaster.ExecuteThis ("UPDATE d_Suppliers SET Loan = 0 WHERE SNo = " & txtSNo)
Else
oSaccoMaster.ExecuteThis ("UPDATE d_Suppliers SET Loan = 1 WHERE SNo = " & txtSNo)
End If
txtSNo_Validate True
MsgBox "Records Saved Successfully!"
End Sub

Private Sub Command1_Click()
txtSNo_Validate True
End Sub

Private Sub txtSNo_Change()
If Trim(txtSNo) = "" Then
cmdChange.Enabled = False
Else
cmdChange.Enabled = True
End If
End Sub

Private Sub txtSNo_KeyPress(KeyAscii As Integer)
If (KeyAscii = 8) Or (KeyAscii = 48) Or (KeyAscii = 49) Or (KeyAscii = 50) Or (KeyAscii = 51) Or (KeyAscii = 52) Or (KeyAscii = 53) Or (KeyAscii = 54) Or (KeyAscii = 55) Or (KeyAscii = 56) Or (KeyAscii = 57) Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
MsgBox "Please enter a number "
End If

End Sub

Private Sub txtSNo_Validate(Cancel As Boolean)
lblSName = ""
lblIdNo = ""
lblAccNo = ""
lblBBranch = ""
lblBName = ""
lblStatus = "Current Status : No Loan"

If Trim(txtSNo) = "" Then
Exit Sub
End If

Set rs = oSaccoMaster.GetRecordset("SELECT [Names], IdNo, AccNo, BBranch, Bcode,Loan From d_Suppliers WHERE SNo = " & txtSNo)
If rs.RecordCount > 0 Then
If Not IsNull(rs.Fields(0)) Then lblSName = rs.Fields(0)
If Not IsNull(rs.Fields(1)) Then lblIdNo = rs.Fields(1)
If Not IsNull(rs.Fields(2)) Then lblAccNo = rs.Fields(2)
If Not IsNull(rs.Fields(3)) Then lblBBranch = rs.Fields(3)
If Not IsNull(rs.Fields(4)) Then lblBName = rs.Fields(4)
If rs.Fields(5) = True Then
lblStatus = "Current Status : Has Loan"
End If
End If
End Sub
