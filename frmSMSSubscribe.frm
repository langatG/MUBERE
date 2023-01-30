VERSION 5.00
Begin VB.Form frmSMSSubscribe 
   BackColor       =   &H00C0C0FF&
   Caption         =   "SMS Subscriptions"
   ClientHeight    =   3330
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6795
   LinkTopic       =   "Form1"
   ScaleHeight     =   3330
   ScaleWidth      =   6795
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Tag             =   "TransCode"
      Top             =   120
      Width           =   6495
      Begin VB.ComboBox cboWhen 
         Height          =   315
         ItemData        =   "frmSMSSubscribe.frx":0000
         Left            =   4200
         List            =   "frmSMSSubscribe.frx":000D
         TabIndex        =   14
         Top             =   2160
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   1215
         Left            =   3600
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   13
         Text            =   "frmSMSSubscribe.frx":0029
         Top             =   240
         Width           =   2535
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         Height          =   375
         Left            =   3000
         TabIndex        =   12
         Top             =   2760
         Width           =   975
      End
      Begin VB.TextBox txtCode 
         Height          =   285
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   11
         Top             =   720
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox txtPhone 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         MaxLength       =   13
         TabIndex        =   10
         Top             =   1200
         Width           =   2175
      End
      Begin VB.ComboBox cboCode 
         Height          =   315
         ItemData        =   "frmSMSSubscribe.frx":009D
         Left            =   1200
         List            =   "frmSMSSubscribe.frx":009F
         TabIndex        =   9
         Top             =   720
         Width           =   2175
      End
      Begin VB.ComboBox cboFreq 
         Height          =   315
         ItemData        =   "frmSMSSubscribe.frx":00A1
         Left            =   1200
         List            =   "frmSMSSubscribe.frx":00AE
         TabIndex        =   8
         Top             =   2160
         Width           =   2175
      End
      Begin VB.ComboBox cboSub 
         Height          =   315
         ItemData        =   "frmSMSSubscribe.frx":00CA
         Left            =   1200
         List            =   "frmSMSSubscribe.frx":00D4
         TabIndex        =   7
         Top             =   1680
         Width           =   2175
      End
      Begin VB.ComboBox cboType 
         Height          =   315
         ItemData        =   "frmSMSSubscribe.frx":00EC
         Left            =   1200
         List            =   "frmSMSSubscribe.frx":00F9
         TabIndex        =   6
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lblWhen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "When"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3600
         TabIndex        =   15
         Top             =   2160
         Width           =   540
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phone :"
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Frequency :"
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   2160
         Width           =   840
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Subscription :"
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   1680
         Width           =   960
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Code :"
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type :"
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   450
      End
   End
End
Attribute VB_Name = "frmSMSSubscribe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cboCode_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Private Sub cboFreq_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cboFreq_Validate(Cancel As Boolean)
lblWhen.Visible = False
cboWhen.Visible = False

If UCase(cboFreq) = "DAILY" Then
cboWhen.Visible = False

End If

If UCase(cboFreq) = "WEEKLY" Then
With cboWhen
.Visible = True
.Clear
.AddItem ("Sunday")
.AddItem ("Monday")
.AddItem ("Tuesday")
.AddItem ("Wednesday")
.AddItem ("Thursday")
.AddItem ("Friday")
End With

cboWhen.Text = Format(Get_Server_Date, "dddd")
lblWhen.Visible = True
End If

End Sub

Private Sub cboSub_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cboType_Change()
cboCode.Clear
If UCase(Trim(cboType)) = "SUPPLIER" Then
cboCode.Visible = True
txtCode.Visible = False
Label2.Caption = "SNo :"
Set rs = oSaccoMaster.GetRecordset("SELECT SNo FROM d_Suppliers ORDER BY SNo")
While Not rs.EOF
cboCode.AddItem (rs.Fields(0))
rs.MoveNext
Wend
Set rs = Nothing
End If

If UCase(Trim(cboType)) = "TRANSPORTER" Then
cboCode.Visible = True
txtCode.Visible = False
Label2.Caption = "Code :"
Set rs = oSaccoMaster.GetRecordset("SELECT TransCode FROM d_Transporters ORDER BY TransCode")
While Not rs.EOF
cboCode.AddItem (rs.Fields(0))
rs.MoveNext
Wend
Set rs = Nothing
End If

If UCase(Trim(cboType)) = "STAFF" Then
cboCode.Visible = False
txtCode.Visible = True
Label2.Caption = "Name :"
End If

End Sub


Private Sub cboType_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cboType_Validate(Cancel As Boolean)
cboType_Change
End Sub

Private Sub cboWhen_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cmdupdate_Click()

If (Trim(cboType) = "") Or (Trim(cboCode) = "") Or (Trim(cboFreq) = "") Or (Trim(cboSub) = "") Or (Trim(txtPhone) = "") Then
MsgBox "Please enter all the details."
Exit Sub
End If

If (Mid(txtPhone, 1, 1) <> "0") And (Mid(txtPhone, 1, 1) <> "+") Then
MsgBox Mid(txtPhone, 1, 1)
MsgBox "Please enter a valid phone number."
txtPhone.SetFocus
Exit Sub
End If

If (Mid(txtPhone, 1, 1) = "0") And (Len(txtPhone) <> 10) Then
MsgBox "Please enter a valid phone number."
txtPhone.SetFocus
Exit Sub
End If

If (Mid(txtPhone, 1, 1) = "+") And (Len(txtPhone) <> 13) Then
MsgBox "Please enter a valid phone number."
txtPhone.SetFocus
Exit Sub
End If

strSQL = "SELECT Active, Code From SMSSubscription WHERE(Freq = '" & cboFreq.Text & "') AND "
strSQL = strSQL & "(Phone = '" & txtPhone & "') AND (Subscription = '" & cboSub & "')"

Set rs = oSaccoMaster.GetRecordset(strSQL)

If rs.RecordCount > 0 Then
  If rs.Fields(0) = 1 Then
  MsgBox cboType.Text & " " & rs.Fields(1) & " is already subscribed to deactivate sms E*" & rs.Fields(1)
  Exit Sub
  End If
  If rs.Fields(0) = 0 Then
  MsgBox cboType.Text & " " & rs.Fields(1) & " is already subscribed to activate sms s*" & rs.Fields(1)
  Exit Sub
  End If
End If

strSQL = "INSERT INTO SMSSubscription ( [Type],[Code],[Phone],[Subscription],[Freq],[AuditId]) "
strSQL = strSQL & " Values ('" & cboType.Text & "','" & cboCode.Text & "','" & txtPhone & "','" & cboSub.Text & "','"
strSQL = strSQL & cboFreq.Text & "','" & User & "')"

oSaccoMaster.ExecuteThis (strSQL)

MsgContent = "Thank you for subscribing to " & cboFreq.Text & " " & cboSub.Text & ", to activate your subscription reply this SMS with S*Transporter Code/S*Supplier Number. From " & cname
If UCase(cboType) = "STAFF" Then
txtCode = "Company"
End If
strSQL = "INSERT INTO Messages(Telephone,Content,ProcessTime, MsgType,Source,Code)"
strSQL = strSQL & "Values ('" & txtPhone & "','" & MsgContent & "',getDate(),'Outbox','" & User & "','" & txtCode & "')"

oSaccoMaster.ExecuteThis (strSQL)
MsgBox "Thank you for subscribing, to deactivate sms to use in activation"

End Sub



Private Sub txtCode_Change()
cboCode.Text = txtCode
End Sub
