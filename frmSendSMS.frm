VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmSendSMS 
   Caption         =   "Send SMS"
   ClientHeight    =   6855
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   ScaleHeight     =   6855
   ScaleWidth      =   6150
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Message to Send"
      Height          =   3375
      Left            =   120
      TabIndex        =   1
      Top             =   3240
      Width           =   5895
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   135
         Left            =   120
         TabIndex        =   13
         Top             =   2760
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   238
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   255
         Left            =   3240
         TabIndex        =   11
         Top             =   3000
         Width           =   855
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "Send"
         Enabled         =   0   'False
         Height          =   255
         Left            =   480
         TabIndex        =   10
         Top             =   3000
         Width           =   855
      End
      Begin VB.TextBox txtMessage 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   2055
         Left            =   600
         MaxLength       =   160
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   0
         Top             =   600
         Width           =   3735
      End
      Begin VB.Label lblOperator 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   5040
         TabIndex        =   12
         Top             =   1080
         Width           =   45
      End
      Begin VB.Label lblNoChars 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "160"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   13.5
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   4800
         TabIndex        =   6
         Top             =   2160
         Width           =   630
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Farmers To Receive SMS"
      Height          =   3135
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5895
      Begin VB.OptionButton optmanualsel 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Enter Phone No."
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1920
         Width           =   1575
      End
      Begin VB.ComboBox cboFarmer 
         Height          =   315
         Left            =   3000
         TabIndex        =   16
         Top             =   1440
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.TextBox txtPhone 
         Height          =   525
         Left            =   840
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   15
         Top             =   2400
         Width           =   4815
      End
      Begin VB.ComboBox cboLoc 
         Height          =   315
         Left            =   3000
         TabIndex        =   9
         Top             =   1080
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.OptionButton optSpecificLocation 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0FF&
         Caption         =   "Only farmers from specific Location"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   2895
      End
      Begin VB.OptionButton optActiveFarmers 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0FF&
         Caption         =   "Active Farmers"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton optAllFarmers 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0FF&
         Caption         =   "All Farmers "
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker DTPPeriod 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   8
         Top             =   720
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         _Version        =   393216
         Format          =   147914753
         CurrentDate     =   40256
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phone :"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   2520
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Specific Farmer"
         Height          =   195
         Left            =   1680
         TabIndex        =   14
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label lblPeriod 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "For Period Ending : "
         Height          =   195
         Left            =   1680
         TabIndex        =   7
         Top             =   720
         Visible         =   0   'False
         Width           =   1395
      End
   End
End
Attribute VB_Name = "frmSendSMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)



Private Sub cboFarmer_Validate(Cancel As Boolean)
If UCase(cboFarmer.Text) = "<SELECT FARMERS>" Then
Exit Sub
End If

If UCase(cboFarmer.Text) <> "ALL FARMERS" Then
strSQL = "SELECT DISTINCT LTRIM(PhoneNo) AS Contact, SNo From d_Suppliers WHERE (LEN(PhoneNo)"
strSQL = strSQL & " = 10 OR LEN(PhoneNo) = 13) AND (Location = '" & cboLoc.Text & "') AND SNo = " & cboFarmer.Text & " GROUP BY SNo, PhoneNo"
End If

If UCase(cboFarmer.Text) = "ALL FARMERS" Then
strSQL = "SELECT DISTINCT LTRIM(PhoneNo) AS Contact, SNo From d_Suppliers WHERE (LEN(PhoneNo)"
strSQL = strSQL & " = 10 OR LEN(PhoneNo) = 13) AND (Location = '" & cboLoc.Text & "')  GROUP BY SNo, PhoneNo"
End If

Set rs = oSaccoMaster.GetRecordset(strSQL)

While Not rs.EOF
Dim j As Long
j = 0

If Trim(txtPhone) <> "" Then

rsNumbers() = Split(Trim(txtPhone), ",")

While j < (UBound(rsNumbers) + 1)
If rsNumbers(j) = rs.Fields(0) Then
GoTo Mwiraria
End If

j = j + 1
Wend
End If

If Trim(txtPhone) = "" Then
txtPhone = Trim(rs.Fields(0))
Else
txtPhone = txtPhone & "," & Trim(rs.Fields(0))
Mwiraria:
rs.MoveNext
End If
Wend

End Sub

Private Sub cboLoc_Validate(Cancel As Boolean)
strSQL = "SELECT DISTINCT LTRIM(PhoneNo) AS Contact, SNo From d_Suppliers WHERE (LEN(PhoneNo)"
strSQL = strSQL & " = 10 OR LEN(PhoneNo) = 13) AND (Location = '" & cboLoc.Text & "')GROUP BY SNo, PhoneNo"

cboFarmer.Clear
txtPhone = ""
Set rs = oSaccoMaster.GetRecordset(strSQL)
With rs
If .RecordCount > 0 Then
While Not .EOF
cboFarmer.AddItem (.Fields(1))
.MoveNext
Wend
End If
cboFarmer.AddItem ("All Farmers")
cboFarmer.Text = "<Select Farmers>"
cboFarmer_Validate True
End With
End Sub

Private Sub cmdClear_Click()
txtMessage = ""
txtMessage.SetFocus
End Sub

Private Sub cmdSend_Click()

MsgContent = txtMessage
If optSpecificLocation.value = True Then GoTo Birgen
'strSQL = "SELECT DISTINCT LTRIM(PhoneNo) AS Contact"
'strSQL = strSQL & "From d_Suppliers WHERE (LEN(PhoneNo) = 10) OR (LEN(PhoneNo) = 13)"

If optAllFarmers.value = True Then
strSQL = "SELECT DISTINCT LTRIM(PhoneNo) AS Contact"
strSQL = strSQL & " From d_Suppliers WHERE (LEN(PhoneNo) = 10) OR (LEN(PhoneNo) = 13)"
End If

If optmanualsel.value = True Then GoTo manualsel

If optActiveFarmers.value = True Then
strSQL = "SET dateformat DMY SELECT DISTINCT ltrim(d_Suppliers.PhoneNo)FROM d_Suppliers INNER JOIN"
strSQL = strSQL & " d_Payroll ON d_Suppliers.SNo = d_Payroll.SNo WHERE d_Payroll.EndofPeriod = '" & dtpPeriod & "' AND"
strSQL = strSQL & " (LEN(PhoneNo) = 10) OR(LEN(PhoneNo) = 13)"
End If

Set rs = oSaccoMaster.GetRecordset(strSQL)
If rs.RecordCount = 0 Then
MsgBox "There are no phone numbers."
Exit Sub
End If

MsgBox "This will send " & rs.RecordCount & " messages now"
With ProgressBar1
.max = rs.RecordCount
.Min = 0
.value = 0
End With
MsgContent = txtMessage
While Not rs.EOF
Dim a As Double
Dim Phonea As String
Phone = rs.Fields(0)
If Len(Phone) = 10 Then
 Phonea = "254"
 a = Phone
 Phone = Phonea + Trim(a)
End If
strSQL = "INSERT INTO Messages(Telephone,Content,ProcessTime, MsgType,Source,code)"
strSQL = strSQL & "Values ('" & Phone & "','" & MsgContent & "',getDate(),'Outbox','" & User & "','SAMLI')"

oSaccoMaster.ExecuteThis (strSQL)

ProgressBar1.value = ProgressBar1.value + 1

rs.MoveNext
Wend

Dim j As Integer
j = 0
Birgen:
If optSpecificLocation.value = True Then
If txtPhone = "" Then
MsgBox "No records to be send"
Exit Sub
End If

rsNumbers() = Split(txtPhone, ",")

MsgBox "This will send " & (UBound(rsNumbers())) & " messages now"

With ProgressBar1
.max = UBound(rsNumbers()) + 1
.Min = 0
.value = 0
End With

MsgContent = txtMessage

While j < (UBound(rsNumbers()) + 1)

Phone = rsNumbers(j)
If Len(Phone) = 10 Then
 Phonea = "254"
 a = Phone
 Phone = Phonea + Trim(a)
End If
strSQL = "INSERT INTO Messages(Telephone,Content,ProcessTime, MsgType,Source,Code)"
strSQL = strSQL & "Values ('" & Phone & "','" & MsgContent & "',getDate(),'Outbox','" & User & "','SAMLI')"

oSaccoMaster.ExecuteThis (strSQL)

ProgressBar1.value = ProgressBar1.value + 1
j = j + 1
Wend
End If
''''''''start manual sms sending
j = 0
manualsel:
    If optmanualsel.value = True Then
    If txtPhone = "" Then
    MsgBox "No records to be send"
    Exit Sub
    End If
    rsNumbers() = Split(txtPhone, ",")
    
    MsgBox "This will send " & (UBound(rsNumbers())) & " messages now"
    
    With ProgressBar1
    .max = UBound(rsNumbers()) + 1
    .Min = 0
    .value = 0
    End With
    
    MsgContent = txtMessage
    
    While j < (UBound(rsNumbers()) + 1)
    
    Phone = rsNumbers(j)
    If Len(Phone) = 10 Then
      Phonea = "254"
      a = Phone
      Phone = Phonea + Trim(a)
    End If
    
    strSQL = "INSERT INTO Messages(Telephone,Content,ProcessTime, MsgType,Source,Code)"
    strSQL = strSQL & "Values ('" & Phone & "','" & MsgContent & "',getDate(),'Outbox','" & User & "','SAMLI')"
    
    oSaccoMaster.ExecuteThis (strSQL)
    
    ProgressBar1.value = ProgressBar1.value + 1
    j = j + 1
    Wend
    End If
''''''end
MsgBox "Messages are ready to be forwarded by SMS monitor."
txtPhone = ""
ProgressBar1.value = 0
Set rs = Nothing
End Sub

Private Sub DTPPeriod_Validate(Cancel As Boolean)
dtpPeriod = DateSerial(Year(dtpPeriod), month(dtpPeriod) + 1, 1 - 1)
End Sub





Private Sub Form_GotFocus()
txtMessage.SetFocus

End Sub

Private Sub Form_Load()
rsNumbers = Split(cname, " ")
txtMessage = vbNewLine & "From " & UCase(mid(rsNumbers(0), 1, 1)) & "" & LCase(mid(rsNumbers(0), 2, (Len(rsNumbers(0)) - 1))) + " Dairy"
txtMessage_Change


dtpPeriod = Get_Server_Date

DTPPeriod_Validate True

cboLoc.Clear

strSQL = "SELECT DISTINCT Location From "
strSQL = strSQL & "d_Suppliers WHERE (Location <> '') ORDER BY Location"

Set rs = oSaccoMaster.GetRecordset(strSQL)
While Not rs.EOF
    cboLoc.AddItem (rs.Fields(0))
    rs.MoveNext
Wend



End Sub


Private Sub optActiveFarmers_Click()

If optActiveFarmers = True Then
lblPeriod.Visible = True
dtpPeriod.Visible = True
cboLoc.Visible = False
Label1.Visible = False
cboFarmer.Visible = False

Else
lblPeriod.Visible = False
dtpPeriod.Visible = False
End If
txtPhone.Locked = True
End Sub

Private Sub optAllFarmers_Click()

optActiveFarmers_Click
optSpecificLocation_Click
txtPhone.Locked = True

End Sub

Private Sub optmanualsel_Click()
optActiveFarmers_Click
optSpecificLocation_Click
txtPhone.Locked = False
End Sub

Private Sub optSpecificLocation_Click()

If optSpecificLocation = True Then
cboLoc.Visible = True
Label1.Visible = True
cboFarmer.Visible = True
Else
cboLoc.Visible = False
Label1.Visible = False
cboFarmer.Visible = False
End If

optActiveFarmers_Click
txtPhone.Locked = True
End Sub

Private Sub txtMessage_Change()
lblNoChars = 160 - Len(txtMessage)

If Trim(txtMessage) = "" Then
cmdSend.Enabled = False
Else
cmdSend.Enabled = True
End If
End Sub

