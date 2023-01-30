VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDUserSummary 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Form1"
   ClientHeight    =   1740
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   ScaleHeight     =   1740
   ScaleWidth      =   4800
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   240
      Top             =   840
   End
   Begin VB.ComboBox cboUsers 
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1560
      TabIndex        =   3
      Top             =   600
      Width           =   2895
   End
   Begin VB.CommandButton cmdBanks 
      Caption         =   "Show Farmers"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   960
      Width           =   2055
   End
   Begin MSComctlLib.ProgressBar progress 
      Height          =   135
      Left            =   0
      TabIndex        =   0
      Top             =   1560
      Visible         =   0   'False
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComCtl2.DTPicker dtpEndPeriod 
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   120
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      Format          =   108658689
      CurrentDate     =   40505
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Select Audit Id :"
      Height          =   195
      Left            =   120
      LinkItem        =   "&H00C0FFC0&"
      TabIndex        =   6
      Top             =   600
      Width           =   1125
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter End of Period :"
      Height          =   210
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1560
   End
   Begin VB.Label lblProgress 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Loading Please wait ... .. ."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   600
      TabIndex        =   4
      Top             =   1320
      Visible         =   0   'False
      Width           =   2025
   End
End
Attribute VB_Name = "frmDUserSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBanks_Click()
If Trim(cboUsers.Text) = "" Then
MsgBox "Please select Audit Id"
Exit Sub
End If
dtpEndPeriod = DateSerial(year(dtpEndPeriod), month(dtpEndPeriod) + 1, 1 - 1)
Startdate = DateSerial(year(dtpEndPeriod), month(dtpEndPeriod), 1)

cmdBanks.Caption = "Wait Loading"
Timer1.Enabled = True
'sleep 300

lblProgress.Visible = True
progress.Visible = True
progress.Max = 100
progress.Min = 0
progress.value = 10
'sleep (100)
oSaccoMaster.ExecuteThis ("d_sp_UpdateClerk '" & Startdate & "','" & dtpEndPeriod & "','" & cboUsers.Text & "'")
progress.value = 95

    reportname = "d_DailySummaryPerClerk.rpt"
    ReportTitle = dtpEndPeriod & "  BY  " & UCase(cboUsers.Text)
    '{d_Payroll.NPay} > 0 and {d_Payroll.Bank} <> '' and month({d_Payroll.EndofPeriod})= month(30/09/2010)  AND year({d_Payroll.EndofPeriod}) = Year(30/09/2010)
    'STRFORMULA = "{d_TransportersPayRoll.NetPay} > 0 and {d_TransportersPayRoll.BankName} = '" & cboBank & "' and month({d_TransportersPayRoll.EndPeriod})=" & month(dtpEndPeriod) & " AND year({d_TransportersPayRoll.EndPeriod}) =" & Year(dtpEndPeriod)
    Show_Sales_Crystal_Report STRFORMULA, reportname, ReportTitle
progress.value = 100
lblProgress.Visible = False
progress.Visible = False

cmdBanks.Caption = "Show Farmers"
Timer1.Enabled = True

'd_sp_UpdateClerk @StartDate varchar(15),@EndDate varchar(15),@AuditId varchar(35)

End Sub

Private Sub dtpEndPeriod_LostFocus()
dtpEndPeriod = DateSerial(year(dtpEndPeriod), month(dtpEndPeriod) + 1, 1 - 1)
End Sub
Private Sub Form_Load()
dtpEndPeriod = Get_Server_Date
dtpEndPeriod = DateSerial(year(dtpEndPeriod), month(dtpEndPeriod) + 1, 1 - 1)
Set rs = oSaccoMaster.GetRecordset("SELECT UserLoginIDs FROM UserAccounts")

Dim j As Integer
    
cboUsers.Clear

If rs.RecordCount > 0 Then
progress.Max = rs.RecordCount
lblProgress.Visible = True
progress.Visible = True
End If

progress.Min = 0
progress.value = 0

j = 1
While Not rs.EOF
progress.value = j
    cboUsers.AddItem rs.Fields(0)
    j = j + 1
    rs.MoveNext
Wend
    lblProgress.Visible = False
    progress.Visible = False



End Sub

Private Sub Timer1_Timer()
If cmdBanks.Caption = "Wait Loading..." Then
cmdBanks.Caption = "Wait Loading... "
Exit Sub
End If

If cmdBanks.Caption = "Wait Loading" Then
cmdBanks.Caption = "Wait Loading."
Exit Sub
End If

If cmdBanks.Caption = "Wait Loading... .." Then
cmdBanks.Caption = "Wait Loading"
Exit Sub
End If

If cmdBanks.Caption = "Wait Loading... ." Then
cmdBanks.Caption = "Wait Loading... .."
Exit Sub
End If

If cmdBanks.Caption = "Wait Loading... " Then
cmdBanks.Caption = "Wait Loading... ."
Exit Sub
End If

If cmdBanks.Caption = "Wait Loading.." Then
cmdBanks.Caption = "Wait Loading..."
Exit Sub
End If

If cmdBanks.Caption = "Wait Loading." Then
cmdBanks.Caption = "Wait Loading.."
Exit Sub
End If

End Sub
