VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPayBanks 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Pay Banks"
   ClientHeight    =   2175
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4875
   BeginProperty Font 
      Name            =   "Sylfaen"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   4875
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton optTransporter 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Caption         =   "Transporters"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      Top             =   0
      Width           =   2055
   End
   Begin VB.OptionButton optSuppliers 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Caption         =   "Suppliers"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   120
      MaskColor       =   &H00FFC0C0&
      TabIndex        =   7
      Top             =   0
      Value           =   -1  'True
      Width           =   1695
   End
   Begin MSComctlLib.ProgressBar progress 
      Height          =   135
      Left            =   0
      TabIndex        =   6
      Top             =   1080
      Visible         =   0   'False
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdBanks 
      Caption         =   "Show Farmers"
      Default         =   -1  'True
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   1680
      Width           =   2055
   End
   Begin MSComCtl2.DTPicker dtpEndPeriod 
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   480
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      Format          =   118095873
      CurrentDate     =   40505
   End
   Begin VB.ComboBox cboBank 
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
      TabIndex        =   2
      Top             =   1200
      Width           =   2895
   End
   Begin VB.Label lblProgress 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "Please wait ... .. . loading Banks"
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
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Visible         =   0   'False
      Width           =   2550
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "Enter End of Period :"
      Height          =   330
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2040
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "Select Bank :"
      Height          =   330
      Left            =   120
      LinkItem        =   "&H00C0FFC0&"
      TabIndex        =   0
      Top             =   1200
      Width           =   1260
   End
End
Attribute VB_Name = "frmPayBanks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBanks_Click()
If optTransporter.value = True Then
    reportname = "d_GroupBankTrans.rpt"
    ReportTitle = "TO :" & UCase(cboBank) & " ; " & vbNewLine & " Please pay the following transporters the amount indicated: (Our Ref is code)"
    '{d_Payroll.NPay} > 0 and {d_Payroll.Bank} <> '' and month({d_Payroll.EndofPeriod})= month(30/09/2010)  AND year({d_Payroll.EndofPeriod}) = Year(30/09/2010)
    STRFORMULA = "{d_TransportersPayRoll.NetPay} > 0 and {d_TransportersPayRoll.BankName} = '" & cboBank & "' and month({d_TransportersPayRoll.EndPeriod})=" & month(dtpEndPeriod) & " AND year({d_TransportersPayRoll.EndPeriod}) =" & year(dtpEndPeriod)
    Show_Sales_Crystal_Report STRFORMULA, reportname, ReportTitle
End If
If optSuppliers.value = True Then
    reportname = "d_GroupBank.rpt"
    ReportTitle = "TO :" & UCase(cboBank) & " ; " & vbNewLine & " Please pay the following farmers the amount indicated: (Our Ref is SNo)"
    '{d_Payroll.NPay} > 0 and {d_Payroll.Bank} <> '' and month({d_Payroll.EndofPeriod})= month(30/09/2010)  AND year({d_Payroll.EndofPeriod}) = Year(30/09/2010)
    STRFORMULA = "{d_Payroll.NPay} > 0 and {d_Payroll.Bank} = '" & cboBank & "' and month({d_Payroll.EndofPeriod})=" & month(dtpEndPeriod) & " AND year({d_Payroll.EndofPeriod}) =" & year(dtpEndPeriod)
    Show_Sales_Crystal_Report STRFORMULA, reportname, ReportTitle
End If
'    d_StmtA4
End Sub

Private Sub dtpEndPeriod_LostFocus()
LoadBanks
End Sub

Private Sub Form_Load()
dtpEndPeriod = Get_Server_Date
dtpEndPeriod = DateSerial(year(dtpEndPeriod), month(dtpEndPeriod) + 1, 1 - 1)
LoadBanks
End Sub
Private Sub LoadBanks()
Dim j As Integer
    lblProgress.Visible = True
    progress.Visible = True
    dtpEndPeriod = DateSerial(year(dtpEndPeriod), month(dtpEndPeriod) + 1, 1 - 1)
    
cboBank.Clear

If optSuppliers.value = True Then
strSQL = "SET dateformat dmy SELECT DISTINCT LTRIM(Bank) AS Bank"
strSQL = strSQL & " From d_Payroll WHERE (NPay > 0) AND (Bank <> '')AND endofperiod = '" & dtpEndPeriod & "'"
End If

If optTransporter.value = True Then
strSQL = "SET dateformat dmy SELECT DISTINCT LTRIM(BankName) AS Bank"
strSQL = strSQL & " From d_TransportersPayRoll WHERE (NetPay > 0) AND (BankName <> '')AND EndPeriod = '" & dtpEndPeriod & "'"
End If

Set rs = oSaccoMaster.GetRecordset(strSQL)

If rs.RecordCount > 0 Then
progress.Max = rs.RecordCount
End If
progress.Min = 0
progress.value = 0

j = 1
While Not rs.EOF
progress.value = j
    cboBank.AddItem rs.Fields(0)
    j = j + 1
    rs.MoveNext
Wend
    lblProgress.Visible = False
    progress.Visible = False
End Sub

Private Sub optSuppliers_Click()
cmdBanks.Caption = "Show Farmers"
LoadBanks
End Sub

Private Sub optTransporter_Click()
cmdBanks.Caption = "Show Transporters"
LoadBanks
End Sub
