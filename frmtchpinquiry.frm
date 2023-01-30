VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmtchpinquiry 
   Caption         =   "TCHP TRACKERS"
   ClientHeight    =   8625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   12435
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "CREATE SMS  WARNING LIST"
      Height          =   495
      Left            =   10440
      TabIndex        =   34
      Top             =   5400
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "TCHP Status Report"
      Height          =   375
      Left            =   12600
      TabIndex        =   33
      Top             =   1200
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdtchpautomaticdeductionmade 
      Caption         =   "Automatic Dedcution Report"
      Height          =   495
      Left            =   12600
      TabIndex        =   32
      Top             =   2400
      Width           =   1935
   End
   Begin VB.CommandButton cmdudatestatus 
      Caption         =   "TCHP Status Update"
      Height          =   255
      Left            =   12600
      TabIndex        =   31
      Top             =   360
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton cmdsmsreport 
      Caption         =   "SMS Transmission Report"
      Height          =   615
      Left            =   10440
      TabIndex        =   30
      Top             =   6840
      Width           =   1935
   End
   Begin VB.CommandButton cmdtimelineauditreport 
      Caption         =   "TCHP Audit Report"
      Height          =   375
      Left            =   10440
      TabIndex        =   29
      Top             =   3600
      Width           =   1815
   End
   Begin VB.CommandButton cmdtrackerreports 
      Caption         =   "TCHP Tracker Report"
      Height          =   375
      Left            =   10440
      TabIndex        =   28
      Top             =   3000
      Width           =   1815
   End
   Begin VB.CommandButton cmdtracker 
      Caption         =   "CREATE SMS CONFIRMATON LIST"
      Height          =   495
      Left            =   10440
      TabIndex        =   27
      Top             =   6120
      Width           =   1935
   End
   Begin VB.CommandButton cmdfind 
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1800
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   0
      Width           =   375
   End
   Begin MSComctlLib.ListView LV 
      Height          =   5415
      Left            =   120
      TabIndex        =   20
      Top             =   2160
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   9551
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Date Of Transaction"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Type of Transaction"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Debits"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Deductions(Credit)"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Cash Receipts(Credits)"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "TCHP Balance"
         Object.Width           =   2469
      EndProperty
   End
   Begin VB.TextBox txtSNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   480
      TabIndex        =   9
      Top             =   0
      Width           =   1335
   End
   Begin VB.TextBox txtTotalDeductions 
      Appearance      =   0  'Flat
      Height          =   495
      Index           =   0
      Left            =   1200
      TabIndex        =   8
      Top             =   8040
      Width           =   1575
   End
   Begin VB.TextBox txtTotalCashReceipts 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   4320
      TabIndex        =   7
      Top             =   8040
      Width           =   1575
   End
   Begin VB.TextBox txtTCHPBalances 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   7200
      TabIndex        =   6
      Top             =   8040
      Width           =   1815
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   0
      Width           =   2535
   End
   Begin VB.TextBox txtMilkAccountBalance 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox txtTCHPMonthlyPremium 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   0
      Width           =   2175
   End
   Begin VB.TextBox txtTCHPMembershipDate 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   480
      Width           =   2175
   End
   Begin VB.TextBox txtTCHPCurrentStatus 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   960
      Width           =   2175
   End
   Begin VB.TextBox txtTCHPBALANCE 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1320
      Width           =   2175
   End
   Begin MSComCtl2.DTPicker DTPEndDate 
      Height          =   255
      Left            =   2280
      TabIndex        =   22
      Top             =   1440
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      _Version        =   393216
      Format          =   102432769
      CurrentDate     =   40096
   End
   Begin MSComCtl2.DTPicker DTPStartDate 
      Height          =   255
      Left            =   240
      TabIndex        =   23
      Top             =   1440
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
      _Version        =   393216
      Format          =   102432769
      CurrentDate     =   40096
   End
   Begin VB.Label Label13 
      Caption         =   "Current Milk Period"
      Height          =   255
      Left            =   360
      TabIndex        =   26
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      Caption         =   "End Date"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2280
      TabIndex        =   25
      Top             =   1200
      Width           =   675
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      Caption         =   "Start Date"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   24
      Top             =   1200
      Width           =   720
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   10080
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   10200
      Y1              =   7680
      Y2              =   7680
   End
   Begin VB.Label Label1 
      Caption         =   "SNo:"
      Height          =   255
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Total Deductions"
      Height          =   495
      Index           =   0
      Left            =   0
      TabIndex        =   18
      Top             =   8040
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Total Cash Receipts"
      Height          =   375
      Left            =   3000
      TabIndex        =   17
      Top             =   8040
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "TCHP Balance"
      Height          =   375
      Left            =   6120
      TabIndex        =   16
      Top             =   8040
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Name:"
      Height          =   255
      Left            =   2400
      TabIndex        =   15
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "Milk Account Balance"
      Height          =   255
      Left            =   2400
      TabIndex        =   14
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label7 
      Caption         =   "TCHP Monthly Premium"
      Height          =   255
      Left            =   5880
      TabIndex        =   13
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label Label8 
      Caption         =   "TCHP Membership Date"
      Height          =   255
      Left            =   5880
      TabIndex        =   12
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label9 
      Caption         =   "TCHP Current status"
      Height          =   255
      Left            =   5880
      TabIndex        =   11
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label10 
      Caption         =   "TCHP BALANCE"
      Height          =   255
      Left            =   5880
      TabIndex        =   10
      Top             =   1440
      Width           =   1455
   End
End
Attribute VB_Name = "frmtchpinquiry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DTPDDeduction As Date
Private Sub cmdfind_Click()
        Me.MousePointer = vbHourglass
        txtMilkAccountBalance = 0
        txttchpbalance = 0
        txtTCHPCurrentStatus = ""
        txttchpbalance = 0
        txtTCHPMembershipDate = ""
        txtTCHPMonthlyPremium = 0
        txtTotalCashReceipts = 0
        'txtTotalDeductions = 0
        
        frmSearchSupplier.Show vbModal
        txtSNo = sel
        txtSNo_Validate True
        Me.MousePointer = 0
End Sub

Private Sub cmdsmsreport_Click()
'//DO THE REPORT HERE.
'SMSTransmissionReport
reportname = "SMSTransmissionReport.rpt"
Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub cmdtchpautomaticdeductionmade_Click()
'automaticdeductions
reportname = "automaticdeductions.rpt"
 Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub cmdtimelineauditreport_Click()
'TCHPTimelineAuditReport
'tchpauditreport99
'frmtchpauditrange.Show vbModal
'Exit Sub
reportname = "tchpauditreport99.rpt"
 Show_Sales_Crystal_Report STRFORMULA, reportname, ""
'reportname = "TCHPTimelineAuditReport.rpt"
 'Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub cmdtracker_Click()
'generate them into the same table
On Error GoTo ErrorHandler
Dim sno As String
Dim premium As Double
Dim tmdate As Date
Dim aarno As String
'//before all this let us clear this table called
sql = "truncate table         tchp_trxsreport"
oSaccoMaster.ExecuteThis (sql)

sql = ""
sql = "select sno,mpremium,Tmdate,statusr,aarno from tchp_members WHERE tchpactive=1 order by sno "
Set rs = oSaccoMaster.GetRecordset(sql)
While Not rs.EOF
'tchp_trxsreport
sno = rs.Fields(0)
Debug.Print sno
premium = rs.Fields(1)
tmdate = rs.Fields(2)
Dim a As Double, b As Double, C As Double, balance As Double, status As String
status = Trim(rs.Fields(3))
aarno = rs.Fields(4)
'tchp_trxs
Set Rst = oSaccoMaster.GetRecordset("SELECT     SUM(Debits) AS a, SUM(CreditsD) AS b, SUM(CreditsC) AS c  FROM         tchp_trxs  WHERE     (sno = '" & sno & "')  GROUP BY sno")
If Not Rst.EOF Then
            a = Rst.Fields(0)
            b = Rst.Fields(1)
            C = Rst.Fields(2)
            
            '//get balance here
            
                    sql = "SELECT     balance   FROM         tchp_trxs  WHERE     sno ='" & sno & "'  ORDER BY transdate DESC, id DESC "
                    Dim rr As New ADODB.Recordset
                    Set rr = oSaccoMaster.GetRecordset(sql)
                    If Not rr.EOF Then
                    balance = rr.Fields(0)
                    Else
                    balance = 0
                    End If
                  
                  
kapjoel:
            Dim MsgContent As String, Phone As String, rt As New ADODB.Recordset
            sql = ""
            sql = "select phoneno from d_suppliers where sno='" & sno & "'"
            Set rt = oSaccoMaster.GetRecordset(sql)
            If Not rt.EOF Then
            Phone = rt.Fields(0)
            Else
            Phone = ""
            End If
            Dim dprem As Double
            If Phone <> "" Then
            If Len(Phone) >= 10 Then
            
                    If status = "Terminate" Then
                    dprem = 2 * premium
                    If balance >= dprem Then
                        MsgContent = "Supplier No. " & sno & ", You have an outstanding TCHP balance of " & Format(balance, "###,###.00") & " and will be terminated from the scheme. You can rejoin the scheme in 12 months time. We will not deduct any more money from your Milk account"
                       
                    End If
                    End If
                    
                    If status = "Suspend" Then
                    If balance = premium Then
                        MsgContent = "Supplier No. " & sno & ", You have an outstanding TCHP balance of " & Format(balance, "###,###.00") & " and will be Suspended from cover next month. Please pay two premiums next month to regain cover"
                    End If
                    End If
                    If status = "status" Then
                    MsgContent = ""
                    End If
                    Else
                    Phone = ""
                    MsgContent = ""
            End If
            End If
                  'status here

            sql = ""
            sql = "INSERT INTO tchp_trxsreport"
            sql = sql & "                   (sno, Debits, CreditsD, CreditsC, Balance, status,premium,phone,content,msgtype,aarno)"
            sql = sql & "  VALUES     ('" & sno & "'," & a & "," & b & "," & C & "," & balance & ",'" & status & "'," & premium & ",'" & Phone & "','" & MsgContent & "','Outbox','" & aarno & "')"
            oSaccoMaster.ExecuteThis (sql)
            sql = ""
            sql = "UPDATE    tchp_members  SET      statusr='" & status & "'         where sno='" & sno & "'"
            oSaccoMaster.ExecuteThis (sql)
            
            '//insert into audit reports
            Dim rm As New ADODB.Recordset
            Set rm = oSaccoMaster.GetRecordset("select sno from tchp_audit where sno='" & sno & "'")
            If rm.EOF Then
             sql = ""
            sql = "INSERT INTO tchp_audit"
            sql = sql & "                   (sno, Debits, CreditsD, CreditsC, Balance, status,premium,jan2012,Febstatus)"
            sql = sql & "  VALUES     ('" & sno & "'," & a & "," & b & "," & C & "," & balance & ",'" & status & "'," & premium & "," & balance & ",'" & status & "')"
            oSaccoMaster.ExecuteThis (sql)
            Else
           
            If month(DTPEndDate) = 2 And Year(DTPEndDate) = 2012 Then
                sql = ""
                sql = "update tchp_audit set Feb2012=" & balance & ",Marstatus='" & status & "' where sno='" & sno & "' "
                oSaccoMaster.ExecuteThis (sql)
            End If
            
            If month(DTPEndDate) = 3 And Year(DTPEndDate) = 2012 Then
                sql = ""
                sql = "update tchp_audit set Mar2012=" & balance & ",Aprilstatus='" & status & "' where sno='" & sno & "' "
                oSaccoMaster.ExecuteThis (sql)
            End If
            
            If month(DTPEndDate) = 4 And Year(DTPEndDate) = 2012 Then
                sql = ""
                sql = "update tchp_audit set Apr2012=" & balance & ",Maystatus='" & status & "' where sno='" & sno & "' "
                oSaccoMaster.ExecuteThis (sql)
            End If
            
            If month(DTPEndDate) = 5 And Year(DTPEndDate) = 2012 Then
                sql = ""
                sql = "update tchp_audit set May2012=" & balance & ",junstatus='" & status & "' where sno='" & sno & "' "
                oSaccoMaster.ExecuteThis (sql)
            End If
            
            If month(DTPEndDate) = 6 And Year(DTPEndDate) = 2012 Then
                sql = ""
                sql = "update tchp_audit set june2012=" & balance & ",julstatus='" & status & "' where sno='" & sno & "' "
                oSaccoMaster.ExecuteThis (sql)
            End If
            
            End If
            '//got and update all the months
            
            
End If
MsgContent = ""
sql = ""
Phone = ""
sno = ""
rs.MoveNext
Wend
MsgBox "Records successfully generated", vbInformation


frmtchpsmslist.Show vbModal, Me
Exit Sub

ErrorHandler:
MsgBox err.description

End Sub

Private Sub cmdtrackerreports_Click()
'TCHPTrackerReport
 reportname = "TCHPTrackerReport.rpt"
 Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub cmdudatestatus_Click()
frmtchppremiumchange.Show vbModal, Me
End Sub



Private Sub Command1_Click()
frmtchpsmswarninglist.Show vbModal, Me
End Sub

Private Sub Command2_Click()
'TCHPTimelineAuditReport
reportname = "TCHP_MemberList_status.rpt"
 Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub Form_Load()
cmdtchpautomaticdeductionmade.Visible = False
DTPDDeduction = Format(Get_Server_Date, "dd/mm/yyyy")
DTPStartDate = DateSerial(Year(DTPDDeduction), month(DTPDDeduction), 1)
DTPEndDate = DateSerial(Year(DTPStartDate), month(DTPStartDate) + 1, 1 - 1)
End Sub

Private Sub txtSNo_Validate(Cancel As Boolean)
On Error GoTo ErrorHandler
Dim tchpa As Integer
Set rs = New ADODB.Recordset
sql = "d_sp_SelectSuppliers '" & txtSNo & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
If Not IsNull(rs.Fields(2)) Then txtName.Text = rs.Fields(2)
If Not IsNull(rs.Fields(24)) Then txtTCHPMembershipDate = rs.Fields(24)
Else
txtName.Text = ""
End If
LV.ListItems.Clear
''tchp_tchpmember
''SELECT     sno, aarno, mpremium, premium, tchpactive,balance,Tmdate    FROM         tchp_members where sno=@sno
Set Rst = New ADODB.Recordset
sql = "tchp_tchpmember '" & txtSNo & "'"
Set Rst = oSaccoMaster.GetRecordset(sql)
If Not Rst.EOF Then
txtTCHPMonthlyPremium = Rst.Fields(2)
txtTCHPBalances = Rst.Fields(5)
txtTCHPMembershipDate = Rst.Fields(7)
tchpa = Rst.Fields(4)
'
txtTCHPCurrentStatus = IIf(IsNull(Rst.Fields(6)), "", Rst.Fields(6))
If txtTCHPCurrentStatus = "" Then
If tchpa = 1 Then
txtTCHPCurrentStatus = ""
Else
txtTCHPCurrentStatus = ""
End If
End If

'//
End If

'//get the milk balance at this screen
'//get the milk balance at all the time
Dim Startdate As Date, NetP As Double
Dim Enddate As Date
Startdate = DateSerial(Year(DTPDDeduction), month(DTPDDeduction), 1)
Enddate = DateSerial(Year(DTPDDeduction), month(DTPDDeduction) + 1, 1 - 1)

Set rs = oSaccoMaster.GetRecordset("d_sp_SupNet '" & txtSNo & "','" & Startdate & "','" & Enddate & "', 0")
If rs.EOF Then GoTo jericho
If Not IsNull(rs.Fields(1)) Then
If Not IsNull(rs.Fields(1)) Then
NetP = rs.Fields(1)
Else
NetP = "0.00"
End If
jericho:
Set rs = oSaccoMaster.GetRecordset("d_sp_SupNet '" & txtSNo & "','" & Startdate & "','" & Enddate & "', 1")
If Not IsNull(rs.Fields(0)) Then
NetP = NetP - rs.Fields(0)
Else
NetP = NetP - 0
End If
txtMilkAccountBalance = NetP
End If
'//get the balance from the trx table
sql = "SELECT     balance   FROM         tchp_trxs  WHERE     sno ='" & txtSNo & "'  ORDER BY transdate DESC, id DESC "
Dim rr As New ADODB.Recordset
Set rr = oSaccoMaster.GetRecordset(sql)
If Not rr.EOF Then
txtTCHPBalances = rr.Fields(0)
txttchpbalance = rr.Fields(0)
End If

'//put data on the listview
sql = "SELECT     sno, transdate, description, Debits, CreditsD, CreditsC, Balance   FROM         tchp_trxs where sno='" & txtSNo & "' order by transdate ,id "
Set rs = oSaccoMaster.GetRecordset(sql)
With LV
        
        
 
    
        While Not rs.EOF
        
        If Not IsNull(rs.Fields("SNo")) Then
        
            Set li = .ListItems.Add(, , Trim(rs.Fields("transdate")))
            End If
            If Not IsNull(rs.Fields("description")) Then
            li.ListSubItems.Add , , Trim(rs.Fields("description"))
            End If
            If Not IsNull(rs.Fields("Debits")) Then
            li.ListSubItems.Add , , Trim(rs.Fields("Debits"))
            Else
            li.ListSubItems.Add , , "0.00"
            End If
            If Not IsNull(rs.Fields("CreditsD")) Then
            li.ListSubItems.Add , , Trim(rs.Fields("CreditsD"))
            Else
            li.ListSubItems.Add , , "0.00"
            End If
            If Not IsNull(rs.Fields("CreditsC")) Then
             li.ListSubItems.Add , , Trim(rs.Fields("CreditsC"))
             Else
             li.ListSubItems.Add , , "0.00"
            End If
            If Not IsNull(rs.Fields("Balance")) Then
             li.ListSubItems.Add , , Trim(rs.Fields("Balance"))
             Else
             li.ListSubItems.Add , , "0.00"
            End If

            
                    rs.MoveNext
        
        Wend
        
    End With
'//GET THE TOTAL DEDUCTIONS
sql = "SELECT     SUM(Debits) AS A, SUM(CreditsD) AS B, SUM(CreditsC) AS C  FROM         tchp_trxs  WHERE     (sno = '" & txtSNo & "')  GROUP BY sno"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
txtTotalDeductions(0) = rs.Fields(1)
txtTotalCashReceipts = rs.Fields(2)
End If


Exit Sub
ErrorHandler:
MsgBox err.description
End Sub

