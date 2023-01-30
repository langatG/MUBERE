VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmstandingorders 
   BackColor       =   &H00FFFF80&
   Caption         =   "STANDING ORDER SET UP"
   ClientHeight    =   7710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   7440
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkaddnew 
      Caption         =   "Topup Standing Order"
      BeginProperty Font 
         Name            =   "News701 BT"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5400
      TabIndex        =   34
      Top             =   2640
      Width           =   1935
   End
   Begin VB.CommandButton cmdreport3 
      Caption         =   "Pause Standing Order Report"
      Height          =   375
      Left            =   2760
      TabIndex        =   33
      Top             =   7320
      Width           =   2415
   End
   Begin VB.CommandButton cmdrepor2 
      Caption         =   "Complete Stnding order Report"
      Height          =   375
      Left            =   120
      TabIndex        =   32
      Top             =   7320
      Width           =   2535
   End
   Begin VB.CommandButton cmdresume 
      Caption         =   "Resume Standing Order"
      Height          =   375
      Left            =   4080
      TabIndex        =   31
      Top             =   3120
      Width           =   1935
   End
   Begin VB.CheckBox chkinactive 
      Caption         =   "Resume Standing Order"
      BeginProperty Font 
         Name            =   "News701 BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   30
      Top             =   2640
      Width           =   2415
   End
   Begin VB.TextBox txtNETO 
      Height          =   375
      Left            =   1560
      TabIndex        =   27
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Standing Order Reports"
      Height          =   375
      Left            =   5280
      TabIndex        =   26
      Top             =   7320
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF80&
      Caption         =   "Stop Standing Order"
      Height          =   375
      Left            =   4080
      TabIndex        =   25
      Top             =   3120
      Width           =   1935
   End
   Begin VB.CommandButton cmdpostall 
      BackColor       =   &H00FFFF80&
      Caption         =   "Post All Suppliers"
      Height          =   375
      Left            =   6000
      TabIndex        =   24
      Top             =   3120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtmaximumamount 
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """Kshs ""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   0
      EndProperty
      Height          =   285
      Left            =   3840
      TabIndex        =   23
      Text            =   "0"
      Top             =   1560
      Width           =   1575
   End
   Begin VB.PictureBox Picture5 
      Height          =   255
      Left            =   1560
      Picture         =   "frmstandingorders.frx":0000
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   10
      Top             =   720
      Width           =   255
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   2640
      TabIndex        =   8
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Default         =   -1  'True
      Height          =   375
      Left            =   960
      TabIndex        =   6
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00FFFF80&
      Caption         =   "Close"
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox txtSNames 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1800
      TabIndex        =   4
      Top             =   720
      Width           =   5535
   End
   Begin VB.TextBox txtAmount 
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """Kshs ""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   0
      EndProperty
      Height          =   285
      Left            =   2040
      TabIndex        =   3
      Text            =   "0"
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox txtSNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox txtRemarks 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   2160
      Width           =   3135
   End
   Begin VB.ComboBox cboDeductionType 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmstandingorders.frx":02C2
      Left            =   0
      List            =   "frmstandingorders.frx":02DB
      TabIndex        =   0
      Top             =   1560
      Width           =   1815
   End
   Begin MSComCtl2.DTPicker DTPEndDate 
      Height          =   255
      Left            =   5640
      TabIndex        =   11
      Top             =   2280
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
      _Version        =   393216
      Format          =   122617857
      CurrentDate     =   40096
   End
   Begin MSComCtl2.DTPicker DTPStartDate 
      Height          =   255
      Left            =   5640
      TabIndex        =   12
      Top             =   1560
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
      _Version        =   393216
      Format          =   122617857
      CurrentDate     =   40096
   End
   Begin MSComCtl2.DTPicker DTPDDeduction 
      Height          =   255
      Left            =   5760
      TabIndex        =   13
      Top             =   360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      _Version        =   393216
      Format          =   122617857
      CurrentDate     =   40096
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3735
      Left            =   120
      TabIndex        =   29
      Top             =   3600
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   6588
      View            =   3
      MultiSelect     =   -1  'True
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
      BackColor       =   65280
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "NETPAY:"
      BeginProperty Font 
         Name            =   "News706 BT"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   28
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFF80&
      Caption         =   "Maximum Amount"
      BeginProperty Font 
         Name            =   "News701 BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   22
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      Caption         =   "End Date"
      BeginProperty Font 
         Name            =   "News701 BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   5640
      TabIndex        =   21
      Top             =   2040
      Width           =   960
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      Caption         =   "Start Date"
      BeginProperty Font 
         Name            =   "News701 BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   5640
      TabIndex        =   20
      Top             =   1320
      Width           =   1065
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      Caption         =   "Type of Deduction"
      BeginProperty Font 
         Name            =   "News701 BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   0
      TabIndex        =   19
      Top             =   1200
      Width           =   1860
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      Caption         =   "Supplier Name"
      BeginProperty Font 
         Name            =   "News701 BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   2520
      TabIndex        =   18
      Top             =   360
      Width           =   1500
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      Caption         =   "Date of Standing Order"
      BeginProperty Font 
         Name            =   "News701 BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   4920
      TabIndex        =   17
      Top             =   120
      Width           =   2385
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      Caption         =   "Supplier Numer"
      BeginProperty Font 
         Name            =   "News701 BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   0
      TabIndex        =   16
      Top             =   360
      Width           =   1605
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      Caption         =   "Amount"
      BeginProperty Font 
         Name            =   "News701 BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   2280
      TabIndex        =   15
      Top             =   1200
      Width           =   825
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      Caption         =   "Remarks"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   14
      Top             =   2160
      Width           =   990
   End
End
Attribute VB_Name = "frmstandingorders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs1 As New ADODB.Recordset, rs2 As New ADODB.Recordset, rs3 As New ADODB.Recordset
Dim rs4 As New ADODB.Recordset
Dim r As Integer
Dim myclass As cdbase
Dim Transport As Currency, agrovet As Currency, AI As Currency, TMShares As Currency, FSA As Currency, HShares As Currency, Advance As Currency, Others As Currency

Private Sub chkinactive_Click()
If chkinactive = vbChecked Then
cmdresume.Visible = True
Command1.Visible = False
Else
cmdresume.Visible = False
Command1.Visible = True
End If
End Sub

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmddelete_Click()
Dim DESCR As String

DESCR = cboDeductionType.Text

 If txtSNo = "" Then
   MsgBox "Please select the supply number", vbInformation
 Exit Sub
 Else
 End If
  If cboDeductionType = "" Then
   MsgBox "Please select Type of Deduction", vbInformation
 Exit Sub
 Else
 End If


'//check if supplier is in standing order table.
sql = ""
sql = "set dateformat dmy select * from d_supplier_standingorder where sno='" & txtSNo & "'  And description = '" & DESCR & "'"
Set rst = oSaccoMaster.GetRecordset(sql)
If Not rst.EOF Then
'//Update deductions
    Set cn = New ADODB.Connection
'    sql = "set dateformat dmy INSERT INTO d_supplier_standingorder"
'    sql = sql & "           (SNo, Date_Deduc, Description, Amount, MaxAmount, Period, StartDate, EndDate, auditid,  yyear, Remarks)"
'    sql = sql & "  VALUES     (" & txtSNo & ",'" & DTPDDeduction & "','" & DESCR & "'," & txtamount & "," & txtmaximumamount & ",'" & Format(DTPDDeduction, "mmm-YYYY") & "','" & DTPStartDate & "','" & DTPEndDate & "','" & User & "'," & year(DTPEndDate) & ",'" & txtRemarks & "')"
'    oSaccoMaster.ExecuteThis (sql)
    sql = ""
    sql = "delete from d_supplier_standingorder where sno='" & txtSNo & "' And description = '" & DESCR & "'"
    oSaccoMaster.ExecuteThis (sql)
    'cn.Execute sql

    Else
    MsgBox "You have not make any standingorder deduction with this member - " & txtSNo & "", vbInformation, "Standing Order Set Up"
    Exit Sub
End If


txtamount = ""
txtSNo = ""
txtSNo_Validate True

txtSNo.SetFocus
'Form_Load
MsgBox "Records successively Deleted."
loadBranchesTypes
Exit Sub

End Sub

Private Sub cmdedit_Click()
r = 1
End Sub

Private Sub cmdNew_Click()
txtamount = ""
txtSNames = ""
txtSNo = ""
cboDeductionType = ""

txtamount.Locked = False
txtSNo.Locked = False
cboDeductionType.Locked = False

cmdsave.Enabled = True
cmdnew.Enabled = False
r = 0
DTPDDeduction = Format(Get_Server_Date, "dd/mm/yyyy")
dtpStartDate = Format(Get_Server_Date, "dd/mm/yyyy")
DTPEndDate = DateSerial(year(dtpStartDate), month(dtpStartDate) + 1, 1 - 1)

End Sub

Private Sub cmdpostall_Click()
Dim ans As String
Dim NetP As Currency
Dim rshast As New ADODB.Recordset
Dim rs As New ADODB.Recordset

Dim DESCR As String

'DESCR = cboDeductionType.Text

Dim D As Date
Dim sno As Long
Dim kgs As Double
Startdate = DateSerial(year(DTPDDeduction), month(DTPDDeduction), 1)
D = Startdate + 9
Enddate = DateSerial(year(DTPDDeduction), month(DTPDDeduction) + 1, 1 - 1)
 Set rshast = oSaccoMaster.GetRecordset("select sno from d_suppliers order by sno")
While Not rshast.EOF
sno = rshast.Fields(0)
'check active
Set rs = New ADODB.Recordset
sql = " set dateformat dmy select  isnull(sum(qsupplied),0)qsupplied from d_milkintake  where sno='" & sno & "' and transdate>='" & dtpStartDate & "' and transdate<='" & DTPEndDate & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
kgs = rs!QSupplied
End If
'//check if it had another deductions of the same nature.
sql = ""
sql = "select description from d_supplier_deduc where sno='" & sno & "' and description ='shares' and date_deduc='" & D & "'"
Set rst = oSaccoMaster.GetRecordset(sql)
If rst.EOF And kgs >= 1 Then
Set cn = New ADODB.Connection
sql = "d_sp_SupplierDeduct " & sno & ",'" & D & "','shares','200','" & dtpStartDate & "','" & DTPEndDate & "','" & year(DTPEndDate) & "','" & User & "','shares','head office'"
oSaccoMaster.ExecuteThis (sql)

'strSQL = "set dateformat dmy insert into [d_Shares]([IdNo],[SNO],[Code],[Name],[Sex],[Loc],[Type],[TransDate],[pmode],[Period],[Amnt],[amount],[AuditId], [AuditDateTime])"
'strSQL = strSQL & " values( '0','" & sno & "','" & sno & "','checkoff','other','olenguruone','shares','"
'strSQL = strSQL & Enddate & "',' 0','" & Enddate & "','200','200','" & User & "','" & Get_Server_Date & "')"
'oSaccoMaster.ExecuteThis (strSQL)

sql = ""
sql = "set dateformat dmy insert into d_sconribution(sno, transdate, amount, bal, transdescription, auditid)"
sql = sql & " values('" & sno & "','" & Enddate & "','200','200','shares','" & User & "') "
oSaccoMaster.ExecuteThis (sql)
End If
'Else
rshast.MoveNext
Wend
MsgBox "Records Successfully Updated", vbInformation
End Sub

Public Sub loadBranchesTypes()
    
    With ListView1
    
        .ListItems.Clear
        
        .ColumnHeaders.Clear
    
    End With

    Set rs2 = CreateObject("adodb.recordset")
    
    sql = "set dateformat dmy Select d.SNO,m.Names, d.Date_Deduc, d.Description, d.Amount, d.MaxAmount,d.Remarks from d_supplier_standingorder AS d INNER JOIN d_Suppliers AS m ON d.SNO = m.SNO  WHERE d.Active='0' order by d.Date_Deduc desc"
'    sql = ""
'    sql = "set dateformat dmy SELECT d.RefNo,m.DName, d.DispDate, d.DispQnty,d.Amount,d.PaidAmount FROM d_MilkControl AS d INNER JOIN d_Debtors AS m ON d.DCode = m.DCode WHERE     (DispDate = '" & txtdateenterered & "') and vehicleno='" & cboVehicle & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    
    Set rs2 = CreateObject("adodb.recordset")
    
    Set clsClass = New cdbase
    
    Provider = clsClass.OpenCon
    
    Set cn = CreateObject("adodb.connection")
    
   cn.Open Provider, "atm", "atm"
    
    rs2.Open sql, cn
    
    With ListView1
        .ColumnHeaders.Add , , "SNo"
        .ColumnHeaders.Add , , "Name"
        .ColumnHeaders.Add , , "Date"
        .ColumnHeaders.Add , , "Description"
        .ColumnHeaders.Add , , "Amount"
        .ColumnHeaders.Add , , "MaxAmount"
        .ColumnHeaders.Add , , "Remarks"
'        .ColumnHeaders.Add , , "Outlet"
        While Not rs2.EOF
        
            Set li = .ListItems.Add(, , Trim(rs2.Fields("SNO")))
            li.ListSubItems.Add , , Trim(rs2.Fields("Names"))
            li.ListSubItems.Add , , Trim(rs2.Fields("Date_Deduc"))
            li.ListSubItems.Add , , Trim(rs2.Fields("Description"))
            li.ListSubItems.Add , , Trim(rs2.Fields("Amount"))
            li.ListSubItems.Add , , Trim(rs2.Fields("MaxAmount"))
            li.ListSubItems.Add , , Trim(rs2.Fields("Remarks"))
            rs2.MoveNext
        
        Wend
        
    End With
    
    rs2.Close
    
    Set rs2 = Nothing
    
ListView1.View = lvwReport
End Sub

Private Sub cmdrepor2_Click()
reportname = "StandingOrdersCOM.rpt"
Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub cmdreport3_Click()
reportname = "StandingOrdersSTOP.rpt"
Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub cmdresume_Click()
Dim DESCR As String

DESCR = cboDeductionType.Text
sql = ""
sql = "select description from d_supplier_standingorder where sno='" & txtSNo & "' and description ='" & DESCR & "' and active=1 and complete=0 and Date_stop IS NULL"
Set rst = oSaccoMaster.GetRecordset(sql)
If Not rst.EOF Then
sql = ""
sql = "update  d_supplier_standingorder set active=0 where sno='" & txtSNo & "' and description='" & Trim(DESCR) & "' and complete=0 and Date_stop IS NULL"
'update  d_supplier_standingorder set active=1 where sno=5 and description='CBO'
oSaccoMaster.ExecuteThis (sql)
Else
MsgBox "Supplier " & txtSNo & " is not in Standing Order Stop List", vbInformation, "Standing Order Set Up."
Exit Sub
End If
MsgBox "Standing Order Successfully Resume", vbInformation, "Standing Order Set Up."
txtamount = ""
txtSNames = ""
txtSNo = ""
cboDeductionType = ""
loadBranchesTypes
End Sub

Private Sub cmdsave_Click()
On Error GoTo ErrorHandler

'//Validation

Dim ans As String
Dim NetP As Currency

Startdate = DateSerial(year(DTPDDeduction), month(DTPDDeduction), 1)
Enddate = DateSerial(year(DTPDDeduction), month(DTPDDeduction) + 1, 1 - 1)

Dim DESCR As String

DESCR = cboDeductionType.Text

Startdate = DateSerial(year(DTPDDeduction), month(DTPDDeduction), 1)
Enddate = DateSerial(year(DTPDDeduction), month(DTPDDeduction) + 1, 1 - 1)

'check if the supplier is bringing milk
Dim rss As New Recordset
Set rss = oSaccoMaster.GetRecordset("select sno from d_Milkintake where sno=" & txtSNo & "")
If Not rss.EOF Then


Set rs = oSaccoMaster.GetRecordset("d_sp_SupNet " & txtSNo & ",'" & Startdate & "','" & Enddate & "', 0")

'If Not rs.EOF Then

If Not IsNull(rs.Fields(1)) Then
NetP = rs.Fields(1)
Else
NetP = "0.00"
End If

Set rs = oSaccoMaster.GetRecordset("d_sp_SupNet " & txtSNo & ",'" & Startdate & "','" & Enddate & "', 1")
If Not IsNull(rs.Fields(0)) Then
NetP = NetP - rs.Fields(0)
Else
NetP = NetP - 0
End If
'End With
End If
If NetP < CCur(txtamount) Then
ans = MsgBox("The supplier number " & txtSNo & " has; " & vbNewLine & "Gross pay of " & Format((NetP + rs.Fields(0)), "#,##0.00") & vbNewLine & " Total Deductions " & Format(rs.Fields(0), "#,##0.00") & vbNewLine & "NetPay " & Format(NetP, "#,##0.00") & "." & vbNewLine & "Continue anyway?", vbYesNo, "LESS NET AMOUNT")

If ans = vbNo Then
Exit Sub
End If
End If

'Else

Set rs = oSaccoMaster.GetRecordset("d_sp_IsClosed '" & Enddate & "'")
If Not rs.EOF Then
    MsgBox "The period " & Enddate & " has been closed by " & rs.Fields(0)
    Exit Sub
End If

'//check if it had another deductions of the same nature.
If chkaddnew = vbUnchecked Then
    sql = ""
    sql = "select description from d_supplier_standingorder where sno='" & txtSNo & "' and description ='" & DESCR & "' and Active='0' and topup='0'"
    Set rst = oSaccoMaster.GetRecordset(sql)
    If rst.EOF Then
    '//Update deductions
        Set cn = New ADODB.Connection
        sql = "set dateformat dmy INSERT INTO d_supplier_standingorder"
        sql = sql & "           (SNo, Date_Deduc, Description, Amount, MaxAmount, Period, StartDate, EndDate, auditid,  yyear, Remarks,topup)"
        sql = sql & "  VALUES     (" & txtSNo & ",'" & DTPDDeduction & "','" & DESCR & "'," & txtamount & "," & txtmaximumamount & ",'" & Format(DTPDDeduction, "mmm-YYYY") & "','" & dtpStartDate & "','" & DTPEndDate & "','" & User & "'," & year(DTPEndDate) & ",'" & txtremarks & "','0')"
        oSaccoMaster.ExecuteThis (sql)

     Else
        '''''check if it is to be edited
        If r = 1 Then
          sql = ""
          sql = "update  d_supplier_standingorder set active=1,complete=1,MaxAmount=" & txtmaximumamount & " where sno='" & txtSNo & "' and description='" & DESCR & "'and active='0' and topup='0'"
          oSaccoMaster.ExecuteThis (sql)
        ''''end
        Else
         MsgBox "The Deduction Code Has Been Defined for this Member " & txtSNo & "", vbInformation, "Standing Order Set Up"
         Exit Sub
        End If
    End If
Else
    Set cn = New ADODB.Connection
    sql = "set dateformat dmy INSERT INTO d_supplier_standingorder"
    sql = sql & "           (SNo, Date_Deduc, Description, Amount, MaxAmount, Period, StartDate, EndDate, auditid,  yyear, Remarks,topup)"
    sql = sql & "  VALUES     (" & txtSNo & ",'" & DTPDDeduction & "','" & DESCR & "'," & txtamount & "," & txtmaximumamount & ",'" & Format(DTPDDeduction, "mmm-YYYY") & "','" & dtpStartDate & "','" & DTPEndDate & "','" & User & "'," & year(DTPEndDate) & ",'" & txtremarks & "','1')"
    oSaccoMaster.ExecuteThis (sql)
End If
txtamount = ""
txtSNo = ""
r = 0
cmdEdit.Enabled = False
txtSNo_Validate True
txtmaximumamount = ""
txtamount = ""
txtSNo.SetFocus
'Form_Load
MsgBox "Records successively updated."
loadBranchesTypes
Exit Sub

'End If

ErrorHandler:
MsgBox err.description

End Sub

Private Sub Command1_Click()
Dim DESCR As String

DESCR = cboDeductionType.Text
sql = ""
sql = "select description from d_supplier_standingorder where sno='" & txtSNo & "' and description ='" & DESCR & "' and active=0"
Set rst = oSaccoMaster.GetRecordset(sql)
If Not rst.EOF Then
sql = ""
sql = "update  d_supplier_standingorder set active=1,complete=0 where sno='" & txtSNo & "' and description='" & Trim(DESCR) & "'and active=0"
'update  d_supplier_standingorder set active=1 where sno=5 and description='CBO'
oSaccoMaster.ExecuteThis (sql)
Else
MsgBox "Supplier " & txtSNo & " is not in Standing Order List", vbInformation, "Standing Order Set Up."
Exit Sub
End If
MsgBox "Standing Order Successfully Stopped", vbInformation, "Standing Order Set Up."
txtamount = ""
txtSNames = ""
txtSNo = ""
cboDeductionType = ""
loadBranchesTypes
End Sub

Private Sub Command2_Click()
reportname = "StandingOrders.rpt"
Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub Command3_Click()
Dim DESCR As String

DESCR = cboDeductionType.Text
sql = ""
sql = "select description from d_supplier_standingorder where sno='" & txtSNo & "' and description ='" & DESCR & "' and active=0"
Set rst = oSaccoMaster.GetRecordset(sql)
If Not rst.EOF Then
sql = ""
sql = "update  d_supplier_standingorder set active=1 where sno='" & txtSNo & "' and description='" & Trim(DESCR) & "'"
'update  d_supplier_standingorder set active=1 where sno=5 and description='CBO'
oSaccoMaster.ExecuteThis (sql)
Else
MsgBox "Supplier " & txtSNo & " is not in Standing Order List", vbInformation, "Standing Order Set Up."
Exit Sub
End If
MsgBox "Standing Order Successfully Stopped", vbInformation, "Standing Order Set Up."
End Sub

Private Sub DTPDDeduction_Change()
dtpStartDate = DateSerial(year(DTPDDeduction), month(DTPDDeduction), 1)
DTPEndDate = DateSerial(year(DTPDDeduction), month(DTPDDeduction) + 1, 1 - 1)
loadBranchesTypes
End Sub

Private Sub Form_Load()
txtamount = ""
txtSNames = ""
txtSNo = ""
txtremarks = "STANDING ORDER"
loadBranchesTypes
cboDeductionType = ""

txtamount.Locked = True
txtSNames.Locked = True
txtSNo.Locked = True
cboDeductionType.Locked = True

cmdnew.Enabled = True
cmdsave.Enabled = False
cmdEdit.Enabled = False
cmdDelete.Enabled = True
cmdresume.Visible = False

DTPDDeduction = Format(Get_Server_Date, "dd/mm/yyyy")
dtpStartDate = DateSerial(year(DTPDDeduction), month(DTPDDeduction), 1)
'DTPStartDate = Format(Get_Server_Date, "dd/mm/yyyy")
DTPEndDate = DateSerial(year(dtpStartDate), month(dtpStartDate) + 1, 1 - 1)

    cboDeductionType.Clear
    Set myclass = New cdbase

    Provider = myclass.OpenCon

    Set cn = CreateObject("adodb.connection")

   cn.Open Provider, "atm", "atm"

    Set rs = CreateObject("adodb.recordset")

    rs.Open "SELECT Description FROM d_DCodes order by 1 ", cn

    If rs.EOF Then Exit Sub

    With rs

        While Not .EOF

         cboDeductionType.AddItem rs.Fields("Description")

         .MoveNext

        Wend

    End With
'///////////////// for active suppliers/////////////////
''''sql = ""
''''sql = "select SNO,Active,Date_Deduc from d_supplier_standingorder where Active='0'"
''''Set rst = oSaccoMaster.GetRecordset(sql)
Dim C As Integer
Dim M As Double
''''While Not rst.EOF
''''C = rst.Fields(0)
''''sql = ""
''''sql = "set dateformat dmy select SUM(Amount)as m from  d_supplier_deduc where SNo='" & C & "' and Remarks like'%STANDING ORDER%' and Date_Deduc>='" & rst.Fields(2) & "'"
''''Set rs = oSaccoMaster.GetRecordset(sql)
''''If rs.Fields(0) <> "" Then
''''M = rs.Fields(0)
''''sql = ""
''''sql = "update d_supplier_standingorder set Deducted='" & M & "' where sno='" & C & "' and Active='0'"
''''oSaccoMaster.ExecuteThis (sql)
''''Else
''''sql = ""
''''sql = "update d_supplier_standingorder set Deducted='0' where sno='" & C & "' and Active='0'"
''''oSaccoMaster.ExecuteThis (sql)
''''End If
''''  rst.MoveNext
''''Wend
'////////////////end

'''''//////////////////suppliers WHO HAVE COMPLETED THEIR STANDING////////////////
''''sql = ""
''''sql = "select SNO,Active,Date_Deduc,Date_stop from d_supplier_standingorder where Active='1' AND complete='1'"
''''Set rsv = oSaccoMaster.GetRecordset(sql)
''''''Dim C As Integer
''''''Dim M As Double
''''While Not rsv.EOF
''''C = rsv.Fields(0)
''''sql = ""
''''sql = "set dateformat dmy select SUM(Amount)as m from  d_supplier_deduc where SNo='" & C & "' and Remarks like'%STANDING ORDER%' and Date_Deduc>='" & rsv.Fields(2) & "' and Date_Deduc<='" & rsv.Fields(3) & "' "
''''Set rsg = oSaccoMaster.GetRecordset(sql)
''''If rsg.Fields(0) <> "" Then
''''M = rsg.Fields(0)
''''sql = ""
''''sql = "update  d_supplier_standingorder set Deducted='" & M & "' where sno='" & C & "' and Active='1' AND complete='1'"
''''oSaccoMaster.ExecuteThis (sql)
''''Else
''''sql = ""
''''sql = "update  d_supplier_standingorder set Deducted='0' where sno='" & C & "'"
''''oSaccoMaster.ExecuteThis (sql)
''''End If
''''  rsv.MoveNext
''''Wend
'''''////////////////end
'//////////////////suppliers WHO HAVE BEEN STOP THEIR STANDING////////////////
sql = ""
sql = "select SNO,Active,Date_Deduc,Date_stop,MaxAmount from d_supplier_standingorder where Active='0' AND complete='0' AND topup='0'"
Set rsp = oSaccoMaster.GetRecordset(sql)
Dim g As Double
Dim k As Double
While Not rsp.EOF
C = rsp.Fields(0)
k = rsp.Fields(4)
sql = ""
sql = "set dateformat dmy select isnull(SUM(Amount),0)as m from  d_supplier_deduc where SNo='" & C & "' and Remarks like'%STANDING ORDER%' and Date_Deduc>='" & rsp.Fields(2) & "' "
Set rsgm = oSaccoMaster.GetRecordset(sql)
If Not rsgm.EOF Then
 M = rsgm.Fields(0)
 g = k - M
 If g <= 0 Then
  sql = ""
  sql = "update  d_supplier_standingorder set Deducted='" & M & "', Active='1',complete='1' where sno='" & C & "' and Active='0' AND complete='0' AND topup='0'"
  oSaccoMaster.ExecuteThis (sql)
'''''''''check if its in new standing order allocated
   sql = ""
   sql = "set dateformat dmy select * from  d_supplier_standingorder where SNo='" & C & "' and Remarks like'%STANDING ORDER%' and topup='1' "
   Set rsgm = oSaccoMaster.GetRecordset(sql)
   If Not rsgm.EOF Then
    sql = ""
    sql = "update  d_supplier_standingorder set topup='0' where sno='" & C & "' and Active='0' AND complete='0' AND topup='1'"
    oSaccoMaster.ExecuteThis (sql)
   End If
'''''''''end
 Else
  sql = ""
  sql = "update  d_supplier_standingorder set Deducted='" & M & "' where sno='" & C & "' and Active='0' AND complete='0' AND topup='0'"
  oSaccoMaster.ExecuteThis (sql)
 End If
Else
sql = ""
sql = "update  d_supplier_standingorder set Deducted='0' where sno='" & C & "' AND topup='0'"
oSaccoMaster.ExecuteThis (sql)
End If
  rsp.MoveNext
Wend
'////////////////end

End Sub

Private Sub Form_LostFocus()
'txtAmount.DataFormat = FormatCurrency("'Kshs '#,##0.00", Val(txtAmount))
End Sub

Private Sub Form_Unload(Cancel As Integer)
'oSaccoMaster.ExecuteThis ("d_sp_GDedNet '" & DTPStartDate & "', '" & DTPEndDate & "'")
End Sub

Private Sub listview1_DblClick()
txtamount = ListView1.SelectedItem.SubItems(4)
txtmaximumamount = ListView1.SelectedItem.SubItems(5)
txtremarks = ListView1.SelectedItem.SubItems(6)
cboDeductionType = ListView1.SelectedItem.SubItems(3)
txtSNo.Text = ListView1.SelectedItem
cmdEdit.Enabled = True
txtSNo_Validate True
End Sub

Private Sub Picture5_Click()
        Me.MousePointer = vbHourglass
        frmSearchSupplier.Show vbModal
        txtSNo = sel
        txtSNo_Validate True
        Me.MousePointer = 0
End Sub

Private Sub txtSNo_Validate(Cancel As Boolean)
Dim ans As String
Dim NetP As Currency
Startdate = DateSerial(year(DTPDDeduction), month(DTPDDeduction), 1)
Enddate = DateSerial(year(DTPDDeduction), month(DTPDDeduction) + 1, 1 - 1)
If txtSNo = "" Then Exit Sub
Set rs = oSaccoMaster.GetRecordset("d_sp_SupNet '" & txtSNo & "','" & Startdate & "','" & Enddate & "',0")
If Not rs.EOF Then
If Not IsNull(rs.Fields(1)) Then
If Not IsNull(rs.Fields(1)) Then
NetP = IIf(IsNull(rs.Fields(1)), 0, rs.Fields(1))
Else
NetP = "0.00"
End If

Set rs = oSaccoMaster.GetRecordset("d_sp_SupNet '" & txtSNo & "','" & Startdate & "','" & Enddate & "', 1")
If Not IsNull(rs.Fields(0)) Then
NetP = NetP - rs.Fields(0)
'- rs.Fields(0)
txtNETO.Text = NetP
txtNETO.Text = NetP
Else
'txtnetp.Tex = NetP
NetP = NetP
txtNETO.Text = NetP
'- 0
End If
End If
End If



Dim a, t As Boolean
Set rs = New ADODB.Recordset
sql = "d_sp_SelectSuppliers '" & txtSNo & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
If Not IsNull(rs.Fields(0)) Then txtSNames = rs.Fields(2)
Else
txtSNames = ""
End If
End Sub

