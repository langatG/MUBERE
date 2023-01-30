VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmFarmerDeductionAssign 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFF80&
   Caption         =   "Assign Deductions To the Farmer"
   ClientHeight    =   7920
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   9060
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtnetp 
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   24
      Text            =   "0"
      Top             =   2160
      Width           =   1335
   End
   Begin VB.ComboBox cbobranches 
      Height          =   315
      ItemData        =   "frmFarmerDeductionAssign.frx":0000
      Left            =   120
      List            =   "frmFarmerDeductionAssign.frx":0002
      TabIndex        =   22
      Top             =   2400
      Width           =   3015
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
      ItemData        =   "frmFarmerDeductionAssign.frx":0004
      Left            =   120
      List            =   "frmFarmerDeductionAssign.frx":0006
      TabIndex        =   21
      Top             =   1560
      Width           =   2175
   End
   Begin VB.TextBox txtRemarks 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   5760
      TabIndex        =   20
      Top             =   2760
      Width           =   3135
   End
   Begin VB.TextBox txtSNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "+++++++++++++++++++++++++++++++++++++++"
      Top             =   720
      Width           =   1575
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
      Left            =   3000
      TabIndex        =   2
      Text            =   "0"
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox txtSNames 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1920
      TabIndex        =   10
      Top             =   720
      Width           =   5775
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00FFFF80&
      Caption         =   "Close"
      Height          =   495
      Left            =   4440
      TabIndex        =   9
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   495
      Left            =   1200
      TabIndex        =   8
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   495
      Left            =   3360
      TabIndex        =   7
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   3240
      Width           =   975
   End
   Begin VB.PictureBox Picture5 
      Height          =   255
      Left            =   1680
      Picture         =   "frmFarmerDeductionAssign.frx":0008
      ScaleHeight     =   195
      ScaleWidth      =   315
      TabIndex        =   0
      Top             =   720
      Width           =   375
   End
   Begin MSComCtl2.DTPicker DTPEndDate 
      Height          =   375
      Left            =   7320
      TabIndex        =   4
      Top             =   1560
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   119668737
      CurrentDate     =   40096
   End
   Begin MSComCtl2.DTPicker DTPStartDate 
      Height          =   375
      Left            =   5280
      TabIndex        =   5
      Top             =   1560
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   119668737
      CurrentDate     =   40096
   End
   Begin MSComCtl2.DTPicker DTPDDeduction 
      Height          =   255
      Left            =   5880
      TabIndex        =   11
      Top             =   240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      _Version        =   393216
      Format          =   119668737
      CurrentDate     =   40096
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3855
      Left            =   120
      TabIndex        =   26
      Top             =   3840
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   6800
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
      BackColor       =   &H00FFFF80&
      Caption         =   "NET  PAY"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   25
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      Caption         =   "Branches"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   240
      TabIndex        =   23
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      Caption         =   "Remarks"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4440
      TabIndex        =   19
      Top             =   2760
      Width           =   1050
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      Caption         =   "Amount"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   3000
      TabIndex        =   18
      Top             =   1200
      Width           =   945
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      Caption         =   "Supplier Numer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   120
      TabIndex        =   17
      Top             =   360
      Width           =   1875
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      Caption         =   "Date of deduction"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   5880
      TabIndex        =   16
      Top             =   0
      Width           =   2085
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      Caption         =   "Supplier Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   3120
      TabIndex        =   15
      Top             =   360
      Width           =   1755
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      Caption         =   "Type of Deduction"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   120
      TabIndex        =   14
      Top             =   1200
      Width           =   2160
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      Caption         =   "Start Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   5280
      TabIndex        =   13
      Top             =   1320
      Width           =   1065
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      Caption         =   "End Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   7320
      TabIndex        =   12
      Top             =   1320
      Width           =   960
   End
End
Attribute VB_Name = "frmFarmerDeductionAssign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs1 As New ADODB.Recordset, rs2 As New ADODB.Recordset, rs3 As New ADODB.Recordset
Dim rs4 As New ADODB.Recordset
Dim myclass As cdbase
Dim Transport As Currency, agrovet As Currency, BONUS As Currency, TMShares As Currency, FSA As Currency, HShares As Currency, Advance As Currency, Others As Currency

Private Sub cboDeductionType_Click()
txtremarks.Text = cboDeductionType
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
sql = "set dateformat dmy select * from d_supplier_deduc where SNo='" & txtSNo & "'  And Description = '" & DESCR & "' And Date_Deduc = '" & DTPDDeduction & "'"
Set rst = oSaccoMaster.GetRecordset(sql)
If Not rst.EOF Then
'//Update deductions
    Set cn = New ADODB.Connection
'    sql = "set dateformat dmy INSERT INTO d_supplier_standingorder"
'    sql = sql & "           (SNo, Date_Deduc, Description, Amount, MaxAmount, Period, StartDate, EndDate, auditid,  yyear, Remarks)"
'    sql = sql & "  VALUES     (" & txtSNo & ",'" & DTPDDeduction & "','" & DESCR & "'," & txtamount & "," & txtmaximumamount & ",'" & Format(DTPDDeduction, "mmm-YYYY") & "','" & DTPStartDate & "','" & DTPEndDate & "','" & User & "'," & year(DTPEndDate) & ",'" & txtRemarks & "')"
'    oSaccoMaster.ExecuteThis (sql)
    sql = ""
    sql = "delete from d_supplier_deduc where sno='" & txtSNo & "' And Description = '" & DESCR & "' And Date_Deduc = '" & DTPDDeduction & "'"
    oSaccoMaster.ExecuteThis (sql)
    'cn.Execute sql

    Else
    MsgBox "You have not make any Deduction with this supplier " & txtSNo & " on Date " & DTPDDeduction & " ", vbInformation
    Exit Sub
End If

txtamount = ""
txtSNames = ""
txtSNo = ""
cboDeductionType = ""
txtSNo_Validate True

txtSNo.SetFocus
'Form_Load
MsgBox "Records successively Deleted."
Exit Sub
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
cmdDelete.Enabled = True
DTPDDeduction = Format(Get_Server_Date, "dd/mm/yyyy")
dtpStartDate = Format(Get_Server_Date, "dd/mm/yyyy")
DTPEndDate = DateSerial(Year(dtpStartDate), month(dtpStartDate) + 1, 1 - 1)

End Sub

Private Sub cmdsave_Click()
On Error GoTo ErrorHandler

'//Validation
Set cn = New ADODB.Connection
sql = "SET dateformat dmy SELECT NPay,GPay,TDeductions from d_Payroll WHERE SNo=" & txtSNo & " AND EndofPeriod ='" & DateSerial(Year(DTPDDeduction), month(DTPDDeduction) + 1, 1 - 1) & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
Dim ans As String
Dim NetP As Currency
If rs.RecordCount > 0 Then
 If Not IsNull(rs.Fields("NPay")) Then
  NetP = rs.Fields("NPay")
 End If
 If Not IsNull(rs.Fields("GPay")) Then
  NetP = rs.Fields("GPay")
 End If

Startdate = DateSerial(Year(DTPDDeduction), month(DTPDDeduction), 1)
Enddate = DateSerial(Year(DTPDDeduction), month(DTPDDeduction) + 1, 1 - 1)
If txtSNo = "" Then Exit Sub
Set rs = oSaccoMaster.GetRecordset("d_sp_SupNet " & txtSNo & ",'" & Startdate & "','" & Enddate & "', 0")
If Not rs.EOF Then
 If Not IsNull(rs.Fields(1)) Then
  If Not IsNull(rs.Fields(1)) Then
   NetP = IIf(IsNull(rs.Fields(1)), 0, rs.Fields(1))
  Else
   NetP = "0.00"
  End If

 Set rs = oSaccoMaster.GetRecordset("d_sp_SupNet " & txtSNo & ",'" & Startdate & "','" & Enddate & "', 1")
 If Not IsNull(rs.Fields(0)) Then
  NetP = NetP - rs.Fields(0)
 Else
  NetP = NetP - 0
 End If
 End If

 If NetP < CCur(txtamount) Then
    ans = MsgBox("The supplier number " & txtSNo & " has; " & vbNewLine & "Gross pay of " & Format((NetP + rs.Fields(0)), "#,##0.00") & vbNewLine & " Total Deductios " & Format(rs.Fields(0), "#,##0.00") & vbNewLine & "NetPay " & Format(NetP, "#,##0.00") & "." & vbNewLine & "You Can Continue anyway?", vbQuestion + vbYesNo, "LESS NET AMOUNT")
    
    If ans = vbNo Then
     Exit Sub
    End If
    
    If ans = vbYes Then
     MsgBox "Please let the supplier apply an amount less or equal to " & Format(NetP, "#,##0.00") & ""
     txtamount.SetFocus
     Exit Sub
    End If
 End If
End If
Else
 MsgBox "There is no record for supplier number " & txtSNo & " for period ending " & DateSerial(Year(DTPDDeduction), month(DTPDDeduction) + 1, 1 - 1) & ""
 txtSNo.SetFocus
 Exit Sub
End If


If cboDeductionType.Text = "Others" And txtremarks = "" Then
MsgBox "Please enter the remarks."
txtremarks.SetFocus
Exit Sub
End If
If txtremarks = "" Then
txtremarks = " "
End If
If cbobranches = "" Then
MsgBox "Please select the Branch"
Exit Sub
End If
If txtremarks = "" Then
txtremarks = " "
End If
Dim DESCR As String

DESCR = cboDeductionType.Text

'If Trim(cboDeductionType.Text) = "Shares" Then
'DESCR = "HShares"
'End If
'If Trim(cboDeductionType.Text) = "Registration" Then
'DESCR = "TMShares"
'End If

Startdate = DateSerial(Year(DTPDDeduction), month(DTPDDeduction), 1)
Enddate = DateSerial(Year(DTPDDeduction), month(DTPDDeduction) + 1, 1 - 1)

Set rs = oSaccoMaster.GetRecordset("d_sp_IsClosed '" & Enddate & "'")
If Not rs.EOF Then
    MsgBox "The period " & Enddate & " has been closed by " & rs.Fields(0)
    Exit Sub
End If
'//Update deductions
Set cn = New ADODB.Connection
sql = "d_sp_SupplierDeduct " & txtSNo & ",'" & DTPDDeduction & "','" & DESCR & "'," & txtamount & ",'" & dtpStartDate & "','" & DTPEndDate & "'," & Year(DTPEndDate) & ",'" & User & "','" & txtremarks & "','" & cbobranches & "'"
oSaccoMaster.ExecuteThis (sql)

'UPDATE Shares Chekoff
If UCase$(DESCR) = "SHARES" Then
Set rs = New ADODB.Recordset
sql = "d_sp_SelectSuppliers '" & txtSNo & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
With rs
 If Not rs.EOF Then
 Dim idno As String, sex As String, Location As String
 idno = IIf(IsNull(!idno), "", !idno)
 sex = IIf(IsNull(!Type), "", !Type)
 Location = IIf(IsNull(!Location), "", !Location)
 End If
End With
sex = Left(sex, 1)
strSQL = "set dateformat dmy insert into [d_Shares]([IdNo],[SNO],[Code],[Name],[Sex],[Loc],[Type],[TransDate],[pmode],[Period],[Amnt],[amount],[AuditId], [AuditDateTime])"
strSQL = strSQL & " values( '" & Trim$(idno) & "','" & txtSNo & "','" & txtSNo & "','" & txtSNames & "','" & sex & "','" & Location & "','" & DESCR & "','"
strSQL = strSQL & Enddate & "',' 0','" & Enddate & "'," & txtamount & "," & txtamount & ",'" & User & "','" & Get_Server_Date & "')"
oSaccoMaster.ExecuteThis (strSQL)

sql = ""
sql = "set dateformat dmy insert into d_sconribution(sno, transdate, amount, bal, transdescription, auditid)"
sql = sql & " values('" & txtSNo & "','" & Enddate & "'," & txtamount & "," & txtamount & ",'" & sex & "','" & User & "') "
oSaccoMaster.ExecuteThis (sql)
End If

'//Update payroll
'Dim Startdate As String, Enddate As String
Set rs2 = New ADODB.Recordset
Dim qnty As Currency, GPay As Currency
'Startdate = DateSerial(DTPMilkDate, cboMonth, 1)

sql = "d_sp_UpdateGPAYQnty '" & Startdate & "','" & Enddate & "'," & txtSNo & ""
Set rs2 = oSaccoMaster.GetRecordset(sql)
If Not rs2.EOF Then
If Not IsNull(rs2.Fields(0)) Then qnty = rs2.Fields(0)
If Not IsNull(rs2.Fields(1)) Then GPay = rs2.Fields(1)
End If


Set Rs1 = New ADODB.Recordset
sql = "d_sp_TotalDeduct " & txtSNo & "," & month(DTPDDeduction) & "," & Year(DTPDDeduction) & ""
Set Rs1 = oSaccoMaster.GetRecordset(sql)
If Not Rs1.EOF Then
Dim TotalDed As Currency
If Not IsNull(Rs1.Fields(0)) Then TotalDed = Rs1.Fields(0)
End If
'//Update payroll -- @SNo bigint,@EndPeriod varchar(15),@Kgs float,@GPay money,@NPay money,@TDeductions money,@auditid  varchar(35)
Set cn = New ADODB.Connection
sql = "d_sp_UpdatePayroll  " & txtSNo & ",'" & Enddate & "'," & qnty & "," & GPay & "," & GPay - TotalDed & "," & TotalDed & ",'" & User & "'"
oSaccoMaster.ExecuteThis (sql)



Set rs3 = New ADODB.Recordset
'Dim Startdate As String, Enddate As String
Dim desc As String
Dim Amnt As Currency
'Startdate = DateSerial(Year(DTPDDeduction), month(DTPDDeduction), 1)
'Enddate = DateSerial(Year(DTPMilkDate), month(DTPMilkDate) + 1, 1 - 1)
sql = "d_sp_SupDed " & txtSNo & ",'" & Startdate & "','" & Enddate & "'"
Set rs3 = oSaccoMaster.GetRecordset(sql)
If Not rs3.EOF Then
While Not rs3.EOF
If Not IsNull(rs3.Fields(0)) Then desc = Trim(rs3.Fields(0))
Amnt = 0
If Not IsNull(rs3.Fields(1)) Then Amnt = rs3.Fields(1)
sql = "SET dateformat DMY SELECT     Transport, Agrovet, BONUS, TMShares, FSA, HShares, Advance, Others FROM d_Payroll WHERE SNo=" & txtSNo & " AND EndofPeriod ='" & Enddate & "'"
Set rs4 = oSaccoMaster.GetRecordset(sql)
If UCase(rs4.Fields(0).name) = UCase(desc) Then
Transport = Amnt
End If
If UCase(rs4.Fields(1).name) = UCase(desc) Then
agrovet = Amnt
End If
If UCase(rs4.Fields(2).name) = UCase(desc) Then
BONUS = Amnt
End If
If UCase(rs4.Fields(3).name) = UCase(desc) Then
TMShares = Amnt
End If
If UCase(rs4.Fields(4).name) = UCase(desc) Then
FSA = Amnt
End If
If UCase(rs4.Fields(5).name) = UCase(desc) Then
HShares = Amnt
End If
If UCase(rs4.Fields(6).name) = UCase(desc) Then
Advance = Amnt
End If
If UCase(rs4.Fields(7).name) = UCase(desc) Then
Others = Amnt
End If

'//d_sp_UpdatePayrollDed  @SNo bigint,@EndPeriod varchar(15),@Transport money,@Agrovet money,@AI money,@TMShares money,@FSA money,@HShares money,@Advance money,@Others
rs3.MoveNext
Wend
'//Update Deductions -- d_sp_UpdatePayrollDed  @SNo bigint,@EndPeriod varchar(15),@Transport money,@Agrovet money,@AI money,@TMShares money,@FSA money,@HShares money,@Advance money,@Others money
Set cn = New ADODB.Connection
sql = "d_sp_UpdatePayrollDed  " & txtSNo & ",'" & Enddate & "'," & Transport & "," & agrovet & "," & BONUS & "," & TMShares & "," & FSA & "," & HShares & "," & Advance & "," & Others & ""
oSaccoMaster.ExecuteThis (sql)
End If

Dim txtTCHPBalances As Double
If UCase(Trim(cboDeductionType)) = UCase("Shares") Then

Set rst = New ADODB.Recordset
sql = "select bal from d_shares where sno= '" & txtSNo & "'"
Set rst = oSaccoMaster.GetRecordset(sql)
If Not rst.EOF Then
'txtTCHPBalances = Rst.Fields(0)

 '//get the balance

sql = "SELECT     bal   FROM         d_sconribution  WHERE     sno ='" & txtSNo & "'  ORDER BY transdate DESC, id DESC "
Dim rr As New ADODB.Recordset
Set rr = oSaccoMaster.GetRecordset(sql)
If Not rr.EOF Then
txtTCHPBalances = txtTCHPBalances + CCur(txtamount)
',[sno],[transdate],[amount],[bal],[transdescription],[auditid],[auditdate],[mno]
  'From [EASYTEA].[dbo].[d_sconribution]
  sql = ""
  sql = "set dateformat dmy insert into d_sconribution([sno],[transdate],[amount],[bal],[transdescription],[auditid])"
  sql = sql & " values ('" & txtSNo & "','" & DTPDDeduction & "'," & txtamount & "," & txtTCHPBalances & ",'Shares','" & User & "') "
  oSaccoMaster.ExecuteThis (sql)
  
  sql = ""
  sql = "update d_shares set bal=" & txtTCHPBalances & " where sno='" & txtSNo & "' "
  oSaccoMaster.ExecuteThis (sql)
'txtTCHPBALANCE = rr.Fields(0)
End If
Else
'//add new one
txtTCHPBalances = 0
sql = "insert into d_Shares(sno, Cash,bal,auditid)"
sql = sql & " values('" & txtSNo & "',1," & txtamount & ",'" & User & "')"
oSaccoMaster.ExecuteThis (sql)
sql = ""
sql = "set dateformat dmy insert into d_sconribution([sno],[transdate],[amount],[bal],[transdescription],[auditid])"
sql = sql & " values ('" & txtSNo & "','" & DTPDDeduction & "'," & txtamount & "," & txtamount & ",'Shares','" & User & "') "
oSaccoMaster.ExecuteThis (sql)

End If
End If


Transport = 0
agrovet = 0
BONUS = 0
TMShares = 0
FSA = 0
HShares = 0
Advance = 0
Others = 0

'Dim Yr As Integer

'Yr = Year(DTPDDeduction)
'vbHourglass
'Fixed deductions update
'oSaccoMaster.ExecuteThis ("d_sp_PresetDeductAssign_99 '" & DTPStartDate & "','" & DTPEndDate & "'," & Yr & ",'" & User & "', " & txtSNo)

'Payroll update
'd_sp_GDedNet @StartDate varchar(10) , @endPeriod varchar(10)
'oSaccoMaster.ExecuteThis ("d_sp_GDedNet_99 '" & DTPStartDate & "','" & DTPEndDate & "'," & txtSNo)

'Update transporters
'd_sp_TransUpdate @StartDate varchar(10),@EndPeriod varchar(10),@User varchar(35) AS
'oSaccoMaster.ExecuteThis ("d_sp_TransUpdate_99 '" & DTPStartDate & "','" & DTPEndDate & "','" & User & "'," & txtSNo)


'oSaccoMaster.ExecuteThis ("d_sp_TransPRoll '" & DTPStartDate & "','" & DTPEndDate & "','" & User & "'")
'Lock period

txtamount = ""
txtSNo = ""
txtnetp = 0
txtSNo_Validate True

txtSNo.SetFocus
'Form_Load
MsgBox "Records successively saved."
loadBranchesTypes
'//
Exit Sub
ErrorHandler:
MsgBox err.description

End Sub

Private Sub Combo1_Change()

End Sub

Private Sub DTPDDeduction_Change()
On Error GoTo ErrorHandler
dtpStartDate = DateSerial(Year(DTPDDeduction), month(DTPDDeduction), 1)
DTPEndDate = DateSerial(Year(DTPDDeduction), month(DTPDDeduction) + 1, 1 - 1)
loadBranchesTypes
Exit Sub
ErrorHandler:

MsgBox err.description
End Sub



Private Sub Form_Load()
On Error GoTo ErrorHandler
txtamount = ""
txtSNames = ""
txtSNo = ""
txtremarks = ""

cboDeductionType = ""
txtamount.Locked = True
txtSNames.Locked = True
txtSNo.Locked = True
cboDeductionType.Locked = True

cmdnew.Enabled = True
cmdsave.Enabled = False
cmdEdit.Enabled = False
cmdDelete.Enabled = False

DTPDDeduction = Format(Get_Server_Date, "dd/mm/yyyy")
dtpStartDate = DateSerial(Year(DTPDDeduction), month(DTPDDeduction), 1)
'DTPStartDate = Format(Get_Server_Date, "dd/mm/yyyy")
DTPEndDate = DateSerial(Year(dtpStartDate), month(dtpStartDate) + 1, 1 - 1)
loadBranchesTypes
    cboDeductionType.Clear
    Set myclass = New cdbase

    Provider = myclass.OpenCon

    Set cn = CreateObject("adodb.connection")

   cn.Open Provider, "atm", "atm"

    Set rs = CreateObject("adodb.recordset")

    rs.Open "SELECT Description FROM d_DCodes", cn

    If rs.EOF Then Exit Sub

    With rs

        While Not .EOF

         cboDeductionType.AddItem rs.Fields("Description")

         .MoveNext

        Wend

    End With
    
    Set rst = New Recordset
    Set cn = New ADODB.Connection
    Provider = "MAZIWA"
    'Provider = cnn
    cn.Open Provider, "atm", "atm"
    Set rst = New Recordset
    
    sql = "Select Bname from   d_Branch order by Bname"
    rst.Open sql, cn, adOpenKeyset, adLockOptimistic
    While Not rst.EOF
    cbobranches.AddItem rst.Fields(0)
    rst.MoveNext
    Wend
    
Exit Sub
ErrorHandler:

MsgBox err.description

End Sub
Public Sub loadBranchesTypes()
    
    With ListView1
    
        .ListItems.Clear
        
        .ColumnHeaders.Clear
    
    End With

    Set rs2 = CreateObject("adodb.recordset")
    
    sql = "set dateformat dmy Select SNo, Date_Deduc, Description, Amount, Remarks, Branch from d_supplier_deduc where Date_Deduc='" & DTPDDeduction & "' order by auditdatetime desc"
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
        .ColumnHeaders.Add , , "Date"
        .ColumnHeaders.Add , , "Description"
        .ColumnHeaders.Add , , "Amount"
        .ColumnHeaders.Add , , "Remarks"
        .ColumnHeaders.Add , , "Branch"
'        .ColumnHeaders.Add , , "Mpesa"
'        .ColumnHeaders.Add , , "Outlet"
        While Not rs2.EOF
        
            Set li = .ListItems.Add(, , Trim(rs2.Fields("SNo")))
            
            li.ListSubItems.Add , , Trim(rs2.Fields("Date_Deduc"))
            li.ListSubItems.Add , , Trim(rs2.Fields("Description"))
            li.ListSubItems.Add , , Trim(rs2.Fields("Amount"))
            li.ListSubItems.Add , , Trim(rs2.Fields("Remarks"))
            li.ListSubItems.Add , , Trim(rs2.Fields("Branch"))
            rs2.MoveNext
        
        Wend
        
    End With
    
    rs2.Close
    
    Set rs2 = Nothing
    
ListView1.View = lvwReport
End Sub

Private Sub Form_LostFocus()
'txtAmount.DataFormat = FormatCurrency("'Kshs '#,##0.00", Val(txtAmount))
End Sub

Private Sub Form_Unload(Cancel As Integer)
'oSaccoMaster.ExecuteThis ("d_sp_GDedNet '" & DTPStartDate & "', '" & DTPEndDate & "'")
End Sub

Private Sub Picture5_Click()
        Me.MousePointer = vbHourglass
        frmSearchSupplier.Show vbModal
        txtSNo = sel
        txtSNo_Validate True
        Me.MousePointer = 0
End Sub

Private Sub txtRemarks_Change()
txtremarks.Text = cboDeductionType
End Sub

Private Sub txtSNo_Validate(Cancel As Boolean)
Dim ans As String
Dim NetP As Currency
'If rs.RecordCount > 0 Then
'If Not IsNull(rs.Fields("NPay")) Then
'NetP = rs.Fields("NPay")
'End If
'If Not IsNull(rs.Fields("GPay")) Then
'NetP = rs.Fields("GPay")
'End If

Startdate = DateSerial(Year(DTPDDeduction), month(DTPDDeduction), 1)
Enddate = DateSerial(Year(DTPDDeduction), month(DTPDDeduction) + 1, 1 - 1)
If txtSNo = "" Then Exit Sub
Set rs = oSaccoMaster.GetRecordset("d_sp_SupNet " & txtSNo & ",'" & Startdate & "','" & Enddate & "',0")
If Not rs.EOF Then
If Not IsNull(rs.Fields(1)) Then
If Not IsNull(rs.Fields(1)) Then
NetP = IIf(IsNull(rs.Fields(1)), 0, rs.Fields(1))
Else
NetP = "0.00"
End If

Set rs = oSaccoMaster.GetRecordset("d_sp_SupNet " & txtSNo & ",'" & Startdate & "','" & Enddate & "', 1")
If Not IsNull(rs.Fields(0)) Then
NetP = NetP - rs.Fields(0)
'- rs.Fields(0)
txtnetp.Text = NetP
txtnetp.Text = NetP
Else
'txtnetp.Tex = NetP
NetP = NetP
txtnetp.Text = NetP
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
