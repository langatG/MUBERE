VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmstockbalance 
   Caption         =   "Stock Balance Processing"
   ClientHeight    =   2010
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   ScaleHeight     =   2010
   ScaleWidth      =   5415
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdfinalre 
      Caption         =   "Process"
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   1935
   End
   Begin VB.CommandButton cmdcarry 
      Caption         =   "Carry forward"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   4
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdreport 
      Caption         =   "Stock Balance Report"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   3
      Top             =   1440
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Process Balance"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker DTPedate 
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   360
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      Format          =   120520705
      CurrentDate     =   42680
   End
   Begin MSComctlLib.ProgressBar prgStatus 
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1080
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label lbledate 
      Caption         =   "End Month Date"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "frmstockbalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcarry_Click()
Dim lastdate As Date
Dim lastdateofsale As Date
Dim pcode As String
Dim dy As Integer
Dim grade As String
Dim curamt As Double
Dim pprice As Double
Dim sprice As Double

Dim rsd As New ADODB.Recordset
Startdate = DateSerial(year(DTPedate), month(DTPedate), 1)
Enddate = DateSerial(year(DTPedate), month(DTPedate) + 1, 1)
sql = ""
sql = "set dateformat dmy DELETE FROM ag_sales where Code='CONFIRM' and Date='" & Enddate & "' "
Set rsd = oSaccoMaster.GetRecordset(sql)

sql = ""
sql = "set dateformat dmy SELECT pcode,pname,Quantity,branch,Bprice, SPrice From ag_sales WHERE Date>='" & Startdate & "' and Date <= '" & DTPedate & "' and Code='RAW' "
Set rs = oSaccoMaster.GetRecordset(sql)
While Not rs.EOF
sql = ""
sql = "set dateformat dmy insert into  ag_sales (pcode,pname,Quantity,branch,Date,Bprice, SPrice,Code)"
sql = sql & "values('" & rs.Fields(0) & "','" & rs.Fields(1) & "','" & rs.Fields(2) & "','" & rs.Fields(3) & "','" & Enddate & "','" & rs.Fields(4) & "','" & rs.Fields(5) & "','CONFIRM') "
oSaccoMaster.ExecuteThis (sql)
rs.MoveNext
Wend

Startdate = DateSerial(year(DTPedate), month(DTPedate), 1)
Enddate = DateSerial(year(DTPedate), month(DTPedate) + 1, 1 - 1)

sql = ""
sql = "set dateformat dmy DELETE FROM ag_sales where Code='RAW' and Date>='" & Startdate & "' and Date<='" & Enddate & "' "
Set rstg = oSaccoMaster.GetRecordset(sql)

MsgBox "Records successfully Inserted", vbInformation

End Sub

Private Sub cmdfinalre_Click()
Dim lastdate As Date
Dim lastdateofsale As Date
Dim pcode As String
Dim dy As Currency
Dim grade As String
Dim curamt As Double
Dim code As Double
Dim Openning As Double

Dim rsd, rsd1, rsd2, rsd3, rsd4 As New ADODB.Recordset
Startdate = DateSerial(year(DTPedate), month(DTPedate), 1)
Enddate = DateSerial(year(DTPedate), month(DTPedate) + 1, 1 - 1)
'check the user
sql = "SELECT     UserLoginIDs,levels, UserGroup, SUPERUSER From UserAccounts where UserLoginIDs='" & User & "'"
Set rsd3 = oSaccoMaster.GetRecordset(sql)
If Not rsd3.EOF Then
 If rsd3!Levels <> "Manager" Then
  If rsd3!Levels <> "Accounts" Then
   MsgBox "You are not allowed to Process", vbInformation
   Exit Sub
  End If
 End If
End If


sql = ""
sql = "set dateformat dmy DELETE FROM ag_sales where  Date>='" & Startdate & "' and Date<='" & Enddate & "'"
Set rsd4 = oSaccoMaster.GetRecordset(sql)
prgStatus.Visible = True
prgStatus.Max = 100
prgStatus.Min = 0
I = 0
sql = ""
sql = "set dateformat dmy Select count(distinct(p_code)) as u from ag_Products "
Set rs5 = cn.Execute(sql)
Dim a As Double
a = rs5.Fields(0)
''''''SELECT DISTICT PRODUCT CODE
sql = ""
sql = "set dateformat dmy SELECT distinct(p_code),p_name From ag_Products order by p_code asc"
Set rs = oSaccoMaster.GetRecordset(sql)
'dy = Trim(rs.Fields(0))
While Not rs.EOF
 'code = Trim(rs.Fields(0))
   I = I + 1
prgStatus = Round((I / a) * 100, 0)
 ''''''SELECT DISTINCT BRANCHES
 sql = ""
 sql = "set dateformat dmy SELECT distinct(Branch) From ag_Products where Branch<>'' AND P_code='" & rs.Fields(0) & "' order by Branch asc"
 Set rsd = oSaccoMaster.GetRecordset(sql)
 While Not rsd.EOF
  ''''SUM GOOD SOLD THAT MONTH
  sql = ""
  'sql = "set dateformat dmy SELECT isnull(sum(Qua),0) FROM ag_Receipts where Branch='" & rsd.Fields(0) & "'AND P_code='" & code & "' and T_Date>='" & Startdate & "' and T_Date<='" & Enddate & "' "
  sql = "d_sp_stockprocessing1 '" & rsd.Fields(0) & "','" & rs.Fields(0) & "','" & Startdate & "','" & Enddate & "'"
  Set rstg = oSaccoMaster.GetRecordset(sql)
  ''''CAL BAL
  sql = ""
  sql = "set dateformat dmy SELECT top(1)Qout FROM ag_Products where Branch='" & rsd.Fields(0) & "'AND p_code='" & rs.Fields(0) & "' order by Last_D_Updated desc "
  Set rsd1 = oSaccoMaster.GetRecordset(sql)
  ''''SUM QNTY IN
  sql = ""
  'sql = "set dateformat dmy SELECT isnull(SUM(Qin),0) FROM ag_Products3 where Branch ='" & rsd.Fields(0) & "'AND p_code='" & code & "' and audit_date>='" & Startdate & "' and audit_date<='" & Enddate & "' "
  sql = "d_sp_stockprocessing2 '" & rsd.Fields(0) & "','" & rs.Fields(0) & "','" & Startdate & "','" & Enddate & "'"
  Set rsd2 = oSaccoMaster.GetRecordset(sql)
  Dim purchases, sales, balance As Double
    purchases = rsd2.Fields(0)
    sales = rstg.Fields(0)
    balance = rsd1.Fields(0)
    Openning = balance + sales - purchases
      ''''''insert to the table now
      sql = ""
      sql = "set dateformat dmy insert into  ag_sales (pcode,pname,Openning,Purchases, Sales,Balance,Branch,Date)"
      sql = sql & "values('" & rs.Fields(0) & "','" & rs.Fields(1) & "','" & Openning & "','" & purchases & "','" & sales & "','" & balance & "','" & rsd.Fields(0) & "','" & Enddate & "') "
      oSaccoMaster.ExecuteThis (sql)
    
 rsd.MoveNext
 Wend
rs.MoveNext
Wend
MsgBox "Records successfully done", vbInformation
Exit Sub
End Sub

Private Sub cmdreport_Click()
 reportname = "stobal.rpt"
 Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub Command1_Click()
Dim lastdate As Date
Dim lastdateofsale As Date
Dim pcode As String
Dim dy As Integer
Dim grade As String
Dim curamt As Double
Dim pprice As Double
Dim sprice As Double

Dim rsd As New ADODB.Recordset
Startdate = DateSerial(year(DTPedate), month(DTPedate), 1)
Enddate = DateSerial(year(DTPedate), month(DTPedate) + 1, 1 - 1)

sql = ""
sql = "set dateformat dmy DELETE FROM ag_sales where Code='RAW' and Date>='" & Startdate & "' and Date<='" & Enddate & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
sql = ""
sql = "set dateformat dmy SELECT SUM(Qin) AS Quantity, p_code,p_name,branch,pprice, sprice From ag_Products3 WHERE Date_Entered => '" & Startdate & "' and Date_Entered <= '" & Enddate & "' GROUP BY p_code,p_name,branch,pprice, sprice"
Set rs = oSaccoMaster.GetRecordset(sql)
While Not rs.EOF
pcode = rs!p_code
Quantity = IIf(IsNull(rs!Quantity), 0, rs!Quantity)
pname = rs!p_name
Branch = rs!Branch
BPRICE = rs!pprice
sprice = rs!sprice
sql = ""
sql = "set dateformat dmy SELECT SUM(Qua) AS Qty, p_code,Branch From ag_Receipts WHERE T_Date => '" & Startdate & "' and T_Date <= '" & Enddate & "' and p_code='" & rs!p_code & "'and branch='" & rs!Branch & "'  GROUP BY p_code,Branch"
Set rsd = oSaccoMaster.GetRecordset(sql)
If Not rsd.EOF Then
curamt = IIf(IsNull(rs!Quantity), 0, rs!Quantity) - IIf(IsNull(rsd!qty), 0, rsd!qty)
Else
curamt = IIf(IsNull(rs!Quantity), 0, rs!Quantity)
End If
'curamt = IIf(IsNull(rs!Quantity), 0, rs!Quantity) - IIf(IsNull(rsd!qty), 0, rsd!qty)
'select pcode,ldate,dy,auditdate,audit,grade from ag_paging
sql = ""
sql = "set dateformat dmy insert into  ag_sales (pcode,pname,Quantity,branch,Date,Bprice, SPrice,Code)"
sql = sql & "values('" & pcode & "','" & pname & "','" & curamt & "','" & Branch & "','" & Enddate & "','" & BPRICE & "','" & sprice & "','RAW') "
oSaccoMaster.ExecuteThis (sql)


rs.MoveNext
Wend
MsgBox "Records successfully done", vbInformation
Exit Sub
'//give him the report here
'agrovetagingreport
'reportname = "stobal.rpt"
''reportname = "evans.rpt"
'
'
' Show_Sales_Crystal_Report STRFORMULA, reportname, ""
'//we look for receipts tables
'//get the number of days
'/// insert into the number of days
'//give us a report

End Sub

Private Sub Form_Load()
DTPedate = Get_Server_Date
Startdate = DateSerial(year(txtransdate), month(txtransdate), 1)
Enddate = DateSerial(year(txtransdate), month(txtransdate) + 1, 1 - 1)
End Sub
