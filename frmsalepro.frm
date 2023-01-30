VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmsalepro 
   Caption         =   "Society Income Process"
   ClientHeight    =   5355
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6795
   LinkTopic       =   "Form1"
   ScaleHeight     =   5355
   ScaleWidth      =   6795
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdParchase 
      Caption         =   "Process Sales statement"
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
      Left            =   240
      TabIndex        =   7
      Top             =   720
      Width           =   2775
   End
   Begin VB.CommandButton cmdincomdairy 
      Caption         =   "Dairy Income Statement"
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
      TabIndex        =   6
      Top             =   3480
      Width           =   2415
   End
   Begin VB.CommandButton cmdpareport 
      Caption         =   "Purchase Report"
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
      Left            =   0
      TabIndex        =   5
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton cmdSalessta 
      Caption         =   "Sales Report"
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
      Left            =   1680
      TabIndex        =   4
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton cmddailys 
      Caption         =   "Daily Sales Report"
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
      Left            =   5640
      TabIndex        =   3
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Statements Process"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   6495
      Begin VB.Line Line1 
         BorderWidth     =   3
         X1              =   120
         X2              =   6360
         Y1              =   840
         Y2              =   840
      End
   End
   Begin MSComCtl2.DTPicker txtdateenterered 
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   120
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   139657217
      CurrentDate     =   38814
   End
   Begin MSComctlLib.ProgressBar prgStatus 
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   2280
      Visible         =   0   'False
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      Caption         =   "End Date to be process"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmsalepro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmddailys_Click()
Dim ans As String
ans = MsgBox("Do you Want a Report as per price??", vbYesNo)
If ans = vbYes Then
reportname = "SALES PER DAY.rpt"
Else
reportname = "dailysales.rpt"
End If
Show_Sales_Crystal_Report "", reportname, ""
End Sub

Private Sub cmdincomdairy_Click()
On Error GoTo ErrorHandler
Dim rst, rstg, rsa As Recordset
Startdate = DateSerial(year(txtdateenterered), month(txtdateenterered), 1)
Enddate = DateSerial(year(txtdateenterered), month(txtdateenterered) + 1, 1 - 1)
sql = ""
sql = "set dateformat dmy delete from d_incomestate where Date >= '" & Startdate & "' And Date<='" & Enddate & "'"
cn.Execute sql
     
     sql = ""
     sql = "set dateformat dmy Select distinct(TransDate) from d_Milkintake WHERE TransDate >='" & Startdate & "' And TransDate<='" & Enddate & "' order by TransDate asc"
     Set rstg = cn.Execute(sql)
     While Not rstg.EOF
      sql = ""
      sql = "set dateformat dmy Select isnull(sum(PAmount),0) from d_Milkintake WHERE TransDate ='" & rstg.Fields(0) & "'"
  '  sql = "set dateformat dmy SELECT d.DispQnty,m.DName, d.Price, d.DispQnty,d.DCode FROM d_MilkControl AS d INNER JOIN d_Debtors AS m ON d.DCode = m.DCode WHERE " & C & " and DispDate between " & Startdate & " And " & Enddate & """"
      Set rst = cn.Execute(sql)
      If Not rst.EOF Then
       sql = ""
       sql = "set dateformat dmy Select isnull(sum(Amount),0) from d_Debtorsparchases WHERE Date ='" & rstg.Fields(0) & "'"
       Set rsa = cn.Execute(sql)
        If Not rsa.EOF Then
         sql = ""
         sql = "set dateformat dmy insert into  d_incomestate(Date, Sales, Purchases,Diff)"
         sql = sql & "  values('" & rstg.Fields(0) & "','" & rsa.Fields(0) & "'," & rst.Fields(0) & ",'" & rsa.Fields(0) - rst.Fields(0) & "')"
         cn.Execute sql
        End If
      End If
       rstg.MoveNext
      Wend
reportname = "Incomestatement.rpt"
Show_Sales_Crystal_Report STRFORMULA, reportname, ""
Exit Sub
ErrorHandler:
MsgBox err.description
End Sub

Private Sub cmdParchase_Click()
On Error GoTo ErrorHandler
Startdate = DateSerial(year(txtdateenterered), month(txtdateenterered), 1)
Enddate = DateSerial(year(txtdateenterered), month(txtdateenterered) + 1, 1 - 1)
sql = ""
sql = "set dateformat dmy delete from d_Debtorsparchases where Date >= '" & Startdate & "' And Date<='" & Enddate & "'"
cn.Execute sql
prgStatus.Visible = True
txtlbl.Visible = 1
If txtlbl.Visible = True Then
 txtlbl = "Please wait as it precess"
End If
      ' MsgBox "Please wait as it precess  "
sql = ""
sql = "set dateformat dmy Select count(distinct(DCode)) as j  from   d_MilkControl where DispDate >= '" & Startdate & "' And DispDate<='" & Enddate & "' "
Set rs = cn.Execute(sql)
'Set rs = oSaccoMaster.GetRecordset(sql)
Dim a As Double
a = rs.Fields(0)
j = rs.Fields(0)
prgStatus.max = 100
prgStatus.Min = 0
I = 0
'baet
 sql = ""
 sql = "set dateformat dmy Select distinct(DCode) from   d_MilkControl where DispDate >= '" & Startdate & "' And DispDate<='" & Enddate & "' order by DCode asc"
 Set rsd = cn.Execute(sql)
  While Not rsd.EOF
  Do While Not j = 0
I = I + 1
prgStatus = Round((I / a) * 100, 0)
  If Not rsd.EOF Then
  'C = "CB003A22"
  C = rsd.Fields(0)

'     If C = "CB003A22" Then
'        MsgBox "Warning:Please wait " & C & "", vbInformation
'     End If
 '  Label40 = "Please wait as it precess"

   sql = ""
   sql = "set dateformat dmy Select distinct(Price) from   d_MilkControl where DCode='" & C & "'and DispDate >= '" & Startdate & "' And DispDate<='" & Enddate & "' "
   Set rs = cn.Execute(sql)
    sql = ""
    sql = "set dateformat dmy Select count(distinct(Price)) from   d_MilkControl where DCode='" & C & "'and DispDate >= '" & Startdate & "' And DispDate<='" & Enddate & "' "
    Set rsg = cn.Execute(sql)
     'g = rsg.Fields(0)
     'q = 1
   'For q = 1 To g
   Do While Not rs.EOF
    If Not rs.EOF Then
     k = rs.Fields(0)
     sql = ""
     sql = "set dateformat dmy Select distinct(DispDate) from   d_MilkControl WHERE DCode= '" & C & "'and Price='" & k & "' and DispDate >='" & Startdate & "' And DispDate<='" & Enddate & "'"
     Set rstg = cn.Execute(sql)
     While Not rstg.EOF
      sql = ""
      sql = "set dateformat dmy Select sum(DispQnty) from   d_MilkControl WHERE DCode= '" & C & "'and Price='" & k & "' and DispDate ='" & rstg.Fields(0) & "'"
  '  sql = "set dateformat dmy SELECT d.DispQnty,m.DName, d.Price, d.DispQnty,d.DCode FROM d_MilkControl AS d INNER JOIN d_Debtors AS m ON d.DCode = m.DCode WHERE " & C & " and DispDate between " & Startdate & " And " & Enddate & """"
      Set rst = cn.Execute(sql)
      If Not rst.EOF Then
        sql = ""
        sql = "Select distinct(DName) from   d_Debtors where DCode='" & C & "' "
        Set rsa = cn.Execute(sql)
        If Not rsa.EOF Then
         p = rsa.Fields(0)
           sql = ""
           sql = "Select distinct(Locations) from   d_Debtors where DCode='" & C & "' "
           Set rsv = cn.Execute(sql)
         sql = ""
         sql = "set dateformat dmy insert into  d_Debtorsparchases(Debtor, Name, Kgs, Price,Amount,Description,Branch,Date)"
         sql = sql & "  values('" & C & "','" & p & "'," & rst.Fields(0) & ",'" & k & "','" & rst.Fields(0) * k & "','OUTLET SALES','" & rsv.Fields(0) & "','" & rstg.Fields(0) & "')"
         cn.Execute sql
        End If
      End If
       rstg.MoveNext
      Wend
    End If
  k = rs.MoveNext
  Loop
  'Next q
    Else
    'baet
    'kamorok
  txtlbl.Visible = False
  
  
  MsgBox "Completed succesfully ", vbInformation
    Exit Sub
    End If

  j = j - 1

  'MsgBox "Warning:Please wait " & j & "", vbInformation
 rsd.MoveNext

Loop
Wend
  'baet
  'kamorok
  txtlbl.Visible = False
MsgBox "Completed succesfully ", vbInformation
Exit Sub
ErrorHandler:
MsgBox err.description
End Sub

Private Sub cmdpareport_Click()
reportname = "MILK PURCHASE REPORT.rpt"
Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub cmdSalessta_Click()
Dim ans As String
ans = MsgBox("Do you Want a Report as per price??", vbYesNo)
If ans = vbYes Then
 'reportname = "d_Dailysummary2.rpt"
reportname = "MILK SALES REPORT1.rpt"
 Else
 'reportname = "d_Dailysummary.rpt"
reportname = "MILK SALES REPORT.rpt"
 End If
 Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub
