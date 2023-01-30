VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAccounts 
   Caption         =   "Generate Trial Balance"
   ClientHeight    =   3585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7695
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   7695
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command7 
      Caption         =   "Print  Balance Sheet"
      Height          =   405
      Left            =   5520
      TabIndex        =   14
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Print Both Assets & Liabilities"
      Height          =   615
      Left            =   4080
      TabIndex        =   13
      Top             =   2760
      Width           =   1575
   End
   Begin VB.ComboBox Cboactivity 
      Height          =   330
      ItemData        =   "frmAccounts.frx":0000
      Left            =   4680
      List            =   "frmAccounts.frx":0010
      TabIndex        =   12
      Top             =   480
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Process"
      Height          =   330
      Left            =   120
      TabIndex        =   11
      Top             =   2280
      Width           =   1335
   End
   Begin MSComctlLib.ProgressBar prgStatus 
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1800
      Visible         =   0   'False
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Print Income Statement"
      Height          =   375
      Left            =   3300
      TabIndex        =   7
      Top             =   2205
      Width           =   1935
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print TB"
      Height          =   375
      Left            =   1635
      TabIndex        =   6
      Top             =   2205
      Width           =   1545
   End
   Begin MSComCtl2.DTPicker dtpFinishDate 
      Height          =   345
      Left            =   2970
      TabIndex        =   2
      Top             =   510
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   " dd-MM-yyyy"
      Format          =   138543107
      CurrentDate     =   39705
   End
   Begin MSComCtl2.DTPicker dtpStartDate 
      Height          =   360
      Left            =   1290
      TabIndex        =   1
      Top             =   510
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   " dd-MM-yyyy"
      Format          =   138543107
      CurrentDate     =   39705
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Process"
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   3000
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Label lblAccount 
      BackColor       =   &H00FFFF80&
      Height          =   375
      Left            =   6360
      TabIndex        =   10
      Top             =   1320
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblStatus 
      BackColor       =   &H00FFFF80&
      Height          =   375
      Left            =   3000
      TabIndex        =   9
      Top             =   1320
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Finish Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3000
      TabIndex        =   5
      Top             =   285
      Width           =   945
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Start Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1320
      TabIndex        =   4
      Top             =   270
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   600
      TabIndex        =   3
      Top             =   2865
      Visible         =   0   'False
      Width           =   645
   End
End
Attribute VB_Name = "frmAccounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim I As Long


Private Sub cmdEOY_Click()
Call EOY_Processing(dtpFinishDate)
End Sub

Private Sub cmdPrint_Click()
 
 reportname = "Trial Balance.rpt"
 
 Show_Sales_Crystal_Report "", reportname, CompanyName

    
End Sub

Private Sub EOY_Processing(EOYDate As Date)
Dim ACCNO As String, amount As String, transdate As Date, Glacctype As String

Set rs = Nothing
Set rs = oSaccoMaster.GetRecordset("set dateformat dmy select * from TBBalance where TransDate='" & EOYDate & "'")
If rs.EOF Then
    MsgBox "Trial Balance has not been generated, Please generate it before proceeding.", vbCritical, Me.Caption
    Exit Sub
End If

Set rs = oSaccoMaster.GetRecordset("select AccNo,Glacctype from GlSetup order by accno")
With rs
    If Not .EOF Then
      While Not .EOF
      Me.Caption = !ACCNO
        Set rst = oSaccoMaster.GetRecordset("set dateformat dmy select AccNo,Amount,transdate From TBBalance" _
        & " where AccNO='" & !ACCNO & "' and transdate='" & EOYDate & "' order by AccNO")
         If Not rst.EOF Then
         
          If !Glacctype = "Income Statement" Then
          amount = 0
          Set rst1 = oSaccoMaster.GetRecordset("set dateformat dmy update GLSETUP set NewGLOpeningBal=0,NewGLOpeningBalDate='" & EOYDate & "',CurrentBal=" & amount & " where AccNo='" & !ACCNO & "'")

          Else
          amount = rst!amount
          Set rst1 = oSaccoMaster.GetRecordset("set dateformat dmy update GLSETUP set NewGLOpeningBal=" & amount & ",NewGLOpeningBalDate='" & EOYDate & "',CurrentBal=" & amount & " where AccNo='" & !ACCNO & "'")

         End If

    End If
    .MoveNext
    Wend
End If
End With

End Sub

Private Sub Command1_Click()
    On Error Resume Next
    Dim ACCNO As String
    Dim suspense As Double
    Dim Debits As Double, Credits As Double
    Dim ACCBAL As Double
    Dim transtype As String, DocumentNo As String, accType As String, AccGroup As String, AccName As String, OpeningBal As Double, TotalDr As Double, TotalCr As Double
    
    If Not oSaccoMaster.Execute("Truncate table TBBALANCE") Then
        GoTo SysError
    End If
    sql = "SELECT Accno,NormalBal,glaccType,glaccname,glaccGroup FROM GLSETUP ORDER BY ACCNO"
    Set rst = oSaccoMaster.GetRecordset(sql)
    With rst
        If Not .EOF Then
            prgStatus.Visible = True
            lblStatus.Visible = True
            lblAccount.Visible = True
            prgStatus.max = 100
            'prgStatus.Min = 0
            I = 0
            While Not .EOF
                DoEvents
                I = I + 1
                lblStatus.Caption = CStr((I / .RecordCount)) * 100 & " %"
                prgStatus.value = Round((I / .RecordCount) * 100, 0)
                ACCNO = !ACCNO
                'If Accno = "I005" Then MsgBox "Here"
                lblAccount = ACCNO
                accType = !Glacctype
                AccGroup = !GLAccGroup
                AccName = !GlAccName
                
                'OpeningBal = getGlBalance(Accno, dtpStartDate, dtpStartDate)
           
                ACCBAL = getGlBalance(ACCNO, DTPStartDate, dtpFinishDate)
                If Not success Then
                    GoTo SysError
                End If
                
                If UCase(accType) = UCase("Income Statement") Then
                    ACCBAL = ACCBAL - OpeningBal
                End If
                
                transtype = IIf(!NormalBal = "Debit", "DR", "CR")
                'save
                If ACCBAL <> 0 Then
                    sql = "Set DateFormat DMY INSERT INTO [tbbalance] ([AccNo],[AccName], [Amount],[Transtype],[StartDate], [EndDate], [AuditID], [AccType], [AccGroup], [BudgetAmount],OBAL,DR,CR)"
                    sql = sql & " Values('" & ACCNO & "','" & AccName & "'," & ACCBAL & ",'" & transtype & "','" & DTPStartDate.value & "','" & dtpFinishDate.value & _
                    "','" & User & "','" & accType & "','" & AccGroup & "',0," & OpeningBal & "," & TotalDr & "," & TotalCr & ")"
                        
                    If Not oSaccoMaster.Execute(sql) Then
                        GoTo SysError
                    End If
                End If
                
                .MoveNext
            Wend
        Else
            prgStatus.Visible = False
            lblStatus.Visible = False
            lblAccount.Visible = False
        End If
    End With
    
    Set rst = oSaccoMaster.GetRecordset("select REArningsAcc,SuspenseAcc from sysparam")
If Not rst.EOF Then
    REarningAcc = rst(0)
    SuspenseAcc = rst(1)
End If
    
    Set rst = oSaccoMaster.GetRecordset("SELECT  (SELECT     isnull(SUM(Amount),0) FROM  tbbalance WHERE transtype = 'DR') AS Debits, (SELECT     isnull(SUM(Amount),0) FROM  tbbalance WHERE transtype = 'CR') AS Credits")
    If Not rst.EOF Then
        If rst("Debits") > rst("Credits") Then
            Credits = rst("Debits") - rst("Credits")
            ACCBAL = rst("Debits") - rst("Credits")
            transtype = "CR"
        Else
            Debits = rst("Credits") - rst("Debits")
            ACCBAL = rst("Credits") - rst("Debits")
            transtype = "DR"
        End If
        
        If ACCBAL > 0 Then
            sql = "Set DateFormat DMY INSERT INTO [tbbalance] ([AccNo],[AccName], [Amount],[Transtype], [Closed],[StartDate], [EndDate], [AuditID], [AccType], [AccGroup], [BudgetAmount])"
            
            sql = sql & " Values('" & SuspenseAcc & "','" & AccName & "'," & ACCBAL & ",'" & transtype & "',0,'" & DTPStartDate & "','" & dtpFinishDate.value & _
            "','" & User & "','" & accType & "','" & AccGroup & "',0)"
                
            If Not oSaccoMaster.Execute(sql) Then
                GoTo SysError
            End If
        End If
        
        'For BalanceSheet Items, check whether they balance
        Set rst = oSaccoMaster.GetRecordset("SELECT  isnull((SELECT     SUM(Amount) FROM  tbbalance WHERE transtype = 'DR' and acctype='Balance Sheet'),0) AS Debits, isnull((SELECT     SUM(Amount) FROM  tbbalance WHERE transtype = 'CR' and acctype='Balance Sheet'),0) AS Credits")
        If Not rst.EOF Then
            If rst("Debits") > rst("Credits") Then
                Credits = rst("Debits") - rst("Credits")
                ACCBAL = rst("Debits") - rst("Credits")
                transtype = "CR"
            Else
                Debits = rst("Credits") - rst("Debits")
                ACCBAL = rst("Credits") - rst("Debits")
                transtype = "DR"
            End If
            
        If ACCBAL <> 0 Then
            sql = "Set DateFormat DMY INSERT INTO [tbbalance] ([AccNo],[AccName], [Amount],[Transtype], [Closed],[StartDate], [EndDate], [AuditID], [AccType], [AccGroup], [BudgetAmount])"
            
            sql = sql & " Values('" & REarningAcc & "','" & UCase("Retained Earnings") & "'," & ACCBAL & ",'" & transtype & "',0,'" & DTPStartDate & "','" & dtpFinishDate.value & _
            "','" & User & "','" & accType & "','" & AccGroup & "',0)"
                
            If Not oSaccoMaster.Execute(sql) Then
                GoTo SysError
            End If
        End If
        
        End If
    End If
    Exit Sub
SysError:
    Command1.Enabled = True
    MsgBox IIf(ErrorMessage = "", err.description, ErrorMessage), vbInformation
End Sub

Private Sub Command2_Click()
    reportname = "incomeandexpenditure.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, CompanyName
    Exit Sub
    
  Command1_Click
  Dim rsAccounts As Recordset
  Dim rsBudgets As Recordset
  '// get the budget amount and then variances
  Set rsAccounts = oSaccoMaster.GetRecordset("Select * From GLSETUP order by AccNo")
         While Not rsAccounts.EOF
  
            Set rsBudgets = oSaccoMaster.GetRecordset("Set DateFormat DMY Select Budgetted" _
            & " As BudgetAmount From BUDGETS where AccNo='" & rsAccounts.Fields("accno") & "' and mmonth='" _
            & month(dtpFinishDate) & "' and yyear='" & year(dtpFinishDate) & "'")
            If Not rsBudgets.EOF Then
            Dim b As Currency
            '//updates on the tem
            b = rsBudgets.Fields(0) '* CDbl(month(dtpFinishDate))
            sql = ""
            sql = "UPDATE    TBBALANCE  SET budgetAMOUNT =" & b & " where accno='" & rsAccounts.Fields("accno") & "'"
            oSaccoMaster.ExecuteThis (sql)
            Else
                        sql = ""
            sql = "UPDATE    TBBALANCE  SET budgetAMOUNT=0 where accno='" & rsAccounts.Fields("accno") & "'"
            oSaccoMaster.ExecuteThis (sql)
            End If
            b = 0
            rsAccounts.MoveNext
        Wend
        
        '//GET THE NETINCOME
        Dim totexpenses As Currency, TotIncome  As Currency
        Dim totexpensesB As Currency, TotIncomeB  As Currency
       
    Set rs = oSaccoMaster.GetRecordset("SELECT     SUM(TBBALANCE.Amount) AS Expr1,sum(budgetamount) as AI  FROM         TBBALANCE TBBALANCE INNER JOIN  GLSETUP GLSETUP ON TBBALANCE.AccNo = GLSETUP.AccNo WHERE     (GLSETUP.GlAccMainGroup = 'INCOME')")
        If Not rs.EOF Then
            If Not IsNull(rs.Fields(0)) Then TotIncome = rs.Fields(0)
            If Not IsNull(rs.Fields(1)) Then TotIncomeB = rs.Fields(1)
        End If
    Set rst = oSaccoMaster.GetRecordset("SELECT     SUM(TBBALANCE.Amount) AS Expr1,sum(budgetamount) as AE  FROM         TBBALANCE TBBALANCE INNER JOIN  GLSETUP GLSETUP ON TBBALANCE.AccNo = GLSETUP.AccNo WHERE     (GLSETUP.GlAccMainGroup = 'EXPENSES')")
        If Not rst.EOF Then
            If Not IsNull(rst.Fields(0)) Then totexpenses = rst.Fields(0)
            If Not IsNull(rst.Fields(1)) Then totexpensesB = rst.Fields(1)
        End If
 
 reportname = "kimincomeandexpenditure.rpt"
 
 Show_Sales_Crystal_Report "", reportname, "& Company_Name &"

   
    MousePointer = vbDefault

    Command1.Enabled = True
    Exit Sub
SysError:
    Command1.Enabled = True
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub Command3_Click()
    '//kimberbalancesheet
    reportname = "BalanceSheeetA.rpt"
    STRFORMULA = ""
    Show_Sales_Crystal_Report STRFORMULA, reportname, CompanyName
Exit Sub

'IF YOU WANT TO DO A CSV FILE
Command1_Click

 
End Sub

Private Sub Command4_Click()
    On Error Resume Next
    Dim ACCNO As String
    Dim suspense As Double
    Dim Debits As Double, Credits As Double
    Dim ACCBAL As Double
    Dim OpeningBal As Double
    Dim transtype As String, DocumentNo As String, accType As String, AccGroup As String, AccName As String, TotalDr As Double, TotalCr As Double
    
    
    If year(DTPStartDate) < year(dtpFinishDate) Then
        MsgBox "The choosen period crosses the definition of the funancial period", vbCritical, "INVALID PERIOD"
        Exit Sub
    End If
    
    
    If Not oSaccoMaster.Execute("Truncate table TBBALANCE") Then
        GoTo SysError
    End If
    Dim rst89 As New ADODB.Recordset
    sql = "SELECT Accno,NormalBal,glaccType,glaccname,glaccGroup FROM GLSETUP WHERE  subtype='" & Cboactivity & "' ORDER BY ACCNO"
    Set rst89 = oSaccoMaster.GetRecordset(sql)
    With rst89
        If Not .EOF Then
            prgStatus.Visible = True
            lblStatus.Visible = True
            lblAccount.Visible = True
            prgStatus.max = 100
            prgStatus.Min = 0
            I = 0
            
            While Not .EOF
                I = I + 1
                lblStatus.Caption = CStr(Round((I / .RecordCount) * 100, 0)) & " %"
                prgStatus.value = Round((I / .RecordCount) * 100, 0)
                ACCNO = !ACCNO
'                If ACCNO = "11-302B" Then
'                    MsgBox "Here"
'                End If
                lblAccount = ACCNO
                accType = !Glacctype
                AccGroup = !GLAccGroup
                AccName = !GlAccName
                
                'OpeningBal = getGlBalance(ACCNO, DTPStartDate, DTPStartDate)
                ACCBAL = getGlBalance(ACCNO, DTPStartDate, dtpFinishDate)
                If Not success Then
                    GoTo SysError
                End If
                If UCase(accType) = UCase("Income Statement") Then
                 ACCBAL = getGlBalance1(ACCNO, DTPStartDate, dtpFinishDate)
                    ACCBAL = ACCBAL - OpeningBal
                End If

                If !NormalBal = "Debit" Then
                    If ACCBAL >= 0 Then
                        transtype = "DR"
                    Else
                        transtype = "CR"
                    End If
                Else
                    If ACCBAL >= 0 Then
                        transtype = "CR"
                    Else
                        transtype = "DR"
                    End If
                End If
                
                ACCBAL = Abs(ACCBAL)
                
                'save
                
                If ACCBAL <> 0 Then
                    sql = "Set DateFormat DMY INSERT INTO [tbbalance] ([AccNo],[AccName], [Amount],[Transtype],[StartDate], [EndDate], [AuditID], [AccType], [AccGroup], [BudgetAmount],OBAL,DR,CR)"
                    sql = sql & " Values('" & ACCNO & "','" & AccName & "'," & ACCBAL & ",'" & transtype & "','" & DTPStartDate.value & "','" & dtpFinishDate.value & _
                    "','" & User & "','" & accType & "','" & AccGroup & "',0,0," & TotalDr & "," & TotalCr & ")"
                        
                    If Not oSaccoMaster.Execute(sql) Then
                        GoTo SysError
                    End If
                End If
                
                .MoveNext
            Wend
       Else
            prgStatus.Visible = False
            lblStatus.Visible = False
            lblAccount.Visible = False
        End If
    End With
    
    Set rst = oSaccoMaster.GetRecordset("SELECT  (SELECT     isnull(SUM(Amount),0) FROM  tbbalance WHERE transtype = 'DR') AS Debits, (SELECT     isnull(SUM(Amount),0) FROM  tbbalance WHERE transtype = 'CR') AS Credits")
    If Not rst.EOF Then
        If rst("Debits") > rst("Credits") Then
            Credits = rst("Debits") - rst("Credits")
            ACCBAL = rst("Debits") - rst("Credits")
            transtype = "CR"
            
            
        Else
            Debits = rst("Credits") - rst("Debits")
            ACCBAL = rst("Credits") - rst("Debits")
            transtype = "DR"
        End If
        
        If ACCBAL > 0 Then
            sql = "Set DateFormat DMY INSERT INTO [tbbalance] ([AccNo],[AccName], [Amount],[Transtype], [Closed],[StartDate], [EndDate], [AuditID], [AccType], [AccGroup], [BudgetAmount])"
            
            sql = sql & " Values('" & SuspenseAcc & "','" & AccName & "'," & ACCBAL & ",'" & transtype & "',0,'" & DTPStartDate & "','" & dtpFinishDate.value & _
            "','" & User & "','" & accType & "','" & AccGroup & "',0)"
                
            If Not oSaccoMaster.Execute(sql) Then
                GoTo SysError
            End If
        End If
        
        'For BalanceSheet Items, check whether they balance
        Set rst = oSaccoMaster.GetRecordset("SELECT  isnull((SELECT     SUM(Amount) FROM  tbbalance WHERE transtype = 'DR' and acctype='Balance Sheet'),0) AS Debits, isnull((SELECT     SUM(Amount) FROM  tbbalance WHERE transtype = 'CR' and acctype='Balance Sheet'),0) AS Credits")
        If Not rst.EOF Then
            If rst("Debits") > rst("Credits") Then
                Credits = rst("Debits") - rst("Credits")
                ACCBAL = rst("Debits") - rst("Credits")
                transtype = "CR"
            Else
                Debits = rst("Credits") - rst("Debits")
                ACCBAL = rst("Credits") - rst("Debits")
                transtype = "DR"
            End If
        'retained Earnings
'        If ACCBAL <> 0 Then
'            sql = "Set DateFormat DMY INSERT INTO [tbbalance] ([AccNo],[AccName], [Amount],[Transtype], [Closed],[StartDate], [EndDate], [AuditID], [AccType], [AccGroup], [BudgetAmount])"
'
'            sql = sql & " Values('" & REarningsAcc & "','" & UCase("Retained Earnings") & "'," & ACCBAL & ",'" & transtype & "',0,'" & dtpStartDate & "','" & dtpFinishDate.value & _
'            "','" & User & "','" & accType & "','" & AccGroup & "',0)"
'
'            If Not oSaccoMaster.Execute(sql) Then
'                GoTo SysError
'            End If
'        End If
        
        End If
    End If
    
    MsgBox "Process Done", vbInformation
    Exit Sub
SysError:
    Command1.Enabled = True
    MsgBox IIf(ErrorMessage = "", err.description, ErrorMessage), vbInformation

End Sub

Private Sub Command5_Click()
    '//kimberbalancesheet
    reportname = "BalanceSheeetL.rpt"
    STRFORMULA = ""
    Show_Sales_Crystal_Report STRFORMULA, reportname, CompanyName
Exit Sub

'IF YOU WANT TO DO A CSV FILE
Command1_Click
End Sub

Private Sub Command6_Click()
frmcmbine.Show vbModal
End Sub

Private Sub Command7_Click()
    '//kimberbalancesheet
    reportname = "BalanceSheeet.rpt"
    STRFORMULA = ""
    Show_Sales_Crystal_Report STRFORMULA, reportname, CompanyName
Exit Sub

'IF YOU WANT TO DO A CSV FILE
Command1_Click

End Sub

Private Sub Form_Load()
    DTPStartDate = Date
    dtpFinishDate = Date
End Sub

