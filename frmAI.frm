VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAI 
   Caption         =   "ALL AGROVET  Reports"
   ClientHeight    =   4785
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5625
   LinkTopic       =   "Form1"
   ScaleHeight     =   4785
   ScaleWidth      =   5625
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "PRINT AGROVET REPORTS HERE:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      Begin VB.CommandButton cmdstock1 
         Caption         =   "Stock Receive"
         Height          =   735
         Left            =   2040
         TabIndex        =   11
         Top             =   3840
         Width           =   1335
      End
      Begin VB.CommandButton cmdreportagr 
         Caption         =   "Monthly sales per Farmers"
         Height          =   735
         Left            =   240
         TabIndex        =   10
         Top             =   3840
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton cmddispatch 
         Caption         =   "Dispatch Report"
         Height          =   615
         Left            =   3720
         TabIndex        =   9
         Top             =   3120
         Width           =   1335
      End
      Begin VB.CommandButton cmdcashstaff 
         Caption         =   "Cash Staff Report"
         Height          =   615
         Left            =   240
         TabIndex        =   8
         Top             =   3120
         Width           =   1455
      End
      Begin VB.CommandButton cmdanalysis 
         Caption         =   "Sales Analysis"
         Height          =   615
         Left            =   2040
         TabIndex        =   7
         Top             =   3120
         Width           =   1335
      End
      Begin VB.CommandButton cmdbalance 
         Caption         =   "Stock Balance Report"
         Height          =   735
         Left            =   3720
         TabIndex        =   6
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CommandButton cmdall 
         Caption         =   "All Sales Report"
         Height          =   735
         Left            =   2040
         TabIndex        =   5
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CommandButton cmdstaff 
         Caption         =   "Staff Report"
         Height          =   735
         Left            =   240
         TabIndex        =   4
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CommandButton cmdcheck 
         Caption         =   "Check Off Reports"
         Height          =   735
         Left            =   3720
         TabIndex        =   3
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton cmdmpesa 
         Caption         =   "Mpesa Report"
         Height          =   735
         Left            =   2040
         TabIndex        =   2
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton cmdash 
         BackColor       =   &H0080FF80&
         Caption         =   "Cash Report"
         Height          =   735
         Left            =   240
         MaskColor       =   &H00FFFF00&
         TabIndex        =   1
         Top             =   1200
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   3360
         TabIndex        =   12
         Top             =   600
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         Format          =   172752897
         CurrentDate     =   44283
      End
      Begin VB.Label Label1 
         Caption         =   "Select the month for the report"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   13
         Top             =   720
         Width           =   2775
      End
   End
End
Attribute VB_Name = "frmAI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdall_Click()
    reportname = "all agrovet sales.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub cmdanalysis_Click()
    reportname = "Sales analysis.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub cmdash_Click()
    reportname = "CASH agrovet sales.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub cmdbalance_Click()
Dim ans As String
ans = MsgBox("Do you Want List of all Branches", vbYesNo)
If ans = vbYes Then
 reportname = "d_StockBal1.rpt"
 Else
 reportname = "d_StockBal.rpt"
 End If
  Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub cmdcashstaff_Click()
    reportname = "Agrovet staffc.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub cmdcheck_Click()
    reportname = "CHECK OFF agrovet sales.rpt"
    Show_Sales_Crystal_Report "", reportname, ""
End Sub
Private Sub cmddispatch_Click()

'Startdate = DateSerial(year(DTPicker1), month(DTPicker1), 1)
'Enddate = DateSerial(year(DTPicker1), month(DTPicker1) + 1, 1)

'sql = ""
'sql = "set dateformat dmy delete from ag_ReceiptsProcess1 where Date>= '" & Startdate & "' And Date<'" & Enddate & "'"
'cn.Execute sql
'
'Dim U As Integer
'Dim C As String
'sql = ""
'sql = "set dateformat dmy Select count(distinct(Branch)) as u from DRAWNSTOCK where Date >= '" & Startdate & "' And Date<'" & Enddate & "' "
'Set rs = cn.Execute(sql)
'
'  U = rs.Fields(0)
''MsgBox "Please wait " & U & ""
'  sql = ""
'  sql = "set dateformat dmy Select distinct(Branch) as y  from   DRAWNSTOCK  where Date >= '" & Startdate & "' And Date<'" & Enddate & "' order by Branch asc  "
'  Set rsg = cn.Execute(sql)
'  While Not rs.EOF
'  Do While Not U <= 0
'  If Not rsg.EOF Then
'    C = rsg.Fields(0)
'
'       sql = ""
'       sql = "set dateformat dmy Select distinct(PRODUCTNAME) from DRAWNSTOCK where Branch='" & C & "' and Date >= '" & Startdate & "' And Date<'" & Enddate & "' "
'       Set rsb = cn.Execute(sql)
'       sql = ""
'       sql = "set dateformat dmy Select  count(distinct(PRODUCTNAME))from DRAWNSTOCK where Branch='" & C & "' and Date >= '" & Startdate & "' And Date<'" & Enddate & "'"
'       Set rsh = cn.Execute(sql)
'           sql = ""
'           sql = "set dateformat dmy Select  PRODUCTID,PRODUCTNAME,QUANTITY, DATE,Branch from DRAWNSTOCK where Branch='" & C & "' and PRODUCTNAME='" & rsb.Fields(0) & "' and Date >= '" & Startdate & "' And Date<'" & Enddate & "' order by PRODUCTID asc"
'           Set rst = cn.Execute(sql)
'       Do While Not rsh.EOF
'      If Not rsb.EOF Then
'       If Not rsh.EOF Then
'
'         If Not rst.EOF Then
'         'Do While Not rst.EOF
'
'           sql = ""
'           sql = "set dateformat dmy select * from ag_ReceiptsProcess1 where Branch='" & C & "' and Date >= '" & Startdate & "' And Date<'" & Enddate & "'"
'           Set rss = oSaccoMaster.GetRecordset(sql)
'           If rss.EOF Then
'             sql = ""
'             sql = "set dateformat dmy insert into  ag_ReceiptsProcess1(Date, bPro1, bPro2, bPro3, bPro4, bPro5, bPro6, bPro7, Branch)"
'             sql = sql & "  values('" & rst.Fields(3) & "','0','0','0','0','0','0','0','" & rst.Fields(4) & "')"
'             cn.Execute sql
'            Else
'            End If
'           sql = ""
'           sql = "set dateformat dmy select bPro1, bPro2, bPro3, bPro4, bPro5, bPro6, bPro7,Branch from ag_ReceiptsProcess1 where Branch='" & C & "' and Date >= '" & Startdate & "' And Date<'" & Enddate & "'"
'           Set rsl = oSaccoMaster.GetRecordset(sql)
'
'          sql = ""
'           sql = "select PRODUCTID,PRODUCTNAME from DRAWNSTOCK where Branch='" & C & "' and PRODUCTNAME='" & rst.Fields(1) & "' ORDER BY PRODUCTID asc"
'           Set rsm = oSaccoMaster.GetRecordset(sql)
'           If Not rsm.EOF Then
'           Dim strong As Integer
'           strong = rsm.Fields(0)
'            Select Case strong
'             Case "1"
'              sql = ""
'              sql = "set dateformat dmy Update ag_ReceiptsProcess1 SET bPro1 ='" & rsl.Fields(0) + rst.Fields(2) & "' WHERE Branch='" & rst.Fields(4) & "' and Branch ='" & rsl.Fields(7) & "' and Date >= '" & Startdate & "'And Date<'" & Enddate & "'"
'              cn.Execute sql
'             Case "2"
'              sql = ""
'              sql = "set dateformat dmy Update ag_ReceiptsProcess1 SET bPro2 ='" & rsl.Fields(1) + rst.Fields(2) & "' WHERE Branch='" & rst.Fields(4) & "' and Branch ='" & rsl.Fields(7) & "' and Date >= '" & Startdate & "'And Date<'" & Enddate & "'"
'              cn.Execute sql
'             Case "3"
'              sql = ""
'              sql = "set dateformat dmy Update ag_ReceiptsProcess1 SET bPro3 ='" & rsl.Fields(2) + rst.Fields(2) & "' WHERE Branch='" & rst.Fields(4) & "'and Branch ='" & rsl.Fields(7) & "' and Date >= '" & Startdate & "'And Date<'" & Enddate & "'"
'              cn.Execute sql
'             Case "4"
'              sql = ""
'              sql = "set dateformat dmy Update ag_ReceiptsProcess1 SET bPro4 ='" & rsl.Fields(3) + rst.Fields(2) & "' WHERE Branch='" & rst.Fields(4) & "' and Branch ='" & rsl.Fields(7) & "' and Date >= '" & Startdate & "'And Date<'" & Enddate & "'"
'              cn.Execute sql
'             Case "5"
'              sql = ""
'              sql = "set dateformat dmy Update ag_ReceiptsProcess1 SET bPro5 ='" & rsl.Fields(4) + rst.Fields(2) & "' WHERE Branch='" & rst.Fields(4) & "' and Branch ='" & rsl.Fields(7) & "' and Date >= '" & Startdate & "'And Date<'" & Enddate & "'"
'              cn.Execute sql
'             Case "6"
'              sql = ""
'              sql = "set dateformat dmy Update ag_ReceiptsProcess1 SET bPro6 ='" & rsl.Fields(5) + rst.Fields(2) & "' WHERE Branch='" & rst.Fields(4) & "' and Branch ='" & rsl.Fields(7) & "' and Date >= '" & Startdate & "'And Date<'" & Enddate & "'"
'              cn.Execute sql
'             Case "7"
'              sql = ""
'              sql = "set dateformat dmy Update ag_ReceiptsProcess1 SET bPro7 ='" & rsl.Fields(6) + rst.Fields(2) & "' WHERE Branch='" & rst.Fields(4) & "' and Branch ='" & rsl.Fields(7) & "' and Date >= '" & Startdate & "'And Date<'" & Enddate & "'"
'              cn.Execute sql
'             Case Else
'            End Select
'
'           End If
'           'Loop
'           rst.MoveNext
'         End If
'           rsb.MoveNext
'           End If
'         Else
'           rsh.MoveNext
'          End If
'         Loop
'      Else
'    End If
'   U = U - 1
' rsg.MoveNext
'Loop
'Wend
'
'Dim ans As String
'ans = MsgBox("Do you Want Dranstock sammary?", vbYesNo)
' If ans = vbYes Then
'   reportname = "COMBINESALES1.rpt"
' Else
   reportname = "Transfer agrovet sales1.rpt"
' End If
  Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub cmdinstitution_Click()
    reportname = "Others.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub cmdfeedsre_Click()
    reportname = "feeds.rpt"
    Show_Sales_Crystal_Report "", reportname, ""
End Sub

Private Sub cmdmpesa_Click()
    reportname = "Mpesa agrovet sales.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub cmdSnpartly_Click()
    reportname = "CheckOff_PartlyPayment.rpt"
    Show_Sales_Crystal_Report "", reportname, ""
End Sub

Private Sub cmdOk_Click()
txtransdate = "& txtransdate &"
fraAlteradv.Visible = False

End Sub

Private Sub cmdreportagr_Click()
Dim ans As String
ans = MsgBox("Do you Want Summary Report as per Outlet?", vbYesNo)
 If ans = vbYes Then
     reportname = "COMBINESALES1.rpt"
 Else
   reportname = "COMBINESALES.rpt"
 End If
' Show_Sales_Crystal_Report STRFORMULA, reportname, ""
    Show_Sales_Crystal_Report "", reportname, ""
    
End Sub

Private Sub cmdstaff_Click()
    reportname = "Agrovet staffs.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub cmdstoock_Click()
    reportname = "stoock Report.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, ""
  
  End Sub

Private Sub cmdTransp_Click()
    reportname = "tran_PartlyPayment.rpt"
    Show_Sales_Crystal_Report "", reportname, ""
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub cmdstock1_Click()
    reportname = "stoock Report.rpt"
    Show_Sales_Crystal_Report "", reportname, ""
End Sub

