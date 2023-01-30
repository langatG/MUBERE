VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmargive 
   Caption         =   "Argive suppliers"
   ClientHeight    =   3795
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   ScaleHeight     =   3795
   ScaleWidth      =   5910
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cmdargrpt 
      Caption         =   "Argive Report"
      Height          =   735
      Left            =   4320
      TabIndex        =   9
      Top             =   2280
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DTPedate 
      Height          =   375
      Left            =   1800
      TabIndex        =   8
      Top             =   1080
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      Format          =   38862849
      CurrentDate     =   42965
   End
   Begin MSComCtl2.DTPicker DTPstdate 
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Top             =   360
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      Format          =   38862849
      CurrentDate     =   42965
   End
   Begin VB.TextBox txtsno 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton Cmdclose 
      Caption         =   "Close"
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Activate"
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton Cmdargive 
      Caption         =   "&Argive "
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Sno"
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label lblstartdate 
      Caption         =   "Start date"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label lbledate 
      Caption         =   "End date"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   1080
      Width           =   1095
   End
End
Attribute VB_Name = "frmargive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmdargive_Click()
Dim lastdate As Date
Dim lastdateofsale As Date
Dim pcode As String
Dim NetPay As Double
Dim dy As Integer
Dim grade As String
Dim bank As String
Dim bcode As String
Dim BBranch As String
Dim rsd As New ADODB.Recordset
'check the user
sql = "SELECT     UserLoginIDs,levels, UserGroup, SUPERUSER From UserAccounts where UserLoginIDs='" & User & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
If rs!Levels <> "Manager" And rs!Levels <> "Accounts" Then
MsgBox "You are not allowed to Approve", vbInformation
Exit Sub
End If
End If

sql = ""
'sql = "set dateformat dmy SELECT     s.SNo,s.Names,s.AccNo,s.Bcode,s.BBranch,d.Remarks, SUM(Amount) AS Netpay From d_supplier_deduc d inner join d_Suppliers s on d.sno=s.sno WHERE   d.Date_Deduc >= '" & DTPstdate & "' and d.Date_Deduc <= '" & DTPedate & "' and d.Remarks LIKE '%bonus%' GROUP BY s.sno, s.names,s.AccNo,s.Bcode,s.BBranch,d.Remarks ORDER BY s.sno asc"
sql = "set dateformat dmy SELECT     SNo, active, Names From d_Suppliers ORDER BY SNo"
Set rs = oSaccoMaster.GetRecordset(sql)
While Not rs.EOF
pcode = rs!sno
'pcode = 3
'NetPay = rs!NetPay
'pname = rs!NAMES
'bank = rs!ACCNO
'bcode = rs!bcode
'BBranch = rs!BBranch

'select pcode,ldate,dy,auditdate,audit,grade from ag_paging
sql = ""
sql = "set dateformat dmy SELECT     SNo, SUM(QSupplied) AS qty From d_Milkintake Where (transdate >= '" & DTPstdate & "') And (sno = " & pcode & ") And (transdate <= '" & DTPedate & "')GROUP BY SNo"

Set rsd = oSaccoMaster.GetRecordset(sql)
If Not rsd.EOF Then
'sql = "update d_Suppliers set active=0 where sno=" & pcode & ""
Else


sql = "update d_Suppliers set active=0 where sno=" & pcode & ""
oSaccoMaster.ExecuteThis (sql)
End If


'sql = "update d_Suppliers set active=0 where sno=" & pcode & ""
'oSaccoMaster.ExecuteThis (sql)
rs.MoveNext

Wend
MsgBox "Records successfully Archived", vbInformation

End Sub

Private Sub Cmdargrpt_Click()
reportname = "Argive sup.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub Command1_Click()
'check the user
sql = "SELECT     UserLoginIDs,levels, UserGroup, SUPERUSER From UserAccounts where UserLoginIDs='" & User & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
If rs!Levels <> "Manager" And rs!Levels <> "Accounts" Then
MsgBox "You are not allowed to Approve", vbInformation
Exit Sub
End If
End If
If txtSNo = "" Then
MsgBox "enter supplier number", vbInformation
Exit Sub
End If
sql = "update d_Suppliers set active=1 where sno=" & txtSNo & ""
oSaccoMaster.ExecuteThis (sql)
MsgBox "Records successfully done", vbInformation
txtSNo = ""
End Sub
