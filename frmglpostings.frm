VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmglpostings 
   Caption         =   "AC - General Ledger Postings"
   ClientHeight    =   2625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6810
   Icon            =   "frmglpostings.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2625
   ScaleWidth      =   6810
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      Height          =   375
      Left            =   3600
      TabIndex        =   8
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CommandButton cmdpost 
      Caption         =   "Post"
      Height          =   375
      Left            =   1080
      TabIndex        =   7
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Frame Frame3 
      Height          =   1575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6375
      Begin VB.TextBox txtrate 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3840
         TabIndex        =   10
         Text            =   "1"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton cmdfinder 
         Height          =   285
         Left            =   5040
         Picture         =   "frmglpostings.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Add New record"
         Top             =   1200
         Width           =   375
      End
      Begin VB.ComboBox cbocurrency 
         Height          =   315
         ItemData        =   "frmglpostings.frx":0704
         Left            =   3840
         List            =   "frmglpostings.frx":071A
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   840
         Width           =   2175
      End
      Begin VB.ComboBox cboyearperiod 
         Height          =   315
         ItemData        =   "frmglpostings.frx":073C
         Left            =   3840
         List            =   "frmglpostings.frx":0797
         TabIndex        =   1
         Text            =   "cboyearperiod"
         Top             =   360
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker txttransdate 
         Height          =   255
         Left            =   720
         TabIndex        =   3
         ToolTipText     =   "Month Calender"
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         Format          =   130416641
         CurrentDate     =   38856
      End
      Begin VB.Label Label30 
         Caption         =   "Rate"
         Height          =   255
         Left            =   2520
         TabIndex        =   11
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "Currency"
         Height          =   255
         Left            =   2520
         TabIndex        =   6
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label14 
         Caption         =   "Year/Period"
         Height          =   255
         Left            =   2520
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label15 
         Caption         =   "Date"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmglpostings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdfinder_Click()
On Error Resume Next
frmsearchcurr.Show vbModal
Dim Y As String
Y = sel
'm = False
If Y <> "" Then
     Dim cn As Connection
    Set cn = New ADODB.Connection
    Provider = SelectedDsn
   cn.Open Provider, "atm", "atm"
sql = ""
'If reportpath = "" Then reportpath = GetSetting("payroll", "AppName", "rptPath", rptPath)
sql = "SELECT     rateaganistsource,currcode From Curr where currcode='" & Y & "'"
Set rs = New ADODB.Recordset
rs.Open sql, cn
If Not rs.EOF Then
If Not IsNull(rs.Fields(0)) Then txtRate = (rs.Fields(0))
If Not IsNull(rs.Fields(1)) Then cbocurrency = Trim((rs.Fields(1)))
End If


End If
End Sub

Private Sub cmdPost_Click()
'// START WITH PAYABLES FIRST.
Set cn = CreateObject("adodb.connection")
Dim sourcecode As String
Dim myclass As cdbase
sourcecode = "AP"
Set myclass = New cdbase
Set cn = CreateObject("adodb.connection")
Provider = myclass.OpenCon
cn.Open Provider, "atm", "atm"
Set rs = CreateObject("ADODB.Recordset")
 Dim desc As String
 Dim id
   Dim rst As Recordset, rss As Recordset
    Dim name1  As String, name2  As String, AVAIL1 As Currency, AVAIL2 As Currency
    
sql = ""
sql = "set dateformat dmy SELECT    accno AS DR,  p_accno AS CR,p_amount,  particulars, chequeno AS chq, Purpose, paidto,pid  FROM         AP WHERE     (posted = 0) AND (chequestatus = 'Proceed')"
rs.Open sql, cn
    If Not rs.EOF Then
    While Not rs.EOF
    '// get the details of dr and cr accounts
   id = rs.Fields(7)
    desc = rs.Fields(3) & "-" & rs.Fields(4)
 
    Set rst = New ADODB.Recordset
    
    sql = "SELECT     AccNo, Name,availablebalance FROM         CUB where accno='" & Trim(rs.Fields(0)) & "'"
    rst.Open sql, cn, adOpenKeyset, adLockOptimistic
    If Not rst.EOF Then
    If Not IsNull(rst.Fields(1)) Then name1 = rst.Fields(1)
    If Not IsNull(rst.Fields(2)) Then AVAIL1 = rst.Fields(2)
    Else
    MsgBox "Account does not exist,No Posting is allowed Further ", vbCritical, "AC"
    Exit Sub
    End If
    '//get for the cr
    sql = "SELECT     AccNo, Name,availablebalance FROM         CUB where accno='" & Trim(rs.Fields(1)) & "'"
    Set rss = New ADODB.Recordset
    rss.Open sql, cn, adOpenKeyset, adLockOptimistic
    If Not rss.EOF Then
    If Not IsNull(rss.Fields(1)) Then name2 = rss.Fields(1)
    If Not IsNull(rss.Fields(2)) Then AVAIL1 = rss.Fields(2)
    Else
    MsgBox "Account does not exist,No Posting is allowed Further ", vbCritical, "AC"
    Exit Sub
    End If
    '// insert into customer balance   start with the DR.
sql = ""
Dim scurr As Currency
scurr = (rs.Fields(2) / txtRate)
If Not Save_GLTRANSACTION(Format(txttransdate, "dd/mm/yyyy"), CDbl(rs.Fields(2)), rs.Fields(0), rs.Fields(1), desc, rs.Fields(5), User, ErrorMessage, desc, 1, 1, rs.Fields(0), TransNo) Then
            If ErrorMessage <> "" Then
                MsgBox ErrorMessage, vbInformation, Me.Caption
                ErrorMessage = ""
            End If
        End If
'sql = "set dateformat dmy insert into customerbalance (customerno,accname,amount,availablebalance,accno,transdescription,transdate,commission,period,transType,posted,locked,status,vno,auditid,auditdate,moduleid,accd,s_code,curr,yyear,rate,dr_scurr,cr_scurr,sbal) "
'sql = sql & " values ('" & rs.Fields(0) & "','" & name1 & "'," & rs.Fields(2) & "," & AVAIL1 + rs.Fields(2) & ",'" & rs.Fields(0) & "','" & DESC & "','" & Format(txttransdate, "dd/mm/yyyy") & "',0,'" & month(txttransdate) & "','DR',0,0,0,'" & rs.Fields(5) & "','" & User & "','" & Get_Server_Date & "','3','" & rs.Fields(1) & "','AP','" & cbocurrency & "'," & cboyearperiod & "," & txtRate & "," & scurr & "," & scurr & "," & scurr & " )"
'cn.Execute sql
'
'sql = ""
'sql = "set dateformat dmy update cub set amount=" & rs.Fields(2) & ",Active=1,transdescription='" & DESC & "',availablebalance=" & AVAIL1 + rs.Fields(2) & ",transdate='" & Format(txttransdate, "dd/mm/yyyy") & "',vno='" & rs.Fields(5) & "',period='" & month(txttransdate) & "',auditid='" & User & "',auditdate='" & Now & "',moduleid=2 where accno='" & rs.Fields(0) & "'"
'cn.Execute sql
''// do the credit before proceeding
'sql = ""
'
'sql = "set dateformat dmy insert into customerbalance (customerno,accname,amount,availablebalance,accno,transdescription,transdate,commission,period,transType,posted,locked,status,vno,auditid,auditdate,moduleid,accd,s_code,curr,yyear,rate,dr_scurr,cr_scurr,sbal) "
'sql = sql & " values ('" & rs.Fields(1) & "','" & name1 & "'," & rs.Fields(2) & "," & AVAIL2 + rs.Fields(2) & ",'" & rs.Fields(1) & "','" & DESC & "','" & Format(txttransdate, "dd/mm/yyyy") & "',0,'" & month(txttransdate) & "','CR',0,0,0,'" & rs.Fields(5) & "','" & User & "','" & Get_Server_Date & "','3','" & rs.Fields(0) & "','AP','" & cbocurrency & "'," & cboyearperiod & "," & txtRate & "," & scurr & "," & scurr & "," & scurr & "  )"
'cn.Execute sql
'
'sql = ""
'sql = "set dateformat dmy update cub set amount=" & rs.Fields(2) & ",Active=1,transdescription='" & DESC & "',availablebalance=" & AVAIL2 + rs.Fields(2) & ",transdate='" & Format(txttransdate, "dd/mm/yyyy") & "',vno='" & rs.Fields(5) & "',period='" & month(txttransdate) & "',auditid='" & User & "',auditdate='" & Now & "',moduleid=2 where accno='" & rs.Fields(1) & "'"
'cn.Execute sql

sql = ""
sql = "set dateformat dmy update ap set posted=1 where pid =" & id & ""
cn.Execute sql
    '// insert into cub
    rs.MoveNext
    Wend
    Else
    MsgBox "No Records To Post", vbInformation
    Exit Sub
    End If

'// WORK ON ACCOUNTS RECEIVABLES.

MsgBox "You have successfully posted the Records"

End Sub

Private Sub Form_Load()
txttransdate = Format(Get_Server_Date, "dd/mm/yyyy")
cbocurrency = "KES"
cboyearperiod = Format(Get_Server_Date, "YYYY") & "-" & Format(Get_Server_Date, "MM")
End Sub
