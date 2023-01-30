VERSION 5.00
Begin VB.Form frmsychronize 
   Caption         =   "Synchronize Stocks"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmsychronize.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdrunall 
      Caption         =   "Run All"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "All"
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   2280
      Width           =   1095
   End
   Begin VB.ComboBox cboproductname 
      Height          =   315
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   480
      Width           =   2175
   End
   Begin VB.CommandButton cmdrun 
      Caption         =   "Run"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Customer No."
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "frmsychronize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public SelectedDsn As String
Public Provider As String
Dim rssy As ADODB.Recordset
Dim Y As Integer

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdrun_Click()
Set rst = New Recordset
Dim openbal As Double
Dim chg As Double
Dim bal As Double
Dim I As Integer
Dim pcode
Provider = cn
Set cn = New ADODB.Connection
cn.Open Provider, "atm", "atm"
sql = ""
sql = "SELECT p_code  From ag_stockbalance1 ORDER BY p_code"
rst.Open sql, cn, adOpenKeyset, adLockOptimistic
'While Not rst.EOF
'If Not IsNull(rst.Fields(0)) Then pcode = rst.Fields(0)
'// open it again but not order by track id again
sql = ""
sql = "SELECT * From ag_stockbalance1  where p_code='" & cboproductname & "' ORDER BY transdate,   trackid asc"
Set rssy = New ADODB.Recordset
rssy.Open sql, cn, adOpenKeyset, adLockOptimistic
I = 1
While Not rssy.EOF
I = I
'// update the balance
'// if it is the first one then leave it
If I = 1 Then '// get the openning balance
If Not IsNull(rssy.Fields("ag_stockbalance")) Then openbal = rssy.Fields("ag_stockbalance")
Else
chg = rssy.Fields("changeinstock")
bal = openbal + rssy.Fields("changeinstock")
sql = ""
sql = "update ag_stockbalance1 set openningstock=" & openbal & ",changeinstock=" & chg & ",stockbalance=" & bal & " where trackid=" & rssy.Fields("trackid") & ""
cn.Execute sql
openbal = bal
End If
I = I + 1
rssy.MoveNext
Wend
If cboproductname <> "" Then
sql = ""
sql = "update ag_products1 set qout=" & bal & " where p_code='" & cboproductname & "'"
cn.Execute sql
End If
I = 0
'rst.Requery
'Set rssy = Nothing
'//THE FINISH LINE
'// ONLY THE MASTER ALONE.
'Wend
MsgBox "Process Complete", vbInformation

End Sub

Private Sub cmdrunall_Click()
Set rst = New Recordset
Dim openbal As Double
Dim chg As Double
Dim bal As Double
Dim I As Integer
Dim pcode
Provider = cn
Set cn = New ADODB.Connection
cn.Open Provider, "atm", "atm"
sql = ""
sql = "SELECT distinct p_code  From ag_stockbalance1 ORDER BY p_code"
rst.Open sql, cn, adOpenKeyset, adLockOptimistic
Dim pp As Integer
pp = 0
While Not rst.EOF
pp = pp + 1

rst.MoveNext
Wend

Dim noofstock As Integer
noofstock = pp

rst.Requery

While Not rst.EOF

If Not IsNull(rst.Fields(0)) Then pcode = rst.Fields(0)
'// open it again but not order by track id again
            sql = ""
            sql = "SELECT * From ag_stockbalance1  where p_code='" & pcode & "' ORDER BY transdate,   trackid asc"
            Set rssy = New ADODB.Recordset
            rssy.Open sql, cn, adOpenKeyset, adLockOptimistic
            Dim rcount As Integer
            rcount = rssy.RecordCount
            I = 1
            While Not rssy.EOF
            I = I
            '// update the balance
            '// if it is the first one then leave it
            If I = 1 Then '// get the openning balance
            Dim openbal1 As Long
            If Not IsNull(rssy.Fields("ag_stockbalance")) Then openbal = rssy.Fields("ag_stockbalance")
             If Not IsNull(rssy.Fields("changeinstock")) Then openbal1 = rssy.Fields("changeinstock")
             If openbal = openbal1 Then openbal = openbal Else openbal = openbal1
            
            sql = ""
            sql = "update ag_stockbalance1 set openningstock=0,changeinstock=" & openbal & ",stockbalance=" & openbal & " where trackid=" & rssy.Fields("trackid") & ""
            cn.Execute sql
            
            Else
            chg = rssy.Fields("changeinstock")
            bal = openbal + rssy.Fields("changeinstock")
            sql = ""
            sql = "update ag_stockbalance1 set openningstock=" & openbal & ",changeinstock=" & chg & ",stockbalance=" & bal & " where trackid=" & rssy.Fields("trackid") & ""
            cn.Execute sql
            openbal = bal
            End If
            I = I + 1
            rssy.MoveNext
            Wend
            If rcount = 1 Then
            sql = ""
            sql = "update ag_products1 set qout=" & openbal & " where p_code='" & pcode & "'"
            cn.Execute sql
            Else
            sql = ""
            sql = "update ag_products1 set qout=" & bal & " where p_code='" & pcode & "'"
            cn.Execute sql
            End If
            I = 0


rst.MoveNext
Wend

MsgBox "Process Complete", vbInformation

End Sub

Private Sub Form_Load()
Provider = cn
Set cn = New ADODB.Connection
cn.Open Provider, "atm", "atm"
'If reportpath = "" Then reportpath = GetSetting("payroll", "AppName", "rptPath", rptPath)
sql = "select P_code  from ag_products1 ORDER BY P_code ASC"
Set rs = New ADODB.Recordset
rs.Open sql, cn

While Not rs.EOF
cboproductname.AddItem rs.Fields(0)
rs.MoveNext
Wend
End Sub
