VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmassetsdisposals 
   Caption         =   "AC -Assets Disposals"
   ClientHeight    =   9165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9705
   Icon            =   "frmassetsdisposals.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9165
   ScaleWidth      =   9705
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdpost 
      Caption         =   "Post"
      Height          =   375
      Left            =   3000
      TabIndex        =   20
      Top             =   8880
      Width           =   1095
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      Height          =   375
      Left            =   5760
      TabIndex        =   19
      Top             =   8880
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Height          =   8775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9495
      Begin VB.TextBox txtlosscr 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2160
         MaxLength       =   15
         TabIndex        =   73
         Top             =   8400
         Width           =   1695
      End
      Begin VB.TextBox txtamount10 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7680
         TabIndex        =   72
         Top             =   8400
         Width           =   1695
      End
      Begin VB.CommandButton Command12 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         Height          =   285
         Left            =   3840
         Picture         =   "frmassetsdisposals.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   71
         ToolTipText     =   "Add New record"
         Top             =   8400
         Width           =   375
      End
      Begin VB.TextBox txtproftdr 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2160
         MaxLength       =   15
         TabIndex        =   68
         Top             =   7440
         Width           =   1695
      End
      Begin VB.TextBox txtamount9 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7680
         TabIndex        =   67
         Top             =   7440
         Width           =   1695
      End
      Begin VB.CommandButton Command11 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         Height          =   285
         Left            =   3840
         Picture         =   "frmassetsdisposals.frx":0704
         Style           =   1  'Graphical
         TabIndex        =   66
         ToolTipText     =   "Add New record"
         Top             =   7440
         Width           =   375
      End
      Begin VB.TextBox txtvno 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1680
         TabIndex        =   65
         Top             =   2040
         Width           =   5295
      End
      Begin VB.TextBox txtbankdr 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2160
         MaxLength       =   15
         TabIndex        =   59
         Top             =   2640
         Width           =   1695
      End
      Begin VB.TextBox txtamount7 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7680
         TabIndex        =   58
         Top             =   2640
         Width           =   1695
      End
      Begin VB.CommandButton Command10 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         Height          =   285
         Left            =   3840
         Picture         =   "frmassetsdisposals.frx":09C6
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "Add New record"
         Top             =   2640
         Width           =   375
      End
      Begin VB.TextBox txtdisposalbankcr 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2160
         MaxLength       =   15
         TabIndex        =   56
         Top             =   3240
         Width           =   1695
      End
      Begin VB.TextBox txtamount8 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7680
         TabIndex        =   55
         Top             =   3240
         Width           =   1695
      End
      Begin VB.CommandButton Command9 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         Height          =   285
         Left            =   3840
         Picture         =   "frmassetsdisposals.frx":0C88
         Style           =   1  'Graphical
         TabIndex        =   54
         ToolTipText     =   "Add New record"
         Top             =   3240
         Width           =   375
      End
      Begin VB.TextBox txtrate 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7920
         TabIndex        =   52
         Text            =   "1"
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton Command6 
         Height          =   285
         Left            =   9120
         Picture         =   "frmassetsdisposals.frx":0F4A
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Add New record"
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton Command5 
         Height          =   285
         Left            =   3360
         Picture         =   "frmassetsdisposals.frx":120C
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Add New record"
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txtASSETSNO 
         DataField       =   "assetsno"
         Height          =   285
         Index           =   1
         Left            =   1530
         TabIndex        =   45
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox txtASSETSNAME 
         Appearance      =   0  'Flat
         DataField       =   "assetsname"
         Height          =   285
         Index           =   1
         Left            =   5370
         TabIndex        =   44
         Top             =   1350
         Width           =   3975
      End
      Begin VB.TextBox txtSERIALNO 
         Appearance      =   0  'Flat
         DataField       =   "assetserialno"
         Height          =   285
         Index           =   1
         Left            =   1530
         TabIndex        =   43
         Top             =   1635
         Width           =   3975
      End
      Begin VB.TextBox txtNETREALIABLEVALUE 
         Height          =   285
         Left            =   7050
         TabIndex        =   42
         Top             =   2040
         Width           =   2055
      End
      Begin VB.CommandButton cmdfinder 
         Height          =   285
         Left            =   1530
         Picture         =   "frmassetsdisposals.frx":14CE
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Add New record"
         Top             =   1320
         Width           =   375
      End
      Begin VB.CommandButton Command4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         Height          =   285
         Left            =   3840
         Picture         =   "frmassetsdisposals.frx":1790
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Add New record"
         Top             =   7920
         Width           =   375
      End
      Begin VB.TextBox txtamount6 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7680
         TabIndex        =   37
         Top             =   7920
         Width           =   1695
      End
      Begin VB.TextBox txtloss 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2160
         MaxLength       =   15
         TabIndex        =   36
         Top             =   7920
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         Height          =   285
         Left            =   3840
         Picture         =   "frmassetsdisposals.frx":1A52
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Add New record"
         Top             =   6960
         Width           =   375
      End
      Begin VB.TextBox txtamount5 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7680
         TabIndex        =   32
         Top             =   6960
         Width           =   1695
      End
      Begin VB.TextBox txtprofit 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2160
         MaxLength       =   15
         TabIndex        =   31
         Top             =   6960
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         Height          =   285
         Left            =   3840
         Picture         =   "frmassetsdisposals.frx":1D14
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Add New record"
         Top             =   6120
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         Height          =   285
         Left            =   3840
         MaskColor       =   &H000000FF&
         Picture         =   "frmassetsdisposals.frx":1FD6
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Add New record"
         Top             =   5520
         Width           =   375
      End
      Begin VB.TextBox txtamount4 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7680
         TabIndex        =   24
         Top             =   6120
         Width           =   1695
      End
      Begin VB.TextBox txtamount3 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7680
         TabIndex        =   23
         Top             =   5520
         Width           =   1695
      End
      Begin VB.TextBox txtprovisionacc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2160
         MaxLength       =   15
         TabIndex        =   22
         Top             =   5520
         Width           =   1695
      End
      Begin VB.TextBox txtdisposalcr 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2160
         MaxLength       =   15
         TabIndex        =   21
         Top             =   6120
         Width           =   1695
      End
      Begin VB.CommandButton Command8 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         Height          =   285
         Left            =   3840
         Picture         =   "frmassetsdisposals.frx":2298
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Add New record"
         Top             =   4680
         Width           =   375
      End
      Begin VB.CommandButton Command7 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         Height          =   285
         Left            =   3840
         MaskColor       =   &H000000FF&
         Picture         =   "frmassetsdisposals.frx":255A
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Add New record"
         Top             =   4080
         Width           =   375
      End
      Begin VB.TextBox txtamount2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7680
         TabIndex        =   7
         Top             =   4680
         Width           =   1695
      End
      Begin VB.TextBox txtamount1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7680
         TabIndex        =   6
         Top             =   4080
         Width           =   1695
      End
      Begin VB.TextBox txtassetsacc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2160
         MaxLength       =   15
         TabIndex        =   5
         Top             =   4080
         Width           =   1695
      End
      Begin VB.TextBox txtdisposaldr 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2160
         MaxLength       =   15
         TabIndex        =   4
         Top             =   4680
         Width           =   1695
      End
      Begin VB.ComboBox cbocurrency 
         Height          =   315
         ItemData        =   "frmassetsdisposals.frx":281C
         Left            =   7320
         List            =   "frmassetsdisposals.frx":2832
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Width           =   2175
      End
      Begin VB.ComboBox cboyearperiod 
         Height          =   315
         ItemData        =   "frmassetsdisposals.frx":2854
         Left            =   4560
         List            =   "frmassetsdisposals.frx":28AF
         TabIndex        =   2
         Text            =   "cboyearperiod"
         Top             =   360
         Width           =   1335
      End
      Begin VB.ComboBox cbosourcecode 
         Height          =   315
         ItemData        =   "frmassetsdisposals.frx":29F2
         Left            =   1530
         List            =   "frmassetsdisposals.frx":2A05
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   840
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker txttransdate 
         Height          =   255
         Left            =   1530
         TabIndex        =   10
         ToolTipText     =   "Month Calender"
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         Format          =   162529281
         CurrentDate     =   38856
      End
      Begin VB.Line Line5 
         BorderWidth     =   3
         X1              =   0
         X2              =   9480
         Y1              =   7800
         Y2              =   7800
      End
      Begin VB.Label lbllabel10 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4320
         TabIndex        =   75
         Top             =   8400
         Width           =   3255
      End
      Begin VB.Label Label17 
         Caption         =   "Account to CR"
         Height          =   255
         Left            =   120
         TabIndex        =   74
         Top             =   8400
         Width           =   2055
      End
      Begin VB.Label lbllabel9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4320
         TabIndex        =   70
         Top             =   7440
         Width           =   3255
      End
      Begin VB.Label Label7 
         Caption         =   "Account to DR"
         Height          =   255
         Left            =   120
         TabIndex        =   69
         Top             =   7440
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Ref/Receipts Details"
         Height          =   255
         Left            =   120
         TabIndex        =   64
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label lbllabel7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4320
         TabIndex        =   63
         Top             =   2640
         Width           =   3255
      End
      Begin VB.Label Label5 
         Caption         =   "Bank DR"
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   2640
         Width           =   1935
      End
      Begin VB.Label lbllabel8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4320
         TabIndex        =   61
         Top             =   3240
         Width           =   3255
      End
      Begin VB.Label Label1 
         Caption         =   "Disposal Of Asset  CR"
         Height          =   255
         Left            =   120
         TabIndex        =   60
         Top             =   3240
         Width           =   2055
      End
      Begin VB.Line Line4 
         BorderWidth     =   3
         X1              =   0
         X2              =   9480
         Y1              =   6720
         Y2              =   6720
      End
      Begin VB.Label Label30 
         Caption         =   "Rate"
         Height          =   255
         Left            =   6600
         TabIndex        =   53
         Top             =   720
         Width           =   1095
      End
      Begin VB.Line Line3 
         BorderWidth     =   3
         X1              =   0
         X2              =   9480
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Line Line2 
         BorderWidth     =   3
         X1              =   0
         X2              =   9480
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Line Line1 
         BorderWidth     =   3
         X1              =   0
         X2              =   9480
         Y1              =   5280
         Y2              =   5280
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Asset No:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   49
         Top             =   1320
         Width           =   690
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "RegNo/Serial No.:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   48
         Top             =   1650
         Width           =   1320
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Asset Name:"
         Height          =   195
         Index           =   3
         Left            =   3960
         TabIndex        =   47
         Top             =   1380
         Width           =   900
      End
      Begin VB.Label Label16 
         Caption         =   "Net Realizable Value"
         Height          =   255
         Left            =   7050
         TabIndex        =   46
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label8 
         Caption         =   " Loss on Disposal DR"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   7920
         Width           =   2055
      End
      Begin VB.Label Lblname6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4320
         TabIndex        =   39
         Top             =   7920
         Width           =   3255
      End
      Begin VB.Label Label6 
         Caption         =   "Profit On  Disposal CR"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   6960
         Width           =   1935
      End
      Begin VB.Label Lblname5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4320
         TabIndex        =   34
         Top             =   6960
         Width           =   3255
      End
      Begin VB.Label Label4 
         Caption         =   "Accumulated  Dep DR."
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   5520
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Disposal of of Assets  CR."
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   6120
         Width           =   2055
      End
      Begin VB.Label Lblname3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4320
         TabIndex        =   28
         Top             =   5520
         Width           =   3255
      End
      Begin VB.Label Lblname4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4320
         TabIndex        =   27
         Top             =   6120
         Width           =   3255
      End
      Begin VB.Label Label9 
         Caption         =   "Assets Account  CR."
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   4080
         Width           =   1935
      End
      Begin VB.Label Label10 
         Caption         =   "Disposal of Asset DR."
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   4680
         Width           =   1575
      End
      Begin VB.Label lblname1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4320
         TabIndex        =   16
         Top             =   4080
         Width           =   3255
      End
      Begin VB.Label Lblname2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4320
         TabIndex        =   15
         Top             =   4680
         Width           =   3255
      End
      Begin VB.Label Label12 
         Caption         =   "Currency"
         Height          =   255
         Left            =   6000
         TabIndex        =   14
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "Source Code"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label14 
         Caption         =   "Year/Period"
         Height          =   255
         Left            =   3240
         TabIndex        =   12
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label15 
         Caption         =   "Date"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmassetsdisposals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myclass As cdbase
Dim lab As String
Dim cn As Connection
Dim Provider As String
Dim BAL1 As Currency, bal2 As Currency, bal3 As Currency, bal4 As Currency, bal5 As Currency, bal6 As Currency, bal7 As Currency, bal8 As Currency, BAL9 As Currency, bal10 As Currency

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdfinder_Click()
On Error Resume Next
frmsearchre.Show vbModal
Dim Y As String
Y = sel

If Y <> "" Then
     Dim cn As Connection
    Set cn = New ADODB.Connection
    
    cn.Open frmODBCLogon.cboDSNList
sql = ""

sql = "SELECT     AssetsNo, AssetserialNo, AssetsName from assets where assetsno='" & Y & "' order by AssetsNo"
Set rs = New ADODB.Recordset
rs.Open sql, cn
If Not rs.EOF Then

If Not IsNull(rs.Fields(0)) Then txtASSETSNO(1) = (rs.Fields(0))
If Not IsNull(rs.Fields(1)) Then txtSERIALNO(1) = (rs.Fields(1))
If Not IsNull(rs.Fields(2)) Then txtASSETSNAME(2) = (rs.Fields(2))
'If Not IsNull(rs.Fields(2)) Then txtNETREALIABLEVALUE = (rs.Fields(3))

'get the serial no of the item

sql = "Select * from assetstrans where assetcode='" & txtASSETSNO(1) & "'"
rs.Open sql, cn
If rs.EOF Then
'MsgBox "No data in the such data available"
Exit Sub
Else
    'If .RecordCount > 0 Then
    txtNETREALIABLEVALUE = rs!nrv
    'txtserialno(1) = rs!assetserialno
    End If
'Call cboname_p

End If
End If
End Sub
Private Sub get_namecr(acc As String)

    Set cn = CreateObject("adodb.connection")
    Set myclass = New cdbase
    Provider = myclass.OpenCon
    cn.Open Provider, "atm", "atm"
    sql = ""
    sql = "select * from glsetup where accno='" & acc & "'"
    Set rs = New ADODB.Recordset
    rs.Open sql, cn, adOpenKeyset, adLockOptimistic
    If rs.EOF Then
    Else
    If Not IsNull(rs.Fields("glaccname")) Then lab = rs.Fields("glaccname")
    End If
End Sub

Private Sub cmdPost_Click()
On Error GoTo ErrorHandler
Set myclass = New cdbase
    Set cn = CreateObject("adodb.connection")
    Provider = myclass.OpenCon
   cn.Open Provider, "atm", "atm"
    '// check wether what you have is correct.
    If txtASSETSNO(1) = "" Then
    MsgBox "Asset to dispose required before you proceed", vbInformation
    Exit Sub
    End If
    If txtdisposalcr = "" Or txtdisposaldr = "" Or txtassetsacc = "" Or txtprovisionacc = "" Then
    MsgBox "Enter the Required accounts Before you proceed", vbInformation
    Exit Sub
    End If
    
    '//assets and disposal account
    '//check whether the item you are disposing is still active or has been disposed.
    
sql = ""
sql = "set dateformat dmy select * from assets where  assetsno='" & txtASSETSNO(1) & "'"
Set rs = New ADODB.Recordset
rs.Open sql, cn, adOpenKeyset, adLockOptimistic

If rs.Fields("status") = True Then
MsgBox "Asset is already disposed or does not exists", vbInformation
Exit Sub
End If

    'the first one is credit
    
    sql = ""
    Dim desc As String
    desc = txtASSETSNAME(1)
Dim scurr As Currency
scurr = (txtamount1 / txtRate)
sql = "set dateformat dmy insert into customerbalance (customerno,accname,amount,availablebalance,accno,transdescription,transdate,commission,period,transType,posted,locked,status,vno,auditid,auditdate,moduleid,accd,s_code,curr,yyear,rate,dr_scurr,cr_scurr,sbal) "
sql = sql & " values ('" & txtassetsacc & "','" & lblname1 & "'," & txtamount1 & "," & BAL1 + txtamount1 & ",'" & txtassetsacc & "','" & txtASSETSNAME(1) & "','" & Format(txttransdate, "dd/mm/yyyy") & "',0,'" & Month(txttransdate) & "','CR',0,0,0,'" & txtvno & "','" & User & "','" & Get_Server_Date & "','3','" & txtdisposaldr & "','" & cbosourcecode & "','" & cbocurrency & "'," & cboyearperiod & "," & txtRate & "," & scurr & "," & scurr & "," & scurr & " )"
cn.Execute sql

sql = ""
sql = "set dateformat dmy update cub set amount=" & txtamount1 & ",transdescription='" & desc & "',availablebalance=" & BAL1 + txtamount1 & ",transdate='" & Format(txttransdate, "dd/mm/yyyy") & "',vno='" & txtASSETSNO(1) & "',period='" & Month(txttransdate) & "',auditid='" & User & "',auditdate='" & Get_Server_Date & "',moduleid=2 where accno='" & txtassetsacc & "'"
cn.Execute sql
'// do the credit before proceeding
sql = ""

sql = "set dateformat dmy insert into customerbalance (customerno,accname,amount,availablebalance,accno,transdescription,transdate,commission,period,transType,posted,locked,status,vno,auditid,auditdate,moduleid,accd,s_code,curr,yyear,rate,dr_scurr,cr_scurr,sbal) "
sql = sql & " values ('" & txtdisposaldr & "','" & Lblname2 & "'," & txtamount2 & "," & bal2 + txtamount2 & ",'" & txtdisposaldr & "','" & txtASSETSNAME(1) & "','" & Format(txttransdate, "dd/mm/yyyy") & "',0,'" & Month(txttransdate) & "','DR',0,0,0,'" & txtvno & "','" & User & "','" & Get_Server_Date & "','3','" & txtassetsacc & "','" & cbosourcecode & "','" & cbocurrency & "'," & cboyearperiod & "," & txtRate & "," & scurr & "," & scurr & "," & scurr & " )"
cn.Execute sql

sql = ""
sql = "set dateformat dmy update cub set amount=" & txtamount2 & ",transdescription='" & desc & "',availablebalance=" & bal2 + txtamount2 & ",transdate='" & Format(txttransdate, "dd/mm/yyyy") & "',vno='" & txtASSETSNO(1) & "',period='" & Month(txttransdate) & "',auditid='" & User & "',auditdate='" & Get_Server_Date & "',moduleid=2 where accno='" & txtdisposaldr & "'"
cn.Execute sql


    
    '//provision and disposal of assets plus profit or loss whichever is there.
    
 sql = "set dateformat dmy insert into customerbalance (customerno,accname,amount,availablebalance,accno,transdescription,transdate,commission,period,transType,posted,locked,status,vno,auditid,auditdate,moduleid,accd,s_code,curr,yyear,rate,dr_scurr,cr_scurr,sbal) "
sql = sql & " values ('" & txtprovisionacc & "','" & Lblname3 & "'," & txtamount3 & "," & bal3 + txtamount3 & ",'" & txtprovisionacc & "','" & txtASSETSNAME(1) & "','" & Format(txttransdate, "dd/mm/yyyy") & "',0,'" & Month(txttransdate) & "','DR',0,0,0,'" & txtvno & "','" & User & "','" & Get_Server_Date & "','3','" & txtdisposalcr & "','" & cbosourcecode & "','" & cbocurrency & "'," & cboyearperiod & "," & txtRate & "," & scurr & "," & scurr & "," & scurr & " )"
cn.Execute sql

sql = ""
sql = "set dateformat dmy update cub set amount=" & txtamount3 & ",transdescription='" & desc & "',availablebalance=" & bal3 + txtamount3 & ",transdate='" & Format(txttransdate, "dd/mm/yyyy") & "',vno='" & txtASSETSNO(1) & "',period='" & Month(txttransdate) & "',auditid='" & User & "',auditdate='" & Get_Server_Date & "',moduleid=2 where accno='" & txtprovisionacc & "'"
cn.Execute sql
'// do the credit before proceeding
sql = ""

sql = "set dateformat dmy insert into customerbalance (customerno,accname,amount,availablebalance,accno,transdescription,transdate,commission,period,transType,posted,locked,status,vno,auditid,auditdate,moduleid,accd,s_code,curr,yyear,rate,dr_scurr,cr_scurr,sbal) "
sql = sql & " values ('" & txtdisposalcr & "','" & Lblname4 & "'," & txtamount4 & "," & bal4 + txtamount4 & ",'" & txtdisposalcr & "','" & txtASSETSNAME(1) & "','" & Format(txttransdate, "dd/mm/yyyy") & "',0,'" & Month(txttransdate) & "','CR',0,0,0,'" & txtvno & "','" & User & "','" & Get_Server_Date & "','3','" & txtprovisionacc & "','" & cbosourcecode & "','" & cbocurrency & "'," & cboyearperiod & "," & txtRate & "," & scurr & "," & scurr & "," & scurr & " )"
cn.Execute sql

sql = ""
sql = "set dateformat dmy update cub set amount=" & txtamount4 & ",transdescription='" & desc & "',availablebalance=" & bal4 + txtamount4 & ",transdate='" & Format(txttransdate, "dd/mm/yyyy") & "',vno='" & txtASSETSNO(1) & "',period='" & Month(txttransdate) & "',auditid='" & User & "',auditdate='" & Get_Server_Date & "',moduleid=2 where accno='" & txtdisposalcr & "'"
cn.Execute sql
   '// do the bank and also the disposal account
   sql = "set dateformat dmy insert into customerbalance (customerno,accname,amount,availablebalance,accno,transdescription,transdate,commission,period,transType,posted,locked,status,vno,auditid,auditdate,moduleid,accd,s_code,curr,yyear,rate,dr_scurr,cr_scurr,sbal) "
sql = sql & " values ('" & txtbankdr & "','" & lbllabel7 & "'," & txtamount7 & "," & bal7 + txtamount7 & ",'" & txtbankdr & "','" & txtASSETSNAME(1) & "','" & Format(txttransdate, "dd/mm/yyyy") & "',0,'" & Month(txttransdate) & "','DR',0,0,0,'" & txtvno & "','" & User & "','" & Get_Server_Date & "','3','" & txtbankdr & "','" & cbosourcecode & "','" & cbocurrency & "'," & cboyearperiod & "," & txtRate & "," & scurr & "," & scurr & "," & scurr & " )"
cn.Execute sql

sql = ""
sql = "set dateformat dmy update cub set amount=" & txtamount7 & ",transdescription='" & desc & "',availablebalance=" & bal7 + txtamount7 & ",transdate='" & Format(txttransdate, "dd/mm/yyyy") & "',vno='" & txtvno & "',period='" & Month(txttransdate) & "',auditid='" & User & "',auditdate='" & Get_Server_Date & "',moduleid=2 where accno='" & txtbankdr & "'"
cn.Execute sql
'// do the credit before proceeding
sql = ""

sql = "set dateformat dmy insert into customerbalance (customerno,accname,amount,availablebalance,accno,transdescription,transdate,commission,period,transType,posted,locked,status,vno,auditid,auditdate,moduleid,accd,s_code,curr,yyear,rate,dr_scurr,cr_scurr,sbal) "
sql = sql & " values ('" & txtdisposalbankcr & "','" & lbllabel8 & "'," & txtamount8 & "," & bal8 + txtamount8 & ",'" & txtdisposalbankcr & "','" & txtASSETSNAME(1) & "','" & Format(txttransdate, "dd/mm/yyyy") & "',0,'" & Month(txttransdate) & "','CR',0,0,0,'" & txtvno & "','" & User & "','" & Get_Server_Date & "','3','" & txtbankdr & "','" & cbosourcecode & "','" & cbocurrency & "'," & cboyearperiod & "," & txtRate & "," & scurr & "," & scurr & "," & scurr & " )"
cn.Execute sql

sql = ""
sql = "set dateformat dmy update cub set amount=" & txtamount8 & ",transdescription='" & desc & "',availablebalance=" & bal8 + txtamount8 & ",transdate='" & Format(txttransdate, "dd/mm/yyyy") & "',vno='" & txtvno & "',period='" & Month(txttransdate) & "',auditid='" & User & "',auditdate='" & Get_Server_Date & "',moduleid=2 where accno='" & txtdisposalbankcr & "'"
cn.Execute sql
    
    '//GO TO THE BALANCING FIGURE which it can be a profit or a loss.
    
    If txtamount5 <> "" Then ' for profit
    sql = ""

sql = "set dateformat dmy insert into customerbalance (customerno,accname,amount,availablebalance,accno,transdescription,transdate,commission,period,transType,posted,locked,status,vno,auditid,auditdate,moduleid,accd,s_code,curr,yyear,rate,dr_scurr,cr_scurr,sbal) "
sql = sql & " values ('" & txtprofit & "','" & Lblname5 & "'," & txtamount5 & "," & bal5 + txtamount5 & ",'" & txtprofit & "','" & txtASSETSNAME(1) & "','" & Format(txttransdate, "dd/mm/yyyy") & "',0,'" & Month(txttransdate) & "','CR',0,0,0,'" & txtSERIALNO(1) & "','" & User & "','" & Get_Server_Date & "','3','" & txtdisposalcr & "','" & cbosourcecode & "','" & cbocurrency & "'," & cboyearperiod & "," & txtRate & "," & scurr & "," & scurr & "," & scurr & " )"
cn.Execute sql

sql = ""
sql = "set dateformat dmy update cub set amount=" & txtamount5 & ",transdescription='" & desc & "',availablebalance=" & bal5 + txtamount5 & ",transdate='" & Format(txttransdate, "dd/mm/yyyy") & "',vno='" & txtASSETSNO(1) & "',period='" & Month(txttransdate) & "',auditid='" & User & "',auditdate='" & Get_Server_Date & "',moduleid=2 where accno='" & txtprofit & "'"
cn.Execute sql

sql = "set dateformat dmy insert into customerbalance (customerno,accname,amount,availablebalance,accno,transdescription,transdate,commission,period,transType,posted,locked,status,vno,auditid,auditdate,moduleid,accd,s_code,curr,yyear,rate,dr_scurr,cr_scurr,sbal) "
sql = sql & " values ('" & txtproftdr & "','" & lbllabel9 & "'," & txtamount9 & "," & BAL9 + txtamount9 & ",'" & txtproftdr & "','" & txtASSETSNAME(1) & "','" & Format(txttransdate, "dd/mm/yyyy") & "',0,'" & Month(txttransdate) & "','DR',0,0,0,'" & txtvno & "','" & User & "','" & Get_Server_Date & "','3','" & txtprofit & "','" & cbosourcecode & "','" & cbocurrency & "'," & cboyearperiod & "," & txtRate & "," & scurr & "," & scurr & "," & scurr & " )"
cn.Execute sql

sql = ""
sql = "set dateformat dmy update cub set amount=" & txtamount9 & ",transdescription='" & desc & "',availablebalance=" & BAL9 + txtamount9 & ",transdate='" & Format(txttransdate, "dd/mm/yyyy") & "',vno='" & txtvno & "',period='" & Month(txttransdate) & "',auditid='" & User & "',auditdate='" & Get_Server_Date & "',moduleid=2 where accno='" & txtproftdr & "'"
cn.Execute sql
    End If
    
      If txtamount6 <> "" Then 'for the loss
      sql = ""

sql = "set dateformat dmy insert into customerbalance (customerno,accname,amount,availablebalance,accno,transdescription,transdate,commission,period,transType,posted,locked,status,vno,auditid,auditdate,moduleid,accd,s_code,curr,yyear,rate,dr_scurr,cr_scurr,sbal) "
sql = sql & " values ('" & txtloss & "','" & Lblname6 & "'," & txtamount6 & "," & bal6 + txtamount6 & ",'" & txtloss & "','" & txtASSETSNAME(1) & "','" & Format(txttransdate, "dd/mm/yyyy") & "',0,'" & Month(txttransdate) & "','DR',0,0,0,'" & txtvno & "','" & User & "','" & Get_Server_Date & "','3','" & txtdisposalcr & "','" & cbosourcecode & "','" & cbocurrency & "'," & cboyearperiod & "," & txtRate & "," & scurr & "," & scurr & "," & scurr & " )"
cn.Execute sql

sql = ""
sql = "set dateformat dmy update cub set amount=" & txtamount6 & ",transdescription='" & desc & "',availablebalance=" & bal6 + txtamount6 & ",transdate='" & Format(txttransdate, "dd/mm/yyyy") & "',vno='" & txtASSETSNO(1) & "',period='" & Month(txttransdate) & "',auditid='" & User & "',auditdate='" & Get_Server_Date & "',moduleid=2 where accno='" & txtloss & "'"
cn.Execute sql
Dim CU As Currency
CU = bal10 + txtamount10
sql = "set dateformat dmy insert into customerbalance (customerno,accname,amount,availablebalance,accno,transdescription,transdate,commission,period,transType,posted,locked,status,vno,auditid,auditdate,moduleid,accd,s_code,curr,yyear,rate,dr_scurr,cr_scurr,sbal) "
sql = sql & " values ('" & txtlosscr & "','" & lbllabel10 & "'," & txtamount10 & "," & bal10 + (txtamount10.Text) & ",'" & txtlosscr & "','" & txtASSETSNAME(1) & "','" & Format(txttransdate, "dd/mm/yyyy") & "',0,'" & Month(txttransdate) & "','CR',0,0,0,'" & txtvno & "','" & User & "','" & Get_Server_Date & "','3','" & txtloss & "','" & cbosourcecode & "','" & cbocurrency & "'," & cboyearperiod & "," & txtRate & "," & scurr & "," & scurr & "," & scurr & " )"
cn.Execute sql

sql = ""
sql = "set dateformat dmy update cub set amount=" & txtamount10 & ",transdescription='" & desc & "',availablebalance=" & bal10 + txtamount10 & ",transdate='" & Format(txttransdate, "dd/mm/yyyy") & "',vno='" & txtvno & "',period='" & Month(txttransdate) & "',auditid='" & User & "',auditdate='" & Get_Server_Date & "',moduleid=2 where accno='" & txtlosscr & "'"
cn.Execute sql
'// update the bank account and disposal account
  
    
'//update the asset register by putting status=1

sql = ""
sql = "set dateformat dmy Update assets set status=1,disposaldate='" & txttransdate & "',d_amount=" & txtamount7 & " where assetsno='" & txtASSETSNO(1) & "'"
cn.Execute sql

    End If
    MsgBox "Transactions Posted successfully", vbInformation, "AC- Assets Disposals"
    
    Exit Sub
ErrorHandler:
    MsgBox err.description
End Sub

Private Sub cmdPrint_Click()

End Sub

Private Sub Command1_Click()
Dim Z, S, U
    Dim rs As Recordset
      frmsearchacc.Show vbModal
    
    Z = sel
    If Z <> "" Then
        txtprovisionacc = Z
   
        End If
          
    Set myclass = New cdbase
    Set cn = CreateObject("adodb.connection")
    Provider = myclass.OpenCon
   cn.Open Provider, "atm", "atm"
sql = "select * from cuB where ACCno='" & Z & "'"
Set rs = New ADODB.Recordset
rs.Open sql, cn

If Not rs.EOF Then
If Not IsNull(rs.Fields("name")) Then Lblname3 = rs.Fields("name")
If Not IsNull(rs.Fields("availablebalance")) Then bal3 = rs.Fields("availablebalance")
End If
End Sub

Private Sub Command10_Click()
Dim Z, S, U
    Dim rs As Recordset
      frmsearchacc.Show vbModal
       ' txtredirect = Sel
   ' frmsearchaccounts.Show vbModal
    Z = sel
    If Z <> "" Then
        txtbankdr = Z
   
        End If
          
    Set myclass = New cdbase
    Set cn = CreateObject("adodb.connection")
    Provider = myclass.OpenCon
   cn.Open Provider, "atm", "atm"
sql = "select * from cuB where ACCno='" & Z & "'"
Set rs = New ADODB.Recordset
rs.Open sql, cn

If Not rs.EOF Then
If Not IsNull(rs.Fields("name")) Then lbllabel7 = rs.Fields("name")
If Not IsNull(rs.Fields("availablebalance")) Then bal7 = rs.Fields("availablebalance")
End If
'bookba1 = cub_balance(txtassetsacc)
End Sub

Private Sub Command11_Click()
Dim Z, S, U
    Dim rs As Recordset
      frmsearchacc.Show vbModal
        Z = sel
    If Z <> "" Then
        txtproftdr = Z
   
        End If
          
    Set myclass = New cdbase
    Set cn = CreateObject("adodb.connection")
    Provider = myclass.OpenCon
   cn.Open Provider, "atm", "atm"
sql = "select * from cuB where ACCno='" & Z & "'"
Set rs = New ADODB.Recordset
rs.Open sql, cn

If Not rs.EOF Then
If Not IsNull(rs.Fields("name")) Then lbllabel9 = rs.Fields("name")
If Not IsNull(rs.Fields("availablebalance")) Then BAL9 = rs.Fields("availablebalance")
End If
End Sub

Private Sub Command12_Click()
Dim Z, S, U
    Dim rs As Recordset
      frmsearchacc.Show vbModal
        Z = sel
    If Z <> "" Then
        txtlosscr = Z
   
        End If
          
    Set myclass = New cdbase
    Set cn = CreateObject("adodb.connection")
    Provider = myclass.OpenCon
   cn.Open Provider, "atm", "atm"
sql = "select * from cuB where ACCno='" & Z & "'"
Set rs = New ADODB.Recordset
rs.Open sql, cn

If Not rs.EOF Then
If Not IsNull(rs.Fields("name")) Then lbllabel10 = rs.Fields("name")
If Not IsNull(rs.Fields("availablebalance")) Then bal10 = rs.Fields("availablebalance")
End If
End Sub

Private Sub Command2_Click()
Dim Z, S, U
    Dim rs As Recordset
      frmsearchacc.Show vbModal
      Z = sel
    If Z <> "" Then
        txtdisposalcr = Z
   
        End If
          
    Set myclass = New cdbase
    Set cn = CreateObject("adodb.connection")
    Provider = myclass.OpenCon
   cn.Open Provider, "atm", "atm"
sql = "select * from cuB where ACCno='" & Z & "'"
Set rs = New ADODB.Recordset
rs.Open sql, cn

If Not rs.EOF Then
If Not IsNull(rs.Fields("name")) Then Lblname4 = rs.Fields("name")
If Not IsNull(rs.Fields("availablebalance")) Then bal4 = rs.Fields("availablebalance")
End If
End Sub

Private Sub Command3_Click()
Dim Z, S, U
    Dim rs As Recordset
      frmsearchacc.Show vbModal
        Z = sel
    If Z <> "" Then
        txtprofit = Z
   
        End If
          
    Set myclass = New cdbase
    Set cn = CreateObject("adodb.connection")
    Provider = myclass.OpenCon
   cn.Open Provider, "atm", "atm"
sql = "select * from cuB where ACCno='" & Z & "'"
Set rs = New ADODB.Recordset
rs.Open sql, cn

If Not rs.EOF Then
If Not IsNull(rs.Fields("name")) Then Lblname5 = rs.Fields("name")
If Not IsNull(rs.Fields("availablebalance")) Then bal5 = rs.Fields("availablebalance")
End If
End Sub

Private Sub Command4_Click()
Dim Z, S, U
    Dim rs As Recordset
      frmsearchacc.Show vbModal
       ' txtredirect = Sel
   ' frmsearchaccounts.Show vbModal
    Z = sel
    If Z <> "" Then
        txtloss = Z
   
        End If
          
    Set myclass = New cdbase
    Set cn = CreateObject("adodb.connection")
    Provider = myclass.OpenCon
   cn.Open Provider, "atm", "atm"
sql = "select * from cuB where ACCno='" & Z & "'"
Set rs = New ADODB.Recordset
rs.Open sql, cn

If Not rs.EOF Then
If Not IsNull(rs.Fields("name")) Then Lblname6 = rs.Fields("name")
If Not IsNull(rs.Fields("availablebalance")) Then bal6 = rs.Fields("availablebalance")
End If
End Sub

Private Sub Command5_Click()
On Error Resume Next
txtNETREALIABLEVALUE = ""
frmsearchassets.Show vbModal
Dim Y As String
Y = sel
'm = False
If Y <> "" Then
     Dim cn As Connection
    Set cn = New ADODB.Connection
    
    cn.Open frmODBCLogon.cboDSNList, "atm"
sql = ""
'If reportpath = "" Then reportpath = GetSetting("payroll", "AppName", "rptPath", rptPath)
sql = "SELECT     AssetsNo, AssetserialNo, AssetsName from assets where assetsno='" & Y & "' order by AssetsNo"
Set rs = New ADODB.Recordset
rs.Open sql, cn
If Not rs.EOF Then

If Not IsNull(rs.Fields(0)) Then txtASSETSNO(1) = (rs.Fields(0))
If Not IsNull(rs.Fields(1)) Then txtSERIALNO(1) = (rs.Fields(1))
If Not IsNull(rs.Fields(2)) Then txtASSETSNAME(1) = (rs.Fields(2))
'If Not IsNull(rs.Fields(2)) Then txtNETREALIABLEVALUE = (rs.Fields(3))

'get the serial no of the item
Set rs = New ADODB.Recordset
sql = "Select TOP 1  * from assetstrans where assetcode='" & txtASSETSNO(1) & "' ORDER BY ASSETTRANSID DESC"
rs.Open sql, cn
If rs.EOF Then
'MsgBox "No data in the such data available"
Exit Sub
Else
    'If .RecordCount > 0 Then
    txtNETREALIABLEVALUE = rs!nrv
    'txtserialno(1) = rs!assetserialno
    End If
'Call cboname_p
txtamount1 = txtNETREALIABLEVALUE
txtamount2 = txtNETREALIABLEVALUE
txtamount3 = txtNETREALIABLEVALUE
txtamount4 = txtNETREALIABLEVALUE
txtamount7 = txtNETREALIABLEVALUE
txtamount8 = txtNETREALIABLEVALUE
txtamount9 = txtNETREALIABLEVALUE
txtamount10 = txtNETREALIABLEVALUE
Set rs = Nothing
End If
End If
End Sub

Private Sub Command6_Click()
On Error Resume Next
frmsearchcurr.Show vbModal
Dim Y As String
Y = sel
'm = False
If Y <> "" Then
     Dim cn As Connection
    Set cn = New ADODB.Connection
    
    cn.Open frmODBCLogon.cboDSNList
sql = ""

sql = "SELECT     rateaganistsource,currcode From Curr where currcode='" & Y & "'"
Set rs = New ADODB.Recordset
rs.Open sql, cn
If Not rs.EOF Then
If Not IsNull(rs.Fields(0)) Then txtRate = (rs.Fields(0))
If Not IsNull(rs.Fields(1)) Then cbocurrency = Trim((rs.Fields(1)))
End If


End If
End Sub

Private Sub Command7_Click()
Dim Z, S, U
    Dim rs As Recordset
      frmsearchacc.Show vbModal
       ' txtredirect = Sel
   ' frmsearchaccounts.Show vbModal
    Z = sel
    If Z <> "" Then
        txtassetsacc = Z
   
        End If
          
    Set myclass = New cdbase
    Set cn = CreateObject("adodb.connection")
    Provider = myclass.OpenCon
   cn.Open Provider, "atm", "atm"
sql = "select * from cuB where ACCno='" & Z & "'"
Set rs = New ADODB.Recordset
rs.Open sql, cn

If Not rs.EOF Then
If Not IsNull(rs.Fields("name")) Then lblname1 = rs.Fields("name")
If Not IsNull(rs.Fields("availablebalance")) Then BAL1 = rs.Fields("availablebalance")
End If
'bookba1 = cub_balance(txtassetsacc)
End Sub

Private Sub Command8_Click()
Dim Z, S, U
    Dim rs As Recordset
      frmsearchacc.Show vbModal
       ' txtredirect = Sel
   ' frmsearchaccounts.Show vbModal
    Z = sel
    If Z <> "" Then
        txtdisposaldr = Z
   
        End If
          
    Set myclass = New cdbase
    Set cn = CreateObject("adodb.connection")
    Provider = myclass.OpenCon
   cn.Open Provider, "atm", "atm"
sql = "select * from cuB where ACCno='" & Z & "'"
Set rs = New ADODB.Recordset
rs.Open sql, cn

If Not rs.EOF Then
If Not IsNull(rs.Fields("name")) Then Lblname2 = rs.Fields("name")
If Not IsNull(rs.Fields("availablebalance")) Then bal2 = rs.Fields("availablebalance")
End If
End Sub

Private Sub Command9_Click()
Dim Z, S, U
    Dim rs As Recordset
      frmsearchacc.Show vbModal
       ' txtredirect = Sel
   ' frmsearchaccounts.Show vbModal
    Z = sel
    If Z <> "" Then
        txtdisposalbankcr = Z
   
        End If
          
    Set myclass = New cdbase
    Set cn = CreateObject("adodb.connection")
    Provider = myclass.OpenCon
   cn.Open Provider, "atm", "atm"
sql = "select * from cuB where ACCno='" & Z & "'"
Set rs = New ADODB.Recordset
rs.Open sql, cn

If Not rs.EOF Then
If Not IsNull(rs.Fields("name")) Then lbllabel8 = rs.Fields("name")
If Not IsNull(rs.Fields("availablebalance")) Then bal8 = rs.Fields("availablebalance")
End If
'bookba1 = cub_balance(txtassetsacc)
End Sub

Private Sub Form_Load()
lab = DatePart("YYYY", Date)
'DTPtransdate = Format(Get_Server_Date, "dd/mm/yyyy")
txttransdate = Format(Get_Server_Date, "dd/mm/yyyy")
cbosourcecode = "GL-DP"
cboyearperiod.Text = Format(Get_Server_Date, "mm") & "-" & Format(Get_Server_Date, "yyyy")
cbocurrency = "KES"
'//get the asset accounts
'Set cn = CreateObject("adodb.connection")
'      Set myclass = New cdbase
'    Provider = myclass.OpenCon
'cn.Open Provider
'   sql = ""
'   sql = ""
'     sql = "select p_dep,a_dep from param"
'     Set rs = New ADODB.Recordset
'     rs.Open sql, cn, adOpenKeyset, adLockOptimistic
'
'     If Not rs.EOF Then
'     txtassetsacc = rs.Fields(0)
'     txtdisposaldr = rs.Fields(1)
'     'txtredirect = rs.Fields(1)
'     get_namecr txtassetsacc
'     'get_namedr
'     Lblname2 = lab
'     get_namecr txtdisposaldr
'     lblname1 = lab
'     End If
End Sub

