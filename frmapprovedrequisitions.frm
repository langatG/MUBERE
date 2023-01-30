VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Begin VB.Form frmapprovedrequisitions 
   Caption         =   "Approved Requisitions - 2 ND LEVEL APPROVALS"
   ClientHeight    =   6750
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10440
   LinkTopic       =   "Form1"
   ScaleHeight     =   6750
   ScaleWidth      =   10440
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdshowall 
      Caption         =   "Show All"
      Height          =   375
      Left            =   3840
      TabIndex        =   6
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtrno 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton cmdprocess 
      Caption         =   "Process"
      Height          =   375
      Left            =   10560
      TabIndex        =   4
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdrefresh 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   9360
      TabIndex        =   3
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton cmdedit 
      Caption         =   "Edit"
      Height          =   375
      Left            =   8040
      TabIndex        =   2
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      Height          =   375
      Left            =   8400
      TabIndex        =   1
      Top             =   6240
      Width           =   1095
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      DrawStyle       =   2  'Dot
      DrawWidth       =   17015
      Height          =   360
      Left            =   3120
      Picture         =   "frmapprovedrequisitions.frx":0000
      ScaleHeight     =   360
      ScaleWidth      =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   240
   End
   Begin MSComctlLib.ListView lvwrequisition 
      Height          =   3015
      Left            =   0
      TabIndex        =   7
      Top             =   2520
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   5318
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
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ItemNo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Transdate"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Cost Centre"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Description"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Make"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Quantity"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Comments"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Status"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label3 
      Caption         =   "Requsition No."
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "New Requisitions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label5 
      Caption         =   "List of Requisitions Pending Approval "
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   2775
   End
End
Attribute VB_Name = "frmapprovedrequisitions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command3_Click()

End Sub

Private Sub Command6_Click()
Unload Me
End Sub

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub Form_Load()

Provider = cn
Set cn = New ADODB.Connection
cn.Open Provider, "atm", "atm"
sql = ""
'If reportpath = "" Then reportpath = GetSetting("payroll", "AppName", "rptPath", rptPath)
sql = "SELECT * FROm d_Requisition where status='Approve' order by rno"
Set rs = New ADODB.Recordset
rs.Open sql, cn

'// load it into the sq
    While Not rs.EOF
        Set li = lvwrequisition.ListItems.Add(, , rs!Rno)
    li.SubItems(1) = (rs.Fields("transdate"))
    li.SubItems(2) = rs.Fields("costcentre")
    li.SubItems(3) = rs.Fields("iname")
    li.SubItems(4) = rs.Fields("make")
    li.SubItems(5) = rs.Fields("qnty")
    li.SubItems(6) = rs.Fields("description")
    li.SubItems(7) = rs.Fields("Status")
    rs.MoveNext
    Wend
    
End Sub





Private Sub lvwrequisition_DblClick()

frmrequisitionapproval2.lblrno = lvwrequisition.SelectedItem
frmrequisitionapproval2.lblname = lvwrequisition.SelectedItem.ListSubItems(3)

frmrequisitionapproval2.Show vbModal
End Sub

Private Sub Picture2_Click()
On Error Resume Next
frmsearchrequisition.Show vbModal
Dim Y As String
Y = sel

If Y <> "" Then
Provider = cn
Set cn = New ADODB.Connection
cn.Open Provider, "atm", "atm"
sql = ""
sql = "SELECT    * FROM         d_Requisition where rno=" & Y & ""
Set rs = New ADODB.Recordset
rs.Open sql, cn
lvwrequisition.ListItems.Clear
If Not rs.EOF Then
If Not IsNull(rs.Fields(0)) Then txtrno = (rs.Fields(0))
End If
End If
'// load it into the sq
While Not rs.EOF
    Set li = lvwrequisition.ListItems.Add(, , txtrno)
    li.SubItems(1) = (rs.Fields("transdate"))
    li.SubItems(2) = rs.Fields("costcentre")
    li.SubItems(3) = rs.Fields("iname")
    li.SubItems(4) = rs.Fields("make")
    li.SubItems(5) = rs.Fields("quantity")
    li.SubItems(6) = rs.Fields("description")
    rs.MoveNext
    Wend
End Sub

