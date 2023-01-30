VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Begin VB.Form Frmapproval 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Approvals"
   ClientHeight    =   4710
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   ScaleHeight     =   4710
   ScaleWidth      =   6465
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Revoke multi intake"
      Height          =   555
      Left            =   5640
      TabIndex        =   6
      Top             =   4080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Allow multi intake"
      Height          =   495
      Left            =   4560
      TabIndex        =   5
      Top             =   4080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtsno 
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Top             =   4200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton Cmdapprove 
      Caption         =   "Approve"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3840
      Width           =   1215
   End
   Begin MSComctlLib.ListView Lvwapprol 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   5530
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
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Sno"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Regdate"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "IdNo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Names"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "SNO"
      Height          =   255
      Left            =   1800
      TabIndex        =   4
      Top             =   3840
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "Frmapproval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmdapprove_Click()
'check the user
sql = "SELECT     UserLoginIDs,levels, UserGroup, SUPERUSER From UserAccounts where UserLoginIDs='" & User & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
If rs!Levels <> "Manager" And rs!Levels <> "Accounts" Then
MsgBox "You are not allowed to Approve", vbInformation
Exit Sub
End If
End If
For I = Lvwapprol.ListItems.Count To 1 Step -1
Set li = Lvwapprol.ListItems(I)
If li.Checked = True Then
       
        DoEvents
'If Lvwapprol.SelectedItem.Checked = True Then
oSaccoMaster.ExecuteThis ("UPDATE d_Suppliers SET Approval = 1 where sno=" & Lvwapprol.ListItems(I) & "")
    DoEvents
        End If
        
    Next I

MsgBox "Record Updated successfully."
cmdclose_Click
'End If
End Sub

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub Command1_Click()
sql = ""
sql = "UPDATE  d_Suppliers  SET MASS='1'where sno=" & txtSNo & ""
 oSaccoMaster.GetRecordset (sql)
 MsgBox "Supplier updated Successfully"
 txtSNo = ""
End Sub

Private Sub Command2_Click()
sql = ""
sql = "UPDATE  d_Suppliers  SET MASS='0'where sno=" & txtSNo & ""
 oSaccoMaster.GetRecordset (sql)
 MsgBox "Supplier Removed Successfully"
 txtSNo = ""
End Sub

Private Sub Command3_Click()

End Sub

Private Sub Form_Load()
Provider = cn
Set cn = New ADODB.Connection
cn.Open Provider, "atm", "atm"
sql = ""
'If reportpath = "" Then reportpath = GetSetting("payroll", "AppName", "rptPath", rptPath)
'sql = "SELECT * FROm d_Requisition where status='Approve' order by rno"

sql = "SELECT     SNo, Regdate, IdNo, Names,  active, Approval From d_Suppliers where Approval=0"


Set rs = New ADODB.Recordset
rs.Open sql, cn

'// load it into the sq
    While Not rs.EOF
        Set li = Lvwapprol.ListItems.Add(, , rs!sno)
    'li.SubItems(1) = (rs.Fields("sno"))
    li.SubItems(1) = (rs.Fields("Regdate"))
    li.SubItems(2) = rs.Fields("IdNo")
    li.SubItems(3) = rs.Fields("Names")
    'li.SubItems(4) = rs.Fields("active")
'    li.SubItems(5) = rs.Fields("qnty")
'    li.SubItems(6) = rs.Fields("description")
'    li.SubItems(7) = rs.Fields("Status")
    rs.MoveNext
    Wend
End Sub
