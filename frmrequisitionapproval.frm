VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmrequisitionapproval 
   Caption         =   "ACTION - APPROVAL"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   5865
   ScaleWidth      =   8145
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      Height          =   375
      Left            =   5400
      TabIndex        =   11
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
      Height          =   375
      Left            =   2280
      TabIndex        =   10
      Top             =   5400
      Width           =   1095
   End
   Begin MSComctlLib.ListView lvwapprovals 
      Height          =   1575
      Left            =   120
      TabIndex        =   9
      Top             =   3600
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   2778
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "User"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Level"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox txtcomments 
      Appearance      =   0  'Flat
      Height          =   615
      Left            =   2040
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   2280
      Width           =   4695
   End
   Begin VB.TextBox txtestimate 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Text            =   "0"
      Top             =   1680
      Width           =   2055
   End
   Begin VB.ComboBox cboaction 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmrequisitionapproval.frx":0000
      Left            =   2040
      List            =   "frmrequisitionapproval.frx":000D
      TabIndex        =   4
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label LBLNAME 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4560
      TabIndex        =   13
      Top             =   720
      Width           =   3495
   End
   Begin VB.Label lblrno 
      Height          =   255
      Left            =   2160
      TabIndex        =   12
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label6 
      Caption         =   "Comments"
      Height          =   255
      Left            =   600
      TabIndex        =   7
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Estimate Cost"
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Action"
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "History Option"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "General Option"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "           PROCESS APPROVAL LEVELS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   5175
   End
End
Attribute VB_Name = "frmrequisitionapproval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdsave_Click()
 sql = ""
             sql = "d_insert_d_Approve2 '" & lblrno & "','0','Level 2','" & User & "'"
             oSaccoMaster.ExecuteThis (sql)
             
             sql = ""
             sql = "UPDATE    d_Approve1  SET approved=1 , description='" & cboaction & "' where rno='" & lblrno & "'"
             
              oSaccoMaster.ExecuteThis (sql)
              
              sql = ""
             sql = "UPDATE    d_Requisition  SET status='" & cboaction & "'  where rno='" & lblrno & "'"
             
              oSaccoMaster.ExecuteThis (sql)
              MsgBox "Records Saved Successfully !!"
              cmdClose_Click
End Sub

Private Sub Form_Load()

'//put in here the history

sql = ""
sql = "SELECT     RNo, Approved, description, auditid  FROM d_Approve1 where rno='" & lblrno & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
 While Not rs.EOF
  Set li = lvwapprovals.ListItems.Add(, , rs.Fields(0))
  li.ListSubItems.Add , , rs.Fields(1)
  li.ListSubItems.Add , , rs.Fields(2)
  li.ListSubItems.Add , , rs.Fields(3)
  

 rs.MoveNext
 Wend
End Sub

