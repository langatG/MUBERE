VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Begin VB.Form frmrequisitionapproval2 
   Caption         =   "ACTION REQUISITION APPROVAL 2"
   ClientHeight    =   6570
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8085
   LinkTopic       =   "Form2"
   ScaleHeight     =   6570
   ScaleWidth      =   8085
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture3 
      Height          =   285
      Left            =   2520
      Picture         =   "frmrequisitionapproval2.frx":0000
      ScaleHeight     =   225
      ScaleWidth      =   195
      TabIndex        =   19
      Top             =   1080
      Width           =   255
   End
   Begin VB.TextBox txtAccNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      MouseIcon       =   "frmrequisitionapproval2.frx":02C2
      TabIndex        =   17
      Top             =   1080
      Width           =   1095
   End
   Begin VB.ComboBox cboaction 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmrequisitionapproval2.frx":0B8C
      Left            =   1440
      List            =   "frmrequisitionapproval2.frx":0B99
      TabIndex        =   5
      Top             =   2760
      Width           =   2055
   End
   Begin VB.TextBox txtestimate 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1440
      TabIndex        =   4
      Text            =   "0.00"
      Top             =   1560
      Width           =   2175
   End
   Begin VB.TextBox txtcomments 
      Appearance      =   0  'Flat
      Height          =   615
      Left            =   1080
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   3240
      Width           =   4695
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      Height          =   375
      Left            =   5280
      TabIndex        =   0
      Top             =   5880
      Width           =   1215
   End
   Begin MSComctlLib.ListView lvwapproval 
      Height          =   1575
      Left            =   120
      TabIndex        =   2
      Top             =   4200
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   2778
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
   Begin VB.Label lblAccNo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   2880
      TabIndex        =   18
      Top             =   1080
      Width           =   5115
   End
   Begin VB.Label lblBudget 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   1440
      TabIndex        =   16
      Top             =   2160
      Width           =   2235
   End
   Begin VB.Label Label8 
      Caption         =   "GL Account :"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "Budget :"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2280
      Width           =   735
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
      Left            =   1080
      TabIndex        =   13
      Top             =   0
      Width           =   5175
   End
   Begin VB.Label Label2 
      Caption         =   "General Option"
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "History Option"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Action"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Estimate Cost"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Comments"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label lblrno 
      Height          =   255
      Left            =   2040
      TabIndex        =   7
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label LBLNAME 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4440
      TabIndex        =   6
      Top             =   600
      Width           =   3495
   End
End
Attribute VB_Name = "frmrequisitionapproval2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdclose_Click()

Unload Me
End Sub

Private Sub cmdsave_Click()
' sql = ""
'             sql = "d_insert_d_Approve2 '" & lblrno & "','0','" & cboaction & "','" & User & "'"
'             oSaccoMaster.ExecuteThis (sql)

         
             sql = ""
             sql = "d_sp_UpdateApprove2 '" & cboaction & "','" & lblrno & "','" & Txtaccno & "'," & txtestimate & ""
              oSaccoMaster.ExecuteThis (sql)
              
                sql = ""
                sql = "UPDATE    d_Requisition  SET              status='" & cboaction & "'  where rno='" & lblrno & "'"
             
                oSaccoMaster.ExecuteThis (sql)
                
                If cboaction = "Approved" Then
                '//call requisition
                '//requisition
                 
               reportname = "requisition.rpt"
                STRFORMULA = "{d_Requisition.Status} = 'Approved' and {d_Requisition.RNo}='" & Trim(lblrno) & "'"
                
                Show_Sales_Crystal_Report STRFORMULA, reportname, ""
                cmdclose_Click
                
                End If
                
End Sub

Private Sub Form_Load()
Dim q As Double
Dim Rno As String

'//get the quantity for the same first
frmrequisitionapproval2.lblrno = Trim(frmapprovedrequisitions.lvwrequisition.SelectedItem)
frmrequisitionapproval2.lblname = Trim(frmapprovedrequisitions.lvwrequisition.SelectedItem.ListSubItems(3))
Rno = frmrequisitionapproval2.lblrno
sql = ""
sql = "SELECT qnty,pricing   FROM d_Requisition  where rno='" & Rno & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
q = rs.Fields(0)
txtestimate.Text = q * CCur(rs.Fields(1))

End If
'//put in here the history

sql = ""
sql = "SELECT     RNo, Approved, description, auditid  FROM         d_Approve2 where rno='" & lblrno & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
 While Not rs.EOF
   Set li = lvwapproval.ListItems.Add(, , rs.Fields(0))
  li.ListSubItems.Add , , rs.Fields(1)
  li.ListSubItems.Add , , rs.Fields(2)
  li.ListSubItems.Add , , rs.Fields(3)
  
  
 rs.MoveNext
 Wend
End Sub

Private Sub get_namecr()
    Dim myclass As cdbase
    Set cn = CreateObject("adodb.connection")
    Set myclass = New cdbase
    Provider = myclass.OpenCon
   cn.Open Provider, "atm", "atm"
    sql = ""
    sql = "select * from cub where accno='" & sel & "'"
    Set rs = New ADODB.Recordset
    rs.Open sql, cn, adOpenKeyset, adLockOptimistic
    If rs.EOF Then
    Else
    If Not IsNull(rs.Fields("name")) Then lblAccNo = rs.Fields("name")
    End If
End Sub

Private Sub Picture3_Click()
Me.MousePointer = vbHourglass
    frmsearchacc.Show vbModal
    Txtaccno = sel
    get_namecr
    Me.MousePointer = 0
    txtAccNo_Validate True
    
End Sub

Private Sub txtAccNo_Validate(Cancel As Boolean)
Set rs = oSaccoMaster.GetRecordset("d_sp_getBudget '" & Txtaccno & "'")
If Not rs.EOF Then
lblBudget = Format(rs.Fields(0), "#,##0.00")
Else
lblBudget = "0.00"
End If


End Sub
