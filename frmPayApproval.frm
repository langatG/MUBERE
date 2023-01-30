VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPayApproval 
   BackColor       =   &H00FFC0FF&
   Caption         =   "Payment Approval"
   ClientHeight    =   7590
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9150
   LinkTopic       =   "Form1"
   ScaleHeight     =   7590
   ScaleWidth      =   9150
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   7680
      TabIndex        =   2
      Top             =   7080
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   3240
      Width           =   855
   End
   Begin MSComctlLib.ListView lvwItems 
      Height          =   2655
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   4683
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "InvId"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "RNo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "LPO#"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Invoice Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Amount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Description"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Vendor"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   135
      Left            =   -120
      TabIndex        =   4
      Top             =   6960
      Visible         =   0   'False
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.ListView lvwSelectedItems 
      Height          =   2895
      Left            =   120
      TabIndex        =   7
      Top             =   3960
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   5106
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "InvId"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "RNo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "LPO#"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Invoice Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Amount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Description"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Vendor"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Items Awaiting Payment Approval"
      Height          =   615
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   8535
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Selected Items To Approve"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   3600
      Width           =   8895
   End
End
Attribute VB_Name = "frmPayApproval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdAdd_Click()
If lvwItems.ListItems.Count = 0 Then
    MsgBox "There are no items", vbInformation, "NO ITEMS"
        lvwItems.SetFocus
    Exit Sub
End If

Set li = lvwSelectedItems.ListItems.Add(, , lvwItems.SelectedItem)
                        li.SubItems(1) = lvwItems.SelectedItem.ListSubItems(1) & ""
                        li.SubItems(2) = lvwItems.SelectedItem.ListSubItems(2) & ""
                        li.SubItems(3) = lvwItems.SelectedItem.ListSubItems(3) & ""
                        li.SubItems(4) = lvwItems.SelectedItem.ListSubItems(4) & ""
                        li.SubItems(5) = lvwItems.SelectedItem.ListSubItems(5) & ""

lvwItems.ListItems.Remove (lvwItems.SelectedItem.Index)
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdRemove_Click()

If lvwSelectedItems.ListItems.Count = 0 Then
    MsgBox "There are no items", vbInformation, "NO ITEMS"
        lvwSelectedItems.SetFocus
    Exit Sub
End If

Set li = lvwItems.ListItems.Add(, , lvwSelectedItems.SelectedItem)
                        li.SubItems(1) = lvwSelectedItems.SelectedItem.ListSubItems(1) & ""
                        li.SubItems(2) = lvwSelectedItems.SelectedItem.ListSubItems(2) & ""
                        li.SubItems(3) = lvwSelectedItems.SelectedItem.ListSubItems(3) & ""
                        li.SubItems(4) = lvwSelectedItems.SelectedItem.ListSubItems(4) & ""
                        li.SubItems(5) = lvwSelectedItems.SelectedItem.ListSubItems(5) & ""

lvwSelectedItems.ListItems.Remove (lvwSelectedItems.SelectedItem.Index)

End Sub

Private Sub cmdsave_Click()
End Sub

Private Sub Form_Load()
Set rs = oSaccoMaster.GetRecordset("d_sp_PayReqApprove")

If rs.EOF Then
Exit Sub
End If

lvwItems.ListItems.Clear

While Not rs.EOF
If Not IsNull(rs.Fields(0)) Then
Set li = lvwItems.ListItems.Add(, , rs.Fields(0))
End If
                    If Not IsNull(rs.Fields(6)) Then li.SubItems(1) = rs.Fields(6) & ""
                        
                    If Not IsNull(rs.Fields(5)) Then li.SubItems(2) = rs.Fields(5) & ""
                    If Not IsNull(rs.Fields(2)) Then li.SubItems(3) = rs.Fields(2) & ""
                    If Not IsNull(rs.Fields(3)) Then li.SubItems(4) = rs.Fields(3) & ""
                    If Not IsNull(rs.Fields(4)) Then li.SubItems(5) = rs.Fields(4) & ""
                    If Not IsNull(rs.Fields(1)) Then li.SubItems(6) = rs.Fields(1) & ""
rs.MoveNext
Wend
 


End Sub


Private Sub lvwSelectedItems_DblClick()

If lvwSelectedItems.ListItems.Count = 0 Then
     Exit Sub
End If

With frmApprovePayReq
.lblInvId = lvwSelectedItems.SelectedItem
.lblLPONo = lvwSelectedItems.SelectedItem.ListSubItems(2) & ""
.txtestimate = lvwSelectedItems.SelectedItem.ListSubItems(4) & ""
End With

lvwSelectedItems.ListItems.Remove (lvwSelectedItems.SelectedItem.Index)
frmApprovePayReq.Show vbModal
End Sub
