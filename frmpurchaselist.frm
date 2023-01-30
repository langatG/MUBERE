VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmpurchaselist 
   Caption         =   "Items Already Ordered"
   ClientHeight    =   5430
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   7590
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   135
      Left            =   0
      TabIndex        =   9
      Top             =   4800
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   2520
      Width           =   1095
   End
   Begin MSComctlLib.ListView Lvwitems 
      Height          =   1815
      Left            =   0
      TabIndex        =   7
      Top             =   480
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   3201
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      MousePointer    =   4
      MouseIcon       =   "frmpurchaselist.frx":0000
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ReqID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Item Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Quantity"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Price"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lvwselecteditems 
      Height          =   1755
      Left            =   0
      TabIndex        =   8
      Top             =   3000
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   3096
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ReqID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Item Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Quantity"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "POCOST"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C000C0&
      Caption         =   "Items to be Ordered"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   0
      TabIndex        =   1
      Top             =   2280
      Width           =   7575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C000C0&
      Caption         =   "Order List Awaiting Approval"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   555
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9240
   End
End
Attribute VB_Name = "frmpurchaselist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
If Lvwitems.ListItems.Count = 0 Then
    MsgBox "There are no items", vbInformation, "NO ITEMS"
        Lvwitems.SetFocus
    Exit Sub
End If

Set li = lvwselecteditems.ListItems.Add(, , Lvwitems.SelectedItem)
                        li.SubItems(1) = Lvwitems.SelectedItem.ListSubItems(1) & ""
                        li.SubItems(2) = Lvwitems.SelectedItem.ListSubItems(2) & ""
                        li.SubItems(3) = Lvwitems.SelectedItem.ListSubItems(3) & ""

Lvwitems.ListItems.Remove (Lvwitems.SelectedItem.Index)
End Sub
Private Sub cmdClear_Click()
Form_Load
End Sub
Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
Dim j As Integer
j = 1

If lvwselecteditems.ListItems.Count = 0 Then
    MsgBox "There are no records to save."
        cmdSave.SetFocus
    Exit Sub
End If


    ProgressBar1.Max = lvwselecteditems.ListItems.Count
 ProgressBar1.value = 0
 
    Do While Not j > lvwselecteditems.ListItems.Count
        
        ProgressBar1.value = j
        lvwselecteditems.ListItems.Item(j).selected = True
        oSaccoMaster.ExecuteThis ("Update d_Requisition SET Status='Ordered' WHERE RNo='" & lvwselecteditems.SelectedItem & "'")
       
        j = j + 1
    Loop
   MsgBox "Records saved successively."
   Set rs = oSaccoMaster.GetRecordset("SELECT Vendor FROM d_LPO WHERE RefNo ='" & lvwselecteditems.SelectedItem & "'")
    reportname = "d_LPO.rpt"
    If Not IsNull(rs.Fields(0)) Then
    STRFORMULA = "{d_LPO.Vendor}='" & rs.Fields(0) & "' and {d_Requisition.Status}='Ordered'"
    Show_Sales_Crystal_Report STRFORMULA, reportname, ""
    End If
    
    lvwselecteditems.ListItems.Clear
  
  'd_LPO.rpt
  
  
  ProgressBar1.value = 0
            

End Sub
Private Sub Form_Load()
lvwselecteditems.ListItems.Clear
Lvwitems.ListItems.Clear

 sql = "SELECT     rno,transdate,iname,qnty,pricing  FROM d_Requisition where  status='Order'"
Set rs = oSaccoMaster.GetRecordset(sql)
            
            With rs
                If .RecordCount > 0 Then
                    .MoveFirst
                    Do While Not .EOF
                    Set li = Lvwitems.ListItems.Add(, , !Rno)
                        li.SubItems(1) = !iname & ""
                        li.SubItems(2) = !qnty & ""
                        li.SubItems(3) = (!pricing) * (!qnty) & ""
                        .MoveNext
                    Loop
                End If
            End With
End Sub

