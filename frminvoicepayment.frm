VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frminvoicepayment 
   Caption         =   "Invoice Payment"
   ClientHeight    =   9795
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   ScaleHeight     =   9795
   ScaleWidth      =   6870
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkCash 
      Caption         =   "Cash"
      Height          =   255
      Left            =   1920
      TabIndex        =   23
      Top             =   3120
      Width           =   735
   End
   Begin VB.ComboBox cboChkAcc 
      Height          =   315
      Left            =   2760
      TabIndex        =   20
      Top             =   3600
      Width           =   2655
   End
   Begin VB.TextBox txtchkVNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2760
      TabIndex        =   19
      Top             =   3120
      Width           =   2655
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   4800
      TabIndex        =   12
      Top             =   9360
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   1560
      TabIndex        =   11
      Top             =   9360
      Width           =   1095
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   9360
      Width           =   975
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   375
      Left            =   4320
      TabIndex        =   9
      Top             =   6480
      Width           =   1335
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   2760
      TabIndex        =   8
      Top             =   6480
      Width           =   1335
   End
   Begin VB.TextBox txtAmnt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2760
      TabIndex        =   7
      Text            =   "0.00"
      Top             =   1680
      Width           =   2655
   End
   Begin VB.ComboBox cboVendor 
      Height          =   315
      Left            =   2760
      TabIndex        =   6
      Top             =   1200
      Width           =   2655
   End
   Begin VB.TextBox txtRef 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2760
      TabIndex        =   5
      Top             =   720
      Width           =   2655
   End
   Begin MSComCtl2.DTPicker dtpPayDate 
      Height          =   255
      Left            =   2760
      TabIndex        =   18
      Top             =   2160
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   450
      _Version        =   393216
      Format          =   90701825
      CurrentDate     =   40108
   End
   Begin MSComctlLib.ListView lvwSelectedItems 
      Height          =   2415
      Left            =   120
      TabIndex        =   21
      Top             =   6960
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   4260
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
      BackColor       =   -2147483624
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "InvId"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "LPO#"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Inv Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Amount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Balance"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lvwItems 
      Height          =   2055
      Left            =   120
      TabIndex        =   22
      Top             =   4320
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   3625
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "InvId"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "LPO#"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Invoice Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Amount"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label4 
      Caption         =   "Bank Account/Paying Account"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   3720
      Width           =   2295
   End
   Begin VB.Label lblAccChk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2760
      TabIndex        =   24
      Top             =   2640
      Width           =   2655
   End
   Begin VB.Label Label10 
      Caption         =   "Payment Date"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label9 
      Caption         =   "Check/Voucher No"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label8 
      Caption         =   "Checking Account"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Seleceted Invoice Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   14
      Top             =   6480
      Width           =   6615
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Invoice Avalable"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   4080
      Width           =   6495
   End
   Begin VB.Label Label5 
      Caption         =   "Amount"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Vendor 
      Caption         =   "Vendor"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Reference"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Payment Details"
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
      TabIndex        =   1
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Process Invoice Payment "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "frminvoicepayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim glaccno999 As String


Private Sub cboVendor_Change()
cmdClear_Click
Lvwitems.ListItems.Clear
If Trim(cboVendor) = "" Then
Exit Sub
End If
Set rs = oSaccoMaster.GetRecordset("d_sp_InvVendor '" & cboVendor & "'")
While Not rs.EOF
If Not IsNull(rs.Fields(0)) Then
Set li = Lvwitems.ListItems.Add(, , rs.Fields(0))
End If
If Not IsNull(rs.Fields(1)) Then li.SubItems(1) = rs.Fields(1) & ""
If Not IsNull(rs.Fields(2)) Then li.SubItems(2) = rs.Fields(2) & ""
If Not IsNull(rs.Fields(3)) Then li.SubItems(3) = rs.Fields(3) & ""
rs.MoveNext
Wend
End Sub


Private Sub cboVendor_Click()
cboVendor_Change
End Sub

Private Sub Check1_Click()

End Sub

Private Sub chkCash_Click()
If chkCash = vbChecked Then
txtchkVNo = "CASH"
txtchkVNo.Enabled = False
Else
txtchkVNo = ""
txtchkVNo.Enabled = True
End If

End Sub

Private Sub chkCash_Validate(Cancel As Boolean)
chkCash_Click
End Sub

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
                        li.SubItems(4) = 0# & ""
                        
Lvwitems.ListItems.Remove (Lvwitems.SelectedItem.Index)
'//get the checking account
sql = ""
sql = "SELECT     *   FROM         d_Approve2  WHERE     (RNo = '" & li & "') and approved=1   ORDER BY id DESC"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
lblAccChk = rs.Fields("glacc")
End If
lvwSelectedItems_DblClick
End Sub

Private Sub cmdClear_Click()
lvwselecteditems.ListItems.Clear
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdRemove_Click()

If lvwselecteditems.ListItems.Count = 0 Then
    MsgBox "There are no items", vbInformation, "NO ITEMS"
        lvwselecteditems.SetFocus
    Exit Sub
End If

Set li = Lvwitems.ListItems.Add(, , lvwselecteditems.SelectedItem)
                        li.SubItems(1) = lvwselecteditems.SelectedItem.ListSubItems(1) & ""
                        li.SubItems(2) = lvwselecteditems.SelectedItem.ListSubItems(2) & ""
                        li.SubItems(3) = lvwselecteditems.SelectedItem.ListSubItems(3) & ""
                        
                        
lvwselecteditems.ListItems.Remove (lvwselecteditems.SelectedItem.Index)

End Sub

Private Sub cmdsave_Click()
Dim j As Integer
j = 1

If lvwselecteditems.ListItems.Count = 0 Then
    MsgBox "There are no records to save."
        cmdsave.SetFocus
    Exit Sub
End If

If Trim(txtAmnt) = "" Then
    MsgBox "Please enter amount."
        txtAmnt.SetFocus
    Exit Sub
End If

If Trim(txtRef) = "" Then
    MsgBox "Please enter the reference number."
        txtRef.SetFocus
    Exit Sub
End If

If (chkCash = vbUnchecked) And (Trim(txtchkVNo) = "") Then
    MsgBox "Please enter the voucher/cheque number."
        txtchkVNo.SetFocus
    Exit Sub
End If

If CCur(txtAmnt) > CCur(lvwselecteditems.SelectedItem.ListSubItems(3)) Then
    MsgBox "The amount cannot be greater than the invoiced amount."
        txtAmnt.SetFocus
    Exit Sub
End If

If Len(cboChkAcc) > 5 Then
glaccno999 = Mid(cboChkAcc, 1, 6)
End If



Do While Not j > lvwselecteditems.ListItems.Count

    
'        ProgressBar1.Visible = True
'        ProgressBar1.value = J
        Dim cate As String
        If chkCash = vbChecked Then
        cate = "CASH"
        Else
        cate = "CHEQUE"
        End If
        
        
        
    If Not Save_GLTRANSACTION(Format(dtpPayDate, "dd/mm/yyyy"), (CCur(txtAmnt)), lblAccChk, glaccno999, txtRef, "Invoice Payment", User, ErrorMessage, cboVendor, 1, 1, txtchkVNo, TransNo) Then
            If ErrorMessage <> "" Then
                MsgBox ErrorMessage, vbInformation, Me.Caption
                ErrorMessage = ""
            End If
        End If
        
    
        
       sql = "[d_sp_Payment] '" & txtRef & "','" & lvwselecteditems.SelectedItem & "'," & lvwselecteditems.SelectedItem.ListSubItems(1)
       sql = sql & ", " & txtAmnt & ",'" & cate & "','" & cboVendor & "','" & dtpPayDate
       sql = sql & "','" & glaccno999 & "','" & txtchkVNo & "','" & cate & "'," & CCur(lvwselecteditems.SelectedItem.ListSubItems(4)) & ",'" & User & "'"
       oSaccoMaster.ExecuteThis (sql)
'     " & lvwSelectedItems.SelectedItem.ListSubItems(1) & " ',' " & lvwSelectedItems.SelectedItem & "', '" & lvwSelectedItems.SelectedItem.ListSubItems(2) & "', '" & lvwSelectedItems.SelectedItem.ListSubItems(3) & "'," & lvwSelectedItems.SelectedItem.ListSubItems(4) & ", '" & lvwSelectedItems.SelectedItem.ListSubItems(5) & "','" & User & "'")
        j = j + 1
    Loop

  lvwselecteditems.ListItems.Remove (lvwselecteditems.SelectedItem.Index)
  MsgBox "Records saved successively."

End Sub

Private Sub Form_Load()
cboChkAcc.Clear
cboVendor.Clear
dtpPayDate = Format(Get_Server_Date, "dd/mm/yyyy")

Set rs = oSaccoMaster.GetRecordset("SELECT DISTINCT Vendor FROM d_Invoice WHERE Paid=0")
While Not rs.EOF
If Not IsNull(rs.Fields(0)) Then cboVendor.AddItem (rs.Fields(0))

rs.MoveNext
Wend

cboVendor = "<Select Vendor>"



Set rs = oSaccoMaster.GetRecordset("d_sp_InvAcc")
While Not rs.EOF
If Not IsNull(rs.Fields(0)) Then cboChkAcc.AddItem (rs.Fields(0) & "-" & rs.Fields(1))

rs.MoveNext
Wend

cboChkAcc = "<Select Checking Account>"
End Sub

Private Sub lvwSelectedItems_DblClick()
Set rs = oSaccoMaster.GetRecordset("SELECT RNo From d_Invoice WHERE (InvId = '" & lvwselecteditems.SelectedItem & "')")
If rs.EOF Then Exit Sub
If Not IsNull(rs.Fields(0)) Then txtRef = rs.Fields(0)
txtAmnt = CCur(lvwselecteditems.SelectedItem.ListSubItems(3))

Set rs = oSaccoMaster.GetRecordset("SELECT glAcc From d_Approve2 WHERE (RNo = '" & txtRef & "')")
If Not IsNull(rs.Fields(0)) Then
lblAccChk = rs.Fields(0)
End If

End Sub

Private Sub txtAmnt_Click()
If txtAmnt = "0.00" Then
txtAmnt = ""
txtAmnt.SetFocus
End If
End Sub

Private Sub txtAmnt_KeyPress(KeyAscii As Integer)
If (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii = 48) Or (KeyAscii = 49) Or (KeyAscii = 50) Or (KeyAscii = 51) Or (KeyAscii = 52) Or (KeyAscii = 53) Or (KeyAscii = 54) Or (KeyAscii = 55) Or (KeyAscii = 56) Or (KeyAscii = 57) Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
MsgBox "Please enter a number "
End If
End Sub

Private Sub txtAmnt_Validate(Cancel As Boolean)
If Trim(txtAmnt) = "" Then
txtAmnt = "0"
End If
txtAmnt = Format(txtAmnt, "#0.00")

End Sub

Private Sub txtRef_Validate(Cancel As Boolean)
' d_sp_SelPayment @RNo
If Trim(txtRef) = "" Then
Exit Sub
End If


Set rs = oSaccoMaster.GetRecordset("d_sp_SelPayment '" & txtRef & "'")
If Not rs.EOF Then

If Not IsNull(rs.Fields(2)) Then txtAmnt = rs.Fields(2)
If Not IsNull(rs.Fields(4)) Then dtpPayDate = rs.Fields(4)
If Not IsNull(rs.Fields(5)) Then cboChkAcc = rs.Fields(5)
If Not IsNull(rs.Fields(6)) Then
If rs.Fields(6) = "CASH" Then
chkCash.value = vbChecked
txtchkVNo = "CASH"
Else
chkCash.value = vbUnchecked
txtchkVNo = rs.Fields(6)
End If
If Not IsNull(rs.Fields(1)) Then cboVendor = rs.Fields(1)
End If
Else

cboVendor = "<Select Vendor>"
txtAmnt = "0.00"
dtpPayDate = Get_Server_Date
cboChkAcc = ""
txtchkVNo = ""
chkCash.value = vbUnchecked



End If


End Sub
