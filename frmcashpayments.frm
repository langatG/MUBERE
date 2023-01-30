VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmcashpayments 
   Caption         =   "Cash Payments"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8160
   ScaleWidth      =   5880
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture3 
      Height          =   285
      Left            =   2280
      Picture         =   "frmcashpayments.frx":0000
      ScaleHeight     =   225
      ScaleWidth      =   195
      TabIndex        =   24
      Top             =   4440
      Width           =   255
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   3120
      TabIndex        =   23
      Top             =   7680
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   1800
      TabIndex        =   22
      Top             =   7680
      Width           =   975
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   375
      Left            =   2880
      TabIndex        =   21
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   1320
      TabIndex        =   20
      Top             =   5280
      Width           =   1215
   End
   Begin VB.TextBox txtDrAmount 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   19
      Top             =   4920
      Width           =   2175
   End
   Begin VB.TextBox txtDRacc 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   17
      Top             =   4440
      Width           =   855
   End
   Begin VB.ComboBox cboBCashAcc 
      Height          =   315
      Left            =   1800
      TabIndex        =   13
      Text            =   "<Select Bank/Cash Account>"
      Top             =   1920
      Width           =   2535
   End
   Begin MSComCtl2.DTPicker dtpTransDate 
      Height          =   255
      Left            =   1800
      TabIndex        =   12
      Top             =   1440
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   450
      _Version        =   393216
      Format          =   122355713
      CurrentDate     =   40108
   End
   Begin MSComctlLib.ListView lvwItems 
      Height          =   1935
      Left            =   120
      TabIndex        =   11
      Top             =   5760
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   3413
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   15787756
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Pay#"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "TransDate"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "CrAccNo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "DrAccNo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "VNo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Payee"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Description"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Amount"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox txtVNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1800
      TabIndex        =   10
      Top             =   3360
      Width           =   2535
   End
   Begin VB.TextBox txtDesc 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1800
      TabIndex        =   8
      Top             =   2880
      Width           =   2535
   End
   Begin VB.TextBox txtPayee 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1800
      TabIndex        =   6
      Top             =   2400
      Width           =   2535
   End
   Begin VB.TextBox txtPayNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Top             =   960
      Width           =   2535
   End
   Begin VB.Label lblAccNo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   2520
      TabIndex        =   25
      Top             =   4440
      Width           =   3315
   End
   Begin VB.Label Label12 
      Caption         =   "Amount"
      Height          =   375
      Left            =   0
      TabIndex        =   18
      Top             =   4920
      Width           =   855
   End
   Begin VB.Label Label11 
      Caption         =   "Account"
      Height          =   375
      Left            =   0
      TabIndex        =   16
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ITEMS(DR)"
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
      Height          =   495
      Left            =   0
      TabIndex        =   15
      Top             =   3840
      Width           =   5895
   End
   Begin VB.Label Label8 
      Caption         =   "Voucher/Check No"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label7 
      Caption         =   "Description"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "Payee"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Bank/Cash Account"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Date"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Payment Number"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Payment Details (CR)"
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
      TabIndex        =   1
      Top             =   360
      Width           =   5775
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Cash Payments"
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
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5895
   End
End
Attribute VB_Name = "frmcashpayments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
If cboBCashAcc = "<Select Bank/Cash Account>" Then
MsgBox "Please select the credit account."
cboBCashAcc.SetFocus
Exit Sub
End If

If Trim(txtPayNo) = "" Then
MsgBox "Please enter the Pay Number."
txtPayNo.SetFocus
Exit Sub
End If

If Trim(txtPayee) = "" Then
MsgBox "Please enter the Payee."
txtPayee.SetFocus
Exit Sub
End If

If Trim(txtdesc) = "" Then
MsgBox "Please enter the Description."
txtdesc.SetFocus
Exit Sub
End If

If Trim(txtvno) = "" Then
MsgBox "Please enter the Voucher Number."
txtvno.SetFocus
Exit Sub
End If

If Trim(txtdracc) = "" Then
MsgBox "Please enter the debit account."
txtdracc.SetFocus
Exit Sub
End If

If Len(cboBCashAcc) > 5 Then
GlAccNo1 = Mid(cboBCashAcc, 1, 4)
End If

Set li = Lvwitems.ListItems.Add(, , txtPayNo)
                        li.SubItems(1) = Format(DTPTransdate, "dd/mm/yyyy") & ""
                        li.SubItems(2) = GlAccNo1 & ""
                        li.SubItems(3) = txtdracc & ""
                        li.SubItems(4) = txtvno & ""
                        li.SubItems(5) = txtPayee & ""
                        li.SubItems(6) = txtdesc & ""
                        li.SubItems(7) = txtDrAmount & ""
End Sub

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdsave_Click()
Dim j As Integer
j = 1

If Lvwitems.ListItems.Count = 0 Then
    MsgBox "There are no records to save."
        cmdAdd.SetFocus
    Exit Sub
End If


   Do While Not j > Lvwitems.ListItems.Count
   
        If Not Save_GLTRANSACTION(Format(Lvwitems.SelectedItem.SubItems(1), "dd/mm/yyyy"), (CCur(Lvwitems.SelectedItem.SubItems(7))), Lvwitems.SelectedItem.SubItems(3), Lvwitems.SelectedItem.SubItems(2), Lvwitems.SelectedItem.SubItems(4), Lvwitems.SelectedItem.SubItems(5), User, ErrorMessage, Lvwitems.SelectedItem.SubItems(6), 1, 1, Lvwitems.SelectedItem.SubItems(4), TransNo) Then
            If ErrorMessage <> "" Then
                MsgBox ErrorMessage, vbInformation, Me.Caption
                ErrorMessage = ""
            End If
        End If
  
       oSaccoMaster.ExecuteThis ("d_sp_CashPay '" & Lvwitems.SelectedItem & "','" & Lvwitems.SelectedItem.SubItems(1) & "','" & Lvwitems.SelectedItem.SubItems(2) & "','" & Lvwitems.SelectedItem.SubItems(3) & "','" & Lvwitems.SelectedItem.SubItems(5) & "','" & Lvwitems.SelectedItem.SubItems(6) & "','" & Lvwitems.SelectedItem.SubItems(4) & "'," & Lvwitems.SelectedItem.SubItems(7) & "")
        j = j + 1
    Loop
   MsgBox "Records saved successively."
    
    Lvwitems.ListItems.Clear
  

End Sub

Private Sub Form_Load()
DTPTransdate = Get_Server_Date

Set rs = oSaccoMaster.GetRecordset("d_sp_InvAcc")
While Not rs.EOF
If Not IsNull(rs.Fields(0)) Then cboBCashAcc.AddItem (rs.Fields(0) & "-" & rs.Fields(1))

rs.MoveNext
Wend

cboBCashAcc = "<Select Bank/Cash Account>"

End Sub

Private Sub Picture3_Click()
Me.MousePointer = vbHourglass
    frmsearchacc.Show vbModal
    txtdracc = sel
    get_namecr
    Me.MousePointer = 0

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
