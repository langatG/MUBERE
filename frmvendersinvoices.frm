VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmvendersinvoices 
   Caption         =   "Vendors Invoices"
   ClientHeight    =   8985
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10725
   LinkTopic       =   "Form1"
   ScaleHeight     =   8985
   ScaleWidth      =   10725
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   7800
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   9360
      TabIndex        =   6
      Top             =   7800
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   8040
      TabIndex        =   5
      Top             =   7800
      Width           =   1095
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7215
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   12726
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Invoice Header"
      TabPicture(0)   =   "frmvendersinvoices.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label6"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label7"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label8"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label9"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label10"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label11"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblRecords"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label12"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblGlAcc"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtIId"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtRef"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cboVendor"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "DTPInvDate"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "DTPInvDue"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtInvAmnt"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtdesc"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "cmdNext"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).ControlCount=   20
      TabCaption(1)   =   "LPO Details"
      TabPicture(1)   =   "frmvendersinvoices.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "LvwItems"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "LvwSelected"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdRemove"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmdAdd"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      Begin VB.CommandButton cmdNext 
         Appearance      =   0  'Flat
         Caption         =   "Next"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2880
         MaskColor       =   &H000000C0&
         MouseIcon       =   "frmvendersinvoices.frx":0038
         Picture         =   "frmvendersinvoices.frx":018A
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   5520
         Width           =   615
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   375
         Left            =   -70440
         TabIndex        =   25
         Top             =   3000
         Width           =   1095
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove"
         Height          =   375
         Left            =   -69120
         TabIndex        =   24
         Top             =   3000
         Width           =   1095
      End
      Begin VB.TextBox txtdesc 
         Appearance      =   0  'Flat
         Height          =   1215
         Left            =   1800
         TabIndex        =   23
         Top             =   4260
         Width           =   4215
      End
      Begin VB.TextBox txtInvAmnt 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1800
         TabIndex        =   21
         Top             =   3780
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker DTPInvDue 
         Height          =   375
         Left            =   1800
         TabIndex        =   18
         Top             =   3180
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         Format          =   131334145
         CurrentDate     =   40112
      End
      Begin MSComCtl2.DTPicker DTPInvDate 
         Height          =   375
         Left            =   1800
         TabIndex        =   16
         Top             =   1980
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         Format          =   131334145
         CurrentDate     =   40112
      End
      Begin VB.ComboBox cboVendor 
         Height          =   315
         Left            =   1800
         TabIndex        =   15
         Text            =   "<Select Vendor>"
         Top             =   1020
         Width           =   1935
      End
      Begin VB.TextBox txtRef 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   375
         Left            =   1800
         TabIndex        =   13
         Top             =   540
         Width           =   1935
      End
      Begin VB.TextBox txtIId 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1800
         TabIndex        =   11
         Top             =   1500
         Width           =   1935
      End
      Begin MSComctlLib.ListView LvwSelected 
         Height          =   2415
         Left            =   -74760
         TabIndex        =   4
         Top             =   3540
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   4260
         View            =   3
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "LPO NO"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Item name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "GL Account"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Quantity"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Amount"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView LvwItems 
         Height          =   2415
         Left            =   -74760
         TabIndex        =   3
         Top             =   420
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   4260
         View            =   3
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "LPO NO"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "LPO Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Item"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Quantity"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Cost"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lblGlAcc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1800
         TabIndex        =   29
         Top             =   2520
         Width           =   1935
      End
      Begin VB.Label Label12 
         Caption         =   "GL Account :"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label lblRecords 
         AutoSize        =   -1  'True
         Caption         =   "0 of 0 records selected"
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   3600
         TabIndex        =   27
         Top             =   5760
         Width           =   1620
      End
      Begin VB.Label Label11 
         Caption         =   "Description :"
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   4260
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Invoice Amount :"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   3780
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Invoice Due Date :"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   3300
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "Invoice Date :"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   2100
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Invoice ID :"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   1500
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Vendor :"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   1020
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Reference :"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   540
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Others"
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
         Left            =   240
         TabIndex        =   9
         Top             =   2940
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "General"
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
         Left            =   240
         TabIndex        =   8
         Top             =   180
         Width           =   1335
      End
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000009&
      Caption         =   "Recieve Vendor Invoices"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   10695
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Caption         =   "Vendor Invoices"
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
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10695
   End
End
Attribute VB_Name = "frmvendersinvoices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 Dim j As Integer
 Dim RecNo As Integer
 
Private Sub cmdAdd_Click()

If Lvwitems.ListItems.Count = 0 Then
    MsgBox "There are no items", vbInformation, "NO ITEMS"
        Lvwitems.SetFocus
    Exit Sub
End If

sql = "SELECT     dbo.d_Approve2.glAcc FROM dbo.d_LPO INNER JOIN "
sql = sql & "dbo.d_Approve2 ON dbo.d_LPO.RefNo = dbo.d_Approve2.RNo AND dbo.d_LPO.PNo = " & Lvwitems.SelectedItem & ""
Set rs = oSaccoMaster.GetRecordset(sql)

Set li = LvwSelected.ListItems.Add(, , Lvwitems.SelectedItem)
                        li.SubItems(1) = Lvwitems.SelectedItem.ListSubItems(2) & ""
                        li.SubItems(2) = rs.Fields(0) & ""
                        li.SubItems(3) = Lvwitems.SelectedItem.ListSubItems(3) & ""
                        li.SubItems(4) = Lvwitems.SelectedItem.ListSubItems(4) & ""

Lvwitems.ListItems.Remove (Lvwitems.SelectedItem.Index)




End Sub

Private Sub cmdClear_Click()
Form_Load
End Sub

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdNext_Click()
j = j + 1

If j > RecNo Then
    MsgBox "End of records."
    Exit Sub
End If

 lblRecords = j & " of " & RecNo & " records selected."
 LvwSelected.ListItems.Item(j).selected = True
 LvwSelected_DblClick
End Sub

Private Sub cmdRemove_Click()

If LvwSelected.ListItems.Count = 0 Then
    MsgBox "There are no items", vbInformation, "NO ITEMS"
        Lvwitems.SetFocus
    Exit Sub
End If

sql = "SELECT TransDate FROM dbo.d_LPO where PNo = " & LvwSelected.SelectedItem & ""
Set rs = oSaccoMaster.GetRecordset(sql)

Set li = Lvwitems.ListItems.Add(, , LvwSelected.SelectedItem)
                        li.SubItems(1) = rs.Fields(0) & ""
                        li.SubItems(2) = LvwSelected.SelectedItem.ListSubItems(1) & ""
                        li.SubItems(3) = LvwSelected.SelectedItem.ListSubItems(3) & ""
                        li.SubItems(4) = LvwSelected.SelectedItem.ListSubItems(4) & ""

LvwSelected.ListItems.Remove (LvwSelected.SelectedItem.Index)

End Sub

Private Sub cmdsave_Click()
If LvwSelected.ListItems.Count = 0 Then
    MsgBox "No item is selected."
    Exit Sub
End If


LvwSelected_DblClick

    If txtInvAmnt = "0.00" Then
        MsgBox "Please enter the amount" & vbNewLine & "Amount cannot be Kshs 0.00", vbCritical, "MISSING DETAILS"
            txtInvAmnt.SetFocus
        Exit Sub
    End If
    
    If txtRef = "" Then
        MsgBox "Please enter the reference number.", vbCritical, "MISSING DETAILS"
            txtRef.SetFocus
        Exit Sub
    End If
    
    If txtIId = "" Then
        MsgBox "Please enter the invoice Id", vbCritical, "MISSING DETAILS"
            txtIId.SetFocus
        Exit Sub
    End If
    
    If cboVendor = "<Select Vendor>" Then
        MsgBox "Please select the vendor", vbCritical, "MISSING DETAILS"
            cboVendor.SetFocus
        Exit Sub
    End If
    'd_sp_Invoice  @InvId varchar(35), @RNo varchar(35), @Vendor varchar(85), @InvDate varchar(12), @DueDate varchar(12), @Amount money, @Desc varchar(150),
'@auditid varchar(35) AS
sql = "d_sp_Invoice '" & txtIId & "','" & txtRef & "','" & cboVendor & "'"
sql = sql & ",'" & DTPInvDate & "','" & DTPInvDue & "'," & Format(txtInvAmnt, "#0.00") & ",'" & txtdesc & "','" & User & "','" & lblGlAcc & "'"

oSaccoMaster.ExecuteThis (sql)

LvwSelected.ListItems.Remove (LvwSelected.SelectedItem.Index)
txtIId = ""

MsgBox "Records saved successfully!!"

End Sub

Private Sub Form_Load()
txtRef = ""
txtInvAmnt = "0.00"
txtdesc = ""
txtIId = ""
DTPInvDate = Format(Get_Server_Date, "dd/MM/YYYy")
DTPInvDue = Format(Get_Server_Date + 14, "dd/MM/YYYy")
Lvwitems.ListItems.Clear
cboVendor.Clear


sql = "SELECT  CompanyName  FROM  ag_Supplier1 order by companyname"
Set rs = oSaccoMaster.GetRecordset(sql)
                Do While Not rs.EOF
                cboVendor.AddItem rs.Fields(0)
                        rs.MoveNext
                    Loop

Set rs = oSaccoMaster.GetRecordset("d_sp_InvoiceDet")

    Do While Not rs.EOF
        Set li = Lvwitems.ListItems.Add(, , rs.Fields(0))
                        li.SubItems(1) = rs.Fields(1) & ""
                        li.SubItems(2) = rs.Fields(2) & ""
                        li.SubItems(3) = rs.Fields(3) & ""
                        li.SubItems(4) = rs.Fields(4) & ""
                        rs.MoveNext
                        Loop

cboVendor.AddItem ("<Select Vendor>")

End Sub

Private Sub lvwItems_DblClick()
cmdAdd_Click
End Sub


Private Sub LvwSelected_DblClick()
Set rs = oSaccoMaster.GetRecordset("SELECT RefNo, Vendor From dbo.d_LPO Where PNo = " & LvwSelected.SelectedItem)
If Not IsNull(rs.Fields(0)) Then
txtRef = Trim(rs.Fields(0))
End If
If Not IsNull(rs.Fields(1)) Then
cboVendor = Trim(rs.Fields(1))
End If
lblGlAcc = LvwSelected.SelectedItem.ListSubItems(2)
txtInvAmnt = LvwSelected.SelectedItem.ListSubItems(4)
txtInvAmnt_Validate True
RecNo = LvwSelected.ListItems.Count
j = LvwSelected.SelectedItem.Index
lblRecords = j & " of " & RecNo & " records selected"
SSTab1.Tab = 0
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Tab = 0 Then
j = 1

If LvwSelected.ListItems.Count = 0 Then
        cmdsave.SetFocus
        cmdNext.Enabled = False
    Exit Sub
End If

    cmdNext.Enabled = True
    
LvwSelected_DblClick
   
    
End If
End Sub

Private Sub txtInvAmnt_Click()
txtInvAmnt = FormatCurrency(txtInvAmnt, 0, vbFalse, vbFalse, vbFalse)
If Trim(txtInvAmnt) = "0.00" Then
txtInvAmnt = ""
End If


End Sub

Private Sub txtInvAmnt_Validate(Cancel As Boolean)
If Trim(txtInvAmnt) = "" Then
txtInvAmnt = "0"
End If

txtInvAmnt = FormatCurrency(txtInvAmnt, 2, vbTrue, vbFalse, vbUseDefault)

End Sub
