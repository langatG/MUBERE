VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FRM 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Purchase Order"
   ClientHeight    =   6915
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   9150
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdLPO 
      Caption         =   "LPO"
      Height          =   375
      Left            =   5520
      TabIndex        =   26
      Top             =   6480
      Width           =   735
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      Height          =   375
      Left            =   4320
      TabIndex        =   17
      Top             =   6480
      Width           =   975
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
      Height          =   375
      Left            =   3240
      TabIndex        =   16
      Top             =   6480
      Width           =   975
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   2160
      TabIndex        =   15
      Top             =   6480
      Width           =   975
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   10610
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Purchase Order Info"
      TabPicture(0)   =   "frmpurchaseorders.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtRemarks"
      Tab(0).Control(1)=   "txtLPOSerial"
      Tab(0).Control(2)=   "DTPduedate"
      Tab(0).Control(3)=   "DTPlpodate"
      Tab(0).Control(4)=   "cbovendors"
      Tab(0).Control(5)=   "txtPoNo"
      Tab(0).Control(6)=   "lblIName"
      Tab(0).Control(7)=   "Label12"
      Tab(0).Control(8)=   "Label9"
      Tab(0).Control(9)=   "Label8"
      Tab(0).Control(10)=   "Label7"
      Tab(0).Control(11)=   "Label6"
      Tab(0).Control(12)=   "Label5"
      Tab(0).Control(13)=   "Label4"
      Tab(0).Control(14)=   "Label3"
      Tab(0).ControlCount=   15
      TabCaption(1)   =   "Purchase Order Items"
      TabPicture(1)   =   "frmpurchaseorders.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label10"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label11"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Lvwitems"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lvwselecteditems"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmdadd"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdremove"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      Begin VB.CommandButton cmdremove 
         Caption         =   "Remove"
         Height          =   375
         Left            =   6240
         TabIndex        =   23
         Top             =   2700
         Width           =   975
      End
      Begin VB.CommandButton cmdadd 
         Caption         =   "Add"
         Height          =   375
         Left            =   5040
         TabIndex        =   22
         Top             =   2700
         Width           =   975
      End
      Begin MSComctlLib.ListView lvwselecteditems 
         Height          =   1755
         Left            =   240
         TabIndex        =   19
         Top             =   3240
         Width           =   7815
         _ExtentX        =   13785
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
      Begin MSComctlLib.ListView Lvwitems 
         Height          =   1815
         Left            =   240
         TabIndex        =   18
         Top             =   780
         Width           =   7935
         _ExtentX        =   13996
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
      Begin VB.TextBox txtRemarks 
         Height          =   2175
         Left            =   -74640
         TabIndex        =   13
         Top             =   3300
         Width           =   7575
      End
      Begin VB.TextBox txtLPOSerial 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   -70680
         TabIndex        =   12
         Top             =   2820
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker DTPduedate 
         Height          =   375
         Left            =   -70680
         TabIndex        =   8
         Top             =   2340
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         Format          =   118685697
         CurrentDate     =   40110
      End
      Begin MSComCtl2.DTPicker DTPlpodate 
         Height          =   375
         Left            =   -70680
         TabIndex        =   7
         Top             =   1860
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         Format          =   118685697
         CurrentDate     =   40110
      End
      Begin VB.ComboBox cbovendors 
         Height          =   315
         Left            =   -70680
         TabIndex        =   6
         Top             =   1380
         Width           =   2295
      End
      Begin VB.TextBox txtPoNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   375
         Left            =   -70680
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   420
         Width           =   2295
      End
      Begin VB.Label lblIName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -70680
         TabIndex        =   25
         Top             =   960
         Width           =   4455
      End
      Begin VB.Label Label12 
         Caption         =   "Item Name :"
         Height          =   375
         Left            =   -72480
         TabIndex        =   24
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "Selected Items"
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
         TabIndex        =   21
         Top             =   2940
         Width           =   1695
      End
      Begin VB.Label Label10 
         Caption         =   "Available Items"
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
         TabIndex        =   20
         Top             =   420
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Remarks"
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
         Left            =   -74640
         TabIndex        =   14
         Top             =   3060
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "LPO Serial No"
         Height          =   375
         Left            =   -72480
         TabIndex        =   11
         Top             =   2940
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Due Date"
         Height          =   375
         Left            =   -72480
         TabIndex        =   10
         Top             =   2460
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "LPO Date"
         Height          =   375
         Left            =   -72480
         TabIndex        =   9
         Top             =   1980
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Vendor"
         Height          =   375
         Left            =   -72480
         TabIndex        =   5
         Top             =   1500
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "PO No"
         Height          =   375
         Left            =   -72480
         TabIndex        =   3
         Top             =   540
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Header"
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
         Left            =   -74760
         TabIndex        =   2
         Top             =   420
         Width           =   855
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Purchase Orders"
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
      TabIndex        =   1
      Top             =   0
      Width           =   9015
   End
End
Attribute VB_Name = "FRM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Sub Command3_Click()
Unload Me
End Sub

Private Sub cmdadd_Click()
On Error GoTo ErrorHandler
Set li = lvwselecteditems.ListItems.Add(, , Lvwitems.SelectedItem)
                        li.SubItems(1) = Lvwitems.SelectedItem.ListSubItems(1) & ""
                        li.SubItems(2) = Lvwitems.SelectedItem.ListSubItems(2) & ""
                        li.SubItems(3) = Lvwitems.SelectedItem.ListSubItems(3) & ""

Lvwitems.ListItems.Remove (Lvwitems.SelectedItem.Index)
Exit Sub

ErrorHandler:
MsgBox err.description
End Sub

Private Sub cmdClear_Click()
Form_Load
End Sub

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdLPO_Click()
frmLPO.Show vbModal
Unload Me
End Sub

Private Sub cmdRemove_Click()
On Error GoTo ErrorHandler
Set li = Lvwitems.ListItems.Add(, , lvwselecteditems.SelectedItem)
                        li.SubItems(1) = lvwselecteditems.SelectedItem.ListSubItems(1) & ""
                        li.SubItems(2) = lvwselecteditems.SelectedItem.ListSubItems(2) & ""
                        li.SubItems(3) = lvwselecteditems.SelectedItem.ListSubItems(3) & ""

lvwselecteditems.ListItems.Remove (lvwselecteditems.SelectedItem.Index)  '// removes the selected item
Exit Sub
ErrorHandler:
MsgBox err.description
End Sub

Private Sub cmdsave_Click()

If lblIName = "" Then
    MsgBox "No item is selected."
        SSTab1.Tab = 1
    Exit Sub
End If

If cbovendors.Text = "" Then
    MsgBox "Please select vendor."
        cbovendors.SetFocus
    Exit Sub
End If

'd_sp_LPO @PNo bigint, @TransDate varchar(12),@DueDate varchar(12), @Serial varchar(100),@RefNo varchar(50),@user varchar(35) as
oSaccoMaster.ExecuteThis ("d_sp_LPO " & txtPoNo & ",'" & DTPlpodate & "','" & DTPduedate & "','" & txtLPOSerial & "','" & lvwselecteditems.SelectedItem & "','" & User & "','" & txtRemarks & "','" & cbovendors & "'")
MsgBox "Records saved successfully!"
Form_Load

End Sub

Private Sub Form_Load()

lvwselecteditems.ListItems.Clear
Lvwitems.ListItems.Clear
txtLPOSerial = ""
lblIName = ""
cbovendors.Text = ""
txtRemarks = ""

DTPduedate = Format(Get_Server_Date + 7, "dd/mm/yyyy")
DTPlpodate = Format(Get_Server_Date, "dd/mm/yyyy")



 sql = "SELECT     rno,transdate,iname,qnty,pricing  FROM         d_Requisition where  status='Approved'"
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
            '//LOAD THE VENDORS
sql = "SELECT  CompanyName  FROM         ag_Supplier1 order by companyname"
Set rs = oSaccoMaster.GetRecordset(sql)
                Do While Not rs.EOF
                cbovendors.AddItem rs.Fields(0)
                        rs.MoveNext
                    Loop

Set rs = oSaccoMaster.GetRecordset("d_sp_PoNo")
If Not rs.EOF Then
txtPoNo = CCur(rs.Fields(0)) + 1
Else
txtPoNo = "1"
End If

End Sub

Private Sub lvwSelectedItems_DblClick()
lblIName = lvwselecteditems.SelectedItem.ListSubItems(1)
SSTab1.Tab = 0
End Sub
