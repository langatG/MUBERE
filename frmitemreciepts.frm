VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmitemreciepts 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Item Reciepts"
   ClientHeight    =   8835
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8835
   ScaleWidth      =   10665
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkPrint 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Print Receipt"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6120
      TabIndex        =   32
      Top             =   8280
      Value           =   2  'Grayed
      Width           =   1935
   End
   Begin VB.ComboBox ports 
      Height          =   315
      ItemData        =   "frmitemreciepts.frx":0000
      Left            =   9480
      List            =   "frmitemreciepts.frx":0010
      TabIndex        =   30
      Text            =   "COM1"
      Top             =   8400
      Width           =   855
   End
   Begin VB.CommandButton cmdClose 
      Appearance      =   0  'Flat
      Caption         =   "Close"
      Height          =   375
      Left            =   4440
      TabIndex        =   17
      Top             =   8280
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   2880
      TabIndex        =   16
      Top             =   8280
      Width           =   1095
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   8280
      Width           =   975
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7815
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   13785
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   8421631
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmitemreciepts.frx":002C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtOrdered"
      Tab(0).Control(1)=   "txtQnty"
      Tab(0).Control(2)=   "txtRemarks"
      Tab(0).Control(3)=   "cboStore"
      Tab(0).Control(4)=   "txtDelNo"
      Tab(0).Control(5)=   "cboVendor"
      Tab(0).Control(6)=   "txtRef"
      Tab(0).Control(7)=   "dtprecDate"
      Tab(0).Control(8)=   "Label12"
      Tab(0).Control(9)=   "Label11"
      Tab(0).Control(10)=   "Label8"
      Tab(0).Control(11)=   "Label7"
      Tab(0).Control(12)=   "Label6"
      Tab(0).Control(13)=   "Label5"
      Tab(0).Control(14)=   "Label4"
      Tab(0).Control(15)=   "Label3"
      Tab(0).Control(16)=   "Label2"
      Tab(0).ControlCount=   17
      TabCaption(1)   =   "LPO Items"
      TabPicture(1)   =   "frmitemreciepts.frx":0048
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label9"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label10"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "LBLTOTAL"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label14"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label15"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "lvwItems"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "LvwselectedItems"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "cmdAdd"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "cmdRemove"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "txtcomment"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).ControlCount=   10
      Begin VB.TextBox txtcomment 
         Height          =   615
         Left            =   7320
         TabIndex        =   34
         Top             =   4080
         Width           =   2895
      End
      Begin VB.TextBox txtOrdered 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         Height          =   375
         Left            =   -69840
         Locked          =   -1  'True
         TabIndex        =   27
         Text            =   "0"
         Top             =   2640
         Width           =   975
      End
      Begin VB.TextBox txtQnty 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   -72360
         TabIndex        =   25
         Text            =   "0"
         Top             =   2520
         Width           =   975
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove"
         Height          =   375
         Left            =   5160
         TabIndex        =   23
         Top             =   4200
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   375
         Left            =   3720
         TabIndex        =   22
         Top             =   4200
         Width           =   1215
      End
      Begin MSComctlLib.ListView LvwselectedItems 
         Height          =   2535
         Left            =   240
         TabIndex        =   19
         Top             =   4680
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   4471
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   7
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
            Text            =   "Item name"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Ordered Qty"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Delivery Qty"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Rejected Qty"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Buying Price"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lvwItems 
         Height          =   3735
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   6588
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ReceiptNo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Ord Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ProductName"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Price"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Ordered Qyt"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Delivery Qty"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Balance"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.TextBox txtRemarks 
         Height          =   1815
         Left            =   -74760
         TabIndex        =   13
         Top             =   4200
         Width           =   6975
      End
      Begin VB.ComboBox cboStore 
         Height          =   315
         Left            =   -72360
         TabIndex        =   11
         Text            =   "<Select Store>"
         Top             =   3480
         Width           =   2175
      End
      Begin VB.TextBox txtDelNo 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   -72360
         TabIndex        =   9
         Top             =   1920
         Width           =   2175
      End
      Begin VB.ComboBox cboVendor 
         Height          =   315
         Left            =   -72360
         TabIndex        =   8
         Text            =   "<Select Vendor>"
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox txtRef 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         Height          =   375
         Left            =   -72360
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   960
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker dtprecDate 
         Height          =   375
         Left            =   -72360
         TabIndex        =   10
         Top             =   3000
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         Format          =   85983233
         CurrentDate     =   40110
      End
      Begin VB.Label Label15 
         Caption         =   "Comment"
         Height          =   255
         Left            =   6600
         TabIndex        =   33
         Top             =   4320
         Width           =   735
      End
      Begin VB.Label Label14 
         Caption         =   "TOTAL"
         Height          =   375
         Left            =   7440
         TabIndex        =   29
         Top             =   7320
         Width           =   615
      End
      Begin VB.Label LBLTOTAL 
         Caption         =   "0"
         Height          =   255
         Left            =   8160
         TabIndex        =   28
         Top             =   7320
         Width           =   2175
      End
      Begin VB.Label Label12 
         Caption         =   "Quantity Ordered :"
         Height          =   255
         Left            =   -71280
         TabIndex        =   26
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "Quantity Delivered :"
         Height          =   255
         Left            =   -74640
         TabIndex        =   24
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "This Delivery"
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
         Top             =   4320
         Width           =   1575
      End
      Begin VB.Label Label9 
         Caption         =   "LPO ITEM"
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
         Left            =   5400
         TabIndex        =   20
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label8 
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
         Left            =   -74760
         TabIndex        =   14
         Top             =   3960
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Store"
         Height          =   255
         Left            =   -74640
         TabIndex        =   12
         Top             =   3480
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Recieved Date"
         Height          =   255
         Left            =   -74640
         TabIndex        =   7
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "Delivery No"
         Height          =   255
         Left            =   -74640
         TabIndex        =   6
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Vendor"
         Height          =   255
         Left            =   -74640
         TabIndex        =   5
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Reference"
         Height          =   255
         Left            =   -74640
         TabIndex        =   3
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Delivery Details"
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
         Left            =   -74760
         TabIndex        =   2
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.Label Label13 
      Caption         =   "Printer Port"
      Height          =   375
      Left            =   8400
      TabIndex        =   31
      Top             =   8400
      Width           =   855
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Item Reciept"
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
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   7575
   End
End
Attribute VB_Name = "frmitemreciepts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objLabelEdit As LabelEdit
Dim objLabelEdit2 As LabelEdit
Dim objLabelEdit3 As LabelEdit

Private Sub cboStore_KeyPress(KeyAscii As Integer)
KeyAscii = 0
Beep
End Sub

Private Sub cboVendor_Change()
Lvwitems.ListItems.Clear
Set rs = oSaccoMaster.GetRecordset("d_sp_SupplierOrderedGoods '" & cboVendor & "'")
'Set rs = oSaccoMaster.GetRecordset("d_sp_IOrdered")
While Not rs.EOF
Set li = Lvwitems.ListItems.Add(, , rs.Fields(0))
                li.SubItems(1) = rs.Fields(1) & ""
                li.SubItems(2) = rs.Fields(2) & ""
                li.SubItems(3) = rs.Fields(3) & ""
                li.SubItems(4) = rs.Fields(4) & ""
                li.SubItems(5) = "0" & ""
                li.SubItems(6) = rs.Fields(4) & ""
                
rs.MoveNext
Wend
End Sub

Private Sub cboVendor_Click()
cboVendor_Change
End Sub

Private Sub cboVendor_KeyPress(KeyAscii As Integer)
KeyAscii = 0
Beep
End Sub

Private Sub cmdAdd_Click()
 
If Lvwitems.ListItems.Count = 0 Then
    MsgBox "There is no records to add"
        cmdadd.SetFocus
    Exit Sub
End If

Set li = lvwselecteditems.ListItems.Add(, , Lvwitems.SelectedItem)
    
    Dim pp As Double
    
    Set Rst = oSaccoMaster.GetRecordset("select pprice from ag_products where p_name='" & Lvwitems.SelectedItem.ListSubItems(2) & "'")
    If Not Rst.EOF Then
    pp = Rst.Fields(0)
    End If
                        li.SubItems(1) = Lvwitems.SelectedItem.ListSubItems(1) & ""
                        li.SubItems(2) = Lvwitems.SelectedItem.ListSubItems(2) & ""
                        li.SubItems(3) = Lvwitems.SelectedItem.ListSubItems(4) & ""
                        li.SubItems(4) = Lvwitems.SelectedItem.ListSubItems(5) & ""
                        li.SubItems(5) = "0" & ""
                        li.SubItems(6) = pp

Lvwitems.ListItems.Remove (Lvwitems.SelectedItem.Index)
Calculate_Total
Exit Sub

End Sub
Sub Calculate_Total()

    Dim Total As Double, amt As Double, Price As Double, qnty As Integer
    Dim ccount As Integer
    On Error Resume Next
    Total = 0
    With lvwselecteditems
        If .ListItems.Count > 0 Then
            ccount = .ListItems.Count
            For I = 1 To ccount
                With .ListItems(I)
                        Price = CDbl(.ListSubItems(6))
                        qnty = CDbl(.ListSubItems(4))
                        amt = Price * qnty
                        Total = Total + amt
                End With
            Next I

        Else
            Total = 0
        End If
    End With
    LBLTOTAL.Caption = Total
End Sub

Private Sub cmdClear_Click()
Form_Load
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub Command4_Click()

End Sub

Private Sub cmdRemove_Click()
If lvwselecteditems.ListItems.Count = 0 Then
    MsgBox "There is no records to Remove"
        cmdRemove.SetFocus
    Exit Sub
End If


Set li = Lvwitems.ListItems.Add(, , lvwselecteditems.SelectedItem)
                        li.SubItems(1) = lvwselecteditems.SelectedItem.ListSubItems(1) & ""
                        li.SubItems(2) = lvwselecteditems.SelectedItem.ListSubItems(2) & ""
                        li.SubItems(3) = lvwselecteditems.SelectedItem.ListSubItems(3) & ""
                        li.SubItems(4) = lvwselecteditems.SelectedItem.ListSubItems(4) & ""
                        li.SubItems(5) = "0" & ""

lvwselecteditems.ListItems.Remove (lvwselecteditems.SelectedItem.Index)
Calculate_Total


End Sub

Private Sub cmdSave_Click()
Dim pcode, ReceiptNo As String, lenght As Integer
'If txtRef = "" Then
'    MsgBox "Please enter the reference number."
'        txtRef.SetFocus
'    Exit Sub
'End If
'
'If txtDelNo = "" Then
'    MsgBox "Please enter the delivery number."
'        txtDelNo.SetFocus
'    Exit Sub
'End If
'
'If CCur(txtQnty) = 0 Then
'    MsgBox "Please enter the quantity received."
'        txtQnty.SetFocus
'    Exit Sub
'End If
Dim j As Integer

'Do While Not j > (lvwselecteditems.ListItems.Count)
' 'LvwselectedItems.ListItems.Item(j).selected = True
' j = j + 1
'Loop

j = 1
For j = 1 To lvwselecteditems.ListItems.Count
' LvwselectedItems.ListItems.Item(j).selected = True

Set li = lvwselecteditems.ListItems(j)
'Refno = lvwselecteditems.ListItems(j).selected
''d_sp_Receipts @R varchar(35), @V varchar(80), @D varchar(35), @Q float, @T varchar(12), @re varchar(85), @A varchar(35) AS
'oSaccoMaster.ExecuteThis ("d_sp_Receipts '" & Refno & "','" & cboVendor & "','" & txtDelNo & "'," & lvwselecteditems.ListItems(j).SubItems(4) & ",'" & dtprecDate & "','" & txtremarks & "','" & User & "'")
''MsgBox "Records saved successfully."
''//add it to the items
'Provider = cn
'Set cn = New ADODB.Connection
'cn.Open Provider
''//get the name available
'Dim rsg As New ADODB.Recordset, rsh As New ADODB.Recordset,
'Dim namee As String
'sql = "SELECT     IName  FROM         d_Requisition  WHERE     (RNo = '" & Refno & "' and iname='" & lvwselecteditems.ListItems(j).SubItems(2) & "')"
'Set rsg = oSaccoMaster.GetRecordset(sql)
'If Not rsg.EOF Then
'namee = IIf(IsNull(rsg.Fields(0)), "", rsg.Fields(0))
'If namee <> "" Then
'        sql = ""
'        sql = "select P_CODE,p_name from ag_products where p_name like '" & lvwselecteditems.ListItems(j).SubItems(2) & "%'"
'        Set rsh = oSaccoMaster.GetRecordset(sql)
'        If Not rsh.EOF Then
'        pcode = rsh.Fields(0)
'        End If
'End If
'Else
'MsgBox "Product not available in the database, key in first in the agrovet module"
'Exit Sub
'End If

sql = "select P_CODE,qout,unserialized,pprice,sprice,o_bal from ag_products where p_name='" & lvwselecteditems.ListItems(j).SubItems(2) & "'"
Set rs = New ADODB.Recordset
rs.Open sql, cn
pcode = rs.Fields("p_code")
sql = ""
sql = "set dateformat DMY update ag_products set qin=" & lvwselecteditems.ListItems(j).SubItems(4) & ",qout=" & lvwselecteditems.ListItems(j).SubItems(4) + rs.Fields("qout") & ",o_bal=" & lvwselecteditems.ListItems(j).SubItems(4) + rs.Fields("qout") & ",last_d_updated='" & dtprecDate & "',user_id='" & User & "',audit_date='" & Get_Server_Date & "',unserialized=0,SERIA=0,pprice=" & rs.Fields("pprice") & ",sprice=" & rs.Fields("sprice") & " where p_code='" & pcode & "'"
cn.Execute sql

Dim rsst As Recordset
sql = ""
sql = "set dateformat DMY select top 1 * from ag_stockbalance where p_code='" & pcode & "' order by trackid desc "
Set rsst = New ADODB.Recordset
rsst.Open sql, cn
If Not rsst.EOF Then
sql = ""
sql = "set dateformat DMY INSERT INTO ag_stockbalance"
sql = sql & " (p_code, productname, openningstock, changeinstock, stockbalance, transdate,companyid)"
sql = sql & " VALUES     ('" & pcode & "', '" & lvwselecteditems.ListItems(j).SubItems(2) & "', '" & rs.Fields("o_bal") & "', '" & lvwselecteditems.ListItems(j).SubItems(4) & "', '" & lvwselecteditems.ListItems(j).SubItems(4) + rs.Fields("qout") & "', '" & Format(Get_Server_Date, "dd/mm/yyyy") & "',1)"
cn.Execute sql
End If

'
ReceiptNo = Trim$(li)
sql = "Update d_Requisition set [status]='Receipt' ,Qnty=" & lvwselecteditems.ListItems(j).SubItems(4) & "  where RNo='" & ReceiptNo & "'"
  oSaccoMaster.ExecuteThis (sql)

'If LvwselectedItems.ListItems.Count > 0 Then
'   LvwselectedItems.ListItems.Remove (LvwselectedItems.SelectedItem.Index)
'End If

Next j
If chkPrint.value = vbChecked Then
Dim Total As Double
Dim ttt
     Dim escFeedAndCut As String
     Dim escNewLine As String
     Dim escUnerLineON As String
     Dim escUnerLineOnX2 As String
     Dim escUnerLineOff As String
     Dim escBoldOn As String
     Dim escBoldOff As String
     Dim escNegativeOn As String
     Dim escNegativeOff As String
     Dim esc8CpiOn As String
     Dim esc8CPiOff As String
     Dim esc16Cpi As String
     Dim esc20Cpi As String
     Dim escAlignLeft As String
     Dim escAlignCenter As String
     Dim escAlignRight As String
    
     
        escNewLine = Chr(10) '//New Line (LF Line Feed)
        escUnerLineON = Chr(27) + Chr(45) + Chr(1) '//Unerline on
        escUnerLineOnX2 = Chr(27) + Chr(45) + Chr(1) '//Unerline on X2
        escUnerLineOff = Chr(27) + Chr(45) + Chr(0) '//unerline off
        escBoldOn = Chr(27) + Chr(69) + Chr(1) '//Bold on
        escBoldOff = Chr(27) + Chr(69) + Chr(0) '//Bold off
        escNegativeOn = Chr(29) + Chr(66) + Chr(1) '//White on Black on
        escNegativeOff = Chr(29) + Chr(66) + Chr(0) '//white on
        esc8CpiOn = Chr(29) + Chr(33) + Chr(16) '//Font Size X2 on
        esc8CPiOff = Chr(29) + Chr(33) + Chr(0) '//Font size X2 off
        esc16Cpi = Chr(27) + Chr(77) + Chr(48) '//Font A - Normal Size
        esc20Cpi = Chr(27) + Chr(77) + Chr(49) '//Font B - Small Font
        escAlignLeft = Chr(27) + Chr(97) + Chr(48) '//Align text to the left
        escAlignCenter = Chr(27) + Chr(97) + Chr(49) '//Align text to the center
        escAlignRight = Chr(27) + Chr(97) + Chr(50) '//Align text to the right
        escFeedAndCut = Chr(29) + Chr(86) + Chr(65) '//Partial cut and feed
       Dim fso, chkPrinter, txtfile
        'ttt = "LPT1" 'LPT1
         Dim PORT As String
        PORT = ports
        'ttt = "LPT1" 'LPT1
        ttt = PORT
        Set fso = CreateObject("Scripting.FileSystemObject")
        Dim strReceipts As String
         
        Set txtfile = fso.CreateTextFile(ttt, True)
        txtfile.WriteLine "      " & cname & ""
        txtfile.WriteLine "      AGROVET"
        txtfile.WriteLine "      " & paddress & ""
        txtfile.WriteLine "      " & town & ""
        txtfile.WriteLine "      " & Phone & ""
        txtfile.WriteLine "      " & Email & ""
        txtfile.WriteLine "---------------------------------------"
        txtfile.WriteLine "    RECEIVING VOUCHER"
        txtfile.WriteLine
        txtfile.WriteLine "    INVOICE NO:" & "-------------------"""
        txtfile.WriteLine "......................................."
        txtfile.WriteLine "Vendor" & cboVendor
        txtfile.WriteLine
        txtfile.WriteLine "---------------------------------------"
        txtfile.WriteLine "ITEM" & vbTab & vbTab & "QNTY" & vbTab & "PRICE" & vbTab & "AMOUNT"
        txtfile.WriteLine "......................................."
       
        j = 1
        strReceipts = ""
        Do While Not j > (lvwselecteditems.ListItems.Count)
            lvwselecteditems.ListItems.Item(j).selected = True
            lenght = Len(lvwselecteditems.SelectedItem.SubItems(2))
            strReceipts = Mid(lvwselecteditems.SelectedItem.SubItems(2), 5, lenght - 5)
            If Len(strReceipts) > 14 Then
            strReceipts = strReceipts & "-"
            Else
            strReceipts = strReceipts & vbTab
            End If
            strReceipts = strReceipts & CDbl(lvwselecteditems.SelectedItem.SubItems(4)) & vbTab & Format(lvwselecteditems.SelectedItem.SubItems(6), "#,##0.00") & vbTab & Format((lvwselecteditems.SelectedItem.SubItems(4) * lvwselecteditems.SelectedItem.SubItems(6)), "#,##0.00") & vbNewLine
            txtfile.WriteLine strReceipts
            j = j + 1
        Loop
      
        txtfile.WriteLine "---------------------------------------" & vbNewLine
        txtfile.WriteLine "RECEIPT TOTAL" & vbTab & vbTab & Format(LBLTOTAL, "#,##0.00") & vbNewLine
        txtfile.WriteLine "======================================="
        txtfile.WriteLine
        txtfile.WriteLine "Remarks" & vbTab & txtcomment
        txtfile.WriteLine
        txtfile.WriteLine "YOU WERE SERVED By " & UCase(username)
        txtfile.WriteLine "    Date :" & Format(Get_Server_Date, "dd/mm/yyyy HH:MM:SS AM/PM")
        txtfile.WriteLine "     THANK YOU AND WELCOME "
        txtfile.WriteLine "****************************************"
        txtfile.WriteLine escFeedAndCut
        txtfile.Close
End If
lvwselecteditems.ListItems.Clear
txtcomment = ""
MsgBox "saved successfully"
Form_Load
End Sub

Private Sub Form_Load()
chkPrint.value = vbChecked

lvwselecteditems.Enabled = True
txtDelNo = ""
txtRef = ""
txtremarks = ""
txtQnty = ""
txtcomment = ""

cboStore.Clear
cboVendor.Clear
dtprecDate = Format(Get_Server_Date, "dd/mm/yyyy")


Set rs = oSaccoMaster.GetRecordset("SELECT  CompanyName  FROM ag_Supplier1 order by companyname")
While Not rs.EOF
cboVendor.AddItem rs.Fields(0)
rs.MoveNext
Wend
cboVendor.Text = "<Select Vendor>"


    cboStore.Clear
    sql = "Select companyname from ag_Supplier1"
    Set rs = oSaccoMaster.GetRecordset(sql)
    While Not rs.EOF
    cboStore.AddItem rs.Fields(0)
    rs.MoveNext
    Wend
cboStore.Text = "<Select Store>"

Lvwitems.ListItems.Clear
Set rs = oSaccoMaster.GetRecordset("d_sp_loadOrderedGoods")
'Set rs = oSaccoMaster.GetRecordset("d_sp_IOrdered")
While Not rs.EOF
Set li = Lvwitems.ListItems.Add(, , rs.Fields(0))
                li.SubItems(1) = rs.Fields(1) & ""
                li.SubItems(2) = rs.Fields(2) & ""
                li.SubItems(3) = rs.Fields(3) & ""
                li.SubItems(4) = rs.Fields(4) & ""
                li.SubItems(5) = "0" & ""
                li.SubItems(6) = rs.Fields(4) & ""
                
rs.MoveNext
Wend

InitSubClass

'    Set objLabelEdit = New LabelEdit
'    objLabelEdit.Init Me, lvwItems
'    Set objLabelEdit2 = New LabelEdit
'    objLabelEdit2.Init Me, lvwselecteditems
'        InitSubClass
  
    'Enable label editing for listview2
    Set objLabelEdit = New LabelEdit
    objLabelEdit.Init Me, Lvwitems
    Set objLabelEdit2 = New LabelEdit
    objLabelEdit2.Init Me, Lvwitems



End Sub

Private Sub lvwItems_DblClick()
cmdAdd_Click
End Sub

Private Sub LvwselectedItems_Click()
  Dim Total As Double, amt As Double, Price As Double, qnty As Integer
    Dim ccount As Integer
    On Error Resume Next
    Total = 0
    With lvwselecteditems
        If .ListItems.Count > 0 Then
            ccount = .ListItems.Count
            For I = 1 To ccount
                With .ListItems(I)
                        Price = CDbl(.ListSubItems(6))
                        qnty = CDbl(.ListSubItems(4))
                        amt = Price * qnty
                        Total = Total + amt
                            Dim objLabelEdit2 As LabelEdit
                        Set objLabelEdit2 = New LabelEdit
                    objLabelEdit2.Init Me, .ListSubItems(4)

                End With
            Next I

        Else
            Total = 0
        End If
    End With
    LBLTOTAL.Caption = Total
    

End Sub

Private Sub lvwSelectedItems_DblClick()
Set rs = oSaccoMaster.GetRecordset("SELECT d_LPO.RefNo,d_Requisition.CostCentre,d_LPO.Vendor FROM d_LPO,d_Requisition WHERE d_Requisition.RNo=d_LPO.RefNo AND  d_LPO.PNo=" & lvwselecteditems.SelectedItem)
If Not rs.EOF Then
If Not IsNull(rs.Fields(0)) Then
txtRef = rs.Fields(0)
End If
If Not IsNull(rs.Fields(1)) Then
cboStore = rs.Fields(1)
End If
If Not IsNull(rs.Fields(2)) Then
cboVendor = rs.Fields(2)
End If

txtOrdered = lvwselecteditems.SelectedItem.ListSubItems(3)
End If
SSTab1.Tab = 0
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Tab = 0 Then
cmdsave.Enabled = True
'cmdClear.Enabled = True
Else
cmdsave.Enabled = True
'cmdClear.Enabled = False
End If
End Sub
Private Sub txtDelNo_Validate(Cancel As Boolean)
    If Trim(txtQnty) = "" Then
        txtQnty = "0"
    End If
End Sub


Private Sub txtQnty_KeyPress(KeyAscii As Integer)
If (KeyAscii = 46) Or (KeyAscii = 8) Or (KeyAscii = 48) Or (KeyAscii = 49) Or (KeyAscii = 50) Or (KeyAscii = 51) Or (KeyAscii = 52) Or (KeyAscii = 53) Or (KeyAscii = 54) Or (KeyAscii = 55) Or (KeyAscii = 56) Or (KeyAscii = 57) Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
MsgBox "Please enter a number "
End If
End Sub

