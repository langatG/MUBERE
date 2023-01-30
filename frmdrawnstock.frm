VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmdrawnstock 
   Caption         =   "Drawn stock"
   ClientHeight    =   8655
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8655
   ScaleWidth      =   12000
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView Lvwdrawn 
      Height          =   3375
      Left            =   120
      TabIndex        =   27
      Top             =   4680
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   5953
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Product code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Product name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "User name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Total Amount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Price"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Quantity"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Branch code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Drawnstock"
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   11775
      Begin VB.CommandButton cmdchange 
         Caption         =   "Change the Price or Quanity"
         BeginProperty Font 
            Name            =   "News701 BT"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4200
         TabIndex        =   41
         Top             =   3240
         Width           =   2775
      End
      Begin VB.ComboBox txtpname 
         Height          =   315
         Left            =   1680
         TabIndex        =   40
         Top             =   1200
         Width           =   3375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "News701 BT"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6600
         TabIndex        =   39
         Top             =   3840
         Width           =   735
      End
      Begin VB.CommandButton cmddelete 
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "News706 BT"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5520
         TabIndex        =   38
         Top             =   3840
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "TO A.I SERVICE PROVIDERS"
         Height          =   615
         Left            =   4080
         TabIndex        =   35
         Top             =   2640
         Visible         =   0   'False
         Width           =   2655
         Begin VB.CheckBox chkbal 
            Caption         =   "Return"
            Height          =   315
            Left            =   1320
            TabIndex        =   37
            Top             =   240
            Width           =   1095
         End
         Begin VB.CheckBox chkservice 
            Caption         =   "Yes"
            Height          =   315
            Left            =   480
            TabIndex        =   36
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.ComboBox cbobranchf 
         Height          =   315
         ItemData        =   "frmdrawnstock.frx":0000
         Left            =   1680
         List            =   "frmdrawnstock.frx":0002
         TabIndex        =   33
         Top             =   3480
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker DTPdateentered 
         Height          =   375
         Left            =   8640
         TabIndex        =   31
         Top             =   480
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         Format          =   130416641
         CurrentDate     =   42648
      End
      Begin VB.ComboBox Cbobrancht 
         Height          =   315
         ItemData        =   "frmdrawnstock.frx":0004
         Left            =   1680
         List            =   "frmdrawnstock.frx":0006
         TabIndex        =   30
         Top             =   3960
         Width           =   2175
      End
      Begin VB.CommandButton cmdsave 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "News701 BT"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4200
         TabIndex        =   28
         Top             =   3840
         Width           =   975
      End
      Begin VB.TextBox txtpcode 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1680
         TabIndex        =   17
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox txtdescription 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1680
         MultiLine       =   -1  'True
         TabIndex        =   16
         Top             =   1800
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.TextBox txtquantity 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   8640
         TabIndex        =   15
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox txtbalance 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Left            =   8640
         TabIndex        =   14
         Top             =   1440
         Width           =   1815
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         DrawStyle       =   2  'Dot
         DrawWidth       =   17015
         Height          =   360
         Left            =   3600
         Picture         =   "frmdrawnstock.frx":0008
         ScaleHeight     =   360
         ScaleWidth      =   240
         TabIndex        =   13
         Top             =   600
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Frame fra1 
         BorderStyle     =   0  'None
         Height          =   1695
         Left            =   7320
         TabIndex        =   4
         Top             =   2520
         Width           =   4335
         Begin VB.PictureBox Picture1 
            Height          =   255
            Left            =   1320
            Picture         =   "frmdrawnstock.frx":018A
            ScaleHeight     =   195
            ScaleWidth      =   195
            TabIndex        =   8
            Top             =   840
            Width           =   255
         End
         Begin VB.PictureBox Picture3 
            Height          =   255
            Left            =   1320
            Picture         =   "frmdrawnstock.frx":0A54
            ScaleHeight     =   195
            ScaleWidth      =   195
            TabIndex        =   7
            Top             =   1440
            Width           =   255
         End
         Begin VB.TextBox txtdracc 
            Height          =   375
            Left            =   1680
            TabIndex        =   6
            Top             =   720
            Width           =   2535
         End
         Begin VB.TextBox txtcracc 
            Height          =   375
            Left            =   1680
            TabIndex        =   5
            Top             =   1320
            Width           =   2535
         End
         Begin VB.Label lbldracc 
            BackColor       =   &H8000000E&
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label lblcracc 
            BackColor       =   &H8000000E&
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Label Label13 
            Caption         =   "DrAccNo"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label Label8 
            Caption         =   "Craccno"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   1200
            Width           =   1095
         End
      End
      Begin VB.TextBox txtpprice 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1680
         TabIndex        =   3
         Top             =   2400
         Width           =   1455
      End
      Begin VB.TextBox txtsellingprice 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1680
         TabIndex        =   2
         Top             =   3000
         Width           =   1455
      End
      Begin VB.TextBox txtprice 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   8640
         TabIndex        =   1
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label lblstatus 
         BackColor       =   &H00C0FFFF&
         Height          =   255
         Left            =   10800
         TabIndex        =   34
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label12 
         Caption         =   "Branch From"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Branch To"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   3960
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Product Code"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Product Name"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Description"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1800
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Quantity To Branch"
         Height          =   375
         Left            =   6960
         TabIndex        =   23
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Balance In Store"
         Height          =   255
         Left            =   7320
         TabIndex        =   22
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Date Entered"
         Height          =   255
         Left            =   7320
         TabIndex        =   21
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Purchase Price "
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Selling Price "
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "Price per goods"
         Height          =   255
         Left            =   7320
         TabIndex        =   18
         Top             =   2160
         Width           =   1215
      End
   End
   Begin VB.Menu mnudrwn 
      Caption         =   "Report"
      Begin VB.Menu mnudrawn 
         Caption         =   "Drawn stock"
      End
   End
End
Attribute VB_Name = "frmdrawnstock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub txtpcode11_Change()
'//TWNG001
Provider = "MAZIWA"
Set cn = New ADODB.Connection
cn.Open Provider, "atm", "atm"
'If reportpath = "" Then reportpath = GetSetting("payroll", "AppName", "rptPath", rptPath)
sql = "select P_CODE,P_NAME,S_NO,QOUT,supplierid,pprice, sprice from ag_products where p_code='" & txtpcode & "'AND Branch='" & cbobranchf & "'"
Set rs = New ADODB.Recordset
rs.Open sql, cn
If Not rs.EOF Then
 txtpcode = (rs.Fields(0))
If Not IsNull(rs.Fields(1)) Then txtpname = (rs.Fields(1))
If Not IsNull(rs.Fields(3)) Then txtbalance = (rs.Fields(3))
If Not IsNull(rs.Fields(4)) Then cbosupplier = (rs.Fields(4))
If Not IsNull(rs.Fields(5)) Then txtpprice = (rs.Fields(5))
If Not IsNull(rs.Fields(6)) Then txtsellingprice = (rs.Fields(6))
If txtbalance <= 0 Then
MsgBox "Warning:Your stock is below zero please reorder", vbInformation
Else

End If
End If
'LstSearch.Refresh
 Lvwdrawn.ListItems.Clear
 'If chkshowallmembers = vbChecked Then
 sql = ""
 sql = "SET dateformat dmy SELECT DATE, DESCRIPTION, QUANTITY, TOTALAMOUNT, PRODUCTID, PRODUCTNAME, USERNAME, PRICEEACH, MONTH, YEAR, Branch, updated From DRAWNSTOCK where (DATE >= '" & Date & "')   order by productid"
 Set rst = oSaccoMaster.GetRecordset(sql)
    'End If
    If rst.RecordCount > 0 Then
    With rst
        If Not .EOF Then
            While Not .EOF
                Set li = Lvwdrawn.ListItems.Add(, , !Date)
               ' li.SubItems(1) = IIf(IsNull(!productid), "", !productid)
                li.SubItems(1) = IIf(IsNull(!description), "", !description)
                li.SubItems(2) = IIf(IsNull(!PRODUCTID), "", !PRODUCTID)
                li.SubItems(3) = IIf(IsNull(!ProductName), "", !ProductName)
                li.SubItems(4) = IIf(IsNull(!username), "", !username)
                li.SubItems(5) = IIf(IsNull(!Totalamount), "", !Totalamount)
                li.SubItems(6) = IIf(IsNull(!PRICEEACH), "", !PRICEEACH)
                li.SubItems(7) = IIf(IsNull(!Quantity), "", !Quantity)
                li.SubItems(8) = IIf(IsNull(!Branch), "", !Branch)
                .MoveNext
            Wend
        End If
    End With
    TxtRecords = rst.RecordCount
    End If
'    cboSearchField.Text = cboSearchField.List(0)
'    cboCriteria.Text = cboCriteria.List(3)
'cmdMemberSearch_Click


'// check with serial no if it exist
End Sub
Private Sub Load_Lvwdrawn()
'LstSearch.Refresh
'Lvwdrawn.ListItems.Clear
 Lvwdrawn.ListItems.Clear
 'If chkshowallmembers = vbChecked Then
      ' Set rst = oSaccoMaster.GetRecordset("SELECT     DATE, DESCRIPTION, QUANTITY, TOTALAMOUNT, PRODUCTID, PRODUCTNAME, USERNAME, PRICEEACH, MONTH, YEAR, Branch, updated From DRAWNSTOCK where date='" & Date & "'  ")
    'End If
 sql = ""
 sql = "SET dateformat dmy SELECT DATE, DESCRIPTION, QUANTITY, TOTALAMOUNT, PRODUCTID, PRODUCTNAME, USERNAME, PRICEEACH, MONTH, YEAR, Branch, updated From DRAWNSTOCK where (DATE >= '" & Date & "')   order by productid"
 Set rst = oSaccoMaster.GetRecordset(sql)
    'End If
    'If rst.RecordCount > 0 Then
    If rst.RecordCount > 0 Then
    With rst
         'Lvwdrawn.ListItems.Clear
        If Not .EOF Then
        'If .EOF Then
        
            While Not .EOF
                Set li = Lvwdrawn.ListItems.Add(, "", !Date)
               ' li.SubItems(1) = IIf(IsNull(!productid), "", !productid)
                li.SubItems(1) = IIf(IsNull(!description), "", !description)
                li.SubItems(2) = IIf(IsNull(!PRODUCTID), "", !PRODUCTID)
                li.SubItems(3) = IIf(IsNull(!ProductName), "", !ProductName)
                li.SubItems(4) = IIf(IsNull(!username), "", !username)
                li.SubItems(5) = IIf(IsNull(!Totalamount), "", !Totalamount)
                li.SubItems(6) = IIf(IsNull(!PRICEEACH), "", !PRICEEACH)
                li.SubItems(7) = IIf(IsNull(!Quantity), "", !Quantity)
                li.SubItems(8) = IIf(IsNull(!Branch), "", !Branch)
                .MoveNext
            Wend
        End If
    End With
    TxtRecords = rst.RecordCount
    End If
End Sub

Private Sub cbobranchf_Click()
    sql = ""
    sql = "select P_NAME from ag_products where Branch='" & cbobranchf & "'"
    'rst.Open sql, cn, adOpenKeyset, adLockOptimistic
    Set rst = oSaccoMaster.GetRecordset(sql)
    'Set rst = ExecuteThis(sql)
    While Not rst.EOF
    txtpname.AddItem rst.Fields(0)
    'Cbobrancht.AddItem rst.Fields(0)
    rst.MoveNext
    Wend
End Sub

Private Sub Cbobrancht_Change()
If chkservice = vbChecked Then
'Cbobrancht = "A.I SERVICE PROVIDERS"
Cbobrancht.Locked = True
Else
Cbobrancht.Locked = False
End If
End Sub

Private Sub chkservice_Click()
'Dim serai As String
If chkservice = vbChecked Then
'Cbobrancht = "A.I SERVICE PROVIDERS"
Cbobrancht.Locked = True
Cbobrancht = "OLENGURUONE"
Else
Cbobrancht.Locked = False
End If
End Sub

Private Sub cmdChange_Click()
FRMCHANGE.Show vbModal
End Sub

Private Sub cmddelete_Click()
On Error GoTo HEREEE
    sql = "SELECT     UserLoginIDs, UserGroup, SUPERUSER From UserAccounts where UserLoginIDs='" & User & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
    If rs!SuperUser <> "1" And rs!SuperUser <> "2" Then
    MsgBox "You are not allowed to Draw stock", vbInformation
    Exit Sub
    End If
    End If
    
If txtquantity = "" Then
MsgBox "Sorry Quanty should not be Zero please re-enter before your proceed", vbInformation
Exit Sub
End If

Set cn = New ADODB.Connection
cn.Open Provider, "atm", "atm"
sql = ""
'If reportpath = "" Then reportpath = GetSetting("payroll", "AppName", "rptPath", rptPath)
sql = "select P_CODE,qin,qout,o_bal,unserialized from ag_products where p_code='" & txtpcode & "' and Branch='" & Cbobrancht & "' "
Set rs = New ADODB.Recordset
rs.Open sql, cn
If rs.EOF Then
'If rs!qout <= 0 Then
MsgBox "Sorry Quanty should not be Zero please re-enter before your proceed", vbInformation
'Exit Sub
'End If
'// insert into ag_products
'If txtserialno = "" Then txtserialno = 0
'sql = ""
'sql = "set dateformat dmy delete from ag_products where p_code='" & txtpcode.Text & "' and p_name='" & txtpname.Text & "' and Qin=" & txtquantity.Text & " and Date_Entered='" & DTPdateentered & "' and Branch='" & Cbobrancht & "')"
'cn.Execute sql
'sql = "set dateformat DMY update ag_products set p_name='" & txtpname & "',qin=" & txtbalance + txtquantity.Text & ", qout= " & txtbalance + txtquantity.Text & ",o_bal=" & txtbalance + txtquantity.Text & "  ,last_d_updated='" & Date & "',user_id='" & User & "',audit_date='" & Date & "',unserialized=" & unsera - D & ",SERIA=" & unsera & ",pprice=" & txtpprice & ",sprice=" & txtsellingprice & " where p_code='" & txtpcode.Text & "' and Branch='" & cbobranchf & "' "
'cn.Execute sql
Else
sql = ""
sql = "set dateformat DMY update ag_products set p_name='" & txtpname & "',qin=" & txtbalance - (-1 * txtquantity.Text) & ", qout= " & txtbalance - (-1 * txtquantity.Text) & ",o_bal=" & txtbalance - (-1 * txtquantity.Text) & ",last_d_updated='" & Date & "',user_id='" & User & "',audit_date='" & Date & "',unserialized=" & unsera - D & ",SERIA='" & unsera & "',pprice=" & txtpprice & ",sprice=" & txtsellingprice & " where p_code='" & txtpcode.Text & "' and Branch='" & cbobranchf & "' "
cn.Execute sql
sql = ""
sql = "set dateformat DMY update ag_products set p_name='" & txtpname & "',Qin= " & rs.Fields("qin") - txtquantity.Text & ", qout= " & rs.Fields("qout") - txtquantity.Text & ",o_bal=" & rs.Fields("o_bal") - txtquantity.Text & " ,last_d_updated='" & Date & "',user_id='" & User & "',audit_date='" & Date & "',unserialized=" & unsera - D & ",SERIA= '" & unsera & "',pprice=" & txtpprice & ",sprice=" & txtsellingprice & " where p_code='" & txtpcode.Text & "' and Branch='" & Cbobrancht & "' "
cn.Execute sql

'sql = ""
'sql = "set dateformat DMY INSERT INTO ag_stockbalance"
'sql = sql & " (p_code, productname, openningstock, changeinstock, stockbalance, transdate,companyid,pprice,sprice,RLevel)"
'sql = sql & " VALUES     ('" & txtpcode.Text & "','" & txtpname & "', " & txtbalance & ", " & txtquantity & ", " & txtbalance.Text + txtquantity.Text & ", '" & txtdateenterered & "',1," & txtpprice & "," & txtsellingprice & "," & txtRLevel & ")"
'cn.Execute sql




'Dim D As Double
'If Not IsNull(rs.Fields(2)) Then D = rs.Fields(2)
'sql = ""
'sql = "set dateformat DMY update ag_products set p_name='" & txtpname & "',qin=" & txtquantity.Text & ", qout= " & rs.Fields("qout") & ",o_bal=" & rs.Fields("qout") & ",last_d_updated='" & Date & "',user_id='" & User & "',audit_date='" & Date & "',unserialized=" & unsera - D & ",SERIA=" & unsera & ",pprice=" & txtpprice & ",sprice=" & txtsellingprice & " where p_code='" & txtpcode.Text & "'"
'cn.Execute sql
End If
Dim rsst As Recordset
sql = ""
sql = "set dateformat DMY select top 1 * from ag_stockbalance where p_code='" & txtpcode & "' and Branch='" & Cbobrancht & "'order by trackid desc "
Set rsst = New ADODB.Recordset
rsst.Open sql, cn
If rsst.EOF Then
sql = ""
sql = "set dateformat DMY delete from ag_stockbalance where p_code='" & txtpcode & "' and productname='" & txtpname & "' and openningstock='" & txtquantity & "' and transdate='" & DTPdateentered & "' and Branch='" & Cbobrancht & "'and AI=" & lblStatus & ")"
cn.Execute sql
sql = ""
sql = "SET DATEFORMAT DMY Update ag_stockbalance"
sql = sql & " SET  productname = '" & txtpname & "', openningstock = " & rsst.Fields("openningstock") + txtquantity & ", changeinstock = " & rsst.Fields("changeinstock") + txtquantity & ", stockbalance = " & rsst.Fields("stockbalance") + txtquantity & ", transdate = '" & DTPdateentered & "' "
sql = sql & " WHERE     (p_code = '" & txtpcode & "' and Branch='" & cbobranchf & "')"
cn.Execute sql
Else
sql = "SET DATEFORMAT DMY Update ag_stockbalance"
sql = sql & " SET              productname = '" & txtpname & "', openningstock = " & txtbalance - (-1 * txtquantity) & ", changeinstock = " & txtbalance - (-1 * txtquantity) & ", stockbalance = " & txtbalance - (-1 * txtquantity) & ", transdate = '" & DTPdateentered & "' "
sql = sql & " WHERE     (p_code = '" & txtpcode & "')  and Branch='" & cbobranchf & "'"
cn.Execute sql

sql = "SET DATEFORMAT DMY Update ag_stockbalance"
sql = sql & " SET              productname = '" & txtpname & "', changeinstock = " & rsst.Fields("changeinstock") - txtquantity & ", stockbalance = " & rsst.Fields("stockbalance") - txtquantity & ", transdate = '" & DTPdateentered & "' "
sql = sql & " WHERE     (p_code = '" & txtpcode & "')  and Branch='" & Cbobrancht & "'"
cn.Execute sql

'// update serialno database

'' ///update gl


End If
If seria = 1 Then
Set rst = Nothing
    sql = ""
   sql = "select * from serialno where serialno='" & txtSERIALNO & "' AND P_CODE='" & txtpcode & "' and used=0"
   Set rst = New ADODB.Recordset
   rst.Open sql, cn, adOpenKeyset, adLockOptimistic

If rst.EOF Then
sql = ""
sql = "set dateformat DMY delete from serialno(serialno,p_code,used)"
sql = sql & " values('" & txtSERIALNO & "','" & txtpcode & "',0)"
cn.Execute sql
Else
MsgBox "Item is in place and not yet used", vbInformation
Exit Sub
End If
End If
If txtPrice = "" Then
txtPrice = txtpprice
End If
sql = ""
sql = "Set dateformat dmy delete from DRAWNSTOCK where DATE>='" & Date & "' and QUANTITY='" & txtquantity & "' and PRODUCTID=" & txtpcode & " and PRODUCTNAME='" & txtpname & "' and Branch='" & Cbobrancht & "' "
cn.Execute sql
'Load_Lvwdrawn
sql = ""
sql = "set dateformat dmy delete from  ag_products3 where p_code='" & txtpcode.Text & "' and audit_date ='" & Date & "' and branch='" & Cbobrancht & "' "
cn.Execute sql
sql = ""
sql = "set dateformat dmy delete from  ag_products3 where p_code='" & txtpcode.Text & "' and qin='" & -1 * txtquantity & "'  and audit_date ='" & Date & "' "
cn.Execute sql
'sql = ""
'sql = "set dateformat dmy delete from  ag_products4 where p_code='" & txtpcode.Text & "' and audit_date ='" & Date & "' and branch='" & cbobranchf & "' "
'insert into  ag_products3(p_code,p_name,s_no,qin,qout,date_entered,last_d_updated,user_id,audit_date,o_bal,supplierid,serialized,unserialized,seria,pprice,sprice,branch,AI ) values('" & txtpcode.Text & "','" & txtpname.Text & "',''," & txtquantity.Text & "*-1," & txtquantity.Text & "*-1,'" & DTPdateentered & "','" & DTPdateentered & "','" & User & "','" & Date & "'," & txtquantity.Text & "*-1,'',0,'',''," & txtpprice & "," & txtsellingprice & ",'" & cbobranchf & "','" & lblstatus & "')"
'cn.Execute sql
'sql = ""
'sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & DTPdateentered & "'," & txtquantity & " *" & txtpprice & ",'" & lbldracc & "','" & lblcracc & "','stock intake','" & cbosupplier & "' ,'stock intake','" & User & "',0,0)"
'oSaccoMaster.ExecuteThis (sql)
Load_Lvwdrawn
txtbalance = ""
txtpcode = ""
txtpname = ""
txtSERIALNO = ""
txtquantity = ""
txtpprice = ""
txtsellingprice = ""
cbosupplier = ""
txtdescription = ""
cbobranch = ""
txtPrice = ""
chkservice = vbUnchecked
chkbal = vbUnchecked
MsgBox "Record Deleted Successfully"
Exit Sub
HEREEE:
MsgBox err.description & " error occured."
End Sub

Private Sub cmdsave_Click()
On Error GoTo HEREEE
Set rst = New Recordset
'If lbldracc = "" Then MsgBox "select the account to Debit": Exit Sub
'
'If lblcracc = "" Then MsgBox "select the account to credit": Exit Sub


'
Dim unsera As Integer
'Dim cn As Connection
If Trim(txtquantity) = "" Then
MsgBox "Quantity cannot be Zero", vbInformation
Exit Sub

End If
If Trim(txtbalance) = "" Then txtbalance = 0
'If chkserialrequired = vbChecked Then
'
'seria = 1
unsera = txtquantity

'// should only be one item
'If txtquantity > 1 Then
'MsgBox "Serialized items should only be added as one", vbCritical
'Exit Sub
'End If
'Else
'seria = 0
'unsera = 0
'End If



Provider = cn
sql = ""
sql = "select P_CODE,qin,qout from ag_products where p_code='" & txtpcode & "' and branch='" & cbobranchf & "'"
Set rsinstock = New ADODB.Recordset
rsinstock.Open sql, cn
'// check the stock if it is less than zero
If rsinstock.Fields(2) <= 0 Then
MsgBox "Sorry Stock is Zero please re-stock before your proceed", vbInformation
Exit Sub
End If
If chkbal = vbChecked Then
If chkservice = vbUnchecked Then
MsgBox "Please mark Yes to continue", vbInformation
Exit Sub
End If
txtquantity = txtquantity * -1
Else
txtquantity = txtquantity
End If

'If chkservice = vbChecked Then
'Cbobrancht = "OLENGURUONE"
'Else
'Cbobrancht
'End If

Set cn = New ADODB.Connection
cn.Open Provider, "atm", "atm"
sql = ""
'If reportpath = "" Then reportpath = GetSetting("payroll", "AppName", "rptPath", rptPath)
sql = "select P_CODE,qin,qout,o_bal,unserialized from ag_products where p_code='" & txtpcode & "' and Branch='" & Cbobrancht & "' "
Set rs = New ADODB.Recordset
rs.Open sql, cn
If rs.EOF Then
'If rs!qout <= 0 Then
'MsgBox "Sorry Stock is Zero please re-stock before your proceed", vbInformation
'Exit Sub
'End If
'// insert into ag_products
'If txtserialno = "" Then txtserialno = 0
sql = ""
sql = "set dateformat dmy insert into  ag_products(p_code,p_name,s_no,qin,qout,date_entered,last_d_updated,user_id,audit_date,o_bal,supplierid,serialized,unserialized,seria,pprice,sprice,branch,AI ) values('" & txtpcode.Text & "','" & txtpname.Text & "','0'," & txtquantity.Text & "," & txtquantity.Text & ",'" & DTPdateentered & "','" & DTPdateentered & "','" & User & "','" & Date & "'," & txtquantity.Text & ",'',0,'',''," & txtpprice & "," & txtsellingprice & ",'" & Cbobrancht & "','" & lblStatus & "')"
cn.Execute sql

'sql = ""


sql = "set dateformat DMY update ag_products set p_name='" & txtpname & "',qin=" & txtbalance - txtquantity.Text & ", qout= " & txtbalance - txtquantity.Text & ",o_bal=" & txtbalance - txtquantity.Text & "  ,last_d_updated='" & Date & "',user_id='" & User & "',audit_date='" & Date & "',unserialized=" & unsera - D & ",SERIA=" & unsera & ",pprice=" & txtpprice & ",sprice=" & txtsellingprice & " where p_code='" & txtpcode.Text & "' and Branch='" & cbobranchf & "' "
cn.Execute sql
Else
sql = ""
sql = "set dateformat DMY update ag_products set p_name='" & txtpname & "',qin=" & rs.Fields("qin") + txtquantity.Text & ", qout= " & rs.Fields("qout") + txtquantity.Text & ",o_bal=" & rs.Fields("o_bal") + txtquantity.Text & "  ,last_d_updated='" & Date & "',user_id='" & User & "',audit_date='" & Date & "',unserialized=" & unsera - D & ",SERIA=" & unsera & ",pprice=" & txtpprice & ",sprice=" & txtsellingprice & " where p_code='" & txtpcode.Text & "' and Branch='" & Cbobrancht & "' "
cn.Execute sql
sql = ""
sql = "set dateformat DMY update ag_products set p_name='" & txtpname & "',qin=" & txtbalance - txtquantity.Text & ", qout= " & txtbalance - txtquantity.Text & ",o_bal=" & txtbalance - txtquantity.Text & " ,last_d_updated='" & Date & "',user_id='" & User & "',audit_date='" & Date & "',unserialized=" & unsera - D & ",SERIA=" & unsera & ",pprice=" & txtpprice & " where p_code='" & txtpcode.Text & "' and Branch='" & cbobranchf & "' "
cn.Execute sql

'sql = ""
'sql = "set dateformat DMY INSERT INTO ag_stockbalance"
'sql = sql & " (p_code, productname, openningstock, changeinstock, stockbalance, transdate,companyid,pprice,sprice,RLevel)"
'sql = sql & " VALUES     ('" & txtpcode.Text & "','" & txtpname & "', " & txtbalance & ", " & txtquantity & ", " & txtbalance.Text + txtquantity.Text & ", '" & txtdateenterered & "',1," & txtpprice & "," & txtsellingprice & "," & txtRLevel & ")"
'cn.Execute sql

'Dim D As Double
'If Not IsNull(rs.Fields(2)) Then D = rs.Fields(2)
'sql = ""
'sql = "set dateformat DMY update ag_products set p_name='" & txtpname & "',qin=" & txtquantity.Text & ", qout= " & rs.Fields("qout") & ",o_bal=" & rs.Fields("qout") & ",last_d_updated='" & Date & "',user_id='" & User & "',audit_date='" & Date & "',unserialized=" & unsera - D & ",SERIA=" & unsera & ",pprice=" & txtpprice & ",sprice=" & txtsellingprice & " where p_code='" & txtpcode.Text & "'"
'cn.Execute sql
End If
Dim rsst As Recordset
sql = ""
sql = "set dateformat DMY select top 1 * from ag_stockbalance where p_code='" & txtpcode & "' and Branch='" & Cbobrancht & "'order by trackid desc "
Set rsst = New ADODB.Recordset
rsst.Open sql, cn
If rsst.EOF Then
sql = ""
sql = "set dateformat DMY INSERT INTO ag_stockbalance"
sql = sql & " (p_code, productname, openningstock, changeinstock, stockbalance, transdate,companyid,Branch,AI)"
sql = sql & " VALUES     ('" & txtpcode & "', '" & txtpname & "', '0', '" & txtquantity & "', '" & txtquantity.Text & "', '" & DTPdateentered & "',1,'" & Cbobrancht & "','" & lblStatus & "')"
cn.Execute sql
sql = ""
sql = "SET DATEFORMAT DMY Update ag_stockbalance"
sql = sql & " SET  productname = '" & txtpname & "', openningstock = " & txtbalance - txtquantity & ", changeinstock = " & txtbalance - txtquantity & ", stockbalance = " & txtbalance - txtquantity & ", transdate = '" & DTPdateentered & "' "
sql = sql & " WHERE     (p_code = '" & txtpcode & "' and Branch='" & cbobranchf & "')"
cn.Execute sql
Else
sql = "SET DATEFORMAT DMY Update ag_stockbalance"
sql = sql & " SET              productname = '" & txtpname & "', openningstock = " & rsst.Fields("openningstock") + txtquantity & ", changeinstock = " & rsst.Fields("changeinstock") + txtquantity & ", stockbalance = " & rsst.Fields("stockbalance") + txtquantity & ", transdate = '" & DTPdateentered & "' "
sql = sql & " WHERE     (p_code = '" & txtpcode & "')  and Branch='" & Cbobrancht & "'"
cn.Execute sql

sql = "SET DATEFORMAT DMY Update ag_stockbalance"
sql = sql & " SET              productname = '" & txtpname & "', openningstock = " & txtbalance - txtquantity & ", changeinstock = " & txtbalance - txtquantity & ", stockbalance = " & txtbalance - txtquantity & ", transdate = '" & DTPdateentered & "' "
sql = sql & " WHERE     (p_code = '" & txtpcode & "')  and Branch='" & cbobranchf & "'"
cn.Execute sql

'// update serialno database

'' ///update gl


End If
If seria = 1 Then
Set rst = Nothing
    sql = ""
   sql = "select * from serialno where serialno='" & txtSERIALNO & "' AND P_CODE='" & txtpcode & "' and used=0"
   Set rst = New ADODB.Recordset
   rst.Open sql, cn, adOpenKeyset, adLockOptimistic

If rst.EOF Then
sql = ""
sql = "set dateformat DMY INSERT INTO serialno(serialno,p_code,used)"
sql = sql & " values('" & txtSERIALNO & "','" & txtpcode & "',0)"
cn.Execute sql
Else
MsgBox "Item is in place and not yet used", vbInformation
Exit Sub
End If
End If
If txtPrice = "" Then
txtPrice = txtpprice
End If
sql = ""
sql = "Set dateformat dmy INSERT INTO DRAWNSTOCK(DATE, DESCRIPTION, QUANTITY, TOTALAMOUNT, PRODUCTID, PRODUCTNAME, USERNAME, PRICEEACH, MONTH, YEAR, Branch,AI,Buying) "
sql = sql & "VALUES     ('" & DTPdateentered & "','" & txtpname & "', '" & txtquantity & "', '" & txtPrice * txtquantity & "', " & txtpcode & ", '" & txtpname & "', '" & User & "', '" & txtsellingprice & "', '" & month(DTPdateentered) & "','" & year(DTPdateentered) & "','" & Cbobrancht & "','" & lblStatus & "', '" & txtpprice & "')"
cn.Execute sql
'Load_Lvwdrawn
sql = ""
sql = "set dateformat dmy insert into  ag_products3(p_code,p_name,s_no,qin,qout,date_entered,last_d_updated,user_id,audit_date,o_bal,supplierid,serialized,unserialized,seria,pprice,sprice,branch,AI ) values('" & txtpcode.Text & "','" & txtpname.Text & "',''," & txtquantity.Text & "," & txtquantity.Text & ",'" & DTPdateentered & "','" & DTPdateentered & "','" & User & "','" & Date & "'," & txtquantity.Text & ",'',0,'',''," & txtpprice & "," & txtsellingprice & ",'" & Cbobrancht & "','" & lblStatus & "')"
cn.Execute sql
sql = ""
sql = "set dateformat dmy insert into  ag_products3(p_code,p_name,s_no,qin,qout,date_entered,last_d_updated,user_id,audit_date,o_bal,supplierid,serialized,unserialized,seria,pprice,sprice,branch,AI ) values('" & txtpcode.Text & "','" & txtpname.Text & "',''," & txtquantity.Text & "*-1," & txtquantity.Text & "*-1,'" & DTPdateentered & "','" & DTPdateentered & "','" & User & "','" & Date & "'," & txtquantity.Text & "*-1,'',0,'',''," & txtpprice & "," & txtsellingprice & ",'" & cbobranchf & "','" & lblStatus & "')"
cn.Execute sql
'sql = ""
'sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & DTPdateentered & "'," & txtquantity & " *" & txtpprice & ",'" & lbldracc & "','" & lblcracc & "','stock intake','" & cbosupplier & "' ,'stock intake','" & User & "',0,0)"
'oSaccoMaster.ExecuteThis (sql)
Load_Lvwdrawn
txtbalance = ""
txtpcode = ""
txtpname = ""
txtSERIALNO = ""
txtquantity = ""
txtpprice = ""
txtsellingprice = ""
cbosupplier = ""
txtdescription = ""
cbobranch = ""
txtPrice = ""
chkservice = vbUnchecked
chkbal = vbUnchecked
MsgBox "Record Saved Successfully"
'Lvwdrawn
Exit Sub
HEREEE:
MsgBox err.description & " error occured."
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
DTPdateentered = Get_Server_Date
    Set rst = New Recordset
   ' Dim cn As Connection
    Set cn = New ADODB.Connection
    Provider = "MAZIWA"
    'Provider = cnn
    cn.Open Provider, "atm", "atm"
    Set rst = New Recordset
    
    sql = "Select Bname from   d_Branch"
    rst.Open sql, cn, adOpenKeyset, adLockOptimistic
    While Not rst.EOF
    cbobranchf.AddItem rst.Fields(0)
    Cbobrancht.AddItem rst.Fields(0)
    rst.MoveNext
    Wend
    
'    sql = "Select Bname from   d_Branch"
'    rst.Open sql, cn, adOpenKeyset, adLockOptimistic
'    While Not rst.EOF
'    Cbobrancht.AddItem rst.Fields(0)
'    rst.MoveNext
'    Wend


Load_Lvwdrawn
End Sub

Private Sub lblcracc_Change()
Set rst = oSaccoMaster.GetRecordset("select glaccname from glsetup where accno='" & lblcracc & "'")
    If Not rst.EOF Then
    txtcracc = rst.Fields("glaccname")
    End If
End Sub

Private Sub lblcracc_Click()
Set rst = oSaccoMaster.GetRecordset("select glaccname from glsetup where accno='" & lblcracc & "'")
    If Not rst.EOF Then
    txtcracc = rst.Fields("glaccname")
    End If
End Sub

Private Sub lbldracc_Change()
Set rst = oSaccoMaster.GetRecordset("select glaccname from glsetup where accno='" & lbldracc & "'")
    If Not rst.EOF Then
    txtdracc = rst.Fields("glaccname")
    End If
End Sub

Private Sub lbldracc_Click()
Set rst = oSaccoMaster.GetRecordset("select glaccname from glsetup where accno='" & lbldracc & "'")
    If Not rst.EOF Then
    txtdracc = rst.Fields("glaccname")
    End If
End Sub

Private Sub mnudrawn_Click()
reportname = "Drawnstockb.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub Picture1_Click()
Me.MousePointer = vbHourglass
        frmSearchGLAccounts.Show vbModal, Me
    If Continue Then
        If SearchValue <> "" Then
            lbldracc = SearchValue
            SearchValue = ""
        End If
    End If
    Me.MousePointer = 0
End Sub

Private Sub Picture2_Click()

If cbobranchf = "" Then
MsgBox "Please select the branch from", vbInformation
Exit Sub
End If
frmSearch.Show vbModal
Dim Y As String
Y = sel

If Y <> "" Then

Provider = "MAZIWA"

Set cn = New ADODB.Connection
cn.Open Provider, "atm", "atm"
'If reportpath = "" Then reportpath = GetSetting("payroll", "AppName", "rptPath", rptPath)
sql = "select P_CODE,P_NAME,S_NO,QOUT,supplierID,pprice,sprice,AI from ag_products where p_code='" & Y & "' AND Branch='" & cbobranchf & "'"
Set rs = New ADODB.Recordset
rs.Open sql, cn
If Not rs.EOF Then
If Not IsNull(rs.Fields(0)) Then txtpcode = (rs.Fields(0))
If Not IsNull(rs.Fields(1)) Then txtpname = (rs.Fields(1))
If Not IsNull(rs.Fields(4)) Then cbosupplier = (rs.Fields(4))
If Not IsNull(rs.Fields(5)) Then txtpprice = (rs.Fields(5))
If Not IsNull(rs.Fields(6)) Then txtsellingprice = (rs.Fields(6))
If Not IsNull(rs.Fields(5)) Then txtPrice = (rs.Fields(5))
If Not IsNull(rs.Fields(3)) Then txtbalance = (rs.Fields(3))
If Not IsNull(rs.Fields(7)) Then lblStatus = (rs.Fields(7))
'txtPrice
If txtbalance <= 0 Then
MsgBox "Your stock is below zero please reorder", vbInformation
End If
'// check with serial no if it exist


End If
End If
End Sub

Private Sub Picture3_Click()
Me.MousePointer = vbHourglass
        frmSearchGLAccounts.Show vbModal, Me
    If Continue Then
        If SearchValue <> "" Then
            lblcracc = SearchValue
            SearchValue = ""
        End If
    End If
    Me.MousePointer = 0
End Sub

Private Sub txtbal_Click()

End Sub

Private Sub txtpcode_KeyPress(KeyAscii As Integer)

If cbobranchf = "" Then
MsgBox "Please select the branch from", vbInformation
Exit Sub
End If
If KeyAscii = 13 Then
txtpcode11_Change
Else
Exit Sub
End If
End Sub

Private Sub txtserialno_Change()

End Sub

Private Sub txtpname_Change()
If cbobranchf = "" Then
MsgBox "Please select the branch From", vbInformation
Exit Sub
End If
If Cbobrancht = "" Then
MsgBox "Please select the branch To", vbInformation
Exit Sub
End If


Set rst = oSaccoMaster.GetRecordset("select p_code from ag_products where p_name ='" & txtpname & "' and branch='" & cbobranchf & "'")
If Not rst.EOF Then
txtpcode = rst.Fields("p_code")
'txtbuyingprice = rst.Fields("pprice")
'txtsellingprice = rst.Fields("sprice")
'txtserai = rst.Fields("AI")
'txtsel
txtpcode11_Change
End If
End Sub
Private Sub txtpname_Click()
If cbobranchf = "" Then
MsgBox "Please select the branch From", vbInformation
Exit Sub
End If
If Cbobrancht = "" Then
MsgBox "Please select the branch To", vbInformation
Exit Sub
End If


Set rst = oSaccoMaster.GetRecordset("select p_code from ag_products where p_name ='" & txtpname & "' and branch='" & cbobranchf & "'")
If Not rst.EOF Then
txtpcode = rst.Fields("p_code")
'txtbuyingprice = rst.Fields("pprice")
'txtsellingprice = rst.Fields("sprice")
'txtserai = rst.Fields("AI")
'txtsel
txtpcode11_Change
End If
End Sub
