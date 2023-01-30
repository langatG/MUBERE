VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FRMCHANGE 
   BackColor       =   &H8000000E&
   Caption         =   "CHANGE THE PRODUCT IN STOCK"
   ClientHeight    =   6570
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6300
   LinkTopic       =   "Form1"
   ScaleHeight     =   6570
   ScaleWidth      =   6300
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkquantity 
      Caption         =   "Quantity to change?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   22
      Top             =   960
      Width           =   2055
   End
   Begin VB.CheckBox chkprice 
      Caption         =   "Price to change?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   21
      Top             =   960
      Width           =   1695
   End
   Begin VB.ComboBox cmbbranch 
      Height          =   315
      ItemData        =   "FRMCHANGE.frx":0000
      Left            =   3960
      List            =   "FRMCHANGE.frx":0002
      TabIndex        =   20
      Top             =   1440
      Width           =   2175
   End
   Begin VB.CommandButton cmdreport 
      Caption         =   "Change Products Report"
      Height          =   375
      Left            =   3000
      TabIndex        =   17
      Top             =   3000
      Width           =   2175
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "Update"
      Height          =   375
      Left            =   1680
      TabIndex        =   16
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton cmdnew 
      Caption         =   "New"
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox txtchangepro 
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   4680
      TabIndex        =   13
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox txtbalnce 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   405
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox txtbuyingprice 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1560
      TabIndex        =   6
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox txtsellingprice 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      Top             =   1920
      Width           =   1455
   End
   Begin VB.TextBox txtpcode 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      DrawStyle       =   2  'Dot
      DrawWidth       =   17015
      Height          =   360
      Left            =   2880
      Picture         =   "FRMCHANGE.frx":0004
      ScaleHeight     =   360
      ScaleWidth      =   240
      TabIndex        =   1
      Top             =   120
      Width           =   240
   End
   Begin VB.ComboBox cboproductname 
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Top             =   600
      Width           =   4575
   End
   Begin MSComCtl2.DTPicker txtransdate 
      Height          =   375
      Left            =   4560
      TabIndex        =   7
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   122224641
      CurrentDate     =   40265
   End
   Begin MSComctlLib.ListView ListView30 
      Height          =   3015
      Left            =   120
      TabIndex        =   18
      Top             =   3480
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   5318
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
      MouseIcon       =   "FRMCHANGE.frx":0186
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ITEM"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "QNTY"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Balance"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Branch"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label5 
      Caption         =   "Branch"
      Height          =   255
      Left            =   3240
      TabIndex        =   19
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Physical Quantity"
      Height          =   375
      Left            =   3120
      TabIndex        =   14
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Balance in Store"
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Trans_Date"
      Height          =   255
      Left            =   3360
      TabIndex        =   10
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label24 
      Caption         =   "Buying Price"
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label25 
      Caption         =   "Selling Price"
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Product Name"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Product Code"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "FRMCHANGE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub chkprice_Click()
If chkprice = 1 Then
txtsellingprice.Visible = True
txtbuyingprice.Visible = True
Label24.Visible = True
Label25.Visible = True
Else
txtsellingprice.Visible = False
txtbuyingprice.Visible = False
Label24.Visible = False
Label25.Visible = False
End If
End Sub

Private Sub chkquantity_Click()
If chkquantity = 1 Then
txtbalnce.Visible = True
txtchangepro.Visible = True
Label4.Visible = True
Label2.Visible = True
Else
txtbalnce.Visible = False
txtchangepro.Visible = False
Label4.Visible = False
Label2.Visible = False
End If
End Sub

Private Sub cmdNew_Click()
txtpcode = ""
cboproductname = ""
txtbuyingprice = ""
txtsellingprice = ""
txtbalnce = ""
txtchangepro = ""
End Sub

Private Sub cmdreport_Click()
    reportname = "changepro.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub cmdsave_Click()
On Error GoTo ErrorHandler
If txtpcode = "" Then
MsgBox "Please insert the product code", vbCritical
Exit Sub
End If
If cboproductname = "" Then
MsgBox "Please select the product Name", vbCritical
Exit Sub
End If

If cmbbranch = "" Then
MsgBox "Please select the Branch", vbCritical
cmbbranch.SetFocus
Exit Sub
End If
'''''insert for change quantity
If chkquantity = vbChecked Then
If txtchangepro = "" Then
MsgBox "Please insert the new Quantity", vbCritical
Exit Sub
End If
Dim rst, rstg, rsa As Recordset
sql = ""
sql = "set dateformat dmy Select Qin, Qout,o_bal from ag_Products WHERE  p_code ='" & txtpcode & "' and Branch='" & cmbbranch & "'"
Set rst = cn.Execute(sql)
If Not rst.EOF Then
    sql = ""
    sql = "set dateformat DMY Update ag_Products SET Qin= '" & txtchangepro & "',Qout='" & txtchangepro & "',o_bal='" & txtchangepro & "' WHERE p_code ='" & txtpcode & "' and Branch='" & cmbbranch & "'"
    oSaccoMaster.ExecuteThis (sql)
End If
txtbuyingprice = "0"
txtsellingprice = "0"
sql = ""
sql = "set dateformat dmy insert into d_changepro(Date, Code, Name, Quantity, [user],PPrice, SPrice, Balance,Branch)values ('" & txtransdate & "','" & txtpcode & "','" & cboproductname & "','" & txtchangepro & "','" & User & "','" & txtbuyingprice & "','" & txtsellingprice & "','" & txtbalnce & "','" & cmbbranch & "')"
oSaccoMaster.ExecuteThis (sql)
End If
''''''end
''''''insert the price
If chkprice = vbChecked Then

If txtbuyingprice = "" Then
MsgBox "Please insert the new Buying price", vbCritical
Exit Sub
End If
If txtsellingprice = "" Then
MsgBox "Please insert the new selling price", vbCritical
Exit Sub
End If

'Dim rst, rstg, rsa As Recordset
sql = ""
sql = "set dateformat dmy Select Qin, Qout,o_bal,pprice, sprice from ag_Products WHERE  p_code ='" & txtpcode & "' and Branch='" & cmbbranch & "'"
Set rst = cn.Execute(sql)
If Not rst.EOF Then
    sql = ""
    sql = "set dateformat DMY Update ag_Products SET pprice= '" & txtbuyingprice & "',sprice='" & txtsellingprice & "'  WHERE p_code ='" & txtpcode & "' and Branch='" & cmbbranch & "'"
    oSaccoMaster.ExecuteThis (sql)
End If
txtchangepro = "0"
txtbalnce = "0"
sql = ""
sql = "set dateformat dmy insert into d_changepro(Date, Code, Name, Quantity, [user],PPrice, SPrice, Balance,Branch)values ('" & txtransdate & "','" & txtpcode & "','" & cboproductname & "','" & txtchangepro & "','" & User & "','" & txtbuyingprice & "','" & txtsellingprice & "','" & txtbalnce & "','" & cmbbranch & "')"
oSaccoMaster.ExecuteThis (sql)

End If
''''''end

 MsgBox "Completed succesfully ", vbInformation
txtpcode = ""
cboproductname = ""
txtbuyingprice = ""
txtsellingprice = ""
txtbalnce = ""
txtchangepro = ""
cmbbranch = ""
chkprice = vbUnchecked
chkquantity = vbUnchecked
loaddispmilk
Exit Sub
ErrorHandler:
MsgBox err.description
End Sub

Private Sub Form_Load()
txtransdate = Format(Date, "dd/mm/yyyy")
Provider = "MAZIWA"
Set cn = New ADODB.Connection
 cn.Open Provider, "atm", "atm"
sql = "select Distinct(P_NAME) from ag_products ORDER BY P_NAME ASC"
Set rs = New ADODB.Recordset
rs.Open sql, cn

While Not rs.EOF
cboproductname.AddItem rs.Fields(0)
rs.MoveNext
Wend
loaddispmilk
Branch
chkprice = 1
chkprice = 0
chkquantity = 1
chkquantity = 0
cboproductname.Enabled = True
'chkPrint.value = vbChecked
End Sub
Private Sub Branch()
    Set cn = New ADODB.Connection
    Provider = "MAZIWA"
    'Provider = cnn
    cn.Open Provider, "atm", "atm"
    'Set rst = New Recordset
    sql = ""
    sql = "Select distinct(Bname) from   d_Branch order by Bname asc"
    Set rst = oSaccoMaster.GetRecordset(sql)
   ' rst.Open sql, cn, adOpenKeyset, adLockOptimistic
    While Not rst.EOF
    cmbbranch.AddItem rst.Fields(0)
    rst.MoveNext
    Wend
'Provider = "MAZIWA"
'Set cn = New ADODB.Connection
' cn.Open Provider, "atm", "atm"
'sql = "select Distinct(Branch) from  ag_Products where Branch<>'' ORDER BY Branch ASC"
'Set rs = New ADODB.Recordset
'rs.Open sql, cn
'
'While Not rs.EOF
'cmbbranch.AddItem rs.Fields(0)
'rs.MoveNext
'Wend
End Sub
Public Sub loaddispmilk()
     
    With ListView30
    
        .ListItems.Clear
        
        .ColumnHeaders.Clear
    
    End With

    Set rs2 = CreateObject("adodb.recordset")
    
    sql = "set dateformat dmy Select * from  d_changepro"
    
    Set rs2 = CreateObject("adodb.recordset")
    
    Set clsClass = New cdbase
    
    Provider = clsClass.OpenCon
    
    Set cn = CreateObject("adodb.connection")
    
   cn.Open Provider, "atm", "atm"
    
    rs2.Open sql, cn
    
    With ListView30
    ', , ,
        
        .ColumnHeaders.Add , , "Date"
        .ColumnHeaders.Add , , "ITEM"
        .ColumnHeaders.Add , , "Name"
        .ColumnHeaders.Add , , "QNTY"
        .ColumnHeaders.Add , , "Balance"
        .ColumnHeaders.Add , , "Branch"
        While Not rs2.EOF
            Set li = .ListItems.Add(, , Trim(rs2.Fields("Date")))
            li.ListSubItems.Add , , Trim(rs2.Fields("Code"))
            li.ListSubItems.Add , , Trim(rs2.Fields("Name"))
            li.ListSubItems.Add , , Trim(rs2.Fields("Quantity"))
            li.ListSubItems.Add , , Trim(rs2.Fields("Balance"))
            li.ListSubItems.Add , , Trim(rs2.Fields("Branch"))
            rs2.MoveNext
        
        Wend
        
    End With
    
    rs2.Close
    
    Set rs2 = Nothing
    
ListView30.View = lvwReport

End Sub
Private Sub cboproductname_Change()
If cmbbranch = "" Then
MsgBox "Please select the Branch", vbCritical
cmbbranch.SetFocus
Exit Sub
End If

Set rst = oSaccoMaster.GetRecordset("select p_code,Qout,Pprice,sprice from ag_products where p_name ='" & cboproductname & "' and Branch='" & cmbbranch & "'")
If Not rst.EOF Then
txtpcode = rst.Fields("p_code")
txtbuyingprice = rst.Fields("Pprice")
txtsellingprice = rst.Fields("sprice")
txtbalnce = rst.Fields("Qout")
' Pprice , sprice
End If


End Sub

Private Sub cboproductname_Click()

If cmbbranch = "" Then
MsgBox "Please select the Branch", vbCritical
cmbbranch.SetFocus
Exit Sub
End If

Set rst = oSaccoMaster.GetRecordset("select p_code,Qout,Pprice,sprice from ag_products where p_name ='" & cboproductname & "' and Branch='" & cmbbranch & "'")
If Not rst.EOF Then
txtpcode = rst.Fields("p_code")
txtbuyingprice = rst.Fields("Pprice")
txtsellingprice = rst.Fields("sprice")
txtbalnce = rst.Fields("Qout")
' Pprice , sprice
End If

End Sub

Private Sub Picture1_Click()
frmSearch.Show vbModal
Dim Y As String
Y = sel
Dim p As Integer
If Y <> "" Then
'Provider = cn
Set cn = New ADODB.Connection
sql = "select P_CODE,P_NAME,QOUT,pprice,sprice from ag_products where p_code='" & Y & "'"
Set rs = New ADODB.Recordset
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
If Not IsNull(rs.Fields(0)) Then txtpcode.Text = (rs.Fields(0))
If Not IsNull(rs.Fields(1)) Then cboproductname = (rs.Fields(1))
If Not IsNull(rs.Fields(0)) Then txtbalnce = (rs.Fields(2))
If Not IsNull(rs.Fields(6)) Then txtbuyingprice = (rs.Fields(3))
If Not IsNull(rs.Fields(7)) Then txtsellingprice = (rs.Fields(4))
End If

'// check if the product have the serial then show the ag_receipts details
'cboproductname_Validate True

End If
End Sub
Private Sub get_serialno(pcode As String)
Provider = cn
Set cn = New ADODB.Connection
cn.Open Provider, "atm", "atm"
Dim RSSE As Recordset
sql = ""
sql = "select top 1 serialno,p_code,used from serialno where p_code='" & txtpcode & "'  order by serialid desc"
Set RSSE = New ADODB.Recordset

RSSE.Open sql, cn, adOpenKeyset, adLockOptimistic
If RSSE.Fields(2) = 1 Then
MsgBox "Serial Number and receipt no used please check again before posting", vbCritical
Exit Sub
End If
End Sub

Private Sub txtpcode_Change()
If cmbbranch = "" Then
MsgBox "Please select the Branch", vbCritical
cmbbranch.SetFocus
Exit Sub
End If

If KeyAscii = 13 Then
Provider = "MAZIWA"
Set cn = New ADODB.Connection
 cn.Open Provider, "atm", "atm"
'If reportpath = "" Then reportpath = GetSetting("payroll", "AppName", "rptPath", rptPath)
sql = "select P_CODE,P_NAME,S_NO,QOUT,supplierid,pprice,sprice from ag_products where p_code='" & txtpcode & "' and Branch='" & cmbbranch & "'"
Set rs = New ADODB.Recordset
rs.Open sql, cn
If Not rs.EOF Then
 
If Not IsNull(rs.Fields(1)) Then cboproductname = (rs.Fields(1))
If Not IsNull(rs.Fields(1)) Then txtbalnce = (rs.Fields(3))
If Not IsNull(rs.Fields(5)) Then txtbuyingprice = (rs.Fields(5))
If Not IsNull(rs.Fields(6)) Then txtsellingprice = (rs.Fields(6))
End If
End If
End Sub
Private Sub txtpcode_KeyPress(KeyAscii As Integer)
'//TWNG001
If cmbbranch = "" Then
MsgBox "Please select the Branch", vbCritical
cmbbranch.SetFocus
Exit Sub
End If

If KeyAscii = 13 Then
Provider = "MAZIWA"
Set cn = New ADODB.Connection
 cn.Open Provider, "atm", "atm"
'If reportpath = "" Then reportpath = GetSetting("payroll", "AppName", "rptPath", rptPath)
sql = "select P_CODE,P_NAME,S_NO,QOUT,supplierid,pprice,sprice from ag_products where p_code='" & txtpcode & "'  and Branch='" & cmbbranch & "'"
Set rs = New ADODB.Recordset
rs.Open sql, cn
If Not rs.EOF Then
 
If Not IsNull(rs.Fields(1)) Then cboproductname = (rs.Fields(1))
If Not IsNull(rs.Fields(1)) Then txtbalnce = (rs.Fields(3))
If Not IsNull(rs.Fields(5)) Then txtbuyingprice = (rs.Fields(5))
If Not IsNull(rs.Fields(6)) Then txtsellingprice = (rs.Fields(6))

End If
End If
'// check with serial no if it exist
End Sub
