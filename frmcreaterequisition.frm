VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmcreaterequisition 
   Caption         =   "Item Requisitions"
   ClientHeight    =   8130
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9165
   LinkTopic       =   "Form1"
   ScaleHeight     =   8130
   ScaleWidth      =   9165
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtcomment 
      Height          =   495
      Left            =   6600
      TabIndex        =   30
      Top             =   2640
      Width           =   2535
   End
   Begin VB.ComboBox ports 
      Height          =   315
      ItemData        =   "frmcreaterequisition.frx":0000
      Left            =   8160
      List            =   "frmcreaterequisition.frx":0010
      TabIndex        =   28
      Text            =   "COM1"
      Top             =   1560
      Width           =   855
   End
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
      Left            =   7080
      TabIndex        =   27
      Top             =   960
      Value           =   2  'Grayed
      Width           =   1935
   End
   Begin VB.TextBox txttotal 
      Height          =   375
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   7680
      Width           =   2295
   End
   Begin VB.TextBox txtpcode 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2640
      TabIndex        =   22
      Top             =   1320
      Width           =   2535
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      DrawStyle       =   2  'Dot
      DrawWidth       =   17015
      Height          =   360
      Left            =   5280
      Picture         =   "frmcreaterequisition.frx":002C
      ScaleHeight     =   360
      ScaleWidth      =   480
      TabIndex        =   21
      Top             =   1320
      Width           =   480
   End
   Begin VB.ComboBox cboproductname 
      Height          =   315
      Left            =   2640
      TabIndex        =   20
      Top             =   1800
      Width           =   4215
   End
   Begin VB.CommandButton cmdnew 
      Caption         =   "New."
      Height          =   375
      Left            =   1200
      TabIndex        =   19
      Top             =   7680
      Width           =   975
   End
   Begin VB.TextBox txtPrice 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7560
      TabIndex        =   18
      Text            =   "0"
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox txtbalance 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5160
      TabIndex        =   15
      Text            =   "0"
      Top             =   2280
      Width           =   1215
   End
   Begin VB.ComboBox cbocostcentre 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2640
      TabIndex        =   14
      Top             =   840
      Width           =   2295
   End
   Begin MSComCtl2.DTPicker dtptransdate 
      Height          =   255
      Left            =   6480
      TabIndex        =   13
      Top             =   480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      Format          =   121831425
      CurrentDate     =   41785
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      Height          =   375
      Left            =   3960
      TabIndex        =   12
      Top             =   7680
      Width           =   975
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
      Height          =   375
      Left            =   2640
      TabIndex        =   11
      Top             =   7680
      Width           =   1095
   End
   Begin VB.CommandButton cmdremove 
      Caption         =   "Remove"
      Height          =   375
      Left            =   4320
      TabIndex        =   10
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "Add"
      Height          =   375
      Left            =   2760
      TabIndex        =   9
      Top             =   2760
      Width           =   1215
   End
   Begin MSComctlLib.ListView lvwrequisition 
      Height          =   4335
      Left            =   360
      TabIndex        =   8
      Top             =   3240
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   7646
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ItemNo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Transdate"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Vendor"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Description"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Quantity"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Price"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Total Price"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Comments"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox txtquantity 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2640
      TabIndex        =   7
      Text            =   "0"
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox txtRNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2640
      TabIndex        =   3
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label8 
      Caption         =   "Comment"
      Height          =   255
      Left            =   5760
      TabIndex        =   31
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "Printer Port"
      Height          =   375
      Left            =   7080
      TabIndex        =   29
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "Total"
      Height          =   255
      Left            =   5520
      TabIndex        =   26
      Top             =   7800
      Width           =   855
   End
   Begin VB.Label Label14 
      Caption         =   "Product Name"
      Height          =   255
      Left            =   360
      TabIndex        =   24
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label13 
      Caption         =   "Product Code"
      Height          =   255
      Left            =   360
      TabIndex        =   23
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label12 
      Caption         =   "Price"
      Height          =   255
      Left            =   6600
      TabIndex        =   17
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label11 
      Caption         =   "Balance"
      Height          =   255
      Left            =   4320
      TabIndex        =   16
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label Label9 
      Caption         =   "Quatitity"
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Requisition Header"
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
      Left            =   360
      TabIndex        =   5
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "ITEM REQUISITIONS"
      Height          =   255
      Left            =   2400
      TabIndex        =   4
      Top             =   0
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "Requisition Number"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Vendors"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Date Required"
      Height          =   375
      Left            =   5160
      TabIndex        =   0
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "frmcreaterequisition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsr As New ADODB.Recordset
Dim Mylength As Integer

Private Sub cbocostcentre_KeyPress(KeyAscii As Integer)
KeyAscii = 0
Beep

End Sub

Private Sub cboproductname_Change()
Set rst = oSaccoMaster.GetRecordset("select p_code from ag_products where p_name ='" & cboproductname & "'")
If Not rst.EOF Then
txtpcode = rst.Fields("p_code")
End If
End Sub

Private Sub cboproductname_Click()
cboproductname_Change
End Sub

'Private Sub chkService_Click()
'If chkService.value = vbChecked Then
'    txtquantity.Enabled = False
'    txtquantity.BackColor = vbInactiveBorder
'Else
'    txtquantity.Enabled = True
'    txtquantity.BackColor = vbWhite
'End If
'End Sub

Private Sub cmdAdd_Click()
On Error GoTo ErrorHandler
If txtquantity = "" Then
   MsgBox "Please capture Product Quantity to be ordered", vbInformation
   txtquantity.SetFocus
Exit Sub
End If
If txtPrice = "" Then
   MsgBox "Please indicate cost of the product", vbInformation
   txtPrice.SetFocus
Exit Sub
End If
If cboproductname = "" Then
   MsgBox "Please Select product to order", vbInformation
   cboproductname.SetFocus
Exit Sub
End If
If txtpcode = "" Then
   MsgBox "Please Select product to order", vbInformation
   txtpcode.SetFocus
Exit Sub
End If

Set li = lvwrequisition.ListItems.Add(, , txtrno)
    li.SubItems(1) = (DTPTransdate)
    li.SubItems(2) = cbocostcentre
    li.SubItems(3) = cboproductname.Text
    li.SubItems(4) = txtquantity
    li.SubItems(5) = txtPrice
    li.SubItems(6) = CDbl(txtPrice) * CDbl(txtquantity)
    li.SubItems(7) = cboproductname.Text
    txtbalance = 0
    txtPrice = 0
    txtquantity = 0
    
    Calculate_Total

    Exit Sub
ErrorHandler:
MsgBox err.description
End Sub
Sub Calculate_Total()
    Dim total As Double, amt As Double, Price As Double, qnty As Integer
    Dim ccount As Integer
    On Error Resume Next
    total = 0
    With lvwrequisition
        If .ListItems.Count > 0 Then
            ccount = .ListItems.Count
            For I = 1 To ccount
                With .ListItems(I)
                        Price = CDbl(.ListSubItems(5))
                        qnty = CDbl(.ListSubItems(4))
                        amt = Price * qnty
                        total = total + amt
                End With
            Next I

        Else
            total = 0
        End If
    End With
    TXTTOTAL = Format(total, Cfmt)
End Sub

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdNew_Click()
On Error GoTo ErrorHandler
 mysql = ""
        mysql = "select * from Receiptno where receiptno like 'RQ-%' order by Receipthnoid desc"
        
        Set rsr = oSaccoMaster.GetRecordset(mysql)
        
        If Not rsr.EOF Then
            Mylength = CInt(Mid(rsr!ReceiptNo, 5, 10))
            Mylength = Mylength + 1
            txtrno = Padding(Mylength)
            txtrno = "RQ-" & txtrno
        Else
            Mylength = 1
            txtrno = "RQ-" & Padding(Mylength)
            
        End If
        Exit Sub
ErrorHandler:
        MsgBox err.description
End Sub

Private Sub cmdRemove_Click()
 If lvwrequisition.ListItems.Count > 0 Then
        If MsgBox("Are you sure you delete  this records " & lvwrequisition.SelectedItem.Text & "? ", vbYesNo) = vbYes Then
        lvwrequisition.ListItems.Remove (lvwrequisition.SelectedItem.Index)  '// removes the selected item
        End If
    End If
    Calculate_Total
End Sub

Private Sub cmdsave_Click()
   Dim PNo As Double, ReceiptNo As String, lenght As Integer
    If DTPTransdate = "" Then
        MsgBox "Please enter the requistion date.", vbExclamation, "MISSING DETAILS"
            DTPTransdate.SetFocus
        Exit Sub
    End If
    
    If cbocostcentre = "<Select Cost Center>" Then
        MsgBox "Please select the cost center.", vbExclamation, "MISSING DETAILS"
            cbocostcentre.SetFocus
        Exit Sub
    End If
  
    Dim chequeno As String
    Dim Rno As String, tdate As Date, CC As String, RNAME  As String, _
    MAKE As String, q As Double, DocumentNo As String, doc_posted As Integer, Totalprice, pprice As Double
    If lvwrequisition.ListItems.Count > 0 Then
        If MsgBox("Do you want post the entry?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
            Exit Sub
        End If
    Else
        MsgBox "There are no transactions to be posted", vbInformation, Me.Caption
        Exit Sub
    End If
    
    ''''' PURCHASE ORDER ''''''''''
        Set rst = oSaccoMaster.GetRecordset("d_sp_PoNo")
        If Not rst.EOF Then
            PNo = CCur(rst.Fields(0)) + 1
            Else
             PNo = "1"
        End If
        
        ReceiptNo = txtrno
           Me.MousePointer = vbHourglass
    For I = 1 To lvwrequisition.ListItems.Count
        Set li = lvwrequisition.ListItems(I)
        'Receiptno = li
        tdate = (lvwrequisition.ListItems(I).SubItems(1))
        CC = lvwrequisition.ListItems(I).SubItems(2)
        RNAME = lvwrequisition.ListItems(I).SubItems(3)
        MAKE = lvwrequisition.ListItems(I).SubItems(3)
        q = CDbl(lvwrequisition.ListItems(I).SubItems(4))
        DocumentNo = lvwrequisition.ListItems(I).SubItems(7)
        pprice = lvwrequisition.ListItems(I).SubItems(5)
        Totalprice = lvwrequisition.ListItems(I).SubItems(6)

        If I <> 1 Then
            ReceiptNo = ReceiptNo
            Mylength = CInt(Mid(ReceiptNo, 5, 10))
            Mylength = Mylength + 1
            ReceiptNo = Padding(Mylength)
            ReceiptNo = "RQ-" & ReceiptNo
            
        End If
            doc_posted = 0
                
                 '********* SAVE REQUISITION ***********
                sql = ""
                sql = "d_sp_Requisition '" & ReceiptNo & "', '" & tdate & "', '" & CC & "', " & doc_posted & "," & Totalprice & " ,'" & RNAME & "', '" & MAKE & "', " & q & ", '" & DocumentNo & "','" & User & "'," & pprice & ",'" & Format(Get_Server_Date, "dd/mm/yyyy") & "'," & PNo & ""
                oSaccoMaster.ExecuteThis (sql)
                
                    mysql = "set dateformat dmy Insert into Receiptno(Receiptno,Auditdate,auditid)values('" & ReceiptNo & "','" & Format(Get_Server_Date, "dd/MM/yyyy") & "','" & User & "')"
           oSaccoMaster.ExecuteThis (mysql)
    
    
    ' APPROVE REQUISITION
        sql = ""
           sql = "d_insert_d_Approve '" & ReceiptNo & "','0','Order','" & User & "'"
           oSaccoMaster.ExecuteThis (sql)
             
    Next I
    '//clear listview
         '********* SAVE LPO ***********
    oSaccoMaster.ExecuteThis ("d_sp_LPO " & PNo & ",'" & tdate & "','" & tdate & "','" & ReceiptNo & "','" & ReceiptNo & "','" & User & "','Ordered','" & CC & "'")
        
        If chkPrint.value = vbChecked Then
Dim total As Double, j As Long
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
       Dim fso, chkPrinter, txtFile
        'ttt = "LPT1" 'LPT1
         Dim PORT As String
        PORT = ports
        'ttt = "LPT1" 'LPT1
        ttt = PORT
        Set fso = CreateObject("Scripting.FileSystemObject")
        Dim strReceipts As String

        'MsgBox strReceipts
        Set txtFile = fso.CreateTextFile(ttt, True)
    txtFile.WriteLine "      " & cname & ""
    txtFile.WriteLine "      AGROVET"
    txtFile.WriteLine "      " & paddress & ""
    txtFile.WriteLine "      " & town & ""
    txtFile.WriteLine "      " & Phone & ""
    txtFile.WriteLine "      " & Email & ""
    txtFile.WriteLine "---------------------------------------"
    txtFile.WriteLine "    REQUISITION"
    txtFile.WriteLine
    txtFile.WriteLine "---------------------------------------"
    txtFile.WriteLine "ITEM" & vbTab & vbTab & "QNTY" & vbTab & "PRICE" & vbTab & "AMOUNT"
    
         j = 1
       strReceipts = ""
    Do While Not j > (lvwrequisition.ListItems.Count)
        lvwrequisition.ListItems.Item(j).selected = True
        lenght = Len(lvwrequisition.SelectedItem.SubItems(3))
        strReceipts = Mid(lvwrequisition.SelectedItem.SubItems(3), 5, lenght - 5)
        If Len(strReceipts) > 14 Then
        strReceipts = strReceipts & "-"
        Else
        strReceipts = strReceipts & vbTab
        End If
        strReceipts = strReceipts & CDbl(lvwrequisition.SelectedItem.SubItems(4)) & vbTab & Format(lvwrequisition.SelectedItem.SubItems(5), "#,##0.00") & vbTab & Format(lvwrequisition.SelectedItem.SubItems(6), "#,##0.00") & vbNewLine
        txtFile.WriteLine strReceipts
        j = j + 1
    Loop
    txtFile.WriteLine "---------------------------------------" & vbNewLine
    txtFile.WriteLine "RECEIPT TOTAL" & vbTab & vbTab & Format(TXTTOTAL, "#,##0.00")
    txtFile.WriteLine "==============================================="
    txtFile.WriteLine
    txtFile.WriteLine "Remarks" & vbTab & txtComment
    txtFile.WriteLine
    txtFile.WriteLine "YOU WERE SERVED By " & UCase(username)
    txtFile.WriteLine "    Date :" & Format(Get_Server_Date, "dd/mm/yyyy HH:MM:SS AM/PM")
    txtFile.WriteLine "     THANK YOU AND WELCOME "
    txtFile.WriteLine "****************************************"
    txtFile.WriteLine escFeedAndCut
    txtFile.Close
End If

     lvwrequisition.ListItems.Clear
     txtComment = ""
     Form_Load
    
    Me.MousePointer = vbDefault
    MsgBox "Posting Successfull", vbInformation, Me.Caption
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
    Me.MousePointer = vbDefault
    
End Sub

Private Sub Form_Load()
txtComment = ""
DTPTransdate = Format(Get_Server_Date, "dd/mm/yyyy")
txtquantity = Format(CCur(txtquantity), "#,##0.00")
txtPrice = Format(CCur(txtPrice), "#,##0.00")
chkPrint.value = vbChecked
'//LOAD COST CENTRES

'Set rs = oSaccoMaster.GetRecordset("select description from  d_CostCent order by description")
'While Not rs.EOF
'cbocostcentre.AddItem rs.Fields(0)
'rs.MoveNext
'Wend

' LOAD VENDORS
    cbocostcentre.Clear
    sql = "Select companyname from ag_Supplier1"
    Set rs = oSaccoMaster.GetRecordset(sql)
    While Not rs.EOF
    cbocostcentre.AddItem rs.Fields(0)
    rs.MoveNext
    Wend

' LOAD Products
sql = "select P_NAME  from ag_products ORDER BY P_NAME ASC"
Set rs = oSaccoMaster.GetRecordset(sql)
While Not rs.EOF
cboproductname.AddItem rs.Fields(0)
rs.MoveNext
Wend
cboproductname.Enabled = True


'get the new requition nuo

 mysql = ""
        mysql = "select * from Receiptno where receiptno like 'RQ-%' order by Receipthnoid desc"
        
        Set rsr = oSaccoMaster.GetRecordset(mysql)
        
        If Not rsr.EOF Then
            Mylength = CInt(Mid(rsr!ReceiptNo, 5, 10))
            Mylength = Mylength + 1
            txtrno = Padding(Mylength)
            txtrno = "RQ-" & txtrno
        Else
            Mylength = 1
            txtrno = "RQ-" & Padding(Mylength)
            
        End If
End Sub

Private Sub Picture1_Click()
frmSearch.Show vbModal
Dim Y As String
Y = sel
Dim p As Integer
If Y <> "" Then
sql = "select P_CODE,P_NAME,S_NO,QOUT,seria,s_no from ag_products where p_code='" & Y & "'"
Set rs = New ADODB.Recordset
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
 txtpcode = (rs.Fields(0))
 cboproductname.Text = IIf(IsNull(rs.Fields(1)), "", rs.Fields(1))
End If
End If
End Sub

Private Sub Picture2_Click()
frmSearch.Show vbModal
Dim Y As String
Y = sel

If Y <> "" Then
Provider = "MAZIWA"
Set cn = New ADODB.Connection
cn.Open Provider, "atm", "atm"
'If reportpath = "" Then reportpath = GetSetting("payroll", "AppName", "rptPath", rptPath)
sql = "select P_CODE,P_NAME,S_NO,QOUT,pprice from ag_products where p_code='" & Y & "'"
Set rs = New ADODB.Recordset
rs.Open sql, cn
If Not rs.EOF Then
'If Not IsNull(rs.Fields(0)) Then txtpcode = (rs.Fields(0))
If Not IsNull(rs.Fields(1)) Then cboproductname = (rs.Fields(1))
If Not IsNull(rs.Fields(2)) Then txtPrice = (rs.Fields(4))
If Not IsNull(rs.Fields(3)) Then txtbalance = (rs.Fields(3))
'// check with serial no if it exist
End If
End If

End Sub

Private Sub txtpcode_Change()
sql = "select P_CODE,P_NAME,S_NO,QOUT,supplierid,pprice, sprice from ag_products where p_code='" & txtpcode & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
 txtpcode = (rs.Fields(0))
 cboproductname.Text = IIf(IsNull(rs.Fields(1)), "", rs.Fields(1))
 txtbalance = IIf(IsNull(rs.Fields(3)), "", rs.Fields(3))
 txtPrice = IIf(IsNull(rs.Fields(5)), "", rs.Fields(5))
 cbocostcentre = IIf(IsNull(rs.Fields(4)), "", rs.Fields(4))
'If Not IsNull(rs.Fields(1)) Then txtpName = (rs.Fields(1))
'If Not IsNull(rs.Fields(3)) Then txtbalance = (rs.Fields(3))
'If Not IsNull(rs.Fields(4)) Then cbosupplier = (rs.Fields(4))
'If Not IsNull(rs.Fields(5)) Then txtPPrice = (rs.Fields(5))
'If Not IsNull(rs.Fields(6)) Then txtsellingprice = (rs.Fields(6))
End If
End Sub

Private Sub txtPrice_Click()
txtPrice = Format(CCur(txtPrice), "#0")
End Sub

Private Sub txtPrice_Validate(Cancel As Boolean)
txtPrice = Format(CCur(txtPrice), "#,##0.00")
End Sub

Private Sub txtquantity_Click()
'txtquantity = Format(CCur(txtquantity), "#0.00")
End Sub

Private Sub txtquantity_Validate(Cancel As Boolean)
If Trim(txtquantity) = "" Then
txtquantity = "0"
End If
If Not IsNumeric(txtquantity) Then
MsgBox UCase(txtquantity) & " is not a number please enter a valid number.", vbExclamation
txtquantity.SetFocus
Exit Sub
End If
txtquantity = Format(CCur(txtquantity), "#,##0.00")
End Sub


Private Sub txttotal_Change()
TXTTOTAL = Format(TXTTOTAL, Cfmt)
End Sub
