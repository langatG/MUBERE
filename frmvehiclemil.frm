VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmvehiclemil 
   Caption         =   "Plant Milk Sales Receive"
   ClientHeight    =   6885
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8595
   LinkTopic       =   "Form1"
   ScaleHeight     =   6885
   ScaleWidth      =   8595
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtfield 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   2760
      TabIndex        =   16
      Top             =   600
      Width           =   1095
   End
   Begin VB.CheckBox chprint 
      Caption         =   "Use LPT1 Printer"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4920
      TabIndex        =   9
      Top             =   1680
      Width           =   3255
   End
   Begin VB.CheckBox chkPrint 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "Print Receipt"
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
      Height          =   405
      Left            =   0
      TabIndex        =   8
      Top             =   1800
      Width           =   2655
   End
   Begin VB.TextBox txtQnty 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   2760
      TabIndex        =   7
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton cmdReceive 
      Appearance      =   0  'Flat
      Caption         =   "Receive"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   240
      TabIndex        =   6
      ToolTipText     =   "Click to receive the milk"
      Top             =   2760
      Width           =   1335
   End
   Begin VB.ComboBox ports 
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "frmvehiclemil.frx":0000
      Left            =   4440
      List            =   "frmvehiclemil.frx":0010
      TabIndex        =   5
      Text            =   "\\127.0.0.1\E-PoS 80mm Thermal Printer "
      Top             =   2160
      Width           =   3735
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
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
      Left            =   4440
      TabIndex        =   4
      Top             =   2760
      Width           =   1095
   End
   Begin VB.ComboBox cbovb 
      Height          =   315
      Left            =   2760
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton cmdvhreport 
      Caption         =   "Vehicle Report"
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
      Left            =   1680
      TabIndex        =   0
      Top             =   2760
      Width           =   1575
   End
   Begin MSComctlLib.ListView ListView3 
      Height          =   2535
      Left            =   120
      TabIndex        =   2
      Top             =   3600
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   4471
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   65280
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Branch"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Quantity"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Actual"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Varriance"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Vehicle No."
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   6630
      Width           =   8595
      _ExtentX        =   15161
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   7937
            MinWidth        =   7937
            Text            =   "USER : Birgen Gideon K."
            TextSave        =   "USER : Birgen Gideon K."
            Object.ToolTipText     =   "EASYMA User"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
            Picture         =   "frmvehiclemil.frx":002C
            Text            =   "DATE : 07/12/2009"
            TextSave        =   "DATE : 07/12/2009"
            Object.ToolTipText     =   "Today's Date"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "10:15 AM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Enabled         =   0   'False
            TextSave        =   "NUM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog cdgPrint 
      Left            =   5160
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FileName        =   "c:\receipt.txt"
   End
   Begin MSComCtl2.DTPicker DTPMilkDate 
      Height          =   375
      Left            =   6480
      TabIndex        =   11
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      MouseIcon       =   "frmvehiclemil.frx":01C0
      CalendarBackColor=   8454016
      Format          =   174850049
      CurrentDate     =   40095
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Today's collection"
      Height          =   2895
      Left            =   0
      TabIndex        =   3
      Top             =   3360
      Width           =   8415
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Plant Actual (Kgs)"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   15
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Milk Date:"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   14
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Field Milk (Kgs)"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Vehicle No:"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmvehiclemil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdvhreport_Click()
Dim ans As String
ans = MsgBox("Do you Want a Report as per Vehicle??", vbYesNo)
If ans = vbYes Then
reportname = "d_vehicledeliveryper1.rpt"
Else
reportname = "d_vehicledelivery.rpt"
End If
Show_Sales_Crystal_Report "", reportname, ""
End Sub


Private Sub DTPMilkDate_Click()
loadMilk
End Sub
Private Sub DTPMilkDate_Change()
loadMilk
End Sub

Private Sub Form_Load()
DTPMilkDate = Format(Get_Server_Date, "dd/mm/yyyy")
DTPMilkDate.MaxDate = Format(Get_Server_Date, "dd/mm/yyyy")
'DTPcomplaintperiod1 = DTPMilkDate
With StatusBar1.Panels
    .Item(1).Text = "USER : " & username
    .Item(2).Text = "DATE : " & Format(Get_Server_Date, "dd/mm/yyyy")

End With
    NAMES3

loadMilk
End Sub
Public Sub loadMilk()

    txtfield = "0"
    txtQnty = "0"

    '/// to list view//////////
sql = ""
sql = "set dateformat dmy SELECT  Vehicle, Quantity, Actual, Varriance From d_MilkVehicle where Date ='" & DTPMilkDate & "'"
'sql = "set dateformat dmy SELECT d.DCode, d.DName, m.DispQnty,m.DispDate FROM  d_Debtors AS d INNER JOIN d_MilkControl AS m ON d.DCode = m.dcode WHERE     (DispDate = '" & DTPDispatchDate & "') and status=0"
Set rs = oSaccoMaster.GetRecordset(sql)
If rs.EOF Then
Exit Sub
End If
ListView3.ListItems.Clear
While Not rs.EOF
If Not IsNull(rs.Fields(0)) Then
Set li = ListView3.ListItems.Add(, , rs.Fields(0))
End If
                    If Not IsNull(rs.Fields(1)) Then li.SubItems(1) = rs.Fields(1) & ""
                    If Not IsNull(rs.Fields(2)) Then li.SubItems(2) = rs.Fields(2) & ""
                    If Not IsNull(rs.Fields(3)) Then li.SubItems(3) = rs.Fields(3) & ""
'                    If Not IsNull(rs.Fields(4)) Then li.SubItems(4) = rs.Fields(4) & ""
'                    If Not IsNull(rs.Fields(5)) Then li.SubItems(5) = rs.Fields(5) & ""
rs.MoveNext

Wend

'////// end of view
    

End Sub
Private Sub NAMES3()
'Private Sub SSTab1_DblClick()
    cbovb.Clear
    Set rst = New Recordset
    Dim cn As Connection
    Set cn = New ADODB.Connection
    Provider = "MAZIWA"
    'Provider = cnn
    cn.Open Provider, "atm", "atm"
    Set rst = New Recordset
    sql = "Select distinct(Locations) from   d_Debtors order by Locations"
    'Select distinct(Locations) from   d_Debtors
    rst.Open sql, cn, adOpenKeyset, adLockOptimistic
    While Not rst.EOF
    cbovb.AddItem rst.Fields(0)
    rst.MoveNext
    Wend
End Sub
Private Sub cmdReceive_Click()
On Error GoTo ErrorHandler
Dim Price As Currency
Dim Startdate, CummulKgs, CummulKgs1, TRANSPORTER As String
Dim transdate As Date, anss As String
'check the user
 sql = "SELECT UserLoginIDs, UserGroup, SUPERUSER,status From UserAccounts where UserLoginIDs='" & User & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
    If rs!SuperUser <> 1 Then
    MsgBox "You are not allowed to Receive plant milk", vbInformation
     
    Exit Sub
   End If
     End If

If cbovb = "" Then
   MsgBox "PLEASE SELECT THE VEHICLE."
    cbovb.SetFocus
   Exit Sub
 End If

If Trim(txtfield) = "" Then
    MsgBox "Please enter the quantity supplied From the Field."
    txtQnty.SetFocus
Exit Sub
End If

If Not IsNumeric(txtfield) Then
MsgBox "Please enter a number. " & txtQnty & " is not a number", vbExclamation
txtQnty.SetFocus
Exit Sub
End If


If Trim(txtQnty) = "" Then
    MsgBox "Please enter the Actual quantity delivered by the vehicle."
    txtQnty.SetFocus
Exit Sub
End If

If Not IsNumeric(txtQnty) Then
MsgBox "Please enter a number. " & txtQnty & " is not a number", vbExclamation
txtQnty.SetFocus
Exit Sub
End If

Startdate = DateSerial(year(DTPMilkDate), month(DTPMilkDate), 1)
Enddate = DateSerial(year(DTPMilkDate), month(DTPMilkDate) + 1, 1 - 1)

sql = ""
sql = "set dateformat dmy SELECT * From d_MilkVehicle where Date ='" & DTPMilkDate & "' and Vehicle='" & cbovb & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If rs.EOF Then
sql = ""
sql = "set dateformat dmy INSERT INTO d_MilkVehicle"
sql = sql & " ( Vehicle, Quantity, Actual, Varriance, Date)"
sql = sql & " VALUES ('" & cbovb & "','" & txtfield & "'," & txtQnty & ",'" & txtfield.Text - txtQnty.Text & "','" & DTPMilkDate & "')"
oSaccoMaster.Execute (sql)
Else
sql = ""
sql = "SET dateformat DMY Update  d_MilkVehicle SET Quantity='" & rs.Fields(1) + txtfield & "', Actual='" & rs.Fields(2) + txtQnty & "',Variance='" & (rs.Fields(3) + txtQnty) - txtfield & "' WHERE Date ='" & DTPMilkDate & "'and Vehicle='" & cbovb & "'"
oSaccoMaster.Execute (sql)
End If

'loadMilk

'//Print Receipt
    If chkPrint = vbChecked Then
    
'/*Print out
 Dim fso, chkPrinter, txtFile
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
       'cdgPrint.Filter = "*.csv|*.txt"
        'cdgPrint.ShowSave
        ttt = "\\127.0.0.1\E-PoS 80mm Thermal Printer" 'LPT1
        'ttt = "D:\PROJECTS\FOSA\DAILY" & Date & ""
        Set fso = CreateObject("Scripting.FileSystemObject")
        'Set chkPrinter = fso.GetFile(ttt)
       
        
        Set txtFile = fso.CreateTextFile(ttt, True)
    txtFile.WriteLine "      " & cname & ""
    txtFile.WriteLine "             Milk Receipt for the Vehicle"
    txtFile.WriteLine "---------------------------------------"
        
    Set rs2 = New ADODB.Recordset
    sql = "d_sp_ReceiptNumber"
    Set rs2 = oSaccoMaster.GetRecordset(sql)
    
    Dim RNumber As String
    'RNumber = rs2.Fields(0)
    If Not IsNull(rs2.Fields(0)) Then RNumber = rs2.Fields(0)
    'Else
    'RNumber = "0"
    'End If
    
    'txtFile.WriteLine "SNo :" & txtSNo
    txtFile.WriteLine "Vehicle :" & cbolocation
    txtFile.WriteLine "Feld Supplied :" & txtfield & " Kgs"
    txtFile.WriteLine "Actual Supplied :" & txtQnty & " Kgs"
    'txtFile.WriteLine "Price" & Price & " Per Kgs"
    Set rs = New ADODB.Recordset
    sql = ""
    sql = "set dateformat dmy select  sum(Quantity), sum(Actual) FROM  d_MilkVehicle where Vehicle='" & cbovb & "' and Date>='" & Startdate & "' and Date<='" & Enddate & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
     If Not IsNull(rs.Fields(0)) Then
      CummulKgs = rs.Fields(0)
     Else
      CummulKgs = "0.00"
     End If
     If Not IsNull(rs.Fields(1)) Then
      CummulKgs1 = rs.Fields(1)
     Else
      CummulKgs1 = "0.00"
     End If
    End If
    Dim varia As Double
    varia = CummulKgs - CummulKgs1
    txtFile.WriteLine "Field kgs This Month : " & Format(CummulKgs, "#,##0.00" & " Kgs")
    txtFile.WriteLine "Actual kgs This Month : " & Format(CummulKgs1, "#,##0.00" & " Kgs")
    txtFile.WriteLine "Varriance This Month : " & Format(varia, "#,##0.00" & " Kgs")
    txtFile.WriteLine "---------------------------------------"
    txtFile.WriteLine "Receipt Number :" & RNumber
  '  txtFile.WriteLine "TRANSPORTER :" & TRANSPORTER
    txtFile.WriteLine "Received by :" & username
    txtFile.WriteLine "---------------------------------------"
    txtFile.WriteLine "  Date :" & Format(DTPMilkDate, "dd/mm/yyyy") & " ,Time : " & Format(Time, "hh:mm:ss AM/PM")
    txtFile.WriteLine "       " & motto & ""
    txtFile.WriteLine "---------------------------------------"
'    If chkComment.value = vbChecked Then
'        txtFile.WriteLine txtComment
'        txtFile.WriteLine "---------------------------------------"
'    End If
    txtFile.WriteLine escFeedAndCut
    
 txtFile.Close
 Reset
End If


loadMilk

txtfield = ""
txtQnty = ""
cbovb = ""
'txtSNo_Validate True
'txtQnty.SetFocus
Exit Sub
ErrorHandler:

MsgBox err.description
End Sub

