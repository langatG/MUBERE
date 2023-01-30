VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmbranchmilk 
   Caption         =   "Receive Milk from Branch"
   ClientHeight    =   6570
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8460
   LinkTopic       =   "Form1"
   ScaleHeight     =   6570
   ScaleWidth      =   8460
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdvhreport 
      Caption         =   "Vehicle Report"
      Height          =   495
      Left            =   1560
      TabIndex        =   22
      Top             =   2760
      Width           =   1335
   End
   Begin VB.ComboBox cbovb 
      Height          =   315
      Left            =   4800
      TabIndex        =   20
      Top             =   1200
      Width           =   2175
   End
   Begin VB.CommandButton cmdnewbra 
      Caption         =   "New Branch"
      Height          =   495
      Left            =   5400
      TabIndex        =   19
      Top             =   2760
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSComctlLib.ListView ListView3 
      Height          =   2535
      Left            =   120
      TabIndex        =   18
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Today's collection"
      Height          =   2895
      Left            =   0
      TabIndex        =   16
      Top             =   3360
      Width           =   8415
   End
   Begin VB.CommandButton cmdreport 
      Caption         =   "Branch Report"
      Height          =   495
      Left            =   120
      TabIndex        =   17
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      Height          =   495
      Left            =   7080
      TabIndex        =   15
      Top             =   2760
      Width           =   1095
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
      ItemData        =   "frmbranchmilk.frx":0000
      Left            =   4440
      List            =   "frmbranchmilk.frx":0010
      TabIndex        =   14
      Text            =   "\\127.0.0.1\E-PoS 80mm Thermal Printer "
      Top             =   2160
      Width           =   3735
   End
   Begin VB.CommandButton cmdReceive 
      Appearance      =   0  'Flat
      Caption         =   "Receive"
      Default         =   -1  'True
      Height          =   525
      Left            =   3720
      TabIndex        =   5
      ToolTipText     =   "Click to receive the milk"
      Top             =   2760
      Width           =   1335
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
      Left            =   2520
      TabIndex        =   4
      Top             =   600
      Width           =   1095
   End
   Begin VB.CheckBox chkPrint 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "Print Receipt"
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   0
      TabIndex        =   3
      Top             =   1800
      Value           =   1  'Checked
      Width           =   2295
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
      TabIndex        =   2
      Top             =   1680
      Width           =   3255
   End
   Begin VB.ComboBox cbolocation 
      Height          =   315
      ItemData        =   "frmbranchmilk.frx":002C
      Left            =   4800
      List            =   "frmbranchmilk.frx":002E
      TabIndex        =   1
      Top             =   360
      Width           =   1815
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   6315
      Width           =   8460
      _ExtentX        =   14923
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
            Picture         =   "frmbranchmilk.frx":0030
            Text            =   "DATE : 07/12/2009"
            TextSave        =   "DATE : 07/12/2009"
            Object.ToolTipText     =   "Today's Date"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "2:04 PM"
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
      Left            =   6720
      TabIndex        =   6
      Top             =   360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      MouseIcon       =   "frmbranchmilk.frx":01C4
      CalendarBackColor=   8454016
      Format          =   149028865
      CurrentDate     =   40095
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
      Left            =   4800
      TabIndex        =   21
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Branch Name:"
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
      Index           =   1
      Left            =   4800
      TabIndex        =   13
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label lblDTotal 
      AutoSize        =   -1  'True
      BackColor       =   &H00004040&
      Caption         =   "0"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00;(#,##0.00)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   2520
      TabIndex        =   12
      Top             =   1080
      Width           =   270
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Today's Total (Kgs)"
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
      Left            =   0
      TabIndex        =   11
      Top             =   1200
      Width           =   2415
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
      Left            =   6720
      TabIndex        =   10
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Actual (Kgs)"
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
      Left            =   0
      TabIndex        =   9
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Branch Intake(kgs)"
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
      Index           =   0
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   2415
   End
   Begin VB.Label lblDTotalb 
      AutoSize        =   -1  'True
      BackColor       =   &H00004040&
      Caption         =   "0"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00;(#,##0.00)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   2520
      TabIndex        =   7
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "frmbranchmilk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs1 As New ADODB.Recordset, rs2 As New ADODB.Recordset, rs3 As New ADODB.Recordset
Dim rs4 As New ADODB.Recordset, CummulKgs As Double, TRANSPORTER As String
Dim Transport As Currency, agrovet As Currency, AI As Currency, TMShares As Currency, FSA As Currency, HShares As Currency, Advance As Currency, Others As Currency

Private Sub cbolocation_Change()
    lblDTotalb = 0
    Startdate = DateSerial(year(DTPMilkDate), month(DTPMilkDate), Day(DTPMilkDate))
    Enddate = DateSerial(year(DTPMilkDate), month(DTPMilkDate), Day(DTPMilkDate) + 1)
    sql = "d_sp_DailyTotal3 '" & Startdate & "','" & Enddate & "','" & cbolocation & "'"
    'sql = "d_sp_DailyTotal3 '" & DTPMilkDate & "','" & cbolocation & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
    If Not IsNull(rs.Fields(0)) Then lblDTotalb.Caption = rs.Fields(0)
    Else
    lblDTotalb.Caption = "0"
    End If

End Sub

Private Sub cbolocation_Click()
    lblDTotalb = 0
    Startdate = DateSerial(year(DTPMilkDate), month(DTPMilkDate), Day(DTPMilkDate))
    Enddate = DateSerial(year(DTPMilkDate), month(DTPMilkDate), Day(DTPMilkDate) + 1)
    sql = "d_sp_DailyTotal3 '" & Startdate & "','" & Enddate & "','" & cbolocation & "'"
    'sql = "d_sp_DailyTotal3 '" & DTPMilkDate & "','" & cbolocation & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
    If Not IsNull(rs.Fields(0)) Then lblDTotalb.Caption = rs.Fields(0)
    Else
    lblDTotalb.Caption = "0"
    End If
End Sub
Private Sub Text1_GotFocus()

End Sub

Private Sub chprint_Click()
ports.Clear
ports = ""
'//If the drivers are installed it won't matter whether the Port is indicated
' or not it will just work.

If chprint.value = vbChecked Then
ports.AddItem "LPT1"
ports = "LPT1"
ports.AddItem "LPT2"
ports.AddItem "LPT3"
ports.AddItem "LPT4"
ports.AddItem "LPT5"
Else
'Share the printer first the use of 127.0.0.1 which is
'standard IP address for a loopback network connection
'instead of getting the computer name or IP Address
'
Dim prnPrinter As Printer
Dim pr As String
ports.Clear

For Each prnPrinter In Printers
   If InStr(prnPrinter.DeviceName, "\\") Then
    ports.AddItem prnPrinter.DeviceName
    If InStr(prnPrinter.DeviceName, "G") Then
    ports.Text = prnPrinter.DeviceName
    End If
    Else
    ports.AddItem "\\127.0.0.1\" & prnPrinter.DeviceName
    If InStr(prnPrinter.DeviceName, "G") Then
    ports.Text = "\\127.0.0.1\" & prnPrinter.DeviceName
    End If
    End If
   
   
Next
End If
End Sub

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdnewbra_Click()
frmBranch.Show vbModal
End Sub

Private Sub cmdReceive_Click()
On Error GoTo ErrorHandler
Dim Price As Currency
Dim Startdate, CummulKgs, TRANSPORTER As String
Dim transdate As Date, anss As String
'check the user
 sql = "SELECT     UserLoginIDs, UserGroup, SUPERUSER,status From UserAccounts where UserLoginIDs='" & User & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
    If rs!SuperUser <> 1 Then
    MsgBox "You are not allowed to make complaint", vbInformation
     
    Exit Sub
   End If
     End If

If cbovb = "" Then
   MsgBox "PLEASE SELECT THE VEHICLE."
    cbovb.SetFocus
   Exit Sub
 End If


If Trim(cbolocation) = "" Then
    MsgBox "PLEASE SELECT THE STATION."
        cbolocation.SetFocus
    Exit Sub
End If
If Trim(txtQnty) = "" Then
    MsgBox "Please enter the quantity supplied From the Branch."
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
sql = "set dateformat dmy SELECT * From d_MilkBranch where Branch ='" & cbolocation & "' and Date ='" & DTPMilkDate & "' and Vehicle='" & cbovb & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If rs.EOF Then
sql = ""
sql = "set dateformat dmy INSERT INTO d_MilkBranch"
sql = sql & " (Branch, Quantity, Date, Actual, Variance,auditdatetime,Vehicle)"
sql = sql & " VALUES ('" & cbolocation & "','" & lblDTotalb & "','" & DTPMilkDate & "'," & txtQnty & ",'" & txtQnty.Text - lblDTotalb & "','" & Now & "','" & cbovb & "')"
oSaccoMaster.Execute (sql)
Else
sql = ""
sql = "SET dateformat DMY Update  d_MilkBranch SET Quantity='" & lblDTotalb & "', Actual='" & rs.Fields(3) + txtQnty & "',Variance='" & (rs.Fields(3) + txtQnty) - lblDTotalb & "' WHERE Branch ='" & cbolocation & "' AND Date ='" & DTPMilkDate & "'and Vehicle='" & cbovb & "'"
oSaccoMaster.Execute (sql)
End If

listmilk

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
    txtFile.WriteLine "  Milk Receipt for the Branch"
    txtFile.WriteLine "----------------------------------"
        
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
    txtFile.WriteLine "Branch Name :" & cbolocation
    txtFile.WriteLine "Quantity Supplied :" & txtQnty & " Kgs"
    'txtFile.WriteLine "Price" & Price & " Per Kgs"
    Set rs = New ADODB.Recordset
    Dim CummulKgs1 As Double
    sql = ""
    'sql = "set dateformat dmy select sum(Quantity),sum(Actual) from d_MilkBranch where Branch='" & cbolocation & "' and Date>='" & Startdate & "' and Date<='" & DTPMilkDate & "' "
    sql = "d_sp_TotalMonth1 '" & cbolocation & "','" & Startdate & "','" & DTPMilkDate & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
    If Not IsNull(rs.Fields(0)) Then CummulKgs = rs.Fields(0)
     If Not IsNull(rs.Fields(1)) Then
      CummulKgs1 = rs.Fields(1)
     Else
      CummulKgs1 = "0.00"
     End If
    Else
    CummulKgs = "0.00"
    End If
    Dim toty As Double
    toty = CummulKgs1 - CummulKgs
    txtFile.WriteLine "Cummulative This Month : " & Format(CummulKgs, "#,##0.00" & " Kgs")
    txtFile.WriteLine "Cummulative Actual This Month : " & Format(CummulKgs1, "#,##0.00" & " Kgs")
    txtFile.WriteLine "---------------------------------------"
    txtFile.WriteLine "Varriance This Month : " & Format(toty, "#,##0.00" & " Kgs")
    txtFile.WriteLine "---------------------------------------"
    txtFile.WriteLine "Receipt Number :" & RNumber
  '  txtFile.WriteLine "TRANSPORTER :" & TRANSPORTER cbovb
    txtFile.WriteLine "Received by Vehicle No :" & username
    txtFile.WriteLine "---------------------------------------"
        txtFile.WriteLine "Transported by :" & cbovb
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

'txtSNo = ""
txtQnty = ""
'cbovb = ""
'txtSNo_Validate True
txtQnty.SetFocus
Exit Sub
ErrorHandler:

MsgBox err.description

End Sub
Public Sub listmilk()
'/// to list view//////////
sql = ""
sql = "set dateformat dmy SELECT Branch, Quantity,Actual, Variance,Vehicle From d_MilkBranch where Date ='" & DTPMilkDate & "'"
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
                    If Not IsNull(rs.Fields(4)) Then li.SubItems(4) = rs.Fields(4) & ""
'                    If Not IsNull(rs.Fields(5)) Then li.SubItems(5) = rs.Fields(5) & ""
rs.MoveNext

Wend

'////// end of view
End Sub

Public Sub loadMilk()
    Startdate = DateSerial(year(DTPMilkDate), month(DTPMilkDate), Day(DTPMilkDate))
    Enddate = DateSerial(year(DTPMilkDate), month(DTPMilkDate), Day(DTPMilkDate) + 1)
    lblDTotal.Caption = "0"
'   lbltoday.Caption = "0"
    Set rs = New ADODB.Recordset
    'sql = "d_sp_DailyTotal '" & DTPMilkDate & "'"
    sql = "d_sp_DailyTotal '" & Startdate & "','" & Enddate & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
    If Not IsNull(rs.Fields(0)) Then lblDTotal.Caption = rs.Fields(0)
    Else
    lblDTotal.Caption = "0"
'    lbltoday.Caption = "0"
    End If
    Set rs = New ADODB.Recordset
    sql = "d_sp_DailyTotal3 '" & Startdate & "','" & Enddate & "','" & cbolocation & "'"
    'sql = "d_sp_DailyTotal3 '" & DTPMilkDate & "','" & cbolocation & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
    If Not IsNull(rs.Fields(0)) Then lblDTotalb.Caption = rs.Fields(0)
    Else
    lblDTotalb.Caption = "0"
    End If
    
   
    rs.Close
    
    Set rs = Nothing
    '/// to list view//////////
sql = ""
sql = "set dateformat dmy SELECT Branch, Quantity,Actual, Variance,Vehicle From d_MilkBranch where Date ='" & DTPMilkDate & "'"
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
                    If Not IsNull(rs.Fields(4)) Then li.SubItems(4) = rs.Fields(4) & ""
'                    If Not IsNull(rs.Fields(5)) Then li.SubItems(5) = rs.Fields(5) & ""
rs.MoveNext

Wend

'////// end of view
    

End Sub

Private Sub cmdreport_Click()
Dim ans As String
ans = MsgBox("Do you Want Report per Branch ?", vbYesNo)
If ans = vbYes Then
 reportname = "d_BranchInvoice.rpt"
Else
reportname = "d_BranchInvoice2.rpt"
End If
Show_Sales_Crystal_Report "", reportname, ""
End Sub
Private Sub cmdvhreport_Click()
reportname = "monthlyvehc.rpt"
Show_Sales_Crystal_Report "", reportname, ""
End Sub

Private Sub DTPMilkDate_Change()
loadMilk
listmilk
End Sub

Private Sub DTPMilkDate_Click()
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
    
    Set rst = New Recordset
    Dim cn As Connection
    Set cn = New ADODB.Connection
    Set cn = CreateObject("adodb.connection")
    Provider = "MAZIWA"
    'Provider = cnn
    cn.Open Provider, "atm", "atm"
    Set rst = New Recordset
    
    sql = "Select Bname from   d_Branch"
    rst.Open sql, cn, adOpenKeyset, adLockOptimistic
    While Not rst.EOF
    cbolocation.AddItem rst.Fields(0)
    rst.MoveNext
    Wend
brnch
loadMilk
End Sub
Private Sub brnch()
    Set rst = New Recordset
    Dim cn As Connection
    Set cn = New ADODB.Connection
    Set cn = CreateObject("adodb.connection")
    Provider = "MAZIWA"
    'Provider = cnn
    cn.Open Provider, "atm", "atm"
    Set rst = New Recordset
    sql = "SET DATEFORMAT DMY Select Branch,Actual,Vehicle from d_MilkBranch where Date='" & DTPMilkDate & "'"
    rst.Open sql, cn, adOpenKeyset, adLockOptimistic
    While Not rst.EOF
    Startdate = DateSerial(year(DTPMilkDate), month(DTPMilkDate), Day(DTPMilkDate))
    Enddate = DateSerial(year(DTPMilkDate), month(DTPMilkDate), Day(DTPMilkDate) + 1)
    sql = "d_sp_DailyTotal3 '" & Startdate & "','" & Enddate & "','" & rst.Fields(0) & "'"
    'sql = "d_sp_DailyTotal3 '" & DTPMilkDate & "','" & rst.Fields(0) & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
    If Not IsNull(rs.Fields(0)) Then lblDTotalb.Caption = rs.Fields(0)
    Else
    lblDTotalb.Caption = "0"
    End If
sql = ""
sql = "SET dateformat DMY Update  d_MilkBranch SET Quantity='" & lblDTotalb & "',Variance='" & (rst.Fields(1)) - lblDTotalb & "' WHERE Branch ='" & rst.Fields(0) & "' AND Date ='" & DTPMilkDate & "'and Vehicle='" & rst.Fields(2) & "'"
oSaccoMaster.Execute (sql)
  rst.MoveNext
   Wend
End Sub
Private Sub NAMES3()
'Private Sub SSTab1_DblClick()
    cbovb.Clear
    Set rst = New Recordset
    Dim cn As Connection
    Set cn = New ADODB.Connection
    Set cn = CreateObject("adodb.connection")
    Provider = "MAZIWA"
    'Provider = cnn
    cn.Open Provider, "atm", "atm"
    Set rst = New Recordset
    sql = "Select Vehicle from d_VehicleTill order by Vehicle"
    'Select distinct(Locations) from   d_Debtors
    rst.Open sql, cn, adOpenKeyset, adLockOptimistic
    While Not rst.EOF
    cbovb.AddItem rst.Fields(0)
    rst.MoveNext
    Wend
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Startdate = DateSerial(year(DTPMilkDate), month(DTPMilkDate), 1)
Enddate = DateSerial(year(DTPMilkDate), month(DTPMilkDate) + 1, 1 - 1)
MainForm.Caption = "EasyMa "
End Sub

Private Sub ListView2_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub txtQnty_Validate(Cancel As Boolean)
txtQnty = Format(txtQnty, "####0.00")
End Sub
