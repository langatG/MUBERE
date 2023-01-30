VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmbranchdisp 
   Caption         =   "Milk Dispatch Form"
   ClientHeight    =   7350
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8670
   BeginProperty Font 
      Name            =   "Sylfaen"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7350
   ScaleWidth      =   8670
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView ListView3 
      Height          =   2535
      Left            =   240
      TabIndex        =   22
      Top             =   4440
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Sylfaen"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
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
   End
   Begin VB.ComboBox cboNames 
      Height          =   390
      Left            =   5640
      TabIndex        =   21
      Top             =   1800
      Width           =   2535
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
      ItemData        =   "frmbranchdisp.frx":0000
      Left            =   4200
      List            =   "frmbranchdisp.frx":0010
      TabIndex        =   20
      Text            =   "\\127.0.0.1\E-PoS 80mm Thermal Printer "
      Top             =   2880
      Width           =   4335
   End
   Begin VB.CheckBox chkPrint 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "Print Receipt"
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   0
      TabIndex        =   19
      Top             =   2880
      Width           =   2295
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
      TabIndex        =   11
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CheckBox chkbranch 
      Caption         =   "Per Branch"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5280
      TabIndex        =   9
      Top             =   480
      Width           =   1215
   End
   Begin VB.CheckBox chkall 
      Caption         =   "All Branches"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3960
      TabIndex        =   8
      Top             =   480
      Width           =   1335
   End
   Begin VB.ComboBox cbolocation 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmbranchdisp.frx":002C
      Left            =   4320
      List            =   "frmbranchdisp.frx":002E
      TabIndex        =   6
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Today's collection"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      TabIndex        =   3
      Top             =   4200
      Width           =   8415
   End
   Begin VB.CommandButton cmdReceive 
      Appearance      =   0  'Flat
      Caption         =   "Receive"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   240
      TabIndex        =   2
      ToolTipText     =   "Click to receive the milk"
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   1
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton cmdreport 
      Caption         =   "Branch Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   3600
      Width           =   1335
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   7095
      Width           =   8670
      _ExtentX        =   15293
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
            Picture         =   "frmbranchdisp.frx":0030
            Text            =   "DATE : 07/12/2009"
            TextSave        =   "DATE : 07/12/2009"
            Object.ToolTipText     =   "Today's Date"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "4:56 PM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
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
   Begin MSComCtl2.DTPicker DTPMilkDate 
      Height          =   375
      Left            =   6720
      TabIndex        =   10
      Top             =   600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      MouseIcon       =   "frmbranchdisp.frx":01C4
      CalendarBackColor=   8454016
      Format          =   178061313
      CurrentDate     =   40095
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Debtor Name"
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
      Left            =   3840
      TabIndex        =   18
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Dispatch Milk For"
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
      Index           =   1
      Left            =   4080
      TabIndex        =   17
      Top             =   120
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
      TabIndex        =   16
      Top             =   840
      Width           =   375
   End
   Begin VB.Label lbltg 
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
      TabIndex        =   15
      Top             =   840
      Width           =   2415
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
      Index           =   1
      Left            =   0
      TabIndex        =   14
      Top             =   1440
      Width           =   2415
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
      Index           =   0
      Left            =   0
      TabIndex        =   13
      Top             =   120
      Width           =   2415
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
      Top             =   120
      Width           =   270
   End
   Begin VB.Label Label12 
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
      Left            =   4320
      TabIndex        =   7
      Top             =   960
      Width           =   1815
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
      TabIndex        =   5
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmbranchdisp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbolocation_Change()
    lblDTotalb = 0
    Startdate = DateSerial(year(DTPMilkDate), month(DTPMilkDate), Day(DTPMilkDate))
    Enddate = DateSerial(year(DTPMilkDate), month(DTPMilkDate), Day(DTPMilkDate) + 1)
    If cbolocation = "MASIMBA" Then
    sql = "set  dateformat dmy SELECT SUM(QSupplied) From d_Milkintake where TransDate>='" & Startdate & "' and TransDate<'" & Enddate & "' and LOCATION NOT LIKE'%SULT%'"
    Else
    sql = "d_sp_DailyTotal3 '" & Startdate & "','" & Enddate & "','" & cbolocation & "'"
    End If
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
    If cbolocation = "MASIMBA" Then
    sql = "set  dateformat dmy SELECT SUM(QSupplied) From d_Milkintake where TransDate>='" & Startdate & "' and TransDate<'" & Enddate & "' and LOCATION NOT LIKE'%SULT%'"
    Else
    sql = "d_sp_DailyTotal3 '" & Startdate & "','" & Enddate & "','" & cbolocation & "'"
    End If
    
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
    If Not IsNull(rs.Fields(0)) Then lblDTotalb.Caption = rs.Fields(0)
    Else
    lblDTotalb.Caption = "0"
    End If
End Sub

Private Sub chkall_Click()
If chkall.value = 1 Then
chkbranch.value = 0
'lblDTotalb = lblDTotal
Else
chkbranch.value = 1
'lblDTotalb = 0
End If
chkbranch_Click
End Sub

Private Sub chkbranch_Click()
If chkbranch.value = 1 Then
'lbltg.Visible = True
cbolocation.Visible = True
lblDTotalb.Visible = True
chkall.value = vbUnchecked
Branch
'Label12.Visible = True
Else
'Label1.Visible = False
cbolocation.Visible = False
lblDTotalb.Visible = False
chkall.value = vbChecked
'Label12.Visible = False
End If
End Sub

Private Sub cmdclose_Click()
Unload Me
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
If Trim(cboNames) = "" Then
    MsgBox "Please select the Debtor."
        cboNames.SetFocus
    Exit Sub
End If
If chkall.value = 0 Then
If Trim(cbolocation) = "" Then
    MsgBox "PLEASE SELECT THE STATION."
        cbolocation.SetFocus
    Exit Sub
End If
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
'Dim bran As String
sql = ""
If chkall.value = 1 Then
 lblDTotalb = lblDTotal
 sql = "set dateformat dmy SELECT * From d_MilkBranch where Date ='" & DTPMilkDate & "' and Vehicle='" & cboNames & "'"
Else
 sql = "set dateformat dmy SELECT * From d_MilkBranch where Branch ='" & cbolocation & "' and Date ='" & DTPMilkDate & "' and Vehicle='" & cboNames & "'"
End If
'sql = ""
'sql = "set dateformat dmy SELECT * From d_MilkBranch where Branch ='" & cbolocation & "' and Date ='" & DTPMilkDate & "' and Vehicle='" & cboNames & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If rs.EOF Then
sql = ""
 If chkall.value = 1 Then
  sql = "set dateformat dmy INSERT INTO d_MilkBranch"
  sql = sql & " (Branch, Quantity, Date, Actual, Variance,auditdatetime,Vehicle)"
  sql = sql & " VALUES ('MASIMBA','" & lblDTotalb & "','" & DTPMilkDate & "'," & txtQnty & ",'" & txtQnty.Text - lblDTotalb & "','" & Now & "','" & cboNames & "')"
 Else
  sql = "set dateformat dmy INSERT INTO d_MilkBranch"
  sql = sql & " (Branch, Quantity, Date, Actual, Variance,auditdatetime,Vehicle)"
  sql = sql & " VALUES ('" & cbolocation & "','" & lblDTotalb & "','" & DTPMilkDate & "'," & txtQnty & ",'" & txtQnty.Text - lblDTotalb & "','" & Now & "','" & cboNames & "')"
 End If
oSaccoMaster.Execute (sql)
Else
 If chkall.value = 1 Then
 sql = ""
 sql = "SET dateformat DMY Update  d_MilkBranch SET Quantity='" & lblDTotalb & "', Actual='" & rs.Fields(3) + txtQnty & "',Variance='" & (rs.Fields(3) + txtQnty) - lblDTotalb & "' WHERE Branch ='MASIMBA' AND Date ='" & DTPMilkDate & "'and Vehicle='" & cboNames & "'"
 Else
 sql = ""
 sql = "SET dateformat DMY Update  d_MilkBranch SET Quantity='" & lblDTotalb & "', Actual='" & rs.Fields(3) + txtQnty & "',Variance='" & (rs.Fields(3) + txtQnty) - lblDTotalb & "' WHERE Branch ='" & cbolocation & "' AND Date ='" & DTPMilkDate & "'and Vehicle='" & cboNames & "'"
 End If
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
    txtFile.WriteLine "             Milk Dispatch Receipt"
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
    txtFile.WriteLine "Name :" & cbolocation
    txtFile.WriteLine "Quantity Supplied :" & txtQnty & " Kgs"
    'txtFile.WriteLine "Price" & Price & " Per Kgs"
    Set rs = New ADODB.Recordset
    If chkall.value = 1 Then
    sql = "d_sp_TotalMonth1 'MASIMBA','" & Startdate & "','" & DTPMilkDate & "'"
    Else
    sql = "d_sp_TotalMonth1 " & cbolocation & ",'" & Startdate & "','" & DTPMilkDate & "'"
    End If
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
    If Not IsNull(rs.Fields(0)) Then CummulKgs = rs.Fields(0)
    Else
    CummulKgs = "0.00"
    End If
    txtFile.WriteLine "Cummulative This Month : " & Format(CummulKgs, "#,##0.00" & " Kgs")

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


'loadMilk

'txtSNo = ""
txtQnty = ""
lblDTotalb = ""
'txtSNo_Validate True
txtQnty.SetFocus
Exit Sub
ErrorHandler:

MsgBox err.description
End Sub
Public Sub loadMilk()
dailykgs
''    Set rs = New ADODB.Recordset
''    Startdate = DateSerial(year(DTPMilkDate), month(DTPMilkDate), Day(DTPMilkDate))
''    Enddate = DateSerial(year(DTPMilkDate), month(DTPMilkDate), Day(DTPMilkDate) + 1)
''    sql = "d_sp_DailyTotal3 '" & Startdate & "','" & Enddate & "','" & cbolocation & "'"
''    'sql = "d_sp_DailyTotal3 '" & DTPMilkDate & "','" & cbolocation & "'"
''    Set rs = oSaccoMaster.GetRecordset(sql)
''    If Not rs.EOF Then
''    If Not IsNull(rs.Fields(0)) Then lblDTotalb.Caption = rs.Fields(0)
''    Else
''    lblDTotalb.Caption = "0"
''    End If
''
''
''    rs.Close
''
''    Set rs = Nothing
''    '/// to list view//////////
''sql = ""
''sql = "set dateformat dmy SELECT Branch, Quantity,Actual, Variance,Vehicle From d_MilkBranch where Date ='" & DTPMilkDate & "'"
'''sql = "set dateformat dmy SELECT d.DCode, d.DName, m.DispQnty,m.DispDate FROM  d_Debtors AS d INNER JOIN d_MilkControl AS m ON d.DCode = m.dcode WHERE     (DispDate = '" & DTPDispatchDate & "') and status=0"
''Set rs = oSaccoMaster.GetRecordset(sql)
''If rs.EOF Then
''Exit Sub
''End If
''ListView3.ListItems.Clear
''While Not rs.EOF
''If Not IsNull(rs.Fields(0)) Then
''Set li = ListView3.ListItems.Add(, , rs.Fields(0))
''End If
''                    If Not IsNull(rs.Fields(1)) Then li.SubItems(1) = rs.Fields(1) & ""
''                    If Not IsNull(rs.Fields(2)) Then li.SubItems(2) = rs.Fields(2) & ""
''                    If Not IsNull(rs.Fields(3)) Then li.SubItems(3) = rs.Fields(3) & ""
''                    If Not IsNull(rs.Fields(4)) Then li.SubItems(4) = rs.Fields(4) & ""
'''                    If Not IsNull(rs.Fields(5)) Then li.SubItems(5) = rs.Fields(5) & ""
''rs.MoveNext
''
''Wend

'////// end of view
    

End Sub

Private Sub cmdreport_Click()
reportname = "d_BranchInvoice3.rpt"
Show_Sales_Crystal_Report "", reportname, ""
End Sub
Private Sub Branch()
'Private Sub SSTab1_DblClick()
    cbolocation.Clear
    Set rst = New Recordset
    Dim cn As Connection
    Set cn = New ADODB.Connection
    Provider = "MAZIWA"
    'Provider = cnn
    cn.Open Provider
    Set rst = New Recordset
    sql = "Select BName from d_Branch where BName like'%MAS%' or BName like'%SUL%' order by BName asc"
    'Select distinct(Locations) from   d_Debtors
    rst.Open sql, cn, adOpenKeyset, adLockOptimistic
    While Not rst.EOF
    cbolocation.AddItem rst.Fields(0)
    rst.MoveNext
    Wend
End Sub
Private Sub dailykgs()
    Startdate = DateSerial(year(DTPMilkDate), month(DTPMilkDate), Day(DTPMilkDate))
    Enddate = DateSerial(year(DTPMilkDate), month(DTPMilkDate), Day(DTPMilkDate) + 1)
    lblDTotal.Caption = "0"
'   lbltoday.Caption = "0"
    Set rs = New ADODB.Recordset
    sql = "d_sp_DailyTotal '" & Startdate & "','" & Enddate & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
    If Not IsNull(rs.Fields(0)) Then lblDTotal.Caption = rs.Fields(0)
    Else
    lblDTotal.Caption = "0"
'    lbltoday.Caption = "0"
    End If
End Sub
Private Sub NAMES()
'Private Sub SSTab1_DblClick()
    cboNames.Clear
    Set rst = New Recordset
    Dim cn As Connection
    Set cn = New ADODB.Connection
    Provider = "MAZIWA"
    'Provider = cnn
    cn.Open Provider
    Set rst = New Recordset
    sql = "Select DName from   d_Debtors order by DName"
    'Select distinct(Locations) from   d_Debtors
    rst.Open sql, cn, adOpenKeyset, adLockOptimistic
    While Not rst.EOF
    cboNames.AddItem rst.Fields(0)
    rst.MoveNext
    Wend
End Sub
Public Sub listmilk()
'/// to list view//////////
dailykgs
sql = ""
sql = "set dateformat dmy SELECT Branch, Quantity,Actual, Variance From d_MilkBranch where Date ='" & DTPMilkDate & "'"
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

Private Sub DTPMilkDate_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
listmilk
End Sub
Private Sub DTPMilkDate_Change()
listmilk
End Sub
Private Sub DTPMilkDate_Click()
listmilk
End Sub
Private Sub Form_Load()
DTPMilkDate = Format(Get_Server_Date, "dd/mm/yyyy")
DTPMilkDate.MaxDate = Format(Get_Server_Date, "dd/mm/yyyy")
'DTPcomplaintperiod1 = DTPMilkDate
With StatusBar1.Panels
    .Item(1).Text = "USER : " & username
    .Item(2).Text = "DATE : " & Format(Get_Server_Date, "dd/mm/yyyy")

End With
chkbranch.value = 0
chkbranch.value = 1
loadMilk
dailykgs
NAMES
End Sub

