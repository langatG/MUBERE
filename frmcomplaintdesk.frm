VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmcomplaintdesk 
   Caption         =   "COMPLAIN DESK"
   ClientHeight    =   8955
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15495
   LinkTopic       =   "Form1"
   ScaleHeight     =   8955
   ScaleWidth      =   15495
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkPrices 
      BackColor       =   &H00FFFFC0&
      Caption         =   "With Different price per day?"
      Enabled         =   0   'False
      Height          =   315
      Left            =   4560
      TabIndex        =   45
      Top             =   480
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Outlet Sales"
      Height          =   495
      Left            =   8760
      TabIndex        =   44
      Top             =   8160
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdDays1 
      Caption         =   "Day's Intake"
      Height          =   495
      Left            =   8760
      TabIndex        =   43
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton cmdsumaintake2 
      Caption         =   "Daily Summary"
      Height          =   495
      Left            =   7200
      TabIndex        =   42
      Top             =   2400
      Width           =   1455
   End
   Begin VB.TextBox txttoday 
      Height          =   375
      Left            =   10920
      Locked          =   -1  'True
      TabIndex        =   37
      Top             =   3720
      Width           =   1095
   End
   Begin VB.ComboBox cbolocation 
      Height          =   315
      ItemData        =   "frmcomplaintdesk.frx":0000
      Left            =   6480
      List            =   "frmcomplaintdesk.frx":0002
      TabIndex        =   36
      Top             =   120
      Width           =   2052
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
      ItemData        =   "frmcomplaintdesk.frx":0004
      Left            =   10200
      List            =   "frmcomplaintdesk.frx":0014
      TabIndex        =   35
      Text            =   "\\127.0.0.1\E-PoS 80mm Thermal Printer "
      Top             =   720
      Width           =   2175
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
      Left            =   9120
      TabIndex        =   34
      Top             =   120
      Width           =   3255
   End
   Begin VB.CommandButton cmdreprintreceipt 
      Caption         =   "Reprint"
      Height          =   405
      Left            =   8760
      TabIndex        =   33
      ToolTipText     =   "Click "
      Top             =   1920
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker DTPcomplaintperiod 
      Height          =   255
      Left            =   10680
      TabIndex        =   32
      Top             =   2400
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
      _Version        =   393216
      Format          =   115015681
      CurrentDate     =   40556
   End
   Begin VB.CommandButton cmdcomplaintreport 
      Caption         =   "Complain Report"
      Height          =   375
      Left            =   10800
      TabIndex        =   30
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton cmdupdatedeductions 
      Caption         =   "Update Deductions"
      Height          =   495
      Left            =   6360
      TabIndex        =   29
      Top             =   8160
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "List of Deductions"
      Height          =   2295
      Left            =   120
      TabIndex        =   27
      Top             =   5880
      Width           =   12255
      Begin MSComctlLib.ListView Lvwdeductions 
         Height          =   1815
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   3201
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   12582912
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Sylfaen"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin VB.TextBox txtComment 
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
      Left            =   5400
      TabIndex        =   19
      Text            =   "<Put your comment here>"
      Top             =   1440
      Width           =   5415
   End
   Begin VB.CheckBox chkComment 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "Add Comment"
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   4440
      TabIndex        =   18
      Top             =   1920
      Value           =   1  'Checked
      Width           =   2535
   End
   Begin VB.CheckBox chkPrint 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "Print Receipt"
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   120
      TabIndex        =   17
      Top             =   1920
      Width           =   2295
   End
   Begin VB.TextBox txtSNo 
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
      Left            =   2400
      TabIndex        =   15
      Top             =   120
      Width           =   975
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
      Left            =   3240
      TabIndex        =   14
      Top             =   720
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Transporter's Receipt"
      Height          =   855
      Left            =   120
      TabIndex        =   9
      Top             =   2880
      Width           =   10335
      Begin VB.TextBox txtTCode 
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
         Height          =   405
         Left            =   2280
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         Height          =   405
         Left            =   8160
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Transporter code"
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
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lblTName 
         AutoSize        =   -1  'True
         BackColor       =   &H0080FF80&
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
         Left            =   3360
         TabIndex        =   12
         Top             =   240
         Width           =   75
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Today's collection"
      Height          =   1695
      Left            =   120
      TabIndex        =   7
      Top             =   4080
      Width           =   12255
      Begin MSComctlLib.ListView lvwMilkCollected 
         Height          =   1335
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   2355
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   12582912
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Sylfaen"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin VB.CommandButton cmdReceive 
      Caption         =   "Receive"
      Default         =   -1  'True
      Height          =   405
      Left            =   7200
      TabIndex        =   6
      ToolTipText     =   "Click to receive the milk"
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   405
      Left            =   11040
      TabIndex        =   5
      ToolTipText     =   "Click "
      Top             =   1320
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   525
      Left            =   11040
      TabIndex        =   4
      Top             =   8160
      Width           =   1095
   End
   Begin VB.CheckBox chkNotepad 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "To Notepad"
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   2400
      TabIndex        =   3
      Top             =   1920
      Width           =   1935
   End
   Begin VB.CommandButton cmdfind 
      Caption         =   "F"
      Height          =   405
      Left            =   3480
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdget 
      Caption         =   "Get DD"
      Height          =   405
      Left            =   12600
      TabIndex        =   0
      ToolTipText     =   "Click to receive the milk"
      Top             =   120
      Width           =   1335
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   8700
      Width           =   15495
      _ExtentX        =   27331
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
            Picture         =   "frmcomplaintdesk.frx":0030
            Text            =   "DATE : 07/12/2009"
            TextSave        =   "DATE : 07/12/2009"
            Object.ToolTipText     =   "Today's Date"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "09:26 PM"
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
      Left            =   4920
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FileName        =   "c:\receipt.txt"
   End
   Begin MSComCtl2.DTPicker DTPMilkDate 
      Height          =   375
      Left            =   7680
      TabIndex        =   16
      Top             =   840
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      MouseIcon       =   "frmcomplaintdesk.frx":01C4
      CalendarBackColor=   8454016
      Format          =   115015681
      CurrentDate     =   40095
   End
   Begin VB.Label Label9 
      Caption         =   "Branch:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5400
      TabIndex        =   41
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label7 
      Caption         =   "Branch intake"
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
      Left            =   120
      TabIndex        =   40
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Branch Total (Kgs)"
      Height          =   375
      Left            =   2520
      TabIndex        =   39
      Top             =   1920
      Width           =   2775
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
      Left            =   2640
      TabIndex        =   38
      Top             =   2280
      Width           =   270
   End
   Begin VB.Label Label6 
      Caption         =   "Complaint Period"
      Height          =   375
      Left            =   10560
      TabIndex        =   31
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Member Number"
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
      Left            =   120
      TabIndex        =   26
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Quantity Supplied (Kgs)"
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
      TabIndex        =   25
      Top             =   720
      Width           =   3015
   End
   Begin VB.Label lblComment 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Reason"
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
      Left            =   3480
      TabIndex        =   24
      Top             =   1320
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Milk Date"
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
      Left            =   6120
      TabIndex        =   23
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label lblNames 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
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
      Left            =   3960
      TabIndex        =   22
      Top             =   120
      Width           =   315
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
      Left            =   120
      TabIndex        =   21
      Top             =   1320
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
      Left            =   2640
      TabIndex        =   20
      Top             =   1200
      Width           =   270
   End
End
Attribute VB_Name = "frmcomplaintdesk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs1 As New ADODB.Recordset, rs2 As New ADODB.Recordset, rs3 As New ADODB.Recordset
Dim rs4 As New ADODB.Recordset, CummulKgs As Double, TRANSPORTER As String
Dim Transport As Currency, agrovet As Currency, BONUS As Currency, TMShares As Currency, FSA As Currency, HShares As Currency, Advance As Currency, Others As Currency

Private Sub cbolocation_Change()
lblDTotalb = "0"
Startdate = DateSerial(Year(DTPMilkDate), month(DTPMilkDate), Day(DTPMilkDate))
Enddate = DateSerial(Year(DTPMilkDate), month(DTPMilkDate), Day(DTPMilkDate) + 1)
sql = "d_sp_DailyTotal3 '" & Startdate & "','" & Enddate & "','" & cbolocation & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
    If Not IsNull(rs.Fields(0)) Then lblDTotalb.Caption = rs.Fields(0)
    Else
    lblDTotalb.Caption = "0"
    End If

End Sub

Private Sub cbolocation_Click()
lblDTotalb = "0"
Startdate = DateSerial(Year(DTPMilkDate), month(DTPMilkDate), Day(DTPMilkDate))
Enddate = DateSerial(Year(DTPMilkDate), month(DTPMilkDate), Day(DTPMilkDate) + 1)
sql = "d_sp_DailyTotal3 '" & Startdate & "','" & Enddate & "','" & cbolocation & "'"
'sql = "d_sp_DailyTotal3 '" & DTPMilkDate & "','" & cbolocation & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
    If Not IsNull(rs.Fields(0)) Then lblDTotalb.Caption = rs.Fields(0)
    Else
    lblDTotalb.Caption = "0"
    End If
End Sub

Private Sub chkComment_Click()
If chkComment.value = vbChecked Then
    lblComment.Visible = True
    txtComment = "<Put your comment here>"
    txtComment.Visible = True
    txtComment.SetFocus
Else
    lblComment.Visible = False
    txtComment.Visible = False

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

Private Sub cmdcomplaintreport_Click()
On Error GoTo ErrorHandler
'//those where dates are inserted to the system days later.,
Dim a As Date
Dim AC As Date
Dim Id As Long
Dim sno As String
Dim kg As Double
sql = ""

sql = "SELECT     id,SNo, TransDate, QSupplied, AuditId, auditdatetime, LR   FROM         d_Milkintake where month(transdate)=" & month(DTPcomplaintperiod) & " and lr<>1 and year(transdate)=" & Year(DTPcomplaintperiod) & ""
Set rs = oSaccoMaster.GetRecordset(sql)


If Not rs.EOF Then
While Not rs.EOF
a = Format(rs.Fields(5), "dd/mm/yyyy")
AC = rs.Fields(2)
Id = rs.Fields(0)
kg = rs.Fields(3)
sno = rs.Fields(1)
If a <> AC Then

sql = ""
sql = "set dateformat dmy update d_milkintake set LR=1 where id =" & Id & ""
oSaccoMaster.ExecuteThis (sql)

End If
rs.MoveNext
Wend


'//put here the report
'complaintreport
reportname = "complaintreport.rpt"
 Show_Sales_Crystal_Report STRFORMULA, reportname, ""
Else
MsgBox "No Records available for the month mentioned", vbInformation
Exit Sub
End If
Exit Sub
ErrorHandler:
MsgBox err.description
End Sub

Private Sub cmdDays1_Click()
 reportname = "d_DailyIntake.rpt"
 Show_Sales_Crystal_Report "", reportname, ""
End Sub

Private Sub cmdFind_Click()
        Me.MousePointer = vbHourglass
        frmSearchSupplier.Show vbModal
        txtSNo = sel
        txtSNo_Validate True
        Me.MousePointer = 0

End Sub

Private Sub cmdget_Click()
On Error GoTo ErrorHandler
Dim Price As Currency
Dim Startdate, CummulKgs, TRANSPORTER As String
Dim transdate As Date

'//open the database of serby
Dim PROVIDER2 As String
PROVIDER2 = "SERBRY"
Dim cn2 As New ADODB.Connection
Set cn2 = New ADODB.Connection
Dim RN2 As New ADODB.Recordset
cn2.Open PROVIDER2, "", "10FLAT"
sql = ""
sql = "select [DATE],ACNO,TIME,TRANSPORTER,KG,USER FROM DELIVERY  ORDER BY TIME"
Set RN2 = New ADODB.Recordset
RN2.Open sql, cn2
If Not RN2.EOF Then
txtSNo = Replace(RN2.Fields(1), "0", "")
txtQnty = RN2.Fields(4)
User = RN2.Fields(5)
Else
GoTo HEL
End If
HEL:
txtSNo_Validate True

If lblNames.Caption = "" Then
MsgBox "Please enter a valid supplier number."
txtSNo.SetFocus
Exit Sub
End If

If Not IsNumeric(txtQnty) Then
MsgBox "Please enter a number. " & txtQnty & " is not a number", vbExclamation
txtQnty.SetFocus
Exit Sub
End If

Startdate = DateSerial(Year(DTPMilkDate), month(DTPMilkDate), 1)
Enddate = DateSerial(Year(DTPMilkDate), month(DTPMilkDate) + 1, 1 - 1)

Set rs = oSaccoMaster.GetRecordset("d_sp_IsClosed '" & Enddate & "'")
If Not rs.EOF Then
    MsgBox "The period " & Enddate & " has been closed by " & rs.Fields(0)
    Exit Sub
End If

Dim ans As String
Set rs = New ADODB.Recordset
 sql = "SET dateformat dmy SELECT d_Milkintake.[ID],d_Milkintake.SNo, d_Milkintake.QSupplied, "
 sql = sql & " d_Milkintake.TransTime , d_Suppliers.[Names] FROM  d_Milkintake INNER JOIN "
 sql = sql & " d_Suppliers ON d_Milkintake.SNo = d_Suppliers.SNo AND  d_Milkintake.SNo= " & txtSNo & " AND (d_milkintake.TransDate = '" & DTPMilkDate & "')"
 sql = sql & " ORDER BY d_Milkintake.Id DESC"

 
Set rs = oSaccoMaster.GetRecordset(sql)
If rs.RecordCount > 0 Then

With lvwMilkCollected
    
       .ListItems.Clear
    
        .ColumnHeaders.Clear

  End With

    With lvwMilkCollected
        
        
        .ColumnHeaders.Add , , "SNo", 2000
        .ColumnHeaders.Add , , "Names", 3000
        .ColumnHeaders.Add , , "QNTY"
        .ColumnHeaders.Add , , "Time"
        .ColumnHeaders.Add , , "Receipt No.", 2000
    
        While Not rs.EOF
        
        If Not IsNull(rs.Fields("SNo")) Then
        
            Set li = .ListItems.Add(, , Trim(rs.Fields("SNo")))
            End If
            If Not IsNull(rs.Fields("Names")) Then
            li.ListSubItems.Add , , Trim(rs.Fields("Names"))
            End If
            If Not IsNull(rs.Fields("QSupplied")) Then
            li.ListSubItems.Add , , Trim(rs.Fields("QSupplied"))
            End If
            If Not IsNull(rs.Fields("TransTime")) Then
            li.ListSubItems.Add , , Trim(rs.Fields("TransTime"))
            End If
            If Not IsNull(rs.Fields("ID")) Then
             li.ListSubItems.Add , , Trim(rs.Fields("ID"))
            End If

            
                    rs.MoveNext
        
        Wend
        
    End With
    
    rs.Close
    
    Set rs = Nothing
    
lvwMilkCollected.View = lvwReport
ans = MsgBox("Supplier number " & txtSNo & " had supplied milk today. Add this?", vbYesNo, "MILK REPEAT")
If ans = vbNo Then
txtSNo.SetFocus
Exit Sub
End If
End If


Set rs = New ADODB.Recordset
sql = "SELECT Price from d_Price"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
Price = rs!Price
End If



'//Update Milk Intake


    
Set cn = New ADODB.Connection
sql = "d_sp_MilkIntake " & txtSNo & ",'" & DTPMilkDate & "'," & txtQnty & "," & Price & "," & Price * CCur(txtQnty) & ",'" & Time & "','" & User & "','Intake Complain Desk'"
oSaccoMaster.ExecuteThis (sql)




Set rs = New ADODB.Recordset
    sql = "d_sp_TransActive " & txtSNo & ",'" & DTPMilkDate & "'"
        Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
    
        transdate = rs!rate
        If transdate < Startdate Then
            transdate = Startdate
        End If
        
        
        '---d_sp_SelTransDet @SNo bigint,@StartDate varchar(10), @Enddate varchar(10)
        Set rst = New ADODB.Recordset
            sql = "d_sp_SelTransDet " & txtSNo & ",'" & transdate & "','" & Enddate & "'"
        Set rst = oSaccoMaster.GetRecordset(sql)
        
    '--d_sp_UpdateDetTrans @SNo bigint,@QNTY float,@Amnt money,@Code varchar(35),@Subsidy money,@EPeriod varchar(10),@user varchar(35)
        If Not rst.EOF Then
            Set cn = New ADODB.Connection
                sql = "d_sp_UpdateDetTrans " & txtSNo & "," & rst!qnty & "," & rst!amount & ",'" & rst!code & "'," & rst!subsidy & ",'" & Enddate & "','" & User & "'"
            oSaccoMaster.ExecuteThis (sql)
            
        
        
        
        '----d_sp_SelTransGPayQnty @EP varchar(10), @Code varchar(35)
        Set rst2 = New ADODB.Recordset
            sql = "d_sp_SelTransGPayQnty '" & Enddate & "','" & rst!code & "'"
        Set rst2 = oSaccoMaster.GetRecordset(sql)
        
    '-- d_sp_UpdateTransPay @Code varchar(35), @Qnty float,@Amnt money,@Subsidy money,@GrossPay money, @EndDate varchar(10)  AS
    If Not rst2.EOF Then
            Set cn = New ADODB.Connection
                sql = "d_sp_UpdateTransPay '" & rst!code & "'," & rst2!qnty & "," & rst2!Amnt & "," & rst2!subsidy & "," & rst2!GPay & ",'" & Enddate & "','" & User & "'"
            oSaccoMaster.ExecuteThis (sql)
            
        End If
 
        '---Get transporters Total Deductions//d_sp_TotalTransDeduct @Code varchar(35),@Month bigint,@Year bigint
        Set Rs1 = New ADODB.Recordset
            sql = "d_sp_TotalTransDeduct '" & rst!code & "'," & month(DTPMilkDate) & "," & Year(DTPMilkDate) & ""
        Set Rs1 = oSaccoMaster.GetRecordset(sql)
    If Not Rs1.EOF Then
    Dim TransTotalDed As Currency
    If Not IsNull(Rs1.Fields(0)) Then TransTotalDed = Rs1.Fields(0)
    End If
    
    Set rs3 = New ADODB.Recordset
    Dim DESCR As String
    Dim amount As Currency
    
    '--d_sp_SelTransDed @Code varchar(35), @startdate varchar(10),@enddate varchar(10)
    sql = "d_sp_SelTransDed '" & rst!code & "','" & Startdate & "','" & Enddate & "'"
    Set rs3 = oSaccoMaster.GetRecordset(sql)
    
    agrovet = 0
    BONUS = 0
    TMShares = 0
    FSA = 0
    HShares = 0
    Advance = 0
    Others = 0
    TransTotalDed = 0
    
    
If Not rs3.EOF Then
    While Not rs3.EOF
    DESCR = Trim(rs3.Fields(0))
    amount = 0
    amount = rs3.Fields(1)
    sql = "SELECT     Agrovet, BONUS, TMShares, FSA, HShares, Advance, Others FROM d_TransportersPayroll WHERE Code='" & rst!code & "' AND EndPeriod ='" & Enddate & "'"
    Set rs4 = oSaccoMaster.GetRecordset(sql)
     If UCase(rs4.Fields(0).name) = UCase(DESCR) Then
        agrovet = amount
    End If
    If UCase(rs4.Fields(1).name) = UCase(DESCR) Then
        BONUS = amount
    End If
    If UCase(rs4.Fields(2).name) = UCase(DESCR) Then
        TMShares = amount
    End If
    If UCase(rs4.Fields(3).name) = UCase(DESCR) Then
        FSA = amount
    End If
    If UCase(rs4.Fields(4).name) = UCase(DESCR) Then
        HShares = amount
    End If
    If UCase(rs4.Fields(5).name) = UCase(DESCR) Then
        Advance = amount
    End If
    If UCase(rs4.Fields(6).name) = UCase(DESCR) Then
        Others = amount
    End If

    '//d_sp_UpdatePayrollDed  @SNo bigint,@EndPeriod varchar(15),@Transport money,@Agrovet money,@AI money,@TMShares money,@FSA money,@HShares money,@Advance money,@Others
    rs3.MoveNext
    Wend
    
End If
     ' d_sp_UpdateTransDed  @Code varchar(35),@EndPeriod varchar(15),@TotalDed money,@Agrovet money,@AI money,@TMShares money,@FSA money,@HShares money,@Advance money,@Others money
   
    Set cn = New ADODB.Connection
    sql = "d_sp_UpdateTransDed  '" & rst!code & "','" & Enddate & "'," & TransTotalDed & "," & agrovet & "," & BONUS & "," & TMShares & "," & FSA & "," & HShares & "," & Advance & "," & Others & ""
    oSaccoMaster.ExecuteThis (sql)
    
 
        '//Update supplier deductions
            Set cn = New ADODB.Connection
            Dim r2  As String
            r2 = "  "
                sql = "d_sp_SupplierDeduct " & txtSNo & ",'" & DTPMilkDate & "','Transport'," & rs!rate * CCur(txtQnty) & ",'" & Startdate & "','" & Enddate & "'," & Year(Enddate) & ",'" & User & "','" & r2 & "',''"
            oSaccoMaster.ExecuteThis (sql)
            



End If
'Check inactive Transporter
Set rs = New ADODB.Recordset
sql = "d_sp_TransInActive " & txtSNo & ",'" & Startdate & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then

'--Update Detailed Trasport
    transdate = rs!Startdate
        If transdate < Startdate Then
            transdate = Startdate
        End If
        
    '---d_sp_SelTransDet @SNo bigint,@StartDate varchar(10), @Enddate varchar(10)
        Set rst = New ADODB.Recordset
            sql = "d_sp_SelTransDet " & txtSNo & ",'" & transdate & "','" & rs!dateinactivate & "'"
        Set rst = oSaccoMaster.GetRecordset(sql)
        
    '--d_sp_UpdateDetTrans @SNo bigint,@QNTY float,@Amnt money,@Code varchar(35),@Subsidy money,@EPeriod varchar(10),@user varchar(35)
        If Not rst.EOF Then
            Set cn = New ADODB.Connection
                sql = "d_sp_UpdateDetTrans " & txtSNo & "," & rst!qnty & "," & rst!amount & ",'" & rst!code & "'," & rst!subsidy & ",'" & Enddate & "','" & User & "'"
            oSaccoMaster.ExecuteThis (sql)
            
            '----d_sp_SelTransGPayQnty @EP varchar(10), @Code varchar(35)
        Set rst2 = New ADODB.Recordset
            sql = "d_sp_SelTransGPayQnty '" & Enddate & "','" & rst!code & "'"
        Set rst2 = oSaccoMaster.GetRecordset(sql)
            
        
        
        
        
        
    '-- d_sp_UpdateTransPay @Code varchar(35), @Qnty float,@Amnt money,@Subsidy money,@GrossPay money, @EndDate varchar(10)  AS
    If Not rst2.EOF Then
            Set cn = New ADODB.Connection
                sql = "d_sp_UpdateTransPay '" & rst!code & "'," & rst2!qnty & "," & rst2!Amnt & "," & rst2!subsidy & "," & rst2!GPay & ",'" & Enddate & "','" & User & "'"
            oSaccoMaster.ExecuteThis (sql)
            
        End If
     
        '---Get transporters Total Deductions//d_sp_TotalTransDeduct @Code varchar(35),@Month bigint,@Year bigint
        Set Rs1 = New ADODB.Recordset
            sql = "d_sp_TotalTransDeduct '" & rst!code & "'," & month(DTPMilkDate) & "," & Year(DTPMilkDate) & ""
        Set Rs1 = oSaccoMaster.GetRecordset(sql)
    If Not Rs1.EOF Then
    'Dim TransTotalDed As Currency
    If Not IsNull(Rs1.Fields(0)) Then TransTotalDed = Rs1.Fields(0)
    End If
    
    Set rs3 = New ADODB.Recordset
    'Dim Startdate As String, Enddate As String
    'Dim DESCR As String
   ' Dim amount As Currency
    
    '--d_sp_SelTransDed @Code varchar(35), @startdate varchar(10),@enddate varchar(10)
    sql = "d_sp_SelTransDed '" & rst!code & "','" & Startdate & "','" & Enddate & "'"
    Set rs3 = oSaccoMaster.GetRecordset(sql)
    
    agrovet = 0
    BONUS = 0
    TMShares = 0
    FSA = 0
    HShares = 0
    Advance = 0
    Others = 0
    TransTotalDed = 0
If Not rs3.EOF Then
    While Not rs3.EOF
    DESCR = Trim(rs3.Fields(0))
    amount = 0
    amount = rs3.Fields(1)
    sql = "SELECT     Agrovet, AI, TMShares, FSA, HShares, Advance, Others FROM d_TransportersPayroll WHERE Code='" & rst!code & "' AND EndPeriod ='" & Enddate & "'"
    Set rs4 = oSaccoMaster.GetRecordset(sql)
     If UCase(rs4.Fields(0).name) = UCase(DESCR) Then
        agrovet = amount
    End If
    If UCase(rs4.Fields(1).name) = UCase(DESCR) Then
        BONUS = amount
    End If
    If UCase(rs4.Fields(2).name) = UCase(DESCR) Then
        TMShares = amount
    End If
    If UCase(rs4.Fields(3).name) = UCase(DESCR) Then
        FSA = amount
    End If
    If UCase(rs4.Fields(4).name) = UCase(DESCR) Then
        HShares = amount
    End If
    If UCase(rs4.Fields(5).name) = UCase(DESCR) Then
        Advance = amount
    End If
    If UCase(rs4.Fields(6).name) = UCase(DESCR) Then
        Others = amount
    End If

    '//d_sp_UpdatePayrollDed  @SNo bigint,@EndPeriod varchar(15),@Transport money,@Agrovet money,@AI money,@TMShares money,@FSA money,@HShares money,@Advance money,@Others
    rs3.MoveNext
    Wend
    ' d_sp_UpdateTransDed  @Code varchar(35),@EndPeriod varchar(15),@TotalDed money,@Agrovet money,@AI money,@TMShares money,@FSA money,@HShares money,@Advance money,@Others money
   
End If
 Set cn = New ADODB.Connection
    sql = "d_sp_UpdateTransDed  '" & rst!code & "','" & Enddate & "'," & TransTotalDed & "," & agrovet & "," & BONUS & "," & TMShares & "," & FSA & "," & HShares & "," & Advance & "," & Others & ""
    oSaccoMaster.ExecuteThis (sql)
    End If

End If

 
End If



'//d_sp_TotalDeduct-Total Deductions
'//d_sp_UpdateGPAYQnty - Total Grosspay and Quantity
'//d_sp_SupDed - Supply Deductions



Set rs2 = New ADODB.Recordset
'Dim Startdate As String, Enddate As String
Dim qnty As Currency, GPay As Currency
'Startdate = DateSerial(DTPMilkDate, cboMonth, 1)
sql = "d_sp_UpdateGPAYQnty '" & Startdate & "','" & Enddate & "'," & txtSNo & ""
Set rs2 = oSaccoMaster.GetRecordset(sql)
If Not rs2.EOF Then
If Not IsNull(rs2.Fields(0)) Then qnty = rs2.Fields(0)
If Not IsNull(rs2.Fields(1)) Then GPay = rs2.Fields(1)
End If


Set Rs1 = New ADODB.Recordset
sql = "d_sp_TotalDeduct " & txtSNo & "," & month(DTPMilkDate) & "," & Year(DTPMilkDate) & ""
Set Rs1 = oSaccoMaster.GetRecordset(sql)
If Not Rs1.EOF Then
Dim TotalDed As Currency
If Not IsNull(Rs1.Fields(0)) Then TotalDed = Rs1.Fields(0)
End If
'//Update payroll -- @SNo bigint,@EndPeriod varchar(15),@Kgs float,@GPay money,@NPay money,@TDeductions money,@auditid  varchar(35)
Set cn = New ADODB.Connection
sql = "d_sp_UpdatePayroll  " & txtSNo & ",'" & Enddate & "'," & qnty & "," & GPay & "," & GPay - TotalDed & "," & TotalDed & ",'" & User & "'"
oSaccoMaster.ExecuteThis (sql)



Set rs3 = New ADODB.Recordset
'Dim Startdate As String, Enddate As String
Dim desc As String
Dim Amnt As Currency
Startdate = DateSerial(Year(DTPMilkDate), month(DTPMilkDate), 1)
Enddate = DateSerial(Year(DTPMilkDate), month(DTPMilkDate) + 1, 1 - 1)
sql = "d_sp_SupDed " & txtSNo & ",'" & Startdate & "','" & Enddate & "'"
Set rs3 = oSaccoMaster.GetRecordset(sql)
If Not rs3.EOF Then
While Not rs3.EOF
desc = Trim(rs3.Fields(0))
Amnt = 0
Amnt = rs3.Fields(1)
sql = "SELECT     Transport, Agrovet, AI, TMShares, FSA, HShares, Advance, Others FROM d_Payroll WHERE SNo=" & txtSNo & " AND EndofPeriod ='" & Enddate & "'"
Set rs4 = oSaccoMaster.GetRecordset(sql)
If UCase(rs4.Fields(0).name) = UCase(desc) Then
Transport = Amnt
End If
If UCase(rs4.Fields(1).name) = UCase(desc) Then
agrovet = Amnt
End If
If UCase(rs4.Fields(2).name) = UCase(desc) Then
BONUS = Amnt
End If
If UCase(rs4.Fields(3).name) = UCase(desc) Then
TMShares = Amnt
End If
If UCase(rs4.Fields(4).name) = UCase(desc) Then
FSA = Amnt
End If
If UCase(rs4.Fields(5).name) = UCase(desc) Then
HShares = Amnt
End If
If UCase(rs4.Fields(6).name) = UCase(desc) Then
Advance = Amnt
End If
If UCase(rs4.Fields(7).name) = UCase(desc) Then
Others = Amnt
End If

'//d_sp_UpdatePayrollDed  @SNo bigint,@EndPeriod varchar(15),@Transport money,@Agrovet money,@AI money,@TMShares money,@FSA money,@HShares money,@Advance money,@Others
rs3.MoveNext
Wend
'//Update Deductions -- d_sp_UpdatePayrollDed  @SNo bigint,@EndPeriod varchar(15),@Transport money,@Agrovet money,@AI money,@TMShares money,@FSA money,@HShares money,@Advance money,@Others money
Set cn = New ADODB.Connection
sql = "d_sp_UpdatePayrollDed  " & txtSNo & ",'" & Enddate & "'," & Transport & "," & agrovet & "," & BONUS & "," & TMShares & "," & FSA & "," & HShares & "," & Advance & "," & Others & ""
oSaccoMaster.ExecuteThis (sql)
End If

Transport = 0
agrovet = 0
BONUS = 0
TMShares = 0
FSA = 0
HShares = 0
Advance = 0
Others = 0
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
        ttt = "LPT1"
'        ttt = cdgPrint.PrinterDefault
        'ttt = "D:\PROJECTS\FOSA\DAILY" & Date & ""
        Set fso = CreateObject("Scripting.FileSystemObject")
        'Set chkPrinter = fso.GetFile(ttt)
       
        
        Set txtFile = fso.CreateTextFile(ttt, True)
    txtFile.WriteLine "         " & cname & ""
    txtFile.WriteLine "            Milk Receipt"
    txtFile.WriteLine "---------------------------------------"
        
    Set rs2 = New ADODB.Recordset
    sql = "d_sp_ReceiptNumber"
    Set rs2 = oSaccoMaster.GetRecordset(sql)
    
    Dim RNumber As String
    'RNumber = rs2.Fields(0)
    If Not IsNull(rs2.Fields(0)) Then RNumber = rs2.Fields(0)
'    Else
'    RNumber = "0"
'    End If
    
    txtFile.WriteLine "SNo :" & txtSNo
    txtFile.WriteLine "Name :" & lblNames
    txtFile.WriteLine "Quantity Supplied :" & txtQnty & " Kgs"
    txtFile.WriteLine "Price" & Price & " Per Kgs"
    Set rs = New ADODB.Recordset
    sql = "d_sp_TotalMonth " & txtSNo & ",'" & Startdate & "','" & DTPMilkDate & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
    If Not IsNull(rs.Fields(0)) Then CummulKgs = rs.Fields(0)
    Else
    CummulKgs = "0.00"
    End If
    txtFile.WriteLine "Cummulative This Month : " & Format(CummulKgs, "#,##0.00" & " Kgs")
    Set rs = New ADODB.Recordset
    sql = "d_sp_TransName '" & txtSNo & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
    If Not IsNull(rs.Fields(0)) Then TRANSPORTER = rs.Fields(0)
    Else
        TRANSPORTER = "Self"
    End If
    txtFile.WriteLine "---------------------------------------"
    txtFile.WriteLine "Receipt Number :" & RNumber
    txtFile.WriteLine "TRANSPORTER :" & TRANSPORTER
    txtFile.WriteLine "Received by :" & username
    txtFile.WriteLine "---------------------------------------"
    txtFile.WriteLine "  Date :" & Format(DTPMilkDate, "dd/mm/yyyy") & " ,Time : " & Format(Time, "hh:mm:ss AM/PM")
    txtFile.WriteLine "         " & motto & ""
    txtFile.WriteLine "---------------------------------------"
    If chkComment.value = vbChecked Then
        txtFile.WriteLine txtComment
        txtFile.WriteLine "---------------------------------------"
    End If
    txtFile.WriteLine escFeedAndCut
    
 txtFile.Close
 Reset
End If


'* writing to notepad
If chkNotepad.value = vbChecked Then

'    Dim fso, txtfile
'    Dim ttt
'     Dim escFeedAndCut As String
     escFeedAndCut = Chr(29) + Chr(86) + Chr(65) '//Partial cut and feed
       cdgPrint.Filter = "*.csv|*.txt"
        cdgPrint.ShowSave
        ttt = cdgPrint.FileName
        If ttt = "" Then
        MsgBox "File should not be blank", vbCritical, "Data transfer"
        Exit Sub
        End If
        'ttt = "D:\PROJECTS\FOSA\DAILY" & Date & ""
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set txtFile = fso.CreateTextFile(ttt, True)
        txtFile.WriteLine
        
    txtFile.WriteLine "---------------------------------------"
    txtFile.WriteLine "" & cname & ""
   ' Printer.Print Tab(0); "Kimathi House Branch"
    txtFile.WriteLine " " & paddress & " "
    txtFile.WriteLine "" & town & ""
    txtFile.WriteLine "Milk Receipt"
    txtFile.WriteLine "---------------------------------------"
'    If cbomemtrans = "Shares" Then
'    DESC = bosanames & " -Member No " & memberno
    txtFile.WriteLine "SNo :" & txtSNo
    txtFile.WriteLine "Name :" & lblNames
'    Else
    txtFile.WriteLine "Quantity Supplied :" & txtQnty & " Kgs"
    txtFile.WriteLine "Price" & Price & " Per Kgs"
    Startdate = DateSerial(Year(DTPMilkDate), month(DTPMilkDate) - 1, 1)
    'sql = "d_sp_TotalMonth " & txtSNo & ",'" & StartDate & "','" & DTPMilkDate & "'"
    Set rs = New ADODB.Recordset
    sql = "d_sp_TotalMonth " & txtSNo & ",'" & Startdate & "','" & DTPMilkDate & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
    If Not IsNull(rs.Fields(0)) Then CummulKgs = rs.Fields(0)
    Else
    CummulKgs = "0.00"
    End If
    txtFile.WriteLine "Cummulative This Month " & Format(CummulKgs, "#,##0.00" & " Kgs")
'    End If
    Set rs = New ADODB.Recordset
    sql = "d_sp_TransName '" & txtSNo & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
    If Not IsNull(rs.Fields(0)) Then TRANSPORTER = rs.Fields(0)
    Else
    TRANSPORTER = "Self"
    End If
    txtFile.WriteLine "---------------------------------------"
    txtFile.WriteLine "Transporter :" & TRANSPORTER
    txtFile.WriteLine "Received by :" & username
    txtFile.WriteLine "---------------------------------------"
    txtFile.WriteLine "Date :" & Format(Get_Server_Date, "dd/mm/yyyy HH:MM:SS AM/PM")
    txtFile.WriteLine "     " & motto & ""
    txtFile.WriteLine "---------------------------------------"
    txtFile.WriteLine escFeedAndCut
    'Printer.Print
    'Printer.CurrentX = 500#
    'Printer.FontSize = 10
    'Printer.CurrentX = 500
    'Printer.FontSize = 8
'    Printer.Print Tab(0); "Date :"; Tab(10); Format(Get_Server_Date, "dd/mm/yyyy HH:MM:SS AM/PM")
'    Printer.Print Tab(0); "     TDL - Improves Your Value "
'    Printer.Print Tab(0); "---------------------------------------"
'    Printer.CurrentX = 500
'    Printer.FontSize = 8
'    Printer.Print
'    Printer.CurrentX = 500
'    Printer.FontSize = 8
'    Printer.Print
txtFile.Close
End If

loadMilk

txtSNo = ""
txtQnty = ""
'txtSNo_Validate True
txtSNo.SetFocus
Exit Sub
ErrorHandler:

MsgBox err.description




End Sub

Private Sub cmdPrint_Click()

Startdate = DateSerial(Year(DTPMilkDate), month(DTPMilkDate), 1)
'Enddate = DateSerial(year(DTPStmts), month(DTPStmts) + 1, 1 - 1)
Enddate = DTPMilkDate
oSaccoMaster.ExecuteThis ("d_sp_TransUpdate '" & Startdate & "','" & Enddate & "','" & User & "','" & txtTCode & "'")

If txtTCode = "" Then
MsgBox "Please enter the transporter code.", vbInformation
txtTCode.SetFocus
Exit Sub
End If

txtTCode_Validate True

If lblTName.Caption = "" Then
MsgBox "Please enter code for a valid transporter. Transporter with code " & txtTCode & " does not exist.", vbInformation
txtTCode.SetFocus
Exit Sub
End If

'//Print Receipt
    Set rs = oSaccoMaster.GetRecordset("d_sp_TransDTotal '" & txtTCode & "','" & DTPMilkDate & "'")
    If IsNull(rs.Fields(0)) Then
        MsgBox "There is no milk supplied by Code " & txtTCode
        txtTCode.SetFocus
        Exit Sub
    End If
    
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
       
        ttt = "\\127.0.0.1\E-PoS 80mm Thermal Printer" 'LPT1
        Set fso = CreateObject("Scripting.FileSystemObject")

        
        Set txtFile = fso.CreateTextFile(ttt, True)
    txtFile.WriteLine "             " & cname & ""
    txtFile.WriteLine "           Transporter's Receipt"
    txtFile.WriteLine "---------------------------------------"
    
    txtFile.WriteLine "Transporter Code :" & txtTCode
    txtFile.WriteLine "Transporter Name :" & lblTName
    
    txtFile.WriteLine "Quantity Transported :" & Format(CCur(rs.Fields(0)), "#,##0.00") & " Kgs "
'    txtFile.WriteLine "Price:" & Price & " Per Kgs"
    Set rs = oSaccoMaster.GetRecordset("d_sp_TransTotal '" & txtTCode & "'," & month(DTPMilkDate) & ", " & Year(DTPMilkDate))
    CummulKgs = 0
    If Not IsNull(rs.Fields(0)) Then CummulKgs = rs.Fields(0)
   
    txtFile.WriteLine "Cummulative This Month " & Format(CummulKgs, "#,##0.00" & " Kgs")
    txtFile.WriteLine "---------------------------------------"
    txtFile.WriteLine " Printed By " & UCase(username)
    txtFile.WriteLine "    Date :" & Format(Get_Server_Date, "dd/mm/yyyy HH:MM:SS AM/PM")
    txtFile.WriteLine "     " & motto & ""
    txtFile.WriteLine "****************************************"
    txtFile.WriteLine escFeedAndCut
    txtFile.Close
    
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
    If rs!status <> 1 Then
    MsgBox "You are not allowed to make complaint", vbInformation
     
    Exit Sub
   End If
     End If

If Trim(txtSNo) = "" Then
    MsgBox "Please enter the supplier number."
        txtSNo.SetFocus
    Exit Sub
End If
If Trim(cbolocation) = "" Then
    MsgBox "PLEASE SELECT THE STATION."
        cbolocation.SetFocus
    Exit Sub
End If
If Trim(txtQnty) = "" Then
    MsgBox "Please enter the quantity supplied."
        txtQnty.SetFocus
Exit Sub
End If

'''*****************************check if supliers branch is correct ****
sql = ""
sql = "select SNo,Branch from d_Suppliers where SNo='" & txtSNo & "'"
Set rss = oSaccoMaster.GetRecordset(sql)
If rss!Branch <> cbolocation Then
MsgBox "This supplier is not from this branch: " & cbolocation & "", vbInformation
Exit Sub
End If

'''*****************************end **************



sql = ""
sql = "select* from d_Suppliers where sno='" & txtSNo & "' and approval=0"
Set rss = oSaccoMaster.GetRecordset(sql)
If Not rss.EOF Then
MsgBox "The supplier need approval first", vbInformation
Exit Sub
End If

Dim rsrepeat As New ADODB.Recordset
Dim nnn As Integer
Dim kk As Integer
'*******************mass transport********************
Set rs = New ADODB.Recordset
sql = "SELECT mass from d_Suppliers  where sno=" & txtSNo & ""
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
kk = rs!mass
End If
'Dim autorize As Integer
Set rsrepeat = New ADODB.Recordset
sql = ""
sql = "[d_sp_sumkgsintake] " & txtSNo & ",'" & DTPMilkDate & "','" & cbolocation & "'"
'sql = "set dateformat dmy SELECT COUNT(SNo) FROM  d_Milkintake WHERE TransDate = '" & DTPMilkDate & "' AND SNo = '" & txtSNo & "' and LOCATION='" & cbolocation & "'"
'sql = "set dateformat dmy SELECT COUNT(sno)   FROM d_Milkintake WHERE     transdate = '" & DTPMilkDate & "' AND sno = '" & txtSNo & "'"
Set rsrepeat = oSaccoMaster.GetRecordset(sql)
If Not rsrepeat.EOF Then
If CLng(rsrepeat.Fields(0)) >= 5 And kk <> 1 Then
MsgBox "You have supplied milk Five times today. You are not allowed to supply for the sixth time ", vbCritical
Exit Sub
End If
End If

txtQnty.SetFocus
Dim k As Double
'''''''neg
If txtQnty < 0 Then
k = txtQnty * -1
'If txttoday < k Then
'    MsgBox "kilos Delivered is less than that, revise the kilos from the supplier."
'        txtSNo.SetFocus
'    Exit Sub
'End If
End If
'End If
txtSNo_Validate True

If Trim(lblNames.Caption) = "" Then
MsgBox "Please enter a valid supplier number."
txtSNo.SetFocus
Exit Sub
End If

If Not IsNumeric(txtQnty) Then
MsgBox "Please enter a number. " & txtQnty & " is not a number", vbExclamation
txtQnty.SetFocus
Exit Sub
End If

Startdate = DateSerial(Year(DTPMilkDate), month(DTPMilkDate), 1)
Enddate = DateSerial(Year(DTPMilkDate), month(DTPMilkDate) + 1, 1 - 1)

Set rs = oSaccoMaster.GetRecordset("d_sp_IsClosed '" & Enddate & "'")
If Not rs.EOF Then
    MsgBox "The period " & Enddate & " has been closed by " & rs.Fields(0)
    Exit Sub
End If

Dim ans As String
Set rs = New ADODB.Recordset
 sql = "SET dateformat dmy SELECT d_Milkintake.[ID],d_Milkintake.SNo, d_Milkintake.QSupplied, "
 sql = sql & " d_Milkintake.TransTime , d_Suppliers.[Names] FROM  d_Milkintake INNER JOIN "
 sql = sql & " d_Suppliers ON d_Milkintake.SNo = d_Suppliers.SNo AND  d_Milkintake.SNo= " & txtSNo & " AND (d_milkintake.TransDate = '" & DTPMilkDate & "')"
 sql = sql & " ORDER BY d_Milkintake.Id DESC"

 
Set rs = oSaccoMaster.GetRecordset(sql)
If rs.RecordCount > 0 Then

With lvwMilkCollected
    
       .ListItems.Clear
    
        .ColumnHeaders.Clear

  End With

    With lvwMilkCollected
        
        
        .ColumnHeaders.Add , , "SNo", 2000
        .ColumnHeaders.Add , , "Names", 3000
        .ColumnHeaders.Add , , "QNTY"
        .ColumnHeaders.Add , , "Time"
        .ColumnHeaders.Add , , "Receipt No.", 2000
    
        While Not rs.EOF
        
        If Not IsNull(rs.Fields("SNo")) Then
        
            Set li = .ListItems.Add(, , Trim(rs.Fields("SNo")))
            End If
            If Not IsNull(rs.Fields("Names")) Then
            li.ListSubItems.Add , , Trim(rs.Fields("Names"))
            End If
            If Not IsNull(rs.Fields("QSupplied")) Then
            li.ListSubItems.Add , , Trim(rs.Fields("QSupplied"))
            End If
            If Not IsNull(rs.Fields("TransTime")) Then
            li.ListSubItems.Add , , Trim(rs.Fields("TransTime"))
            End If
            If Not IsNull(rs.Fields("ID")) Then
             li.ListSubItems.Add , , Trim(rs.Fields("ID"))
            End If

            
                    rs.MoveNext
        
        Wend
        
    End With
    
    rs.Close
    
    Set rs = Nothing
    
lvwMilkCollected.View = lvwReport
ans = MsgBox("Supplier number " & txtSNo & " had supplied milk today. Add this?", vbYesNo, "MILK REPEAT")
If ans = vbNo Then
txtSNo.SetFocus
Exit Sub
End If
End If

''>>>>>>>>>>>>>>>>>>>>>>>check if has different prices a day
'If chkPrices = 0 Then

''....>>>>>>>>>>>>>>>>>>>.select the price for traders and individual......>>>>>>>>>


Set rs = New ADODB.Recordset
'''Dim rsmesg As Recordset
''sql = ""
''sql = "Select Trader  from d_Suppliers  where SNo ='" & txtSNo & "'"
''Set rs = oSaccoMaster.GetRecordset(sql)
'''Set rsmesg = oSaccoMaster.GetRecordset(sql)
''Dim St As Integer
''St = rs!Trader
''If St = True Then
'' St = 1
'' Else
'' St = 0
'' End If
''If St = 1 Then
'St = rsmesg!mes
'If St = 1 Then

''sql = "SELECT Price from d_Debtors2 where Dcode='" & txtSNo & "'"
''Set rs = oSaccoMaster.GetRecordset(sql)
''If Not rs.EOF Then
''Price = rs!Price
''End If
''
''Else

Set rs = New ADODB.Recordset
sql = "SELECT Price from d_Price"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
Price = rs!Price
End If
''End If
''............end of selection...............

''Else
''
''    sql = "SELECT Price from d_Price2 where Sno='" & txtSNo & "'"
''    Set rs = oSaccoMaster.GetRecordset(sql)
''    If Not rs.EOF Then
''      Price = rs!Price
''     Else
''       MsgBox "Please add the Supplier '" & txtSNo & "' to the list.", vbCritical
''     Exit Sub
''     End If
''
''End If

''>>>>>>>>>>>>>>>>>>>>>end of checking


'Set rs = New ADODB.Recordset
'sql = "SELECT Price from d_Price"
'Set rs = oSaccoMaster.GetRecordset(sql)
'If Not rs.EOF Then
'Price = rs!Price
'End If



'//Update Milk Intake

Dim tbl As String
    
'Set cn = New ADODB.Connection
    Set cn = New Connection
    Set rs = CreateObject("adodb.recordset")
    Set cn = CreateObject("adodb.connection")
    Provider = SelectedDsn
    cn.Open Provider, "atm", "atm"

sql = "d_sp_MilkIntake " & txtSNo & ",'" & DTPMilkDate & "'," & txtQnty & "," & Price & "," & Price * CCur(txtQnty) & ",'" & Time & "','" & User & "','Intake Complain Desk','" & cbolocation & "'"
oSaccoMaster.ExecuteThis (sql)
tbl = "d_milk"
sql = ""
sql = "set dateformat dmy INSERT INTO AUDITTRANS"
sql = sql & "         (TransTable, TransDescription, TransDate, Amount, AuditID, AuditTime)"
sql = sql & " VALUES     ('" & tbl & "','" & txtComment & "','" & DTPMilkDate & "'," & txtQnty & ",'" & User & "','" & Get_Server_Date & "')"
oSaccoMaster.Execute (sql)

'//Update Daily Intake

'Set rs = New ADODB.Recordset
'sql = "SET dateformat DMY SELECT [" & Day(DTPMilkDate) & "] AS Milk from d_DailySummary WHERE SNo =" & txtSNo & " AND Endperiod ='" & Enddate & "'"
'Set rs = oSaccoMaster.GetRecordset(sql)
'If Not rs.EOF Then
'sql = "SET dateformat DMY Update  d_DailySummary SET [" & Day(DTPMilkDate) & "] = " & CCur(rs!Milk) + CCur(txtQnty) & " WHERE SNo =" & txtSNo & " AND Endperiod ='" & Enddate & "'"
'oSaccoMaster.ExecuteThis (sql)
'End If
'
'If rs.EOF Then
'sql = "SET dateformat DMY INSERT INTO  d_DailySummary (SNo,[" & Day(DTPMilkDate) & "],EndPeriod) Values(" & txtSNo & "," & txtQnty & " ,'" & Enddate & "')"
'oSaccoMaster.ExecuteThis (sql)
'End If

'//Check Transporter if active

Set rs = New ADODB.Recordset
    sql = "d_sp_TransActive " & txtSNo & ",'" & DTPMilkDate & "'"
        Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then

        transdate = rs!rate
        If transdate < Startdate Then
            transdate = Startdate
        End If


        '---d_sp_SelTransDet @SNo bigint,@StartDate varchar(10), @Enddate varchar(10)
        Set rst = New ADODB.Recordset
            sql = "d_sp_SelTransDet " & txtSNo & ",'" & transdate & "','" & Enddate & "'"
        Set rst = oSaccoMaster.GetRecordset(sql)

    '--d_sp_UpdateDetTrans @SNo bigint,@QNTY float,@Amnt money,@Code varchar(35),@Subsidy money,@EPeriod varchar(10),@user varchar(35)
        If Not rst.EOF Then
            Set cn = New ADODB.Connection
                sql = "d_sp_UpdateDetTrans " & txtSNo & "," & rst!qnty & "," & rst!amount & ",'" & rst!code & "'," & rst!subsidy & ",'" & Enddate & "','" & User & "'"
            oSaccoMaster.ExecuteThis (sql)




        '----d_sp_SelTransGPayQnty @EP varchar(10), @Code varchar(35)
        Set rst2 = New ADODB.Recordset
            sql = "d_sp_SelTransGPayQnty '" & Enddate & "','" & rst!code & "'"
        Set rst2 = oSaccoMaster.GetRecordset(sql)

    '-- d_sp_UpdateTransPay @Code varchar(35), @Qnty float,@Amnt money,@Subsidy money,@GrossPay money, @EndDate varchar(10)  AS
    If Not rst2.EOF Then
            Set cn = New ADODB.Connection
                sql = "d_sp_UpdateTransPay '" & rst!code & "'," & rst2!qnty & "," & rst2!Amnt & "," & rst2!subsidy & "," & rst2!GPay & ",'" & Enddate & "','" & User & "'"
            oSaccoMaster.ExecuteThis (sql)

        End If

        '---Get transporters Total Deductions//d_sp_TotalTransDeduct @Code varchar(35),@Month bigint,@Year bigint
        Set Rs1 = New ADODB.Recordset
            sql = "d_sp_TotalTransDeduct '" & rst!code & "'," & month(DTPMilkDate) & "," & Year(DTPMilkDate) & ""
        Set Rs1 = oSaccoMaster.GetRecordset(sql)
    If Not Rs1.EOF Then
    Dim TransTotalDed As Currency
    If Not IsNull(Rs1.Fields(0)) Then TransTotalDed = Rs1.Fields(0)
    End If

    Set rs3 = New ADODB.Recordset
    Dim DESCR As String
    Dim amount As Currency

    '--d_sp_SelTransDed @Code varchar(35), @startdate varchar(10),@enddate varchar(10)
    sql = "d_sp_SelTransDed '" & rst!code & "','" & Startdate & "','" & Enddate & "'"
    Set rs3 = oSaccoMaster.GetRecordset(sql)

    agrovet = 0
    BONUS = 0
    TMShares = 0
    FSA = 0
    HShares = 0
    Advance = 0
    Others = 0
    TransTotalDed = 0


If Not rs3.EOF Then
    While Not rs3.EOF
    DESCR = Trim(rs3.Fields(0))
    amount = 0
    amount = rs3.Fields(1)
    sql = "SELECT     Agrovet, BONUS, TMShares, FSA, HShares, Advance, Others FROM d_TransportersPayroll WHERE Code='" & rst!code & "' AND EndPeriod ='" & Enddate & "'"
    Set rs4 = oSaccoMaster.GetRecordset(sql)
     If UCase(rs4.Fields(0).name) = UCase(DESCR) Then
        agrovet = amount
    End If
    If UCase(rs4.Fields(1).name) = UCase(DESCR) Then
        BONUS = amount
    End If
    If UCase(rs4.Fields(2).name) = UCase(DESCR) Then
        TMShares = amount
    End If
    If UCase(rs4.Fields(3).name) = UCase(DESCR) Then
        FSA = amount
    End If
    If UCase(rs4.Fields(4).name) = UCase(DESCR) Then
        HShares = amount
    End If
    If UCase(rs4.Fields(5).name) = UCase(DESCR) Then
        Advance = amount
    End If
    If UCase(rs4.Fields(6).name) = UCase(DESCR) Then
        Others = amount
    End If

    '//d_sp_UpdatePayrollDed  @SNo bigint,@EndPeriod varchar(15),@Transport money,@Agrovet money,@AI money,@TMShares money,@FSA money,@HShares money,@Advance money,@Others
    rs3.MoveNext
    Wend

End If
     ' d_sp_UpdateTransDed  @Code varchar(35),@EndPeriod varchar(15),@TotalDed money,@Agrovet money,@AI money,@TMShares money,@FSA money,@HShares money,@Advance money,@Others money

    Set cn = New ADODB.Connection
    sql = "d_sp_UpdateTransDed  '" & rst!code & "','" & Enddate & "'," & TransTotalDed & "," & agrovet & "," & BONUS & "," & TMShares & "," & FSA & "," & HShares & "," & Advance & "," & Others & ""
    oSaccoMaster.ExecuteThis (sql)


        '//Update supplier deductions
            Set cn = New ADODB.Connection
            Dim r2  As String
            r2 = "  "
                sql = "d_sp_SupplierDeduct " & txtSNo & ",'" & DTPMilkDate & "','Transport'," & rs!rate * CCur(txtQnty) & ",'" & Startdate & "','" & Enddate & "'," & Year(Enddate) & ",'" & User & "','" & r2 & "',''"
            oSaccoMaster.ExecuteThis (sql)




End If
'Check inactive Transporter
Set rs = New ADODB.Recordset
sql = "d_sp_TransInActive " & txtSNo & ",'" & Startdate & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then

'--Update Detailed Trasport
    transdate = rs!Startdate
        If transdate < Startdate Then
            transdate = Startdate
        End If

    '---d_sp_SelTransDet @SNo bigint,@StartDate varchar(10), @Enddate varchar(10)
        Set rst = New ADODB.Recordset
            sql = "d_sp_SelTransDet " & txtSNo & ",'" & transdate & "','" & rs!dateinactivate & "'"
        Set rst = oSaccoMaster.GetRecordset(sql)

    '--d_sp_UpdateDetTrans @SNo bigint,@QNTY float,@Amnt money,@Code varchar(35),@Subsidy money,@EPeriod varchar(10),@user varchar(35)
        If Not rst.EOF Then
            Set cn = New ADODB.Connection
                sql = "d_sp_UpdateDetTrans " & txtSNo & "," & rst!qnty & "," & rst!amount & ",'" & rst!code & "'," & rst!subsidy & ",'" & Enddate & "','" & User & "'"
            oSaccoMaster.ExecuteThis (sql)

            '----d_sp_SelTransGPayQnty @EP varchar(10), @Code varchar(35)
        Set rst2 = New ADODB.Recordset
            sql = "d_sp_SelTransGPayQnty '" & Enddate & "','" & rst!code & "'"
        Set rst2 = oSaccoMaster.GetRecordset(sql)






    '-- d_sp_UpdateTransPay @Code varchar(35), @Qnty float,@Amnt money,@Subsidy money,@GrossPay money, @EndDate varchar(10)  AS
    If Not rst2.EOF Then
            Set cn = New ADODB.Connection
                sql = "d_sp_UpdateTransPay '" & rst!code & "'," & rst2!qnty & "," & rst2!Amnt & "," & rst2!subsidy & "," & rst2!GPay & ",'" & Enddate & "','" & User & "'"
            oSaccoMaster.ExecuteThis (sql)

        End If

        '---Get transporters Total Deductions//d_sp_TotalTransDeduct @Code varchar(35),@Month bigint,@Year bigint
        Set Rs1 = New ADODB.Recordset
            sql = "d_sp_TotalTransDeduct '" & rst!code & "'," & month(DTPMilkDate) & "," & Year(DTPMilkDate) & ""
        Set Rs1 = oSaccoMaster.GetRecordset(sql)
    If Not Rs1.EOF Then
    'Dim TransTotalDed As Currency
    If Not IsNull(Rs1.Fields(0)) Then TransTotalDed = Rs1.Fields(0)
    End If

    Set rs3 = New ADODB.Recordset
    'Dim Startdate As String, Enddate As String
    'Dim DESCR As String
   ' Dim amount As Currency

    '--d_sp_SelTransDed @Code varchar(35), @startdate varchar(10),@enddate varchar(10)
    sql = "d_sp_SelTransDed '" & rst!code & "','" & Startdate & "','" & Enddate & "'"
    Set rs3 = oSaccoMaster.GetRecordset(sql)

    agrovet = 0
    BONUS = 0
    TMShares = 0
    FSA = 0
    HShares = 0
    Advance = 0
    Others = 0
    TransTotalDed = 0
If Not rs3.EOF Then
    While Not rs3.EOF
    DESCR = Trim(rs3.Fields(0))
    amount = 0
    amount = rs3.Fields(1)
    sql = "SELECT     Agrovet, BONUS, TMShares, FSA, HShares, Advance, Others FROM d_TransportersPayroll WHERE Code='" & rst!code & "' AND EndPeriod ='" & Enddate & "'"
    Set rs4 = oSaccoMaster.GetRecordset(sql)
     If UCase(rs4.Fields(0).name) = UCase(DESCR) Then
        agrovet = amount
    End If
    If UCase(rs4.Fields(1).name) = UCase(DESCR) Then
        BONUS = amount
    End If
    If UCase(rs4.Fields(2).name) = UCase(DESCR) Then
        TMShares = amount
    End If
    If UCase(rs4.Fields(3).name) = UCase(DESCR) Then
        FSA = amount
    End If
    If UCase(rs4.Fields(4).name) = UCase(DESCR) Then
        HShares = amount
    End If
    If UCase(rs4.Fields(5).name) = UCase(DESCR) Then
        Advance = amount
    End If
    If UCase(rs4.Fields(6).name) = UCase(DESCR) Then
        Others = amount
    End If

    '//d_sp_UpdatePayrollDed  @SNo bigint,@EndPeriod varchar(15),@Transport money,@Agrovet money,@AI money,@TMShares money,@FSA money,@HShares money,@Advance money,@Others
    rs3.MoveNext
    Wend
    ' d_sp_UpdateTransDed  @Code varchar(35),@EndPeriod varchar(15),@TotalDed money,@Agrovet money,@AI money,@TMShares money,@FSA money,@HShares money,@Advance money,@Others money

End If
 Set cn = New ADODB.Connection
    sql = "d_sp_UpdateTransDed  '" & rst!code & "','" & Enddate & "'," & TransTotalDed & "," & agrovet & "," & BONUS & "," & TMShares & "," & FSA & "," & HShares & "," & Advance & "," & Others & ""
    oSaccoMaster.ExecuteThis (sql)
    End If

End If


End If






Set rs2 = New ADODB.Recordset
'Dim Startdate As String, Enddate As String
Dim qnty As Currency, GPay As Currency
'Startdate = DateSerial(DTPMilkDate, cboMonth, 1)
sql = "d_sp_UpdateGPAYQnty '" & Startdate & "','" & Enddate & "'," & txtSNo & ""
Set rs2 = oSaccoMaster.GetRecordset(sql)
If Not rs2.EOF Then
If Not IsNull(rs2.Fields(0)) Then qnty = rs2.Fields(0)
If Not IsNull(rs2.Fields(1)) Then GPay = rs2.Fields(1)
End If

'//insert location to d_dailysummary
           
           Provider = "MAZIWA"
           Set cn = New ADODB.Connection
          cn.Open Provider, "atm", "atm"
'           sql = ""
'           sql = "set dateformat dmy INSERT INTO d_DailySummary(Location) where SNo='" & txtSNo & "' "
'           sql = sql & "('" & cboLocation & "')"
'           'oSaccoMaster.ExecuteThis (sql)
'           cn.Execute sql

           sql = ""
           sql = "set dateformat DMY update d_DailySummary set Location='" & cbolocation & "' where SNo='" & txtSNo & "' "
           cn.Execute sql
'end



Set Rs1 = New ADODB.Recordset
sql = "d_sp_TotalDeduct " & txtSNo & "," & month(DTPMilkDate) & "," & Year(DTPMilkDate) & ""
Set Rs1 = oSaccoMaster.GetRecordset(sql)
If Not Rs1.EOF Then
Dim TotalDed As Currency
If Not IsNull(Rs1.Fields(0)) Then TotalDed = Rs1.Fields(0)
End If
'//Update payroll -- @SNo bigint,@EndPeriod varchar(15),@Kgs float,@GPay money,@NPay money,@TDeductions money,@auditid  varchar(35)
Set cn = New ADODB.Connection
sql = "d_sp_UpdatePayroll  " & txtSNo & ",'" & Enddate & "'," & qnty & "," & GPay & "," & GPay - TotalDed & "," & TotalDed & ",'" & User & "'"
oSaccoMaster.ExecuteThis (sql)



Set rs3 = New ADODB.Recordset
'Dim Startdate As String, Enddate As String
Dim desc As String
Dim Amnt As Currency
Startdate = DateSerial(Year(DTPMilkDate), month(DTPMilkDate), 1)
Enddate = DateSerial(Year(DTPMilkDate), month(DTPMilkDate) + 1, 1 - 1)
sql = "d_sp_SupDed " & txtSNo & ",'" & Startdate & "','" & Enddate & "'"
Set rs3 = oSaccoMaster.GetRecordset(sql)
If Not rs3.EOF Then
While Not rs3.EOF
desc = Trim(rs3.Fields(0))
Amnt = 0
Amnt = rs3.Fields(1)
sql = "SELECT     Transport, Agrovet, BONUS, TMShares, FSA, HShares, Advance, Others FROM d_Payroll WHERE SNo=" & txtSNo & " AND EndofPeriod ='" & Enddate & "'"
Set rs4 = oSaccoMaster.GetRecordset(sql)
If UCase(rs4.Fields(0).name) = UCase(desc) Then
Transport = Amnt
End If
If UCase(rs4.Fields(1).name) = UCase(desc) Then
agrovet = Amnt
End If
If UCase(rs4.Fields(2).name) = UCase(desc) Then
BONUS = Amnt
End If
If UCase(rs4.Fields(3).name) = UCase(desc) Then
TMShares = Amnt
End If
If UCase(rs4.Fields(4).name) = UCase(desc) Then
FSA = Amnt
End If
If UCase(rs4.Fields(5).name) = UCase(desc) Then
HShares = Amnt
End If
If UCase(rs4.Fields(6).name) = UCase(desc) Then
Advance = Amnt
End If
If UCase(rs4.Fields(7).name) = UCase(desc) Then
Others = Amnt
End If

'//d_sp_UpdatePayrollDed  @SNo bigint,@EndPeriod varchar(15),@Transport money,@Agrovet money,@AI money,@TMShares money,@FSA money,@HShares money,@Advance money,@Others
rs3.MoveNext
Wend
'//Update Deductions -- d_sp_UpdatePayrollDed  @SNo bigint,@EndPeriod varchar(15),@Transport money,@Agrovet money,@AI money,@TMShares money,@FSA money,@HShares money,@Advance money,@Others money
Set cn = New ADODB.Connection
sql = "d_sp_UpdatePayrollDed  " & txtSNo & ",'" & Enddate & "'," & Transport & "," & agrovet & "," & BONUS & "," & TMShares & "," & FSA & "," & HShares & "," & Advance & "," & Others & ""
oSaccoMaster.ExecuteThis (sql)
End If

' '************message***************'
     ''''''''//check settings  for sms alert///''''''''''

'Dim rsmesg As Recordset
'sql = ""
'sql = "Select mes  from d_Suppliers  where SNo ='" & txtSNo & "'"
'
'Set rsmesg = oSaccoMaster.GetRecordset(sql)
'Dim St, mes As Integer
'St = rsmesg!mes
'If St = 1 Then
'
        'If SendSmsOnSalary = True Then
        'Dim phone As String
        Dim RsPhoneNumbers As New ADODB.Recordset
        Dim MsgContent As String
'        Set RsPhoneNumbers = oSaccoMaster.GetRecordset("select  PhoneNo  from d_Suppliers  where SNo ='" & txtSNo & "' and PhoneNo  is not null and PhoneNo <>'' and LEN(PhoneNo) =10 ")
'
'        If Not RsPhoneNumbers.EOF Then
'            Phone = Trim(RsPhoneNumbers.Fields!PhoneNo & "")
'        Else
'            Phone = ""
'        End If
'
'
'            '//insert into sms alert
'        Set rs = New ADODB.Recordset
'        sql = "d_sp_TotalMonth " & txtSNo & ",'" & Startdate & "','" & DTPMilkDate & "'"
'        Set rs = oSaccoMaster.GetRecordset(sql)
'        If Not rs.EOF Then
'        If Not IsNull(rs.Fields(0)) Then CummulKgs = rs.Fields(0)
'        Else
'        CummulKgs = "0.00"
'        End If
'            If Trim(Phone) <> "" Then
'                 '''''''''''// FOR ONLY NEW FARMERS OR THOSE STARTING TO SUPPLY MILK RECEIVE THIS //'''''''''''''''
'
'               Dim rsmesg1 As Recordset
'''               sql = ""
'''               sql = "Select mez  from Swift_Messages1  where SNo ='" & txtSNo & "'"
'''               Set rsmesg1 = oSaccoMaster.GetRecordset(sql)
'
'''               If rsmesg1.EOF Then
'''                 'Dim stet, mez As Integer
'''                 'stet = rsmesg1!mez
'''                 'If stet = 0 Then
'''                     MsgContent = ""
'''                     MsgContent = "Dear supplier " & lblNames & ", On " & Date & " You have supplied :" & txtQnty & " Kgs ,Current Buying Price: " & Price & "/= per ltr, Cummulative For This Month :  " & Format(CummulKgs, "#,##0.00") & " Kgs. To continue receiving these Notification Daily @ ksh 1.5, Kindly visit our Offices to Subscribe"
'''                     strSQL = ""
'''                     strSQL = "INSERT INTO Swift_Messages(SaccoCode,PhoneNumber,Message, msgstatus,,Transdate)"
'''                     strSQL = strSQL & "Values (14,'" & Phone & "','" & MsgContent & "',0,'" & Date & "')"
'''                     oSaccoMaster.ExecuteThis (strSQL)
'''                         ''''''''' //Insert Table for new farmers //''''''''
'''                        sql = ""
'''                        sql = sql & "SET dateformat DMY INSERT INTO Swift_Messages1(SNO, PhoneNumber, mez, Auditdate) VALUES ("
'''                        sql = sql & txtSNo & ",'" & Phone & "','1','" & Date & "')"
'''                        oSaccoMaster.ExecuteThis (sql)
'''
'''                 Else
'
'        '''''''''''''''// end of NEW FARMERS //''''''''''''''''''
'            MsgContent = ""
'            MsgContent = "Dear " & lblNames & ", On " & Date & " You have supplied :" & txtQnty & " Kgs of milk. Cummulative For this Month  is:  " & Format(CummulKgs, "#,##0.00") & " Kgs.Thanks you for being our customer SAMLI DAIRY"
'            strSQL = ""
'            strSQL = "INSERT INTO Messages(Telephone,Content,ProcessTime, MsgType,Source,code)"
'            strSQL = strSQL & "Values ('" & Phone & "','" & MsgContent & "',getDate(),'Outbox','" & User & "','SAMLI')"
'            oSaccoMaster.ExecuteThis (strSQL)
'
''           MsgContent = ""
''           MsgContent = "Dear supplier " & lblNames & ", On " & Date & " You have supplied :" & txtQnty & " Kgs ,Current Buying Price: " & Price & "/= per ltr, Cummulative For This Month :  " & Format(CummulKgs, "#,##0.00") & " Kgs"
''           oSaccoMaster.ExecuteThis ("SET dateformat DMY Update Swift_Messages SET Message ='" & MsgContent & "' WHERE PhoneNumber= '" & Phone & "'and msgstatus = '0'  and Transdate ='" & Date & "'")
'''       sql = ""
''       sql = "set dateformat dmy delete gltransactions WHERE DocumentNo='Milk Sales'  and (Amount='" & (CCur(Price) * CCur(txtquantity)) & "') and (TransDate = '" & frmMilkControl.DTPDispatchDate & "')"
''       oSaccoMaster.ExecuteThis (sql)
'
''             sql = ""
''             sql = " SET dateformat DMY delete Swift_Messages where PhoneNumber='" & Phone & "' and msgstatus='0' and Auditdate >" & Date & " "
''             oSaccoMaster.ExecuteThis (sql)
''
''
''            MsgContent = ""
''            MsgContent = "Dear supplier " & lblNames & ", On " & Date & " You have supplied :" & txtQnty & " Kgs ,Current Buying Price: " & Price & "/= per ltr, Cummulative For This Month :  " & Format(CummulKgs, "#,##0.00") & " Kgs"
''            strSQL = ""
''            strSQL = "INSERT INTO Swift_Messages(SaccoCode,PhoneNumber,Message, msgstatus)"
''            strSQL = strSQL & "Values (14,'" & Phone & "','" & MsgContent & "',0)"
''
''            oSaccoMaster.ExecuteThis (strSQL)
''            ''''''''' Deduct sms charge''''''''
'
'              'sql = ""
'              'sql = sql & "SET dateformat DMY INSERT INTO d_supplier_deduc(SNo, Date_Deduc, Description, Amount, Period, StartDate, EndDate, auditid, auditdatetime, yyear, Remarks, Branch, AI) VALUES ("
'              'sql = sql & txtSNo & ",'" & Date & "','SMS ','1.5','" & month(DTPMilkDate) & "','" & Startdate & "','" & Enddate & "','" & User & "','" & Date & "','" & year(DTPMilkDate) & "','Charges','" & cbolocation & "','0')"
'              'oSaccoMaster.ExecuteThis (sql)
'
'             '''''''''END OF sms charge''''''''
''           End If
'            End If
' 'Else
'
' 'End If
'    '******END OF SMS*******
'''************message***************'
''              '//check settings  for sms alert
''        'If SendSmsOnSalary = True Then
''        'Dim phone As String
''        Dim RsPhoneNumbers As New ADODB.Recordset
''        Dim MsgContent As String
''        Set RsPhoneNumbers = oSaccoMaster.GetRecordset("select  PhoneNo  from d_Suppliers  where SNo ='" & txtSNo & "' and PhoneNo  is not null and PhoneNo <>'' and LEN(PhoneNo) =10 ")
''
''        If Not RsPhoneNumbers.EOF Then
''            Phone = Trim(RsPhoneNumbers.Fields!PhoneNo & "")
''        Else
''            Phone = ""
''        End If
'''          '//insert into sms alert
'        Set rs = New ADODB.Recordset
'        sql = "d_sp_TotalMonth " & txtSNo & ",'" & Startdate & "','" & DTPMilkDate & "'"
'        Set rs = oSaccoMaster.GetRecordset(sql)
'        If Not rs.EOF Then
'        If Not IsNull(rs.Fields(0)) Then CummulKgs = rs.Fields(0)
'        Else
'        CummulKgs = "0.00"
'        End If
'
'            If Trim(Phone) <> "" Then
'            MsgContent = ""
'            MsgContent = "Dear supplier, You have supplied," & txtQnty & " Kgs Today at " & Time & ",Cummulative This Month :  " & Format(CummulKgs, "#,##0.00" & "Kgs") ' From Soitaran Mcs Dairy"
'            strSQL = ""
'            strSQL = "INSERT INTO Swift_Messages(SaccoCode,PhoneNumber,Message, msgstatus)"
'            strSQL = strSQL & "Values (14,'" & Phone & "','" & MsgContent & "',0)"
'
'            oSaccoMaster.ExecuteThis (strSQL)
'            End If

Transport = 0
agrovet = 0
BONUS = 0
TMShares = 0
FSA = 0
HShares = 0
Advance = 0
Others = 0
'


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
    txtFile.WriteLine "             Milk Receipt"
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
    
    txtFile.WriteLine "SNo :" & txtSNo
    txtFile.WriteLine "Name :" & lblNames
    txtFile.WriteLine "Quantity Supplied :" & txtQnty & " Kgs"
    txtFile.WriteLine "Price" & Price & " Per Kgs"
    Set rs = New ADODB.Recordset
    sql = "d_sp_TotalMonth " & txtSNo & ",'" & Startdate & "','" & DTPMilkDate & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
    If Not IsNull(rs.Fields(0)) Then CummulKgs = rs.Fields(0)
    Else
    CummulKgs = "0.00"
    End If
    txtFile.WriteLine "Cummulative This Month : " & Format(CummulKgs, "#,##0.00" & " Kgs")
    Set rs = New ADODB.Recordset
    sql = "d_sp_TransName '" & txtSNo & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
    If Not IsNull(rs.Fields(0)) Then TRANSPORTER = rs.Fields(0)
    Else
        TRANSPORTER = "Self"
    End If
    txtFile.WriteLine "---------------------------------------"
    txtFile.WriteLine "Receipt Number :" & RNumber
    txtFile.WriteLine "TRANSPORTER :" & TRANSPORTER
    txtFile.WriteLine "Received by :" & username
    txtFile.WriteLine "---------------------------------------"
    txtFile.WriteLine "  Date :" & Format(DTPMilkDate, "dd/mm/yyyy") & " ,Time : " & Format(Time, "hh:mm:ss AM/PM")
    txtFile.WriteLine "       " & motto & ""
    txtFile.WriteLine "---------------------------------------"
    If chkComment.value = vbChecked Then
        txtFile.WriteLine txtComment
        txtFile.WriteLine "---------------------------------------"
    End If
    txtFile.WriteLine escFeedAndCut
    
 txtFile.Close
 Reset
End If


'* writing to notepad
If chkNotepad.value = vbChecked Then

'    Dim fso, txtfile
'    Dim ttt
'     Dim escFeedAndCut As String
     escFeedAndCut = Chr(29) + Chr(86) + Chr(65) '//Partial cut and feed
       cdgPrint.Filter = "*.csv|*.txt"
        cdgPrint.ShowSave
        ttt = cdgPrint.FileName
        If ttt = "" Then
        MsgBox "File should not be blank", vbCritical, "Data transfer"
        Exit Sub
        End If
        'ttt = "D:\PROJECTS\FOSA\DAILY" & Date & ""
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set txtFile = fso.CreateTextFile(ttt, True)
        txtFile.WriteLine
        
    txtFile.WriteLine "---------------------------------------"
    txtFile.WriteLine "" & cname & ""
   ' Printer.Print Tab(0); "Kimathi House Branch"
    txtFile.WriteLine " " & paddress & " "
    txtFile.WriteLine "" & town & ""
    txtFile.WriteLine "Milk Receipt"
    txtFile.WriteLine "---------------------------------------"
'    If cbomemtrans = "Shares" Then
'    DESC = bosanames & " -Member No " & memberno
    txtFile.WriteLine "SNo :" & txtSNo
    txtFile.WriteLine "Name :" & lblNames
'    Else
    txtFile.WriteLine "Quantity Supplied :" & txtQnty & " Kgs"
    txtFile.WriteLine "Price" & Price & " Per Kgs"
    Startdate = DateSerial(Year(DTPMilkDate), month(DTPMilkDate) - 1, 1)
    'sql = "d_sp_TotalMonth " & txtSNo & ",'" & StartDate & "','" & DTPMilkDate & "'"
    Set rs = New ADODB.Recordset
    sql = "d_sp_TotalMonth " & txtSNo & ",'" & Startdate & "','" & DTPMilkDate & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
    If Not IsNull(rs.Fields(0)) Then CummulKgs = rs.Fields(0)
    Else
    CummulKgs = "0.00"
    End If
    txtFile.WriteLine "Cummulative This Month " & Format(CummulKgs, "#,##0.00" & " Kgs")
'    End If
    Set rs = New ADODB.Recordset
    sql = "d_sp_TransName '" & txtSNo & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
    If Not IsNull(rs.Fields(0)) Then TRANSPORTER = rs.Fields(0)
    Else
    TRANSPORTER = "Self"
    End If
    txtFile.WriteLine "---------------------------------------"
    txtFile.WriteLine "Transporter :" & TRANSPORTER
    txtFile.WriteLine "Received by :" & username
    txtFile.WriteLine "---------------------------------------"
    txtFile.WriteLine "Date :" & Format(Get_Server_Date, "dd/mm/yyyy HH:MM:SS AM/PM")
    txtFile.WriteLine "     " & motto & " "
    txtFile.WriteLine "---------------------------------------"
    txtFile.WriteLine escFeedAndCut
    'Printer.Print
    'Printer.CurrentX = 500#
    'Printer.FontSize = 10
    'Printer.CurrentX = 500
    'Printer.FontSize = 8
'    Printer.Print Tab(0); "Date :"; Tab(10); Format(Get_Server_Date, "dd/mm/yyyy HH:MM:SS AM/PM")
'    Printer.Print Tab(0); "     TDL - Improves Your Value "
'    Printer.Print Tab(0); "---------------------------------------"
'    Printer.CurrentX = 500
'    Printer.FontSize = 8
'    Printer.Print
'    Printer.CurrentX = 500
'    Printer.FontSize = 8
'    Printer.Print
txtFile.Close
End If

loadMilk

txtSNo = ""
txtQnty = ""
txttoday = ""
'txtSNo_Validate True
txtSNo.SetFocus
Exit Sub
ErrorHandler:

MsgBox err.description

End Sub

Public Sub loadMilk()
Startdate = DateSerial(Year(DTPMilkDate), month(DTPMilkDate), Day(DTPMilkDate))
Enddate = DateSerial(Year(DTPMilkDate), month(DTPMilkDate), Day(DTPMilkDate) + 1)
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
'    daily supplier
'     lblDTotal.Caption = "0"
    Set rs = New ADODB.Recordset
'    sql = "d_sp_DailyTotal1 '" & DTPMilkDate & "', sno='" & txtSNo & "'"""
'    Set rs = oSaccoMaster.GetRecordset(sql)
'    If Not rs.EOF Then
'    If Not IsNull(rs.Fields(0)) Then lbltoday.Caption = rs.Fields(0)
'    Else
'    lbltoday.Caption = "0"
'    End If

sql = "d_sp_DailyTotal3 '" & Startdate & "','" & Enddate & "','" & cbolocation & "'"
   ' sql = "d_sp_DailyTotal3 '" & DTPMilkDate & "','" & cbolocation & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
    If Not IsNull(rs.Fields(0)) Then lblDTotalb.Caption = rs.Fields(0)
    Else
    lblDTotalb.Caption = "0"
    End If
    
    With lvwMilkCollected
    
       .ListItems.Clear
    
        .ColumnHeaders.Clear

  End With

    Set rs = CreateObject("adodb.recordset")
  
    sql = "d_sp_CurrentList '" & Format(DTPMilkDate, "dd/mm/yyyy") & "','" & User & "'"
    
   Set rs = oSaccoMaster.GetRecordset(sql)
    
    With lvwMilkCollected
        
        
        .ColumnHeaders.Add , , "SNo", 2000
        .ColumnHeaders.Add , , "Names", 3000
        .ColumnHeaders.Add , , "QNTY"
        .ColumnHeaders.Add , , "Time"
        .ColumnHeaders.Add , , "Receipt No.", 2000
    
        While Not rs.EOF
            If Not IsNull(rs.Fields("SNo")) Then
            Set li = .ListItems.Add(, , Trim(rs.Fields("SNo")))
            End If
            If Not IsNull(rs.Fields("Names")) Then
            li.ListSubItems.Add , , Trim(rs.Fields("Names"))
            End If
            If Not IsNull(rs.Fields("QSupplied")) Then
            li.ListSubItems.Add , , Trim(rs.Fields("QSupplied"))
            End If
            If Not IsNull(rs.Fields("TransTime")) Then
            li.ListSubItems.Add , , Trim(rs.Fields("TransTime"))
            End If
            If Not IsNull(rs.Fields("ID")) Then
            li.ListSubItems.Add , , Trim(rs.Fields("ID"))
            End If

            
                    rs.MoveNext
        
        Wend
        
    End With
    
    rs.Close
    
    Set rs = Nothing
    
lvwMilkCollected.View = lvwReport

End Sub

Private Sub cmdRemove_Click()
Dim Valu As String
Dim TransTotalDed As Currency
Dim DESCR As String
Dim amount As Currency
    

'InputBox
Valu = InputBox("Please enter the receipt number", "REMOVE INTAKE", "<Enter receipt number here>")

If Valu = "" Then
    MsgBox "Please enter receipt number "
    Exit Sub
End If
    
If Valu = "<Enter receipt number here>" Then
    MsgBox "Please enter receipt number "
    Exit Sub
End If
    
If Not IsNumeric(Valu) Then
    MsgBox "Please enter receipt number." & Valu & " is not a number", vbCritical
    Exit Sub
End If

Set rs = New ADODB.Recordset
sql = "d_sp_SpecificReceipt " & Valu & ""
Set rs = oSaccoMaster.GetRecordset(sql)

Startdate = DateSerial(Year(rs.Fields("TransDate")), month(rs.Fields("TransDate")), 1)
Enddate = DateSerial(Year(rs.Fields("TransDate")), month(rs.Fields("TransDate")) + 1, 1 - 1)

sql = "d_sp_DeleteMIn " & Valu & ""
oSaccoMaster.ExecuteThis (sql)

'//Delete Daily Intake

Set rs2 = New ADODB.Recordset
sql = "SET dateformat DMY SELECT [" & Day(rs.Fields("TransDate")) & "] AS Milk from d_DailySummary WHERE SNo =" & rs.Fields("SNo") & " AND Endperiod ='" & Enddate & "'"
Set rs2 = oSaccoMaster.GetRecordset(sql)
If Not rs2.EOF Then
If ((rs2.Fields(0)) > 0) Then
sql = "SET dateformat DMY Update  d_DailySummary SET [" & Day(DTPMilkDate) & "] = " & CCur(rs2!milk) - CCur(rs.Fields("QSupplied")) & " WHERE SNo =" & rs.Fields("SNo") & " AND Endperiod ='" & Enddate & "'"
oSaccoMaster.ExecuteThis (sql)
Else
sql = "SET dateformat DMY Update  d_DailySummary SET [" & Day(DTPMilkDate) & "] = '' WHERE SNo =" & rs.Fields("SNo") & " AND Endperiod ='" & Enddate & "'"
oSaccoMaster.ExecuteThis (sql)
End If
End If


Set Rs1 = New ADODB.Recordset
sql = "d_sp_TransActive " & rs.Fields("SNo") & ",'" & rs.Fields("TransDate") & "'"
Set Rs1 = oSaccoMaster.GetRecordset(sql)
    If Not Rs1.EOF Then
        '//Update deductions
        Set cn = New ADODB.Connection

        Dim r2  As String
        r2 = "Transport"
            sql = "d_sp_DeleteSupDeduc " & rs.Fields("SNo") & "," & Rs1!rate * rs.Fields("QSupplied") & ",'" & r2 & "','" & rs.Fields("TransDate") & "'"
        oSaccoMaster.ExecuteThis (sql)
        
        transdate = Rs1!Startdate
        If transdate < Startdate Then
            transdate = Startdate
        End If
        
        
        '---d_sp_SelTransDet @SNo bigint,@StartDate varchar(10), @Enddate varchar(10)
        Set rst = New ADODB.Recordset
            sql = "d_sp_SelTransDet " & rs.Fields("SNo") & ",'" & transdate & "','" & Enddate & "'"
        Set rst = oSaccoMaster.GetRecordset(sql)
        
    '--d_sp_UpdateDetTrans @SNo bigint,@QNTY float,@Amnt money,@Code varchar(35),@Subsidy money,@EPeriod varchar(10),@user varchar(35)
        If Not rst.EOF Then
            Set cn = New ADODB.Connection
                sql = "d_sp_UpdateDetTrans " & rs.Fields("SNo") & "," & rst!qnty & "," & rst!amount & ",'" & rst!code & "'," & rst!subsidy & ",'" & Enddate & "','" & User & "'"
            oSaccoMaster.ExecuteThis (sql)
            
        
        Else
            Set cn = New ADODB.Connection
                sql = "delete d_TransDetailed WHERE SNO = " & rs.Fields("SNo") & " AND EndPeriod ='" & Enddate & "'"
            oSaccoMaster.ExecuteThis (sql)
        End If
        '----d_sp_SelTransGPayQnty @EP varchar(10), @Code varchar(35)
        Set rst2 = New ADODB.Recordset
            sql = "d_sp_SelTransGPayQnty '" & Enddate & "','" & rst!code & "'"
        Set rst2 = oSaccoMaster.GetRecordset(sql)
        
    '-- d_sp_UpdateTransPay @Code varchar(35), @Qnty float,@Amnt money,@Subsidy money,@GrossPay money, @EndDate varchar(10)  AS
    If Not rst2.EOF Then
            Set cn = New ADODB.Connection
                sql = "d_sp_UpdateTransPay '" & rst!code & "'," & rst2!qnty & "," & rst2!Amnt & "," & rst2!subsidy & "," & rst2!GPay & ",'" & Enddate & "','" & User & "'"
            oSaccoMaster.ExecuteThis (sql)
            
        End If
 
        '---Get transporters Total Deductions//d_sp_TotalTransDeduct @Code varchar(35),@Month bigint,@Year bigint
        Set Rs1 = New ADODB.Recordset
            sql = "d_sp_TotalTransDeduct '" & rst!code & "'," & month(rs.Fields("transdate")) & "," & Year(rs.Fields("transdate")) & ""
        Set Rs1 = oSaccoMaster.GetRecordset(sql)
    If Not Rs1.EOF Then
    'Dim TransTotalDed As Currency
    If Not IsNull(Rs1.Fields(0)) Then TransTotalDed = Rs1.Fields(0)
    End If
    
    Set rs3 = New ADODB.Recordset
    'Dim Startdate As String, Enddate As String
    'Dim DESCR As String
   ' Dim amount As Currency
    
    '--d_sp_SelTransDed @Code varchar(35), @startdate varchar(10),@enddate varchar(10)
    sql = "d_sp_SelTransDed '" & rst!code & "','" & Startdate & "','" & Enddate & "'"
    Set rs3 = oSaccoMaster.GetRecordset(sql)
If Not rs3.EOF Then
    While Not rs3.EOF
    DESCR = Trim(rs3.Fields(0))
    amount = 0
    amount = rs3.Fields(1)
    sql = "SELECT     Agrovet, BONUS, TMShares, FSA, HShares, Advance, Others FROM d_TransportersPayroll WHERE Code='" & rst!code & "' AND EndPeriod ='" & Enddate & "'"
    Set rs4 = oSaccoMaster.GetRecordset(sql)
     If UCase(rs4.Fields(0).name) = UCase(DESCR) Then
        agrovet = amount
    End If
    If UCase(rs4.Fields(1).name) = UCase(DESCR) Then
        BONUS = amount
    End If
    If UCase(rs4.Fields(2).name) = UCase(DESCR) Then
        TMShares = amount
    End If
    If UCase(rs4.Fields(3).name) = UCase(DESCR) Then
        FSA = amount
    End If
    If UCase(rs4.Fields(4).name) = UCase(DESCR) Then
        HShares = amount
    End If
    If UCase(rs4.Fields(5).name) = UCase(DESCR) Then
        Advance = amount
    End If
    If UCase(rs4.Fields(6).name) = UCase(DESCR) Then
        Others = amount
    End If

    '//d_sp_UpdatePayrollDed  @SNo bigint,@EndPeriod varchar(15),@Transport money,@Agrovet money,@AI money,@TMShares money,@FSA money,@HShares money,@Advance money,@Others
    rs3.MoveNext
    Wend
    ' d_sp_UpdateTransDed  @Code varchar(35),@EndPeriod varchar(15),@TotalDed money,@Agrovet money,@AI money,@TMShares money,@FSA money,@HShares money,@Advance money,@Others money
    Set cn = New ADODB.Connection
    sql = "d_sp_UpdateTransDed  '" & rst!code & "','" & Enddate & "'," & TransTotalDed & "," & agrovet & "," & BONUS & "," & TMShares & "," & FSA & "," & HShares & "," & Advance & "," & Others & ""
    oSaccoMaster.ExecuteThis (sql)
End If

    agrovet = 0
    BONUS = 0
    TMShares = 0
    FSA = 0
    HShares = 0
    Advance = 0
    Others = 0
    TransTotalDed = 0
 
    End If

Set Rs1 = New ADODB.Recordset
sql = "d_sp_TransInActive " & rs.Fields("SNo") & ",'" & rs.Fields("TransDate") & "'"
Set Rs1 = oSaccoMaster.GetRecordset(sql)
If Not Rs1.EOF Then
''//Update deductions
sql = "d_sp_DeleteSupDeduc " & rs.Fields("SNo") & "," & Rs1!rate * rs.Fields("QSupplied") & ",'Transport','" & rs.Fields("TransDate") & "'"
oSaccoMaster.ExecuteThis (sql)


        '---d_sp_SelTransDet @SNo bigint,@StartDate varchar(10), @Enddate varchar(10)
        Set rst = New ADODB.Recordset
            sql = "d_sp_SelTransDet " & rs.Fields("SNo") & ",'" & transdate & "','" & Enddate & "'"
        Set rst = oSaccoMaster.GetRecordset(sql)
        
    '--d_sp_UpdateDetTrans @SNo bigint,@QNTY float,@Amnt money,@Code varchar(35),@Subsidy money,@EPeriod varchar(10),@user varchar(35)
        If Not rst.EOF Then
            Set cn = New ADODB.Connection
                sql = "d_sp_UpdateDetTrans " & rs.Fields("SNo") & "," & rst!qnty & "," & rst!amount & ",'" & rst!code & "'," & rst!subsidy & ",'" & Enddate & "','" & User & "'"
            oSaccoMaster.ExecuteThis (sql)
            
        Else
            Set cn = New ADODB.Connection
                sql = "delete d_TransDetailed WHERE SNO = " & rs.Fields("SNo") & " AND EndPeriod ='" & Enddate & "'"
            oSaccoMaster.ExecuteThis (sql)
            
        End If
        
                '----d_sp_SelTransGPayQnty @EP varchar(10), @Code varchar(35)
        Set rst2 = New ADODB.Recordset
            sql = "d_sp_SelTransGPayQnty '" & Enddate & "','" & rst!code & "'"
        Set rst2 = oSaccoMaster.GetRecordset(sql)
        
    '-- d_sp_UpdateTransPay @Code varchar(35), @Qnty float,@Amnt money,@Subsidy money,@GrossPay money, @EndDate varchar(10)  AS
    If Not rst2.EOF Then
            Set cn = New ADODB.Connection
                sql = "d_sp_UpdateTransPay '" & rst!code & "'," & rst2!qnty & "," & rst2!Amnt & "," & rst2!subsidy & "," & rst2!GPay & ",'" & Enddate & "','" & User & "'"
            oSaccoMaster.ExecuteThis (sql)
            
        End If
 
        '---Get transporters Total Deductions//d_sp_TotalTransDeduct @Code varchar(35),@Month bigint,@Year bigint
        Set Rs1 = New ADODB.Recordset
            sql = "d_sp_TotalTransDeduct '" & rst!code & "'," & month(rs.Fields("transdate")) & "," & Year(rs.Fields("transdate")) & ""
        Set Rs1 = oSaccoMaster.GetRecordset(sql)
    If Not Rs1.EOF Then
    'Dim TransTotalDed As Currency
    If Not IsNull(Rs1.Fields(0)) Then TransTotalDed = Rs1.Fields(0)
    End If
    
    Set rs3 = New ADODB.Recordset
    'Dim Startdate As String, Enddate As String
    'Dim DESCR As String
   ' Dim amount As Currency
    
    '--d_sp_SelTransDed @Code varchar(35), @startdate varchar(10),@enddate varchar(10)
    sql = "d_sp_SelTransDed '" & rst!code & "','" & Startdate & "','" & Enddate & "'"
    Set rs3 = oSaccoMaster.GetRecordset(sql)
If Not rs3.EOF Then
    While Not rs3.EOF
    DESCR = Trim(rs3.Fields(0))
    amount = 0
    amount = rs3.Fields(1)
    sql = "SELECT     Agrovet, BONUS, TMShares, FSA, HShares, Advance, Others FROM d_TransportersPayroll WHERE Code='" & rst!code & "' AND EndPeriod ='" & Enddate & "'"
    Set rs4 = oSaccoMaster.GetRecordset(sql)
     If UCase(rs4.Fields(0).name) = UCase(DESCR) Then
        agrovet = amount
    End If
    If UCase(rs4.Fields(1).name) = UCase(DESCR) Then
        BONUS = amount
    End If
    If UCase(rs4.Fields(2).name) = UCase(DESCR) Then
        TMShares = amount
    End If
    If UCase(rs4.Fields(3).name) = UCase(DESCR) Then
        FSA = amount
    End If
    If UCase(rs4.Fields(4).name) = UCase(DESCR) Then
        HShares = amount
    End If
    If UCase(rs4.Fields(5).name) = UCase(DESCR) Then
        Advance = amount
    End If
    If UCase(rs4.Fields(6).name) = UCase(DESCR) Then
        Others = amount
    End If

    '//d_sp_UpdatePayrollDed  @SNo bigint,@EndPeriod varchar(15),@Transport money,@Agrovet money,@AI money,@TMShares money,@FSA money,@HShares money,@Advance money,@Others
    rs3.MoveNext
    Wend
    ' d_sp_UpdateTransDed  @Code varchar(35),@EndPeriod varchar(15),@TotalDed money,@Agrovet money,@AI money,@TMShares money,@FSA money,@HShares money,@Advance money,@Others money
    Set cn = New ADODB.Connection
    sql = "d_sp_UpdateTransDed  '" & rst!code & "','" & Enddate & "'," & TransTotalDed & "," & agrovet & "," & BONUS & "," & TMShares & "," & FSA & "," & HShares & "," & Advance & "," & Others & ""
    oSaccoMaster.ExecuteThis (sql)
End If

    agrovet = 0
    BONUS = 0
    TMShares = 0
    FSA = 0
    HShares = 0
    Advance = 0
    Others = 0
    TransTotalDed = 0
End If


'//d_sp_TotalDeduct-Total Deductions
'//d_sp_UpdateGPAYQnty - Total Grosspay and Quantity
'//d_sp_SupDed - Supply Deductions



Set rs2 = New ADODB.Recordset
'Dim Startdate As String, Enddate As String
Dim qnty As Currency, GPay As Currency
'Startdate = DateSerial(DTPMilkDate, cboMonth, 1)
sql = "d_sp_UpdateGPAYQnty '" & Startdate & "','" & Enddate & "'," & rs.Fields("SNo") & ""
Set rs2 = oSaccoMaster.GetRecordset(sql)
If Not rs2.EOF Then
If Not IsNull(rs2.Fields(0)) Then qnty = rs2.Fields(0)
If Not IsNull(rs2.Fields(1)) Then GPay = rs2.Fields(1)
End If


Set Rs1 = New ADODB.Recordset
sql = "d_sp_TotalDeduct " & rs.Fields("SNo") & "," & month(DTPMilkDate) & "," & Year(DTPMilkDate) & ""
Set Rs1 = oSaccoMaster.GetRecordset(sql)
If Not Rs1.EOF Then
Dim TotalDed As Currency
If Not IsNull(Rs1.Fields(0)) Then TotalDed = Rs1.Fields(0)
End If
'//Update payroll -- @SNo bigint,@EndPeriod varchar(15),@Kgs float,@GPay money,@NPay money,@TDeductions money,@auditid  varchar(35)
Set cn = New ADODB.Connection
sql = "d_sp_UpdatePayroll  " & rs.Fields("SNo") & ",'" & Enddate & "'," & qnty & "," & GPay & "," & GPay - TotalDed & "," & TotalDed & ",'" & User & "'"
oSaccoMaster.ExecuteThis (sql)



Set rs3 = New ADODB.Recordset
''Dim Startdate As String, Enddate As String
Dim desc As String
Dim Amnt As Currency
'Startdate = DateSerial(Year(DTPMilkDate), month(DTPMilkDate), 1)
'Enddate = DateSerial(Year(DTPMilkDate), month(DTPMilkDate) + 1, 1 - 1)
sql = "d_sp_SupDed " & rs.Fields("SNo") & ",'" & Startdate & "','" & Enddate & "'"
Set rs3 = oSaccoMaster.GetRecordset(sql)
If Not rs3.EOF Then
While Not rs3.EOF
desc = Trim(rs3.Fields(0))
Amnt = 0
Amnt = rs3.Fields(1)
sql = "SELECT     Transport, Agrovet, BONUS, TMShares, FSA, HShares, Advance, Others FROM d_Payroll WHERE SNo=" & rs.Fields("SNo") & " AND EndofPeriod ='" & Enddate & "'"
Set rs4 = oSaccoMaster.GetRecordset(sql)
If UCase(rs4.Fields(0).name) = UCase(desc) Then
Transport = Amnt
End If
If UCase(rs4.Fields(1).name) = UCase(desc) Then
agrovet = Amnt
End If
If UCase(rs4.Fields(2).name) = UCase(desc) Then
BONUS = Amnt
End If
If UCase(rs4.Fields(3).name) = UCase(desc) Then
TMShares = Amnt
End If
If UCase(rs4.Fields(4).name) = UCase(desc) Then
FSA = Amnt
End If
If UCase(rs4.Fields(5).name) = UCase(desc) Then
HShares = Amnt
End If
If UCase(rs4.Fields(6).name) = UCase(desc) Then
Advance = Amnt
End If
If UCase(rs4.Fields(7).name) = UCase(desc) Then
Others = Amnt
End If

'//d_sp_UpdatePayrollDed  @SNo bigint,@EndPeriod varchar(15),@Transport money,@Agrovet money,@AI money,@TMShares money,@FSA money,@HShares money,@Advance money,@Others
rs3.MoveNext
Wend
Else
Transport = 0
'//Update Deductions -- d_sp_UpdatePayrollDed  @SNo bigint,@EndPeriod varchar(15),@Transport money,@Agrovet money,@AI money,@TMShares money,@FSA money,@HShares money,@Advance money,@Others money
Set cn = New ADODB.Connection
sql = "d_sp_UpdatePayrollDed  " & rs.Fields("SNo") & ",'" & Enddate & "'," & Transport & "," & agrovet & "," & BONUS & "," & TMShares & "," & FSA & "," & HShares & "," & Advance & "," & Others & ""
oSaccoMaster.ExecuteThis (sql)
End If

Dim Yr As Integer

Yr = Year(DTPMilkDate)
'vbHourglass
'Fixed deductions update
oSaccoMaster.ExecuteThis ("d_sp_PresetDeductAssign_99 '" & Startdate & "','" & Enddate & "'," & Yr & ",'" & User & "', " & rs.Fields("SNo"))

'Payroll update
'd_sp_GDedNet @StartDate varchar(10) , @endPeriod varchar(10)
oSaccoMaster.ExecuteThis ("d_sp_GDedNet_99 '" & Startdate & "','" & Enddate & "'," & rs.Fields("SNo"))

'Update transporters
'd_sp_TransUpdate @StartDate varchar(10),@EndPeriod varchar(10),@User varchar(35) AS
oSaccoMaster.ExecuteThis ("d_sp_TransUpdate_99 '" & Startdate & "','" & Enddate & "','" & User & "'," & rs.Fields("SNo"))


oSaccoMaster.ExecuteThis ("d_sp_TransPRoll '" & Startdate & "','" & Enddate & "','" & User & "'")
'Lock period
'[d_sp_Periods]

loadMilk
MsgBox "Record removed successfully."

Transport = 0
agrovet = 0
BONUS = 0
TMShares = 0
FSA = 0
HShares = 0
Advance = 0
Others = 0

    
End Sub

Private Sub cmdreprintreceipt_Click()
On Error GoTo ErrorHandler
Dim Price As Currency
Dim Startdate, CummulKgs, TRANSPORTER As String
Dim transdate As Date
Startdate = DateSerial(Year(DTPMilkDate), month(DTPMilkDate), 1)
Enddate = DateSerial(Year(DTPMilkDate), month(DTPMilkDate) + 1, 1 - 1)
    
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
        Dim PORT As String
 '      PORT = ports
        'ttt = "LPT1" 'LPT1
 '      ttt = PORT
        ttt = ports
        'ttt = "LPT1" 'LPT1
        'ttt = "D:\PROJECTS\FOSA\DAILY" & Date & ""
        Set fso = CreateObject("Scripting.FileSystemObject")
        'Set chkPrinter = fso.GetFile(ttt)
       
        
    Set txtFile = fso.CreateTextFile(ttt, True)
    txtFile.WriteLine "      " & cname & ""
    txtFile.WriteLine "      " & paddress & " " & town & ""
    txtFile.WriteLine "      " & Phone & ""
    txtFile.WriteLine "             Milk Receipt"
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
    
    txtFile.WriteLine "SNo :" & txtSNo
    txtFile.WriteLine "Name :" & lblNames
    '//get the kilo supplied
      Set rs = New ADODB.Recordset
    sql = "d_sp_Totalday  " & txtSNo & ",'" & DTPMilkDate & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
    If Not IsNull(rs.Fields(0)) Then txtQnty = rs.Fields(0)
    Else
    txtQnty = "0.00"
    End If
    txtFile.WriteLine "Quantity Supplied :" & txtQnty & " Kgs"
    txtFile.WriteLine "Price" & Price & " Per Kgs"
    Set rs = New ADODB.Recordset
    sql = "d_sp_TotalMonth " & txtSNo & ",'" & Startdate & "','" & DTPMilkDate & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
    If Not IsNull(rs.Fields(0)) Then CummulKgs = rs.Fields(0)
    Else
    CummulKgs = "0.00"
    End If
    txtFile.WriteLine "Cummulative This Month : " & Format(CummulKgs, "#,##0.00" & " Kgs")
    Set rs = New ADODB.Recordset
    sql = "d_sp_TransName '" & txtSNo & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
    If Not IsNull(rs.Fields(0)) Then TRANSPORTER = rs.Fields(0)
    Else
        TRANSPORTER = "Self"
    End If
    txtFile.WriteLine "---------------------------------------"
    txtFile.WriteLine "Receipt Number :" & RNumber
    txtFile.WriteLine "TRANSPORTER :" & TRANSPORTER
    txtFile.WriteLine "Received by :" & username
    txtFile.WriteLine "---------------------------------------"
    txtFile.WriteLine "  Date :" & DTPMilkDate & " ,Time : " & Format(Time, "hh:mm:ss AM/PM")
    txtFile.WriteLine "       " & motto & ""
    txtFile.WriteLine "---------------------------------------"
     txtFile.WriteLine "********POWERED BY EASYMA***************"
    If chkComment.value = vbChecked Then
        txtFile.WriteLine txtComment
        txtFile.WriteLine "---------------------------------------"
    End If
    txtFile.WriteLine escFeedAndCut
    
 txtFile.Close
 Reset
Exit Sub
ErrorHandler:
MsgBox err.description, vbCritical
End Sub

Private Sub cmdsumaintake2_Click()
'cmdsumaintake.Show vbModal
Dim ans As String
ans = MsgBox("Do you Want List of Suppliers from All Branches?", vbYesNo)
If ans = vbYes Then
 reportname = "d_Dailysummary2.rpt"

 Else
 reportname = "d_Dailysummary.rpt"
'  ReportTitle = "TO :" & UCase(Location) & " ; " & vbNewLine & " Please pay the following transporters the amount indicated: (Our Ref is code)"
'    '{d_Payroll.NPay} > 0 and {d_Payroll.Bank} <> '' and month({d_Payroll.EndofPeriod})= month(30/09/2010)  AND year({d_Payroll.EndofPeriod}) = Year(30/09/2010)
' STRFORMULA = "{d_TransportersPayRoll.NetPay} > 0 and {d_TransportersPayRoll.BankName} = '" & cboBank & "' and month({d_TransportersPayRoll.EndPeriod})=" & month(dtpEndPeriod) & " AND year({d_TransportersPayRoll.EndPeriod}) =" & year(dtpEndPeriod)
' Show_Sales_Crystal_Report STRFORMULA, reportname, ReportTitle
 End If
 Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub cmdupdatedeductions_Click()
On Error GoTo ErrorHandler
Dim IDENTI As Long


Set rs = oSaccoMaster.GetRecordset("d_sp_IsClosed '" & DTPMilkDate & "'")
If Not rs.EOF Then
    MsgBox "The period " & Enddate & " has been closed by " & rs.Fields(0)
    Exit Sub
End If
Dim sno As String, Date_Deduc As Date, description As String, amount As Double, tbl As String
'Id, SNo, Date_Deduc, Description, Amount, auditid
 For I = 1 To Lvwdeductions.ListItems.Count
 
        If Lvwdeductions.ListItems.Item(I).Checked = True Then
        Set li = Lvwdeductions.ListItems(I)
        IDENTI = li
        tbl = "d_supplier_deduc"
        amount = CDbl(Lvwdeductions.ListItems(I).SubItems(4))
        Date_Deduc = Lvwdeductions.ListItems(I).SubItems(2)
        sno = Lvwdeductions.ListItems(I).SubItems(1)
        description = Lvwdeductions.ListItems(I).SubItems(3)
        '//before delete update audit report
            sql = ""
            sql = "set dateformat dmy INSERT INTO AUDITTRANS"
             sql = sql & "         (TransTable, TransDescription, TransDate, Amount, AuditID, AuditTime)"
             sql = sql & " VALUES     ('" & tbl & "','" & description & "','" & Date_Deduc & "'," & amount & ",'" & User & "','" & Get_Server_Date & "')"
             oSaccoMaster.ExecuteThis (sql)
             
        'identi = Lvwdeductions.ListItems(I).SubItems(1)
            sql = ""
            sql = "delete from d_supplier_deduc  where id =" & IDENTI & ""
            oSaccoMaster.ExecuteThis (sql)
        End If
    Next I
MsgBox "Record Sucessfully Updated"
Exit Sub
ErrorHandler:
MsgBox err.description
End Sub

Private Sub Command1_Click()
frmOutletsale.Show vbModal
End Sub

Private Sub DTPMilkDate_Change()
loadMilk
End Sub

Private Sub DTPMilkDate_Click()
loadMilk
End Sub

Private Sub Form_Load()

DTPMilkDate = Format(Get_Server_Date, "dd/mm/yyyy")
DTPMilkDate.MaxDate = Format(Get_Server_Date, "dd/mm/yyyy")
DTPcomplaintperiod = DTPMilkDate
With StatusBar1.Panels
    .Item(1).Text = "USER : " & username
    .Item(2).Text = "DATE : " & Format(Get_Server_Date, "dd/mm/yyyy")

End With
'   Set rst = New Recordset
'    Dim cn As Connection
'    Set cn = New ADODB.Connection
'    Provider = "MAZIWA"
'    'Provider = cnn

'    cn.Open Provider, "atm", "atm"
    Set rst = New Recordset
    
    sql = "Select Bname from   d_Branch"
    'rst.Open sql, cn, adOpenKeyset, adLockOptimistic
    Set rst = oSaccoMaster.GetRecordset(sql)
    While Not rst.EOF
    cbolocation.AddItem rst.Fields(0)
    rst.MoveNext
    Wend
    
loadMilk


End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'lblWait.Visible = True
'Timer1.Enabled = True
Startdate = DateSerial(Year(DTPMilkDate), month(DTPMilkDate), 1)
Enddate = DateSerial(Year(DTPMilkDate), month(DTPMilkDate) + 1, 1 - 1)
'
'oSaccoMaster.ExecuteThis ("d_sp_TransUpdate '" & Startdate & "','" & Enddate & "','" & User & "'")
'oSaccoMaster.ExecuteThis ("d_sp_TransPRoll '" & Startdate & "','" & Enddate & "','" & User & "'")

MainForm.Caption = "EasyMa "
End Sub

Private Sub Timer1_Timer()
End Sub

Private Sub txtComment_Click()
If txtComment = "<Put your comment here>" Then
txtComment = ""
End If
End Sub




Private Sub txtQnty_KeyPress(KeyAscii As Integer)
'txttoday = "0"
On Error GoTo ErrorHandler
If txtSNo = "" Then
MsgBox "Please key in the Member Number"
Exit Sub
End If
Set rs = New ADODB.Recordset
Startdate = DateSerial(Year(DTPMilkDate), month(DTPMilkDate), Day(DTPMilkDate))
Enddate = DateSerial(Year(DTPMilkDate), month(DTPMilkDate), Day(DTPMilkDate) + 1)
sql = "d_sp_DailyTotal1 '" & Startdate & "','" & Enddate & "','" & txtSNo & "'"
'sql = "d_sp_DailyTotal1 '" & DTPMilkDate & "','" & txtSNo & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
    If Not IsNull(rs.Fields(0)) Then txttoday = rs.Fields(0)
    Else
    txttoday = "0"
    End If
    
'''*****************************check if supliers branch is correct ****
sql = ""
sql = "select SNo,Branch from d_Suppliers where SNo='" & txtSNo & "'"
Set rss = oSaccoMaster.GetRecordset(sql)
If rss!Branch <> cbolocation Then
MsgBox "This supplier is not from this branch: " & cbolocation & "", vbInformation
Exit Sub
End If

'''*****************************end **************
Exit Sub
ErrorHandler:

MsgBox err.description
End Sub

Private Sub txtQnty_Validate(Cancel As Boolean)
txtQnty = Format(txtQnty, "####0.00")
End Sub
Public Sub load_deduc()
If txtSNo = "" Then Exit Sub
Lvwdeductions.ListItems.Clear
sql = ""
sql = "SELECT     Id, SNo, Date_Deduc, Description, Amount, auditid   FROM         d_supplier_deduc   WHERE     SNo = " & txtSNo & " and year(date_deduc)=" & Year(DTPMilkDate) & " and month(date_deduc)=" & month(DTPMilkDate) & ""
Set rs = oSaccoMaster.GetRecordset(sql)
If rs.EOF Then Exit Sub
If rs.RecordCount > 0 Then

With Lvwdeductions
    
       .ListItems.Clear
    
        .ColumnHeaders.Clear

  End With

    With Lvwdeductions
        
        .ColumnHeaders.Add , , "IDNo", 1500
        .ColumnHeaders.Add , , "SNo", 2000
        .ColumnHeaders.Add , , "Date Deduction Made", 3000
        .ColumnHeaders.Add , , "Description", 3000
        .ColumnHeaders.Add , , "Amount"
        .ColumnHeaders.Add , , "auditid"
      
    
        While Not rs.EOF
        
           If Not IsNull(rs.Fields("SNo")) Then
        
            Set li = .ListItems.Add(, , Trim(rs.Fields("id")))
            End If
            If Not IsNull(rs.Fields("sno")) Then
            li.ListSubItems.Add , , Trim(rs.Fields("sno"))
            End If
            If Not IsNull(rs.Fields("Date_Deduc")) Then
            li.ListSubItems.Add , , Trim(rs.Fields("Date_Deduc"))
            End If
            If Not IsNull(rs.Fields("description")) Then
            li.ListSubItems.Add , , Trim(rs.Fields("description"))
            End If
            If Not IsNull(rs.Fields("amount")) Then
             li.ListSubItems.Add , , Trim(rs.Fields("amount"))
            End If

             If Not IsNull(rs.Fields("auditid")) Then
             li.ListSubItems.Add , , Trim(rs.Fields("auditid"))
            End If
rs.MoveNext
        
        Wend
        
    End With
    
    rs.Close
    End If
    Set rs = Nothing
    
Lvwdeductions.View = lvwReport
End Sub

Private Sub txtSNo_Change()
txtSNo_Validate True
End Sub
Private Sub txtSNo_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorHandler
If (KeyAscii = 8) Or (KeyAscii = 48) Or (KeyAscii = 49) Or (KeyAscii = 50) Or (KeyAscii = 51) Or (KeyAscii = 52) Or (KeyAscii = 53) Or (KeyAscii = 54) Or (KeyAscii = 55) Or (KeyAscii = 56) Or (KeyAscii = 57) Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
MsgBox "Please enter a number "
End If

'''''*****************************check if supliers branch is correct ****
''sql = ""
''sql = "select SNo,Location from d_Suppliers where SNo='" & txtSNo & "'"
''Set rss = oSaccoMaster.GetRecordset(sql)
''If rss!Location <> cbolocation Then
''MsgBox "This supplier is not from this branch: " & cbolocation & "", vbInformation
''Exit Sub
''End If
''
'''''*****************************end **************

Exit Sub
ErrorHandler:
MsgBox err.description
'txttoday = "0"
Set rs = New ADODB.Recordset
Startdate = DateSerial(Year(DTPMilkDate), month(DTPMilkDate), Day(DTPMilkDate))
Enddate = DateSerial(Year(DTPMilkDate), month(DTPMilkDate), Day(DTPMilkDate) + 1)
sql = "d_sp_DailyTotal1 '" & Startdate & "','" & Enddate & "','" & txtSNo & "'"
'sql = "d_sp_DailyTotal1 '" & DTPMilkDate & "','" & txtSNo & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
    If Not IsNull(rs.Fields(0)) Then txttoday = rs.Fields(0)
    Else
    txttoday = "0"
    End If
    

End Sub

Private Sub txtSNo_Validate(Cancel As Boolean)

Set rs = New ADODB.Recordset

sql = "d_sp_SelectSuppliers '" & txtSNo & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
If Not IsNull(rs.Fields(2)) Then lblNames.Caption = rs.Fields(2)
load_deduc
Else
lblNames.Caption = ""
End If

'dailyfarmer
'txttoday = "0"
Set rs = New ADODB.Recordset
Startdate = DateSerial(Year(DTPMilkDate), month(DTPMilkDate), Day(DTPMilkDate))
Enddate = DateSerial(Year(DTPMilkDate), month(DTPMilkDate), Day(DTPMilkDate) + 1)
sql = "d_sp_DailyTotal1 '" & Startdate & "','" & Enddate & "','" & txtSNo & "'"
'sql = "d_sp_DailyTotal1 '" & DTPMilkDate & "','" & txtSNo & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
    If Not IsNull(rs.Fields(0)) Then txttoday = rs.Fields(0)
    Else
    txttoday = "0"
    End If

If txtSNo = "" Then Exit Sub
sql = ""
sql = "SELECT     TOP 1 Trans_Code  FROM         d_Transport  WHERE     (Sno = " & txtSNo & ") AND (Active = 1)  ORDER BY ID DESC"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
txtTCode = rs.Fields(0)


'//PUT THE NAME OF THE TRANSPORTER
Set rst = New ADODB.Recordset
sql = "d_sp_SelectTrans '" & txtTCode & "'"
Set rst = oSaccoMaster.GetRecordset(sql)
If Not rst.EOF Then
If Not IsNull(rst.Fields(0)) Then lblTName = rst.Fields(0)
Else
lblTName = ""
End If
Else
txtTCode = ""
End If
sql = ""
sql = "select* from d_Suppliers where sno='" & txtSNo & "' and approval=0"
Set rss = oSaccoMaster.GetRecordset(sql)
If Not rss.EOF Then
MsgBox "The supplier need approval first", vbInformation
Exit Sub
End If

'If rs.RecordCount = 0 Then
'lblNames.Caption = ""
'End If
End Sub

Private Sub txtTCode_Validate(Cancel As Boolean)
Set rs = New ADODB.Recordset
sql = "d_sp_SelectTrans '" & txtTCode & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
If Not IsNull(rs.Fields(0)) Then lblTName = rs.Fields(0)
Else
lblTName = ""
End If

End Sub


