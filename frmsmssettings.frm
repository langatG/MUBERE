VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmsmssettings 
   Caption         =   "Setup sms"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5730
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   5730
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFC0FF&
      Height          =   495
      Left            =   1800
      TabIndex        =   6
      Top             =   1800
      Width           =   1815
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
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton smsre 
      Caption         =   "Sms Report"
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
      Left            =   1440
      TabIndex        =   4
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton cmdupdate 
      Caption         =   "Update"
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
      TabIndex        =   3
      Top             =   2760
      Width           =   1095
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   6240
      Width           =   5730
      _ExtentX        =   10107
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   2645
            MinWidth        =   2645
            Text            =   "USER : Birgen Gideon K."
            TextSave        =   "USER : Birgen Gideon K."
            Object.ToolTipText     =   "EASYMA User"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2470
            MinWidth        =   2470
            Picture         =   "frmsmssettings.frx":0000
            Text            =   "DATE : 07/12/2009"
            TextSave        =   "DATE : 07/12/2009"
            Object.ToolTipText     =   "Today's Date"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "9:32 AM"
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
   Begin MSComCtl2.DTPicker DTPMilkDate 
      Height          =   375
      Left            =   2280
      TabIndex        =   10
      Top             =   0
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      MouseIcon       =   "frmsmssettings.frx":0194
      CalendarBackColor=   8454016
      Format          =   160825345
      CurrentDate     =   40095
   End
   Begin MSComctlLib.ListView ListView3 
      Height          =   2535
      Left            =   240
      TabIndex        =   12
      Top             =   3600
      Width           =   5295
      _ExtentX        =   9340
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
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Deposited Amnt"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Send Sms Amnt"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Balance Amnt"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Send sms"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Today's collection"
      Height          =   2895
      Left            =   120
      TabIndex        =   13
      Top             =   3360
      Width           =   5535
   End
   Begin VB.Label Label7 
      Caption         =   "Send sms Today"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   14
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "Date"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label5 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   1800
      TabIndex        =   8
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   1800
      TabIndex        =   7
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Add Amount"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Balance sms Amount"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Amount spent Today"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1575
   End
End
Attribute VB_Name = "frmsmssettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdupdate_Click()
On Error GoTo ErrorHandler
Dim Price As Currency
Dim Startdate, CummulKgs, CummulKgs1, TRANSPORTER As String
Dim transdate As Date, anss As String
'check the user
 sql = "SELECT UserLoginIDs, UserGroup, SUPERUSER,status From UserAccounts where UserLoginIDs='" & User & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
    If rs!SuperUser <> 1 Then
    MsgBox "You are not allowed to update the sms", vbInformation
     
    Exit Sub
   End If
     End If

If Text1 = "" Then
   MsgBox "Please entered the amount."
    Text1.SetFocus
   Exit Sub
 End If


'Startdate = DateSerial(Year(DTPMilkDate), month(DTPMilkDate), 1)
'Enddate = DateSerial(Year(DTPMilkDate), month(DTPMilkDate) + 1, 1 - 1)

sql = ""
sql = "set dateformat dmy SELECT Dr From d_smssettings where Date ='" & DTPMilkDate & "' and Dr>0"
Set rs = oSaccoMaster.GetRecordset(sql)
If rs.EOF Then
sql = ""
sql = "set dateformat dmy INSERT INTO d_smssettings"
sql = sql & " (  Date, Dr, Cr, Balanace)"
sql = sql & " VALUES ('" & DTPMilkDate & "','" & Text1 & "','0','" & Text1 & "')"
oSaccoMaster.Execute (sql)
Else
sql = ""
sql = "SET dateformat DMY Update  d_smssettings SET Dr='" & Text1 & "',Balanace='" & Text1 & "' WHERE Date ='" & DTPMilkDate & "'and Dr>0"
oSaccoMaster.Execute (sql)
End If


NAMES3
loadMilk
Exit Sub
ErrorHandler:

MsgBox err.description
End Sub

Private Sub DTPMilkDate_Change()
NAMES3
loadMilk
End Sub
Private Sub DTPMilkDate_Click()
NAMES3
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
Text1 = "0"
loadMilk
End Sub
Public Sub NAMES3()

'Enddate = DateSerial(Year(DTPMilkDate), month(DTPMilkDate) + 1, 1 - 1)
sql = ""
sql = "set dateformat dmy SELECT Cr,Balanace From d_smssettings where Date ='" & DTPMilkDate & "' order by Date desc"
'sql = "set dateformat dmy SELECT d.DCode, d.DName, m.DispQnty,m.DispDate FROM  d_Debtors AS d INNER JOIN d_MilkControl AS m ON d.DCode = m.dcode WHERE     (DispDate = '" & DTPDispatchDate & "') and status=0"
Set rss = oSaccoMaster.GetRecordset(sql)
If Not rss.EOF Then
 Label4 = rss.Fields(0)
 Label5 = rss.Fields(1)
 Text2 = (rss.Fields(0) / 0.9)
Else
 sql = ""
 sql = "set dateformat dmy SELECT top 1 Cr,Balanace From d_smssettings where Date <'" & DTPMilkDate & "' order by Date desc"
 Set rss = oSaccoMaster.GetRecordset(sql)
 If Not rss.EOF Then
'  If rss.Fields(0) < 0 Then
   Label4 = 0
   Text2 = 0
'  Else
'   Label4 = rss.Fields(0)
'   Text2 = 0
'  End If
  If rss.Fields(1) < 0 Then
   Labe5 = 0
  Else
   Labe5 = rss.Fields(1)
  End If
 Else
 Label4 = 0
 Labe5 = 0
 Text2 = 0
 End If
End If
End Sub
Public Sub loadMilk()
    '/// to list view//////////
    Startdate = DateSerial(Year(DTPMilkDate), Month(DTPMilkDate), 1)
sql = ""
sql = "set dateformat dmy SELECT  Date, Dr, Cr, Balanace From d_smssettings where Date >='" & Startdate & "' and Date <='" & DTPMilkDate & "' order by Date desc"
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
                    If Not IsNull(rs.Fields(2)) Then li.SubItems(4) = rs.Fields(2) / 0.9 & ""
'                    If Not IsNull(rs.Fields(5)) Then li.SubItems(5) = rs.Fields(5) & ""
rs.MoveNext

Wend

'////// end of view
End Sub

Private Sub smsre_Click()
reportname = "SMS BALANCE REPORT.rpt"
Show_Sales_Crystal_Report "", reportname, ""
End Sub
