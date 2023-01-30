VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPeriods 
   Caption         =   "Period Setup"
   ClientHeight    =   5190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7170
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   7170
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   6975
      Begin MSComCtl2.DTPicker DTPdate 
         Height          =   375
         Left            =   4560
         TabIndex        =   9
         Top             =   840
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         Format          =   104660993
         CurrentDate     =   42392
      End
      Begin VB.Frame chkstatus2 
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   2655
         Begin VB.OptionButton optOpen 
            Caption         =   "Open"
            Height          =   375
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton optClosed 
            Caption         =   "Closed"
            Height          =   375
            Left            =   1200
            TabIndex        =   7
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Label lblperiod 
         Caption         =   "EndPeriod"
         Height          =   375
         Left            =   3240
         TabIndex        =   10
         Top             =   840
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   3840
      TabIndex        =   4
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   495
      Left            =   840
      TabIndex        =   3
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   5400
      TabIndex        =   2
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Top             =   4560
      Width           =   1215
   End
   Begin MSComctlLib.ListView lvwPeriod 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   4683
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Month"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Year"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Status"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "StartDate"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "EndDate"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmPeriods"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdadd_Click()
NewRecord = True
cmdSave.Enabled = True
cmdAdd.Enabled = False
cmdEdit.Enabled = True
End Sub

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdEdit_Click()
NewRecord = False
cmdEdit.Enabled = False
cmdSave.Enabled = True
cmdAdd.Enabled = True
End Sub

Private Sub cmdsave_Click()
Dim rsPeriod As New Recordset
Dim status As Integer, description As String

If optClosed = True Then
    status = 1
End If

If optOpen = True Then
    status = 0
End If

'If cboMonth.Text <> "" Then
'    description = MonthName(cboMonth)
'Else
'    MsgBox "Select the Period", vbInformation, Me.Caption
'    cboMonth.SetFocus
'    Exit Sub
'End If
'
'If Trim(txtyear) = "" Then
'    MsgBox "Enter the Year", vbInformation, Me.Caption
'    txtyear.SetFocus
'    Exit Sub
'End If

'If NewRecord = True Then 'Add new
'
'Set rsPeriod = oSaccoMaster.GetRecordset("select * from Periods where Period=" & cboMonth & " and PeriodYear=" & Trim(txtyear) & "")
'If Not rsPeriod.EOF Then
'    MsgBox "The Period is already Created.", vbInformation, Me.Caption
'    Exit Sub
'End If
'
'Set rsPeriod = Nothing
'Set rsPeriod = oSaccoMaster.GetRecordset("set dateformat dmy insert into PERIODS(Period,Description,PeriodYear,StartDate,EndDate,Status) values(" _
'& cboMonth & ",'" & description & "'," & Trim(txtyear) & ",'" & DateSerial(txtyear, cboMonth, 1) & "','" & DateSerial(txtyear, cboMonth + 1, 1 - 1) & "','" & status & "')")
'
'Else 'update
'
'Set rsPeriod = oSaccoMaster.GetRecordset("set dateformat dmy update PERIODS set Period=" _
'& cboMonth & ",Description='" & description & "',PeriodYear=" & txtyear & ",StartDate='" _
'& DateSerial(txtyear, cboMonth, 1) & "',EndDate='" & DateSerial(txtyear, cboMonth + 1, 1 - 1) & "',Status=" & status & " where Period=" & sql & " and PeriodYear=" & strValue & "")
'End If
Dim Vaa As Integer

If optClosed.value = True Then
Vaa = 1
Else
Vaa = 0
End If

oSaccoMaster.ExecuteThis ("d_sp_Periods2 '" & DTPdate & "'," & Vaa & ",'" & User & "'")
'ProgressBar1.value = 100

MsgBox "Completed Payroll"
'vbDefault

'End Sub

Call Form_Load
cmdSave.Enabled = False
cmdAdd.Enabled = True
cmdEdit.Enabled = True
End Sub

Private Sub Form_Load()
Dim m As Long, li As ListItem
Dim rsPeriods As New Recordset
'cboYear.Clear
'cboMonth.Clear
lvwPeriod.ListItems.Clear
cmdAdd.Enabled = True
cmdSave.Enabled = False
cmdEdit.Enabled = True
optOpen = True

'With cboMonth
'    For m = 1 To 12
'        .AddItem m
'    Next m
'End With
Set rsPeriods = oSaccoMaster.GetRecordset("select * from d_Periods order by ID")
If Not rsPeriods.EOF Then
    With rsPeriods
        While Not .EOF
            Set li = lvwPeriod.ListItems.Add(, , !endperiod)
            'li.ListSubItems.Add , , !Closed
            'li.ListSubItems.Add , , !periodyear
            li.ListSubItems.Add , , IIf(!Closed = True, "Closed", "Open")
            'li.ListSubItems.Add , , !Startdate
            li.ListSubItems.Add , , !auditid
        .MoveNext
        Wend
    End With
End If
End Sub

Private Sub lvwPeriod_DblClick()
Dim rsGetPeriod As New Recordset
strValue = lvwPeriod.SelectedItem.ListSubItems(2).Text
sql = lvwPeriod.SelectedItem.Text
 Set rsGetPeriod = oSaccoMaster.GetRecordset("select * from d_PERIODS where endPeriod=" _
 & lvwPeriod.SelectedItem.Text & "")
 If Not rsGetPeriod.EOF Then
    With rsGetPeriod
'        cboMonth = !Period
'        txtyear = !periodyear
        DTPdate = rsGetPeriod!endperiod
        'dtpEndDate.Value = !EndDate
        If !status = True Then
        optClosed.value = True
        Else
        optOpen.value = True
        End If
    End With
 End If
End Sub
