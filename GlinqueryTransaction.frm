VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form GlinqueryTransaction 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "General Ledger Inquiry"
   ClientHeight    =   7980
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9855
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   9855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   2520
      TabIndex        =   20
      Top             =   7560
      Visible         =   0   'False
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdCorrect 
      Caption         =   "GlIssueCorrect"
      Height          =   375
      Left            =   240
      TabIndex        =   19
      Top             =   7560
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "GL TRANSACTIONS"
      Height          =   3495
      Left            =   3600
      TabIndex        =   14
      Top             =   1680
      Visible         =   0   'False
      Width           =   6015
      Begin VB.CommandButton Command1 
         Caption         =   "Close"
         Height          =   375
         Left            =   4680
         TabIndex        =   16
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtDocNo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   360
         Width           =   1815
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   2535
         Left            =   240
         TabIndex        =   17
         Top             =   720
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   4471
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "DRACCNO"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "CRACCNO"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "AMOUNT"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "MEMBERNO"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label30 
         Caption         =   "DocumentNo"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.CommandButton frmAccountStatement 
      Caption         =   "View Statement"
      Height          =   375
      Left            =   7080
      TabIndex        =   13
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton cmdFindacc 
      Caption         =   "<>"
      Height          =   315
      Left            =   4260
      TabIndex        =   9
      Top             =   90
      Width           =   465
   End
   Begin MSComCtl2.DTPicker dtpFromdate 
      Height          =   315
      Left            =   2160
      TabIndex        =   6
      Top             =   600
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   556
      _Version        =   393216
      Format          =   94961665
      CurrentDate     =   40179
   End
   Begin VB.TextBox txtAccno 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2160
      TabIndex        =   1
      Top             =   90
      Width           =   1995
   End
   Begin MSComCtl2.DTPicker dtpTodate 
      Height          =   315
      Left            =   5280
      TabIndex        =   8
      Top             =   600
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   556
      _Version        =   393216
      Format          =   94961665
      CurrentDate     =   40364
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5925
      Left            =   240
      TabIndex        =   12
      Top             =   1560
      Width           =   9420
      _ExtentX        =   16616
      _ExtentY        =   10451
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "TransDate"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "TransDescription"
         Object.Width           =   4057
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Debits"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Credits"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Balance"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Document No"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "TransNo"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Left            =   9120
      TabIndex        =   21
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "As by the Start of Range"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   240
      TabIndex        =   11
      Top             =   1080
      Width           =   2625
   End
   Begin VB.Label txtBalByRange 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2880
      TabIndex        =   10
      Top             =   1080
      Width           =   2235
   End
   Begin VB.Label Label8 
      Caption         =   "To Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4200
      TabIndex        =   7
      Top             =   600
      Width           =   915
   End
   Begin VB.Label Label7 
      Caption         =   "From Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   5
      Top             =   660
      Width           =   1155
   End
   Begin VB.Label lblCurrentbalance 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   7200
      TabIndex        =   4
      Top             =   1080
      Width           =   2235
   End
   Begin VB.Label Label2 
      Caption         =   "Book Balance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5520
      TabIndex        =   3
      Top             =   1080
      Width           =   1545
   End
   Begin VB.Label lblGlname 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   4845
      TabIndex        =   2
      Top             =   90
      Width           =   4590
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Account No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   285
      TabIndex        =   0
      Top             =   165
      Width           =   1200
   End
End
Attribute VB_Name = "GlinqueryTransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NormalBal As String
Dim RangeOpeningBal As Double

Private Sub cmdCorrect_Click()
 On Error GoTo Capture
    Set Rst = oSaccoMaster.GetRecordset("SET DATEFORMAT DMY select trAnsactionno,MEMBERNO from vwglissues where transdate>='" & dtpFromdate.value & "' and transdate<='" & dtpTodate.value & "' ORDER BY TRANSDATE")
    If Rst.EOF Then
        Exit Sub
    End If
    ProgressBar1.Max = 100
    ProgressBar1.Visible = True
    While Not Rst.EOF
        ProgressBar1.value = (Rst.AbsolutePosition / Rst.RecordCount) * 100
        oSaccoMaster.ExecuteThis ("UPDATE GLTRANSACTIONS SET SOURCE='" & Rst!memberno & "' WHERE TRANSACTIONNO='" & Rst!transactionNo & "'")
        If success = False Then
            GoTo Capture
        End If
    Rst.MoveNext
    Wend
    MsgBox "Update Done!"
 Exit Sub
Capture:
 MsgBox IIf(ErrorMessage = "", err.description, ErrorMessage)
End Sub

Private Sub cmdFindacc_Click()
frmSearchGLAccounts.Show vbModal, Me
    If Continue Then
        If SearchValue <> "" Then
            txtAccno = SearchValue
            SearchValue = ""
        End If
    End If
End Sub

Private Sub Command1_Click()
    Frame1.Visible = False
End Sub

Private Sub dtpFromdate_Change()
    txtAccno_Change
End Sub

Private Sub dtpTodate_Change()
    'txtAccNo_KeyPress 13
     txtAccno_Change
End Sub


Private Sub Form_Load()
    dtpFromdate = DateSerial(Year(Get_Server_Date), 1, 1)
    dtpTodate = Get_Server_Date
End Sub

Private Sub frmAccountStatement_Click()
        reportname = "AccountStatement.rpt"
        Show_Sales_Crystal_Report STRFORMULA, reportname, CompanyName
End Sub

Private Sub lblCurrentbalance_Change()
    lblCurrentbalance = Format(lblCurrentbalance, Cfmt)

End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With ListView1
        .Sorted = True
          .SortKey = ColumnHeader.SubItemIndex
          If .SortOrder = lvwAscending Then
              .SortOrder = lvwDescending
          Else
              .SortOrder = lvwAscending
          End If
    End With
End Sub

Private Sub ListView1_DblClick()
    If ListView1.ListItems.Count > 0 Then
        Set li = ListView1.SelectedItem
        mDocNo = li.SubItems(6)
        txtDocNo.Text = mDocNo
        Load_Ledgers mDocNo, txtAccno.Text
    End If
End Sub



Private Sub ListView3_DblClick()
    On Error GoTo Capture
    Dim rsDr As ADODB.Recordset, rsCr As ADODB.Recordset
    If ListView3.ListItems.Count = 0 Then
        Exit Sub
    End If
    If MsgBox("You want to switch the accounts?", vbQuestion + vbYesNo) = vbNo Then
        Exit Sub
    End If
    Dim newDrAcc As String
    Dim NewCrAcc As String
    dracc = ListView3.SelectedItem
    CRAcc = ListView3.SelectedItem.ListSubItems(1)
    
    newDrAcc = InputBox("Enter the New DRAccNo ", "New Debit Acc", "")
    NewCrAcc = InputBox("Enter the New CRAccNo ", "New Crebit Acc", "")
    
    
    
    If newDrAcc = "" And NewCrAcc = "" Then
        MsgBox "You chose not to make any change!"
        Exit Sub
    ElseIf newDrAcc = "" And NewCrAcc <> "" Then
        Set rs = oSaccoMaster.GetRecordset("SELECT * FROM GLSETUP WHERE ACCNO='" & NewCrAcc & "'")
        If rs.EOF Then
            MsgBox "The new Credit account is not a valid Gl Account", vbCritical
            Exit Sub
        End If
        sql = "Update Gltransactions set CRAccNo='" & NewCrAcc & "' where documentno='" & txtDocNo & "' and DRAccNo='" & dracc & "' and CRAccno='" & CRAcc & "'"
    ElseIf newDrAcc <> "" And NewCrAcc = "" Then
        Set Rst = oSaccoMaster.GetRecordset("SELECT * FROM GLSETUP WHERE ACCNO='" & newDrAcc & "'")
        If rs.EOF Then
            MsgBox "The new Debit account is not a valid Gl Account", vbCritical
            Exit Sub
        End If
        sql = "Update Gltransactions set DRAccNo='" & newDrAcc & "' where documentno='" & txtDocNo & "' and DRAccNo='" & dracc & "' and CRAccno='" & CRAcc & "'"
    Else
        If NewCrAcc <> "" Then
            Set rsCr = oSaccoMaster.GetRecordset("SELECT * FROM GLSETUP WHERE ACCNO='" & NewCrAcc & "'")
        End If
        If newDrAcc <> "" Then
            Set rs = oSaccoMaster.GetRecordset("SELECT * FROM GLSETUP WHERE ACCNO='" & newDrAcc & "'")
        End If
        If rsCr.EOF Then
            MsgBox "The new Credit account is not a valid Gl Account", vbCritical
            Exit Sub
        End If
        If rsDr.EOF Then
            MsgBox "The new Debit account is not a valid Gl Account", vbCritical
            Exit Sub
        End If
        sql = "Update Gltransactions set DRAccNo='" & newDrAcc & "',CRAccNo='" & NewCrAcc & "' where documentno='" & txtDocNo & "' and DRAccNo='" & dracc & "' and CRAccno='" & CRAcc & "'"
    End If
    oSaccoMaster.ExecuteThis (sql)
    If success = False Then
        MsgBox ErrorMessage
    Else
        MsgBox "Operation Successfull", vbInformation
    End If
    Exit Sub
Capture:
    MsgBox err.description
End Sub

Private Sub txtAccno_Change()

    If Trim(txtAccno) = "" Then
        Exit Sub
    End If
    If dtpFromdate > dtpTodate Then
        dtpFromdate = dtpTodate
    End If
    Get_GL_AccDetails txtAccno
    If GlAccName = "" Then
        ListView1.ListItems.clear
        lblGlname.Caption = ""
        lblCurrentbalance.Caption = 0
        txtBalByRange.Caption = 0
        Exit Sub
    Else
        lblGlname.Caption = GlAccName
        NormalBal = GlAccNBal
        'dtpFromdate.Value = OpeningBalDate
        lblCurrentbalance.Caption = CurrentBal
        RangeOpeningBal = getGlBalance(txtAccno, dtpFromdate, dtpFromdate)
        txtBalByRange.Caption = getGlBalance(txtAccno, dtpFromdate, dtpTodate)
        LoadTransactions
    End If
End Sub


Private Sub LoadTransactions()
    'On Error GoTo SysError
    Dim rsRecon As New Recordset, BankBal As Double, bCredits As Double, bDebits As Double, _
    RsDesc As New Recordset
    BankBal = RangeOpeningBal
    Set rsRecon = oSaccoMaster.GetRecordset("SET DATEFORMAT DMY EXEC GETgLtRANSACTIONS '" & txtAccno & "','" & dtpFromdate.value & "','" & Format(dtpTodate.value, "DD/MM/YYYY") & "'")
    ListView1.ListItems.clear
    With rsRecon
        While Not .EOF
            'DoEvents
            Set li = ListView1.ListItems.Add(, , !transdate)
jump:
            li.SubItems(1) = !tDescription
            li.SubItems(2) = Format(IIf(!transtype = "DR", !Amount, 0), Cfmt)
                'bDebits = bDebits + CDbl(li.SubItems(2))
                 bDebits = CDbl(li.SubItems(2))
            li.SubItems(3) = Format(IIf(!transtype = "CR", !Amount, 0), Cfmt)
                'bCredits = bCredits + CDbl(li.SubItems(3))
                 bCredits = CDbl(li.SubItems(3))
            If UCase(NormalBal) = UCase("Dr") Then
                BankBal = BankBal + bDebits - bCredits
            Else
                BankBal = BankBal + bCredits - bDebits
            End If
            li.SubItems(4) = Format(BankBal, Cfmt)
            li.SubItems(5) = IIf(IsNull(!chequeno), "", !chequeno)
            li.SubItems(6) = IIf(IsNull(!DocumentNo), "", !DocumentNo)
            'li.Checked = IIf(!recon = True, True, False)
            .MoveNext
        Wend
    End With
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub txtBalByRange_Change()
    txtBalByRange = Format(txtBalByRange, Cfmt)
End Sub
Private Sub Load_Ledgers(DocNo As String, ACCNO As String)
    On Error GoTo SysError
    Dim rsLedger As New Recordset
    Frame1.Visible = True
    ListView3.ListItems.clear
    Set rsLedger = oSaccoMaster.GetRecordset("Select * From gltransactions where " _
    & " documentno='" & _
    DocNo & "' and (drAccNo='" & ACCNO & "'or crAccNo='" & ACCNO & "') order by ID")
    With rsLedger
        If .State = adStateOpen Then
            While Not .EOF
                Set li = ListView3.ListItems.Add(, , IIf(IsNull(!DRaccno), "", !DRaccno))
                li.SubItems(1) = IIf(IsNull(!CRaccno), "", !CRaccno)
                li.SubItems(2) = IIf(IsNull(!Amount), 0, !Amount)
                li.SubItems(3) = IIf(IsNull(!Source), "", !Source)
                .MoveNext
            Wend
        End If
    End With
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

