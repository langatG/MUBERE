VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCompundGLTrans 
   Caption         =   "Detailed GL Transactions"
   ClientHeight    =   5565
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8790
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCompundGLTrans.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   8790
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox lblAccName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1965
      TabIndex        =   21
      Top             =   1200
      Width           =   3225
   End
   Begin MSComctlLib.ListView lvwAccName 
      Height          =   1740
      Left            =   1965
      TabIndex        =   20
      Top             =   1485
      Visible         =   0   'False
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   3069
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "AccName"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "AccNo"
         Object.Width           =   18
      EndProperty
   End
   Begin VB.TextBox txtDocumentNo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6720
      TabIndex        =   18
      Top             =   555
      Width           =   1920
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7305
      TabIndex        =   17
      Top             =   5070
      Width           =   1275
   End
   Begin VB.CommandButton cmdPost 
      Caption         =   "&Post"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5385
      TabIndex        =   16
      Top             =   5085
      Width           =   1275
   End
   Begin VB.TextBox txtTotDebit 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4875
      TabIndex        =   14
      Top             =   4590
      Width           =   1665
   End
   Begin VB.TextBox txtTotCredit 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6870
      TabIndex        =   13
      Top             =   4590
      Width           =   1665
   End
   Begin VB.TextBox txtNaration 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2280
      TabIndex        =   11
      Top             =   555
      Width           =   4335
   End
   Begin MSComCtl2.DTPicker dtpTransDate 
      Height          =   315
      Left            =   465
      TabIndex        =   9
      Top             =   540
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   " dd-MM-yyyy"
      Format          =   163184643
      CurrentDate     =   39400
   End
   Begin VB.TextBox txtCrAmount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7035
      TabIndex        =   4
      Top             =   1200
      Width           =   1665
   End
   Begin VB.TextBox txtDrAmount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5280
      TabIndex        =   3
      Top             =   1200
      Width           =   1665
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "..."
      Height          =   330
      Left            =   135
      TabIndex        =   2
      Top             =   1185
      Width           =   330
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2910
      Left            =   60
      TabIndex        =   1
      Top             =   1665
      Width           =   8700
      _ExtentX        =   15346
      _ExtentY        =   5133
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "AccNo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "AccName"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Dr Amount"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Cr Amount"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "DocumentNo"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Narration"
         Object.Width           =   7056
      EndProperty
   End
   Begin VB.TextBox txtAccNo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   480
      TabIndex        =   0
      Top             =   1200
      Width           =   1440
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Document No"
      Height          =   210
      Left            =   6720
      TabIndex        =   19
      Top             =   315
      Width           =   1140
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Totals"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4200
      TabIndex        =   15
      Top             =   4620
      Width           =   555
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Naration"
      Height          =   210
      Left            =   2280
      TabIndex        =   12
      Top             =   300
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Transaction Date"
      Height          =   210
      Left            =   480
      TabIndex        =   10
      Top             =   300
      Width           =   1395
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Credit Amount"
      Height          =   210
      Left            =   7485
      TabIndex        =   8
      Top             =   960
      Width           =   1200
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Debit Amount"
      Height          =   210
      Left            =   5790
      TabIndex        =   7
      Top             =   960
      Width           =   1155
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "AccName"
      Height          =   210
      Left            =   1995
      TabIndex        =   6
      Top             =   945
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "AccNo"
      Height          =   210
      Left            =   480
      TabIndex        =   5
      Top             =   960
      Width           =   525
   End
End
Attribute VB_Name = "frmCompundGLTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdPost_Click()
    On Error GoTo SysError
    Dim I As Long, ACCNO As String, CrAmt As Double, DrAmt As Double, _
    AccName As String, CnMAZIWA As New Connection
    If Trim$(txtNaration) = "" Then
        MsgBox "Please enter the Naration for this Transaction.", vbInformation, Me.Caption
        txtNaration.SetFocus
        SendKeys "{Home}+{End}"
        Exit Sub
    End If
    If Trim(txtDocumentNo) = "" Then
        MsgBox "Please enter the Document No for this Transaction", vbInformation, Me.Caption
        txtDocumentNo.SetFocus
        SendKeys "{Home}+{End}"
        Exit Sub
    End If
    If ListView1.ListItems.Count > 0 Then
        If MsgBox("Do you want to Post the entries", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
            If CDbl(txtTotCredit) = CDbl(txtTotDebit) Then
                With CnMAZIWA
                    If .State = adStateClosed Then
                        .Open SelectedDsn, "atm"
                        CnMAZIWA.BeginTrans
                    End If
                End With
                For I = 1 To ListView1.ListItems.Count
                    ACCNO = ListView1.ListItems(I)
                    Set li = ListView1.ListItems(I)
                    AccName = ListView1.ListItems(I).SubItems(1)
                    DrAmt = IIf(CDbl(ListView1.ListItems(I).SubItems(2)) > 0, CDbl(ListView1.ListItems(I).SubItems(2)), 0)
                    CrAmt = IIf(CDbl(ListView1.ListItems(I).SubItems(3)) > 0, CDbl(ListView1.ListItems(I).SubItems(3)), 0)
                    If Not Save_CustBalance(ACCNO, ACCNO, ACCNO, AccName, _
                    IIf(DrAmt > CrAmt, DrAmt, CrAmt), 0, ACCNO, CStr(li.SubItems(5)), DTPTransdate, _
                    0, CStr(li.SubItems(4)), Month(DTPTransdate), 0, 0, IIf(DrAmt > CrAmt, "DR", "CR"), 0, _
                    ACCNO, User, "GL Trans", ACCNO, DTPTransdate, 0, 0, "", 0, 1, CnMAZIWA, _
                    ErrorMessage) Then
                        If ErrorMessage <> "" Then
                            CnMAZIWA.RollbackTrans
                            Set CnMAZIWA = Nothing
                            MsgBox ErrorMessage, vbInformation, Me.Caption
                            ErrorMessage = ""
                        End If
                    End If
                Next
                CnMAZIWA.CommitTrans
                MsgBox "Transaction Posted Successfully", vbInformation, Me.Caption
                ListView1.ListItems.Clear
                txtNaration = ""
                txtDocumentNo = ""
                txtTotCredit = "0.00"
                txtTotDebit = "0.00"
                Set CnMAZIWA = Nothing
            Else
                MsgBox "Total Credits must be equal to total Debits", vbInformation, Me.Caption
            End If
        End If
    Else
        MsgBox "There are no entries to be posted", vbInformation, Me.Caption
    End If
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub cmdsearch_Click()
    On Error GoTo SysError
    frmsearchacc.Show vbModal, Me
    If Continue Then
        If sel <> "" Then
            Txtaccno = sel
        End If
        sel = ""
    End If
    Continue = True
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If ListView1.ListItems.Count < 1 Then
                Exit Sub
        End If
            
        If MsgBox("are you sure you want to delete", vbYesNo) = vbYes Then
            
            
            
            ListView1.ListItems.Remove (ListView1.SelectedItem.Index)  '// removes the selected item
            
            
            'ListView1.Refresh
        End If
    End If
End Sub


Private Sub Form_Load()
    On Error GoTo SysError
    DTPTransdate = Format(Get_Server_Date, " dd-MM-yyyy")
    
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub lblAccName_Change()
    On Error GoTo SysError
    Dim rsAcc As New Recordset
    If Trim$(lblaccname) <> "" Then
        lvwAccName.ListItems.Clear
        Set rsAcc = oSaccoMaster.GetRecordset("Select GLAccName,AccNo From GLSETUP " _
        & "where GLAccName like '%" & lblaccname & "%'")
        With rsAcc
            If Not .EOF Then
                lvwAccName.Visible = True
                While Not .EOF
                    Set li = lvwAccName.ListItems.Add(, , !GlAccName)
                    li.SubItems(1) = !ACCNO
                    .MoveNext
                Wend
            Else
                lvwAccName.Visible = False
                lvwAccName.ListItems.Clear
            End If
        End With
    Else
        lvwAccName.Visible = False
        lvwAccName.ListItems.Clear
    End If
    If lvwAccName.ListItems.Count = 1 Then
        lblaccname = lvwAccName.SelectedItem
        Txtaccno = lvwAccName.SelectedItem.SubItems(1)
        lvwAccName.Visible = False
    End If
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub listview1_DblClick()
    On Error GoTo SysError
    Dim CrAmt As Double, DrAmt As Double, ACCNO As String
    If ListView1.ListItems.Count > 0 Then
        ACCNO = ListView1.SelectedItem
        DrAmt = CDbl(ListView1.SelectedItem.SubItems(2))
        CrAmt = CDbl(ListView1.SelectedItem.SubItems(3))
    End If
    Txtaccno = ACCNO
    txtCrAmount = Format(CrAmt, Cfmt)
    txtDrAmount = Format(DrAmt, Cfmt)
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub lvwAccName_Click()
    On Error GoTo SysError
    If lvwAccName.ListItems.Count > 0 Then
        lblaccname = lvwAccName.SelectedItem
        Txtaccno = lvwAccName.SelectedItem.SubItems(1)
        lvwAccName.Visible = False
        lvwAccName.ListItems.Clear
    End If
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub txtAccno_Change()
    On Error GoTo SysError
    If Trim$(Txtaccno) <> "" Then
        Dim rsAccs As New Recordset
        Set rsAccs = oSaccoMaster.GetRecordset("Select * From GLSETUP where AccNo" _
        & "='" & Txtaccno & "'")
        With rsAccs
            If .State = adStateOpen Then
                If Not .EOF Then
                    lblaccname = IIf(IsNull(!GlAccName), "", !GlAccName)
                Else
                    lblaccname = ""
                End If
            End If
        End With
    End If
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
    err.Raise err.number, err.Source, err.description
End Sub

Private Sub txtAccNo_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13
        If Trim$(Txtaccno) <> "" Then
            If lblaccname <> "" Then
                txtDrAmount = "0.00"
                txtDrAmount.SetFocus
                SendKeys "{Home}+{End}"
            End If
        End If
    End Select
End Sub

Private Sub Load_To_LiastView(ACCNO As String, Debit As Boolean, amount As Double)
    On Error GoTo SysError
    Dim DrAmt As Double, CrAmt As Double, I As Long
    Set li = ListView1.ListItems.Add(, , ACCNO)
    li.SubItems(1) = lblaccname
    Select Case Debit
        Case True
        li.SubItems(2) = Format(txtDrAmount, Cfmt)
        li.SubItems(3) = "0.00"
        Case False
        li.SubItems(2) = "0.00"
        li.SubItems(3) = Format(txtCrAmount, Cfmt)
    End Select
    li.SubItems(4) = IIf(Trim(txtDocumentNo) <> "", txtDocumentNo, "")
    li.SubItems(5) = IIf(Trim(txtNaration) <> "", txtNaration, "")
    For I = 1 To ListView1.ListItems.Count
        CrAmt = CrAmt + CDbl(ListView1.ListItems(I).SubItems(3))
        DrAmt = DrAmt + CDbl(ListView1.ListItems(I).SubItems(2))
    Next I
    txtTotCredit = Format(CrAmt, Cfmt)
    txtTotDebit = Format(DrAmt, Cfmt)
    CrAmt = 0
    DrAmt = 0
    Txtaccno = ""
    lblaccname = ""
    txtDrAmount = ""
    txtCrAmount = ""
    Txtaccno.SetFocus
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub txtCrAmount_KeyPress(KeyAscii As Integer)
    On Error GoTo SysError
    Select Case KeyAscii
        Case 48 To 57
        Case Is = 8
        Case Is = 46
        Case 13
        If Trim$(txtCrAmount) <> "" Then
            If CDbl(txtCrAmount) <> 0 Then
                Load_To_LiastView Txtaccno, False, CDbl(txtCrAmount)
            End If
        End If
        Case Else
        KeyAscii = 0
    End Select
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub txtDrAmount_KeyPress(KeyAscii As Integer)
    On Error GoTo SysError
    Select Case KeyAscii
        Case 48 To 57
        Case Is = 8
        Case Is = 46
        Case 13
        If Trim$(txtDrAmount) = "" Then
            txtCrAmount = "0.00"
            txtDrAmount = "0.00"
            txtCrAmount.SetFocus
            SendKeys "{Home}+{End}"
        Else
            If CDbl(txtDrAmount) = 0 Then
                txtCrAmount = "0.00"
                txtDrAmount = "0.00"
                txtCrAmount.SetFocus
                SendKeys "{Home}+{End}"
            Else
                Load_To_LiastView Txtaccno, True, CDbl(txtDrAmount)
            End If
        End If
        Case Else
        KeyAscii = 0
    End Select
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub
