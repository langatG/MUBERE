VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form frmGLTran 
   Caption         =   "GL Tramsactions"
   ClientHeight    =   6075
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10935
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   10935
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   390
      Left            =   8100
      TabIndex        =   7
      Top             =   1545
      Width           =   1275
   End
   Begin VB.ComboBox cboTransType 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmGLTran.frx":0000
      Left            =   6345
      List            =   "frmGLTran.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1620
      Width           =   1680
   End
   Begin VB.TextBox txtNarration 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5010
      TabIndex        =   1
      Top             =   450
      Width           =   4365
   End
   Begin VB.TextBox txtCrAccNo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   420
      TabIndex        =   3
      Top             =   1635
      Width           =   1440
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   105
      TabIndex        =   22
      Top             =   1635
      Width           =   315
   End
   Begin VB.TextBox txtDrAccNo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   420
      TabIndex        =   2
      Top             =   1065
      Width           =   1440
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   13
      Top             =   1080
      Width           =   300
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5640
      TabIndex        =   4
      Top             =   1065
      Width           =   1665
   End
   Begin VB.TextBox txtSource 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1830
      TabIndex        =   0
      Top             =   450
      Width           =   3135
   End
   Begin VB.CommandButton cmdPost 
      Caption         =   "&Post"
      Height          =   405
      Left            =   8190
      TabIndex        =   11
      Top             =   5565
      Width           =   1275
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   405
      Left            =   9540
      TabIndex        =   10
      Top             =   5565
      Width           =   1275
   End
   Begin VB.TextBox txtDocumentNo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7455
      TabIndex        =   5
      Top             =   1065
      Width           =   1920
   End
   Begin VB.TextBox lblDrAccName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2055
      TabIndex        =   8
      Top             =   1065
      Width           =   3225
   End
   Begin MSComctlLib.ListView lvwAccName 
      Height          =   1350
      Left            =   2055
      TabIndex        =   9
      Top             =   1365
      Visible         =   0   'False
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   2381
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
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
   Begin MSComCtl2.DTPicker dtpTransDate 
      Height          =   315
      Left            =   120
      TabIndex        =   12
      Top             =   450
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   " dd-MM-yyyy"
      Format          =   73269251
      CurrentDate     =   39400
   End
   Begin MSComctlLib.ListView lvwTrans 
      Height          =   3390
      Left            =   75
      TabIndex        =   14
      Top             =   2025
      Width           =   10755
      _ExtentX        =   18971
      _ExtentY        =   5980
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "TransDate"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Amount"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Dr AccNo"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Cr AccNo"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "DocumentNo"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Source"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Narration"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "CashBook"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox txtCrAccName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2055
      TabIndex        =   21
      Top             =   1635
      Width           =   3225
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "TransType"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5415
      TabIndex        =   26
      Top             =   1665
      Width           =   870
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Naration"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5010
      TabIndex        =   25
      Top             =   195
      Width           =   705
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Cr AccNo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   465
      TabIndex        =   24
      Top             =   1410
      Width           =   765
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "AccName"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2070
      TabIndex        =   23
      Top             =   1395
      Width           =   795
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Dr AccNo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   465
      TabIndex        =   20
      Top             =   855
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "AccName"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2070
      TabIndex        =   19
      Top             =   840
      Width           =   795
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Amount"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6660
      TabIndex        =   18
      Top             =   855
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Transaction Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   17
      Top             =   195
      Width           =   1410
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Source"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1845
      TabIndex        =   16
      Top             =   195
      Width           =   585
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Document No"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   7470
      TabIndex        =   15
      Top             =   855
      Width           =   1140
   End
End
Attribute VB_Name = "frmGLTran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdadd_Click()
    On Error GoTo SysError
    If Trim$(txtAmount) = "" Then
        MsgBox "Please enter the Amount", vbInformation, Me.Caption
        txtAmount.SetFocus
        SendKeys "{Home}+{End}"
        Exit Sub
    End If
    If Val(txtAmount) = 0 Then
        MsgBox "Please enter a valid Amount", vbInformation, Me.Caption
        txtAmount.SetFocus
        SendKeys "{Home}+{End}"
        Exit Sub
    End If
    If Trim$(txtDrAccNo) = "" Then
        MsgBox "Please enter the Account to Debit.", vbInformation, Me.Caption
        txtDrAccNo.SetFocus
        SendKeys "{Home}+{End}"
        Exit Sub
    End If
    If Trim$(lblDrAccName) = "" Then
        MsgBox "Please verify the Debit Account.", vbInformation, Me.Caption
        lblDrAccName.SetFocus
        SendKeys "{Home}+{End}"
        Exit Sub
    End If
    If Trim$(txtCrAccNo) = "" Then
        MsgBox "Please enter the Account to Credit.", vbInformation, Me.Caption
        txtCrAccNo.SetFocus
        SendKeys "{Home}+{End}"
        Exit Sub
    End If
    If Trim$(txtCrAccName) = "" Then
        MsgBox "Please verify the Credit Account", vbInformation, Me.Caption
        txtCrAccName.SetFocus
        SendKeys "{Home}+{End}"
        Exit Sub
    End If
    If Trim$(txtSource) = "" Then
        MsgBox "Please enter the Transaction Source", vbInformation, Me.Caption
        txtSource.SetFocus
        SendKeys "{Home}+{End}"
        Exit Sub
    End If
    If Trim$(txtNarration) = "" Then
        MsgBox "Please enter the Transaction Description", vbInformation, Me.Caption
        txtNarration.SetFocus
        SendKeys "{Home}+{End}"
        Exit Sub
    End If
    If Trim$(txtDocumentNo) = "" Then
        MsgBox "Please enter the Amount", vbInformation, Me.Caption
        txtDocumentNo.SetFocus
        SendKeys "{Home}+{End}"
        Exit Sub
    End If
    Set li = lvwTrans.ListItems.Add(, , Dtptransdate)
    li.SubItems(1) = Format(CDbl(txtAmount), "#,##0.00")
    li.SubItems(2) = txtDrAccNo
    li.SubItems(3) = txtCrAccNo
    li.SubItems(4) = txtDocumentNo
    li.SubItems(5) = txtSource
    li.SubItems(6) = txtNarration
    If cboTransType.ListIndex = 0 Then
    li.SubItems(7) = 1
    Else
    li.SubItems(7) = 0
    End If
    
    txtAmount = "0"
    txtDrAccNo = ""
    txtCrAccNo = ""
    txtDocumentNo = ""
    txtSource = ""
    txtNarration = ""
    lblDrAccName = ""
    txtCrAccName = ""
    Exit Sub
SysError:
    MsgBox Err.description, vbInformation, Me.Caption
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdPost_Click()
    On Error GoTo SysError
    Dim Cubaccount As Cub_Acc_Details
    Dim Account As Acc_Details
    
    Dim DRaccno As String, CRaccno As String, amount As Double, transdate As Date, _
    TransDescription As String, TransSource As String, DocumentNo As String, CashBook As Long
    If lvwTrans.ListItems.Count > 0 Then
        If MsgBox("Do you want post the entry?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
            Exit Sub
        End If
    Else
        MsgBox "There are no transactions to be posted", vbInformation, Me.Caption
        Exit Sub
    End If
    Me.MousePointer = vbHourglass
    For i = 1 To lvwTrans.ListItems.Count
        Set li = lvwTrans.ListItems(i)
        transdate = li
        amount = CDbl(lvwTrans.ListItems(i).SubItems(1))
        DRaccno = lvwTrans.ListItems(i).SubItems(2)
        CRaccno = lvwTrans.ListItems(i).SubItems(3)
        DocumentNo = lvwTrans.ListItems(i).SubItems(4)
        TransSource = lvwTrans.ListItems(i).SubItems(5)
        TransDescription = lvwTrans.ListItems(i).SubItems(6)
        CashBook = lvwTrans.ListItems(i).SubItems(7)
'        If Not Save_GLTRANSACTION(transdate, Amount, DrAccNo, CrAccNo, DocumentNo, _
'        TransSource, User, ErrorMessage, TransDescription, CashBook) Then
'            If ErrorMessage <> "" Then
'                MsgBox ErrorMessage, vbInformation, Me.Caption
'                ErrorMessage = ""
'            End If
'        End If
        
        
        Cubaccount = Get_Cub_Acc_Details(DRaccno, ErrorMessage)
        
        '//SAVE  TO CUSTOMERBALANCEOLD
        
        If Save_CustBalance(Cubaccount.Accno, Cubaccount.Accno, Cubaccount.payrollno, Cubaccount.AccName, amount, _
            Cubaccount.availablebalance, Cubaccount.Accno, TransDescription, transdate, 0, DocumentNo, month(transdate), 0, 0, "DR", 0, DocumentNo, User, "GL Trans", CRaccno, Get_Server_Date, Cubaccount.availablebalance, 1, "", 0, "1", cn, ErrorMessage) = False Then
            
            If ErrorMessage <> "" Then
                MsgBox ErrorMessage, vbInformation, Me.Caption
                ErrorMessage = ""
            End If
        End If
        '//SAVE TO CUSTOMERBALANCE
        
        If Save_CustBalance_OLD(Cubaccount.CustomerNo, Cubaccount.idno, Cubaccount.payrollno, Cubaccount.AccName, amount, _
            Cubaccount.availablebalance, Cubaccount.Accno, TransDescription, transdate, 0, DocumentNo, month(transdate), 0, 0, "DR", 0, DocumentNo, User, "GL Trans", CRaccno, Get_Server_Date, Cubaccount.availablebalance, 1, "", 0, cn, ErrorMessage) = False Then
            
            If ErrorMessage <> "" Then
                MsgBox ErrorMessage, vbInformation, Me.Caption
                ErrorMessage = ""
            End If
        End If
        
        ''//save  Credit Account
        Cubaccount = Get_Cub_Acc_Details(CRaccno, ErrorMessage)
        
        If Save_CustBalance(Cubaccount.Accno, Cubaccount.Accno, Cubaccount.payrollno, Cubaccount.AccName, amount, _
            Cubaccount.availablebalance, Cubaccount.Accno, TransDescription, transdate, 0, DocumentNo, month(transdate), 0, 0, "CR", 0, DocumentNo, User, "GL Trans", CRaccno, Get_Server_Date, Cubaccount.availablebalance, 1, "", 0, DocumentNo, cn, ErrorMessage) = False Then
            
            If ErrorMessage <> "" Then
                MsgBox ErrorMessage, vbInformation, Me.Caption
                ErrorMessage = ""
            End If
        End If
        '//SAVE TO CUSTOMERBALANCE
        
        If Save_CustBalance_OLD(Cubaccount.CustomerNo, Cubaccount.idno, Cubaccount.payrollno, Cubaccount.AccName, amount, _
            Cubaccount.availablebalance, Cubaccount.Accno, TransDescription, transdate, 0, DocumentNo, month(transdate), 0, 0, "CR", 0, DocumentNo, User, "GL Trans", CRaccno, Get_Server_Date, Cubaccount.availablebalance, 1, "", 0, cn, ErrorMessage) = False Then
            
            If ErrorMessage <> "" Then
                MsgBox ErrorMessage, vbInformation, Me.Caption
                ErrorMessage = ""
            End If
        End If
        
        
        
     
        
        
        
    Next i
    '//clear listview
    
    lvwTrans.ListItems.Clear
    
    Me.MousePointer = vbDefault
    MsgBox "Posting Successfull", vbInformation, Me.Caption
    Exit Sub
SysError:
    MsgBox Err.description, vbInformation, Me.Caption
    Me.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Dtptransdate = Format(Get_Server_Date, " dd-MM-yyyy")
End Sub

Private Sub lblDrAccName_Change()
    On Error GoTo SysError
    Dim rsAccounts As New Recordset
    TSource = "DR"
    lvwAccName.ListItems.Clear
    If Trim$(lblDrAccName) <> "" Then
        If Not Editing Then
            Set rsAccounts = oSaccoMaster.GetRecordSet("Exec Get_Acc_Names '%" & lblDrAccName & "%'")
            With rsAccounts
                If .State = adStateOpen Then
                    If Not .EOF Then
                        lvwAccName.Visible = True
                        lvwAccName.Top = 1365
                    Else
                        lvwAccName.Visible = False
                    End If
                    While Not .EOF
                        Set li = lvwAccName.ListItems.Add(, , !GlAccName)
                        li.SubItems(1) = IIf(IsNull(!Accno), "", !Accno)
                        .MoveNext
                    Wend
                End If
            End With
        End If
    End If
    If lvwAccName.ListItems.Count = 1 Then
        lvwAccName_DblClick
    End If
    Editing = False
    Exit Sub
SysError:
    MsgBox Err.description, vbInformation, Me.Caption
End Sub

Private Sub lblDrAccName_KeyPress(KeyAscii As Integer)
    KeyAscii = To_Upper_Case(KeyAscii)
End Sub

Private Sub lvwAccName_DblClick()
    If lvwAccName.ListItems.Count > 0 Then
        Select Case TSource
            Case "DR"
            lblDrAccName = lvwAccName.SelectedItem
            txtDrAccNo = lvwAccName.SelectedItem.SubItems(1)
            Case "CR"
            txtCrAccName = lvwAccName.SelectedItem
            txtCrAccNo = lvwAccName.SelectedItem.SubItems(1)
        End Select
        lvwAccName.Visible = False
    End If
End Sub


Private Sub lvwTrans_DblClick()
    If lvwTrans.ListItems.Count > 0 Then
        If MsgBox("Are you sure you delete  this records " & lvwTrans.SelectedItem.Text & "? ", vbYesNo) = vbYes Then
        lvwTrans.ListItems.Remove (lvwTrans.SelectedItem.Index)  '// removes the selected item
        End If
    End If
End Sub

Private Sub txtCrAccName_Change()
    On Error GoTo SysError
    Dim rsAccounts As New Recordset
    TSource = "CR"
    lvwAccName.ListItems.Clear
    If Trim$(txtCrAccName) <> "" Then
        If Not Editing Then
            Set rsAccounts = oSaccoMaster.GetRecordSet("Exec Get_Acc_Names '%" & txtCrAccName & "%'")
            With rsAccounts
                If .State = adStateOpen Then
                    If Not .EOF Then
                        lvwAccName.Visible = True
                        lvwAccName.Top = 1935
                    Else
                        lvwAccName.Visible = False
                    End If
                    While Not .EOF
                        Set li = lvwAccName.ListItems.Add(, , !GlAccName)
                        li.SubItems(1) = IIf(IsNull(!Accno), "", !Accno)
                        .MoveNext
                    Wend
                End If
            End With
        End If
    End If
    If lvwAccName.ListItems.Count = 1 Then
        lvwAccName_DblClick
    End If
    Editing = False
    Exit Sub
SysError:
    MsgBox Err.description, vbInformation, Me.Caption
End Sub

Private Sub txtCrAccName_KeyPress(KeyAscii As Integer)
    KeyAscii = To_Upper_Case(KeyAscii)
End Sub

Private Sub txtCrAccNo_Change()
    On Error GoTo SysError
    Dim Account As Acc_Details
        
        Editing = True
    Account = Get_Acc_Details(txtCrAccNo, ErrorMessage)
    If Account.Accno <> "" Then
        txtCrAccName = Account.AccName
    Else
        If ErrorMessage <> "" Then
            MsgBox ErrorMessage, vbInformation, Me.Caption
            ErrorMessage = ""
        End If
        txtCrAccName = ""
    End If
    Exit Sub
SysError:
    MsgBox Err.description, vbInformation, Me.Caption
End Sub

Private Sub txtCrAccNo_KeyPress(KeyAscii As Integer)
    KeyAscii = To_Upper_Case(KeyAscii)
End Sub

Private Sub txtDrAccNo_Change()
    On Error GoTo SysError
    Dim Account As Acc_Details
    Editing = True
    Account = Get_Acc_Details(txtDrAccNo, ErrorMessage)
    If Account.Accno <> "" Then
        lblDrAccName = Account.AccName
    Else
        If ErrorMessage <> "" Then
            MsgBox ErrorMessage, vbInformation, Me.Caption
            ErrorMessage = ""
        End If
        lblDrAccName = ""
    End If
    Exit Sub
SysError:
    MsgBox Err.description, vbInformation, Me.Caption
End Sub

Private Sub txtDrAccNo_KeyPress(KeyAscii As Integer)
    KeyAscii = To_Upper_Case(KeyAscii)
End Sub

Private Sub txtSource_KeyPress(KeyAscii As Integer)
    KeyAscii = To_Upper_Case(KeyAscii)
End Sub
