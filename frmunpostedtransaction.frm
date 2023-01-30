VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmunpostedtransaction 
   Caption         =   "Unposted Transactions"
   ClientHeight    =   7095
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9345
   LinkTopic       =   "Form1"
   ScaleHeight     =   7095
   ScaleWidth      =   9345
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtdocumentno 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1920
      TabIndex        =   22
      Top             =   1680
      Width           =   2415
   End
   Begin VB.TextBox txttransdescription 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1920
      TabIndex        =   20
      Top             =   1320
      Width           =   6735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "..."
      Height          =   315
      Left            =   4200
      TabIndex        =   15
      Top             =   3120
      Width           =   345
   End
   Begin VB.TextBox txtdeductioncontrol 
      Height          =   375
      Left            =   1920
      TabIndex        =   13
      Top             =   3120
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      Caption         =   "..."
      Height          =   315
      Left            =   4200
      TabIndex        =   12
      Top             =   2640
      Width           =   345
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "..."
      Height          =   315
      Left            =   4200
      TabIndex        =   11
      Top             =   2160
      Width           =   345
   End
   Begin VB.CommandButton cmdload 
      Caption         =   "Load "
      Height          =   375
      Left            =   4440
      TabIndex        =   10
      Top             =   720
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DTPendofperiod 
      Height          =   375
      Left            =   1920
      TabIndex        =   9
      Top             =   720
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      Format          =   41746433
      CurrentDate     =   40234
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   375
      Left            =   8040
      TabIndex        =   8
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Post"
      Height          =   375
      Left            =   8040
      TabIndex        =   7
      Top             =   4080
      Width           =   975
   End
   Begin MSComctlLib.ListView Listunposted 
      Height          =   2895
      Left            =   360
      TabIndex        =   6
      Top             =   4080
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   5106
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Transport"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Agrovet "
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "AI"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Tamboche"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "FSA"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Housing"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Advance"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Others "
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Gross Pay"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Net"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox txtdraccount 
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   2640
      Width           =   2295
   End
   Begin VB.TextBox txtcraccount 
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   2160
      Width           =   2295
   End
   Begin VB.Label Label7 
      Caption         =   "Document No."
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "Transaction Details"
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label lbldeductioncontrol 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4560
      TabIndex        =   18
      Top             =   3120
      Width           =   4095
   End
   Begin VB.Label lbldraccount 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4560
      TabIndex        =   17
      Top             =   2640
      Width           =   4095
   End
   Begin VB.Label lblcraccount 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4560
      TabIndex        =   16
      Top             =   2160
      Width           =   4095
   End
   Begin VB.Label Label3 
      Caption         =   "Creditors Farmers/Deduction Control"
      Height          =   615
      Left            =   120
      TabIndex        =   14
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "Purchases Account"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Bank Account"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "End Of Period"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Unposted Transactions"
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
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "frmunpostedtransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdLoad_Click()
On Error GoTo ErrorHandler
Set rs = oSaccoMaster.GetRecordset("d_glpayrolltotals '" & DTPendofperiod & "'")

    Do While Not rs.EOF
        Set li = Listunposted.ListItems.Add(, , rs.Fields(0))
                        li.SubItems(1) = rs.Fields(1) & ""
                        li.SubItems(2) = rs.Fields(2) & ""
                        li.SubItems(3) = rs.Fields(3) & ""
                        li.SubItems(4) = rs.Fields(4) & ""
                        li.SubItems(5) = rs.Fields(5) & ""
                        li.SubItems(6) = rs.Fields(6) & ""
                        li.SubItems(7) = rs.Fields(7) & ""
                        li.SubItems(8) = rs.Fields(8) & ""
                        li.SubItems(9) = rs.Fields(9) & ""
                        rs.MoveNext
                        Loop
Exit Sub
ErrorHandler:
MsgBox err.description
End Sub

Private Sub cmdsearch_Click()
  frmSearchGLAccounts.Show vbModal, Me
    If Continue Then
        If SearchValue <> "" Then
            txtcraccount = SearchValue
            SearchValue = ""
        End If
    End If
End Sub

Private Sub Command1_Click()
    On Error GoTo SysError
    Dim Cubaccount As Cub_Acc_Details
    Dim Account As Acc_Details
    Dim chequeno As String
     If Check_Period_If_Closed(DTPendofperiod) = True Then
         Exit Sub
     End If
     If txtcraccount = "" Then
     MsgBox "Contra Account Missing, Please Enter", vbInformation
     Exit Sub
     End If
     If txtdeductioncontrol = "" Then
      MsgBox "Control Account Missing, Please Enter", vbInformation
     Exit Sub
     End If
    
    If txtdraccount = "" Then
     MsgBox "Contra Account Missing, Please Enter", vbInformation
     Exit Sub
     End If
    Dim Transports As Currency, agrovet As Currency, AI As Currency, Tamboche As Currency, FSA As Currency
    Dim HOUSING As Currency, Advance As Currency, Others As Currency, Gross_Pay As Currency
    Dim DRaccno As String, Craccno As String, amount As Double, transdate As Date, _
    transDescription As String, TransSource As String, DocumentNo As String, CashBook As Long, doc_posted As Integer
    If Listunposted.ListItems.Count > 0 Then
        If MsgBox("Do you want post the entry?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
            Exit Sub
        End If
    Else
        MsgBox "There are no transactions to be posted", vbInformation, Me.Caption
        Exit Sub
    End If
    '// check if the amount and month has been posted
    Dim rsmm As New ADODB.Recordset
    sql = "SELECT     mmonth, yyear, namount, posted  FROM         d_glposting where mmonth=" & month(DTPendofperiod) & " and yyear=" & year(DTPendofperiod) & " and posted=1"
    Set rsmm = oSaccoMaster.GetRecordset(sql)
    If Not rsmm.EOF Then
    MsgBox "The period you are trying to post has been posted"
    Else
   sql = "INSERT INTO d_glposting (mmonth, yyear, namount, posted) VALUES     (" & month(DTPendofperiod) & "," & year(DTPendofperiod) & ",0,0)"
   oSaccoMaster.ExecuteThis (sql)
    End If
    Me.MousePointer = vbHourglass
    For I = 1 To Listunposted.ListItems.Count
        Set li = Listunposted.ListItems(I)
        Transports = li
        agrovet = CDbl(Listunposted.ListItems(I).SubItems(1))
        AI = Listunposted.ListItems(I).SubItems(2)
        Tamboche = Listunposted.ListItems(I).SubItems(3)
        FSA = Listunposted.ListItems(I).SubItems(4)
        HOUSING = Listunposted.ListItems(I).SubItems(5)
        Advance = Listunposted.ListItems(I).SubItems(6)
        Others = Listunposted.ListItems(I).SubItems(7)
        Gross_Pay = Listunposted.ListItems(I).SubItems(8)
        transdate = DTPendofperiod
        CashBook = 1

  
         doc_posted = 1
        If Transports > 0 Then
        DocumentNo = txtdocumentno
        transDescription = txttransdescription
        TransSource = "Transport  for " & Format(DTPendofperiod, "mmm/yyyy") & " "
        DRaccno = txtdeductioncontrol
        Craccno = txtcraccount
        amount = Transports
        If Not Save_GLTRANSACTION(transdate, amount, DRaccno, Craccno, DocumentNo, _
        TransSource, User, ErrorMessage, transDescription, CashBook, doc_posted, chequeno, transactionNo, "", "", 0) Then
            If ErrorMessage <> "" Then
                MsgBox ErrorMessage, vbInformation, Me.Caption
                ErrorMessage = ""
            End If
        End If
        
         End If
       If agrovet > 0 Then
        DocumentNo = txtdocumentno
        transDescription = txttransdescription
        TransSource = "Agrovet  for " & Format(DTPendofperiod, "mmm/yyyy") & " "
        DRaccno = txtdeductioncontrol
        Craccno = txtcraccount
        amount = agrovet
        If Not Save_GLTRANSACTION(transdate, amount, DRaccno, Craccno, DocumentNo, _
        TransSource, User, ErrorMessage, transDescription, CashBook, doc_posted, chequeno, transactionNo, "", "", 0) Then
            If ErrorMessage <> "" Then
                MsgBox ErrorMessage, vbInformation, Me.Caption
                ErrorMessage = ""
            End If
        End If
        
         End If
        
        If AI > 0 Then
        DocumentNo = txtdocumentno
        transDescription = txttransdescription
        TransSource = "AI  for " & Format(DTPendofperiod, "mmm/yyyy") & " "
        DRaccno = txtdeductioncontrol
        Craccno = txtcraccount
        amount = AI
        If Not Save_GLTRANSACTION(transdate, amount, DRaccno, Craccno, DocumentNo, _
        TransSource, User, ErrorMessage, transDescription, CashBook, doc_posted, chequeno, transactionNo, "", "", 0) Then
            If ErrorMessage <> "" Then
                MsgBox ErrorMessage, vbInformation, Me.Caption
                ErrorMessage = ""
            End If
        End If
        
         End If
     
        If Tamboche > 0 Then
        DocumentNo = txtdocumentno
        transDescription = txttransdescription
        TransSource = "Tamboche  for " & Format(DTPendofperiod, "mmm/yyyy") & " "
        DRaccno = txtdeductioncontrol
        Craccno = txtcraccount
        amount = Tamboche
        If Not Save_GLTRANSACTION(transdate, amount, DRaccno, Craccno, DocumentNo, _
        TransSource, User, ErrorMessage, transDescription, CashBook, doc_posted, chequeno, transactionNo, "", "", 0) Then
            If ErrorMessage <> "" Then
                MsgBox ErrorMessage, vbInformation, Me.Caption
                ErrorMessage = ""
            End If
        End If
        
         End If
        
        If FSA > 0 Then
        DocumentNo = txtdocumentno
        transDescription = txttransdescription
        TransSource = "FSA  for " & Format(DTPendofperiod, "mmm/yyyy") & " "
        DRaccno = txtdeductioncontrol
        Craccno = txtcraccount
        amount = FSA
        If Not Save_GLTRANSACTION(transdate, amount, DRaccno, Craccno, DocumentNo, _
        TransSource, User, ErrorMessage, transDescription, CashBook, doc_posted, chequeno, transactionNo, "", "", 0) Then
            If ErrorMessage <> "" Then
                MsgBox ErrorMessage, vbInformation, Me.Caption
                ErrorMessage = ""
            End If
        End If
        
         End If
        
        If HOUSING > 0 Then
        DocumentNo = txtdocumentno
        transDescription = txttransdescription
        TransSource = "HOUSING  for " & Format(DTPendofperiod, "mmm/yyyy") & " "
        DRaccno = txtdeductioncontrol
        Craccno = txtcraccount
        amount = HOUSING
        If Not Save_GLTRANSACTION(transdate, amount, DRaccno, Craccno, DocumentNo, _
        TransSource, User, ErrorMessage, transDescription, CashBook, doc_posted, chequeno, transactionNo, "", "", 0) Then
            If ErrorMessage <> "" Then
                MsgBox ErrorMessage, vbInformation, Me.Caption
                ErrorMessage = ""
            End If
        End If
        
         End If
         
        If Advance > 0 Then
        DocumentNo = txtdocumentno
        transDescription = txttransdescription
        TransSource = "Advance  for " & Format(DTPendofperiod, "mmm/yyyy") & " "
        DRaccno = txtdeductioncontrol
        Craccno = txtcraccount
        amount = Advance
        If Not Save_GLTRANSACTION(transdate, amount, DRaccno, Craccno, DocumentNo, _
        TransSource, User, ErrorMessage, transDescription, CashBook, doc_posted, chequeno, transactionNo, "", "", 0) Then
            If ErrorMessage <> "" Then
                MsgBox ErrorMessage, vbInformation, Me.Caption
                ErrorMessage = ""
            End If
        End If
        
         End If
         
        If Others > 0 Then
        DocumentNo = txtdocumentno
        transDescription = txttransdescription
        TransSource = "Others  for " & Format(DTPendofperiod, "mmm/yyyy") & " "
        DRaccno = txtdeductioncontrol
        Craccno = txtcraccount
        amount = Others
        If Not Save_GLTRANSACTION(transdate, amount, DRaccno, Craccno, DocumentNo, _
        TransSource, User, ErrorMessage, transDescription, CashBook, doc_posted, chequeno, transactionNo, "", "", 0) Then
            If ErrorMessage <> "" Then
                MsgBox ErrorMessage, vbInformation, Me.Caption
                ErrorMessage = ""
            End If
        End If
        
         End If
         
        If Gross_Pay > 0 Then
        DocumentNo = txtdocumentno
        transDescription = txttransdescription
        TransSource = "Total Purchases  for " & Format(DTPendofperiod, "mmm/yyyy") & " "
        DRaccno = txtcraccount ' BANK ACCOUNT
        Craccno = txtdraccount ' PURCHASES
        amount = Gross_Pay
        If Not Save_GLTRANSACTION(transdate, amount, DRaccno, Craccno, DocumentNo, _
        TransSource, User, ErrorMessage, transDescription, CashBook, doc_posted, chequeno, transactionNo, "", "", 0) Then
            If ErrorMessage <> "" Then
                MsgBox ErrorMessage, vbInformation, Me.Caption
                ErrorMessage = ""
            End If
        End If
        
         End If
         
         
    Next I
    '//clear listview
'// update the period
sql = "update  d_glposting set posted =1 where mmonth=" & month(DTPendofperiod) & " and yyear=" & year(DTPendofperiod) & " and posted=0"
    Listunposted.ListItems.Clear
    
    Me.MousePointer = vbDefault
    MsgBox "Posting Successfull", vbInformation, Me.Caption
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
    Me.MousePointer = vbDefault

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
 frmSearchGLAccounts.Show vbModal, Me
    If Continue Then
        If SearchValue <> "" Then
            txtdraccount = SearchValue
            SearchValue = ""
        End If
    End If
End Sub

Private Sub Command4_Click()
 frmSearchGLAccounts.Show vbModal, Me
    If Continue Then
        If SearchValue <> "" Then
            txtdeductioncontrol = SearchValue
            SearchValue = ""
        End If
    End If
End Sub

Private Sub txtcraccount_Change()
  On Error GoTo SysError
    Dim rsAcc As New Recordset
    'lvwCrAccounts.ListItems.clear
    If Trim$(txtcraccount) <> "" Then
        Set rsAcc = oSaccoMaster.GetRecordset("Select GLAccName From GLSETUP " _
        & "where accno='" & txtcraccount & "'")
        With rsAcc
            If Not .EOF Then
                lblcraccount = .Fields(0)
            End If
        End With
    Else
        
    End If
    'lvwCrAccounts.Visible = False
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub txtdeductioncontrol_Change()
  On Error GoTo SysError
    Dim rsAcc As New Recordset
    'lvwCrAccounts.ListItems.clear
    If Trim$(txtcraccount) <> "" Then
        Set rsAcc = oSaccoMaster.GetRecordset("Select GLAccName From GLSETUP " _
        & "where accno='" & txtdeductioncontrol & "'")
        With rsAcc
            If Not .EOF Then
                lbldeductioncontrol = .Fields(0)
            End If
        End With
    Else
        
    End If
    'lvwCrAccounts.Visible = False
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption

End Sub

Private Sub txtdraccount_Change()
  On Error GoTo SysError
    Dim rsAcc As New Recordset
    'lvwCrAccounts.ListItems.clear
    If Trim$(txtcraccount) <> "" Then
        Set rsAcc = oSaccoMaster.GetRecordset("Select GLAccName From GLSETUP " _
        & "where accno='" & txtdraccount & "'")
        With rsAcc
            If Not .EOF Then
                lbldraccount = .Fields(0)
            End If
        End With
    Else
        
    End If
    'lvwCrAccounts.Visible = False
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption

End Sub
