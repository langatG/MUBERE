VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmreversalofcashbookentries 
   Caption         =   "PRINTING"
   ClientHeight    =   7410
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13620
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   13620
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkvoucherno 
      Caption         =   "Use Dates Only"
      Height          =   255
      Left            =   2160
      TabIndex        =   32
      Top             =   6720
      Width           =   2175
   End
   Begin VB.CommandButton cmdprintgl 
      Caption         =   "Print GL"
      Height          =   495
      Left            =   5760
      TabIndex        =   31
      Top             =   6480
      Width           =   1215
   End
   Begin VB.TextBox txtchequeno 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9600
      TabIndex        =   30
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CommandButton cmdPost 
      Caption         =   "&Post"
      Height          =   405
      Left            =   8115
      TabIndex        =   25
      Top             =   6570
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   405
      Left            =   9465
      TabIndex        =   24
      Top             =   6570
      Width           =   1275
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   390
      Left            =   8160
      TabIndex        =   12
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
      ItemData        =   "frmreversalofcashbookentries.frx":0000
      Left            =   6405
      List            =   "frmreversalofcashbookentries.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   1605
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
      Left            =   6660
      TabIndex        =   10
      Top             =   435
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
      Left            =   480
      TabIndex        =   9
      Top             =   1620
      Width           =   1440
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   285
      Left            =   165
      TabIndex        =   8
      Top             =   1620
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
      Left            =   480
      TabIndex        =   7
      Top             =   1050
      Width           =   1440
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "..."
      Height          =   285
      Left            =   180
      TabIndex        =   6
      Top             =   1065
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
      Left            =   5700
      TabIndex        =   5
      Top             =   1050
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
      Left            =   3480
      TabIndex        =   4
      Top             =   435
      Width           =   3135
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
      Left            =   7515
      TabIndex        =   3
      Top             =   1050
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
      Left            =   2115
      TabIndex        =   2
      Top             =   1050
      Width           =   3225
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
      Left            =   2115
      TabIndex        =   1
      Top             =   1620
      Width           =   3225
   End
   Begin MSComctlLib.ListView lvwAccName 
      Height          =   1350
      Left            =   2115
      TabIndex        =   0
      Top             =   1350
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
      Left            =   165
      TabIndex        =   13
      Top             =   390
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
      CustomFormat    =   " dd/MM/yyyy"
      Format          =   132186115
      CurrentDate     =   39400
   End
   Begin MSComctlLib.ListView lvwTrans 
      Height          =   3390
      Left            =   -15
      TabIndex        =   26
      Top             =   3000
      Width           =   13515
      _ExtentX        =   23839
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
      NumItems        =   9
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
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Cheque No"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTPlastdate 
      Height          =   315
      Left            =   1755
      TabIndex        =   27
      Top             =   375
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
      CustomFormat    =   " dd/MM/yyyy"
      Format          =   132186115
      CurrentDate     =   39400
   End
   Begin VB.Label Label12 
      Caption         =   "ChequeNo"
      Height          =   255
      Left            =   9600
      TabIndex        =   29
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Last Date"
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
      Left            =   1755
      TabIndex        =   28
      Top             =   165
      Width           =   795
   End
   Begin VB.Label Label11 
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
      Left            =   5475
      TabIndex        =   23
      Top             =   1650
      Width           =   870
   End
   Begin VB.Label Label7 
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
      Left            =   6660
      TabIndex        =   22
      Top             =   180
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
      Left            =   525
      TabIndex        =   21
      Top             =   1395
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
      Left            =   2130
      TabIndex        =   20
      Top             =   1380
      Width           =   795
   End
   Begin VB.Label Label5 
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
      Left            =   525
      TabIndex        =   19
      Top             =   840
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
      Left            =   2130
      TabIndex        =   18
      Top             =   825
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
      Left            =   6720
      TabIndex        =   17
      Top             =   840
      Width           =   630
   End
   Begin VB.Label Label2 
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
      Left            =   180
      TabIndex        =   16
      Top             =   180
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
      Left            =   3495
      TabIndex        =   15
      Top             =   180
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
      Left            =   7530
      TabIndex        =   14
      Top             =   840
      Width           =   1140
   End
End
Attribute VB_Name = "frmreversalofcashbookentries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rscompany As New ADODB.Recordset


Private Sub cmdAdd_Click()
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
    If Trim$(txtnarration) = "" Then
        MsgBox "Please enter the Transaction Description", vbInformation, Me.Caption
        txtnarration.SetFocus
        SendKeys "{Home}+{End}"
        Exit Sub
    End If
    If Trim$(txtDocumentNo) = "" Then
        MsgBox "Please enter the Amount", vbInformation, Me.Caption
        txtDocumentNo.SetFocus
        SendKeys "{Home}+{End}"
        Exit Sub
    End If
    
    
    txtAmount = "0"
    txtDrAccNo = ""
    txtCrAccNo = ""
    txtDocumentNo = ""
    txtSource = ""
    txtnarration = ""
    lblDrAccName = ""
    txtCrAccName = ""
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdPost_Click()
    On Error GoTo SysError
    Dim Cubaccount As Cub_Acc_Details
    Dim Account As Acc_Details
    
    Dim DRaccno As String, Craccno As String, amount As Double, transdate As Date, _
    transDescription As String, TransSource As String, DocumentNo As String, CashBook As Long, doc_posted As Integer, chequeno As String
    If lvwTrans.ListItems.Count > 0 Then
        If MsgBox("Do you want to reverse this entry?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
            Exit Sub
        End If
    Else
        MsgBox "There are no transactions to be posted", vbInformation, Me.Caption
        Exit Sub
    End If
    
       If lvwTrans.ListItems.Count > 0 Then
        If MsgBox("Are you sure you want to reverse this entry?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
            Exit Sub
        End If

    End If
    
    Me.MousePointer = vbHourglass
    For I = 1 To lvwTrans.ListItems.Count
        Set li = lvwTrans.ListItems(I)
        transdate = li
        amount = CDbl(lvwTrans.ListItems(I).SubItems(1))
        DRaccno = lvwTrans.ListItems(I).SubItems(2)
        Craccno = lvwTrans.ListItems(I).SubItems(3)
        DocumentNo = lvwTrans.ListItems(I).SubItems(4)
        TransSource = lvwTrans.ListItems(I).SubItems(5)
        transDescription = lvwTrans.ListItems(I).SubItems(6)
        chequeno = txtChequeno
        CashBook = 1
        doc_posted = 1
      
        '// check if the doc number has been used
        'sp_doc_used
        
        
         
        Set rs = oSaccoMaster.GetRecordset("sp_doc_used '" & DocumentNo & "'")
        If Not rs.EOF Then
       GoTo anjela
        Else
           
anjela:
sql = ""
sql = "update gltransactions set documentno='" & "R" & DocumentNo & "',doc_posted=1,source='" & TransSource & "',chequeno='" & "R" & chequeno & "' WHERE documentno='" & txtDocumentNo & "'"
oSaccoMaster.ExecuteThis sql
        DocumentNo = "R" & DocumentNo
        chequeno = "R" & chequeno
        TransSource = "Reversal  " & TransSource
        If Not Save_GLTRANSACTION(transdate, amount, DRaccno, Craccno, DocumentNo, _
        TransSource, User, ErrorMessage, transDescription, CashBook, doc_posted, chequeno, transactionNo, "", "", 0) Then
            If ErrorMessage <> "" Then
                MsgBox ErrorMessage, vbInformation, Me.Caption
                ErrorMessage = ""
            End If
        End If
        
         End If
'        Cubaccount = Get_Cub_Acc_Details(DrAccNo, ErrorMessage)
'
'        '//SAVE  TO CUSTOMERBALANCEOLD
'
'        If Save_CustBalance(Cubaccount.AccNo, Cubaccount.AccNo, Cubaccount.payrollno, Cubaccount.AccName, Amount, _
'            Cubaccount.availablebalance, Cubaccount.AccNo, TransDescription, transdate, 0, DocumentNo, month(transdate), 0, 0, "DR", 0, DocumentNo, User, "GL Trans", CrAccNo, Get_Server_Date, Cubaccount.availablebalance, 1, "", 0, cn, ErrorMessage) = False Then
'
'            If ErrorMessage <> "" Then
'                MsgBox ErrorMessage, vbInformation, Me.Caption
'                ErrorMessage = ""
'            End If
'        End If
'        '//SAVE TO CUSTOMERBALANCE
'
'        If Save_CustBalance_OLD(Cubaccount.CustomerNo, Cubaccount.idno, Cubaccount.payrollno, Cubaccount.AccName, Amount, _
'            Cubaccount.availablebalance, Cubaccount.AccNo, TransDescription, transdate, 0, DocumentNo, month(transdate), 0, 0, "DR", 0, DocumentNo, User, "GL Trans", CrAccNo, Get_Server_Date, Cubaccount.availablebalance, 1, "", 0, cn, ErrorMessage) = False Then
'
'            If ErrorMessage <> "" Then
'                MsgBox ErrorMessage, vbInformation, Me.Caption
'                ErrorMessage = ""
'            End If
'        End If
'
'        ''//save  Credit Account
'        Cubaccount = Get_Cub_Acc_Details(CrAccNo, ErrorMessage)
'
'        If Save_CustBalance(Cubaccount.AccNo, Cubaccount.AccNo, Cubaccount.payrollno, Cubaccount.AccName, Amount, _
'            Cubaccount.availablebalance, Cubaccount.AccNo, TransDescription, transdate, 0, DocumentNo, month(transdate), 0, 0, "CR", 0, DocumentNo, User, "GL Trans", CrAccNo, Get_Server_Date, Cubaccount.availablebalance, 1, "", 0, cn, ErrorMessage) = False Then
'
'            If ErrorMessage <> "" Then
'                MsgBox ErrorMessage, vbInformation, Me.Caption
'                ErrorMessage = ""
'            End If
'        End If
'        '//SAVE TO CUSTOMERBALANCE
'
'        If Save_CustBalance_OLD(Cubaccount.CustomerNo, Cubaccount.idno, Cubaccount.payrollno, Cubaccount.AccName, Amount, _
'            Cubaccount.availablebalance, Cubaccount.AccNo, TransDescription, transdate, 0, DocumentNo, month(transdate), 0, 0, "CR", 0, DocumentNo, User, "GL Trans", CrAccNo, Get_Server_Date, Cubaccount.availablebalance, 1, "", 0, cn, ErrorMessage) = False Then
'
'            If ErrorMessage <> "" Then
'                MsgBox ErrorMessage, vbInformation, Me.Caption
'                ErrorMessage = ""
'            End If
'        End If
        
        
        
     
        
        
        
    Next I
    '//clear listview
    
    lvwTrans.ListItems.Clear
    
    Me.MousePointer = vbDefault
    MsgBox "Posting Successfull", vbInformation, Me.Caption
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
    Me.MousePointer = vbDefault
End Sub


Private Sub cmdprintgl_Click()
Dim str
str = Me.Caption

mysql = "delete  from GLTRANSACTIONS2"
oSaccoMaster.ExecuteThis (mysql)
 MousePointer = vbHourglass
 
 '// Get Opening Balances

'//Get Non-Member Transactions
If chkvoucherno = vbChecked Then
mysql = ""
mysql = "Get_Non_member_Transaction '" & Format(DTPTransdate, "dd/MM/yyyy") & "','" & Format(DTPlastdate, "dd/MM/yyyy") & "'"
oSaccoMaster.ExecuteThis (mysql)
Else
mysql = ""
mysql = "Get_Non_member_Transaction_Voucher '" & Format(DTPTransdate, "dd/MM/yyyy") & "','" & Format(DTPlastdate, "dd/MM/yyyy") & "','" & lvwTrans.SelectedItem.SubItems(4) & "'"
oSaccoMaster.ExecuteThis (mysql)
End If


reportname = "new transactions list.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, title
    
End Sub

Private Sub DTPlastdate_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
dtpTransDate_change
End Sub

Private Sub dtpTransDate_change()
lvwTrans.ListItems.Clear
Set rs = oSaccoMaster.GetRecordset("SET              dateformat dmy  SELECT     *   FROM         GLTRANSACTIONS where transdate>='" & DTPTransdate & "' and transdate<='" & DTPlastdate & "'")
While Not rs.EOF
Set li = lvwTrans.ListItems.Add(, , rs.Fields("transdate"))
    li.SubItems(1) = Format(CDbl(rs.Fields("Amount")))
    li.SubItems(2) = rs.Fields("DrAccNo")
    li.SubItems(3) = rs.Fields("CrAccNo")
    li.SubItems(4) = rs.Fields("documentno")
    li.SubItems(5) = rs.Fields("Source")
    li.SubItems(6) = rs.Fields("Transdescript")
    'li.SubItems(8) = rs.Fields("chequeno")
   
    rs.MoveNext
    Wend
End Sub

Private Sub dtpTransDate_DropDown()
dtpTransDate_change
End Sub

Private Sub dtpTransDate_LostFocus()
dtpTransDate_change
End Sub

Private Sub Form_Load()
    DTPTransdate = Format(Get_Server_Date, " dd/mm/yyyy")
    DTPlastdate = Format(Get_Server_Date, "dd/mm/yyyy")
    Set rscompany = Nothing

'Set rsCompany = oSaccoMaster.GetRecordSet("select * from Company order by CompanyCode asc")

'//load todays transactions
Set rs = oSaccoMaster.GetRecordset("SET              dateformat dmy  SELECT     *   FROM         GLTRANSACTIONS where transdate>='" & DTPTransdate & "' and transdate<='" & DTPlastdate & "'")
While Not rs.EOF
Set li = lvwTrans.ListItems.Add(, , rs.Fields("transdate"))
    li.SubItems(1) = Format(CDbl(rs.Fields("Amount")))
    li.SubItems(2) = rs.Fields("DrAccNo")
    li.SubItems(3) = rs.Fields("CrAccNo")
    li.SubItems(4) = rs.Fields("documentno")
    li.SubItems(5) = rs.Fields("Source")
    li.SubItems(6) = rs.Fields("Transdescript")
    If Not IsNull(rs.Fields("chequeno")) Then li.SubItems(8) = rs.Fields("chequeno")
   
    rs.MoveNext
    Wend
End Sub

Private Sub lblDrAccName_Change()
    On Error GoTo SysError
    Dim rsAccounts As New Recordset
    TSource = "DR"
    lvwAccName.ListItems.Clear
    If Trim$(lblDrAccName) <> "" Then
        If Not Editing Then
            Set rsAccounts = oSaccoMaster.GetRecordset("Exec Get_Acc_Names '%" & lblDrAccName & "%'")
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
                        li.SubItems(1) = IIf(IsNull(!ACCNO), "", !ACCNO)
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
    MsgBox err.description, vbInformation, Me.Caption
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
'lvwTrans.ListItems.Clear
Set rs = oSaccoMaster.GetRecordset("SET              dateformat dmy  SELECT     *   FROM         GLTRANSACTIONS where  documentno='" & lvwTrans.SelectedItem.SubItems(4) & "'")
lvwTrans.ListItems.Clear
While Not rs.EOF
Set li = lvwTrans.ListItems.Add(, , rs.Fields("transdate"))
    li.SubItems(1) = Format(CDbl(rs.Fields("Amount")))
    li.SubItems(2) = rs.Fields("CrAccNo")
    li.SubItems(3) = rs.Fields("DrAccNo")
    li.SubItems(4) = rs.Fields("documentno")
    li.SubItems(5) = rs.Fields("Source")
    li.SubItems(6) = rs.Fields("Transdescript")
    
    If Not IsNull(rs.Fields("amount")) Then txtAmount = rs.Fields("amount")
            If Not IsNull(rs.Fields("draccno")) Then txtCrAccNo = rs.Fields("draccno")
            If Not IsNull(rs.Fields("craccno")) Then txtDrAccNo = rs.Fields("craccno")
            If Not IsNull(rs.Fields("documentno")) Then txtDocumentNo = rs.Fields("documentno")
            If Not IsNull(rs.Fields("source")) Then txtSource = rs.Fields("source")
            If Not IsNull(rs.Fields("transdescript")) Then txtnarration = rs.Fields("transdescript")
            If Not IsNull(rs.Fields("chequeno")) Then txtChequeno = rs.Fields("chequeno")
   
    rs.MoveNext
    Wend
    
        If Not rs.EOF Then
           
            
        End If

End Sub

Private Sub txtCrAccName_Change()
    On Error GoTo SysError
    Dim rsAccounts As New Recordset
    TSource = "CR"
    lvwAccName.ListItems.Clear
    If Trim$(txtCrAccName) <> "" Then
        If Not Editing Then
            Set rsAccounts = oSaccoMaster.GetRecordset("Exec Get_Acc_Names '%" & txtCrAccName & "%'")
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
                        li.SubItems(1) = IIf(IsNull(!ACCNO), "", !ACCNO)
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
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub txtCrAccName_KeyPress(KeyAscii As Integer)
    KeyAscii = To_Upper_Case(KeyAscii)
End Sub

Private Sub txtCrAccNo_Change()
    On Error GoTo SysError
    Dim Account As Acc_Details
        
        Editing = True
    Account = Get_Acc_Details(txtCrAccNo, ErrorMessage)
    If Account.ACCNO <> "" Then
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
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub txtCrAccNo_KeyPress(KeyAscii As Integer)
    KeyAscii = To_Upper_Case(KeyAscii)
End Sub

Private Sub txtDrAccNo_Change()
    On Error GoTo SysError
    Dim Account As Acc_Details
    Editing = True
    Account = Get_Acc_Details(txtDrAccNo, ErrorMessage)
    If Account.ACCNO <> "" Then
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
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub txtDrAccNo_KeyPress(KeyAscii As Integer)
    KeyAscii = To_Upper_Case(KeyAscii)
End Sub

Private Sub txtSource_KeyPress(KeyAscii As Integer)
    KeyAscii = To_Upper_Case(KeyAscii)
End Sub




