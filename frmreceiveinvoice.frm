VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmreceiveinvoice 
   Caption         =   "Receive Invoice"
   ClientHeight    =   5730
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7350
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5730
   ScaleWidth      =   7350
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   9
      Top             =   5040
      Width           =   1095
   End
   Begin VB.TextBox txtcontra 
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
      Left            =   1200
      TabIndex        =   7
      Top             =   1920
      Width           =   1170
   End
   Begin VB.PictureBox Picture4 
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
      Left            =   2520
      Picture         =   "frmreceiveinvoice.frx":0000
      ScaleHeight     =   225
      ScaleWidth      =   240
      TabIndex        =   6
      Top             =   1920
      Width           =   300
   End
   Begin VB.TextBox txtnarration 
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
      Left            =   1320
      TabIndex        =   5
      Top             =   3840
      Width           =   5775
   End
   Begin VB.TextBox txtcreditorAcc 
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
      Left            =   1200
      TabIndex        =   4
      Top             =   3240
      Width           =   1170
   End
   Begin VB.PictureBox Picture1 
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
      Left            =   2475
      Picture         =   "frmreceiveinvoice.frx":02C2
      ScaleHeight     =   225
      ScaleWidth      =   240
      TabIndex        =   3
      Top             =   3240
      Width           =   300
   End
   Begin VB.TextBox txtamount 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox txtinvoiceNo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton cmdnew 
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   5040
      Width           =   855
   End
   Begin MSComCtl2.DTPicker DTPtransdate 
      Height          =   375
      Left            =   4920
      TabIndex        =   8
      Top             =   240
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   90701825
      CurrentDate     =   41927
   End
   Begin VB.Label Label1 
      Caption         =   " Invoice Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   19
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label11 
      BackColor       =   &H0080FF80&
      Caption         =   "CREDITOR ACCOUNT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   18
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Label Label10 
      BackColor       =   &H0080FF80&
      Caption         =   "ACCOUNT TO DEBIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   17
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Debit Acc"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label lblcontra 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   2880
      TabIndex        =   15
      Top             =   1920
      Width           =   4170
   End
   Begin VB.Label Label3 
      Caption         =   "Narration"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label lblcreditorname 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   2880
      TabIndex        =   13
      Top             =   3240
      Width           =   4170
   End
   Begin VB.Label Label8 
      Caption         =   "Creditor Acc"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   " Amount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label9 
      Caption         =   " InvoiceNo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmreceiveinvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdNew_Click()
 txtcontra = ""
 lblcontra = ""
 txtcreditorAcc = ""
 lblcreditorname = ""
 txtAmount = 0
 txtNarration = ""
 dtptransdate = Format(Get_Server_Date, "dd/mm/yyyy")
End Sub

Private Sub cmdsave_Click()
Dim amount As Double, DRaccno As String, Craccno As String, _
      TransSource As String, transDescription As String, CashBook As Long, doc_posted As Integer, chequeno As String
   If txtinvoiceNo = "" Then
   MsgBox "Enter Invoice No ", vbInformation, Me.Caption
    txtNarration.SetFocus
  Exit Sub
 End If
  
  If txtcontra = "" Then
   MsgBox "Enter Debit Gl Account Item ", vbInformation, Me.Caption
    txtcontra.SetFocus
  Exit Sub
 End If

  If txtcreditorAcc = "" Then
   MsgBox "Enter Creditor Account ", vbInformation, Me.Caption
    txtcreditorAcc.SetFocus
  Exit Sub
 End If
   If txtNarration = "" Then
   MsgBox "Enter Narration ", vbInformation, Me.Caption
    txtNarration.SetFocus
  Exit Sub
 End If
   
    transdate = Format(dtptransdate, "dd/mm/yyyy")
    If transdate > Format(Get_Server_Date, "dd/mm/yyyy") Then
     MsgBox "  Cant Transact on a future Date"
     dtptransdate.SetFocus
     Exit Sub
    End If
    
    amount = CDbl(txtAmount)
    DRaccno = txtcontra
    Craccno = txtcreditorAcc
    DocumentNo = txtinvoiceNo
    TransSource = lblcreditorname
    transDescription = txtNarration
    CashBook = 1
    doc_posted = 1
       If Not Save_GLTRANSACTION(transdate, amount, DRaccno, Craccno, DocumentNo, _
      TransSource, User, ErrorMessage, transDescription, CashBook, doc_posted, chequeno, TransNo) Then
          If ErrorMessage <> "" Then
              MsgBox ErrorMessage, vbInformation, Me.Caption
              ErrorMessage = ""
          End If
      End If
      
       sql = "set dateformat dmy  INSERT INTO InvoiceReceived"
       sql = sql & " (InvoiceNo,CreditorAccNo, DRAccNo,Amount,Transdate, Transdescription,Auditid) "
       sql = sql & "  VALUES     (" & txtinvoiceNo & ",'" & txtcreditorAcc & "','" & txtcontra & "'," & amount & " ,"
       sql = sql & "  '" & transdate & "','" & transDescription & "','" & User & "')"
       oSaccoMaster.ExecuteThis (sql)
       
       MsgBox "Invoice Received Successfuly", vbInformation, Me.Caption
        
        
       
       cmdNew_Click
       Exit Sub
ErrorHandler:
MsgBox err.description

End Sub



Private Sub Form_Load()
cmdNew_Click
End Sub

Private Sub Picture1_Click()
frmSearchGLAccounts.Show vbModal, Me
    If Continue Then
        If SearchValue <> "" Then
            txtcreditorAcc = SearchValue
            SearchValue = ""
        End If
    End If
End Sub

Private Sub Picture4_Click()
   frmSearchGLAccounts.Show vbModal, Me
    If Continue Then
        If SearchValue <> "" Then
            txtcontra = SearchValue
            SearchValue = ""
        End If
    End If
End Sub

Private Sub txtamount_LostFocus()
    If Val(txtAmount) = 0 Then
        MsgBox "Please enter a valid Amount", vbInformation, Me.Caption
        txtrate.SetFocus
        Beep
        Exit Sub
    End If
End Sub

Private Sub txtcontra_Change()
Dim Account As Acc_Details
    Account = Get_Acc_Details(txtcontra, ErrorMessage)
    If Account.ACCNO <> "" Then
        lblcontra = Account.AccName
    Else
        If ErrorMessage <> "" Then
            MsgBox ErrorMessage, vbInformation, Me.Caption
            ErrorMessage = ""
        End If
        lblcontra = ""
    End If
End Sub

Private Sub txtcreditorAcc_Change()
Dim Account As Acc_Details
    Account = Get_Acc_Details(txtcontra, ErrorMessage)
    If Account.ACCNO <> "" Then
        lblcreditorname = Account.AccName
    Else
        If ErrorMessage <> "" Then
            MsgBox ErrorMessage, vbInformation, Me.Caption
            ErrorMessage = ""
        End If
        lblcreditorname = ""
    End If
End Sub
