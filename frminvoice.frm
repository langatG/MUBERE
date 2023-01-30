VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frminvoice 
   Caption         =   "Create Invoice"
   ClientHeight    =   4965
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7350
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4965
   ScaleWidth      =   7350
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtdocNo 
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
      Left            =   5280
      TabIndex        =   29
      Top             =   4530
      Width           =   1815
   End
   Begin VB.PictureBox Picture5 
      Height          =   255
      Left            =   2400
      Picture         =   "frminvoice.frx":0000
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   28
      Top             =   2280
      Width           =   255
   End
   Begin VB.TextBox txtNames 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   2760
      TabIndex        =   26
      Top             =   2280
      Width           =   4215
   End
   Begin VB.TextBox txtTCode 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   1080
      TabIndex        =   25
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdprint 
      Caption         =   "Print"
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
      Left            =   3720
      TabIndex        =   24
      Top             =   4440
      Width           =   1095
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
      Left            =   1200
      TabIndex        =   23
      Top             =   4440
      Width           =   855
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
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   1440
      Width           =   1815
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
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   1440
      Width           =   1815
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
      Left            =   2355
      Picture         =   "frminvoice.frx":02C2
      ScaleHeight     =   225
      ScaleWidth      =   240
      TabIndex        =   16
      Top             =   3240
      Width           =   300
   End
   Begin VB.TextBox txtdebtorAcc 
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
      Left            =   1080
      TabIndex        =   15
      Top             =   3240
      Width           =   1170
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
      Left            =   1440
      TabIndex        =   14
      Top             =   3840
      Width           =   5535
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
      Left            =   2400
      Picture         =   "frminvoice.frx":0584
      ScaleHeight     =   225
      ScaleWidth      =   240
      TabIndex        =   11
      Top             =   2760
      Width           =   300
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
      Left            =   1080
      TabIndex        =   10
      Top             =   2760
      Width           =   1170
   End
   Begin VB.TextBox txtrate 
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
      Left            =   1680
      TabIndex        =   8
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox txtkilos 
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
      Left            =   5280
      TabIndex        =   4
      Top             =   840
      Width           =   1815
   End
   Begin MSComCtl2.DTPicker DTPtransdate 
      Height          =   375
      Left            =   4320
      TabIndex        =   1
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
      Left            =   2400
      TabIndex        =   0
      Top             =   4440
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1560
      TabIndex        =   5
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
   Begin VB.Label Label10 
      Caption         =   "Debtors"
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
      TabIndex        =   27
      Top             =   2280
      Width           =   855
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
      Left            =   3720
      TabIndex        =   22
      Top             =   1440
      Width           =   975
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
      Left            =   120
      TabIndex        =   20
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label8 
      Caption         =   "Debtors Acc"
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
      TabIndex        =   18
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label lbldebtorname 
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
      Left            =   2760
      TabIndex        =   17
      Top             =   3240
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
      Left            =   120
      TabIndex        =   13
      Top             =   3840
      Width           =   975
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
      Left            =   2760
      TabIndex        =   12
      Top             =   2760
      Width           =   4170
   End
   Begin VB.Label Label1 
      Caption         =   "Rate"
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
      TabIndex        =   9
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Total Kilos"
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
      Left            =   3840
      TabIndex        =   7
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "From"
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
      TabIndex        =   6
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "To"
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
      Left            =   3600
      TabIndex        =   3
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Income Item"
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
      TabIndex        =   2
      Top             =   2760
      Width           =   1095
   End
End
Attribute VB_Name = "frminvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdNew_Click()
    DTPicker1 = Format(Get_Server_Date, "dd/mm/yyyy")
    dtptransdate = DateSerial(year(DTPicker1), month(DTPicker1) + 1, 1 - 1)
    DTPicker1 = DateSerial(year(DTPicker1), month(DTPicker1), 1)
 
    txtNarration = ""
    txtcontra = ""
    txtkilos = 0
    txtAmount = 0
    txtkilos = 0
    txtrate = 0
    txtdebtorAcc = ""
    lblcontra = ""
    lbldebtorname = ""
    
    Generate_InvoiceNo

End Sub

Private Sub cmdPrint_Click()
    STRFORMULA = "{Invoice.InvoiceNo}=" & txtDocNo & ""
        reportname = "Invoice.rpt"
        Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub cmdsave_Click()
Dim amount As Double, DRaccno As String, Craccno As String, _
      TransSource As String, transDescription As String, CashBook As Long, doc_posted As Integer, chequeno As String
  If txtrate = "" Then
   MsgBox "Enter Rate ", vbInformation, Me.Caption
    txtrate.SetFocus
  Exit Sub
 End If
  If txtcontra = "" Then
   MsgBox "Enter Income Item ", vbInformation, Me.Caption
    txtcontra.SetFocus
  Exit Sub
 End If
  If txtkilos = "" Then
   MsgBox "Enter Total  kilos ", vbInformation, Me.Caption
    txtkilos.SetFocus
  Exit Sub
 End If
  If txtdebtorAcc = "" Then
   MsgBox "Enter Debtor Accno ", vbInformation, Me.Caption
    txtdebtorAcc.SetFocus
  Exit Sub
 End If
   If txtNarration = "" Then
   MsgBox "Enter Narration ", vbInformation, Me.Caption
    txtNarration.SetFocus
  Exit Sub
 End If
 
 transdate = Format(dtptransdate, "dd/mm/yyyy")
 amount = CDbl(txtAmount)
 DRaccno = txtdebtorAcc
 Craccno = txtcontra
 DocumentNo = txtinvoiceNo
 TransSource = lblcontra
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
      
      
       sql = " set dateformat dmy  INSERT INTO invoice"
       sql = sql & " (InvoiceNo,Dcode,SupplierAcc, IncomeAcc,Amount,StartDate, EndDate, Transdescription, Rate, Kilos,Auditid) "
       sql = sql & "  VALUES     (" & txtinvoiceNo & ",'" & txtTCode & "','" & txtdebtorAcc & "','" & txtcontra & "'," & amount & " ,"
       sql = sql & "  '" & Format(DTPicker1, "dd/mm/yyyy") & "','" & transdate & "','" & transDescription & "'," & CDbl(txtrate) & "," & CDbl(txtkilos) & ",'" & User & "')"
       oSaccoMaster.ExecuteThis (sql)
       
       MsgBox "Invoice Created Successfuly", vbInformation, Me.Caption
        
        STRFORMULA = "{Invoice.InvoiceNo}=" & txtinvoiceNo & ""
        reportname = "Invoice.rpt"
        Show_Sales_Crystal_Report STRFORMULA, reportname, ""
       
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
            txtsupplierAcc = SearchValue
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

Private Sub Picture5_Click()
     frmSearchDebtors.Show vbModal
        txtTCode = sel
        txtTCode_Change
        Me.MousePointer = 0
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







Private Sub txtdocNo_LostFocus()
If Val(txtDocNo) = 0 Then
        MsgBox "Please enter a valid Invoice No", vbInformation, Me.Caption
        txtrate.SetFocus
        Beep
        Exit Sub
    End If
End Sub

Private Sub txtkilos_LostFocus()
If Val(txtkilos) = 0 Then
        MsgBox "Please enter a valid Amount", vbInformation, Me.Caption
        txtrate.SetFocus
        Beep
        Exit Sub
    End If
If txtrate = "" Then txtrate = 0
If txtkilos = "" Then txtkilos = 0
txtAmount = CDbl(txtrate * txtkilos)
End Sub



Private Sub txtrate_LostFocus()
If Val(txtrate) = 0 Then
        MsgBox "Please enter a valid Amount", vbInformation, Me.Caption
        txtrate.SetFocus
        Beep
        Exit Sub
    End If
If txtrate = "" Then txtrate = 0
If txtkilos = "" Then txtkilos = 0
txtAmount = CDbl(txtrate * txtkilos)
End Sub

Private Sub txtdebtorAcc_Change()
 Dim Account As Acc_Details
    Account = Get_Acc_Details(txtdebtorAcc, ErrorMessage)
    If Account.ACCNO <> "" Then
        lbldebtorname = Account.AccName
    Else
        If ErrorMessage <> "" Then
            MsgBox ErrorMessage, vbInformation, Me.Caption
            ErrorMessage = ""
        End If
        lbldebtorname = ""
    End If
End Sub
Sub Generate_InvoiceNo()
 sql = "select isnull(max(invoiceno),0) from Invoice"
  Set rst = oSaccoMaster.GetRecordset(sql)
   If Not rst.EOF Then
    txtinvoiceNo = rst.Fields(0) + 1
   End If
  
End Sub

Private Sub txtTCode_Change()
sql = "d_sp_Selectdebtors '" & txtTCode & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
If Not IsNull(rs.Fields(0)) Then txtNames = rs.Fields(0)
If Not IsNull(rs.Fields(15)) Then txtdebtorAcc = rs.Fields(15)
Else
txtNames = ""

End If
End Sub

Private Sub txtTCode_Click()
  txtTCode_Change
End Sub
