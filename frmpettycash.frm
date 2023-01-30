VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPettyCash 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PETTY CASH"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12195
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   12195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame FraOtherpayment 
      Caption         =   "Transaction Details"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   120
      TabIndex        =   5
      Top             =   2565
      Width           =   11895
      Begin VB.CheckBox chkperiodicreceipts 
         Caption         =   "Print Period Vouchers"
         Height          =   255
         Left            =   9240
         TabIndex        =   37
         Top             =   1320
         Width           =   2055
      End
      Begin VB.TextBox txtnarations 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "##-##-####"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
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
         Left            =   2970
         TabIndex        =   35
         Top             =   960
         Width           =   8310
      End
      Begin VB.CommandButton cmdprint 
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7800
         TabIndex        =   34
         Top             =   1320
         Width           =   1140
      End
      Begin VB.TextBox txtNarration 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "##-##-####"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
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
         Left            =   3000
         TabIndex        =   29
         ToolTipText     =   "The person who is taking the money"
         Top             =   600
         Width           =   3510
      End
      Begin VB.TextBox TxtOtherPAcc 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "##-##-####"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
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
         Left            =   1320
         TabIndex        =   0
         Top             =   210
         Width           =   1305
      End
      Begin VB.PictureBox Picture21 
         Height          =   285
         Left            =   2685
         Picture         =   "frmpettycash.frx":0000
         ScaleHeight     =   225
         ScaleWidth      =   240
         TabIndex        =   7
         Top             =   225
         Width           =   300
      End
      Begin VB.TextBox TxtOtherPayment 
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
         Left            =   7545
         MaxLength       =   9
         TabIndex        =   1
         Text            =   "0"
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdsave 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3720
         TabIndex        =   2
         Top             =   1320
         Width           =   1110
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6210
         TabIndex        =   6
         Top             =   1305
         Width           =   1215
      End
      Begin VB.CommandButton cmdpost 
         Caption         =   "Post"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4980
         TabIndex        =   3
         Top             =   1320
         Width           =   1140
      End
      Begin MSComctlLib.ListView lvwTrans 
         Height          =   2535
         Left            =   120
         TabIndex        =   33
         Top             =   1680
         Width           =   11685
         _ExtentX        =   20611
         _ExtentY        =   4471
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
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
         NumItems        =   10
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
            Text            =   "DocPosted"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Cheque No."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "ID"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Description:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1950
         TabIndex        =   36
         Top             =   1005
         Width           =   990
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Payee:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2430
         TabIndex        =   30
         Top             =   645
         Width           =   540
      End
      Begin VB.Label lblOtherPaymentAcc 
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
         Left            =   3000
         TabIndex        =   9
         Top             =   210
         Width           =   3495
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   4
         Left            =   6915
         TabIndex        =   8
         Top             =   270
         Width           =   555
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2385
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   11895
      Begin VB.Frame Frame 
         ClipControls    =   0   'False
         Height          =   1695
         Left            =   225
         TabIndex        =   10
         Top             =   690
         Width           =   11415
         Begin VB.OptionButton Optcheque 
            Caption         =   "Cheque"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   1080
            TabIndex        =   28
            Top             =   300
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton Optcash 
            Caption         =   "Cash"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   26
            Top             =   300
            Width           =   855
         End
         Begin VB.Frame Frame7 
            Height          =   570
            Left            =   2130
            TabIndex        =   16
            Top             =   120
            Width           =   5895
            Begin VB.TextBox txtReceiptsno 
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
               TabIndex        =   18
               Top             =   195
               Width           =   1935
            End
            Begin VB.TextBox txtChequeno 
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
               Left            =   4065
               TabIndex        =   17
               Top             =   195
               Visible         =   0   'False
               Width           =   1695
            End
            Begin VB.Label lblCheque 
               AutoSize        =   -1  'True
               Caption         =   "Cheque No"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   3240
               TabIndex        =   20
               Top             =   240
               Visible         =   0   'False
               Width           =   795
            End
            Begin VB.Label lblVoucher 
               AutoSize        =   -1  'True
               Caption         =   "Voucher No"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   210
               TabIndex        =   19
               Top             =   225
               Width           =   870
            End
         End
         Begin VB.TextBox txtAmountDue 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   8100
            Locked          =   -1  'True
            MaxLength       =   9
            TabIndex        =   15
            Top             =   780
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.TextBox TxtDRAmount 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   2130
            Locked          =   -1  'True
            MaxLength       =   9
            TabIndex        =   14
            Text            =   "0"
            Top             =   1215
            Width           =   1410
         End
         Begin VB.PictureBox Picture4 
            Height          =   285
            Left            =   3525
            Picture         =   "frmpettycash.frx":02C2
            ScaleHeight     =   225
            ScaleWidth      =   240
            TabIndex        =   13
            Top             =   780
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
            Height          =   300
            Left            =   2130
            TabIndex        =   12
            Top             =   780
            Width           =   1410
         End
         Begin VB.CommandButton cmdnew 
            Caption         =   "New"
            Height          =   330
            Left            =   8085
            TabIndex        =   11
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label LblStatus 
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
            ForeColor       =   &H000000FF&
            Height          =   300
            Left            =   4905
            TabIndex        =   25
            Top             =   1200
            Width           =   1320
         End
         Begin VB.Label Labal 
            Caption         =   "Avaliable Amount"
            Height          =   255
            Left            =   750
            TabIndex        =   24
            Top             =   1260
            Width           =   1335
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
            Left            =   3825
            TabIndex        =   23
            Top             =   780
            Width           =   4170
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Petty Cash Account"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   570
            TabIndex        =   22
            Top             =   825
            Width           =   1440
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            Caption         =   "Account Status"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   3735
            TabIndex        =   21
            Top             =   1245
            Width           =   1125
         End
      End
      Begin MSComCtl2.DTPicker dtptransdate 
         Height          =   315
         Left            =   7755
         TabIndex        =   31
         Top             =   390
         Width           =   1365
         _ExtentX        =   2408
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
         Format          =   115867651
         CurrentDate     =   39954
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Transaction Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   6465
         TabIndex        =   32
         Top             =   450
         Width           =   1230
      End
      Begin VB.Label Label1 
         Caption         =   "THE PETTY CASH ACCOUNT"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   495
         Left            =   240
         TabIndex        =   27
         Top             =   240
         Width           =   5895
      End
   End
End
Attribute VB_Name = "frmPettyCash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DRaccno As String
Dim Craccno As String
Dim IntAccNo As String
Dim glmemno As String
Dim glnamE1 As String
Dim amount As Currency
Dim ReceiptNo As String
Dim mysql As String
Dim loanno1 As String
Dim loanno2 As String
Dim loanno3 As String
Dim loanno4 As String
Dim loanno5 As String
Dim loanno6 As String
Dim loanno7 As String

Private Sub cmdNew_Click()
Dim rsr As New ADODB.Recordset
Dim rsg As New ADODB.Recordset
Dim I As Object
Dim Mylength As Integer
'//if this record is new then look for receipts no

''//clear all textboxes





mysql = ""
mysql = "select GenerateReceiptno from param"

Set rsg = oSaccoMaster.GetRecordset(mysql)
If Not rsg.EOF Then
    ''''check check
    If rsg!GenerateReceiptno = True Then
    
        mysql = ""
        mysql = "select * from Receiptno where receiptno like 'PC-%' order by Receipthnoid desc"
        
        Set rsr = oSaccoMaster.GetRecordset(mysql)
        
        If Not rsr.EOF Then
            Mylength = CInt(Mid(rsr!ReceiptNo, 5, 10))
            Mylength = Mylength + 1
            txtReceiptsno = Padding(Mylength)
            txtReceiptsno = "PC-" & txtReceiptsno
        Else
            Mylength = 1
            txtReceiptsno = "PC-" & Padding(Mylength)
            
        End If
Else
    ''//receiptno  will be keyed in
End If
End If

End Sub

Private Sub cmdPost_Click()
    On Error GoTo SysError
    If Check_Period_If_Closed(dtptransdate) = True Then
         Exit Sub
     End If
    Dim Cubaccount As Cub_Acc_Details
    Dim Account As Acc_Details
    Dim chequeno As String
    Dim DRaccno As String, Craccno As String, amount As Double, transdate As Date, _
    transDescription As String, TransSource As String, DocumentNo As String, CashBook As Long, doc_posted As Integer, IDENTI As Long
    If lvwTrans.ListItems.Count > 0 Then
        If MsgBox("Do you want post the entry?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
            Exit Sub
        End If
    Else
        MsgBox "There are no transactions to be posted", vbInformation, Me.Caption
        Exit Sub
    End If
    Me.MousePointer = vbHourglass
    For I = 1 To lvwTrans.ListItems.Count
    ' If lvwTrans.ListItems.Item(I).Checked = True Then
        Set li = lvwTrans.ListItems(I)
        transdate = li
        amount = CDbl(lvwTrans.ListItems(I).SubItems(1))
        DRaccno = lvwTrans.ListItems(I).SubItems(2)
        Craccno = lvwTrans.ListItems(I).SubItems(3)
        DocumentNo = lvwTrans.ListItems(I).SubItems(4)
        TransSource = lvwTrans.ListItems(I).SubItems(5)
        transDescription = lvwTrans.ListItems(I).SubItems(6)
        chequeno = lvwTrans.ListItems(I).SubItems(8)
        doc_posted = lvwTrans.ListItems(I).SubItems(7)
'IDENTI = lvwTrans.ListItems(i).SubItems(9)
        CashBook = 1
        If DocumentNo = "" Then DocumentNo = "CB"

                If chknonmemberpostings = vbChecked Then
                doc_posted = 1
                Else
                doc_posted = 0
                End If
        
        Set rs = oSaccoMaster.GetRecordset("sp_chequeno_used '" & Trim(chequeno) & "','" & Trim(TransSource) & "'")
        If Not rs.EOF Then

         End If

       
        
        If Not Save_GLTRANSACTION(transdate, amount, DRaccno, Craccno, DocumentNo, _
        TransSource, User, ErrorMessage, transDescription, CashBook, doc_posted, chequeno, transactionNo, "", "", 0) Then
            If ErrorMessage <> "" Then
                MsgBox ErrorMessage, vbInformation, Me.Caption
                ErrorMessage = ""
            End If
        End If
        
        
         sql = " INSERT INTO PettyCash"
         sql = sql & "             (transdate, AccName, Pvcno, Amount, Naration, auditid)"
         sql = sql & "  VALUES     ('" & transdate & "','" & TransSource & "','" & DocumentNo & "'," & amount & ",'" & transDescription & "','" & User & "')"
         oSaccoMaster.ExecuteThis (sql)
        
       ' End If
    Next I
    
    '//clear listview
    mysql = ""
    mysql = "set dateformat dmy Insert into Receiptno(Receiptno,Auditdate,auditid)values('" & txtReceiptsno & "','" & Format(Get_Server_Date, "dd/MM/yyyy") & "','" & User & "')"
    oSaccoMaster.ExecuteThis (mysql)
    lvwTrans.ListItems.Clear
    
    Me.MousePointer = vbDefault
    MsgBox "Posting Successfull", vbInformation, Me.Caption
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
    Me.MousePointer = vbDefault

End Sub

Private Sub cmdPrint_Click()
'//pettycashvoucher
If chkperiodicreceipts = vbChecked Then
 'STRFORMULA = "{PettyCash.Pvcno}='" & txtReceiptsno & "'"
    reportname = "pettycashvoucherperiodic.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, title
'//periodics
Else
    STRFORMULA = "{PettyCash.Pvcno}='" & txtReceiptsno & "'"
    reportname = "pettycashvoucher.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, title
'//periodics
'pettycashvoucherperiodic
End If
End Sub

Private Sub cmdsave_Click()
    On Error GoTo SysError
    If Trim$(CCur(TxtOtherPayment)) > CCur(TxtDRAmount) Then
        MsgBox "You do not have sufficient Amount in Petty Cash Account", vbInformation, Me.Caption
        Exit Sub
    End If
    
    If Trim(txtChequeno) = "" Then
       ' MsgBox "Please Enter The chequne No", vbInformation, Me.Caption
       ' Exit Sub
    End If
    '// PLEASE TOP UP YOUR IMPREST
    If TxtDRAmount < 0 Then
    MsgBox "Please top up your imprest amount", vbCritical
    Exit Sub
    End If
    
    If Val(TxtOtherPayment) = 0 Then
        MsgBox "Please enter a valid Amount", vbInformation, Me.Caption
        txtamount.SetFocus
        SendKeys "{Home}+{End}"
        Exit Sub
    End If
    If Trim$(txtcontra) = "" Then
        MsgBox "Please enter the Account to Debit.", vbInformation, Me.Caption
        txtDrAccNo.SetFocus
        SendKeys "{Home}+{End}"
        Exit Sub
    End If
    If txtnarations = "" Then
    MsgBox "Please enter the naration", vbCritical
    Exit Sub
    End If
    If Trim$(TxtOtherPAcc) = "" Then
        MsgBox "Please enter the Account to Credit.", vbInformation, Me.Caption
        txtCrAccNo.SetFocus
        SendKeys "{Home}+{End}"
        Exit Sub
    End If
    
    If Trim$(txtReceiptsno) = "" Then
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
   TxtDRAmount = CCur(TxtDRAmount) - CCur(TxtOtherPayment)
    Set li = lvwTrans.ListItems.Add(, , dtptransdate)
    li.SubItems(1) = Format(CDbl(TxtOtherPayment), "#,##0.00")
    li.SubItems(2) = TxtOtherPAcc
    li.SubItems(3) = txtcontra
    li.SubItems(4) = txtReceiptsno
    li.SubItems(5) = txtnarations & "-" & (lblOtherPaymentAcc) & "-" & TxtOtherPAcc
    li.SubItems(6) = txtNarration
    li.SubItems(7) = 1
    li.SubItems(8) = txtChequeno
    TxtOtherPayment = "0"
    TxtOtherPAcc = ""
    
    'txtReceiptsno = ""
    'txtNarration = ""
    txtnarations = ""
    Exit Sub
    'lblDrAc = ""
    lblOtherPaymentAcc = ""
    'txtchequeno = ""
SysError:
    MsgBox err.description, vbInformation, Me.Caption

End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim RsLoancode As New ADODB.Recordset
    Dim RsScheme As New ADODB.Recordset
    Dim rscompany As New ADODB.Recordset
    'optCash_Click
    dtptransdate.value = Format(Get_Server_Date, "dd/MM/yyyy")
    'get_availbalance
    'Headers
    'fraLoanRepayment.Visible = True
    'Frashares.Visible = True
    FraOtherpayment.Visible = True
    Optcash.value = True
    'optCash_Click
    'Load_Data
    txtChequeno.Visible = False
    lblVoucher.Visible = True
End Sub

Private Sub Optcash_Click()
If Optcash = True Then
        txtcontra = ""
        lblcontra = GetLedgerDesc(txtcontra)
        txtChequeno.Visible = False
        lblCheque.Visible = False
        lblVoucher.Visible = True
    End If
End Sub

Private Sub Picture21_Click()
frmSearchGLAccounts.Show vbModal, Me
    If Continue Then
        If SearchValue <> "" Then
            TxtOtherPAcc = SearchValue
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

Private Sub txtcontra_Change()
    On Error GoTo SysError
    Dim Account As Acc_Details
    Editing = True
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
    
     '//GET THE BALANCE AMOUNT
     
'    TxtDRAmount = getGlCurrentBalance(Account.ACCNO)
    
    ''*********** Commented By Koech
'
'    '//GET THE BALANCE AMOUNT
'mysql = "delete  from GLTRANSACTIONS2"
'oSaccoMaster.ExecuteThis (mysql)
' MousePointer = vbHourglass
'
' '// Get Opening Balances
'mysql = ""
'mysql = "Get_OpeningBalances '31/03/2014'"
'oSaccoMaster.ExecuteThis (mysql)
'
''//Get Non-Member Transactions
'Dim bal As Currency, CR As Currency, DR As Currency
'bal = 0
'mysql = ""
'mysql = "Get_Non_member_Transaction '31/03/2014','" & Format(Get_Server_Date, "dd/MM/yyyy") & "'"
'oSaccoMaster.ExecuteThis (mysql)
'sql = "SELECT     SUM(Amount) AS a, Transtype   FROM         GLTRANSACTIONS2   WHERE     (Accno ='" & txtcontra & "')   GROUP BY Transtype order by transtype desc"
'Set rs = oSaccoMaster.GetRecordset(sql)
'While Not rs.EOF
'If rs.Fields(1) = "DR" Then DR = rs.Fields(0)
'If rs.Fields(1) = "CR" Then CR = rs.Fields(0)
'bal = rs.Fields(0) - bal
'
'rs.MoveNext
'Wend
'If DR > CR Then
'TxtDRAmount = Abs(bal)
'Else
'TxtDRAmount = (bal * -1)
'End If

MousePointer = vbNormal
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption

End Sub

Private Sub TxtOtherPAcc_Change()
 On Error GoTo SysError
    Dim Account As Acc_Details
    Editing = True
    Account = Get_Acc_Details(TxtOtherPAcc, ErrorMessage)
    If Account.ACCNO <> "" Then
        lblOtherPaymentAcc = Account.AccName
    Else
        If ErrorMessage <> "" Then
            MsgBox ErrorMessage, vbInformation, Me.Caption
            ErrorMessage = ""
        End If
        lblOtherPaymentAcc = ""
    End If
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption

End Sub


