VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmConfirmPosting 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Confirm Batch Posting"
   ClientHeight    =   7455
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   11610
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   11610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdContinue 
      Caption         =   "Post"
      Height          =   495
      Left            =   5865
      TabIndex        =   9
      Top             =   6900
      Width           =   1215
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Close"
      Height          =   495
      Left            =   8745
      TabIndex        =   8
      Top             =   6885
      Width           =   1215
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "<< Back"
      Height          =   495
      Left            =   3105
      TabIndex        =   7
      Top             =   6885
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Transaction Summary"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   120
      TabIndex        =   2
      Top             =   5520
      Width           =   11415
      Begin MSComCtl2.DTPicker dtpCashBookDate 
         Height          =   375
         Left            =   5280
         TabIndex        =   12
         Top             =   270
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   78446595
         CurrentDate     =   39926
      End
      Begin MSComCtl2.DTPicker dtpTransDate 
         Height          =   375
         Left            =   1785
         TabIndex        =   13
         Top             =   300
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   661
         _Version        =   393216
         Format          =   78446595
         CurrentDate     =   39926
      End
      Begin VB.Label Label8 
         Caption         =   "Statement Date "
         Height          =   255
         Left            =   405
         TabIndex        =   15
         Top             =   360
         Width           =   1440
      End
      Begin VB.Label Label9 
         Caption         =   "CashBook Date"
         Height          =   255
         Left            =   4035
         TabIndex        =   14
         Top             =   330
         Width           =   1440
      End
      Begin VB.Label txtChequeAmount 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5310
         TabIndex        =   11
         Top             =   855
         Width           =   1815
      End
      Begin VB.Label txtChequeNo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1785
         TabIndex        =   10
         Top             =   855
         Width           =   1815
      End
      Begin VB.Label Label12 
         Caption         =   "Total"
         Height          =   195
         Left            =   7455
         TabIndex        =   6
         Top             =   285
         Width           =   720
      End
      Begin VB.Label Label6 
         Caption         =   "Cheque No"
         Height          =   195
         Left            =   690
         TabIndex        =   5
         Top             =   855
         Width           =   1080
      End
      Begin VB.Label Label7 
         Caption         =   "Cheque Amount"
         Height          =   195
         Left            =   4050
         TabIndex        =   4
         Top             =   855
         Width           =   1560
      End
      Begin VB.Label lblTotal 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   8760
         TabIndex        =   3
         Top             =   285
         Width           =   1815
      End
   End
   Begin MSComctlLib.ListView lvwConfirmPost 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   8916
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483646
      BackColor       =   -2147483624
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "MemberNo"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "FullName"
         Object.Width           =   5116
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Amount"
         Object.Width           =   4939
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Do you want to Continue Posting the following Transaction(s)?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   8295
   End
End
Attribute VB_Name = "frmConfirmPosting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
Unload Me
End Sub

Private Sub cmdBack_Click()
Continue = False
Unload Me
End Sub

Private Sub cmdContinue_Click()
On Error GoTo SysError
Dim FullName As String, rsPostShares As New Recordset, rsPostLoans As New Recordset
Dim Prncipal As Double, interest As Double, Shares As Double, totalamount As Double
Dim LoanOption As Society_Parameters

If MsgBox("Do you want to Post the Transaction(s)?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
    Exit Sub
End If

'// check if the period is closed
       If Check_Period_If_Closed(dtpTransDate) = True Then
            
            Exit Sub
          GoTo hell
        End If


With lvwConfirmPost
 If .ListItems.count >= 1 Then
        For i = 1 To .ListItems.count
                memberno = .ListItems(i).Text
                totalamount = .ListItems(i).SubItems(2)
                Me.Caption = "Posting Batch Please wait ....." & memberno
                FullName = .ListItems(i).SubItems(1)
                'xxxxxxxxxxxxxxxxxxx POST THE TRANSACTIONS WRT PRIORITY xxxxxxxxxxxxxxx
                'xxxx Loans (Interest then Principal) and then Shares
                'XXXXXXXXXXXXXXXXXX SYSTEM INTELLIGENCE XXXXXXXXXXXXXXXXXXXXX
                
                Set rsPostLoans = Nothing
                
                'Get the Batch posting Option from SYSPARAM=== either use Loan Priority or Date of Issue(FIFO)-refers to Loan
                LoanOption = Get_Society_Details
                Select Case LoanOption.LoanRecoveryMethod
                
                 Case 0 'Recover using priority i.e. Loans recovered using priorities set, Shares comes Last
                    Set rsPostLoans = oSaccoMaster.GetRecordSet("Select * from TMPPOSTING where MemberNo='" & memberno & "' and Description='Loan' order by MemerNo,Priority asc")
                    
                 Case 1 'use the date of issue of the Loan => First Loan is the First to be Recovered
                    Set rsPostLoans = oSaccoMaster.GetRecordSet("select TP.*,LB.FirstDate from TMPPOSTING TP inner join " _
                        & " LOANBAL LB  on LB.LoanNo=TP.LoanNo where TP.MemberNo='" _
                        & memberno & "' and Description='Loan' order by TP.MemberNo,LB.FirstDate asc")
                        
                End Select
                
                If Not rsPostLoans.EOF Then
                    While Not rsPostLoans.EOF
                        If totalamount > 0 Then
                            interest = CDbl(rsPostLoans!interest)
                            totalamount = CDbl(totalamount - interest) 'get the balance from the received amount -- deduct interest first
                            
                            If Not totalamount <= 0 Then
                                Principal = CDbl(rsPostLoans!Principal)
                                If totalamount <= Principal Then
                                    Principal = totalamount
                                    totalamount = 0
                                Else
                                    totalamount = CDbl(totalamount - Principal) 'get the balance from the received amount -- deduct principal
                                    
                                End If
                                
                            Else
                                Principal = 0
                            End If
                                'post the amount
                                Call frmTransactionPosting.Post_Loan_Repayment(memberno, rsPostLoans!LoanNo, Principal, interest, frmTransactionPosting.txtChequeNo, FullName)
                         End If
                         rsPostLoans.MoveNext
                    Wend
                End If
                
                'Post the remaining amount to Shares
                If Format(totalamount, Cfmt) > 0 Then
                    Set rsPostShares = Nothing
                    Set rsPostShares = oSaccoMaster.GetRecordSet("Select * from TMPPOSTING where MemberNo='" & memberno & "' and Description='Shares'")
                    If Not rsPostShares.EOF Then
                        Shares = CDbl(rsPostShares!amount)
                        If Not totalamount <= 0 Then
                            
                           If rsPostShares.RecordCount > 1 Then
                                While Not rsPostShares.EOF
                                    If Not rsPostShares.EOF Then
                                        maxRec = rsPostShares.RecordCount
                                        If maxRec > 1 Then
                                            Call frmTransactionPosting.Post_Share_contribution(memberno, rsPostShares!SharesCode, Shares, frmTransactionPosting.txtChequeNo, FullName)
                                            totalamount = totalamount - Shares
                                            maxRec = maxRec - 1
                                            
                                        End If
                                    End If
                                    rsPostShares.MoveNext
                                Wend
                                                                
                            Else
                    '//CHECK IF IT IS THE FIRST TIME TO BE RECOVERED THE MONEY
                    Dim RegFees As Double
                    Dim ByLaws As Double
                    
                    mysql = ""
                    mysql = "select * from contrib  where memberno ='" & Trim(memberno) & "'"
                    
                    Set rst = oSaccoMaster.GetRecordSet(mysql)
                    If rst.EOF Then
                        '//this guy might have never conributed the registration fee
                        If totalamount > 600 Then
                        
                            RegFees = 500
                            totalamount = totalamount - RegFees
                            ByLaws = 100
                            totalamount = totalamount - ByLaws
                                                If Not Save_To_GL("L099", "E014", RegFees, memberno, memberno, _
                                                dtpCashBookDate, memberno, memberno, ErrorMessage) Then
                                                    If ErrorMessage <> "" Then
                                                        MsgBox ErrorMessage, vbInformation, Me.Caption
                                                        ErrorMessage = ""
                                                    End If
                                                End If
                                                If Not Save_To_GL("L099", "I003", ByLaws, memberno, memberno, _
                                                dtpCashBookDate, memberno, memberno, ErrorMessage) Then
                                                    If ErrorMessage <> "" Then
                                                        MsgBox ErrorMessage, vbInformation, Me.Caption
                                                        ErrorMessage = ""
                                                    End If
                                                End If
                        Else
                        
                        End If
                   
                    End If
                    
                                Call frmTransactionPosting.Post_Share_contribution(memberno, rsPostShares!SharesCode, totalamount, frmTransactionPosting.txtChequeNo, FullName)
                                '//update gltransactions
                                
                            End If
                        End If
                    End If
                End If
        Next i
 End If
End With
    
    'Post the Batch to Batch Master
    Call frmTransactionPosting.Update_BatchMaster(frmTransactionPosting.txtChequeNo, CDbl(frmTransactionPosting.txtChequeAmount), frmTransactionPosting.cboCompanyCode, frmTransactionPosting.dtpTransDate)
    '//update
       strSQL = "UPDATE    GLTRANSACTIONS  SET     doc_posted=1            WHERE     (DocumentNo = '" & frmTransactionPosting.txtdocumentno & "')"
        oSaccoMaster.ExecuteThis strSQL
        
        lvwConfirmPost.ListItems.Clear
    'Set rs = oSaccoMaster.ExecuteThis()
    frmTransactionPosting.lvwBatch.ListItems.Clear
    sql = ""
    sql = "DELETE   FROM         TMPPOSTING WHERE AUDITID='" & User & "'"
    oSaccoMaster.ExecuteThis (sql)
       MsgBox "Batch Posted Successfully.", vbInformation, Me.Caption
       Me.Caption = "Transaction Posting"
hell:
Exit Sub
SysError:
     MsgBox Err.Description
End Sub

Private Sub Form_Load()
    On Error GoTo SysErr
    With frmTransactionPosting.lvwBatch
        If .ListItems.count >= 1 Then
            For i = 1 To .ListItems.count
                If .ListItems(i).Checked = True Then
                    Set li = lvwConfirmPost.ListItems.Add(, , .ListItems(i).Text)
                    li.SubItems(1) = .ListItems(i).SubItems(1)
                    li.SubItems(2) = .ListItems(i).SubItems(2)
                End If
            Next i
            lblTotal = frmTransactionPosting.lblTotal
            txtChequeNo = frmTransactionPosting.txtChequeNo
            txtChequeAmount = frmTransactionPosting.txtChequeAmount
            dtpCashBookDate = frmTransactionPosting.dtpCashBookDate
            dtpTransDate = frmTransactionPosting.dtpTransDate
            
        End If
    End With
    Exit Sub
SysErr:
    MsgBox Err.Description
End Sub
