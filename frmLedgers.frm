VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLedgers 
   Caption         =   "Detailed Ledger Transactions"
   ClientHeight    =   5730
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11355
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   11355
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lvwTrans 
      Height          =   4350
      Left            =   75
      TabIndex        =   0
      Top             =   840
      Width           =   11220
      _ExtentX        =   19791
      _ExtentY        =   7673
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
         Text            =   "Amount"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "TransType"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "MemberNo"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Member Names"
         Object.Width           =   5292
      EndProperty
   End
End
Attribute VB_Name = "frmLedgers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    On Error GoTo SysError
    lvwTrans.ListItems.Clear
    Load_Ledgers MTransDate, TransNo, mDocNo
    Exit Sub
SysError:
    MsgBox Err.description, vbInformation, Me.Caption
End Sub

Private Sub Load_Ledgers(TransDate As Date, TransNo As String, DocNo As String)
    On Error GoTo SysError
    Dim rsLedger As New Recordset, TotAmount As Double, rsMember As New Recordset
    Set rsLedger = oSaccoMaster.GetRecordSet("Select * From CUSTOMERBALANCE where " _
    & "TransDate='" & TransDate & "' and VNo='" & TransNo & "' and ChequeNo='" & _
    DocNo & "' and AccNo<>'A001' order by CustomerBalanceID")
    With rsLedger
        If .State = adStateOpen Then
            While Not .EOF
                Set li = lvwTrans.ListItems.Add(, , IIf(IsNull(!accno), "", !accno))
                li.SubItems(1) = IIf(IsNull(!accname), "", !accname)
                li.SubItems(2) = Format(IIf(IsNull(!amount), 0, !amount), CfMt)
                li.SubItems(3) = IIf(IsNull(!transtype), "DR", !transtype)
                li.SubItems(4) = IIf(IsNull(!payrollno), "", !payrollno)
                TotAmount = TotAmount + CDbl(li.SubItems(2))
                Set rsMember = oSaccoMaster.GetRecordSet("Select OtherNames + ' ' + " _
                & "SurName as [Names] from MEMBERS where MemberNo='" & !payrollno & "'")
                With rsMember
                    If .State = adStateOpen Then
                        If Not .EOF Then
                            li.SubItems(5) = IIf(IsNull(![names]), "", ![names])
                        End If
                    End If
                End With
                rsMember.Close
                .MoveNext
            Wend
            Set li = lvwTrans.ListItems.Add(, , "Totals")
            li.SubItems(2) = Format(TotAmount, CfMt)
        End If
    End With
    TotAmount = 0
    Exit Sub
SysError:
    MsgBox Err.description, vbInformation, Me.Caption
End Sub
