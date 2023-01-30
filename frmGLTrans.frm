VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGLTrans 
   Caption         =   "GL Transactions"
   ClientHeight    =   3270
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8565
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   8565
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtVoucherNo 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1530
      TabIndex        =   3
      Top             =   435
      Width           =   3075
   End
   Begin VB.TextBox txtTransDate 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   45
      TabIndex        =   1
      Top             =   420
      Width           =   1290
   End
   Begin MSComctlLib.ListView lvwTransactions 
      Height          =   1815
      Left            =   15
      TabIndex        =   0
      Top             =   840
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   3201
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
      Appearance      =   0
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Account Name"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Debit Amount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Credit Amount"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   315
      Left            =   2640
      TabIndex        =   5
      Top             =   2835
      Width           =   1605
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Voucher No"
      Height          =   210
      Left            =   1560
      TabIndex        =   4
      Top             =   180
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Trans Date"
      Height          =   210
      Left            =   60
      TabIndex        =   2
      Top             =   165
      Width           =   900
   End
End
Attribute VB_Name = "frmGLTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
    On Error GoTo SysError
    Dim rsTrans As New Recordset
    Set rsTrans = oSaccoMaster.GetRecordSet("Select * From CustomerBalance where " _
    & "VNo='" & txtVoucherNo & "' and TransDate='" & Format(txtTransDate, "MM/dd/yyyy") & "'")
    With rsTrans
        If .State = adStateOpen Then
            While Not .EOF
                Set li = lvwTransactions.ListItems.Add(, , !accname)
                If !transtype = "DR" Then
                    li.SubItems(1) = Format(!amount, "###,###,###,##0.00")
                    li.SubItems(2) = "0.00"
                Else
                    li.SubItems(1) = "0.00"
                    li.SubItems(2) = Format(!amount, "###,###,###,##0.00")
                End If
                .MoveNext
            Wend
        End If
    End With
    Exit Sub
SysError:
    MsgBox Err.description, vbInformation, Me.Caption
End Sub

