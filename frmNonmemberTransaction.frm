VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmNonmemberTransaction 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Non Member Ledger Transaction"
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10770
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   10770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmddelete 
      Caption         =   "Delete"
      Height          =   360
      Left            =   8265
      TabIndex        =   5
      Top             =   210
      Width           =   1470
   End
   Begin VB.TextBox txtAccno 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1395
      TabIndex        =   3
      Top             =   225
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   7110
      Left            =   15
      TabIndex        =   0
      Top             =   750
      Width           =   10710
      Begin MSComctlLib.ListView lsvTransaction 
         Height          =   6870
         Left            =   15
         TabIndex        =   1
         Top             =   105
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   12118
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
   End
   Begin VB.Label lblname 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2985
      TabIndex        =   4
      Top             =   240
      Width           =   4410
   End
   Begin VB.Label Label1 
      Caption         =   "Accno"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   225
      Width           =   930
   End
End
Attribute VB_Name = "frmNonmemberTransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDelete_Click()
If lsvTransaction.ListItems.Count > 0 Then
    If lsvTransaction.SelectedItem.Text <> "" Then
        If MsgBox("Are you  sure you want to delete  " & lsvTransaction.SelectedItem.ListSubItems(4).Text & " ?", vbYesNo) = vbYes Then
            mysql = "delete  from GLTRANSACTIONS  where id =" & lsvTransaction.SelectedItem.ListSubItems(6).Text & ""
            oSaccoMaster.ExecuteThis (mysql)
        End If
    End If
End If
Load_data

End Sub

Private Sub Form_Load()
headers
End Sub

Private Sub txtAccno_Change()

Dim RsRecords As New ADODB.Recordset
''//get accno details
mysql = "select * from glsetup where accno ='" & txtAccNo & "'"

Set RsRecords = oSaccoMaster.GetRecordSet(mysql)

If Not RsRecords.EOF Then
    lblname = RsRecords!GlAccName & ""
Else
    lblname = ""
End If
mysql = ""
mysql = "select "
mysql = ""
mysql = "select * from GLTRANSACTIONS where Draccno ='" & txtAccNo & "' or CrAccno ='" & txtAccNo & "' "

Set RsRecords = oSaccoMaster.GetRecordSet(mysql)
If Not RsRecords.EOF Then
    Load_data
Else
    If lsvTransaction.ListItems.Count > 0 Then
        lsvTransaction.ListItems.Clear
    End If
End If
End Sub
Private Sub Load_data()
Dim RsRecords As New ADODB.Recordset
 mysql = ""
 mysql = "select * from GLTRANSACTIONS where Draccno ='" & txtAccNo & "' or CrAccno ='" & txtAccNo & "' order by TransDate,id"
 
 Set RsRecords = oSaccoMaster.GetRecordSet(mysql)
 
 If Not RsRecords.EOF Then
    lsvTransaction.ListItems.Clear
    Do While Not RsRecords.EOF
            Set li = lsvTransaction.ListItems.Add(, , RsRecords!transdate)
                li.ListSubItems.Add , , RsRecords!amount
                li.ListSubItems.Add , , RsRecords!DocumentNo & ""
                li.ListSubItems.Add , , RsRecords!Source & ""
                li.ListSubItems.Add , , RsRecords!TransDescript & ""
                li.ListSubItems.Add , , RsRecords!id & ""
        RsRecords.MoveNext
    Loop
 Else
 lsvTransaction.ListItems.Clear
 End If
End Sub

Private Sub headers()
    With lsvTransaction
    .ColumnHeaders.Add , , "Transdate"
    .ColumnHeaders.Add , , "Amount"
    .ColumnHeaders.Add , , "Document No"
    .ColumnHeaders.Add , , "Source"
    .ColumnHeaders.Add , , "Transdescription"
    .View = lvwReport
    .GridLines = True
    .FullRowSelect = True
    .LabelEdit = lvwManual
    
    End With
End Sub
