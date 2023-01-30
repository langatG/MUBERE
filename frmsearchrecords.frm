VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmsearchrecords 
   Caption         =   "SEARCH RECORDS"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4815
   Icon            =   "frmsearchrecords.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4815
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2475
      Left            =   30
      TabIndex        =   2
      Top             =   0
      Width           =   4800
      Begin MSComctlLib.ListView lsvLedgers 
         Height          =   2370
         Left            =   30
         TabIndex        =   3
         Top             =   105
         Width           =   4770
         _ExtentX        =   8414
         _ExtentY        =   4180
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Select"
      Enabled         =   0   'False
      Height          =   375
      Left            =   390
      TabIndex        =   1
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2430
      TabIndex        =   0
      Top             =   2520
      Width           =   1095
   End
End
Attribute VB_Name = "frmsearchrecords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
On Error GoTo 10
    strName = lsvLedgers.SelectedItem.Text
    'frmContributions.cbocharge1glaccno = DataList1.SelectedItem
    Unload Me
    Exit Sub
10:    MsgBox Err.description
End Sub

Private Sub Command2_Click()
 strName = ""
    Unload Me
End Sub

Private Sub DataList1_Click()
On Error Resume Next
    Command1.Enabled = True
End Sub

Private Sub DataList1_DblClick()
On Error Resume Next
    Call Command1_Click
End Sub

Private Sub Form_Load()
   ' On Error GoTo 10
    Dim myclass As cdbase
    Dim RsRecords As New ADODB.Recordset
    Dim strQ
    Dim cn As Connection
    Set cn = CreateObject("adodb.connection")
    Set myclass = New cdbase
        With lsvLedgers
            .Columnheaders.Add , , "Name", lsvLedgers.Width
            .View = lvwReport
            .Gridlines = True
            .FullRowSelect = True
            .LabelEdit = lvwManual
            
        End With
        
        
        mysql = "Select * from CUB where main=1 and ledger=1 order by cuid"
        
        Set RsRecords = oSaccoMaster.GetRecordSet(mysql)
        
        If Not RsRecords.EOF Then
            Do While Not RsRecords.EOF
                Set li = lsvLedgers.ListItems.Add(, , RsRecords!name & "")
                RsRecords.MoveNext
            Loop
        End If
    
    Exit Sub
10:    MsgBox Err.description
End Sub


Private Sub lsvLedgers_Click()
Command1.Enabled = True
End Sub

Private Sub lsvLedgers_DblClick()
Command1_Click
End Sub
