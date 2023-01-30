VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRejectionReasons 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Rejection Reasons"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7065
   Icon            =   "frmRejectionReasons.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Height          =   495
      Left            =   4440
      Picture         =   "frmRejectionReasons.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Cancel Process"
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton cmdUpdate 
      Height          =   495
      Left            =   3960
      Picture         =   "frmRejectionReasons.frx":040C
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Save Record"
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton cmdDelete 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      Picture         =   "frmRejectionReasons.frx":04FE
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Delete Record"
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton cmdEdit 
      Height          =   495
      Left            =   3000
      Picture         =   "frmRejectionReasons.frx":05F0
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Edit Record"
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton cmdAdd 
      Height          =   495
      Left            =   2520
      Picture         =   "frmRejectionReasons.frx":06F2
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Add New record"
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "CLOSE"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   6
      Top             =   3000
      Width           =   1230
   End
   Begin VB.Frame Frame1 
      Height          =   2745
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6840
      Begin VB.TextBox txtRejectionReason 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1440
         TabIndex        =   5
         Top             =   1200
         Width           =   5250
      End
      Begin VB.TextBox txtReasonId 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Reason ID"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   1260
      End
      Begin VB.Label Label1 
         Caption         =   "Rejection Reason"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1440
         TabIndex        =   2
         Top             =   960
         Width           =   1860
      End
   End
   Begin MSComctlLib.ListView lvwRejectionReasons 
      Height          =   2700
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6810
      _ExtentX        =   12012
      _ExtentY        =   4763
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   16711680
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
End
Attribute VB_Name = "frmRejectionReasons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim disablemodifying As Boolean
Private Sub cmdAdd_Click()
On Error GoTo errFix
        action = "addingRecords"
        cmdadd.Enabled = False
        cmdUpdate.Enabled = True
        cmdDelete.Enabled = False
        cmdCancel.Enabled = True
        cmdEdit.Enabled = False
        txtReasonId.Locked = False
        txtRejectionReason = False
        lvwRejectionReasons.Visible = False
        Frame1.Visible = True
        txtReasonId.Text = ""
        txtRejectionReason.Text = ""
Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Rejection Reasons"
End Sub

Private Sub cmdAdd_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    txtReasonId.SetFocus
    cmdAdd_Click
End If
Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Rejection Reasons"
End Sub

Private Sub cmdCancel_Click()
On Error GoTo errFix
    action = ""
    lvwRejectionReasons.Visible = True
    Frame1.Visible = True
    cmdadd.Enabled = True
    cmdEdit.Enabled = True
    cmdUpdate.Enabled = False
    cmdDelete.Enabled = True
    cmdCancel.Enabled = False
    Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Rejection Reasons"
End Sub

Private Sub cmdCancel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdCancel_Click
End If
End Sub

Private Sub cmdclose_Click()
    Unload Me
    action = ""
End Sub

Private Sub cmdClose_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdclose_Click
End If
End Sub

Private Sub cmddelete_Click()
On Error GoTo errFix
    sel = lvwRejectionReasons.SelectedItem
    Set Rst = oSaccoMaster.GetRecordset("select * from reasons where reasonid= " & sel & "")
    If Rst.RecordCount > 0 Then
        If (MsgBox("Do you want to delete rejection reason '" & Rst!description & "'" & "?", vbYesNo, "Reasons") = vbYes) Then
            If Not Rst.EOF Then
                Rst.Delete
                Rst.Update
            End If
            Set Rst = oSaccoMaster.GetRecordset("select * from reasons order by reasonid")
            If Rst.RecordCount > 0 Then
            Rst.MoveFirst
            lvwRejectionReasons.ListItems.Clear
            While Not Rst.EOF
                 Set li = lvwRejectionReasons.ListItems.Add(, , Rst!reasonid & "")
                 li.SubItems(1) = Rst!description & ""
                 Rst.MoveNext
            Wend
            End If
        End If
    End If
    Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Rejection Reasons"
End Sub

Private Sub cmdDelete_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmddelete_Click
End If
End Sub

Private Sub cmdedit_Click()
On Error GoTo errFix
    action = "editingRecords"
    cmdUpdate.Enabled = True
    cmdEdit.Enabled = True
    cmdadd.Enabled = False
    cmdDelete.Enabled = False
    cmdCancel.Enabled = True
    txtRejectionReason.Locked = False
    txtReasonId.Locked = True
    lvwRejectionReasons.Visible = False
    Frame1.Visible = True
        sel = lvwRejectionReasons.SelectedItem
        txtReasonId.Text = sel
        Set Rst = oSaccoMaster.GetRecordset("select * from reasons where reasonid= " & txtReasonId.Text & "")
        If Rst.RecordCount > 0 Then
            txtRejectionReason.Text = Rst!description & ""
        End If
        lvwRejectionReasons.Visible = False
        Frame1.Visible = True
        Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Rejection Reasons"
End Sub

Private Sub cmdEdit_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdedit_Click
    txtRejectionReason.SetFocus
End If
End Sub

Private Sub cmdHelp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    'cmdHelp_Click
End If
End Sub

Private Sub cmdupdate_Click()
On Error GoTo errFix
If action = "addingRecords" Then
    While (txtReasonId = "" Or txtRejectionReason = "")
        MsgBox "Make sure you input both ID and Description", vbInformation, "Rejection Reasons"
        Exit Sub
    Wend
    Set Rst5 = oSaccoMaster.GetRecordset("select reasonid from reasons where reasonid= " & txtReasonId.Text & "")
        If Not Rst5.EOF Then
            MsgBox "Reason with same ID exists", vbInformation, "Reason ID"
            txtReasonId.Text = ""
            txtReasonId.SetFocus
            Exit Sub
        End If
    
    
    Set Rst = oSaccoMaster.GetRecordset("select * from reasons")
    If Not (txtReasonId.Text = "") And Not (txtRejectionReason.Text = "") Then
     With Rst
        .AddNew
       !reasonid = txtReasonId.Text & ""
        !description = txtRejectionReason.Text & ""
        .Update
     End With
    Else
    
    End If
    cmdadd.Enabled = True
    cmdUpdate.Enabled = False
    cmdDelete.Enabled = True
    cmdCancel.Enabled = False
    cmdEdit.Enabled = True
    cmdadd.Enabled = True
    Frame1.Visible = True
    lvwRejectionReasons.Visible = True
    Set Rst = oSaccoMaster.GetRecordset("select * from reasons")
    If Rst.RecordCount > 0 Then
    Rst.MoveFirst
    lvwRejectionReasons.ListItems.Clear
    While Not Rst.EOF
         Set li = lvwRejectionReasons.ListItems.Add(, , Rst!reasonid & "")
         li.SubItems(1) = Rst!description & ""
         Rst.MoveNext
    Wend
    End If
ElseIf action = "editingRecords" Then
    While (txtReasonId = "" Or txtRejectionReason = "")
        MsgBox "Make sure you input both ID and Description", vbOKOnly, "Rejection Reasons"
        Exit Sub
    Wend
        Set Rst = oSaccoMaster.GetRecordset("select * from reasons where reasonid= " & txtReasonId.Text & "")
        With Rst
            !reasonid = txtReasonId.Text & ""
            !description = txtRejectionReason & ""
            !auditid = User
            !audittime = Format(Now, "DD/MM/YYYY hh:mm:ss")
            cmdEdit.Enabled = True
            cmdadd.Enabled = True
            cmdUpdate.Enabled = False
            cmdCancel.Enabled = False
            txtRejectionReason.Locked = True
            txtReasonId.Locked = True
            lvwRejectionReasons.Visible = True
            Frame1.Visible = False
            .Update
        End With
        Set Rst = oSaccoMaster.GetRecordset("select * from reasons order by reasonid")
        If Rst.RecordCount > 0 Then
        Rst.MoveFirst
        lvwRejectionReasons.ListItems.Clear
        While Not Rst.EOF
             Set li = lvwRejectionReasons.ListItems.Add(, , Rst!reasonid & "")
             li.SubItems(1) = Rst!description & ""
             Rst.MoveNext
        Wend
        End If
End If
action = ""
Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Rejection Reasons"
End Sub

Private Sub cmdUpdate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdupdate_Click
    txtReasonId.SetFocus
End If

End Sub

Private Sub Form_Load()
On Error GoTo errFix
    PositionForm Me
    With lvwRejectionReasons
        .ColumnHeaders.Clear
        .ListItems.Clear
        .ColumnHeaders.Add 1, "hReasonId", "Reason ID", 2000
        .ColumnHeaders.Add 2, "hDescription", "Description", 5000
        Set Rst = oSaccoMaster.GetRecordset("select * from reasons order by reasonid")
        If Rst.RecordCount > 0 Then
            Rst.MoveFirst
            txtReasonId.Text = Rst!reasonid & ""
            While Not Rst.EOF
                 Set li = lvwRejectionReasons.ListItems.Add(, , Rst!reasonid & "")
                 li.SubItems(1) = Rst!description & ""
                 Rst.MoveNext
            Wend
        End If
    
    End With
    lvwRejectionReasons.Visible = True
    Frame1.Visible = False
    cmdCancel.Enabled = False
    cmdUpdate.Enabled = False
    'allow only viewing for certain users
          Set Rst = oSaccoMaster.GetRecordset("select groupid from users where userid = '" & User & "'")
           If Not Rst.EOF Then
           Set Rst1 = oSaccoMaster.GetRecordset("select * from usergrps where groupid= '" & Rst!groupid & "'")
           If Not Rst1.EOF Then
            valToEncrOrDecr = Rst1!memreg & vbNullString
            EncryptPassword
            If EncryptPass = "View" Then
                cmdadd.Enabled = False
                cmdEdit.Enabled = False
                cmdUpdate.Enabled = False
                cmdCancel.Enabled = False
                cmdDelete.Enabled = False
            End If
            End If
            End If
            Set Rst4 = oSaccoMaster.GetRecordset("select groupid from users where userid= '" & User & "'")
            If Not Rst4.EOF Then
                Set Rst5 = oSaccoMaster.GetRecordset("select * from usergrps where groupid= '" & Rst4!groupid & "'")
                If Not Rst5.EOF Then
                    valToEncrOrDecr = Rst5!rejreasons & vbNullString
                    EncryptPassword
                    If EncryptPass = "View" Then
                        disablemodifying = True
                        cmdadd.Enabled = False
                        cmdUpdate.Enabled = False
                        cmdCancel.Enabled = False
                        cmdDelete.Enabled = False
                        cmdEdit.Enabled = False
                    Else
                        disablemodifying = False
                    End If
                End If
            End If
            Exit Sub
errFix:
'   MsgBox Err.Description, vbOKOnly, "Rejection Reasons"
End Sub
Function EncryptPassword()
On Error GoTo errFix
    Dim Pwd As Variant
    Dim Temp As String, PwdChr As Long
    Dim EncryptKey As Long
    Pwd = valToEncrOrDecr
    EncryptKey = Int(Sqr(Len(Pwd) * 81)) + 23
    
    For PwdChr = 1 To Len(Pwd)
        Temp = Temp + Chr(Asc(Mid(Pwd, PwdChr, 1)) Xor EncryptKey)
    Next PwdChr
    
    EncryptPass = Temp
Exit Function
errFix:
    MsgBox err.description, vbOKOnly, "Rejection Reasons"
End Function

Private Sub Form_Unload(Cancel As Integer)
action = ""
End Sub

Private Sub lvwRejectionReasons_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If lvwRejectionReasons.ListItems.Count > 0 Then
        sel = lvwRejectionReasons.SelectedItem
    End If
End Sub

Private Sub txtRejectionReason_Change()
txtRejectionReason.Text = UCase(txtRejectionReason.Text)
txtRejectionReason.SelStart = Len(txtRejectionReason.Text)
End Sub



Private Sub txtReasonId_KeyPress(KeyAscii As Integer)
    On Error GoTo errFix
    If KeyAscii = 13 Then
        Set Rst5 = oSaccoMaster.GetRecordset("select reasonid from reasons where reasonid= " & txtReasonId.Text & "")
        If Not Rst5.EOF And Not txtReasonId = "" Then
            MsgBox "Reason with same ID exists", vbInformation, "Reason ID"
            txtReasonId.Text = ""
            txtReasonId.SetFocus
            Exit Sub
        Else
            txtRejectionReason.SetFocus
        End If
    End If
    If Len(txtReasonId.Text) > 1 Then
        Beep
        MsgBox "Can't enter more than 2 characters", vbExclamation
        KeyAscii = 8
    End If
    
  Select Case KeyAscii
    Case Asc("0") To Asc("9")
    Case Is = 8
    Case Else
        Beep
        KeyAscii = 0
  End Select
  Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Rejection Reasons"
End Sub
Private Sub txtReasonId_LostFocus()
On Error GoTo errFix
     If action = "addingRecords" Then
        If Not IsNumeric(txtReasonId.Text) And Not txtReasonId.Text = "" Then
           MsgBox "Input a number for reason ID", vbInformation, "Reason ID"
           txtReasonId.Text = ""
           txtReasonId.SetFocus
        End If
        If Not txtReasonId.Text = "" Then
           Set Rst5 = oSaccoMaster.GetRecordset("select reasonid from reasons where reasonid= " & txtReasonId.Text & "")
           If Not Rst5.EOF Then
               MsgBox "Reason with same ID exists", vbInformation, "Reason ID"
               txtReasonId.Text = ""
               txtReasonId.SetFocus
               Exit Sub
           End If
        End If
    End If
    Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Rejection Reasons"
End Sub

Private Sub txtRejectionReason_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
    If KeyAscii = 13 Then
        cmdUpdate.SetFocus
    End If
    If Len(txtRejectionReason.Text) > 99 Then
        Beep
        MsgBox "Can't enter more than 100 characters", vbExclamation
        KeyAscii = 8
    End If
    Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Rejection Reasons"
End Sub
