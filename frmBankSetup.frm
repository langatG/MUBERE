VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmBankSetup 
   Appearance      =   0  'Flat
   BackColor       =   &H80000004&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "  Bank Setup"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8025
   Icon            =   "frmBankSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   8025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Height          =   495
      Left            =   4680
      Picture         =   "frmBankSetup.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Cancel Process"
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton cmdUpdate 
      Height          =   495
      Left            =   4200
      Picture         =   "frmBankSetup.frx":040C
      Style           =   1  'Graphical
      TabIndex        =   7
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
      Left            =   3720
      Picture         =   "frmBankSetup.frx":050E
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Delete Record"
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton cmdEdit 
      Height          =   495
      Left            =   3240
      Picture         =   "frmBankSetup.frx":0610
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Edit Record"
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1095
      TabIndex        =   2
      ToolTipText     =   "Move to the Next"
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1575
      Picture         =   "frmBankSetup.frx":0712
      TabIndex        =   3
      ToolTipText     =   "Move to Last record"
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   1
      ToolTipText     =   "Move to the Previous record"
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Move to the Last record"
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton cmdAdd 
      Height          =   495
      Left            =   2760
      Picture         =   "frmBankSetup.frx":0814
      Style           =   1  'Graphical
      TabIndex        =   4
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
      Left            =   6600
      TabIndex        =   23
      Top             =   3000
      Width           =   1230
   End
   Begin MSComctlLib.ListView lvwSummary 
      Height          =   2295
      Left            =   120
      TabIndex        =   22
      Top             =   3600
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   4048
      View            =   3
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
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Frame fraBank 
      Height          =   2175
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   7680
      Begin VB.TextBox txtBankCode 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Top             =   375
         Width           =   1575
      End
      Begin VB.TextBox txtBankName 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1800
         TabIndex        =   10
         Top             =   375
         Width           =   5775
      End
      Begin VB.TextBox txtNoOfMembers 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2280
         TabIndex        =   14
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox txtTelephone 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Top             =   1695
         Width           =   2055
      End
      Begin VB.TextBox txtAddress 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   3360
         TabIndex        =   12
         Top             =   960
         Width           =   4215
      End
      Begin VB.TextBox txtBranchName 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "Bank Code"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   105
         TabIndex        =   21
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Bank Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   20
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "No. of Members"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2280
         TabIndex        =   19
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "Phone No"
         BeginProperty Font 
            Name            =   "Verdana"
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
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3315
         TabIndex        =   17
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Branch Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   1815
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6360
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBankSetup.frx":0D46
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBankSetup.frx":0E58
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBankSetup.frx":0F6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBankSetup.frx":107C
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmBankSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim disablemodifying As Boolean
Private Sub cmdadd_Click()
    On Error GoTo errFix
    action = "addingRecords"
    
    Dim ctrl As Control
    For Each ctrl In Me
        If TypeOf ctrl Is TextBox Then
            ctrl.Locked = False
        End If
        
        If TypeOf ctrl Is ComboBox Then
            ctrl.Locked = False
        End If

        If TypeOf ctrl Is CheckBox Then
                ctrl.Enabled = True
        End If
    Next ctrl
    Set Rst3 = oSaccoMaster.GetRecordSet("select * from d_BANKS where bankcode= '" & txtBankCode.Text & "'")
    If Rst3.RecordCount > 0 Then
        MyBookMark = Rst3.Bookmark
    End If
    lvwSummary.Visible = True
    fraBank.Visible = True
    fraBank.Enabled = True
    cmdEdit.Enabled = False
    cmdCancel.Enabled = True
    cmdAdd.Enabled = False
    cmdDelete.Enabled = False
    cmdUpdate.Enabled = True
    txtBankCode.Text = ""
    txtBankName.Text = ""
    txtBranchName.Text = ""
    txtTelephone.Text = ""
    txtAddress.Text = ""
    txtNoOfMembers.Locked = True
    txtBankCode.Locked = False
    txtBankName.Locked = False
    txtTelephone.Locked = False
    txtAddress.Locked = False
    Exit Sub
errFix:
    MsgBox Err.Description, vbOKOnly, "Bank Setup"
End Sub

Private Sub cmdAdd_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    cmdadd_Click
    txtBankCode.SetFocus
End If
Exit Sub
errFix:
    MsgBox Err.Description, vbOKOnly, "Bank Setup"
End Sub

Private Sub cmdCancel_Click()
On Error GoTo errFix
action = ""

Dim ctrl As Control
For Each ctrl In Me
        If TypeOf ctrl Is TextBox Then
            ctrl.Locked = True
        End If
        
        If TypeOf ctrl Is ComboBox Then
            ctrl.Locked = True
        End If
        
        If TypeOf ctrl Is CheckBox Then
                ctrl.Enabled = False
        End If
    Next ctrl
Set Rst3 = oSaccoMaster.GetRecordSet("select * from d_BANKS")
If Rst3.RecordCount > 0 Then
    'Rst3.Bookmark = MyBookMark
    txtBankCode.Text = Rst3!BankCode & ""
    load_records
End If
fraBank.Enabled = False
cmdAdd.Enabled = True
cmdEdit.Enabled = True
cmdCancel.Enabled = False
cmdDelete.Enabled = True
cmdUpdate.Enabled = False
Exit Sub
errFix:
    MsgBox Err.Description, vbOKOnly, "Bank Setup"
End Sub

Private Sub cmdCancel_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    cmdCancel_Click
End If
Exit Sub
Exit Sub
errFix:
    MsgBox Err.Description, vbOKOnly, "Bank Setup"
End Sub

Private Sub cmdclose_Click()
On Error GoTo errFix
Unload Me
action = ""
Exit Sub
errFix:
    MsgBox Err.Description, vbOKOnly, "Bank Setup"
End Sub



Private Sub cmdClose_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    cmdclose_Click
End If
Exit Sub
errFix:
    MsgBox Err.Description, vbOKOnly, "Bank Setup"
End Sub

Private Sub cmddelete_Click()
On Error GoTo errFix
If lvwSummary.ListItems.Count > 0 Then
    sel = lvwSummary.SelectedItem
End If
If lvwSummary.Visible Then

    Set rst = oSaccoMaster.GetRecordSet("select * from banks where bankcode= '" & sel & "'")
    If rst.RecordCount > 0 Then
        If MsgBox("Are you sure you want to delete " & rst!BankName & "" & rst!BranchName & " ? ", vbYesNo, "Bank deletion") = vbYes Then
            Set Rst5 = oSaccoMaster.GetRecordSet("select * from banks where bankcode= '" & sel & "'")
            If Not Rst5.EOF Then
                Rst5.Delete
                Rst5.Update
            End If
             
             
            load_records
         End If
     End If
 
Else
    Set rst = oSaccoMaster.GetRecordSet("select * from d_BANKS where bankcode= '" & txtBankCode.Text & "'")
    If rst.RecordCount > 0 Then
    If MsgBox("Are you sure you want to delete " & rst!BankName & "" & rst!BranchName & "" & " ? ", vbYesNo, "bank deletion") = vbYes Then
        Set rst = oSaccoMaster.GetRecordSet("select * from d_BANKS where bankcode= '" & txtBankCode.Text & "'")
        rst.Delete
        rst.Update
        Set Rst5 = oSaccoMaster.GetRecordSet("select * from d_BANKS")
        Rst5.MoveFirst
        txtBankCode.Text = Rst5!BankCode & ""
        load_records
     End If
 End If
End If
Exit Sub
errFix:
    MsgBox Err.Description, vbOKOnly, "Bank Setup"
End Sub

Private Sub cmdDelete_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    cmddelete_Click
End If
Exit Sub
errFix:
    MsgBox Err.Description, vbOKOnly, "Bank Setup"
End Sub

Private Sub cmdedit_Click()
On Error GoTo errFix
    action = "editingRecords"
    
    Dim ctrl As Control
    For Each ctrl In Me
        If TypeOf ctrl Is TextBox Then
        If Not ctrl = txtNoOfMembers And Not ctrl = txtBankCode Then
            ctrl.Locked = False
        End If
        End If
        
        If TypeOf ctrl Is ComboBox Then
            ctrl.Locked = False
        End If
        
        If TypeOf ctrl Is CheckBox Then
                ctrl.Enabled = True
        End If
    Next ctrl

    Set Rst3 = oSaccoMaster.GetRecordSet("select * from d_BANKS where bankcode= '" & txtBankCode.Text & "'")
    If Rst3.RecordCount > 0 Then
       ' MyBookMark = Rst3.Bookmark
    End If
    
    If lvwSummary.Visible Then
        sel = lvwSummary.SelectedItem
        lvwSummary.Visible = False
        fraBank.Visible = True
        txtBankCode.Text = sel
        load_records
        
    End If
    txtNoOfMembers.Locked = True
    cmdAdd.Enabled = False
    cmdEdit.Enabled = False
    fraBank.Enabled = True
    cmdCancel.Enabled = True
    cmdDelete.Enabled = False
    cmdUpdate.Enabled = True
    txtBankCode.Locked = True
    txtBankName.Locked = False
    txtTelephone.Locked = False
    txtAddress.Locked = False
    Exit Sub
errFix:
    MsgBox Err.Description, vbOKOnly, "Bank Setup"
End Sub

Private Sub cmdEdit_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    cmdedit_Click
    txtBankName.SetFocus
End If
Exit Sub
errFix:
    MsgBox Err.Description, vbOKOnly, "Bank Setup"
End Sub

Private Sub cmdFirst_Click()
On Error GoTo errFix

Dim ctrl As Control
For Each ctrl In Me
        If TypeOf ctrl Is TextBox Then
            ctrl.Locked = True
        End If
        
        If TypeOf ctrl Is ComboBox Then
            ctrl.Locked = True
        End If
        
        If TypeOf ctrl Is CheckBox Then

        End If
    Next ctrl
Set rst1 = oSaccoMaster.GetRecordSet("select bankcode from d_banks order by bankcode")
With rst1
    If .RecordCount > 0 Then
        .MoveFirst
        txtBankCode.Text = rst1!BankCode & ""
        load_records
        cmdFirst.Enabled = False
        cmdPrevious.Enabled = False
        cmdNext.Enabled = True
        cmdLast.Enabled = True
    End If
End With

rst1.Close
action = ""
Exit Sub
errFix:
    MsgBox Err.Description, vbOKOnly, "Bank Setup"
End Sub

Private Sub cmdFirst_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    cmdFirst_Click
End If
Exit Sub
errFix:
    MsgBox Err.Description, vbOKOnly, "Bank Setup"
End Sub

Private Sub cmdHelp_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    'cmdHelp_Click
End If
Exit Sub
errFix:
    MsgBox Err.Description, vbOKOnly, "Bank Setup"
End Sub

Private Sub cmdLast_Click()
On Error GoTo errFix

Dim ctrl As Control
For Each ctrl In Me
        If TypeOf ctrl Is TextBox Then
            ctrl.Locked = True
        End If
        
        If TypeOf ctrl Is ComboBox Then
            ctrl.Locked = True
        End If
        
        If TypeOf ctrl Is CheckBox Then
            
        End If
    Next ctrl
Set rst1 = oSaccoMaster.GetRecordSet("select bankcode from d_banks order by bankcode")
With rst1
    If .RecordCount > 0 Then
        .MoveLast
        txtBankCode.Text = !BankCode & ""
        cmdFirst.Enabled = True
        cmdPrevious.Enabled = True
        cmdNext.Enabled = False
        cmdLast.Enabled = False
        load_records
    End If
End With

If action = "editingRecords" Or action = "addingRecords" Then
    If disablemodifying = False Then
    cmdAdd.Enabled = True
    cmdEdit.Enabled = True
    End If
    cmdCancel.Enabled = False
    cmdDelete.Enabled = True
End If
rst1.Close
action = ""
Exit Sub
errFix:
    MsgBox Err.Description, vbOKOnly, "Bank Setup"
End Sub

Private Sub cmdLast_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    cmdLast_Click
End If
Exit Sub
errFix:
    MsgBox Err.Description, vbOKOnly, "Bank Setup"
End Sub

Private Sub cmdNext_Click()
On Error GoTo errFix

Dim ctrl As Control
For Each ctrl In Me
        If TypeOf ctrl Is TextBox Then
            ctrl.Locked = True
        End If
        
        If TypeOf ctrl Is ComboBox Then
            ctrl.Locked = True
        End If
        
        If TypeOf ctrl Is CheckBox Then
        
        End If
    Next ctrl
Set rst1 = oSaccoMaster.GetRecordSet("select * from d_BANKS order by bankcode")
If cmdUpdate.Enabled = True Then
    If Not rst1.EOF Then
        'rst1.Bookmark = MyBookMark
        txtBankCode.Text = rst1!BankCode & ""
    End If
End If
With rst1
    If .RecordCount > 0 Then
        .Find "bankcode= '" & txtBankCode.Text & "'"
        If Not .EOF Then
            .MoveNext
            If .EOF Then
                .MoveLast
                cmdFirst.Enabled = True
                cmdPrevious.Enabled = True
                cmdNext.Enabled = False
                cmdLast.Enabled = False
            Else
                cmdFirst.Enabled = True
                cmdPrevious.Enabled = True
                cmdNext.Enabled = True
                cmdLast.Enabled = True
            End If
            txtBankCode.Text = !BankCode & ""
            load_records
        End If
    End If
End With
rst1.Close
action = ""
Exit Sub
errFix:
    MsgBox Err.Description, vbOKOnly, "Bank Setup"
End Sub

Private Sub cmdNext_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    cmdNext_Click
End If
Exit Sub
errFix:
    MsgBox Err.Description, vbOKOnly, "Bank Setup"
End Sub

Private Sub cmdPrevious_Click()
On Error GoTo errFix

Dim ctrl As Control
For Each ctrl In Me
        If TypeOf ctrl Is TextBox Then
            ctrl.Locked = True
        End If
        
        If TypeOf ctrl Is ComboBox Then
            ctrl.Locked = True
        End If
        
        If TypeOf ctrl Is CheckBox Then

        End If
    Next ctrl
Set rst1 = oSaccoMaster.GetRecordSet("select bankcode from d_BANKS order by bankcode")

If cmdUpdate.Enabled = True Then
    'rst1.Bookmark = MyBookMark
    txtBankCode.Text = rst1!BankCode & ""
End If
With rst1
    If Not rst1.EOF Then
        .MovePrevious
        .Find ("bankcode= '" & txtBankCode.Text & "'")
        If Not .EOF Then
            .MovePrevious
            If .BOF Then
                .MoveFirst
                cmdFirst.Enabled = False
                cmdPrevious.Enabled = False
                cmdNext.Enabled = True
                cmdLast.Enabled = True
            Else
                cmdFirst.Enabled = True
                cmdPrevious.Enabled = True
                cmdNext.Enabled = True
                cmdLast.Enabled = True
            End If
          txtBankCode.Text = !BankCode & ""
          load_records
        End If
    End If
End With
rst1.Close
action = ""
Exit Sub
errFix:
    MsgBox Err.Description, vbOKOnly, "Bank Setup"
End Sub

Private Sub cmdPrevious_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    cmdPrevious_Click
End If
Exit Sub
errFix:
    MsgBox Err.Description, vbOKOnly, "Bank Setup"
End Sub



Private Sub cmdupdate_Click()
On Error GoTo errFix
If action = "addingRecords" Then
    Set Rst5 = oSaccoMaster.GetRecordSet("select bankcode from d_BANKS where bankcode= '" & txtBankCode.Text & "'")
    If Not Rst5.EOF Then
        MsgBox "Bank with same code already exists", vbInformation, "Bank Code"
        txtBankCode.Text = ""
        txtBankCode.SetFocus
    Else
        txtBankName.SetFocus
    End If
    While txtBankName.Text = "" Or txtTelephone.Text = "" Or txtAddress.Text = ""
        MsgBox "Enter sufficient information", vbInformation, "Insufficient Information"
        Exit Sub
    Wend
    cmdEdit.Enabled = True
    cmdCancel.Enabled = False
    cmdDelete.Enabled = True
    cmdUpdate.Enabled = False
    cmdAdd.Enabled = True
    Set rst = oSaccoMaster.GetRecordSet("select bankcode, auditid, audittime, branchname, bankname,telephone,address from d_BANKS")
     With rst
            .AddNew
            !BankCode = txtBankCode.Text & ""
            !BankName = txtBankName.Text & ""
            !BranchName = txtBranchName.Text & ""
            !Telephone = txtTelephone.Text & ""
            !Address = txtAddress.Text & ""
            !auditid = User
            !audittime = Format(Now, "DD/MM/YYYY hh:mm:ss")
            
            .Update
    End With
ElseIf action = "editingRecords" Then
    While txtBankName.Text = "" Or txtTelephone.Text = "" Or txtAddress.Text = ""
        MsgBox "Enter sufficient information", vbInformation, "Insufficient Information"
        Exit Sub
    Wend
    cmdEdit.Enabled = True
    cmdCancel.Enabled = False
    cmdDelete.Enabled = True
    cmdUpdate.Enabled = False
    cmdAdd.Enabled = True
    Set rst = oSaccoMaster.GetRecordSet("select auditid,audittime,branchname,bankname,telephone,address from d_BANKS where bankcode='" & txtBankCode.Text & "'")
    If rst.RecordCount > 0 Then
     With rst
            !BankName = txtBankName.Text & ""
            !BranchName = txtBranchName.Text & ""
            !Telephone = txtTelephone.Text & ""
            !Address = txtAddress.Text & ""
            !auditid = User
            !audittime = Format(Now, "DD/MM/YYYY hh:mm:ss")
           
            .Update
    End With
    End If
End If
action = ""
fraBank.Enabled = False
Exit Sub
errFix:
    MsgBox Err.Description, vbOKOnly, "Bank Setup"
End Sub

Private Sub cmdUpdate_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    cmdupdate_Click
    txtBankCode.SetFocus
End If
Exit Sub
errFix:
    MsgBox Err.Description, vbOKOnly, "Bank Setup"
End Sub

Private Sub Form_Activate()
On Error GoTo errFix
 load_records
    Exit Sub
errFix:
'   MsgBox Err.Description, vbOKOnly, "Bank Setup"
End Sub

Private Sub Form_Load()
On Error GoTo errFix

With lvwSummary
    .ColumnHeaders.Clear
    .ColumnHeaders.Clear
    .ColumnHeaders.Add 1, "hbankCode", "Bank Code", 2000
    .ColumnHeaders.Add 2, "hbankName", "Bank Name", 2000
    .ColumnHeaders.Add 3, "hbranchName", "Branch Name", 2000
    .ColumnHeaders.Add 4, "hTelephone", "Telephone", 2000
    .ColumnHeaders.Add 5, "hAddress", "Address", 2000
    .ColumnHeaders.Add 6, "hNoOfMem", "No Of Members", 2000
End With
load_records
Set rst = oSaccoMaster.GetRecordSet("select bankcode from d_BANKS order by bankcode")
With rst
    If .RecordCount > 0 Then
        .MoveFirst
        txtBankCode.Text = !BankCode & ""
        lvwSummary.Visible = True
        cmdFirst.Enabled = True
        cmdPrevious.Enabled = True
        cmdNext.Enabled = True
        cmdLast.Enabled = True
        load_records
    End If
End With
Dim ctrl As Control
For Each ctrl In Me
        If TypeOf ctrl Is TextBox Then
            ctrl.Locked = True
        End If
        
        If TypeOf ctrl Is ComboBox Then
            ctrl.Locked = True
        End If
        
        
    Next ctrl
    Set Rst4 = oSaccoMaster.GetRecordSet("select groupid from users where userid= '" & User & "'")
    If Not Rst4.EOF Then
        Set Rst5 = oSaccoMaster.GetRecordSet("select * from usergrps where groupid= '" & Rst4!groupid & "'")
        If Not Rst5.EOF Then
            valToEncrOrDecr = Rst5!banksetup & vbNullString
            EncryptPassword
            If EncryptPass = "View" Then
                disablemodifying = True
                cmdAdd.Enabled = False
                cmdUpdate.Enabled = False
                cmdCancel.Enabled = False
                cmdDelete.Enabled = False
                cmdEdit.Enabled = False
               
            ElseIf EncryptPass = "Mod" Then
                disablemodifying = False
        
            End If
        End If
    Else
       
     
    End If
    
    Exit Sub
errFix:
    MsgBox Err.Description, vbOKOnly, "Bank Setup"
End Sub
Private Sub load_Summary()
On Error GoTo errFix
With lvwSummary
        .ListItems.Clear
    Set rst = oSaccoMaster.GetRecordSet("select bankcode,bankname,branchname,telephone,address from d_BANKS order by bankcode")
    With rst
            Do While Not .EOF
                Set li = lvwSummary.ListItems.Add(, , rst!BankCode & "")
                li.SubItems(1) = rst!BankName & ""
                li.SubItems(2) = rst!BranchName & ""
                li.SubItems(3) = rst!Telephone & ""
                li.SubItems(4) = rst!Address & ""
'                Set rst2 = oSaccoMaster.GetRecordSet("select count(memberno) as memCount from members where bankcode= '" & rst!BankCode & "'")
'                If rst2!memCount > 0 Then
'                    li.SubItems(5) = rst2!memCount
'                Else
'                    li.SubItems(5) = "0"
'                End If
                rst.MoveNext
            Loop
    End With
    End With
    Exit Sub
errFix:
    MsgBox Err.Description, vbOKOnly, "Bank Setup"
End Sub
Public Sub load_records()
On Error GoTo errFix
cmdCancel.Enabled = False
cmdUpdate.Enabled = False
If disablemodifying = False Then
cmdAdd.Enabled = True
cmdEdit.Enabled = True
End If
fraBank.Enabled = False
Dim ctrl As Control
For Each ctrl In Me
        If TypeOf ctrl Is TextBox Then
            ctrl.Locked = True
        End If
        
        If TypeOf ctrl Is ComboBox Then
            ctrl.Locked = True
        End If
        
        If TypeOf ctrl Is CheckBox Then
      
        End If
    Next ctrl
cmdCancel.Enabled = False
Set rst = oSaccoMaster.GetRecordSet("select branchname,bankname,telephone,address from d_BANKS where bankcode= '" & txtBankCode.Text & "'")
With rst
    If .RecordCount > 0 Then
        txtBankName.Text = !BankName & ""
        txtBranchName.Text = !BranchName & ""
        txtTelephone.Text = !Telephone & ""
        txtAddress.Text = !Address & ""
'        Set rst2 = oSaccoMaster.GetRecordSet("select count(memberno) as countMembers from members where bankcode= '" & txtBankCode.Text & "'")
'        If Not IsNull(rst2!countMembers) Then
'            txtNoOfMembers.Text = rst2!countMembers
'        Else
'            If Not txtBankCode.Text = "" Then
'                txtNoOfMembers.Text = "0"
'            End If
        End If
        'rst2.Close
   ' End If
End With
rst.Close

If (lvwSummary.Visible = True) Then
    With lvwSummary
    .ListItems.Clear
    Set rst = oSaccoMaster.GetRecordSet("select bankcode,bankname, branchname,telephone,address from d_BANKS order by bankcode")
    With rst
            Do While Not .EOF
               ' rst.MoveFirst
                Set li = lvwSummary.ListItems.Add(, , rst!BankCode)
                li.SubItems(1) = rst!BankName & ""
                li.SubItems(2) = rst!BranchName & ""
                li.SubItems(3) = rst!Telephone & ""
                li.SubItems(4) = rst!Address & ""
                
'                Set rst2 = oSaccoMaster.GetRecordSet("select count(memberno) as memberCount from members where bankcode= '" & rst!BankCode & "'")
'                If Not IsNull(rst2!memberCount) Then
'                    li.SubItems(5) = rst2!memberCount
'                End If
                rst.MoveNext
            Loop
    End With
    End With
End If
Exit Sub
errFix:
    MsgBox Err.Description, vbOKOnly, "Bank Setup"
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
    MsgBox Err.Description, vbOKOnly, "Bank Setup"
End Function
Private Function fieldFound(FieldName As String) As Boolean
On Error GoTo errFix
fieldFound = False
For i = 0 To rstRecordsImported.Fields.Count - 1
     If UCase(rstRecordsImported.Fields(i).name) = UCase(FieldName) Then
        fieldFound = True
        Exit Function
    End If
Next i
Exit Function
errFix:
    MsgBox Err.Description, vbOKOnly, "Bank Setup"
End Function
Private Function importField(thisField As String) As Boolean
On Error GoTo errFix
importField = False
With frmImport.lvwImportFields
 For counter = 0 To frmImport.lvwImportFields.ListCount - 1
     If frmImport.lvwImportFields.selected(counter) Then
         If frmImport.lvwImportFields.List(counter) = thisField Then
            importField = True
            Exit Function
         End If
     End If
     Next counter
End With
Exit Function
errFix:
    MsgBox Err.Description, vbOKOnly, "Bank Setup"
End Function
Public Sub onFormLoadOfImportFrm()
On Error GoTo errFix
    Dim mylist As String
   PositionForm frmImport
    
    frmImport.Caption = "Data Import: Banks"
    frmImport.chkImportAllFields.value = vbChecked
    
    With frmImport.lvwImportFields
        .AddItem ("Bank Code")
        .AddItem ("Bank Name")
        .AddItem ("Branch Name")
        .AddItem ("Telephone")
        .AddItem ("Address")
        
        
        mylist = frmImport.lvwImportFields.List(1)
        For counter = 0 To frmImport.lvwImportFields.ListCount - 1
            frmImport.lvwImportFields.selected(counter) = True
        Next counter
        
    End With
    frmImport.Show vbModal
    Exit Sub
errFix:
    MsgBox Err.Description, vbOKOnly, "Bank Setup"
End Sub
Public Sub onImpBtnOfImportFrmClick()
On Error GoTo errFixer
While Not rstRecordsImported.EOF
    If fieldFound("BankCode") And importField("Bank Code") Then
    Set Rst5 = oSaccoMaster.GetRecordSet("select bankcode from d_BANKS where bankcode= '" & rstRecordsImported!BankCode & "'")
    If Not Rst5.EOF Then
        frmImport.txtErrorLog.Text = ""
        If fieldFound("bankcode") Then
       frmImport.txtErrorLog.Text = "Bank with code '" & rstRecordsImported!BankCode & "' already exists"
       Exit Sub
       Else
        frmImport.txtErrorLog.Text = "Bank with same bank code exists already exists"
        Exit Sub
       End If
    End If
    End If
    If fieldFound("BankName") And importField("Bank Name") Then
    While rstRecordsImported!BankName = "" Or rstRecordsImported!Telephone = "" Or rstRecordsImported!Address = ""
        frmImport.txtErrorLog.Text = ""
         If fieldFound("bankcode") Then
        frmImport.txtErrorLog.Text = rstRecordsImported!BankCode & "Each record in your excel sheet must have bank name and phone no."
        Else
            frmImport.txtErrorLog.Text = "Each record in your excel sheet must have bank name and phone no."
        End If
        Exit Sub
    Wend
    End If
    If fieldFound("Brachname") And importField("Branch Name") Then
    While rstRecordsImported!BranchName = ""
        frmImport.txtErrorLog.Text = ""
        If fieldFound("bankcode") Then
        frmImport.txtErrorLog.Text = rstRecordsImported!BankCode & "Each record in your excel sheet must have branchname"
        Else
            frmImport.txtErrorLog.Text = "Each record in your excel sheet must have branchname"
        End If
        Exit Sub
    Wend
    End If
    
    Set rst = oSaccoMaster.GetRecordSet("select auditid,audittime,bankcode, branchname, bankname,telephone,address from d_BANKS")
     With rst
            .AddNew
            If importField("Bank Code") And fieldFound("BankCode") Then
            !BankCode = rstRecordsImported!BankCode & ""
            End If
            If importField("Bank Name") And fieldFound("BankName") Then
            !BankName = rstRecordsImported!BankName & ""
            End If
            If fieldFound("Telephone") And importField("Telephone") Then
            !Telephone = rstRecordsImported!Telephone & ""
            End If
            If fieldFound("BranchName") And importField("Branch Name") Then
            !BranchName = rstRecordsImported!BranchName & ""
            End If
            If importField("Address") And fieldFound("Address") Then
            !Address = rstRecordsImported!Address & ""
            End If
            !auditid = User
            !audittime = Format(Now, "DD/MM/YYYY hh:mm:ss")
            .Update
    End With
         
        rstRecordsImported.MoveNext
Wend
MsgBox "Importing Complete", vbInformation, "Complete!"
Exit Sub
errFixer:
    frmImport.txtErrorLog.Text = "Unable to import. Probably file selected is not of excel type"
End Sub

Public Sub onOkOfRangeForm()
On Error GoTo errFix
Set rst = oSaccoMaster.GetRecordSet("select bankcode from d_BANKS order by bankcode")
If Not frmRangeSelection.txtFrom.Text = "" Then
    Set Rst4 = oSaccoMaster.GetRecordSet("select bankcode from d_BANKS where bankcode= '" & frmRangeSelection.txtFrom.Text & "'")
    If Not Rst4.EOF Then
        rangeFrom = frmRangeSelection.txtFrom.Text
    Else
        MsgBox "Enter a bank code.", vbInformation, "Non-existent No"
        frmRangeSelection.txtFrom.Text = ""
        frmRangeSelection.txtFrom.SetFocus
        Exit Sub
    End If
Else
    If Not rst.EOF Then
        rst.MoveFirst
        rangeFrom = rst!BankCode
    End If
End If
If Not frmRangeSelection.txtTo.Text = "" Then
    Set Rst4 = oSaccoMaster.GetRecordSet("select bankcode from d_BANKS where bankcode= '" & frmRangeSelection.txtTo.Text & "'")
    If Not Rst4.EOF Then
        rangeTo = frmRangeSelection.txtTo.Text
    Else
        MsgBox "Enter an existing no.", vbInformation, "Non-existent No"
        frmRangeSelection.txtTo.Text = ""
        frmRangeSelection.txtTo.SetFocus
        Exit Sub
    End If
Else
    
    If Not rst.EOF Then
        rst.MoveLast
        rangeTo = rst!BankCode
    End If
End If
If rangeFrom > rangeTo Then
    MsgBox "'To Bank Code' should be greater than or equal to 'From Bank Code'", vbInformation, "Check Range"
    frmRangeSelection.txtTo.SetFocus
    Exit Sub
End If
Set rst = oSaccoMaster.GetRecordSet("select bankcode from d_BANKS where bankcode= '" & rangeTo & "'")
Set rst1 = oSaccoMaster.GetRecordSet("select bankcode from d_BANKS where bankcode= '" & rangeFrom & "'")
If rst.RecordCount = 0 Then
    MsgBox "Enter an existing bank code for 'From' field", vbInformation, "Bank Code"
End If
If rst1.RecordCount = 0 Then
    MsgBox "Enter an existing bank  code for 'To' field", vbInformation, "Bank Code"
End If
If rst.RecordCount > 0 And rst1.RecordCount > 0 Then
     frmRangeSelection.Visible = False
        Select Case reportType
        Case "Bank List Report"
            Set A = New CRAXDRT.Application
            Set r = A.OpenReport(App.Path & "\Reports\Bank List Report.rpt")
            r.RecordSelectionFormula = "{banks.bankcode}= '" & rangeFrom & "' to '" & rangeTo & "'"
            r.ReadRecords
            Set Rst5 = oSaccoMaster.GetRecordSet("select companyname from sysparam")
            If Not Rst5.EOF Then
                r.ReportTitle = Rst5!CompanyName & ""
            End If
          
                
                With frmReports
                    .Show vbModal, Me
                End With
            
        Case "Bank Member List Report"
            Set A = New CRAXDRT.Application
            Set r = A.OpenReport(App.Path & "\Reports\Bank Memberno Report.rpt")
            r.RecordSelectionFormula = "{banks.bankcode}= '" & rangeFrom & "' to '" & rangeTo & "'"
            r.ReadRecords
            Set Rst5 = oSaccoMaster.GetRecordSet("select companyname from sysparam")
            If Not Rst5.EOF Then
                r.ReportTitle = Rst5!CompanyName & ""
            End If
            
        Case "Bank Staff List Report"
            Set A = New CRAXDRT.Application
            Set r = A.OpenReport(App.Path & "\Reports\Bank Memberno Report.rpt")
            r.RecordSelectionFormula = "{banks.bankcode}= '" & rangeFrom & "' to '" & rangeTo & "'"
            r.ReadRecords
            Set Rst5 = oSaccoMaster.GetRecordSet("select companyname from sysparam")
            If Not Rst5.EOF Then
                r.ReportTitle = Rst5!CompanyName & ""
            End If
          
    End Select
End If
Exit Sub
errFix:
    MsgBox Err.Description, vbOKOnly, "Bank Setup"
End Sub


Public Sub searchSelect()
On Error GoTo errFix
    txtBankCode.Text = sel
    load_records
   
    Exit Sub
errFix:
    MsgBox Err.Description, vbOKOnly, "Bank Setup"
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo errFix
action = "" 'TO cancel edit or add mode
Exit Sub
errFix:
    MsgBox Err.Description, vbOKOnly, "Bank Setup"
End Sub




Private Sub txtAddress_Change()
On Error GoTo errFix
txtAddress.Text = UCase(txtAddress.Text)
txtAddress.SelStart = Len(txtAddress.Text)
Exit Sub
errFix:
    MsgBox Err.Description, vbOKOnly, "Bank Setup"
End Sub

Private Sub txtAddress_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If Len(txtAddress.Text) > 29 Then
      Beep
      MsgBox "Can't enter more than 30 characters", vbExclamation
      KeyAscii = 8
End If
If KeyAscii = 13 Then
    txtTelephone.SetFocus
End If
Exit Sub
errFix:
    MsgBox Err.Description, vbOKOnly, "Bank Setup"
End Sub

Private Sub txtBankCode_Change()
On Error GoTo errFix
txtBankCode.Text = UCase(txtBankCode.Text)
txtBankCode.SelStart = Len(txtBankCode.Text)
Exit Sub
errFix:
    MsgBox Err.number, vbOKOnly, "Bank Setup"
End Sub



Private Sub txtBankCode_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If Len(txtBankCode.Text) > 9 Then
      Beep
      MsgBox "Can't enter more than 10 characters", vbExclamation
      KeyAscii = 8
End If
If KeyAscii = 13 Then
    Set Rst5 = oSaccoMaster.GetRecordSet("select bankcode from banks where bankcode= '" & txtBankCode.Text & "'")
    If Not Rst5.EOF Then
        MsgBox "Bank with same code already exists", vbInformation, "Bank Code"
        txtBankCode.Text = ""
        txtBankCode.SetFocus
    Else
        txtBankName.SetFocus
    End If
End If
Exit Sub
errFix:
    MsgBox Err.Description, vbOKOnly, "Bank Setup"
End Sub

Private Sub txtBankCode_LostFocus()
On Error GoTo errFix
    If action = "addingRecords" Then
        Set Rst5 = oSaccoMaster.GetRecordSet("select bankcode from d_BANKS where bankcode= '" & txtBankCode.Text & "'")
        If Not Rst5.EOF Then
            MsgBox "Bank with same code already exists", vbInformation, "Bank Code"
            txtBankCode.Text = ""
            txtBankCode.SetFocus
        
        End If
    End If
'    Set rst2 = oSaccoMaster.GetRecordSet("select memberno from members where bankcode= '" & txtBankCode.Text & "'")
'    If rst2.RecordCount > 0 Then
'        txtNoOfMembers.Text = rst2.RecordCount
'    Else
'        txtNoOfMembers.Text = "0"
'    End If
    'rst2.Close
    Exit Sub
errFix:
    MsgBox Err.Description, vbOKOnly, "Bank Setup"
End Sub

Private Sub txtBankName_Change()
On Error GoTo errFix
txtBankName.Text = UCase(txtBankName.Text)
txtBankName.SelStart = Len(txtBankName.Text)
Exit Sub
errFix:
    MsgBox Err.Description, vbOKOnly, "Bank Setup"
End Sub

Private Sub txtBankName_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If Len(txtBankName.Text) > 49 Then
      Beep
      MsgBox "Can't enter more than 50 characters", vbExclamation
      KeyAscii = 8
End If
If KeyAscii = 13 Then
    txtBranchName.SetFocus
End If
Exit Sub
errFix:
    MsgBox Err.Description, vbOKOnly, "Bank Setup"
End Sub

Private Sub txtBranchName_Change()
On Error GoTo errFix
txtBranchName.Text = UCase(txtBranchName.Text)
txtBranchName.SelStart = Len(txtBranchName.Text)
Exit Sub
errFix:
    MsgBox Err.Description, vbOKOnly, "Bank Setup"
End Sub

Private Sub txtBranchName_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If Len(txtBranchName.Text) > 49 Then
      Beep
      MsgBox "Can't enter more than 50 characters", vbExclamation
      KeyAscii = 8
End If
If KeyAscii = 13 Then
    txtAddress.SetFocus
End If
Exit Sub
errFix:
    MsgBox Err.Description, vbOKOnly, "Bank Setup"
End Sub

Private Sub txtTelephone_Change()
On Error GoTo errFix
txtTelephone.Text = UCase(txtTelephone.Text)
txtTelephone.SelStart = Len(txtTelephone.Text)
Exit Sub
errFix:
    MsgBox Err.Description, vbOKOnly, "Bank Setup"
End Sub

Private Sub txtTelephone_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    cmdUpdate.SetFocus
End If
Select Case KeyAscii
    Case Asc("0") To Asc("9")
    Case Asc("-")
    Case Asc("+")
    Case Asc(")")
    Case Asc(" ")
    Case Asc(",")
    Case Asc("(")
    Case Is = 8
    Case Else
        Beep
        KeyAscii = 0
  End Select
If Len(txtTelephone.Text) > 14 Then
      Beep
      MsgBox "Can't enter more than 15 characters", vbExclamation
      KeyAscii = 8
End If

  Select Case KeyAscii
    Case Asc("0") To Asc("9")
    Case Asc("-")
    Case Asc("+")
    Case Asc(")")
    Case Asc(" ")
    Case Asc(",")
    Case Asc("(")
    Case Is = 8
    Case Else
        Beep
        KeyAscii = 0
  End Select
  Exit Sub
errFix:
    MsgBox Err.Description, vbOKOnly, "Bank Setup"
End Sub
