VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmLoanTypes 
   Caption         =   "Deductions code"
   ClientHeight    =   6555
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7710
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLoanTypes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   7710
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lvwContraAcc 
      Height          =   1215
      Left            =   3600
      TabIndex        =   26
      Top             =   3720
      Visible         =   0   'False
      Width           =   3990
      _ExtentX        =   7038
      _ExtentY        =   2143
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "AccNo"
         Object.Width           =   18
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Acc Name"
         Object.Width           =   10583
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "GL Integration"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   18
      Top             =   2400
      Width           =   7455
      Begin MSComctlLib.ListView lvwLoanAcc 
         Height          =   945
         Left            =   3480
         TabIndex        =   25
         Top             =   360
         Visible         =   0   'False
         Width           =   3990
         _ExtentX        =   7038
         _ExtentY        =   1667
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "AccNo"
            Object.Width           =   18
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Acc Name"
            Object.Width           =   10583
         EndProperty
      End
      Begin VB.TextBox txtContraAccName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3240
         TabIndex        =   24
         Top             =   1200
         Width           =   3975
      End
      Begin VB.TextBox txtLoanAccName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3270
         TabIndex        =   21
         Top             =   225
         Width           =   3975
      End
      Begin VB.TextBox txtLoanAcc 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtContraAcc 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Contra Account"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   23
         Top             =   1200
         Width           =   1125
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Loan Account"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   22
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.PictureBox CrystalReport1 
      Height          =   480
      Left            =   5175
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   17
      Top             =   5940
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Height          =   495
      Left            =   4680
      Picture         =   "frmLoanTypes.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Cancel Process"
      Top             =   6000
      Width           =   495
   End
   Begin VB.CommandButton cmdUpdate 
      Height          =   495
      Left            =   4200
      Picture         =   "frmLoanTypes.frx":040C
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Save Record"
      Top             =   6000
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
      Picture         =   "frmLoanTypes.frx":04FE
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Delete Record"
      Top             =   6000
      Width           =   495
   End
   Begin VB.CommandButton cmdEdit 
      Height          =   495
      Left            =   3240
      Picture         =   "frmLoanTypes.frx":05F0
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Edit Record"
      Top             =   6000
      Width           =   495
   End
   Begin VB.CommandButton cmdNext 
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
      Picture         =   "frmLoanTypes.frx":06F2
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Move to the Next"
      Top             =   6000
      Width           =   495
   End
   Begin VB.CommandButton cmdLast 
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
      Left            =   1605
      Picture         =   "frmLoanTypes.frx":0A34
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Move to Last record"
      Top             =   6000
      Width           =   495
   End
   Begin VB.CommandButton cmdPrev 
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
      Picture         =   "frmLoanTypes.frx":0D76
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Move to the Previous record"
      Top             =   6000
      Width           =   495
   End
   Begin MSComctlLib.ListView lvwLoanApplications 
      Height          =   1665
      Left            =   90
      TabIndex        =   15
      Top             =   4200
      Width           =   7530
      _ExtentX        =   13282
      _ExtentY        =   2937
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
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.CommandButton cmdAdd 
      Height          =   495
      Left            =   2760
      Picture         =   "frmLoanTypes.frx":10B8
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Add New record"
      Top             =   6000
      Width           =   495
   End
   Begin VB.CommandButton cmdFirst 
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
      Picture         =   "frmLoanTypes.frx":11AA
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Move to the Last record"
      Top             =   6000
      Width           =   495
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   510
      Left            =   6360
      TabIndex        =   11
      Top             =   6000
      Width           =   1230
   End
   Begin VB.Frame frameLoanTypes 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   120
      TabIndex        =   12
      Top             =   600
      Width           =   7470
      Begin VB.TextBox txtLoanCode 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtLoanType 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   1605
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   480
         Width           =   4215
      End
      Begin VB.Label Label7 
         Caption         =   "&Loan Code"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Loan &Type"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1605
         TabIndex        =   13
         Top             =   240
         Width           =   1575
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2160
      Top             =   5865
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLoanTypes.frx":14EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLoanTypes.frx":15FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLoanTypes.frx":1710
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblLoanApp 
      Caption         =   "Loan Applications"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   3960
      Width           =   2085
   End
End
Attribute VB_Name = "frmLoanTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim disablemodifying As Boolean
Private Sub cmdAdd_Click()
On Error GoTo errFix
action = "addingRecords"
Toolbar1.Buttons("bSearch").Enabled = False
Toolbar1.Buttons("bView").Enabled = False
Toolbar1.Buttons("bPrint").Enabled = False
txtLoanCode.Locked = False
txtLoanType.Locked = False
txtMaxAmmount.Locked = False
txtInterestRate.Locked = False
txtRepayPeriod.Locked = False
frameLoanTypes.Visible = True
frameLoanTypes.Enabled = True
cmdEdit.Enabled = False
cmdDelete.Enabled = False
Set Rst3 = oSaccoMaster.GetRecordSet("select * from loantype where loancode= '" & txtLoanCode.Text & "'")
If Rst3.RecordCount > 0 Then
    MyBookMark = Rst3.Bookmark
End If
cmdCancel.Enabled = True
txtLoanCode.Locked = False
txtLoanType.Locked = False
txtMaxAmmount.Locked = False
txtInterestRate.Locked = False
txtRepayPeriod.Locked = False

Set Rst3 = oSaccoMaster.GetRecordSet("select * from loantype where loancode= '" & txtLoanCode.Text & "'")
If Rst3.RecordCount > 0 Then
    MyBookMark = Rst3.Bookmark
End If
    cmdEdit.Enabled = False
    lvwSummary.Visible = False
    lvwLoanApplications.Visible = False
    lblLoanApp.Visible = False
    txtLoanCode.Text = ""
    txtLoanType.Text = ""
    txtRepayPeriod.Text = ""
    txtMaxAmmount.Text = ""
    txtInterestRate.Text = ""
    txtNumOfLoans.Text = ""
    cmdAdd.Enabled = False
    cmdUpdate.Enabled = True
    txtNumOfLoans.Text = "0"
     Exit Sub
errFix:
    MsgBox Err.Description, vbOKOnly, "Loan Types"
End Sub

Private Sub cmdAdd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtLoanCode.SetFocus
End If
End Sub

Private Sub cmdCancel_Click()
On Error GoTo errFix
Toolbar1.Buttons("bSearch").Enabled = True
Toolbar1.Buttons("bView").Enabled = True
Toolbar1.Buttons("bPrint").Enabled = True
txtLoanCode.Locked = True
txtLoanType.Locked = True
txtMaxAmmount.Locked = True
frameLoanTypes.Enabled = False
txtInterestRate.Locked = True
txtRepayPeriod.Locked = True
txtNumOfLoans.Locked = True
cmdFirst.Enabled = True
cmdPrev.Enabled = True
cmdNext.Enabled = True
cmdLast.Enabled = True
cmdCancel.Enabled = False
cmdEdit.Enabled = True
cmdAdd.Enabled = True
cmdDelete.Enabled = True
Set Rst3 = oSaccoMaster.GetRecordSet("select * from loantype")
If Rst3.RecordCount > 0 Then
    Rst3.Bookmark = MyBookMark
End If
txtLoanCode.Text = Rst3!Loancode & ""
load_records
 Exit Sub
errFix:
    MsgBox Err.Description, vbOKOnly, "Loan Types"
End Sub

Private Sub cmdCancel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtLoanCode.SetFocus
End If
End Sub

Private Sub cmdClose_Click()
Unload Me
action = ""
End Sub
Private Sub cmdDelete_Click()
On Error GoTo errFix
If lvwSummary.Visible = False Then
    Set rst = oSaccoMaster.GetRecordSet("select * from loantype where loancode= '" & txtLoanCode.Text & "'")
    If rst.RecordCount > 0 Then
        If MsgBox("Do you want to delete loan type " & txtLoanCode.Text & "  ?", vbOKCancel, "Loan Code") = vbOK Then
            rst.Delete
            rst.Update
            load_records
        End If
    End If
Else
    Set rst = oSaccoMaster.GetRecordSet("select * from loantype where loancode= '" & sel & "'")
    If rst.RecordCount > 0 Then
        If MsgBox("Do you want to delete loan type " & txtLoanCode.Text & "  ?", vbOKCancel, "Loan Code") = vbOK Then
            rst.Delete
            rst.Update
            load_Summary
        End If
        
    End If
End If
 Exit Sub
errFix:
    MsgBox Err.Description, vbOKOnly, "Loan Types"
End Sub

Private Sub cmdDelete_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdNext.SetFocus
End If
End Sub

Private Sub cmdEdit_Click()
On Error GoTo errFix
action = "editingRecords"
Toolbar1.Buttons("bSearch").Enabled = False
Toolbar1.Buttons("bView").Enabled = False
Toolbar1.Buttons("bPrint").Enabled = False
frameLoanTypes.Enabled = True
cmdAdd.Enabled = False
cmdDelete.Enabled = False
cmdCancel.Enabled = True
cmdUpdate.Enabled = True
cmdEdit.Enabled = False
    If lvwSummary.Visible = True Then
        Set Rst3 = oSaccoMaster.GetRecordSet("select * from loantype where loancode= '" & sel & "'")
    Else
       Set Rst3 = oSaccoMaster.GetRecordSet("select * from loantype where loancode= '" & txtLoanCode.Text & "'")
    End If
    If Rst3.RecordCount > 0 Then
        MyBookMark = Rst3.Bookmark
    End If
    If lvwSummary.Visible Then
       txtLoanCode.Text = sel
       load_records
    End If
    lvwSummary.Visible = False
    lvwLoanApplications.Visible = True
    frameLoanTypes.Visible = True
    lblLoanApp.Visible = True
    txtLoanType.Locked = False
    txtMaxAmmount.Locked = False
    txtInterestRate.Locked = False
    txtRepayPeriod.Locked = False
    Editing = True
     Exit Sub
errFix:
    MsgBox Err.Description, vbOKOnly, "Loan Types"
End Sub

Private Sub cmdEdit_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtLoanCode.SetFocus
End If
End Sub

Private Sub cmdFirst_Click()
    On Error GoTo errFix
    Toolbar1.Buttons("bSearch").Enabled = True
    Toolbar1.Buttons("bView").Enabled = True
    Toolbar1.Buttons("bPrint").Enabled = True
    lvwSummary.Visible = False
    lvwLoanApplications.Visible = True
    lblLoanApp.Visible = True
    
    Set rst = oSaccoMaster.GetRecordSet("select loancode from loantype order by loancode")
    
    If rst.RecordCount > 0 Then
        rst.MoveFirst
        txtLoanCode.Text = rst!Loancode & ""
        If lvwSummary.Visible = False Then
        LoadLoanType
        cmdFirst.Enabled = False
        cmdPrev.Enabled = False
        cmdNext.Enabled = True
        cmdLast.Enabled = True
        End If
    End If
    rst.Close
    action = ""
    Exit Sub
errFix:
    MsgBox Err.Description, vbOKOnly, "Loan Types"
End Sub

Private Sub cmdLast_Click()
On Error GoTo errFix
Toolbar1.Buttons("bSearch").Enabled = True
Toolbar1.Buttons("bView").Enabled = True
Toolbar1.Buttons("bPrint").Enabled = True
lvwSummary.Visible = False
lvwLoanApplications.Visible = True
lblLoanApp.Visible = True

Set rst = oSaccoMaster.GetRecordSet("select loancode from loantype order by loancode")
With rst
    .MoveLast
    txtLoanCode.Text = rst!Loancode & ""
    If lvwSummary.Visible = False Then
    LoadLoanType
    cmdFirst.Enabled = True
    cmdPrev.Enabled = True
    cmdNext.Enabled = False
    cmdLast.Enabled = False
    End If
End With
action = ""
 Exit Sub
errFix:
    MsgBox Err.Description, vbOKOnly, "Loan Types"
End Sub
Private Sub cmdNext_Click()
On Error GoTo errFix
Toolbar1.Buttons("bSearch").Enabled = True
Toolbar1.Buttons("bView").Enabled = True
Toolbar1.Buttons("bPrint").Enabled = True
lvwSummary.Visible = False
lvwLoanApplications.Visible = True
lblLoanApp.Visible = True
Set rst = oSaccoMaster.GetRecordSet("select loancode from loantype order by loancode")
If cmdUpdate.Enabled = True Then
    If rst.RecordCount > 0 Then
        rst.Bookmark = MyBookMark
        txtLoanCode.Text = rst!Loancode & ""
    End If

End If
With rst
If rst.RecordCount > 0 Then
    rst.MoveFirst
    rst.Find ("loancode= '" & txtLoanCode.Text & "'")
    If Not .EOF Then
     rst.MoveNext
        If .EOF Then
           .MoveLast
           cmdFirst.Enabled = True
           cmdPrev.Enabled = True
           cmdNext.Enabled = False
           cmdLast.Enabled = False
        Else
           cmdFirst.Enabled = True
           cmdPrev.Enabled = True
           cmdNext.Enabled = True
           cmdLast.Enabled = True
        End If
     txtLoanCode.Text = rst!Loancode & ""
     If lvwSummary.Visible = False Then
        LoadLoanType
     End If
    End If
 End If
 End With
 rst.Close
 action = ""
  Exit Sub
errFix:
    MsgBox Err.Description, vbOKOnly, "Loan Types"
End Sub

Private Sub cmdPrev_Click()
On Error GoTo errFix
Toolbar1.Buttons("bSearch").Enabled = True
Toolbar1.Buttons("bView").Enabled = True
Toolbar1.Buttons("bPrint").Enabled = True
lvwSummary.Visible = False
lvwLoanApplications.Visible = True
lblLoanApp.Visible = True
Set rst = oSaccoMaster.GetRecordSet("select loancode from loantype order by loancode")

If cmdUpdate.Enabled = True Then
    If rst.RecordCount > 0 Then
        rst.Bookmark = MyBookMark
        txtLoanCode.Text = rst!Loancode & ""
    End If
End If
With rst
If rst.RecordCount > 0 Then
    rst.MoveFirst
    rst.Find ("loancode= '" & txtLoanCode.Text & "'")
    If Not .EOF Then
        .MovePrevious
        If .BOF Then
            .MoveFirst
            cmdFirst.Enabled = False
            cmdPrev.Enabled = False
            cmdNext.Enabled = True
            cmdLast.Enabled = True
        Else
            cmdFirst.Enabled = True
            cmdPrev.Enabled = True
            cmdNext.Enabled = True
            cmdLast.Enabled = True
        End If
    txtLoanCode.Text = rst!Loancode & ""
    If lvwSummary.Visible = False Then
    LoadLoanType
    End If
    End If
End If
End With
rst.Close
action = ""
 Exit Sub
errFix:
    MsgBox Err.Description, vbOKOnly, "Loan Types"
End Sub
Private Sub CommandButton5_Click()
Unload Me
End Sub

Private Sub cmdPrevious_Click()

End Sub

Private Sub cmdUpdate_Click()
On Error GoTo errFix
If action = "addingRecords" Then
    Set Rst5 = oSaccoMaster.GetRecordSet("select loancode from loantype where loancode= '" & txtLoanCode.Text & "'")
    If Not Rst5.EOF Then
        MsgBox "Loan type with same code exists", vbInformation, "Loan Code"
        txtLoanCode.Text = ""
        txtLoanCode.SetFocus
        Exit Sub
    End If
    While txtLoanCode.Text = "" Or txtLoanType.Text = "" Or txtMaxAmmount.Text = "" Or txtInterestRate.Text = "" Or txtRepayPeriod.Text = ""
        MsgBox "Enter full information", , "Insufficient Information"
        Exit Sub
    Wend
    Set rst = oSaccoMaster.GetRecordSet("select * from loantype")
    rst.AddNew
    cmdCancel.Enabled = False
    cmdUpdate.Enabled = False
    cmdDelete.Enabled = True
    txtLoanCode.Locked = True
    txtLoanType.Locked = True
    txtMaxAmmount.Locked = True
    txtInterestRate.Locked = True
    txtRepayPeriod.Locked = True
    txtNumOfLoans.Locked = True
    cmdEdit.Enabled = True
    cmdAdd.Enabled = True
    rst!Loancode = txtLoanCode.Text & ""
    rst!LoanType = txtLoanType.Text & ""
    If ChkGuarantors.Value = vbChecked Then
        rst!Guarantor = "Yes"
    Else
        rst!Guarantor = "No"
    End If
    rst!auditid = User
    rst!audittime = Get_Server_Date
    rst!repayperiod = txtRepayPeriod.Text & ""
    rst!Priority = txtpriority.Text & ""
    rst!MaxAmount = IIf(Trim(txtMaxAmmount) <> "", txtMaxAmmount, 0)
    rst!interest = txtInterestRate.Text & ""
    rst!MaxLoans = IIf(txtMaxLoans = "", 0, txtMaxLoans)
    rst!audittime = Now
    Toolbar1.Buttons("bSearch").Enabled = True
    Toolbar1.Buttons("bView").Enabled = True
    Toolbar1.Buttons("bPrint").Enabled = True
    rst.Update
ElseIf action = "editingRecords" Then
    While txtLoanCode.Text = "" Or txtLoanType.Text = "" Or txtMaxAmmount.Text = "" Or txtInterestRate.Text = "" Or txtRepayPeriod.Text = ""
        MsgBox "Enter full information", , "Insufficient Information"
        Exit Sub
    Wend
    cmdFirst.Enabled = True
    cmdPrev.Enabled = True
    cmdNext.Enabled = True
    cmdLast.Enabled = True
    cmdEdit.Enabled = True
    cmdUpdate.Enabled = False
    cmdAdd.Enabled = True
    cmdCancel.Enabled = False
    cmdDelete.Enabled = True
    txtLoanType.Locked = True
    txtMaxAmmount.Locked = True
    txtInterestRate.Locked = True
    txtRepayPeriod.Locked = True
    txtNumOfLoans.Locked = True
    Set rst = oSaccoMaster.GetRecordSet("Select * from loantype where loancode= '" & Rst3!Loancode & "'")
    rst!Loancode = txtLoanCode.Text & ""
    rst!LoanType = txtLoanType.Text & ""
    rst!repayperiod = txtRepayPeriod.Text & ""
    'rst!MaxLoans = IIf(txtMaxLoans = "", 0, txtMaxLoans)
    If ChkGuarantors.Value = vbChecked Then
        rst!Guarantor = "Yes"
    Else
        rst!Guarantor = "No"
    End If
    rst!auditid = User
    rst!Priority = txtpriority.Text & ""
    rst!audittime = Format(Now, "DD/MM/YYYY hh:mm:ss")
    rst!MaxAmount = txtMaxAmmount.Text & ""
    rst!interest = txtInterestRate.Text & ""
    rst!audittime = Now
    Toolbar1.Buttons("bSearch").Enabled = True
    Toolbar1.Buttons("bView").Enabled = True
    Toolbar1.Buttons("bPrint").Enabled = True
    rst.Update
    End If
        frameLoanTypes.Enabled = False
        action = ""
    Exit Sub
errFix:
    MsgBox Err.Description, vbOKOnly, "Loan Types"
End Sub

Private Sub cmdUpdate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtLoanCode.SetFocus
End If
End Sub

Private Sub Form_Load()
    On Error GoTo errFix
    PositionForm Me
    cmdCancel.Enabled = False
    cmdUpdate.Enabled = False
    lvwSummary.Visible = False
    Toolbar1.Buttons.Add 1, "bSearch", "&Search", 0, 2
    Toolbar1.Buttons.Add 2, "bView", "&View", 5, 3
    Toolbar1.Buttons.Add 3, "bPrint", "&Print", 5, 1
    Toolbar1.Buttons("bView").ButtonMenus.Add 1, "mDetails", "Details"
    Toolbar1.Buttons("bView").ButtonMenus.Add 2, "mSummary", "Summary"
    Toolbar1.Buttons("bPrint").ButtonMenus.Add 1, "mLoanTypes", "loan Types"
    Toolbar1.Buttons("bPrint").ButtonMenus.Add 2, "mLoanApplications", "Loan Applications"
    With lvwLoanApplications
        .ColumnHeaders.Clear
        .ListItems.Clear
        .ColumnHeaders.Add 1, "hLoanCode", "Loan Code", 2000
        .ColumnHeaders.Add 2, "hLoanNo", "Loan No", 2000
        .ColumnHeaders.Add 3, "hMemberNo", "Member No", 2000
        .ColumnHeaders.Add 4, "hApplicant", "Applicant", 3000
        .ColumnHeaders.Add 5, "hLoanAmount", "Loan Amt", 2000, 1
        .ColumnHeaders.Add 6, "hPurpose", "Purpose", 3000
    End With
    Set rst = oSaccoMaster.GetRecordSet("select top 1 * from loantype")
    If Not rst.EOF Then
        txtLoanCode.Text = rst!Loancode
        LoadLoanType
    End If
    Exit Sub
errFix:
   MsgBox Err.Description, vbOKOnly, "Loan Types"
    Exit Sub
End Sub

Public Sub LoadLoanType()
'    On Error GoTo errFix
Dim yye As Integer, Account As Acc_Details
yye = 0
    lvwLoanApplications.ListItems.Clear
    Set rst = oSaccoMaster.GetRecordSet("select * from LOANTYPE" _
    & " where LoanCode='" & txtLoanCode.Text & "'")
    With rst
        If Not .EOF Then
            txtLoanCode.Text = !Loancode
            txtLoanType.Text = !LoanType
            txtInterestRate.Text = !interest
            txtMaxAmmount.Text = Format(!MaxAmount, Cfmt)
            txtRepayPeriod.Text = !repayperiod
            txtpriority.Text = !Priority
            
            Editing = False
            txtContraAcc = !ContraAcc
            txtContraAccName = Get_Acc_Details(txtContraAcc, ErrorMessage).AccName
            
            txtInterestAcc = !InterestAcc
            txtInterestAccName = Get_Acc_Details(txtInterestAcc, ErrorMessage).AccName
            txtLoanAcc = !LoanAcc
            txtLoanAccName = Get_Acc_Details(txtLoanAcc, ErrorMessage).AccName
            'txtMaxLoans = IIf(IsNull(!MaxLoans), 0, !MaxLoans)
            
            If !Guarantor = "Yes" Then
                ChkGuarantors.Value = 1
            Else
                ChkGuarantors.Value = 0
            End If
            Set rsLoanGuar = oSaccoMaster.GetRecordSet("select LB.LoanNo,C.Amount," _
            & "LB.MemberNo,LB.LoanCode,M.SurName,M.OtherNames from (LOANBAL LB inner join " _
            & "CHEQUES C on LB.LoanNo=C.LoanNo)inner join MEMBERS M on LB.MemberNo" _
            & "=M.MemberNo where LB.LoanCode='" & txtLoanCode & "'")
            With rsLoanGuar
                If Not .EOF Then
                    While Not .EOF
                        Set li = lvwLoanApplications.ListItems.Add(, , !Loancode)
                        li.SubItems(1) = !LoanNo
                        li.SubItems(2) = !memberno
                        li.SubItems(3) = !othernames & " " & !surname
                        li.SubItems(4) = Format(!amount, Cfmt)
                        yye = yye + 1
                        .MoveNext
                    Wend
                End If
            End With
        End If
    End With
    txtNumOfLoans = yye
    Exit Sub
errFix:
    MsgBox Err.Description, vbOKOnly, "Loan Types"
    Exit Sub
End Sub
Private Sub load_Summary()
On Error GoTo errFix
cmdCancel.Enabled = False
cmdUpdate.Enabled = False
lvwLoanApplications.Visible = True
frameLoanTypes.Visible = False
lblLoanApp.Visible = False
With lvwSummary
    .ColumnHeaders.Clear
    .ListItems.Clear
    .ColumnHeaders.Add 1, "hLoanCode", "Loan Code", 2000
    .ColumnHeaders.Add 2, "hLoanType", "Loan Type", 3000
    .ColumnHeaders.Add 3, "hRepayPeriod", "Repay Period", 2000
    .ColumnHeaders.Add 4, "hInterestRate", "Interest Rate", 2000
    .ColumnHeaders.Add 5, "hMaxAmt", "Maximum Amount", 2000
    .ColumnHeaders.Add 6, "hNumLoans", "Num Loans", 2000
    Set rst = oSaccoMaster.GetRecordSet("SELECT LOANTYPE.LoanCode, LOANTYPE.LoanType, LOANTYPE.RepayPeriod, LOANTYPE.MaxAmount, LOANTYPE.Interest, Count(LOANS.LoanNo) AS CountOfLoanNo FROM LOANTYPE LEFT JOIN LOANS ON LOANTYPE.LoanCode = LOANS.LoanCode GROUP BY LOANTYPE.LoanCode, LOANTYPE.LoanType, LOANTYPE.RepayPeriod, LOANTYPE.MaxAmount, LOANTYPE.Interest order by LOANTYPE.loancode")
    If rst.RecordCount > 0 Then
    rst.MoveFirst
    While Not rst.EOF
        Set li = .ListItems.Add(, , rst!Loancode & "")
        li.SubItems(1) = rst!LoanType & ""
        li.SubItems(2) = rst!repayperiod & ""
        li.SubItems(3) = rst!interest & ""
        li.SubItems(4) = Format(rst!MaxAmount & "", Cfmt)
        If Not IsNull(rst!CountOfLoanNo) Then
        li.SubItems(5) = rst!CountOfLoanNo
        Else
        li.SubItems(5) = "0"
        End If
        rst.MoveNext
    Wend
    End If
End With
lvwSummary.Visible = True
 Exit Sub
errFix:
    MsgBox Err.Description, vbOKOnly, "Loan Types"
End Sub
Public Sub load_records()
On Error GoTo errFix
cmdUpdate.Enabled = False
If disablemodifying = False Then
cmdAdd.Enabled = True
cmdEdit.Enabled = True
cmdDelete.Enabled = True
End If
cmdCancel.Enabled = False
lvwSummary.Visible = False
lvwLoanApplications.Visible = True
frameLoanTypes.Visible = True
frameLoanTypes.Enabled = False
lblLoanApp.Visible = True
Set rst = oSaccoMaster.GetRecordSet("SELECT LOANTYPE.LoanCode, LOANTYPE.LoanType,loantype.guarantor, LOANTYPE.Interest, LOANTYPE.MaxAmount, LOANTYPE.RepayPeriod, LOANS.LoanAmt, Count(LOANS.LoanNo) AS CountOfLoanNo FROM LOANTYPE LEFT JOIN LOANS ON LOANTYPE.LoanCode = LOANS.LoanCode WHERE (((LOANTYPE.LoanCode)= '" & txtLoanCode.Text & "')) GROUP BY LOANTYPE.LoanCode, LOANTYPE.LoanType, LOANTYPE.Interest,loantype.guarantor, LOANTYPE.MaxAmount, LOANTYPE.RepayPeriod, LOANS.LoanAmt order by LOANTYPE.loancode")
With rst
    txtLoanCode.Text = !Loancode & ""
    txtLoanType.Text = !LoanType & ""
    txtRepayPeriod.Text = !repayperiod & ""
    txtMaxAmmount.Text = Format(!MaxAmount & "", Cfmt)
    txtInterestRate.Text = !interest & ""
    If !Guarantor = "Yes" Then
        ChkGuarantors.Value = vbChecked
    ElseIf !Guarantor = "No" Then
        ChkGuarantors.Value = vbUnchecked
    End If
    If Not IsNull(rst!CountOfLoanNo) Then
        txtNumOfLoans.Text = rst!CountOfLoanNo
    Else
        txtNumOfLoans.Text = "0"
    End If
End With

    
'    Set Rst1 = osaccomaster.GetRecordSet("SELECT LOANS.MemberNo, LOANS.LoanCode, LOANS.LoanAmt, LOANS.LoanNo, MEMBERS.Surname, MEMBERS.OtherNames, LOANS.Purpose FROM LOANS LEFT JOIN MEMBERS ON LOANS.MemberNo = MEMBERS.MemberNo where LOANS.loancode= '" & txtLoanCode.Text & "' GROUP BY LOANS.MemberNo, LOANS.LoanCode, LOANS.LoanAmt, LOANS.LoanNo, MEMBERS.Surname, MEMBERS.OtherNames, LOANS.Purpose order by LOANS.loancode")
'
'    Do While Not Rst1.EOF
'         Set Li = .ListItems.Add(, , Rst!Loancode & "")
'        Li.SubItems(1) = Rst1!Loanno & ""
'        Li.SubItems(2) = Rst1!MemberNo & ""
'        Li.SubItems(3) = Rst1!surname & " " & Rst1!othernames
'        Li.SubItems(4) = Format(Rst1!LoanAmt & "", Cfmt)
'        Li.SubItems(5) = Rst1!purpose & ""
'        Rst1.MoveNext
'    Loop
'End With
     Exit Sub
errFix:
    MsgBox Err.Description, vbOKOnly, "Loan Types"
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
    MsgBox Err.Description, vbOKOnly, "Loan Types"
End Function
Private Sub Form_Unload(Cancel As Integer)
action = ""
End Sub

Private Sub lvwContraAcc_DblClick()
On Error GoTo SysError
    If lvwContraAcc.ListItems.Count > 0 Then
        txtContraAcc = lvwContraAcc.SelectedItem
        txtContraAccName = lvwContraAcc.SelectedItem.SubItems(1)
        lvwContraAcc.Visible = False
    End If
    Exit Sub
SysError:
    MsgBox Err.Description, vbInformation, Me.Caption
End Sub

Private Sub lvwInterestAcc_DblClick()
On Error GoTo SysError
    If lvwInterestAcc.ListItems.Count > 0 Then
        txtInterestAcc = lvwInterestAcc.SelectedItem
        txtInterestAccName = lvwInterestAcc.SelectedItem.SubItems(1)
        lvwInterestAcc.Visible = False
    End If
    Exit Sub
SysError:
    MsgBox Err.Description, vbInformation, Me.Caption
End Sub

Private Sub lvwLoanAcc_DblClick()
On Error GoTo SysError
    If lvwLoanAcc.ListItems.Count > 0 Then
        txtLoanAcc = lvwLoanAcc.SelectedItem
        txtLoanAccName = lvwLoanAcc.SelectedItem.SubItems(1)
        lvwLoanAcc.Visible = False
    End If
    Exit Sub
SysError:
    MsgBox Err.Description, vbInformation, Me.Caption
End Sub

Private Sub lvwSummary_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo errFix
    sel = lvwSummary.SelectedItem
    With lvwLoanApplications
    .ColumnHeaders.Clear
    .ListItems.Clear
    .ColumnHeaders.Add 1, "hLoanCode", "Loan Code", 2000
    .ColumnHeaders.Add 2, "hLoanNo", "Loan No", 2000
    .ColumnHeaders.Add 3, "hMemberNo", "Member No", 2000
    .ColumnHeaders.Add 4, "hApplicant", "Applicant", 3000
    .ColumnHeaders.Add 5, "hLoanAmount", "Loan Amt", 2000
    .ColumnHeaders.Add 6, "hPurpose", "Purpose", 3000
    Set rst1 = oSaccoMaster.GetRecordSet("SELECT LOANS.MemberNo, LOANS.LoanCode, LOANS.LoanAmt, LOANS.LoanNo, MEMBERS.Surname, MEMBERS.OtherNames, LOANS.Purpose FROM LOANS LEFT JOIN MEMBERS ON LOANS.MemberNo = MEMBERS.MemberNo where LOANS.loancode= '" & sel & "' GROUP BY LOANS.MemberNo, LOANS.LoanCode, LOANS.LoanAmt, LOANS.LoanNo, MEMBERS.Surname, MEMBERS.OtherNames, LOANS.Purpose order by LOANS.loancode")
    
    Do While Not rst1.EOF
         Set li = .ListItems.Add(, , sel & "")
        li.SubItems(1) = rst1!LoanNo & ""
        li.SubItems(2) = rst1!memberno & ""
        li.SubItems(3) = rst1!surname & " " & rst1!othernames
        li.SubItems(4) = Format(rst1!LoanAmt & "", Cfmt)
        li.SubItems(5) = rst1!purpose & ""
        rst1.MoveNext
    Loop
End With
 Exit Sub
errFix:
    MsgBox Err.Description, vbOKOnly, "Loan Types"
End Sub
Public Sub searchSelect()
    txtLoanCode.Text = sel
    load_records
    frmSearch.Visible = False
End Sub
Public Sub onRefreshOfSearchFrm()
On Error GoTo errFix
    frmSearch.lstSearch.ListItems.Clear
    strSQL = "Select * from loantype order by loancode"
    Set rst = oSaccoMaster.GetRecordSet(strSQL)
    With rst
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                Set li = frmSearch.lstSearch.ListItems.Add(, , !Loancode)
                li.SubItems(1) = !LoanType & ""
            Loop
        End If
        .Close
    End With
    Set rst = Nothing
    frmSearch.txtFrom.Text = ""
    frmSearch.txtTo.Text = ""
    frmSearch.cboCrieria.Text = "="
    frmSearch.cboField.Text = "Loan Code"
     Exit Sub
errFix:
    MsgBox Err.Description, vbOKOnly, "Loan Types"
End Sub
Public Sub onFormLoadOfSearchFrm()
On Error GoTo errFix
    PositionForm Me
    
    With frmSearch.lstSearch
        .ListItems.Clear
        .ColumnHeaders.Clear
        .ColumnHeaders.Add 1, , "Loan Code", 2000
        .ColumnHeaders.Add 2, , "Loan Type", 2000
        .View = lvwReport
        .GridLines = True
    End With
    frmSearch.Caption = "Search Loan Types"
    With frmSearch.cboField
        .AddItem ("Loan Code")
        .AddItem ("Loan Type")
    End With
    strSQL = "Select * from loantype order by loancode"
    Set rst = oSaccoMaster.GetRecordSet(strSQL)
    With rst
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                Set li = frmSearch.lstSearch.ListItems.Add(, , !Loancode)
                li.SubItems(1) = !LoanType & ""
                .MoveNext
            Loop
        End If
        .Close
    End With
    Set rst = Nothing
    frmSearch.cboCrieria.Text = frmSearch.cboCrieria.List(0)
    frmSearch.cboField.Text = frmSearch.cboField.List(0)
    frmSearch.Show vbModal
     Exit Sub
errFix:
    MsgBox Err.Description, vbOKOnly, "Loan Types"
End Sub

Public Sub onOkOfRangeForm()
On Error GoTo errFix
Set rst = oSaccoMaster.GetRecordSet("select loancode from loantype order by loancode")
If Not frmRangeSelection.txtFrom.Text = "" Then
    Set Rst4 = oSaccoMaster.GetRecordSet("select loancode from loantype where loancode= '" & frmRangeSelection.txtFrom.Text & "'")
    If Not Rst4.EOF Then
        rangeFrom = frmRangeSelection.txtFrom.Text
    Else
        MsgBox "Enter an existing No.", vbInformation, "Non-existent No"
        frmRangeSelection.txtFrom.Text = ""
        frmRangeSelection.txtFrom.SetFocus
        Exit Sub
    End If
Else
    If Not rst.EOF Then
        rst.MoveFirst
        rangeFrom = rst!Loancode
    End If
End If
If Not frmRangeSelection.txtTo.Text = "" Then
    Set Rst4 = oSaccoMaster.GetRecordSet("select loancode from loantype where loancode= '" & frmRangeSelection.txtTo.Text & "'")
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
        rangeTo = rst!Loancode
    End If
End If
If rangeFrom > rangeTo Then
    MsgBox "'To loan code' should be greater than or equal to 'From loan code'", vbInformation, "Check Range"
    frmRangeSelection.txtTo.SetFocus
    Exit Sub
End If
Set rst = oSaccoMaster.GetRecordSet("select loancode from loantype where loancode= '" & rangeTo & "'")
Set rst1 = oSaccoMaster.GetRecordSet("select loancode from loantype where loancode= '" & rangeFrom & "'")
If rst.RecordCount = 0 Then
    MsgBox "Enter an existing loan code for 'From' field", vbInformation, "Loan Code"
End If
If rst1.RecordCount = 0 Then
    MsgBox "Enter an existing Loan code for 'To' field", vbInformation, "Loan Code"
End If
If rst.RecordCount > 0 And rst1.RecordCount > 0 Then
     frmRangeSelection.Visible = False
      Dim reportname As String, SchemeType As String, STRFORMULA As String
            Dim title As String, Year As Integer
        Select Case reportType
        Case "Loan Applications Report"
            'Set a = New CRAXDRT.APPLICATION
            'Set R = a.OpenReport(App.path & "\Reports\Loan Applications Report.rpt")
            'R.RecordSelectionFormula = "{Loantype.loancode}= '" & rangeFrom & "' to '" & rangeTo & "'"
            'R.ReadRecords
            Set Rst5 = oSaccoMaster.GetRecordSet("select companyname from sysparam")
            If Not Rst5.EOF Then
                title = Rst5!CompanyName & ""
            End If
           
       
             Set cn = New ADODB.Connection
            Provider = "DSN=BOSA"
            cn.Open Provider
   ' Membership Details Report
            reportname = "Loan Applications Report.rpt"
             STRFORMULA = "{Loantype.loancode}= '" & rangeFrom & "' to '" & rangeTo & "'"
            Show_Sales_Crystal_Report STRFORMULA, reportname, title
            If chkPreviewReport.Value = vbUnchecked Then
                'R.PrintOut
            Else
                
                'With frmReports
                   ' .Show vbModal, Me
                'End With
            End If
        Case "Loan Types Report"
            Set a = New CRAXDRT.Application
            Set r = a.OpenReport(App.Path & "\Reports\Loan Types Report.rpt")
            r.RecordSelectionFormula = "{Loantype.loancode}= '" & rangeFrom & "' to '" & rangeTo & "'"
            r.ReadRecords
            Set Rst5 = oSaccoMaster.GetRecordSet("select companyname from sysparam")
            If Not Rst5.EOF Then
                title = Rst5!CompanyName & ""
            End If
                Set cn = New ADODB.Connection
                Provider = "DSN=BOSA"
                cn.Open Provider
                ' Membership Details Report
                reportname = "Loan Types Report.rpt"
                STRFORMULA = "{Loantype.loancode}= '" & rangeFrom & "' to '" & rangeTo & "'"
                Show_Sales_Crystal_Report STRFORMULA, reportname, title
                
                With frmReports
                    '.Show vbModal, Me
                End With
            'End If
    End Select
End If
 Exit Sub
errFix:
    MsgBox Err.Description, vbOKOnly, "Loan Types"
End Sub

Public Sub onFindOfSearchFrmClick()
On Error GoTo errFix
    Dim Find As Integer
    If frmSearch.cboField.Text = "Loan Types" Then
        searchField = "loantype"
    End If
    If frmSearch.cboField.Text = "Loan Code" Then
        searchField = "loancode"
    End If
    frmSearch.lstSearch.ListItems.Clear
    If Not frmSearch.cboField.Text = "" Then
        If Not frmSearch.cboCrieria.Text = "" Then
            If Not frmSearch.cboCrieria.Text = "Between" And Not frmSearch.cboCrieria.Text = "Like" Then
                strSQL = "Select * from loantype where " & searchField & " " & frmSearch.cboCrieria.Text & " '" & frmSearch.txtFrom.Text & "'"
                Set rst = oSaccoMaster.GetRecordSet(strSQL)
                With rst
                    If .RecordCount > 0 Then
                        .MoveFirst
                        Do While Not .EOF
                            Set li = frmSearch.lstSearch.ListItems.Add(, , !Loancode & "")
                            li.ListSubItems.Add , , !LoanType & ""
                            .MoveNext
                        Loop
                    End If
                End With
                Set rst = Nothing
            ElseIf frmSearch.cboCrieria.Text = "Like" Then
                sql = "Select * from loantype order by loancode"
                Set rst = oSaccoMaster.GetRecordSet(strSQL)
                With rst
                    If .RecordCount > 0 Then
                        .MoveFirst
                        Do While Not .EOF
                            .Find "" & searchField & " " & frmSearch.cboCrieria.Text & " '" & frmSearch.txtFrom.Text & "%'", , adSearchForward
                            If Not .EOF Then
                                Set li = frmSearch.lstSearch.ListItems.Add(, , !Loancode & "")
                                li.ListSubItems.Add , , !LoanType & ""
                                .MoveNext
                            End If
                        Loop
                    End If
                End With
                Set rst = Nothing
            Else
                    strSQL = "select * from loantype where " & searchField & " between'" & frmSearch.txtFrom.Text & "' And '" & frmSearch.txtTo.Text & " '"
                    Set rst = oSaccoMaster.GetRecordSet(strSQL)
                With rst
                    If .RecordCount > 0 Then
                        .MoveFirst
                        Do While Not .EOF
                            
                            Set li = frmSearch.lstSearch.ListItems.Add(, , !Loancode & "")
                            li.SubItems(1) = !LoanType & ""
                            .MoveNext
                        Loop
                    End If
                End With
                Set rst = Nothing
            End If
        Else
            MsgBox "Select the search criteria.", vbExclamation
        End If
    Else
        MsgBox "Select the search field.", vbExclamation
    End If
     Exit Sub
errFix:
    MsgBox Err.Description, vbOKOnly, "Loan Types"
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "bSearch"
        Set formCallingSearch = frmLoanTypes
        onFormLoadOfSearchFrm
    End Select
End Sub
Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
On Error GoTo errFix
    Select Case ButtonMenu.Key
    Case "mDetails"
        load_records
        frameLoanTypes.Visible = True
        lvwSummary.Visible = False
        lvwLoanApplications.Visible = True
        lblLoanApp.Visible = True
        cmdLast.Enabled = True
        cmdPrev.Enabled = True
        cmdNext.Enabled = True
        cmdFirst.Enabled = True
    Case "mSummary"
        frameLoanTypes.Visible = False
        lvwSummary.Visible = True
        lvwLoanApplications.Visible = True
        lblLoanApp.Visible = False
        load_Summary
        cmdLast.Enabled = False
        cmdPrev.Enabled = False
        cmdNext.Enabled = False
        cmdFirst.Enabled = False
        cmdCancel.Enabled = False
    Case "mLoanTypes"
        reportType = "Loan Types Report"
        Set formCallingRangeSelector = Me
        frmRangeSelection.txtTo.Text = ""
        frmRangeSelection.txtFrom.Text = ""
        
        frmRangeSelection.Caption = "Select Loan Range"
        frmRangeSelection.lblFrom.Caption = "From Loan No"
        frmRangeSelection.lblTo.Caption = "To Loan No"
        frmRangeSelection.Show vbModal
    Case "mLoanApplications"
        reportType = "Loan Applications Report"
        Set formCallingRangeSelector = Me
       frmRangeSelection.txtTo.Text = ""
        frmRangeSelection.txtFrom.Text = ""
       
        frmRangeSelection.Caption = "Select Loan Range"
        frmRangeSelection.lblFrom.Caption = "From Loan No"
        frmRangeSelection.lblTo.Caption = "To Loan No"
        frmRangeSelection.Show vbModal
    End Select
     Exit Sub
errFix:
    MsgBox Err.Description, vbOKOnly, "Loan Types"
End Sub


Private Sub txtContraAccName_Change()
On Error GoTo SysError
    Dim rsAccount As New Recordset
    lvwContraAcc.ListItems.Clear
    
    If Trim$(txtContraAccName) <> "" Then
        If Editing = True Then
            Set rsAccount = oSaccoMaster.GetRecordSet("Select * From GLSETUP where " _
            & "GLAccName Like '%" & txtContraAccName & "%'")
            With rsAccount
                If .State = adStateOpen Then
                    If Not .EOF Then
                        lvwContraAcc.Visible = True
                        If .RecordCount = 1 Then
                            txtContraAcc = IIf(IsNull(!AccNo), "", !AccNo)
                            Editing = True
                            txtContraAccName = IIf(IsNull(!GlAccName), "", !GlAccName)
                            lvwInterestAcc.Visible = False
                            Exit Sub
                        End If
                    Else
                        lvwContraAcc.Visible = False
                    End If
                    While Not .EOF
                        Set li = lvwContraAcc.ListItems.Add(, , IIf(IsNull(!AccNo), "", !AccNo))
                        li.SubItems(1) = IIf(IsNull(!GlAccName), "", !GlAccName)
                        .MoveNext
                    Wend
                End If
            End With
        End If
    End If
    Exit Sub
SysError:
    MsgBox Err.Description, vbInformation, Me.Caption
End Sub

Private Sub txtInterestAccName_Change()
On Error GoTo SysError
    Dim rsAccount As New Recordset
    lvwInterestAcc.ListItems.Clear
    If Trim$(txtInterestAccName) <> "" Then
        If Editing = True Then
            Set rsAccount = oSaccoMaster.GetRecordSet("Select * From GLSETUP where " _
            & "GLAccName Like '%" & txtInterestAccName & "%'")
            With rsAccount
                If .State = adStateOpen Then
                    If Not .EOF Then
                        lvwInterestAcc.Visible = True
                        If .RecordCount = 1 Then
                            txtInterestAcc = IIf(IsNull(!AccNo), "", !AccNo)
                            txtInterestAccName = IIf(IsNull(!GlAccName), "", !GlAccName)
                            Editing = True
                           ' txtAccName_Change
                            lvwInterestAcc.Visible = False
                            Exit Sub
                        End If
                    Else
                        lvwInterestAcc.Visible = False
                    End If
                    While Not .EOF
                        Set li = lvwInterestAcc.ListItems.Add(, , IIf(IsNull(!AccNo), "", !AccNo))
                        li.SubItems(1) = IIf(IsNull(!GlAccName), "", !GlAccName)
                        .MoveNext
                    Wend
                End If
            End With
        End If
    End If
    Exit Sub
SysError:
    MsgBox Err.Description, vbInformation, Me.Caption
End Sub

Private Sub txtInterestRate_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    txtNumOfLoans.SetFocus
End If
If KeyAscii = 13 Then
    txtMaxAmmount.SetFocus
End If
If Len(Trim(txtInterestRate.Text)) > 2 Then
    If Not KeyAscii = 8 Then
    Beep
    MsgBox "Can't enter more than 3 characters", vbExclamation
    End If
    KeyAscii = 8
End If
  Select Case KeyAscii
    Case Asc("0") To Asc("9")
    Case Is = 8
    Case Is = 46
    Case Else

        Beep
        KeyAscii = 0
  End Select
 Exit Sub
errFix:
    MsgBox Err.Description, vbOKOnly, "Loan Types"
End Sub



Private Sub txtLoanAccName_Change()
 On Error GoTo SysError
    Dim rsAccount As New Recordset
    lvwLoanAcc.ListItems.Clear
    If Trim$(txtLoanAccName) <> "" Then
        If Editing = True Then
            Set rsAccount = oSaccoMaster.GetRecordSet("Select * From GLSETUP where " _
            & "GLAccName Like '%" & txtLoanAccName & "%'")
            With rsAccount
                If .State = adStateOpen Then
                    If Not .EOF Then
                        lvwLoanAcc.Visible = True
                        If .RecordCount = 1 Then
                            txtLoanAcc = IIf(IsNull(!AccNo), "", !AccNo)
                            txtLoanAccName = IIf(IsNull(!GlAccName), "", !GlAccName)
                            Editing = True
                            lvwLoanAcc.Visible = False
                            Exit Sub
                        End If
                    Else
                        lvwLoanAcc.Visible = False
                    End If
                    While Not .EOF
                        Set li = lvwLoanAcc.ListItems.Add(, , IIf(IsNull(!AccNo), "", !AccNo))
                        li.SubItems(1) = IIf(IsNull(!GlAccName), "", !GlAccName)
                        .MoveNext
                    Wend
                End If
            End With
        End If
    End If
    Exit Sub
SysError:
    MsgBox Err.Description, vbInformation, Me.Caption
End Sub

Private Sub txtLoanCode_Change()
On Error GoTo errFix
txtLoanCode.Text = UCase(txtLoanCode.Text)
txtLoanCode.SelStart = Len(txtLoanCode.Text)
 Exit Sub
errFix:
    MsgBox Err.Description, vbOKOnly, "Loan Types"
End Sub



Private Sub txtLoanCode_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If Len(txtLoanCode.Text) > 4 Then
      Beep
      MsgBox "Can't enter more than 5 characters", vbExclamation
      KeyAscii = 8
  End If
If KeyAscii = 13 Then
    If action = "addingRecords" Then
        Set Rst5 = oSaccoMaster.GetRecordSet("select loancode from loantype where loancode= '" & txtLoanCode.Text & "'")
        If Not Rst5.EOF Then
            MsgBox "Loan type with same code exists", vbInformation, "Loan Code"
            txtLoanCode.Text = ""
            txtLoanCode.SetFocus
            Exit Sub
        Else
            txtLoanType.SetFocus
        End If
    End If
End If
 Exit Sub
errFix:
    MsgBox Err.Description, vbOKOnly, "Loan Types"
End Sub

Private Sub txtLoanType_Change()
On Error GoTo errFix
txtLoanType.Text = UCase(txtLoanType.Text)
txtLoanType.SelStart = Len(txtLoanType.Text)
 Exit Sub
errFix:
    MsgBox Err.Description, vbOKOnly, "Loan Types"
End Sub

Private Sub txtLoanType_KeyPress(KeyAscii As Integer)
    On Error GoTo errFix
    If KeyAscii <> vbKeyReturn Then 'Catch the Enter key
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
    Exit Sub
errFix:
    MsgBox Err.Description, vbOKOnly, "Loan Types"
End Sub

Private Sub txtMaxAmmount_KeyPress(KeyAscii As Integer)
    On Error GoTo errFix
    If KeyAscii = 13 Then
        txtInterestRate.SetFocus
    End If
    Select Case KeyAscii
        Case Asc("0") To Asc("9")
        Case Is = 8
        Case Is = 46
        Case Else
        Beep
        KeyAscii = 0
    End Select
    Exit Sub
errFix:
    MsgBox Err.Description, vbOKOnly, "Loan Types"
End Sub

Private Sub txtNumOfLoans_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdUpdate.SetFocus
End If
End Sub

Private Sub txtRepayPeriod_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    txtMaxAmmount.SetFocus
End If
If Len(txtRepayPeriod.Text) > 1 Then
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
    MsgBox Err.Description, vbOKOnly, "Loan Types"
End Sub
