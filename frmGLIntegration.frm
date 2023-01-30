VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGLIntegration 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "GL Integration"
   ClientHeight    =   3000
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   6645
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView lvwCRAccount 
      Height          =   1440
      Left            =   3270
      TabIndex        =   14
      Top             =   1215
      Visible         =   0   'False
      Width           =   3105
      _ExtentX        =   5477
      _ExtentY        =   2540
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
   Begin VB.CommandButton cmdAdd 
      Caption         =   "New"
      Height          =   390
      Left            =   1020
      TabIndex        =   12
      Top             =   2505
      Width           =   960
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   390
      Left            =   2310
      TabIndex        =   9
      Top             =   2520
      Width           =   960
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   420
      Left            =   4875
      TabIndex        =   8
      Top             =   2505
      Width           =   1020
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2430
      Left            =   60
      TabIndex        =   1
      Top             =   30
      Width           =   6555
      Begin VB.TextBox txtAmount 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1485
         TabIndex        =   16
         Top             =   1995
         Width           =   1515
      End
      Begin VB.TextBox txtDescription 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3195
         TabIndex        =   15
         Top             =   435
         Width           =   3105
      End
      Begin VB.TextBox txtDRAccName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3210
         TabIndex        =   11
         Top             =   1560
         Width           =   3105
      End
      Begin VB.TextBox txtCRAccName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3210
         TabIndex        =   10
         Top             =   915
         Width           =   3105
      End
      Begin VB.ComboBox cboDeductionCode 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmGLIntegration.frx":0000
         Left            =   1530
         List            =   "frmGLIntegration.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   405
         Width           =   1605
      End
      Begin VB.TextBox txtDRAccount 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1500
         TabIndex        =   3
         Top             =   1545
         Width           =   1515
      End
      Begin VB.TextBox txtCRAccount 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1500
         TabIndex        =   2
         Top             =   930
         Width           =   1485
      End
      Begin MSComctlLib.ListView lvwDRAccount 
         Height          =   1455
         Left            =   3210
         TabIndex        =   13
         Top             =   1815
         Visible         =   0   'False
         Width           =   3105
         _ExtentX        =   5477
         _ExtentY        =   2566
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
      Begin VB.Label lblAmount 
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   17
         Top             =   1995
         Width           =   1395
      End
      Begin VB.Label Label18 
         Caption         =   "Debit Account"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   105
         TabIndex        =   7
         Top             =   1545
         Width           =   1395
      End
      Begin VB.Label Label17 
         Caption         =   "Credit Account"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   90
         TabIndex        =   6
         Top             =   960
         Width           =   1605
      End
      Begin VB.Label Label16 
         Caption         =   "Deduction Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   90
         TabIndex        =   5
         Top             =   420
         Width           =   1515
      End
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   405
      Left            =   3630
      TabIndex        =   0
      Top             =   2520
      Width           =   1020
   End
End
Attribute VB_Name = "frmGLIntegration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cboDeductionCode_Click()
Select Case cboDeductionCode
Case "REGFEE"
    txtDescription = "Registration Fees"
Case "BLW"
    txtDescription = "By-Laws"
Case "PREM"
    txtDescription = "Premium"
End Select
End Sub

Private Sub cmdAdd_Click()
NewRecord = True
cboDeductionCode.ListIndex = -1
txtCRAccName = ""
txtCRAccount = ""
txtDRAccName = ""
txtDRAccount = ""
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdEdit_Click()
NewRecord = False
End Sub

Private Sub cmdSave_Click()
On Error GoTo Syserr

    If Not Save_OtherGLAccounts(cboDeductionCode, txtDescription, txtCRAccount, txtDRAccount, ErrorMessage, txtAmount, NewRecord) Then
        If ErrorMessage <> "" Then
            MsgBox ErrorMessage, vbInformation, Me.Caption
            ErrorMessage = ""
            Exit Sub
        End If
    End If
Exit Sub
Syserr:
MsgBox Err.Description
End Sub

Private Sub lvwCRAccount_DblClick()
On Error GoTo SysError
    If lvwCRAccount.ListItems.count > 0 Then
        txtCRAccount = lvwCRAccount.SelectedItem
        txtCRAccName = lvwCRAccount.SelectedItem.SubItems(1)
        lvwCRAccount.Visible = False
    End If
    Exit Sub
SysError:
    MsgBox Err.Description, vbInformation, Me.Caption
End Sub

Private Sub lvwDRAccount_DblClick()
On Error GoTo SysError
    If lvwDRAccount.ListItems.count > 0 Then
        txtDRAccount = lvwDRAccount.SelectedItem
        txtDRAccName = lvwDRAccount.SelectedItem.SubItems(1)
        lvwDRAccount.Visible = False
    End If
    Exit Sub
SysError:
    MsgBox Err.Description, vbInformation, Me.Caption
End Sub

Private Sub txtCRAccName_Change()
On Error GoTo SysError
    Dim rsAccount As New Recordset
    lvwCRAccount.ListItems.Clear
    If Trim$(txtCRAccName) <> "" Then
      '  If Editing = True Then
            Set rsAccount = oSaccoMaster.GetRecordSet("Select * From GLSETUP where " _
            & "GLAccName Like '%" & txtCRAccName & "%'")
            With rsAccount
                If .State = adStateOpen Then
                    If Not .EOF Then
                        lvwCRAccount.Visible = True
                        If .RecordCount = 1 Then
                            txtCRAccount = IIf(IsNull(!Accno), "", !Accno)
                            Editing = True
                          
                            lvwCRAccount.Visible = False
                            Exit Sub
                        End If
                    Else
                        lvwCRAccount.Visible = False
                    End If
                    While Not .EOF
                        Set li = lvwCRAccount.ListItems.Add(, , IIf(IsNull(!Accno), "", !Accno))
                        li.SubItems(1) = IIf(IsNull(!GlAccName), "", !GlAccName)
                        .MoveNext
                    Wend
                End If
            End With
      '  End If
    End If
    Exit Sub
SysError:
    MsgBox Err.Description, vbInformation, Me.Caption
End Sub

Private Sub txtDrAccName_Change()
On Error GoTo SysError
    Dim rsAccount As New Recordset
    lvwDRAccount.ListItems.Clear
    If Trim$(txtDRAccName) <> "" Then
        'If NewRecord = True Then
            Set rsAccount = oSaccoMaster.GetRecordSet("Select * From GLSETUP where " _
            & "GLAccName Like '%" & txtDRAccName & "%'")
            With rsAccount
                If .State = adStateOpen Then
                    If Not .EOF Then
                        lvwDRAccount.Visible = True
                        If .RecordCount = 1 Then
                            txtDRAccount = IIf(IsNull(!Accno), "", !Accno)
                            Editing = True
                          
                            lvwDRAccount.Visible = False
                            Exit Sub
                        End If
                    Else
                        lvwDRAccount.Visible = False
                    End If
                    While Not .EOF
                        Set li = lvwDRAccount.ListItems.Add(, , IIf(IsNull(!Accno), "", !Accno))
                        li.SubItems(1) = IIf(IsNull(!GlAccName), "", !GlAccName)
                        .MoveNext
                    Wend
                End If
            End With
       ' End If
    End If
    Exit Sub
SysError:
    MsgBox Err.Description, vbInformation, Me.Caption
End Sub
