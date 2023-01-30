VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSearchGLAccounts 
   Caption         =   "Search GL Accounts"
   ClientHeight    =   4245
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6780
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   6780
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3375
      TabIndex        =   4
      Top             =   3675
      Width           =   1290
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   1965
      TabIndex        =   3
      Top             =   3690
      Width           =   1290
   End
   Begin VB.TextBox txtValue 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1770
      TabIndex        =   0
      Top             =   345
      Width           =   3645
   End
   Begin VB.ComboBox cboField 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmSearchGLAccounts.frx":0000
      Left            =   90
      List            =   "frmSearchGLAccounts.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   345
      Width           =   1605
   End
   Begin MSComctlLib.ListView lvwAccounts 
      Height          =   2655
      Left            =   75
      TabIndex        =   1
      Top             =   840
      Width           =   6675
      _ExtentX        =   11774
      _ExtentY        =   4683
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "GL AccountNo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "GL Account Name"
         Object.Width           =   7056
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Search Value"
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
      Left            =   1785
      TabIndex        =   6
      Top             =   90
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Search Field"
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
      Left            =   105
      TabIndex        =   5
      Top             =   90
      Width           =   1005
   End
End
Attribute VB_Name = "frmSearchGLAccounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Continue = False
    Unload Me
End Sub

Private Sub cmdOk_Click()
    On Error GoTo SysError
    If lvwAccounts.ListItems.Count > 0 Then
        SearchValue = lvwAccounts.SelectedItem
        Continue = True
    Else
        SearchValue = ""
    End If
    Unload Me
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub Form_Load()
    On Error GoTo SysError
    cboField.ListIndex = 0
    txtValue_Change
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub lvwAccounts_DblClick()
    cmdOk_Click
End Sub

Private Sub txtValue_Change()
    On Error GoTo SysError
    lvwAccounts.ListItems.Clear
    Select Case cboField.ListIndex
        Case 0
        GetRecords ("Select AccNo,GLAccName From GLSETUP Where GLAccName Like '%" & txtValue.Text & "%'")
        Case 1
        GetRecords ("Select AccNo,GLAccName From GLSETUP Where AccNo Like '%" & txtValue & "%'")
    End Select
    With Rst
        If .State = adStateOpen Then
            While Not .EOF
                Set li = lvwAccounts.ListItems.Add(, , IIf(IsNull(!accno), "", !accno))
                li.SubItems(1) = IIf(IsNull(!GlAccName), "", !GlAccName)
                .MoveNext
            Wend
        End If
    End With
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub
