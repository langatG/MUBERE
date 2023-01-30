VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmGLMainSetup 
   Caption         =   "Main GL Account Setup"
   ClientHeight    =   5160
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6555
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
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   6555
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lvwMainAcc 
      Height          =   2160
      Left            =   120
      TabIndex        =   11
      Top             =   2295
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   3810
      View            =   3
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "AccNo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "AccType"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   375
      Left            =   705
      TabIndex        =   10
      Top             =   4650
      Width           =   1230
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   375
      Left            =   2025
      TabIndex        =   9
      Top             =   4650
      Width           =   1230
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   4665
      TabIndex        =   8
      Top             =   4650
      Width           =   1230
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   3345
      TabIndex        =   7
      Top             =   4650
      Width           =   1230
   End
   Begin VB.Frame Frame1 
      Caption         =   "Main GL Accounts"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2070
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      Begin VB.ComboBox cboAccType 
         Appearance      =   0  'Flat
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
         ItemData        =   "frmGLMainSetup.frx":0000
         Left            =   1890
         List            =   "frmGLMainSetup.frx":000D
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1335
         Width           =   3735
      End
      Begin VB.TextBox txtaccno 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1905
         MaxLength       =   20
         TabIndex        =   2
         Top             =   480
         Width           =   2160
      End
      Begin VB.TextBox txtaccname 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1905
         MaxLength       =   50
         TabIndex        =   1
         Top             =   915
         Width           =   3615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Account Type"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   480
         TabIndex        =   6
         Top             =   1380
         Width           =   1170
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Account No"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   480
         TabIndex        =   5
         Top             =   525
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Account name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   480
         TabIndex        =   4
         Top             =   930
         Width           =   1200
      End
   End
End
Attribute VB_Name = "frmGLMainSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdedit_Click()
    NewRecord = False
End Sub

Private Sub cmdnew_Click()
    NewRecord = True
    txtaccname = ""
    Txtaccno = ""
    cboAccType.ListIndex = -1
    Txtaccno.SetFocus
End Sub

Private Sub cmdsave_Click()
    On Error GoTo SysError
    If Trim(Txtaccno) = "" Then
        MsgBox "Enter the Main Account Number"
        Txtaccno.SetFocus
        Exit Sub
    End If
    If Trim(txtaccname) = "" Then
        MsgBox "Enter the Main Account Name"
        txtaccname.SetFocus
        Exit Sub
    End If
    If NewRecord = True Then
        If Not SAVE_GLMAINACCOUNTS(Txtaccno, txtaccname, cboAccType, _
        ErrorMessage) Then
            If ErrorMessage <> "" Then
            MsgBox ErrorMessage, vbInformation, Me.Caption
            ErrorMessage = ""
            End If
        End If
    Else
        If Not UPDATE_GLMAINACCOUNTS(Txtaccno, txtaccname, cboAccType, _
        ErrorMessage) Then
            If ErrorMessage <> "" Then
            MsgBox ErrorMessage, vbInformation, Me.Caption
            ErrorMessage = ""
            End If
        End If
    End If
    NewRecord = False
    Populate_List
    Exit Sub
SysError:
    MsgBox Err.description
End Sub

Private Sub Form_Load()
    Populate_List
End Sub

Private Sub Populate_List()
    On Error GoTo SysError
    Dim rsMain As New Recordset
    lvwMainAcc.ListItems.Clear
    Set rsMain = Get_Records("Select * From GLMAINACCOUNTS order by AccountType", ErrorMessage)
    With rsMain
        If .State = adStateClosed Then
            If ErrorMessage <> "" Then
                MsgBox ErrorMessage, vbInformation, Me.Caption
                ErrorMessage = ""
            End If
        Else
            While Not .EOF
                Set li = lvwMainAcc.ListItems.Add(, , IIf(IsNull(!MainAccNo), "", !MainAccNo))
                li.SubItems(1) = IIf(IsNull(!MainAccName), "", !MainAccName)
                li.SubItems(2) = IIf(IsNull(!AccountType), "", !AccountType)
                .MoveNext
            Wend
        End If
    End With
    Exit Sub
SysError:
    MsgBox Err.description, vbInformation, Me.Caption
End Sub

Private Sub lvwMainAcc_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error GoTo SysError
    If lvwMainAcc.ListItems.Count > 0 Then
        txtaccname = lvwMainAcc.SelectedItem.SubItems(1)
        Txtaccno = lvwMainAcc.SelectedItem
        cboAccType = lvwMainAcc.SelectedItem.SubItems(2)
    End If
    Exit Sub
SysError:
    MsgBox Err.description, vbInformation, Me.Caption
End Sub

Private Sub txtaccname_KeyPress(KeyAscii As Integer)
    KeyAscii = To_Upper_Case(KeyAscii)
End Sub

Private Sub txtAccNo_Change()
    On Error GoTo SysError
    Dim GLMainAcc As GL_MainAcc
    GLMainAcc = Get_GLMainAcc_Details(Txtaccno, ErrorMessage)
    If GLMainAcc.AccCode <> "" Then
        txtaccname = GLMainAcc.accname
        cboAccType = GLMainAcc.AccountType
    Else
        txtaccname = ""
        cboAccType.ListIndex = -1
        If ErrorMessage <> "" Then
            MsgBox ErrorMessage, vbInformation, Me.Caption
            ErrorMessage = ""
        End If
    End If
    Exit Sub
SysError:
    MsgBox Err.description, vbInformation, Me.Caption
End Sub

Private Sub txtAccNo_KeyPress(KeyAscii As Integer)
    KeyAscii = To_Upper_Case(KeyAscii)
End Sub
