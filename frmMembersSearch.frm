VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMembersSearch 
   Caption         =   "Members Search Form"
   ClientHeight    =   2670
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6810
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMembersSearch.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   6810
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lvwNames 
      Height          =   1170
      Left            =   1680
      TabIndex        =   9
      Top             =   1515
      Visible         =   0   'False
      Width           =   3360
      _ExtentX        =   5927
      _ExtentY        =   2064
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
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
         Text            =   "Names"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "MemNo"
         Object.Width           =   18
      EndProperty
   End
   Begin VB.TextBox txtNames 
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
      Left            =   1665
      TabIndex        =   8
      Top             =   1200
      Width           =   3390
   End
   Begin VB.TextBox txtMemberNo 
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
      Left            =   1665
      TabIndex        =   7
      Top             =   705
      Width           =   1470
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   405
      Left            =   3435
      TabIndex        =   4
      Top             =   2085
      Width           =   1620
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   405
      Left            =   1680
      TabIndex        =   3
      Top             =   2085
      Width           =   1620
   End
   Begin VB.ComboBox cboOrgName 
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
      Left            =   3120
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   225
      Width           =   3615
   End
   Begin VB.ComboBox cboOrgCode 
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
      Left            =   1680
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   225
      Width           =   1335
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Member Names:"
      Height          =   210
      Left            =   315
      TabIndex        =   6
      Top             =   1230
      Width           =   1320
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Member No:"
      Height          =   210
      Left            =   630
      TabIndex        =   5
      Top             =   735
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Organisation:"
      Height          =   210
      Left            =   570
      TabIndex        =   0
      Top             =   255
      Width           =   1065
   End
End
Attribute VB_Name = "frmMembersSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboOrgCode_Change()
    On Error GoTo SysError
    Dim rsComp As New Recordset, CompCode As String, CompName As String
    'cboOrgCode.Clear
    'cboOrgName.Clear
    If Trim$(cboOrgCode) <> "" Then
        Set rsComp = oSaccoMaster.GetRecordSet("Select * From COMPANY where " _
        & "CompanyCode Like '" & cboOrgCode & "%'")
        With rsComp
            If .State = adStateOpen Then
                While Not .EOF
                    cboOrgCode.AddItem IIf(IsNull(!companycode), "", !companycode)
                    cboOrgName.AddItem IIf(IsNull(!CompanyName), "", !CompanyName)
                    .MoveNext
                Wend
            End If
        End With
    End If
    Exit Sub
SysError:
    MsgBox Err.description, vbInformation, Me.Caption
End Sub

Private Sub cboOrgCode_Click()
    On Error GoTo SysError
    Dim rsOrg As New Recordset
    If Trim$(cboOrgCode) <> "" Then
        Set rsOrg = oSaccoMaster.GetRecordSet("Select CompanyName From COMPANY where " _
        & "CompanyCode='" & cboOrgCode & "'")
        With rsOrg
            If .State = adStateOpen Then
                If Not .EOF Then
                    cboOrgName = IIf(IsNull(!CompanyName), "", !CompanyName)
                End If
            End If
        End With
    End If
    Exit Sub
SysError:
    MsgBox Err.description, vbInformation, Me.Caption
End Sub

Private Sub cmdOk_Click()
    If Trim$(txtMemberno) <> "" Then
        SearchValue = txtMemberno
        Group = False
    Else
        If cboOrgCode <> "All" Then
            SearchValue = cboOrgCode
            Group = True
        Else
            SearchValue = "All"
        End If
    End If
    Continue = True
    Unload Me
End Sub

Private Sub Form_Load()
    On Error GoTo SysError
    Dim rsMem As New Recordset
    cboOrgCode.Clear
    cboOrgName.Clear
    cboOrgCode.AddItem "All"
    cboOrgName.AddItem "All"
    Set rsMem = oSaccoMaster.GetRecordSet("Select * From Company order By CompanyCode")
    With rsMem
        If .State = adStateOpen Then
            While Not .EOF
                cboOrgCode.AddItem IIf(IsNull(!companycode), "", !companycode)
                cboOrgName.AddItem IIf(IsNull(!CompanyName), "", !CompanyName)
                .MoveNext
            Wend
        End If
    End With
    cboOrgCode.ListIndex = 0
    cboOrgName.ListIndex = 0
    Exit Sub
SysError:
    MsgBox Err.description, vbInformation, Me.Caption
End Sub

Private Sub lvwNames_Click()
    On Error GoTo SysError
    Dim mMemberNo As String, mName As String
    If lvwNames.ListItems.Count > 0 Then
        mName = lvwNames.SelectedItem
        mMemberNo = lvwNames.SelectedItem.SubItems(1)
        If Not li Is Nothing Then
            Editing = True
            txtNames = mName
            txtMemberno = mMemberNo
            lvwNames.ListItems.Clear
            lvwNames.Visible = False
        End If
    End If
    Exit Sub
SysError:
    MsgBox Err.description, vbInformation, Me.Caption
End Sub

Private Sub txtMemberno_Change()
    On Error GoTo SysError
    Dim RsMembers As New Recordset
    If Trim$(txtMemberno) <> "" Then
        Set RsMembers = oSaccoMaster.GetRecordSet("Select * From MEMBERS where MemberNo='" & txtMemberno & "'")
        With RsMembers
            If .State = adStateOpen Then
                If Not .EOF Then
                    txtNames = IIf(IsNull(!othernames), "", !othernames) & " " _
                    & IIf(IsNull(!surname), "", !surname)
                Else
                    txtNames = ""
                End If
            End If
        End With
    Else
        txtNames = ""
    End If
    Exit Sub
SysError:
    MsgBox Err.description, vbInformation, Me.Caption
End Sub

Private Sub txtNames_Change()
    On Error GoTo SysError
    Dim RsNames As New Recordset
    lvwNames.ListItems.Clear
    If Not Editing Then
        If Trim$(txtNames) <> "" Then
            If Trim$(cboOrgCode) <> "All" Then
                Set RsNames = oSaccoMaster.GetRecordSet("Select OtherNames + ' ' + SurName " _
                & "as [Names],MemberNo From MEMBERS where CompanyCode='" & cboOrgCode & "' and " _
                & "OtherNames Like '%" & txtNames & "%'")
            Else
                Set RsNames = oSaccoMaster.GetRecordSet("Select OtherNames + ' ' + SurName " _
                & "as [Names],MemberNo From MEMBERS where OtherNames Like '%" & txtNames & "%'")
            End If
            With RsNames
                If .State = adStateOpen Then
                    If Not .EOF Then
                        lvwNames.Visible = True
                        While Not .EOF
                            Set li = lvwNames.ListItems.Add(, , IIf(IsNull(![names]), "", ![names]))
                            li.SubItems(1) = IIf(IsNull(!memberno), "", !memberno)
                            .MoveNext
                        Wend
                    Else
                        lvwNames.Visible = False
                    End If
                End If
            End With
        End If
    End If
    Editing = False
    Exit Sub
SysError:
    MsgBox Err.description, vbInformation, Me.Caption
End Sub
