VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSearchLoan 
   Caption         =   "Search"
   ClientHeight    =   4875
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6210
   Icon            =   "Search.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   4875
   ScaleWidth      =   6210
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   16
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   15
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   ">>|"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   14
      Top             =   4440
      Width           =   615
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   13
      Top             =   4440
      Width           =   615
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   12
      Top             =   4440
      Width           =   615
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "|<<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   4440
      Width           =   615
   End
   Begin MSComctlLib.ListView lstSearch 
      Height          =   2535
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   4471
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
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
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   3840
      TabIndex        =   4
      Top             =   720
      Width           =   2295
      Begin VB.TextBox txtTo 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox txtFrom 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.ComboBox cboCrieria 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "Search.frx":0442
      Left            =   2280
      List            =   "Search.frx":045B
      TabIndex        =   3
      Top             =   960
      Width           =   1455
   End
   Begin VB.ComboBox cboField 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "Search.frx":047F
      Left            =   120
      List            =   "Search.frx":0481
      TabIndex        =   2
      Top             =   960
      Width           =   2055
   End
   Begin MSComctlLib.Toolbar tbrSearch 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6210
      _ExtentX        =   10954
      _ExtentY        =   1005
      ButtonWidth     =   714
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Find"
            Key             =   "find"
         EndProperty
      EndProperty
      BorderStyle     =   1
      MousePointer    =   1
   End
   Begin VB.Label Label3 
      Caption         =   "Criteria"
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
      Left            =   2280
      TabIndex        =   10
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label lblRecords 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1440
      TabIndex        =   8
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Records Found"
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
      TabIndex        =   7
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Search Field"
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
      TabIndex        =   1
      Top             =   720
      Width           =   1575
   End
End
Attribute VB_Name = "frmSearchLoan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboCrieria_Click()
  If cboCrieria.Text = "Between" Then
    txtTo.Visible = True
  Else
    txtTo.Visible = False
  End If
End Sub

Private Sub cboField_Click()
  Connect
  lstSearch.ListItems.Clear
    With Rst
    If SItem = 1 Then
      If cboField.Text = "MemberNo" Then
        .Open "select loanno,memberno,loanamt from loans order by memberno", Cnn, adOpenKeyset, adLockOptimistic
      ElseIf cboField.Text = "Surname" Then
        .Open "select loanno,memberno,loanamt from loans order by loanno", Cnn, adOpenKeyset, adLockOptimistic
      Else
        .Open "select loanno,memberno,loanamt from loans order by loanamt", Cnn, adOpenKeyset, adLockOptimistic
      End If
    ElseIf SItem = 2 Or SItem = 0 Then
      If cboField.Text = "Member No" Then
        .Open "select memberno,surname,othernames from members order by memberno", Cnn, adOpenKeyset, adLockOptimistic
      ElseIf cboField.Text = "Surname" Then
        .Open "select memberno,surname,othernames from members order by surname", Cnn, adOpenKeyset, adLockOptimistic
      Else
        .Open "select memberno,surname,othernames from members order by othernames", Cnn, adOpenKeyset, adLockOptimistic
      End If
    End If
      If .RecordCount > 0 Then
        .MoveFirst
        While Not .EOF
          With lstSearch
            If SItem = 2 Or SItem = 0 Then
              Set Li = .ListItems.Add(, , Rst!memberno)
              Li.SubItems(1) = Rst!surname
              Li.SubItems(2) = Rst!othernames
            Else
              Set Li = .ListItems.Add(, , Rst!LoanNo)
              Li.SubItems(1) = Rst!memberno
              Li.SubItems(2) = Format(Rst!loanamt, "###,###,###,##0.00")
            End If
          End With
          .MoveNext
        Wend
      End If
      .Close
    End With
  
End Sub

Private Sub cmdSelect_Click()
  Connect
  
  With Rst
    If SItem > 0 Then
      .Open "select loanno from loans where loanno='" & lstSearch.SelectedItem.Text & "'", Cnn, adOpenKeyset, adLockOptimistic
      If .RecordCount > 0 Then
        frmApplic.txtLoanNo.Text = !LoanNo
        frmApplic.LoadRec
      End If
    Else
      .Open "select memberno from members where memberno='" & lstSearch.SelectedItem.Text & "'", Cnn, adOpenKeyset, adLockOptimistic
      If .RecordCount > 0 Then
        frmApplic.txtMemberno.Text = !memberno
        LoadMember
        LoadLoan
      End If
    End If
    .Close
  End With
  Unload Me
End Sub

Private Sub Form_Load()
  Me.Top = (Screen.Height - Height) / 2
  Me.Left = (Screen.Width - Width) / 2
End Sub

Private Sub tbrSearch_ButtonClick(ByVal Button As MSComctlLib.Button)
  SItem = Button.Index
  If SItem = 1 Then
    Connect
    With Rst
      If cboField.Text <> "Amount" Then
        lstSearch.ListItems.Clear
        If cboCrieria.Text = "Between" Then
          .Open "select loanno,memberno,loanamt from loans", Cnn, adOpenKeyset, adLockOptimistic
          If .RecordCount > 0 Then
            While Not .EOF
              If .RecordCount > 0 Then
              If cboField.Text = "LoanNo" Then
                If Rst!LoanNo >= txtFrom.Text And Rst!LoanNo <= txtTo.Text Then
                  With lstSearch
                    Set Li = .ListItems.Add(, , Rst!LoanNo)
                    Li.SubItems(1) = Rst!memberno
                    Li.SubItems(2) = Format(Rst!loanamt, "###,###,###,##0.00")
                  End With
                End If
              ElseIf cboField.Text = "MemberNo" Then
                If Rst!memberno >= txtFrom.Text And Rst!memberno <= txtTo Then
                  With lstSearch
                    Set Li = .ListItems.Add(, , Rst!LoanNo)
                    Li.SubItems(1) = Rst!memberno
                    Li.SubItems(2) = Format(Rst!loanamt, "###,###,###,##0.00")
                  End With
                End If
              End If
              End If
              .MoveNext
            Wend
            lblRecords.Caption = lstSearch.ListItems.Count
          End If
        ElseIf cboCrieria = "Like" Then
          .Open "select loanno,memberno,loanamt from loans where " & cboField.Text & " " & cboCrieria.Text & " '" & txtFrom.Text & "%'", Cnn, adOpenKeyset, adLockOptimistic
          If .RecordCount > 0 Then
            While Not .EOF
              '.Find "" & cboField.Text & " " & cboCrieria.Text & " '" & txtFrom.Text & "%'", , adSearchForward, adBookmarkCurrent
              'If .RecordCount > 0 Then
                With lstSearch
                  Set Li = .ListItems.Add(, , Rst!LoanNo)
                  Li.SubItems(1) = Rst!memberno
                  Li.SubItems(2) = Format(Rst!loanamt, "###,###,###,##0.00")
                End With
              'End If
              .MoveNext
            Wend
          End If
          lblRecords.Caption = .RecordCount
        Else
          .Open "select loanno,memberno,loanamt from loans", Cnn, adOpenKeyset, adLockOptimistic
          If .RecordCount > 0 Then
            .MoveFirst
            While Not .EOF
              .Find "" & cboField.Text & " " & cboCrieria.Text & " '" & txtFrom & "'", , adSearchForward, adBookmarkCurrent
              If Not .EOF Then
                With lstSearch
                  Set Li = .ListItems.Add(, , Rst!LoanNo)
                  Li.SubItems(1) = Rst!memberno
                  Li.SubItems(2) = Format(Rst!loanamt, "###,###,###,##0.00")
                End With
                .MoveNext
              End If
            Wend
          End If
        End If
      Else
        lstSearch.ListItems.Clear
        If cboCrieria.Text = "Between" Then
          .Open "select loanno,memberno,loanamt from loans", Cnn, adOpenKeyset, adLockOptimistic
          If .RecordCount > 0 Then
            .MoveFirst
            While Not .EOF
              If Rst!loanamt >= CCur(txtFrom.Text) And Rst!loanamt <= CCur(txtTo.Text) Then
                With lstSearch
                  Set Li = .ListItems.Add(, , Rst!LoanNo)
                  Li.SubItems(1) = Rst!memberno
                  Li.SubItems(2) = Format(Rst!loanamt, "###,###,###,##0.00")
                End With
              End If
              .MoveNext
            Wend
          End If
        ElseIf cboCrieria = "Like" Then
          lstSearch.ListItems.Clear
          .Open "select loanno,memberno,loanamt from loans where loanamt like '" & CCur(txtFrom.Text) & "%'", Cnn, adOpenKeyset, adLockOptimistic
          If .RecordCount > 0 Then
            .MoveFirst
            While Not .EOF
              With lstSearch
                Set Li = .ListItems.Add(, , Rst!LoanNo)
                Li.SubItems(1) = Rst!memberno
                Li.SubItems(2) = Format(Rst!loanamt, "###,###,###,##0.00")
              End With
              .MoveNext
            Wend
            lblRecords.Caption = lstSearch.ListItems.Count
          End If
        Else
          lstSearch.ListItems.Clear
          
          .Open "select loanno,memberno,loanamt from loans where loanamt " & cboCrieria.Text & " " & CCur(txtFrom.Text) & "", Cnn, adOpenKeyset, adLockOptimistic
          If .RecordCount > 0 Then
            While Not .EOF
              With lstSearch
                Set Li = .ListItems.Add(, , Rst!LoanNo)
                Li.SubItems(1) = Rst!memberno
                Li.SubItems(2) = Format(Rst!loanamt, "###,###,###,##0.00")
              End With
              .MoveNext
            Wend
          End If
        End If
      End If
      lblRecords.Caption = lstSearch.ListItems.Count
      .Close
    End With
    Set Rst = Nothing
  End If
End Sub

Private Sub txtFrom_Change()
  If txtFrom.Text = "" Then
    tbrSearch.Buttons(1).Enabled = False
  Else
    tbrSearch.Buttons(1).Enabled = True
  End If
End Sub
