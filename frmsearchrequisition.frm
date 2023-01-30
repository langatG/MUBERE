VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmsearchrequisition 
   Caption         =   "Search Requisition"
   ClientHeight    =   4845
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6960
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   6960
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   4695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6855
      Begin VB.CommandButton cmdRef 
         Height          =   495
         Left            =   3840
         Picture         =   "frmsearchrequisition.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Refresh"
         Top             =   4080
         Width           =   495
      End
      Begin VB.ComboBox cboField 
         BackColor       =   &H00E0E0E0&
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
         Height          =   330
         ItemData        =   "frmsearchrequisition.frx":030A
         Left            =   120
         List            =   "frmsearchrequisition.frx":030C
         TabIndex        =   8
         Top             =   360
         Width           =   2055
      End
      Begin VB.ComboBox cboCrieria 
         BackColor       =   &H00E0E0E0&
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
         Height          =   330
         ItemData        =   "frmsearchrequisition.frx":030E
         Left            =   2280
         List            =   "frmsearchrequisition.frx":0327
         TabIndex        =   7
         Top             =   360
         Width           =   1455
      End
      Begin VB.Frame Frame2 
         Height          =   1215
         Left            =   3840
         TabIndex        =   3
         Top             =   120
         Width           =   2895
         Begin VB.CommandButton cmdFind 
            Enabled         =   0   'False
            Height          =   375
            Left            =   2400
            Picture         =   "frmsearchrequisition.frx":034B
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox txtFrom 
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
            Height          =   315
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   2175
         End
         Begin VB.TextBox txtTo 
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
            Height          =   315
            Left            =   120
            TabIndex        =   4
            Top             =   720
            Visible         =   0   'False
            Width           =   2175
         End
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "SELECT"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5400
         TabIndex        =   2
         Top             =   4065
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancel 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3240
         Picture         =   "frmsearchrequisition.frx":044D
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Cancel"
         Top             =   4080
         Width           =   495
      End
      Begin MSComctlLib.ListView lstSearch 
         Height          =   2535
         Left            =   120
         TabIndex        =   10
         ToolTipText     =   "Search Accounts Record"
         Top             =   1440
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   4471
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
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
      Begin VB.Label Label1 
         Caption         =   "Search Field"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Records Found"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label lblRecords 
         Height          =   255
         Left            =   1800
         TabIndex        =   12
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Criteria"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   11
         Top             =   120
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmsearchrequisition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SelectedDsn As String
Dim CConnect As MAZIWA.cdbase
Dim rst As Recordset
Dim rst1 As Recordset
Dim li As ListItem
Dim recordfound As String


Private Sub cboCrieria_Click()
  If cboCrieria.Text = "Between" Then
    txtTo.Visible = True
  Else
    txtTo.Visible = False
  End If
End Sub

Private Sub cboCrieria_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cboField_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmdCancel_Click()
    sel = ""
    frmsearchrequisition.Visible = False
    
End Sub

Private Sub cmdFind_Click()
Dim Find As Long
Dim li As ListItem
Dim Field As String
Set cn = New Connection
Set rst = New Recordset
Field = cboField.Text
'Find = Button.Index
lstSearch.ListItems.Clear
  
'If Find = 1 Then
'CConnect.cnnConnect

If Not cboField.Text = "" Then
    If Not cboCrieria.Text = "" Then
        If Not cboCrieria.Text = "Between" And Not cboCrieria.Text = "Like" Then
        If cboField.Text = "IName" Then
        sql = " SELECT     RNo, IName, Make, Description FROM         d_Requisition where " & cboField.Text & "" & cboCrieria.Text & "" & txtFrom.Text & ""
 
        Else
            sql = "SELECT     RNo, IName, Make, Description FROM         d_Requisition where  " & cboField.Text & "" & cboCrieria.Text & "'" & txtFrom.Text & "'"
            End If
           CConnect.Openrs
            
            With rs
                If .RecordCount > 0 Then
                    .MoveFirst
                    Do While Not .EOF
Set li = frmsearchrequisition.lstSearch.ListItems.Add(, , .Fields(0))
                li.SubItems(1) = .Fields(1) & ""

                
                        .MoveNext
                    Loop
                End If
            End With
            
            Set rs = Nothing
            
        ElseIf cboCrieria.Text = "Like" Then
         'If cboField.Text = "R_No" Then
            sql = "SELECT     RNo, IName, Make, Description FROM         d_Requisition where " & cboField & " LIKE '%" & txtFrom & "%' ORDER BY Names"
            'Else
            'end if
           CConnect.Openrs
            
            With rs
                If .RecordCount > 0 Then
                    .MoveFirst
                    Do While Not .EOF
                     
                        If Not .EOF Then
Set li = frmsearchrequisition.lstSearch.ListItems.Add(, , .Fields(0))
                li.SubItems(1) = .Fields(1) & ""

                            .MoveNext
                        End If
                        
                    Loop
                End If
            End With
            
            Set rs = Nothing
            
            
        Else
            If cboField.Text = "INAME" Then
                sql = "SELECT     RNo, IName, Make, Description FROM         d_Requisition where " & cboField.Text & " " & cboCrieria & " " & txtFrom.Text & " And " & txtTo.Text & ""
               CConnect.Openrs
            Else
                sql = "SELECT     RNo, IName, Make, Description FROM         d_Requisition where " & cboField.Text & " " & cboCrieria & " '" & txtFrom.Text & "' And '" & txtTo.Text & "'"
               CConnect.Openrs
            End If
            
            With rs
                If .RecordCount > 0 Then
                    .MoveFirst
                    Do While Not .EOF
Set li = frmsearchrequisition.lstSearch.ListItems.Add(, , .Fields(0))
                li.SubItems(1) = .Fields(1) & ""

                        .MoveNext
                    Loop
                End If
            End With
            
            Set rs = Nothing
            
        End If
        
    Else
        MsgBox "Select the search criteria.", vbExclamation
    End If
Else
    MsgBox "Select the search field.", vbExclamation
End If

'Set cnnPayroll = Nothing


End Sub

Private Sub cmdRef_Click()
    Call SRefresh
    
End Sub

Private Sub cmdSelect_Click()
    sel = ""
        If lstSearch.ListItems.Count > 0 Then
        sel = lstSearch.SelectedItem
        Me.Visible = False
        Continue = True
    Else
        MsgBox "No record selected.", vbExclamation
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    With frmsearchrequisition.lstSearch
        .ListItems.Clear
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "Requisition No", 1500
        .ColumnHeaders.Add , , "Item Names", 2800
        .ColumnHeaders.Add , , "Make", 2800
        .ColumnHeaders.Add , , "Description", 2800
        .View = lvwReport
        .GridLines = True
    End With
    With frmsearchrequisition.cboField
        .AddItem "RNO"
        .AddItem "INAME"
        .AddItem "MAKE"
        .AddItem "Description"

    End With
    Set CConnect = New MAZIWA.cdbase
    CConnect.cnnConnect
    sql = "SELECT     d_Approve1.RNo, d_Requisition.IName, d_Requisition.AuditDatetime, d_Requisition.TransDate, d_Requisition.CostCentre, "
    sql = sql & " d_Requisition.qnty FROM d_Approve1 INNER JOIN"
    sql = sql & " dbo.d_Requisition ON dbo.d_Approve1.RNo = dbo.d_Requisition.RNo AND dbo.d_Approve1.description = 'New'"
    'CConnect.Openrs
    Set rs = oSaccoMaster.GetRecordset(sql)
       
    With rs
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                Set li = frmsearchrequisition.lstSearch.ListItems.Add(, , .Fields(0))
                li.SubItems(1) = .Fields(1) & ""
               
                .MoveNext
            Loop
        End If
        .Close
    End With
    Set rs = Nothing
    'Set cnnPayroll = Nothing
    cboCrieria.Text = cboCrieria.List(0)
    cboField.Text = cboField.List(0)
    'txtFrom.SetFocus
    Me.Top = (Screen.Height - Height) / 2
    Me.Left = (Screen.Width - Width) / 1.4
End Sub



Private Sub lstSearch_DblClick()
    If lstSearch.ListItems.Count > 0 Then
        sel = lstSearch.SelectedItem.Text
        Continue = True
    End If
    Unload Me
End Sub

Private Sub txtFrom_Change()
    If txtFrom.Text = "" Then
        cmdFind.Enabled = False
    Else
        cmdFind.Enabled = True
    End If
    
End Sub





Public Sub SRefresh()
lstSearch.ListItems.Clear

    CConnect.cnnConnect
   sql = "SELECT     RNo, IName, Make, Description FROM         d_Requisition order by rno "
    CConnect.Openrs
    
       
    With rs
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                Set li = frmsearchrequisition.lstSearch.ListItems.Add(, , .Fields(0))
                li.SubItems(1) = .Fields(1) & ""
                
                .MoveNext
                
            Loop
            
        End If
        .Close
    End With
    
    Set rs = Nothing
    'Set cnnPayroll = Nothing
    
    txtFrom.Text = ""
    txtTo.Text = ""
    cboCrieria.Text = "="
    cboField.Text = "INAME"
End Sub

Private Sub txtFrom_KeyPress(KeyAscii As Integer)
    If Len(Trim(txtFrom.Text)) > 20 Then
        Beep
        MsgBox "Can't enter more than 20 characters", vbExclamation
        KeyAscii = 8
    End If
  Select Case KeyAscii
    'Case Asc("vbBack")
    Case Asc("A") To Asc("Z")
    Case Asc("a") To Asc("z")
    Case Asc("0") To Asc("9")
    Case Asc("/")
    Case Asc("-")
    Case Asc("(")
    Case Asc(")")
    Case Asc(" ")
    Case Asc(".")
    'Case Asc("'")
    Case Is = 8
    
    Case Else
    Beep
    KeyAscii = 0
  End Select
End Sub

Private Sub txtTo_KeyPress(KeyAscii As Integer)
    If Len(Trim(txtTo.Text)) > 20 Then
        Beep
        MsgBox "Can't enter more than 20 characters", vbExclamation
        KeyAscii = 8
    End If
  Select Case KeyAscii
    'Case Asc("vbBack")
    Case Asc("A") To Asc("Z")
    Case Asc("a") To Asc("z")
    Case Asc("0") To Asc("9")
    Case Asc("/")
    Case Asc("-")
    Case Asc("(")
    Case Asc(")")
    Case Asc(" ")
    Case Asc(".")
    Case Is = 8
    Case Else
    Beep
    KeyAscii = 0
  End Select
End Sub






