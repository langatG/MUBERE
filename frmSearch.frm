VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Begin VB.Form frmSearch 
   BackColor       =   &H80000001&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7245
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   7245
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame Frame1 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6855
      Begin VB.CommandButton cmdRef 
         Height          =   495
         Left            =   3840
         Picture         =   "frmSearch.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   14
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
         ItemData        =   "frmSearch.frx":074C
         Left            =   120
         List            =   "frmSearch.frx":074E
         TabIndex        =   8
         Text            =   "P_Name"
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
         ItemData        =   "frmSearch.frx":0750
         Left            =   2280
         List            =   "frmSearch.frx":0766
         TabIndex        =   7
         Top             =   360
         Width           =   1455
      End
      Begin VB.Frame Frame2 
         Height          =   1215
         Left            =   3840
         TabIndex        =   4
         Top             =   120
         Width           =   2895
         Begin VB.CommandButton cmdFind 
            Enabled         =   0   'False
            Height          =   375
            Left            =   2400
            Picture         =   "frmSearch.frx":0787
            Style           =   1  'Graphical
            TabIndex        =   13
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
            TabIndex        =   6
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
            TabIndex        =   5
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
         Top             =   4080
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
         Picture         =   "frmSearch.frx":0889
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Cancel"
         Top             =   4080
         Width           =   495
      End
      Begin MSComctlLib.ListView lstSearch 
         Height          =   2535
         Left            =   120
         TabIndex        =   3
         Top             =   1440
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   4471
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
         TabIndex        =   12
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
         TabIndex        =   11
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label lblRecords 
         Height          =   255
         Left            =   1800
         TabIndex        =   10
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
         TabIndex        =   9
         Top             =   120
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public SelectedDsn As String

Dim rst As Recordset
Dim rst1 As Recordset
Dim li As ListItem
Dim recordfound As String
Dim CConnect As MAZIWA.CConnect

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

Private Sub cmdcancel_Click()
    sel = ""
    frmSearch.Visible = False
    
End Sub

Private Sub cmdFind_Click()
On Error GoTo ErrorHandler
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
            sql = "Select * from ag_products where " & cboField.Text & "" & cboCrieria.Text & "'" & txtFrom.Text & "'"
           CConnect.Openrs
            
            With rs
                If .RecordCount > 0 Then
                    .MoveFirst
                    Do While Not .EOF
                       Set li = frmSearch.lstSearch.ListItems.Add(, , !p_code)
                li.SubItems(1) = !p_name & ""
                li.SubItems(2) = !S_no & ""
                        
                        .MoveNext
                    Loop
                End If
            End With
            
            Set rs = Nothing
            
        ElseIf cboCrieria.Text = "Like" Then
        
            sql = "Select * from ag_products order by P_code"
           CConnect.Openrs
            
            With rs
                If .RecordCount > 0 Then
                    .MoveFirst
                    Do While Not .EOF
                    On Error Resume Next
                        .Find "" & cboField.Text & " " & cboCrieria.Text & " '" & txtFrom.Text & "%'", , adSearchForward
                        If Not .EOF Then
                          Set li = frmSearch.lstSearch.ListItems.Add(, , !p_code)
                              li.SubItems(1) = !p_name & ""
                              li.SubItems(2) = !S_no & ""
                            
                            .MoveNext
                        End If
                        
                    Loop
                End If
            End With
            
            Set rs = Nothing
            
            
        Else
            If cboField.Text = "Amount" Then
                sql = "select * from products where " & cboField.Text & " " & cboCrieria & " " & txtFrom.Text & " And " & txtTo.Text & ""
               CConnect.Openrs
            Else
                sql = "select * from products where " & cboField.Text & " " & cboCrieria & " '" & txtFrom.Text & "' And '" & txtTo.Text & "'"
               CConnect.Openrs
            End If
            
            With rs
                If .RecordCount > 0 Then
                    .MoveFirst
                    Do While Not .EOF
                 Set li = frmSearch.lstSearch.ListItems.Add(, , !p_code)
                         li.SubItems(1) = !p_name & ""
                         li.SubItems(2) = !S_no & ""
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

Set cnnPayroll = Nothing
Exit Sub
ErrorHandler:
MsgBox err.description

End Sub

Private Sub cmdRef_Click()
    Call SRefresh
    
End Sub

Private Sub cmdSelect_Click()
sel = ""
    If lstSearch.ListItems.Count > 0 Then
        sel = lstSearch.SelectedItem
        Me.Visible = False
       ' Me.Unload Me
    Else
        MsgBox "No record selected.", vbExclamation
    End If
Unload Me
End Sub

Private Sub Form_Load()
Set CConnect = New MAZIWA.CConnect

    DSource = "MAZIWA"

    With frmSearch.lstSearch
        .ListItems.Clear
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "Product Code", 1500
        .ColumnHeaders.Add , , "Product Name", 2800
        .ColumnHeaders.Add , , "Serial No", 2800
        .View = lvwReport
        .GridLines = True
    End With
    
    With frmSearch.cboField
        .AddItem "P_code"
        .AddItem "P_name"
        .AddItem "S_no"
    End With
    
    CConnect.cnnConnect
    sql = "Select * from ag_Products order by P_NAME"
    CConnect.Openrs
    
    With rs
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                Set li = frmSearch.lstSearch.ListItems.Add(, , !p_code)
                li.SubItems(1) = !p_name & ""
                li.SubItems(2) = !S_no & ""
                .MoveNext
                
            Loop
            
        End If
        .Close
    End With
    
    Set rs = Nothing
    Set cnnPayroll = Nothing
        
    cboCrieria.Text = cboCrieria.List(0)
    cboField.Text = cboField.List(0)
    cboField.Text = "P_NAME"
    'txtFrom.SetFocus
    
    Me.Top = (Screen.Height - Height) / 2
    Me.Left = (Screen.Width - Width) / 1.4
    
End Sub



Private Sub txtFrom_Change()
    If txtFrom.Text = "" Then
        cmdFind.Enabled = False
    Else
        cmdFind.Enabled = True
    End If
    cmdFind_Click
End Sub





Public Sub SRefresh()
lstSearch.ListItems.Clear

    CConnect.cnnConnect
    sql = "Select * from products order by P_code"
    CConnect.Openrs
    
       
    With rs
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
     Set li = frmSearch.lstSearch.ListItems.Add(, , !p_code)
                li.SubItems(1) = !p_name & ""
                li.SubItems(2) = !S_no & ""
                .MoveNext
                
            Loop
            
        End If
        .Close
    End With
    
    Set rs = Nothing
    Set cnnPayroll = Nothing
    
    txtFrom.Text = ""
    txtTo.Text = ""
    cboCrieria.Text = "="
    cboField.Text = "PayrollNo"
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
