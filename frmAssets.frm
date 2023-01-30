VERSION 5.00
Begin VB.Form frmAssets 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Assets Types"
   ClientHeight    =   2205
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   6030
   Icon            =   "frmAssets.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   6030
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtFields 
      Appearance      =   0  'Flat
      DataField       =   "rate"
      Height          =   285
      Index           =   2
      Left            =   2040
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   0
      ScaleHeight     =   540
      ScaleWidth      =   6030
      TabIndex        =   12
      Top             =   1125
      Width           =   6030
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   300
         Left            =   1213
         TabIndex        =   18
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   300
         Left            =   59
         TabIndex        =   3
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   300
         Left            =   4675
         TabIndex        =   17
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Print"
         Height          =   300
         Left            =   3521
         TabIndex        =   16
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   300
         Left            =   2367
         TabIndex        =   15
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   300
         Left            =   1213
         TabIndex        =   14
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   300
         Left            =   59
         TabIndex        =   13
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.PictureBox picStatBox 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   0
      ScaleHeight     =   540
      ScaleWidth      =   6030
      TabIndex        =   6
      Top             =   1665
      Width           =   6030
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4875
         Picture         =   "frmAssets.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4530
         Picture         =   "frmAssets.frx":0784
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   825
         Picture         =   "frmAssets.frx":0AC6
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   480
         Picture         =   "frmAssets.frx":0E08
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1170
         TabIndex        =   11
         Top             =   0
         Width           =   3360
      End
   End
   Begin VB.TextBox txtFields 
      Appearance      =   0  'Flat
      DataField       =   "AssetNAME"
      Height          =   285
      Index           =   1
      Left            =   2040
      TabIndex        =   1
      Top             =   360
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      Appearance      =   0  'Flat
      DataField       =   "AssetCODE"
      Height          =   285
      Index           =   0
      Left            =   2040
      TabIndex        =   0
      Top             =   60
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Depreciation Rate"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   19
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Asset Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   555
      TabIndex        =   5
      Top             =   375
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Asset Code:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   600
      TabIndex        =   4
      Top             =   60
      Width           =   1035
   End
End
Attribute VB_Name = "frmAssets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
    On Error GoTo 10
    Set rst = New Recordset
    Dim cn As Connection
    Set cn = New ADODB.Connection
    
    cn.Open frmODBCLogon.cboDSNList, "atm"
    rst.Open "select AssetCODE,AssetNAME,rate from assetcode", cn, adOpenStatic, adLockOptimistic
    Dim oText As TextBox
    'Bind the text boxes to the data provider
    For Each oText In Me.txtFields
        Set oText.DataSource = rst
    Next
   ' mbDataChanged = False
    Exit Sub
10:    MsgBox err.description
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  lblStatus.Width = Me.Width - 1500
  cmdNext.Left = lblStatus.Width + 700
  cmdLast.Left = cmdNext.Left + 340
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  'If mbEditFlag Or mbAddNewFlag Then Exit Sub

  Select Case KeyCode
    Case vbKeyEscape
      cmdclose_Click
    Case vbKeyEnd
      cmdLast_Click
    Case vbKeyHome
      cmdFirst_Click
    Case vbKeyUp, vbKeyPageUp
      If Shift = vbCtrlMask Then
        cmdFirst_Click
      Else
        cmdPrevious_Click
      End If
    Case vbKeyDown, vbKeyPageDown
      If Shift = vbCtrlMask Then
        cmdLast_Click
      Else
        cmdNext_Click
      End If
  End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub



Private Sub cmdAdd_Click()
  On Error GoTo AddErr
  With rst
    If Not (.BOF And .EOF) Then
      'mvBookMark = .Bookmark
    End If
    .AddNew
    lblStatus.Caption = "Add record"
   ' mbAddNewFlag = True
    SetButtons False
  End With

  Exit Sub
AddErr:
  MsgBox err.description
End Sub

Private Sub cmddelete_Click()
  On Error GoTo DeleteErr
  With rst
    .Delete
    .MoveNext
    If .EOF Then .MoveLast
  End With
  Exit Sub
DeleteErr:
  MsgBox err.description
End Sub

Private Sub cmdrefresh_Click()
  'This is only needed for multi user apps
  On Error GoTo RefreshErr
  'rptAssets.Show vbModal
  Exit Sub
RefreshErr:
  MsgBox err.description
End Sub

Private Sub cmdedit_Click()
  On Error GoTo EditErr

  lblStatus.Caption = "Edit record"
  'mbEditFlag = True
  SetButtons False
  Exit Sub

EditErr:
  MsgBox err.description
End Sub
Private Sub cmdcancel_Click()
  On Error Resume Next

  SetButtons True
  'mbEditFlag = False
  'mbAddNewFlag = False
  rst.CancelUpdate
  'If mvBookMark > 0 Then
   ' rst.Bookmark = mvBookMark
  'Else
    'rst.MoveFirst
 ' End If
  'mbDataChanged = False

End Sub

Private Sub cmdupdate_Click()
  On Error GoTo UpdateErr

  rst.UpdateBatch adAffectAll

  'If mbAddNewFlag Then
    'rst.MoveLast              'move to the new record
  'End If

 ' mbEditFlag = False
  'mbAddNewFlag = False
  SetButtons True
 ' mbDataChanged = False

  Exit Sub
UpdateErr:
  MsgBox err.description
End Sub

Private Sub cmdclose_Click()
  Unload Me
End Sub

Private Sub cmdFirst_Click()
  On Error GoTo GoFirstError

  rst.MoveFirst
 ' mbDataChanged = False

  Exit Sub

GoFirstError:
  MsgBox err.description
End Sub

Private Sub cmdLast_Click()
  On Error GoTo GoLastError

  rst.MoveLast
  'mbDataChanged = False

  Exit Sub

GoLastError:
  MsgBox err.description
End Sub

Private Sub cmdNext_Click()
  On Error GoTo GoNextError

  If Not rst.EOF Then rst.MoveNext
  If rst.EOF And rst.RecordCount > 0 Then
    Beep
     'moved off the end so go back
    rst.MoveLast
  End If
  'show the current record
  'mbDataChanged = False

  Exit Sub
GoNextError:
  MsgBox err.description
End Sub

Private Sub cmdPrevious_Click()
  On Error GoTo GoPrevError

  If Not rst.BOF Then rst.MovePrevious
  If rst.BOF And rst.RecordCount > 0 Then
    Beep
    'moved off the end so go back
    rst.MoveFirst
  End If
  'show the current record
 ' mbDataChanged = False

  Exit Sub

GoPrevError:
  MsgBox err.description
End Sub

Private Sub SetButtons(bVal As Boolean)
  cmdAdd.Visible = bVal
  cmdEdit.Visible = bVal
  cmdupdate.Visible = Not bVal
  cmdCancel.Visible = Not bVal
  cmdDelete.Visible = bVal
  cmdclose.Visible = bVal
  cmdRefresh.Visible = bVal
  cmdNext.Enabled = bVal
  cmdFirst.Enabled = bVal
  cmdLast.Enabled = bVal
  cmdPrevious.Enabled = bVal
End Sub

