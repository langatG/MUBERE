VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form frmRangeSelection 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Range Selection"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4275
   Icon            =   "frmRangeSelection.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   4275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   2775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   4895
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Range Selector"
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblTo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblFrom"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdCancel"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdOK"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtTo"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtFrom"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdFindFrom"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdFindTo"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      Begin VB.CommandButton cmdFindTo 
         Height          =   345
         Left            =   3765
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1665
         Width           =   345
      End
      Begin VB.CommandButton cmdFindFrom 
         Height          =   345
         Left            =   3750
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   975
         Width           =   330
      End
      Begin VB.Frame Frame2 
         Height          =   25
         Left            =   0
         TabIndex        =   8
         Top             =   2160
         Width           =   4335
      End
      Begin VB.Frame Frame1 
         Height          =   25
         Left            =   0
         TabIndex        =   7
         Top             =   320
         Width           =   4215
      End
      Begin VB.TextBox txtFrom 
         Height          =   350
         Left            =   240
         TabIndex        =   1
         Top             =   960
         Width           =   3495
      End
      Begin VB.TextBox txtTo 
         Height          =   350
         Left            =   240
         TabIndex        =   3
         Top             =   1680
         Width           =   3495
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   350
         Left            =   600
         TabIndex        =   5
         Top             =   2280
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   350
         Left            =   2040
         TabIndex        =   6
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label lblFrom 
         AutoSize        =   -1  'True
         Caption         =   "From Member No."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   1260
      End
      Begin VB.Label lblTo 
         AutoSize        =   -1  'True
         Caption         =   "To Member No."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   1440
         Width           =   1110
      End
   End
End
Attribute VB_Name = "frmRangeSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdcancel_Click()
    cancelActionInvolvingRange = True
    Unload Me
End Sub

'To use the range form create a function called onOkOfRangeForm within your form to load reports
' Your form should also have subs onFormLoadOfSearchFrm,onRefreshOfSearchForm, onFindOfSearchFrmClick in your form module. Look at frmMemRegistration module
'when you load range form from your form set things like caption for lables and caption of form
Private Sub cmdFindFrom_Click()
On Error GoTo errFix
Set formCallingSearch = Me
theTextBoxOnRange = "txtFrom" 'for search form to know where to drop value
onFormLoadOfSearchFrm
 Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "Range Selection"
End Sub

Private Sub cmdFindTo_Click()
On Error GoTo errFix
theTextBoxOnRange = "txtTo" 'for search form to know where to drop value
onFormLoadOfSearchFrm
Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "Range Selection"
End Sub

Private Sub cmdOK_Click()
cancelActionInvolvingRange = False
formCallingRangeSelector.onOkOfRangeForm
End Sub
Public Sub searchSelect() 'Drop value on range form
On Error GoTo errFix
    If theTextBoxOnRange = "txtFrom" Then
        Txtfrom.Text = Sel
    ElseIf theTextBoxOnRange = "txtTo" Then
        Txtto.Text = Sel
    End If
    frmSearch.Visible = False
  Exit Sub
errFix:
    MsgBox Err.description, vbOKOnly, "Range Selection"
End Sub



Public Sub onRefreshOfSearchFrm()
    formCallingRangeSelector.onRefreshOfSearchFrm 'refresh search form depending on form from which you are printing reports(formCallingRangeSelector)
End Sub


Public Sub onFindOfSearchFrmClick()
 formCallingRangeSelector.onFindOfSearchFrmClick 'execute find on  search form depending on form from which you are printing reports(formCallingRangeSelector)
End Sub
Public Sub onFormLoadOfSearchFrm()
PositionForm frmSearch
formCallingRangeSelector.onFormLoadOfSearchFrm 'load search form depending on form from which you are printing reports(formCallingRangeSelector)
End Sub


Private Sub cmdOk_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdOK_Click
End If
End Sub

Private Sub Form_Load()
PositionForm Me
Txtto.Text = ""
Txtfrom.Text = ""
End Sub

Private Sub txtFrom_Change()
Txtfrom.Text = UCase(Txtfrom.Text)
Txtfrom.SelStart = Len(Txtfrom.Text)
End Sub

Private Sub txtFrom_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Txtto.SetFocus
End If
End Sub

Private Sub txtTo_Change()
Txtto.Text = UCase(Txtto.Text)
Txtto.SelStart = Len(Txtto.Text)
End Sub

Private Sub txtTo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdOk.SetFocus
End If
End Sub
