VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Begin VB.Form frmlogout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Managing Login Users"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   5640
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLogOut 
      Caption         =   "Log Out"
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   2760
      Width           =   1335
   End
   Begin MSComctlLib.ListView lvwlogout 
      Height          =   2415
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   4260
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "User Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Work Station"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Logging Status"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmlogout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdLogOut_Click()
'Dim cnlog As New Connection
'If cnlog.State = adStateClosed Then
'cnlog.Open "fosa"
'End If
For I = 1 To lvwlogout.ListItems.Count
    If lvwlogout.ListItems(I).Checked = True Then
        Dim RSUPLOG As New Recordset, yy As String
        Set li = lvwlogout.ListItems(I)
                   sql = "SELECT * From UserAccounts where UserLoginIDs='" & User & "'"
                    Set rs = oSaccoMaster.GetRecordset(sql)
                     If Not rs.EOF Then
                     If rs!UserLoginIDs = li Then
                       MsgBox "Use another account", vbInformation
                       Unload Me
                     Exit Sub
                     End If
                  End If
          oSaccoMaster.Execute ("UPDATE LOGINS SET LOGedout='Yes' where UserLoginIDs='" & Trim(lvwlogout.ListItems(I).Text) & "'") 'and wkstation='" & yy & "'")
    End If
          
Next I
'lvwlogout.ListItems(I).clear
'End
Form_Load
End Sub

Private Sub Form_Load()
Dim rslogin As New ADODB.Recordset, log As String
lvwlogout.ListItems.Clear
Set rslogin = oSaccoMaster.GetRecordset("SELECT     UserLoginIDs, LogedOut, WkStation From LOGINS WHERE     (LogedOut = N'no')ORDER BY UserLoginIDs")
With rslogin
While Not .EOF
    Set li = lvwlogout.ListItems.Add(, , Trim(IIf(IsNull(!UserLoginIDs), "", !UserLoginIDs)))
    li.SubItems(1) = Trim(IIf(IsNull(!WkStation), "", !WkStation))
    li.SubItems(2) = Trim(IIf(IsNull(!LOGedout), "", !LOGedout))
    .MoveNext
Wend
End With

End Sub

