VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Begin VB.Form frmMilkTestSettings 
   BackColor       =   &H0080FF80&
   Caption         =   "Milk Test Settings"
   ClientHeight    =   6660
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6060
   BeginProperty Font 
      Name            =   "Garamond"
      Size            =   15.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   6060
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Rejection Reasons Set Up"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      Begin VB.CommandButton cmddelete 
         Caption         =   "Delete"
         Height          =   360
         Left            =   3240
         TabIndex        =   22
         Top             =   3840
         Width           =   975
      End
      Begin VB.ComboBox cboOrganoleptic 
         Height          =   480
         ItemData        =   "frmMilkTestSettings.frx":0000
         Left            =   3120
         List            =   "frmMilkTestSettings.frx":000A
         TabIndex        =   20
         Top             =   2160
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.ComboBox cboAlcohol 
         Height          =   480
         ItemData        =   "frmMilkTestSettings.frx":0019
         Left            =   3120
         List            =   "frmMilkTestSettings.frx":0023
         TabIndex        =   19
         Top             =   2160
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   360
         Left            =   4200
         TabIndex        =   18
         Top             =   3840
         Width           =   975
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit"
         Height          =   360
         Left            =   1320
         TabIndex        =   17
         Top             =   3840
         Width           =   975
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "New"
         Height          =   360
         Left            =   360
         TabIndex        =   16
         Top             =   3840
         Width           =   975
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   360
         Left            =   2280
         TabIndex        =   15
         Top             =   3840
         Width           =   975
      End
      Begin VB.TextBox txtRejDescription 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2160
         TabIndex        =   14
         Top             =   1680
         Width           =   3135
      End
      Begin VB.TextBox txtRejPCValue 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   3120
         TabIndex        =   13
         Top             =   2160
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.ComboBox cboRejLact 
         Height          =   480
         ItemData        =   "frmMilkTestSettings.frx":003B
         Left            =   3120
         List            =   "frmMilkTestSettings.frx":00A5
         TabIndex        =   12
         Top             =   2160
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.ComboBox cboRejRez 
         Height          =   480
         ItemData        =   "frmMilkTestSettings.frx":0197
         Left            =   3120
         List            =   "frmMilkTestSettings.frx":01B0
         TabIndex        =   11
         Top             =   2160
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.ComboBox cboRejCriteria 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         ItemData        =   "frmMilkTestSettings.frx":01C9
         Left            =   1200
         List            =   "frmMilkTestSettings.frx":01D6
         TabIndex        =   10
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox txtRejReasons 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1080
         TabIndex        =   9
         Top             =   3240
         Width           =   4095
      End
      Begin VB.ComboBox cboRejType 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         ItemData        =   "frmMilkTestSettings.frx":01E5
         Left            =   2160
         List            =   "frmMilkTestSettings.frx":01F8
         TabIndex        =   8
         Top             =   1080
         Width           =   3135
      End
      Begin VB.TextBox txtRejId 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   480
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Milk Quality is Good"
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   1080
         TabIndex        =   23
         Top             =   2760
         Width           =   2775
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Value"
         Height          =   375
         Left            =   2400
         TabIndex        =   6
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Reason"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   3240
         Width           =   975
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Criteria"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Description"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Rejection Type"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Rejection Id"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1575
      End
   End
   Begin MSComctlLib.ListView lvWMilkQSettings 
      Height          =   1935
      Left            =   120
      TabIndex        =   21
      Top             =   4560
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   3413
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "transcode"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmMilkTestSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myclass As cdbase
Dim value As String
Private Sub cboRejType_Click()

If cboRejType.Text = "Lactometer" Then
    cboRejLact.Visible = True
    cboRejRez.Visible = False
    txtRejPCValue.Visible = False
    cboAlcohol.Visible = False
    cboOrganoleptic.Visible = False
    cboRejCriteria.Enabled = True
    value = cboRejLact.Text
    
Else
If cboRejType.Text = "Rezasurin" Then
    cboRejLact.Visible = False
    cboRejRez.Visible = True
    txtRejPCValue.Visible = False
    cboAlcohol.Visible = False
    cboOrganoleptic.Visible = False
    cboRejCriteria.Enabled = True
    value = cboRejRez.Text

Else
If cboRejType.Text = "Plate Count" Then
    cboRejLact.Visible = False
    cboRejRez.Visible = False
    txtRejPCValue.Visible = True
    cboAlcohol.Visible = False
    cboOrganoleptic.Visible = False
    cboRejCriteria.Enabled = True
    value = txtRejPCValue

Else
If cboRejType.Text = "Alcohol" Then
    cboRejLact.Visible = False
    cboRejRez.Visible = False
    txtRejPCValue.Visible = False
    cboAlcohol.Visible = True
    cboOrganoleptic.Visible = False
    cboRejCriteria.Enabled = False
    value = cboAlcohol.Text

Else
If cboRejType.Text = "Organoleptic" Then
    cboRejLact.Visible = False
    cboRejRez.Visible = False
    txtRejPCValue.Visible = False
    cboAlcohol.Visible = False
    cboOrganoleptic.Visible = True
    cboRejCriteria.Enabled = False
    value = cboOrganoleptic.Text

End If
End If
End If
End If
End If
End Sub

Private Sub Command2_Click()
End Sub

Private Sub Command4_Click()
End Sub

Private Sub cmdclose_Click()
Unload Me

End Sub

Private Sub cmddelete_Click()
If txtRejId = "" Then
MsgBox "No reasons to be deleted", vbInformation
Exit Sub
End If
Set myclass = New cdbase
sql = ""
sql = "delete from d_M_QSettings where rejid='" & txtRejId & "'"
myclass.Delete sql

txtRejId = ""
txtRejDescription = ""
txtRejReasons = ""
cboRejCriteria = ""
cboRejLact = ""
cboRejRez = ""
cboRejType = ""
cboOrganoleptic = ""
txtRejPCValue.Locked = True
txtRejId.Locked = True
txtRejDescription.Locked = True
txtRejReasons.Locked = True
cboRejCriteria.Locked = True
cboRejLact.Locked = True
cboRejRez.Locked = True
cboRejType.Locked = True
cboOrganoleptic.Locked = True
txtRejPCValue.Locked = True
cmdnew.Enabled = True
cmdEdit.Enabled = True

loadMilkQSettings
End Sub

Private Sub cmdedit_Click()
txtRejPCValue.Locked = False
txtRejId.Locked = False
txtRejDescription.Locked = False
txtRejReasons.Locked = False
cboRejCriteria.Locked = False
cboRejLact.Locked = False
cboRejRez.Locked = False
cboRejType.Locked = False
cboOrganoleptic.Locked = False
txtRejPCValue.Locked = False
cmdnew.Enabled = False
cmdEdit.Enabled = False
End Sub

Private Sub cmdNew_Click()
txtRejId = ""
txtRejDescription = ""
txtRejReasons = ""
cboRejCriteria = ""
cboRejLact = ""
cboRejRez = ""
cboRejType = ""
cboOrganoleptic = ""
txtRejPCValue.Locked = False
txtRejId.Locked = False
txtRejDescription.Locked = False
txtRejReasons.Locked = False
cboRejCriteria.Locked = False
cboRejLact.Locked = False
cboRejRez.Locked = False
cboRejType.Locked = False
cboOrganoleptic.Locked = False
txtRejPCValue.Locked = False
cmdnew.Enabled = False
cmdEdit.Enabled = False
End Sub

Private Sub cmdsave_Click()
On Error GoTo ErrorHandler
cboRejType_Click
Set cn = New ADODB.Connection
sql = "d_sp_MilkQ'" & txtRejId & "','" & cboRejType.Text & "','" & txtRejDescription & "','" & cboRejCriteria.Text & "','" & value & "','" & txtRejReasons & "','" & User & "'"
oSaccoMaster.ExecuteThis (sql)
txtRejId = ""
txtRejDescription = ""
txtRejReasons = ""
cboRejCriteria = ""
cboRejLact = ""
cboRejRez = ""
cboRejType = ""
cboOrganoleptic = ""
txtRejPCValue.Locked = True
txtRejId.Locked = True
txtRejDescription.Locked = True
txtRejReasons.Locked = True
cboRejCriteria.Locked = True
cboRejLact.Locked = True
cboRejRez.Locked = True
cboRejType.Locked = True
cboOrganoleptic.Locked = True
txtRejPCValue.Locked = True
cmdnew.Enabled = True
cmdEdit.Enabled = True
loadMilkQSettings
MsgBox "Records successively updated."
Exit Sub
ErrorHandler:
MsgBox err.description


End Sub
Public Sub loadMilkQSettings()
    
    With lvWMilkQSettings
    
        .ListItems.Clear
        
        .ColumnHeaders.Clear
    
    End With

    Set rs = CreateObject("adodb.recordset")
    
    sql = "Select * from d_M_QSettings"
    
    Set rs = CreateObject("adodb.recordset")
    
    Set clsClass = New cdbase
    
    Provider = clsClass.OpenCon
    
    Set cn = CreateObject("adodb.connection")
    
   cn.Open Provider, "atm", "atm"
    
    rs.Open sql, cn
    
    With lvWMilkQSettings
        
        .ColumnHeaders.Add , , "Rejection Id"
        .ColumnHeaders.Add , , "Rejection Type"
        .ColumnHeaders.Add , , "Description"
        .ColumnHeaders.Add , , "Criteria"
        .ColumnHeaders.Add , , "Value"
        .ColumnHeaders.Add , , "Reason"
    
        While Not rs.EOF
        
            Set li = .ListItems.Add(, , Trim(rs.Fields("RejId")))
            
            li.ListSubItems.Add , , Trim(rs.Fields("Type"))
            li.ListSubItems.Add , , Trim(rs.Fields("Description"))
            li.ListSubItems.Add , , Trim(rs.Fields("Criteria"))
            li.ListSubItems.Add , , Trim(rs.Fields("dvalue"))
            li.ListSubItems.Add , , Trim(rs.Fields("Reasons"))
            
        
            
            rs.MoveNext
        
        Wend
        
    End With
    
    rs.Close
    
    Set rs = Nothing
    
lvWMilkQSettings.View = lvwReport

End Sub
Private Sub Form_Load()
txtRejPCValue.Locked = True
txtRejId.Locked = True
txtRejDescription.Locked = True
txtRejReasons.Locked = True
cboRejCriteria.Locked = True
cboRejLact.Locked = True
cboRejRez.Locked = True
cboRejType.Locked = True
cboOrganoleptic.Locked = True
txtRejPCValue.Locked = True
loadMilkQSettings

End Sub
Public Sub edit(selected As String)

Set myclass = New cdbase
Set cn = CreateObject("adodb.connection")
Provider = myclass.OpenCon
cn.Open Provider, "atm", "atm"
Set rs = CreateObject("adodb.recordset")
sql = "select * from d_M_QSettings where RejId='" & selected & "'"
rs.Open sql, cn
If Not rs.EOF Then
txtRejId = selected
txtRejDescription = rs!description
cboRejType = rs!Type
cboRejCriteria.Text = rs!Criteria
txtRejReasons = rs!Reasons
cboRejType_Click
If cboAlcohol.Visible = True Then
cboAlcohol.Text = rs!dvalue
Else
If cboOrganoleptic.Visible = True Then
cboOrganoleptic.Text = rs!dvalue
Else
If cboRejRez.Visible = True Then
cboRejRez.Text = rs!dvalue
Else
If cboRejLact.Visible = True Then
cboRejLact.Text = rs!dvalue
Else
If txtRejPCValue.Visible = True Then
txtRejPCValue = rs!dvalue
End If
End If
End If
End If
End If
End If
'cmdDelete.Enabled = True

End Sub

Private Sub lvWMilkQSettings_DblClick()
edit lvWMilkQSettings.SelectedItem
End Sub
