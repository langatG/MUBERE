VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Begin VB.Form frmContainer 
   Caption         =   "Container Details"
   ClientHeight    =   4140
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5010
   LinkTopic       =   "Form1"
   ScaleHeight     =   4140
   ScaleWidth      =   5010
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtCCode 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      TabIndex        =   7
      Top             =   120
      Width           =   2055
   End
   Begin VB.TextBox txtCName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      TabIndex        =   6
      Top             =   600
      Width           =   2055
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   375
      Left            =   960
      TabIndex        =   4
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   1080
      Width           =   855
   End
   Begin MSComctlLib.ListView lvWContainer 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   4260
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Container Code"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Container Name"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   1140
   End
End
Attribute VB_Name = "frmContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmddelete_Click()
Set myclass = New cdbase
Provider = myclass.OpenCon
Set cn = CreateObject("adodb.connection")
cn.Open Provider, "atm", "atm"
sql = "delete from d_CType where ContCode='" & txtCCode & "'"

myclass.Delete sql
loadContainerTypes
txtCCode = ""
txtCName = ""

End Sub

Private Sub cmdedit_Click()
txtCCode.Locked = False
txtCName.Locked = False
cmdnew.Enabled = False
cmdEdit.Enabled = False
End Sub

Private Sub cmdNew_Click()
txtCCode = ""
txtCName = ""
txtCCode.Locked = False
txtCName.Locked = False
cmdnew.Enabled = False
cmdEdit.Enabled = False

End Sub

Public Sub loadContainerTypes()
    
    With lvWContainer
    
        .ListItems.Clear
        
        .ColumnHeaders.Clear
    
    End With

    Set rs2 = CreateObject("adodb.recordset")
    
    sql = "Select * from d_CType"
    
    Set rs2 = CreateObject("adodb.recordset")
    
    Set clsClass = New cdbase
    
    Provider = clsClass.OpenCon
    
    Set cn = CreateObject("adodb.connection")
    
   cn.Open Provider, "atm", "atm"
    
    rs2.Open sql, cn
    
    With lvWContainer
        
        .ColumnHeaders.Add , , "Container Code"
        .ColumnHeaders.Add , , "Container Name"
    
        While Not rs2.EOF
        
            Set li = .ListItems.Add(, , Trim(rs2.Fields("ContCode")))
            
            li.ListSubItems.Add , , Trim(rs2.Fields("ContName"))
            
        
            
            rs2.MoveNext
        
        Wend
        
    End With
    
    rs2.Close
    
    Set rs2 = Nothing
    
lvWContainer.View = lvwReport

End Sub

Private Sub cmdsave_Click()
On Error GoTo ErrorHandler

Set cn = New ADODB.Connection
sql = "d_sp_CType '" & txtCCode & "','" & txtCName & "','" & User & "'"
oSaccoMaster.ExecuteThis (sql)
txtCCode = ""
txtCName = ""
txtCCode.Locked = True
txtCName.Locked = True
cmdnew.Enabled = True
cmdEdit.Enabled = True
loadContainerTypes
MsgBox "Records successively updated."
Exit Sub
ErrorHandler:
MsgBox err.description

End Sub
Public Sub edit(selected As String)

Set myclass = New cdbase
Set cn = CreateObject("adodb.connection")
Provider = myclass.OpenCon
cn.Open Provider, "atm", "atm"
Set rs = CreateObject("adodb.recordset")
sql = "select * from d_CType where ContCode='" & selected & "'"
rs.Open sql, cn
If Not rs.EOF Then
txtCCode = selected
txtCName = rs!ContName
End If
cmdDelete.Enabled = True

End Sub

Private Sub Form_Load()
txtCCode.Locked = True
txtCName.Locked = True
loadContainerTypes
End Sub

Private Sub lvWContainer_DblClick()
edit lvWContainer.SelectedItem
End Sub
