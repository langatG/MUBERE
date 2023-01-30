VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Begin VB.Form frmtchprates 
   Caption         =   "TCHP RATES"
   ClientHeight    =   4155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4890
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   4890
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cbotype 
      Height          =   315
      ItemData        =   "frmtchprates.frx":0000
      Left            =   1320
      List            =   "frmtchprates.frx":000D
      TabIndex        =   11
      Text            =   "OTHERS"
      Top             =   3120
      Width           =   2415
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   3600
      TabIndex        =   6
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   3720
      Width           =   855
   End
   Begin VB.TextBox txtBName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   2640
      Width           =   3375
   End
   Begin VB.TextBox txtBCode 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   2040
      Width           =   3375
   End
   Begin MSComctlLib.ListView lvWBranch 
      Height          =   1695
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   2990
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
   Begin VB.Label Label3 
      Caption         =   "Type"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Rate"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   2640
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Code"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   375
   End
End
Attribute VB_Name = "frmtchprates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myclass As cdbase
Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmddelete_Click()
Set myclass = New cdbase
Provider = myclass.OpenCon
Set cn = CreateObject("adodb.connection")
cn.Open Provider, "atm", "atm"
sql = "delete from Tchp_Rate where Code='" & txtBCode & "'"
oSaccoMaster.ExecuteThis (sql)

myclass.Delete sql
loadBranchesTypes
txtBCode = ""
txtBName = ""
End Sub

Private Sub cmdedit_Click()
txtBCode.Locked = False
txtBName.Locked = False
cmdnew.Enabled = False
cmdEdit.Enabled = False
cmdsave.Enabled = True
End Sub

Private Sub cmdNew_Click()
txtBCode = ""
txtBName = ""
cboType = ""
txtBCode.Locked = False
txtBName.Locked = False
cmdnew.Enabled = False
cmdEdit.Enabled = False
cmdsave.Enabled = True
End Sub

Public Sub loadBranchesTypes()
    
    With lvWBranch
    
        .ListItems.Clear
        
        .ColumnHeaders.Clear
    
    End With

    Set rs2 = CreateObject("adodb.recordset")
    
    sql = "Select * from Tchp_Rate"
    
    Set rs2 = CreateObject("adodb.recordset")
    
    Set clsClass = New cdbase
    
    Provider = clsClass.OpenCon
    
    Set cn = CreateObject("adodb.connection")
    
   cn.Open Provider, "atm", "atm"
    
    rs2.Open sql, cn
    
    With lvWBranch
        
        .ColumnHeaders.Add , , "Code", 1400
        .ColumnHeaders.Add , , "Rate", 1500
        .ColumnHeaders.Add , , "Type", 2500
    
        While Not rs2.EOF
        
            Set li = .ListItems.Add(, , Trim(rs2.Fields("Code")))
            
            li.ListSubItems.Add , , Trim(rs2.Fields("Rate"))
             li.ListSubItems.Add , , IIf(IsNull(Trim(rs2.Fields("Type"))), "NA", rs2.Fields("type"))
        
            
            rs2.MoveNext
        
        Wend
        
    End With
    
    rs2.Close
    
    Set rs2 = Nothing
    
lvWBranch.View = lvwReport

End Sub

Private Sub cmdsave_Click()
On Error GoTo ErrorHandler

Set cn = New ADODB.Connection
sql = "d_sp_tchprates '" & txtBCode & "'," & txtBName & ",'" & cboType & "'"
oSaccoMaster.ExecuteThis (sql)
txtBCode = ""
txtBName = ""
txtBCode.Locked = True
txtBName.Locked = True
cmdnew.Enabled = True
cmdEdit.Enabled = False
cmdsave.Enabled = False
loadBranchesTypes
MsgBox "Records successively updated."
Exit Sub
ErrorHandler:
MsgBox err.description
End Sub

Private Sub Form_Load()
txtBCode.Locked = True
txtBName.Locked = True
cmdDelete.Enabled = False
loadBranchesTypes
End Sub

Public Sub edit(selected As String)

Set myclass = New cdbase
Set cn = CreateObject("adodb.connection")
Provider = myclass.OpenCon
cn.Open Provider, "atm", "atm"
Set rs = CreateObject("adodb.recordset")
sql = "select * from Tchp_Rate where Code='" & selected & "'"
rs.Open sql, cn
If Not rs.EOF Then
txtBCode = selected
txtBName = rs!rate
cboType = IIf(IsNull(rs!Type), "NA", rs!Type)
End If
cmdDelete.Enabled = True

End Sub
Private Sub lvWBranch_DblClick()
cmdEdit.Enabled = True
edit lvWBranch.SelectedItem
End Sub


