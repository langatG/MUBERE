VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Begin VB.Form frmBranch 
   Caption         =   "Branches"
   ClientHeight    =   3915
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5085
   Icon            =   "frmBranch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   5085
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   3000
      Width           =   855
   End
   Begin VB.TextBox txtBName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   2520
      Width           =   3375
   End
   Begin VB.TextBox txtBCode 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Top             =   1920
      Width           =   3375
   End
   Begin MSComctlLib.ListView lvWBranch 
      Height          =   1695
      Left            =   120
      TabIndex        =   7
      Top             =   120
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Branch Name"
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Branch Code"
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   1920
      Width           =   930
   End
End
Attribute VB_Name = "frmBranch"
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
sql = "delete from d_Branch where BCode='" & txtBCode & "'"

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
End Sub

Private Sub cmdNew_Click()
txtBCode = ""
txtBName = ""
txtBCode.Locked = False
txtBName.Locked = False
cmdnew.Enabled = False
cmdEdit.Enabled = False
sql = ""
sql = "select count(BCode) from d_Branch"
Set rs = oSaccoMaster.GetRecordset(sql)

If Not rs.EOF Then
txtBCode = rs.Fields(0) + 1
Else
txtBCode = 1
End If

End Sub

Public Sub loadBranchesTypes()
    
    With lvWBranch
    
        .ListItems.Clear
        
        .ColumnHeaders.Clear
    
    End With

    Set rs2 = CreateObject("adodb.recordset")
    
    sql = "Select * from d_Branch"
    
    Set rs2 = CreateObject("adodb.recordset")
    
    Set clsClass = New cdbase
    
    Provider = clsClass.OpenCon
    
    Set cn = CreateObject("adodb.connection")
    
   cn.Open Provider, "atm", "atm"
    
    rs2.Open sql, cn
    
    With lvWBranch
        
        .ColumnHeaders.Add , , "Branch Code"
        .ColumnHeaders.Add , , "Branch Name"
    
        While Not rs2.EOF
        
            Set li = .ListItems.Add(, , Trim(rs2.Fields("BCode")))
            
            li.ListSubItems.Add , , Trim(rs2.Fields("BName"))
            
        
            
            rs2.MoveNext
        
        Wend
        
    End With
    
    rs2.Close
    
    Set rs2 = Nothing
    
lvWBranch.View = lvwReport

End Sub

Private Sub cmdsave_Click()
On Error GoTo ErrorHandler
If txtBCode = "" Then
MsgBox "Enter the Branch Code", vbInformation
Exit Sub
End If
Set cn = New ADODB.Connection
sql = "d_sp_branch '" & txtBCode & "','" & txtBName & "','" & User & "'"
oSaccoMaster.ExecuteThis (sql)
txtBCode = ""
txtBName = ""
txtBCode.Locked = True
txtBName.Locked = True
cmdnew.Enabled = True
cmdEdit.Enabled = False
'cmdsave.Enabled = True
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
sql = "select * from d_Branch where BCode='" & selected & "'"
rs.Open sql, cn
If Not rs.EOF Then
txtBCode = selected
txtBName = rs!Bname
End If
cmdDelete.Enabled = True

End Sub
Private Sub lvWBranch_DblClick()
cmdEdit.Enabled = True
edit lvWBranch.SelectedItem
End Sub
