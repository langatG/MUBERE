VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Begin VB.Form frmtypes 
   Caption         =   "Types"
   ClientHeight    =   3585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4860
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   4860
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtBCode 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      TabIndex        =   6
      Top             =   2040
      Width           =   3375
   End
   Begin VB.TextBox txtBName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      TabIndex        =   5
      Top             =   2640
      Width           =   3375
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   3120
      Width           =   855
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Type Code"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   780
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Type Name"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   2640
      Width           =   825
   End
End
Attribute VB_Name = "frmtypes"
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
sql = "delete from d_Type where BCode='" & txtBCode & "'"

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
cmdsave.Enabled = True
End Sub

Public Sub loadBranchesTypes()
    
    With lvWBranch
    
        .ListItems.Clear
        
        .ColumnHeaders.Clear
    
    End With

    Set rs2 = CreateObject("adodb.recordset")
    
    sql = "Select * from d_Type"
    
    Set rs2 = CreateObject("adodb.recordset")
    
    Set clsClass = New cdbase
    
    Provider = clsClass.OpenCon
    
    Set cn = CreateObject("adodb.connection")
    
   cn.Open Provider, "atm", "atm"
    
    rs2.Open sql, cn
    
    With lvWBranch
        
        .ColumnHeaders.Add , , "Type Code"
        .ColumnHeaders.Add , , "Type Name"
    
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

Set cn = New ADODB.Connection
sql = "d_sp_Type '" & txtBCode & "','" & txtBName & "','" & User & "'"
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
sql = "select * from d_Type where BCode='" & selected & "'"
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

