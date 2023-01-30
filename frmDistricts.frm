VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Begin VB.Form frmDistricts 
   Caption         =   "Districts"
   ClientHeight    =   3495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   4950
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtDCode 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      TabIndex        =   7
      Top             =   1920
      Width           =   2055
   End
   Begin VB.TextBox txtDName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      TabIndex        =   6
      Top             =   2520
      Width           =   2055
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   3000
      Width           =   855
   End
   Begin MSComctlLib.ListView lvWDistrict 
      Height          =   1695
      Left            =   120
      TabIndex        =   0
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "District Code"
      Height          =   195
      Left            =   480
      TabIndex        =   9
      Top             =   1920
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "District Name"
      Height          =   195
      Left            =   480
      TabIndex        =   8
      Top             =   2520
      Width           =   945
   End
End
Attribute VB_Name = "frmDistricts"
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
sql = "delete from d_Districts where DCode='" & txtdcode & "'"

myclass.Delete sql
loadDistrictTypes
txtdcode = ""
txtDName = ""
End Sub

Private Sub cmdedit_Click()
txtdcode.Locked = False
txtDName.Locked = False
cmdnew.Enabled = False
cmdEdit.Enabled = False
End Sub

Private Sub cmdNew_Click()
txtdcode = ""
txtDName = ""
txtdcode.Locked = False
txtDName.Locked = False
cmdnew.Enabled = False
cmdEdit.Enabled = False
End Sub

Public Sub loadDistrictTypes()
    
    With lvWDistrict
    
        .ListItems.Clear
        
        .ColumnHeaders.Clear
    
    End With

    Set rs2 = CreateObject("adodb.recordset")
    
    sql = "Select * from d_Districts"
    
    Set rs2 = CreateObject("adodb.recordset")
    
    Set clsClass = New cdbase
    
    Provider = clsClass.OpenCon
    
    Set cn = CreateObject("adodb.connection")
    
   cn.Open Provider, "atm", "atm"
    
    rs2.Open sql, cn
    
    With lvWDistrict
        
        .ColumnHeaders.Add , , "District Code"
        .ColumnHeaders.Add , , "District Name"
    
        While Not rs2.EOF
        
            Set li = .ListItems.Add(, , Trim(rs2.Fields("DCode")))
            
            li.ListSubItems.Add , , Trim(rs2.Fields("DName"))
            
        
            
            rs2.MoveNext
        
        Wend
        
    End With
    
    rs2.Close
    
    Set rs2 = Nothing
    
lvWDistrict.View = lvwReport

End Sub

Private Sub cmdsave_Click()
On Error GoTo ErrorHandler

Set cn = New ADODB.Connection
sql = "d_sp_District '" & txtdcode & "','" & txtDName & "','" & User & "'"
oSaccoMaster.ExecuteThis (sql)
txtdcode = ""
txtDName = ""
txtdcode.Locked = True
txtDName.Locked = True
cmdnew.Enabled = True
cmdEdit.Enabled = True
loadDistrictTypes
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
sql = "select * from d_Districts where DCode='" & selected & "'"
rs.Open sql, cn
If Not rs.EOF Then
txtdcode = selected
txtDName = rs!DName
End If
cmdDelete.Enabled = True

End Sub
Private Sub Form_Load()
txtdcode.Locked = True
txtDName.Locked = True
loadDistrictTypes
End Sub

Private Sub lvWDistrict_DblClick()
edit lvWDistrict.SelectedItem
End Sub
