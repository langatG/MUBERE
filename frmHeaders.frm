VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Begin VB.Form frmHeaders 
   Caption         =   "Account Headers"
   ClientHeight    =   3525
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   5025
   LinkTopic       =   "Form1"
   ScaleHeight     =   3525
   ScaleWidth      =   5025
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtHCode 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      MaxLength       =   2
      TabIndex        =   6
      Top             =   2040
      Width           =   3375
   End
   Begin VB.TextBox txtHName 
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "Text"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   285
      Left            =   1320
      TabIndex        =   5
      Top             =   2520
      Width           =   3375
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   3000
      Width           =   855
   End
   Begin MSComctlLib.ListView lvwHeaders 
      Height          =   1695
      Left            =   0
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Header Code"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   945
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Header"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   2520
      Width           =   525
   End
End
Attribute VB_Name = "frmHeaders"
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
sql = "delete from d_Headers where HCode='" & txtHCode & "'"

myclass.Delete sql
loadBranchesTypes
txtHCode = ""
txtHName = ""
End Sub

Private Sub cmdedit_Click()
txtHCode.Locked = False
txtHName.Locked = False
cmdnew.Enabled = False
cmdEdit.Enabled = False
End Sub

Private Sub cmdNew_Click()
txtHCode = ""
txtHName = ""
txtHCode.Locked = False
txtHName.Locked = False
cmdnew.Enabled = False
cmdEdit.Enabled = False
End Sub

Public Sub loadBranchesTypes()
    
    With lvwHeaders
    
        .ListItems.Clear
        
        .ColumnHeaders.Clear
    
    End With

    Set rs2 = CreateObject("adodb.recordset")
    
    sql = "Select * from d_Headers order by Id"
    
    Set rs2 = CreateObject("adodb.recordset")
    
    Set clsClass = New cdbase
    
    Provider = clsClass.OpenCon
    
    Set cn = CreateObject("adodb.connection")
    
   cn.Open Provider, "atm", "atm"
    
    rs2.Open sql, cn
    
    With lvwHeaders
        
        .ColumnHeaders.Add , , "Header Code"
        .ColumnHeaders.Add , , "Header Name"
    
        While Not rs2.EOF
        
            Set li = .ListItems.Add(, , Trim(rs2.Fields("HCode")))
            
            li.ListSubItems.Add , , Trim(rs2.Fields("HName"))
            
        
            
            rs2.MoveNext
        
        Wend
        
    End With
    
    rs2.Close
    
    Set rs2 = Nothing
    
lvwHeaders.View = lvwReport

End Sub

Private Sub cmdsave_Click()
On Error GoTo ErrorHandler

Set cn = New ADODB.Connection
sql = "d_sp_Header '" & txtHCode & "','" & txtHName & "','" & User & "'"
oSaccoMaster.ExecuteThis (sql)
txtHCode = ""
'txtHName = ""
'txtHCode.Locked = True
'txtHName.Locked = True
cmdnew.Enabled = True
cmdEdit.Enabled = False
loadBranchesTypes
MsgBox "Records successively updated."

Exit Sub
ErrorHandler:
MsgBox err.description
End Sub

Private Sub Form_Load()
txtHCode.Locked = True
txtHName.Locked = True
cmdDelete.Enabled = False
loadBranchesTypes
End Sub

Public Sub edit(selected As String)

Set myclass = New cdbase
Set cn = CreateObject("adodb.connection")
Provider = myclass.OpenCon
cn.Open Provider, "atm", "atm"
Set rs = CreateObject("adodb.recordset")
sql = "select * from d_Headers where HCode='" & selected & "'"
rs.Open sql, cn
If Not rs.EOF Then
txtHCode = selected
txtHName = rs!HName
End If
cmdDelete.Enabled = True

End Sub

Private Sub lvwHeaders_BeforeLabelEdit(Cancel As Integer)
cmdEdit.Enabled = True
edit lvwHeaders.SelectedItem
End Sub



Private Sub lvwHeaders_DblClick()
cmdEdit.Enabled = True
edit lvwHeaders.SelectedItem
End Sub

Private Sub txtHName_Validate(Cancel As Boolean)
txtHName = UCase(txtHName)
End Sub
