VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Begin VB.Form frmMainAccounts 
   Caption         =   "Main Accounts"
   ClientHeight    =   4200
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   5490
   LinkTopic       =   "Form1"
   ScaleHeight     =   4200
   ScaleWidth      =   5490
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboAccGroup 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmMainAccounts.frx":0000
      Left            =   1800
      List            =   "frmMainAccounts.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   11
      Top             =   2280
      Width           =   3375
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   3720
      Width           =   855
   End
   Begin VB.TextBox txtMName 
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
      Left            =   1800
      TabIndex        =   1
      Top             =   3240
      Width           =   3375
   End
   Begin VB.TextBox txtMCode 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Top             =   2760
      Width           =   3375
   End
   Begin MSComctlLib.ListView lvwMainAcc 
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
   Begin VB.Label Label3 
      Caption         =   "Header"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Account Name"
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   3360
      Width           =   1065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Main account Code"
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   2760
      Width           =   1395
   End
End
Attribute VB_Name = "frmMainAccounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myclass As cdbase

Private Sub cboAccGroup_Validate(Cancel As Boolean)
If Trim(cboAccGroup) = "" Then
Exit Sub
End If
Set rs = oSaccoMaster.GetRecordset("SELECT HCode FROM d_Headers WHERE HName='" & cboAccGroup & "'")
If Not IsNull(rs.Fields(0)) Then txtMCode = Trim(rs.Fields(0))

End Sub

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmddelete_Click()
Set myclass = New cdbase
Provider = myclass.OpenCon
Set cn = CreateObject("adodb.connection")
cn.Open Provider, "atm", "atm"
sql = "delete from d_MainAccount where MCode='" & txtMCode & "'"

myclass.Delete sql
loadBranchesTypes
txtMCode = ""
txtMName = ""
End Sub

Private Sub cmdedit_Click()
txtMCode.Locked = False
txtMName.Locked = False
cmdnew.Enabled = False
cmdEdit.Enabled = False
End Sub

Private Sub cmdNew_Click()
txtMCode = ""
txtMName = ""
txtMCode.Locked = False
txtMName.Locked = False
cmdnew.Enabled = False
cmdEdit.Enabled = False
cboAccGroup_Validate True
End Sub

Public Sub loadBranchesTypes()
    
    With lvwMainAcc
    
        .ListItems.Clear
        
        .ColumnHeaders.Clear
    
    End With

    Set rs2 = CreateObject("adodb.recordset")
    
    sql = "Select * from d_MainAccount order by Id"
    
    Set rs2 = CreateObject("adodb.recordset")
    
    Set clsClass = New cdbase
    
    Provider = clsClass.OpenCon
    
    Set cn = CreateObject("adodb.connection")
    
   cn.Open Provider, "atm", "atm"
    
    rs2.Open sql, cn
    
    With lvwMainAcc
        
        .ColumnHeaders.Add , , "Header Code"
        .ColumnHeaders.Add , , "Header Name"
    
        While Not rs2.EOF
        
            Set li = .ListItems.Add(, , Trim(rs2.Fields("MCode")))
            
            li.ListSubItems.Add , , Trim(rs2.Fields("MName"))
            
        
            
            rs2.MoveNext
        
        Wend
        
    End With
    
    rs2.Close
    
    Set rs2 = Nothing
    
lvwMainAcc.View = lvwReport

End Sub

Private Sub cmdsave_Click()
On Error GoTo ErrorHandler
Dim Header As String

If Trim(cboAccGroup) = "" Then
Exit Sub
End If

Set rs = oSaccoMaster.GetRecordset("SELECT HCode FROM d_Headers WHERE HName='" & cboAccGroup & "'")
If Not IsNull(rs.Fields(0)) Then Header = Trim(rs.Fields(0))

Set cn = New ADODB.Connection
sql = "d_sp_MainAccount '" & txtMCode & "','" & txtMName & "','" & User & "','" & Header & "'"
oSaccoMaster.ExecuteThis (sql)
txtMCode = ""
'txtHName = ""
'txtHCode.Locked = True
'txtHName.Locked = True
cmdnew.Enabled = True
cmdEdit.Enabled = False
loadBranchesTypes
MsgBox "Records successively updated."
cboAccGroup_Validate True
Exit Sub
ErrorHandler:
MsgBox err.description
End Sub

Private Sub Form_Load()

Set rs = oSaccoMaster.GetRecordset("SELECT HName FROM d_Headers order by ID")
While Not rs.EOF
If Not IsNull(rs.Fields(0)) Then cboAccGroup.AddItem (rs.Fields(0))

    rs.MoveNext
    Wend

txtMCode.Locked = True
txtMName.Locked = True
cmdDelete.Enabled = False
loadBranchesTypes
End Sub

Public Sub edit(selected As String)

Set myclass = New cdbase
Set cn = CreateObject("adodb.connection")
Provider = myclass.OpenCon
cn.Open Provider, "atm", "atm"
Set rs = CreateObject("adodb.recordset")
sql = "select * from d_MainAccount where MCode='" & selected & "'"
rs.Open sql, cn
If Not rs.EOF Then
txtMCode = selected
txtMName = rs!mName
End If
cmdDelete.Enabled = True

End Sub

Private Sub lvwHeaders_BeforeLabelEdit(Cancel As Integer)
cmdEdit.Enabled = True
edit lvwMainAcc.SelectedItem
End Sub

Private Sub lvwMainAcc_DblClick()
cmdEdit.Enabled = True
edit lvwMainAcc.SelectedItem
End Sub

Private Sub txtMName_Validate(Cancel As Boolean)
txtMName = UCase(txtMName)
End Sub

