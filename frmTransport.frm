VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Begin VB.Form frmTransport 
   Caption         =   "Transport"
   ClientHeight    =   4155
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4935
   Icon            =   "frmTransport.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4155
   ScaleWidth      =   4935
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   3480
      TabIndex        =   6
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox txtTransName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   600
      Width           =   2055
   End
   Begin VB.TextBox txtTCode 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin MSComctlLib.ListView lvWTransport 
      Height          =   2415
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   4260
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "transcode"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Transport Mode"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   600
      Width           =   1125
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Transport Code"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmTransport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myclass As cdbase
Dim ed As Boolean
Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmddelete_Click()
Set myclass = New cdbase
Provider = myclass.OpenCon
Set cn = CreateObject("adodb.connection")
cn.Open Provider, "atm", "atm"
sql = "delete from d_TransMode where TransCode='" & txtTCode & "'"

myclass.Delete sql
txtTCode = ""
txtTransName = ""
txtTCode.Locked = True
txtTransName.Locked = True
cmdnew.Enabled = True
cmdEdit.Enabled = True

loadTransportModes

End Sub

Private Sub cmdedit_Click()
txtTCode.Locked = False
txtTransName.Locked = False
cmdnew.Enabled = False
cmdEdit.Enabled = False
cmdsave.Enabled = True
ed = True
Set rs = CreateObject("adodb.recordset")
    
    sql = "Select * from d_TransMode"
    
    Set rs = CreateObject("adodb.recordset")
    
    Set clsClass = New cdbase
    
    Provider = clsClass.OpenCon
    
    Set cn = CreateObject("adodb.connection")
    
   cn.Open Provider, "atm", "atm"
    
    rs.Open sql, cn
    If Not rs.EOF Then
edit lvWTransport.SelectedItem
End If
End Sub

Private Sub cmdNew_Click()
txtTCode = ""
txtTransName = ""
txtTCode.Locked = False
txtTransName.Locked = False
cmdnew.Enabled = False
cmdEdit.Enabled = False

End Sub
Public Sub loadTransportModes()
    
    With lvWTransport
    
        .ListItems.Clear
        
        .ColumnHeaders.Clear
    
    End With

    Set rs = CreateObject("adodb.recordset")
    
    sql = "Select * from d_TransMode"
    
    Set rs = CreateObject("adodb.recordset")
    
    Set clsClass = New cdbase
    
    Provider = clsClass.OpenCon
    
    Set cn = CreateObject("adodb.connection")
    
   cn.Open Provider, "atm", "atm"
    
    rs.Open sql, cn
    
    With lvWTransport
        
        .ColumnHeaders.Add , , "Transport Code"
        .ColumnHeaders.Add , , "Transport Name"
    
        While Not rs.EOF
        
            Set li = .ListItems.Add(, , Trim(rs.Fields("TransCode")))
            
            li.ListSubItems.Add , , Trim(rs.Fields("Transport"))
            
        
            
            rs.MoveNext
        
        Wend
        
    End With
    
    rs.Close
    
    Set rs = Nothing
    
lvWTransport.View = lvwReport

End Sub
Public Sub edit(selected As String)

Set myclass = New cdbase
Set cn = CreateObject("adodb.connection")
Provider = myclass.OpenCon
cn.Open Provider, "atm", "atm"
Set rs = CreateObject("adodb.recordset")
sql = "select * from d_TransMode where TransCode='" & selected & "'"
rs.Open sql, cn
If Not rs.EOF Then
txtTCode = selected
txtTransName = rs!Transport
End If
cmdDelete.Enabled = True

End Sub
Private Sub cmdsave_Click()
On Error GoTo ErrorHandler

Set cn = New ADODB.Connection
sql = "d_sp_TransMode '" & txtTCode & "','" & txtTransName & "','" & User & "'"
oSaccoMaster.ExecuteThis (sql)
txtTCode = ""
txtTransName = ""
txtTCode.Locked = True
txtTransName.Locked = True
cmdnew.Enabled = True
cmdEdit.Enabled = True
loadTransportModes
MsgBox "Records successively updated."
Exit Sub
ErrorHandler:
MsgBox err.description
End Sub

Private Sub Form_Load()
txtTCode.Locked = True
txtTransName.Locked = True
loadTransportModes
End Sub

Private Sub lvWTransport_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
lvWTransport.Sorted = True
lvWTransport.SortKey = ColumnHeader.Index

End Sub

Private Sub lvWTransport_DblClick()
edit lvWTransport.SelectedItem

End Sub

