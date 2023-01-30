VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Begin VB.Form frmLocation 
   Caption         =   "Locations"
   ClientHeight    =   3495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   4950
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ListView lvWLocations 
      Height          =   1695
      Left            =   120
      TabIndex        =   9
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
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   3720
      TabIndex        =   8
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   2880
      TabIndex        =   7
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   2880
      Width           =   855
   End
   Begin VB.TextBox txtLName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Top             =   2400
      Width           =   2055
   End
   Begin VB.TextBox txtLCode 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Location Name"
      Height          =   195
      Left            =   480
      TabIndex        =   1
      Top             =   2400
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Location Code"
      Height          =   195
      Left            =   480
      TabIndex        =   0
      Top             =   1800
      Width           =   1035
   End
End
Attribute VB_Name = "frmLocation"
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
sql = "delete from d_Location where LCode='" & txtLCode & "'"

myclass.Delete sql
loadLocations

End Sub

Private Sub cmdedit_Click()
txtLCode.Locked = False
txtLName.Locked = False
cmdnew.Enabled = False
cmdEdit.Enabled = False
End Sub

Private Sub cmdNew_Click()
txtLCode = ""
txtLName = ""
txtLCode.Locked = False
txtLName.Locked = False
cmdnew.Enabled = False
cmdEdit.Enabled = False

End Sub
Public Sub loadLocations()

    With lvWLocations
    
       .ListItems.Clear
    
        .ColumnHeaders.Clear

  End With

    Set rs = CreateObject("adodb.recordset")
  
    sql = "Select * from d_Location"
    
    Set rs = CreateObject("adodb.recordset")
   
    Set clsClass = New cdbase
    
    Provider = clsClass.OpenCon
    
    Set cn = CreateObject("adodb.connection")
    
   cn.Open Provider, "atm", "atm"
    
    rs.Open sql, cn
    
    With lvWLocations
        
        .ColumnHeaders.Add , , "Location Code"
        .ColumnHeaders.Add , , "Location Name"
    
        While Not rs.EOF
        
            Set li = .ListItems.Add(, , Trim(rs.Fields("LCode")))
            
            li.ListSubItems.Add , , Trim(rs.Fields("LName"))
            
                    rs.MoveNext
        
        Wend
        
    End With
    
    rs.Close
    
    Set rs = Nothing
    
lvWLocations.View = lvwReport

End Sub

Private Sub cmdsave_Click()
On Error GoTo ErrorHandler
If txtLCode = "" Then
MsgBox "Please enter the location code", vbInformation
Exit Sub
End If

Set cn = New ADODB.Connection
sql = "d_sp_Location '" & txtLCode & "','" & txtLName & "','" & User & "'"
oSaccoMaster.ExecuteThis (sql)
txtLCode = ""
txtLName = ""
txtLCode.Locked = True
txtLName.Locked = True
cmdnew.Enabled = True
cmdEdit.Enabled = True
loadLocations
MsgBox "Records successively updated."
Exit Sub
ErrorHandler:
MsgBox err.description

End Sub

Private Sub Form_Load()
txtLCode.Locked = True
txtLName.Locked = True
loadLocations
End Sub
Public Sub edit(selected As String)

Set myclass = New cdbase
Set cn = CreateObject("adodb.connection")
Provider = myclass.OpenCon
cn.Open Provider, "atm", "atm"
Set rs = CreateObject("adodb.recordset")
sql = "select * from d_Location where LCode='" & selected & "'"
rs.Open sql, cn
If Not rs.EOF Then
txtLCode = selected
txtLName = rs!LName
End If
cmdDelete.Enabled = True

End Sub
Private Sub lvWLocations_DblClick()
edit lvWLocations.SelectedItem
End Sub
