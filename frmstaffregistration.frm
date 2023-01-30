VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmstaffregistration 
   Caption         =   "Staff registration"
   ClientHeight    =   5820
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7965
   LinkTopic       =   "Form1"
   ScaleHeight     =   5820
   ScaleWidth      =   7965
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      Height          =   375
      Left            =   4320
      TabIndex        =   10
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton cmdedit 
      Caption         =   "Edit"
      Height          =   375
      Left            =   3360
      TabIndex        =   9
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
      Height          =   375
      Left            =   2280
      TabIndex        =   8
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton cmdnew 
      Caption         =   "New"
      Height          =   375
      Left            =   1200
      TabIndex        =   7
      Top             =   5160
      Width           =   1095
   End
   Begin MSComctlLib.ListView Lvwstaffs 
      Height          =   2895
      Left            =   360
      TabIndex        =   6
      Top             =   1920
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   5106
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComCtl2.DTPicker DTPregd 
      Height          =   375
      Left            =   5760
      TabIndex        =   5
      Top             =   600
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   122355713
      CurrentDate     =   42647
   End
   Begin VB.TextBox txtstaffname 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   840
      Width           =   3015
   End
   Begin VB.TextBox txtstaffnumber 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "Reg Date"
      Height          =   255
      Left            =   5760
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lblstname 
      Caption         =   "Staff Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Staff Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "frmstaffregistration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub loadstaffs()
    
    With Lvwstaffs
    
        .ListItems.Clear
        
        .ColumnHeaders.Clear
    
    End With

    Set rs = CreateObject("adodb.recordset")
    
    sql = "Select * from staffs"
    
    Set rs = CreateObject("adodb.recordset")
    
    Set clsClass = New cdbase
    
    Provider = clsClass.OpenCon
    
    Set cn = CreateObject("adodb.connection")
    
   cn.Open Provider, "atm", "atm"
    
    rs.Open sql, cn
    
    With Lvwstaffs
        
        .ColumnHeaders.Add , , "staffno"
        .ColumnHeaders.Add , , "staffname"
    
        While Not rs.EOF
        
            Set li = .ListItems.Add(, , Trim(rs.Fields("staffno")))
            
            li.ListSubItems.Add , , Trim(rs.Fields("staffname"))
            
        
            
            rs.MoveNext
        
        Wend
        
    End With
    
    rs.Close
    
    Set rs = Nothing
    
Lvwstaffs.View = lvwReport

End Sub
Private Sub cmdNew_Click()
cmdsave.Enabled = True
Dim I As Object
    For Each I In Controls
        If TypeOf I Is TextBox Then I.Text = ""
        
     Next
End Sub

Private Sub cmdsave_Click()
On Error GoTo ErrorHandler
If txtstaffnumber = "" Then
    MsgBox "please enter your staffnumber", vbInformation
    Exit Sub
End If

If txtstaffname = "" Then
    MsgBox "please enter your staffnumber", vbInformation
    Exit Sub
End If
Set cn = New ADODB.Connection
sql = "insert into staffs(staffno,staffname,date)values('" & txtstaffnumber & "','" & txtstaffname & "','" & DTPregd & "')"

oSaccoMaster.ExecuteThis (sql)
txtstaffnumber = ""
txtstaffname = ""
txtstaffnumber.Locked = True
txtstaffname.Locked = True
cmdnew.Enabled = True
cmdEdit.Enabled = True
loadstaffs
'MsgBox "Records successively updated."
Exit Sub
ErrorHandler:
MsgBox err.description
End Sub

Private Sub Form_Load()
Set cn = New ADODB.Connection
sql = "select * from staffs"
oSaccoMaster.ExecuteThis (sql)
txtstaffname = ""
txtstaffnumber = ""
'txtTCode.Locked = True
'txtTransName.Locked = True
cmdnew.Enabled = True
cmdEdit.Enabled = True
loadstaffs
'MsgBox "Records successively updated."
Exit Sub
ErrorHandler:
MsgBox err.description
End Sub
