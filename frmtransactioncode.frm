VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Begin VB.Form frmtransactioncode 
   Caption         =   "TRANSACTION CODES"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6555
   Icon            =   "frmtransactioncode.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   6555
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdcancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   11
      Top             =   6120
      Width           =   735
   End
   Begin VB.CommandButton cmdnew 
      Caption         =   "&New"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   10
      Top             =   6120
      Width           =   735
   End
   Begin VB.CommandButton cmdedit 
      Caption         =   "&Edit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   9
      Top             =   6120
      Width           =   735
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   8
      Top             =   6120
      Width           =   735
   End
   Begin VB.CommandButton cmddelete 
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   7
      Top             =   6120
      Width           =   735
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   6
      Top             =   6120
      Width           =   735
   End
   Begin VB.Frame Fracategory 
      Caption         =   "Categories"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   1
      Top             =   4440
      Width           =   6135
      Begin VB.TextBox txtdescription 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1800
         TabIndex        =   3
         Top             =   960
         Width           =   2895
      End
      Begin VB.TextBox txtcid 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1800
         MaxLength       =   6
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Transaction Codes"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1695
      End
   End
   Begin MSComctlLib.ListView Lvwcategory 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   6800
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
End
Attribute VB_Name = "frmtransactioncode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim I As Object
Dim ed As Boolean
Dim myclass As cdbase
Public Sub edit(selected As String)
Set myclass = New cdbase
Set cn = CreateObject("adodb.connection")
Provider = myclass.OpenCon
cn.Open Provider, "atm", "atm"
Set rs = CreateObject("adodb.recordset")
sql = "select * from TransCode where transactioncode='" & selected & "'"
rs.Open sql, cn
If Not rs.EOF Then
txtcid = selected
txtdescription = rs!description
End If
cmdDelete.Enabled = True

End Sub

Private Sub cmdcancel_Click()
Form_Load
End Sub

Private Sub cmdclose_Click()
Unload Me
End Sub
Private Sub cmddelete_Click()
Set myclass = New cdbase
Provider = myclass.OpenCon
Set cn = CreateObject("adodb.connection")
cn.Open Provider, "atm", "atm"
sql = "delete from TransCode where transactioncode='" & txtcid & "'"

myclass.Delete sql


End Sub

Private Sub cmdedit_Click()
cmdsave.Enabled = True
ed = True
Set rs = CreateObject("adodb.recordset")
    
    sql = "Select * from TransCode"
    
    Set rs = CreateObject("adodb.recordset")
    
    Set clsClass = New cdbase
    
    Provider = clsClass.OpenCon
    
    Set cn = CreateObject("adodb.connection")
    
   cn.Open Provider, "atm", "atm"
    
    rs.Open sql, cn
    If Not rs.EOF Then
edit Lvwcategory.SelectedItem
End If
End Sub
Public Sub loadCategory()
    
    With Lvwcategory
    
        .ListItems.Clear
        
        .ColumnHeaders.Clear
    
    End With

    Set rs = CreateObject("adodb.recordset")
    
    sql = "Select * from TransCode"
    
    Set rs = CreateObject("adodb.recordset")
    
    Set clsClass = New cdbase
    
    Provider = clsClass.OpenCon
    
    Set cn = CreateObject("adodb.connection")
    
   cn.Open Provider, "atm", "atm"
    
    rs.Open sql, cn
    
    With Lvwcategory
        
        .ColumnHeaders.Add , , "transaction ID"
        .ColumnHeaders.Add , , "Description"
    
        While Not rs.EOF
        
            Set li = .ListItems.Add(, , Trim(rs.Fields("transactioncode")))
            
            li.ListSubItems.Add , , Trim(rs.Fields("Description"))
            
        
            
            rs.MoveNext
        
        Wend
        
    End With
    
    rs.Close
    
    Set rs = Nothing
    
Lvwcategory.View = lvwReport

End Sub
Private Sub cmdNew_Click()

For Each I In Controls
  If TypeOf I Is TextBox Then I.Enabled = True
Next
txtcid = ""
txtdescription = ""
cmdsave.Enabled = True

End Sub

Private Sub cmdsave_Click()

  Set myclass = New cdbase

    Provider = myclass.OpenCon

    Set cn = CreateObject("adodb.connection")

   cn.Open Provider, "atm", "atm"

    Set rs = CreateObject("ADODB.Recordset")
    
    sql = ""
    
   sql = "SELECT * from TransCode WHERE transactioncode='" & txtcid & "'"
   
  ' MsgBox "you cann't have two savings account", vbInformation, "FOSA": Exit Sub
    rs.Open sql, cn
If ed = True Then '// update the bank details
        sql = "update TransCode "
        sql = sql & " set Description='" & txtdescription
        sql = sql & "'  where transactioncode='" & txtcid & "'"
        
        myclass.save sql

Else
'// check if all the it already exist
        If Not rs.EOF Then
            
            MsgBox "The Transaction Code already exist Please input a new one.", vbInformation, "Sets categories"
            txtcid.SetFocus
            Exit Sub
          
        End If
       
    

Set myclass = New cdbase

sql = "insert into TransCode(transactioncode,Description)values('" & txtcid & "','" & txtdescription & "')"

myclass.save sql
End If


Form_Load
 
End Sub
Private Sub txtcid_KeyPress(KeyAscii As Integer)
If ValidChar(KeyAscii) = False Then KeyAscii = 0
End Sub
Private Sub txtDescription_KeyPress(KeyAscii As Integer)
If ValidChar(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub Form_Load()
For Each I In Controls
 If TypeOf I Is TextBox Then I.Text = ""
Next
 cmdsave.Enabled = False
 cmdDelete.Enabled = False
 ed = False
 loadCategory
End Sub


