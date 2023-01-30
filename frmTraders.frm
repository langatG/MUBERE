VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTraders 
   Caption         =   "Price Assignment"
   ClientHeight    =   5205
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7425
   LinkTopic       =   "Form1"
   ScaleHeight     =   5205
   ScaleWidth      =   7425
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPrice 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6240
      TabIndex        =   14
      Text            =   "0.00"
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   2280
      TabIndex        =   13
      Top             =   4560
      Width           =   735
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   3240
      TabIndex        =   12
      Top             =   4560
      Width           =   855
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   375
      Left            =   720
      TabIndex        =   11
      Top             =   4560
      Width           =   735
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFF80&
      Caption         =   "Suppliers Details"
      Height          =   5175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7455
      Begin VB.CheckBox chkActive 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Active"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5040
         TabIndex        =   20
         Top             =   4560
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.PictureBox Picture5 
         Height          =   255
         Left            =   960
         Picture         =   "frmTraders.frx":0000
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   18
         Top             =   480
         Width           =   255
      End
      Begin VB.CommandButton cmdedit 
         Caption         =   "Edit"
         Height          =   375
         Left            =   1560
         TabIndex        =   17
         Top             =   4560
         Width           =   615
      End
      Begin VB.ComboBox cboLocation 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1920
         TabIndex        =   4
         Top             =   1080
         Width           =   2415
      End
      Begin VB.TextBox txtId 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5040
         TabIndex        =   3
         Top             =   480
         Width           =   2295
      End
      Begin VB.TextBox txtNames 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   2
         Top             =   480
         Width           =   3615
      End
      Begin VB.TextBox txtSNo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   855
      End
      Begin MSComCtl2.DTPicker DTPRegDate 
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         _Version        =   393216
         Format          =   122355713
         CurrentDate     =   40096
      End
      Begin MSComctlLib.ListView lvWBranch2 
         Height          =   2775
         Left            =   120
         TabIndex        =   19
         Top             =   1680
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   4895
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
         BackStyle       =   0  'Transparent
         Caption         =   "Buying Price:"
         Height          =   375
         Left            =   4920
         TabIndex        =   16
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF80&
         Caption         =   "Location"
         Height          =   195
         Left            =   1920
         TabIndex        =   10
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF80&
         Caption         =   "Id Number"
         Height          =   195
         Left            =   5040
         TabIndex        =   9
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF80&
         Caption         =   "Date registered"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF80&
         Caption         =   "Names"
         Height          =   195
         Left            =   1200
         TabIndex        =   7
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF80&
         Caption         =   "Supplier:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      Caption         =   "Price (Per Kg)"
      Height          =   195
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "frmTraders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim newa As Integer
Dim cessapp As Integer

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdedit_Click()
newa = 0
txtId.Locked = False
txtNames.Locked = False
txtSNo.Locked = False
cbolocation.Locked = False
cmdsave.Enabled = True
End Sub

Private Sub cmdNew_Click()
'On Error GoTo ErrorHandler
newa = 1
txtId = ""
txtNames = ""
txtSNo = ""
cbolocation.Text = ""
txtPrice = "0.00"
txtId.Locked = False
txtNames.Locked = False
txtSNo.Locked = False
cbolocation.Locked = False
cmdEdit.Enabled = True
cmdsave.Enabled = True
'ErrorHandler:
'MsgBox err.description
End Sub

Private Sub cmdsave_Click()
Dim Active As String
On Error GoTo ErrorHandler

If txtSNo = "" Then
MsgBox "Please enter the Suppliers Namber ", vbInformation, "Missing Information"
txtSNo.SetFocus
Exit Sub
End If
If newa = 1 Then
Dim rss As ADODB.Recordset
sql = ""
sql = "select* from d_Price2 where sno=" & txtSNo & " "
Set rss = oSaccoMaster.GetRecordset(sql)
If Not rss.EOF Then
MsgBox "The supplier You have selected already in the List", vbInformation
Exit Sub
End If

If chkActive.value = vbChecked Then
    Active = "1"
Else
    Active = "0"
End If


'///////////

''Set cn = New ADODB.Connection
''sql = ""
''sql = "d_sp_Debtors2 '" & txtSNo & "','" & txtNames & "','" & txtId & "','" & cboLocation & "','" & DTPRegDate & "'," & CCur(txtPrice) & ""
''oSaccoMaster.ExecuteThis (sql)
''Else
''End If

sql = ""
sql = "set dateformat dmy insert into  d_Price2(Sno, Name, Price, Date, UserId, Branch,Active)"
sql = sql & "  values('" & txtSNo & "','" & txtNames & "'," & CCur(txtPrice) & ",'" & DTPRegDate & "','" & User & "','" & cbolocation & "','" & Active & "')"
cn.Execute sql

Else
sql = ""
sql = "set dateformat DMY update d_Price2 set Name='" & txtNames & "',Price=" & CCur(txtPrice) & ",Date=" & DTPRegDate & ",UserId='" & User & "',Branch='" & cbolocation & "',Active='" & Active & "' where Sno='" & txtSNo & "' "
cn.Execute sql

End If



cmdNew_Click
cmdsave.Enabled = False

MsgBox "Supplier Records successively updated."
loadpro
Exit Sub
ErrorHandler:
MsgBox err.description

End Sub
Private Sub Form_Load()

DTPRegDate = Format(Get_Server_Date, "dd/mm/yyyy")
DTPRegDate.MaxDate = Format(Get_Server_Date, "dd/mm/yyyy")
'DTPcomplaintperiod = DTPRegDate
loadpro
Dim myclass As cdbase
cessapp = 0
txtId.Locked = True
txtNames.Locked = True
cbolocation.Locked = True
cmdEdit.Enabled = False
cmdsave.Enabled = False

    
    Set myclass = New cdbase

    Provider = myclass.OpenCon
    
    Set cn = CreateObject("adodb.connection")
    
   cn.Open Provider, "atm", "atm"

Set rs = CreateObject("adodb.recordset")
    
    rs.Open "SELECT LName FROM d_Location", cn
    
    If rs.EOF Then Exit Sub
    
    With rs
        
        While Not .EOF
         
         If Not IsNull(rs.Fields("LName")) Then cbolocation.AddItem rs.Fields("LName")
         
         .MoveNext
        
        Wend
    
    End With
    
    
DTPRegDate = Format(Get_Server_Date, "dd/mm/yyyy")
End Sub
Public Sub loadpro()
    
    With lvWBranch2
    
        .ListItems.Clear
        
        .ColumnHeaders.Clear
    
    End With

    Set rs2 = CreateObject("adodb.recordset")
    
    sql = "set dateformat dmy Select Sno, Name, Date,Price, Branch from d_Price2 order by Date desc"
    
    Set rs2 = CreateObject("adodb.recordset")
    
    Set clsClass = New cdbase
    
    Provider = clsClass.OpenCon
    
    Set cn = CreateObject("adodb.connection")
    
   cn.Open Provider, "atm", "atm"
    
    rs2.Open sql, cn
    
    With lvWBranch2
        
        .ColumnHeaders.Add , , "Sno"
        .ColumnHeaders.Add , , "Name"
        .ColumnHeaders.Add , , "Date"
        .ColumnHeaders.Add , , "Price"
        '.ColumnHeaders.Add , , "R.Price"
        .ColumnHeaders.Add , , "Branch"
        While Not rs2.EOF
        
            Set li = .ListItems.Add(, , Trim(rs2.Fields("Sno")))
            
            'li.ListSubItems.Add , , Trim(rs2.Fields("Sno"))
            li.ListSubItems.Add , , Trim(rs2.Fields("Name"))
            li.ListSubItems.Add , , Trim(rs2.Fields("Date"))
            li.ListSubItems.Add , , Trim(rs2.Fields("Price"))
            li.ListSubItems.Add , , Trim(rs2.Fields("Branch"))
            rs2.MoveNext
        
        Wend
        
    End With
    
    rs2.Close
    
    Set rs2 = Nothing
    
lvWBranch2.View = lvwReport

End Sub

Private Sub Picture5_Click()
        Me.MousePointer = vbHourglass
        frmSearchSupplier.Show vbModal
        txtSNo = sel
'///////////
Dim rss As ADODB.Recordset
sql = ""
sql = "select* from d_Suppliers where sno=" & txtSNo & ""
Set rss = oSaccoMaster.GetRecordset(sql)
If Not rss.EOF Then
   txtSNo_Validate True
   Me.MousePointer = 0
End If
End Sub
Private Sub Text1_Change()

End Sub
Private Sub txtPrice_Click()
If Trim(txtPrice) = "0.00" Then
txtPrice = ""
End If
End Sub

Private Sub txtPrice_Validate(Cancel As Boolean)
If Trim(txtPrice) = "" Then
txtPrice = "0.00"
End If

txtPrice = Format(txtPrice, "#,##0.00")

End Sub
Private Sub txtSNo_Validate(Cancel As Boolean)
If Trim(txtSNo) = "" Then
Exit Sub
End If
Dim mthcp As Integer, thcpactive As Integer, thcppremium   As Double
Dim a, t As Boolean
Set rs = New ADODB.Recordset
sql = "d_sp_SelectSuppliers '" & txtSNo & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
If Not IsNull(rs.Fields(2)) Then txtNames = rs.Fields(2)
If Not IsNull(rs.Fields(1)) Then txtId = rs.Fields(1)
If Not IsNull(rs.Fields(8)) Then cbolocation = rs.Fields(8)
If Not IsNull(rs.Fields(0)) Then DTPRegDate = rs.Fields(0)

End If
cmdEdit.Enabled = True
cmdsave.Enabled = True


End Sub
