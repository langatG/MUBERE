VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Begin VB.Form frmVehicleReg 
   BackColor       =   &H00FFC0FF&
   Caption         =   "Vehicle Registration"
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5460
   LinkTopic       =   "Form1"
   ScaleHeight     =   6540
   ScaleWidth      =   5460
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fra1 
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   1080
      TabIndex        =   14
      Top             =   4800
      Width           =   4215
      Begin VB.TextBox txtcracc 
         Height          =   375
         Left            =   1680
         TabIndex        =   18
         Top             =   600
         Width           =   2895
      End
      Begin VB.TextBox txtdracc 
         Height          =   375
         Left            =   1680
         TabIndex        =   17
         Top             =   120
         Width           =   3375
      End
      Begin VB.PictureBox Picture3 
         Height          =   255
         Left            =   1320
         Picture         =   "frmVehicleReg.frx":0000
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   16
         Top             =   720
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Height          =   255
         Left            =   1320
         Picture         =   "frmVehicleReg.frx":08CA
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   15
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblcracc 
         BackColor       =   &H8000000E&
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lbldracc 
         BackColor       =   &H8000000E&
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdNew1 
      Caption         =   "New"
      Height          =   495
      Left            =   240
      TabIndex        =   13
      Top             =   6000
      Width           =   855
   End
   Begin VB.TextBox txtBName1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   7
      Top             =   3000
      Width           =   3735
   End
   Begin VB.TextBox txtBCode1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   6
      Top             =   2520
      Width           =   3735
   End
   Begin VB.TextBox txttill 
      Appearance      =   0  'Flat
      CausesValidation=   0   'False
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Top             =   3480
      Width           =   3735
   End
   Begin VB.TextBox txtPhoNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Top             =   3960
      Width           =   3735
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   4080
      TabIndex        =   3
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3120
      TabIndex        =   2
      Top             =   6000
      Width           =   855
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1200
      TabIndex        =   1
      Top             =   6000
      Width           =   855
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   2160
      TabIndex        =   0
      Top             =   6000
      Width           =   855
   End
   Begin MSComctlLib.ListView lvWBranch1 
      Height          =   2295
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   4048
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
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Dr Stock"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   23
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Cr Sales"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   22
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "Ledgers only"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   21
      Top             =   4320
      Width           =   2295
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle/Plant"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   12
      Top             =   3000
      Width           =   1230
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   11
      Top             =   2520
      Width           =   435
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "TILL No."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone No."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3960
      Width           =   1215
   End
End
Attribute VB_Name = "frmVehicleReg"
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
sql = "delete from d_VehicleTill where Code='" & txtBCode1 & "'"

myclass.Delete sql
loadBranchesTypes
txtBCode1 = ""
txtBName1 = ""
txttill = ""
txtPhoNo = ""
lbldracc = ""
txtdracc = ""
lblcracc = ""
txtcracc = ""
End Sub

Private Sub cmdedit_Click()
txtBCode1.Locked = False
txtBName1.Locked = False
txttill.Locked = False
txtPhoNo.Locked = False
'lbldracc.Locked = False
'lblcracc.Locked = False
cmdNew1.Enabled = False
cmdEdit.Enabled = False
End Sub
Private Sub cmdNew1_Click()
txtBCode1 = ""
txtBName1 = ""
txttill = ""
txtPhoNo = ""
lbldracc = ""
txtdracc = ""
lblcracc = ""
txtcracc = ""
txtBCode1.Locked = False
txtBName1.Locked = False
txttill.Locked = False
txtPhoNo.Locked = False
cmdNew1.Enabled = False
cmdEdit.Enabled = False
cmdsave.Enabled = True
sql = ""
sql = "select count(Vehicle) from d_VehicleTill"
Set rs = oSaccoMaster.GetRecordset(sql)

If Not rs.EOF Then
txtBCode1 = rs.Fields(0) + 1
Else
txtBCode1 = 1
End If

End Sub

Public Sub loadBranchesTypes()
    
    With lvWBranch1
    
        .ListItems.Clear
        
        .ColumnHeaders.Clear
    
    End With

    Set rs2 = CreateObject("adodb.recordset")
    
    sql = "Select * from d_VehicleTill"
    
    Set rs2 = CreateObject("adodb.recordset")
    
    Set clsClass = New cdbase
    
    Provider = clsClass.OpenCon
    
    Set cn = CreateObject("adodb.connection")
    
   cn.Open Provider, "atm", "atm"
    
    rs2.Open sql, cn
    
    With lvWBranch1
        
        .ColumnHeaders.Add , , "Vehicle Code"
        .ColumnHeaders.Add , , "Vehicle Name"
        .ColumnHeaders.Add , , "Vehicle Till"
        .ColumnHeaders.Add , , "Vehicle PhoneNo"
        .ColumnHeaders.Add , , "Outlet Dr"
        .ColumnHeaders.Add , , "Outlet Cr"
        While Not rs2.EOF
        
            Set li = .ListItems.Add(, , Trim(rs2.Fields("Code")))
            
            li.ListSubItems.Add , , Trim(rs2.Fields("Vehicle"))
            li.ListSubItems.Add , , Trim(rs2.Fields("TillNo"))
            li.ListSubItems.Add , , Trim(rs2.Fields("Phone"))
            li.ListSubItems.Add , , Trim(rs2.Fields("Dr"))
            li.ListSubItems.Add , , Trim(rs2.Fields("Cr"))
            rs2.MoveNext
        
        Wend
        
    End With
    
    rs2.Close
    
    Set rs2 = Nothing
    
lvWBranch1.View = lvwReport

End Sub

Private Sub cmdsave_Click()

On Error GoTo ErrorHandler
If txtBCode1 = "" Then
MsgBox "Enter the Branch Code", vbInformation
Exit Sub 'txtBName
End If
If txtBName1 = "" Then
MsgBox "Enter the Branch Code", vbInformation
Exit Sub 'txtBName
End If
If txttill = "" Then
MsgBox "Enter the Agent Till No", vbInformation
Exit Sub 'txtBName
End If
If txtPhoNo = "" Then
MsgBox "Enter the Outlet agent Phone", vbInformation
Exit Sub 'txtBName
End If
'Code, Vehicle, TillNo, Phone, Auditdatetime, UserID
Set cn = New ADODB.Connection
sql = "d_sp_VehicleTill '" & txtBCode1 & "','" & txtBName1 & "','" & txttill & "','" & txtPhoNo & "','" & User & "','" & lbldracc & "','" & lblcracc & "'"
oSaccoMaster.ExecuteThis (sql)
txtBCode1 = ""
txtBName1 = ""
txttill = ""
txtPhoNo = ""
lbldracc = ""
lblcracc = ""
txtdracc = ""
txtcracc = ""
txtBCode1.Locked = True
txtBName1.Locked = True
cmdNew1.Enabled = True
cmdEdit.Enabled = False
cmdsave.Enabled = True
loadBranchesTypes
MsgBox "Records successively updated."
Exit Sub
ErrorHandler:
MsgBox err.description
End Sub

Private Sub Form_Load()
txtBCode1.Locked = True
txtBName1.Locked = True
txttill.Locked = True
txtPhoNo.Locked = True
cmdDelete.Enabled = False
loadBranchesTypes
End Sub

Public Sub edit(selected As String)
Set myclass = New cdbase
Set cn = CreateObject("adodb.connection")
Provider = myclass.OpenCon
cn.Open Provider, "atm", "atm"
Set rs = CreateObject("adodb.recordset")
sql = "select * from d_VehicleTill where Code='" & selected & "'"
rs.Open sql, cn
If Not rs.EOF Then
', ,
txtBCode1 = selected
txtBName1 = rs!Vehicle
txttill = rs!TillNo
txtPhoNo = rs!Phone
lbldracc = rs!dr
lblcracc = rs!cr
End If
cmdDelete.Enabled = True

End Sub
Private Sub lvWBranch_DblClick()
cmdEdit.Enabled = True
edit lvWBranch1.SelectedItem
End Sub

Private Sub lblcracc_Change()
    Set rst = oSaccoMaster.GetRecordset("select glaccname from glsetup where accno='" & lblcracc & "'")
    If Not rst.EOF Then
    txtcracc = rst.Fields("glaccname")
    End If
End Sub
Private Sub lbldracc_Change()
    Set rst = oSaccoMaster.GetRecordset("select glaccname from glsetup where accno='" & lbldracc & "'")
    If Not rst.EOF Then
    txtdracc = rst.Fields("glaccname")
    End If
End Sub
Private Sub lvWBranch1_DblClick()
cmdEdit.Enabled = True
edit lvWBranch1.SelectedItem
End Sub

Private Sub Picture1_Click()
Me.MousePointer = vbHourglass
        frmSearchGLAccounts.Show vbModal, Me
    If Continue Then
        If SearchValue <> "" Then
            lbldracc = SearchValue
            SearchValue = ""
        End If
    End If
    Me.MousePointer = 0
End Sub

Private Sub Picture3_Click()
Me.MousePointer = vbHourglass
        frmSearchGLAccounts.Show vbModal, Me
    If Continue Then
        If SearchValue <> "" Then
            lblcracc = SearchValue
            SearchValue = ""
        End If
    End If
    Me.MousePointer = 0
End Sub
'

