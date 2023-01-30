VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Begin VB.Form frmOutletreg 
   Caption         =   "Outlet Registration"
   ClientHeight    =   5445
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   6120
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fra1 
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   960
      TabIndex        =   15
      Top             =   3720
      Width           =   4215
      Begin VB.PictureBox Picture1 
         Height          =   255
         Left            =   1320
         Picture         =   "frmOutletreg.frx":0000
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   19
         Top             =   240
         Width           =   255
      End
      Begin VB.PictureBox Picture3 
         Height          =   255
         Left            =   1320
         Picture         =   "frmOutletreg.frx":08CA
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   18
         Top             =   720
         Width           =   255
      End
      Begin VB.TextBox txtdracc 
         Height          =   375
         Left            =   1680
         TabIndex        =   17
         Top             =   120
         Width           =   3375
      End
      Begin VB.TextBox txtcracc 
         Height          =   375
         Left            =   1680
         TabIndex        =   16
         Top             =   600
         Width           =   2895
      End
      Begin VB.Label lbldracc 
         BackColor       =   &H8000000E&
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   120
         Width           =   975
      End
      Begin VB.Label lblcracc 
         BackColor       =   &H8000000E&
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.TextBox txtPhoNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      TabIndex        =   14
      Top             =   2880
      Width           =   3855
   End
   Begin VB.TextBox txttill 
      Appearance      =   0  'Flat
      CausesValidation=   0   'False
      Height          =   285
      Left            =   1320
      TabIndex        =   10
      Top             =   2520
      Width           =   3855
   End
   Begin VB.TextBox txtBCode1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      TabIndex        =   6
      Top             =   1800
      Width           =   3855
   End
   Begin VB.TextBox txtBName1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      TabIndex        =   5
      Top             =   2160
      Width           =   3855
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   5040
      Width           =   855
   End
   Begin MSComctlLib.ListView lvWBranch1 
      Height          =   1695
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   5895
      _ExtentX        =   10398
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
   Begin VB.Label Label7 
      Caption         =   "Outlet Ledgers only"
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
      Left            =   240
      TabIndex        =   23
      Top             =   3360
      Width           =   2295
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
      Left            =   120
      TabIndex        =   22
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Agent Phno"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2880
      Width           =   975
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
      Left            =   120
      TabIndex        =   12
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "TILL No."
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Branch Code"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   930
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Branch Name"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   975
   End
End
Attribute VB_Name = "frmOutletreg"
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
sql = "delete from d_Outletbranch where BCode='" & txtBCode1 & "'"

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
'lbldracc = False
'lblcracc = False
cmdnew.Enabled = False
cmdEdit.Enabled = False
End Sub

Private Sub cmdNew_Click()
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
cmdnew.Enabled = False
cmdEdit.Enabled = False
cmdsave.Enabled = True
sql = ""
sql = "select count(BCode1) from d_Outletbranch"
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
    
    sql = "Select * from d_Outletbranch"
    
    Set rs2 = CreateObject("adodb.recordset")
    
    Set clsClass = New cdbase
    
    Provider = clsClass.OpenCon
    
    Set cn = CreateObject("adodb.connection")
    
   cn.Open Provider, "atm", "atm"
    
    rs2.Open sql, cn
    
    With lvWBranch1
        
        .ColumnHeaders.Add , , "Outlet Code"
        .ColumnHeaders.Add , , "Outlet Name"
        .ColumnHeaders.Add , , "Outlet Till"
        .ColumnHeaders.Add , , "Outlet PhoneNo"
        .ColumnHeaders.Add , , "Outlet Dr"
        .ColumnHeaders.Add , , "Outlet Cr"
        While Not rs2.EOF
        
            Set li = .ListItems.Add(, , Trim(rs2.Fields("BCode1")))
            
            li.ListSubItems.Add , , Trim(rs2.Fields("BName1"))
            li.ListSubItems.Add , , Trim(rs2.Fields("Till"))
            li.ListSubItems.Add , , Trim(rs2.Fields("PhoneNo"))
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
If lbldracc = "" Then
MsgBox "Enter the Outlet Ledger", vbInformation
Exit Sub 'txtBName
End If

Set cn = New ADODB.Connection
sql = "d_sp_OutletBranch '" & txtBCode1 & "','" & txtBName1 & "','" & User & "','" & Date & "','" & txttill & "','" & txtPhoNo & "','" & lbldracc & "','" & lblcracc & "'"
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
cmdnew.Enabled = True
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
sql = "select * from d_Outletbranch where BCode1='" & selected & "'"
rs.Open sql, cn
If Not rs.EOF Then
', ,
txtBCode1 = selected
txtBName1 = rs!Bname1
txttill = rs!Till
txtPhoNo = rs!PhoneNo
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

