VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Begin VB.Form frmDCodes 
   Caption         =   "Deductions code"
   ClientHeight    =   5970
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7710
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDCodes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   7710
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lvwContraAcc 
      Height          =   1215
      Left            =   3360
      TabIndex        =   24
      Top             =   2640
      Visible         =   0   'False
      Width           =   3990
      _ExtentX        =   7038
      _ExtentY        =   2143
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "AccNo"
         Object.Width           =   18
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Acc Name"
         Object.Width           =   10583
      EndProperty
   End
   Begin MSComctlLib.ListView lvwDeductionAcc 
      Height          =   945
      Left            =   3360
      TabIndex        =   25
      Top             =   720
      Visible         =   0   'False
      Width           =   3990
      _ExtentX        =   7038
      _ExtentY        =   1667
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "AccNo"
         Object.Width           =   18
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Acc Name"
         Object.Width           =   10583
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "GL Integration"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   18
      Top             =   1440
      Width           =   7455
      Begin VB.TextBox txtContraAccName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3240
         TabIndex        =   26
         Top             =   840
         Width           =   3975
      End
      Begin VB.TextBox txtDeductionAccName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3270
         TabIndex        =   21
         Top             =   240
         Width           =   3975
      End
      Begin VB.TextBox txtDeductionAcc 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtContraAcc 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Contra Account"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   23
         Top             =   840
         Width           =   1125
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Deduction Account"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   22
         Top             =   240
         Width           =   1350
      End
   End
   Begin VB.PictureBox CrystalReport1 
      Height          =   480
      Left            =   5175
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   17
      Top             =   5340
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Height          =   495
      Left            =   4680
      Picture         =   "frmDCodes.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Cancel Process"
      Top             =   5400
      Width           =   495
   End
   Begin VB.CommandButton cmdUpdate 
      Height          =   495
      Left            =   4200
      Picture         =   "frmDCodes.frx":040C
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Save Record"
      Top             =   5400
      Width           =   495
   End
   Begin VB.CommandButton cmdDelete 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      Picture         =   "frmDCodes.frx":04FE
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Delete Record"
      Top             =   5400
      Width           =   495
   End
   Begin VB.CommandButton cmdEdit 
      Height          =   495
      Left            =   3240
      Picture         =   "frmDCodes.frx":05F0
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Edit Record"
      Top             =   5400
      Width           =   495
   End
   Begin VB.CommandButton cmdNext 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1095
      Picture         =   "frmDCodes.frx":06F2
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Move to the Next"
      Top             =   5400
      Width           =   495
   End
   Begin VB.CommandButton cmdLast 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1605
      Picture         =   "frmDCodes.frx":0A34
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Move to Last record"
      Top             =   5400
      Width           =   495
   End
   Begin VB.CommandButton cmdPrev 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      Picture         =   "frmDCodes.frx":0D76
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Move to the Previous record"
      Top             =   5400
      Width           =   495
   End
   Begin MSComctlLib.ListView lvwdeducitonsApplications 
      Height          =   1665
      Left            =   90
      TabIndex        =   15
      Top             =   3600
      Width           =   7530
      _ExtentX        =   13282
      _ExtentY        =   2937
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   16711680
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.CommandButton cmdAdd 
      Height          =   495
      Left            =   2760
      Picture         =   "frmDCodes.frx":10B8
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Add New record"
      Top             =   5400
      Width           =   495
   End
   Begin VB.CommandButton cmdFirst 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Picture         =   "frmDCodes.frx":11AA
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Move to the Last record"
      Top             =   5400
      Width           =   495
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   510
      Left            =   6360
      TabIndex        =   11
      Top             =   5400
      Width           =   1230
   End
   Begin VB.Frame frameLoanTypes 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   7470
      Begin VB.TextBox txtdCode 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtdType 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   1605
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   480
         Width           =   4215
      End
      Begin VB.Label Label7 
         Caption         =   "&Deduction Code"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Deduction &Type"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1605
         TabIndex        =   13
         Top             =   240
         Width           =   1575
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2160
      Top             =   5265
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDCodes.frx":14EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDCodes.frx":15FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDCodes.frx":1710
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblLoanApp 
      Caption         =   "Deduction Applications"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   3360
      Width           =   2085
   End
End
Attribute VB_Name = "frmDCodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myclass As cdbase
Private Sub cmdAdd_Click()
Dim I As Integer
I = 1
Set rs = oSaccoMaster.GetRecordset("SELECT DCode FROM d_DCodes WHERE DCode = '" & I & "'")
While Not rs.EOF
Set rs = oSaccoMaster.GetRecordset("SELECT DCode FROM d_DCodes WHERE DCode = '" & I & "'")
I = I + 1
Wend
I = I - 1
txtdcode = I
txtdType = ""
txtContraAcc = ""
txtContraAccName = ""
txtContraAcc.Locked = False
txtContraAccName.Locked = False
txtDeductionAcc.Locked = False
txtdcode.Locked = True
txtdType.Locked = False
cmdAdd.Enabled = False
cmdEdit.Enabled = False
cmdupdate.Enabled = True
cmdDelete.Enabled = False
cmdCancel.Enabled = True
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
sql = "delete from d_DCodes where DCode='" & txtdcode & "'"
myclass.Delete sql

Form_Load
End Sub

Private Sub cmdedit_Click()
txtContraAcc.Locked = False
txtContraAccName.Locked = False
txtdcode.Locked = False
txtdType.Locked = False
cmdAdd.Enabled = False
cmdEdit.Enabled = False
cmdupdate.Enabled = True
cmdDelete.Enabled = False
cmdCancel.Enabled = True
End Sub

Private Sub cmdupdate_Click()
On Error GoTo ErrorHandler

If txtdcode = "" Then
MsgBox "Please enter the deduction code", vbInformation
txtdcode.SetFocus
Exit Sub
End If
If txtContraAcc = "" Then
MsgBox "Please enter the contra account", vbInformation
txtContraAcc.SetFocus
Exit Sub
End If
If txtDeductionAcc = "" Then
MsgBox "Please enter the deduction account", vbInformation
txtDeductionAcc.SetFocus
Exit Sub
End If

Set cn = New ADODB.Connection
sql = "SET dateformat dmy SELECT DCode from d_DCodes WHERE DCode='" & txtdcode & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
Dim ans As String
If Not rs.EOF Then
ans = MsgBox("The code is already in use. Update?", vbYesNo, "EXISTING CODES")
If ans = vbNo Then
Exit Sub
End If
End If

Set cn = New ADODB.Connection
sql = "d_sp_DCodes '" & txtdcode & "','" & txtdType & "','" & txtDeductionAcc & "','" & txtContraAcc & "','" & User & "'"
oSaccoMaster.ExecuteThis (sql)
txtdcode = ""
txtdType = ""
txtContraAcc = ""
txtContraAccName = ""
txtContraAcc.Locked = False
txtContraAccName.Locked = False
txtdcode.Locked = False
txtdType.Locked = False
cmdAdd.Enabled = False
cmdEdit.Enabled = False
cmdupdate.Enabled = True
cmdDelete.Enabled = False
cmdCancel.Enabled = True

loadDCodes

MsgBox "Records successively updated."
Exit Sub
ErrorHandler:
MsgBox err.description

End Sub
Public Sub loadDCodes()
    
    With lvwdeducitonsApplications
    
        .ListItems.Clear
        
        .ColumnHeaders.Clear
    
    End With

    Set rs = CreateObject("adodb.recordset")
    
    sql = "Select * from d_DCodes"
    
    Set rs = CreateObject("adodb.recordset")
    
    Set clsClass = New cdbase
    
    Provider = clsClass.OpenCon
    
    Set cn = CreateObject("adodb.connection")
    
   cn.Open Provider, "atm", "atm"
    
    rs.Open sql, cn
    
    With lvwdeducitonsApplications
        
        .ColumnHeaders.Add , , "Deduction Code"
        .ColumnHeaders.Add , , "Deduction Name"
        .ColumnHeaders.Add , , "Deduction Account"
        .ColumnHeaders.Add , , "Contra Account"
    
        While Not rs.EOF
        
            Set li = .ListItems.Add(, , Trim(rs.Fields("DCode")))
            
            li.ListSubItems.Add , , Trim(rs.Fields("Description"))
            li.ListSubItems.Add , , Trim(rs.Fields("Dedaccno"))
            li.ListSubItems.Add , , Trim(rs.Fields("Contraacc"))
            
        
            
            rs.MoveNext
        
        Wend
        
    End With
    
    rs.Close
    
    Set rs = Nothing
    
lvwdeducitonsApplications.View = lvwReport

End Sub

Private Sub Form_Load()
txtdcode = ""
txtdType = ""
txtContraAcc = ""
txtContraAccName = ""
txtContraAcc.Locked = True
txtContraAcc.Locked = True
txtdcode.Locked = False
txtdType.Locked = False
cmdAdd.Enabled = True
cmdEdit.Enabled = True
cmdupdate.Enabled = True
'cmdDelete.Enabled = False
cmdCancel.Enabled = False
loadDCodes
End Sub

Private Sub lvwContraAcc_DblClick()
Dim rsAccount As New ADODB.Recordset
txtContraAcc = lvwContraAcc.SelectedItem
Set rsAccount = oSaccoMaster.GetRecordset("Select * From GLSETUP where " _
            & "accno= '" & txtContraAcc & "'")
If Not rsAccount.EOF Then
   txtContraAccName = IIf(IsNull(rsAccount!GlAccName), "", rsAccount!GlAccName)
  
 
End If


lvwContraAcc.Visible = False
End Sub

Public Sub edit(selected As String)

Set myclass = New cdbase
Set cn = CreateObject("adodb.connection")
Provider = myclass.OpenCon
cn.Open Provider, "atm", "atm"
Set rs = CreateObject("adodb.recordset")
sql = "select * from d_DCodes where DCode='" & selected & "'"
rs.Open sql, cn
If Not rs.EOF Then
txtdcode = selected
txtContraAcc = rs!ContraAcc
Set rst = New ADODB.Recordset
sql = "select glaccname from glsetup where accno='" & txtContraAcc & "'"
rst.Open sql, cn
If Not rst.EOF Then
txtContraAccName = rst.Fields(0)
End If

txtdType = rs!description
txtDeductionAcc = rs!DedAccno

Set rst = New ADODB.Recordset
sql = "select glaccname from glsetup where accno='" & txtDeductionAcc & "'"
rst.Open sql, cn
If Not rst.EOF Then
txtDeductionAccName = rst.Fields(0)
End If
End If
cmdDelete.Enabled = True

End Sub

Private Sub lvwdeducitonsApplications_DblClick()
edit lvwdeducitonsApplications.SelectedItem
lvwContraAcc.Visible = False
lvwDeductionAcc.Visible = False
End Sub

Private Sub txtContraAccName_Change()
On Error GoTo SysError
    Dim rsAccount As New Recordset
    lvwContraAcc.ListItems.Clear
    
    If Trim$(txtContraAccName) <> "" Then
        'If Editing = True Then
            Set rsAccount = oSaccoMaster.GetRecordset("Select * From GLSETUP where " _
            & "GLAccName Like '%" & txtContraAccName & "%'")
            With rsAccount
                If .State = adStateOpen Then
                    If Not .EOF Then
                        'lvwContraAcc.Visible = True
                        If .RecordCount = 1 Then
                            txtContraAcc = IIf(IsNull(!ACCNO), "", !ACCNO)
                            Editing = True
                            txtContraAccName = IIf(IsNull(!GlAccName), "", !GlAccName)
                            lvwContraAcc.Visible = False
                            Else
                            lvwContraAcc.Visible = False
                            
                        End If
                    Else
                        lvwContraAcc.Visible = False
                    End If
                    'lvwDeductionAcc.Visible = True
                    While Not .EOF
                        lvwContraAcc.Visible = True
                        Set li = lvwContraAcc.ListItems.Add(, , IIf(IsNull(!ACCNO), "", !ACCNO))
                        li.SubItems(1) = IIf(IsNull(!GlAccName), "", !GlAccName)
                        .MoveNext
                    Wend
                    'lvwDeductionAcc.Visible = False
                End If
            End With
        'End If
    End If
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption

End Sub



Private Sub txtDeductionAccName_Change()
On Error GoTo SysError
    Dim rsAccount As New Recordset
    lvwDeductionAcc.ListItems.Clear
    
    If Trim$(txtDeductionAccName) <> "" Then
        'If Editing = True Then
            Set rsAccount = oSaccoMaster.GetRecordset("Select * From GLSETUP where " _
            & "GLAccName Like '%" & txtDeductionAccName & "%'")
            With rsAccount
                If .State = adStateOpen Then
                    If Not .EOF Then
                        'lvwContraAcc.Visible = True
                        If .RecordCount = 1 Then
                            txtDeductionAcc = IIf(IsNull(!ACCNO), "", !ACCNO)
                            Editing = True
                            txtDeductionAccName = IIf(IsNull(!GlAccName), "", !GlAccName)
                            lvwDeductionAcc.Visible = False
                            Else
                            lvwDeductionAcc.Visible = False
                            
                        End If
                    Else
                        lvwDeductionAcc.Visible = False
                    End If
                    'lvwDeductionAcc.Visible = True
                    While Not .EOF
                        lvwDeductionAcc.Visible = True
                        Set li = lvwDeductionAcc.ListItems.Add(, , IIf(IsNull(!ACCNO), "", !ACCNO))
                        li.SubItems(1) = IIf(IsNull(!GlAccName), "", !GlAccName)
                        .MoveNext
                    Wend
                    'lvwDeductionAcc.Visible = False
                End If
            End With
        'End If
    End If
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub lvwDeductionAcc_DblClick()
Dim rsAccount As New ADODB.Recordset
txtDeductionAcc = lvwDeductionAcc.SelectedItem
Set rsAccount = oSaccoMaster.GetRecordset("Select * From GLSETUP where " _
            & "accno= '" & txtDeductionAcc & "'")
If Not rsAccount.EOF Then
   txtDeductionAccName = IIf(IsNull(rsAccount!GlAccName), "", rsAccount!GlAccName)
  
 
End If


lvwDeductionAcc.Visible = False
End Sub
