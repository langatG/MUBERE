VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTransportersDetails 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Transporters Details"
   ClientHeight    =   8220
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   7605
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFF80&
      Caption         =   "Contacts"
      Height          =   1695
      Left            =   120
      TabIndex        =   20
      Top             =   3720
      Width           =   7455
      Begin VB.TextBox txtTown 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2880
         TabIndex        =   28
         Top             =   1200
         Width           =   3735
      End
      Begin VB.TextBox txtEmail 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   23
         Top             =   480
         Width           =   2655
      End
      Begin VB.TextBox txtPhone 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2880
         TabIndex        =   22
         Top             =   480
         Width           =   3735
      End
      Begin VB.TextBox txtPAddress 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   21
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF80&
         Caption         =   "E - Mail"
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   525
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF80&
         Caption         =   "Phone"
         Height          =   195
         Left            =   2880
         TabIndex        =   26
         Top             =   240
         Width           =   465
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF80&
         Caption         =   "Postal Address"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   960
         Width           =   1050
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF80&
         Caption         =   "Town"
         Height          =   195
         Left            =   2880
         TabIndex        =   24
         Top             =   960
         Width           =   405
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFF80&
      Caption         =   "Other Details"
      Height          =   1335
      Left            =   120
      TabIndex        =   17
      Top             =   6960
      Width           =   7455
      Begin VB.CommandButton cmdNew 
         Caption         =   "New"
         Height          =   375
         Left            =   0
         TabIndex        =   46
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   375
         Left            =   4680
         TabIndex        =   45
         Top             =   840
         Width           =   855
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   375
         Left            =   1440
         TabIndex        =   44
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit"
         Height          =   375
         Left            =   720
         TabIndex        =   43
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtflatratet 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4560
         TabIndex        =   35
         Top             =   360
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CheckBox chkflatrate 
         Caption         =   "Flat Rate"
         Height          =   255
         Left            =   2640
         TabIndex        =   34
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtSubsidy 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   18
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label16 
         Caption         =   "Rate"
         Height          =   255
         Left            =   3840
         TabIndex        =   37
         Top             =   360
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label15 
         Caption         =   "Per  Day"
         Height          =   255
         Left            =   6240
         TabIndex        =   36
         Top             =   360
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF80&
         Caption         =   "Subsidy (Per Kg)"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   1170
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFF80&
      Caption         =   "Personal Details"
      Height          =   3495
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   7455
      Begin VB.TextBox txtcanno 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5040
         TabIndex        =   49
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txtptransporter 
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   2640
         TabIndex        =   47
         Top             =   3120
         Width           =   3615
      End
      Begin VB.TextBox txtttrate 
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   240
         TabIndex        =   42
         Top             =   3120
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CheckBox chktt 
         Caption         =   "Is Transporter Transporter"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   240
         TabIndex        =   39
         Top             =   2400
         Width           =   2295
      End
      Begin VB.ComboBox cboparent 
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   2640
         TabIndex        =   38
         Top             =   2520
         Width           =   2055
      End
      Begin VB.PictureBox Picture5 
         Height          =   255
         Left            =   960
         Picture         =   "frmTransportersDetails.frx":0000
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   33
         Top             =   480
         Width           =   255
      End
      Begin VB.ComboBox cboBranch 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   31
         Top             =   1680
         Width           =   2895
      End
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
         Left            =   3360
         TabIndex        =   29
         Top             =   1680
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.ComboBox cboLocation 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1920
         TabIndex        =   16
         Top             =   1080
         Width           =   2895
      End
      Begin VB.TextBox txtId 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5040
         TabIndex        =   13
         Top             =   480
         Width           =   2295
      End
      Begin VB.TextBox txtNames 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   12
         Top             =   480
         Width           =   3615
      End
      Begin VB.TextBox txtTCode 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   855
      End
      Begin MSComCtl2.DTPicker DTPRegDate 
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         _Version        =   393216
         Format          =   122355713
         CurrentDate     =   40096
      End
      Begin VB.Label lblcanno 
         BackColor       =   &H00FFFF80&
         Caption         =   "Canno"
         Height          =   255
         Left            =   5040
         TabIndex        =   48
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label18 
         Caption         =   "TT Rate"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   2880
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label Label17 
         Caption         =   "Parent Transporter"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2640
         TabIndex        =   40
         Top             =   2280
         Width           =   2055
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF80&
         Caption         =   "Branch"
         Height          =   195
         Left            =   120
         TabIndex        =   32
         Top             =   1440
         Width           =   510
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF80&
         Caption         =   "Location"
         Height          =   195
         Left            =   1920
         TabIndex        =   15
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF80&
         Caption         =   "Id Number/Business No"
         Height          =   195
         Left            =   5040
         TabIndex        =   10
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF80&
         Caption         =   "Date registered"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF80&
         Caption         =   "Names"
         Height          =   195
         Left            =   1200
         TabIndex        =   8
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF80&
         Caption         =   "Trans Code"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   825
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF80&
      Caption         =   "Bank Details"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   5520
      Width           =   7455
      Begin VB.ComboBox cboBBranch 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2520
         TabIndex        =   30
         Top             =   720
         Width           =   2295
      End
      Begin VB.ComboBox cboBName 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox txtAccNo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4920
         TabIndex        =   4
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF80&
         Caption         =   "Account Number"
         Height          =   195
         Left            =   4920
         TabIndex        =   3
         Top             =   360
         Width           =   1200
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF80&
         Caption         =   "Bank Name"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF80&
         Caption         =   "Bank Branch"
         Height          =   195
         Left            =   2640
         TabIndex        =   1
         Top             =   360
         Width           =   930
      End
   End
End
Attribute VB_Name = "frmTransportersDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim flat As Integer

Private Sub cboBName_Change()
Dim supno As String
Dim ACCNO As String
'If cboBName = "FSA" Then
'    Txtaccno.Enabled = False
'    supno = txtTCode
'    supno = Format(supno, "00000")
'    ACCNO = "010100" & supno & "00"
'    Txtaccno = ACCNO
'Else
'    Txtaccno.Enabled = True
'End If

End Sub

Private Sub cboBName_Click()
cboBName_Change
End Sub

Private Sub cboparent_Change()
On Error Resume Next
Dim NAMES As String
Set rst = oSaccoMaster.GetRecordset("select transname from d_transporters where transcode='" & cboparent & "'")
If Not rst.EOF Then
txtptransporter = rst.Fields("transname")
End If

End Sub

Private Sub cboparent_KeyPress(KeyAscii As Integer)
On Error Resume Next
Dim NAMES As String
Set rst = oSaccoMaster.GetRecordset("select transname from d_transporters where transcode='" & cboparent & "'")
If Not rst.EOF Then
txtptransporter = rst.Fields("transname")
End If
End Sub

Private Sub chkflatrate_Click()
If chkflatrate = vbChecked Then
Label15.Visible = True
Label16.Visible = True
txtflatratet.Visible = True
flat = 1
Else
Label15.Visible = False
Label16.Visible = False
txtflatratet.Visible = False
flat = 0
End If
End Sub

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdedit_Click()
Txtaccno.Locked = False
txtEmail.Locked = False
txtId.Locked = False
txtNames.Locked = False
txtPAddress.Locked = False
txtPhone.Locked = False
txtsubsidy.Locked = False
txtTCode.Locked = False
txtTown.Locked = False
cboBBranch.Locked = False
cboBName.Locked = False
cbolocation.Locked = False

'cmdEdit.Enabled = False
'cmdSave.Enabled = False
cmdsave.Enabled = True
End Sub

Private Sub cmdNew_Click()
Txtaccno = ""
txtEmail = ""
txtId = ""
txtNames = ""
txtPAddress = ""
txtPhone = ""
txtsubsidy = ""
txtTCode = ""
txtTown = ""
cboBBranch.Text = ""
cboBName.Text = ""
cbolocation.Text = ""
cbobranch.Text = ""

Txtaccno.Locked = False
txtEmail.Locked = False
txtId.Locked = False
txtNames.Locked = False
txtPAddress.Locked = False
txtPhone.Locked = False
txtsubsidy.Locked = False
txtTCode.Locked = False
txtTown.Locked = False
cboBBranch.Locked = False
cboBName.Locked = False
cbolocation.Locked = False
'cmdEdit.Enabled = False
'cmdSave.Enabled = False
cmdsave.Enabled = True

End Sub

Private Sub cmdsave_Click()
Dim Active As String
Dim rate As Currency
On Error GoTo ErrorHandler

If txtTCode = "" Then
MsgBox "Please enter the transporters code ", vbInformation, "Missing Information"
txtTCode.SetFocus
Exit Sub
End If

If txtsubsidy = "" Then
txtsubsidy = "0"
End If

If chkActive.value = vbChecked Then
    Active = "1"
Else
    Active = "0"
End If
If chkflatrate = vbChecked Then
flat = 1
If txtflatratet = "" Then
MsgBox "The rate per day is not privided for yet you have select the use of Flat Rate", vbInformation, "EASYMA"
Exit Sub
End If
rate = txtflatratet
Else
If txtflatratet = "" Then txtflatratet = 0
rate = txtflatratet
flat = flat
End If

'If chktt = vbChecked Then
'    If cboparent = "" Or txtttrate = "" Then MsgBox "Kindly select the parent transporter and put the rate ", vbCritical
'    Exit Sub
'End If

Set cn = New ADODB.Connection
sql = "d_sp_Transporter '" & txtTCode & "','" & txtNames & "','" & txtId & "','" & cbolocation & "','" & DTPRegDate & "','" & txtEmail & "','" & txtPhone & "','" & txtTown & "','" & txtPAddress & "'," & txtsubsidy & ",'" & Txtaccno & "','" & cboBName & "'," & Active & ",'" & Replace(cboBBranch, "'", "") & "','" & cbobranch & "','" & User & "'," & flat & "," & rate & ",'" & txtcanno & "'"
oSaccoMaster.ExecuteThis (sql)


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'transpoter transporter
Dim tt As Integer
If chktt.value = vbChecked Then
    tt = 1
Else
    tt = 0
End If
oSaccoMaster.ExecuteThis ("update d_transporters set tt='" & tt & "' ,parentT='" & txtptransporter & "',ttrate='" & txtttrate & "' where transcode='" & txtTCode & "'")
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX



cmdNew_Click
cmdsave.Enabled = False

MsgBox "Records successively updated."
Exit Sub
ErrorHandler:
MsgBox err.description

End Sub

Private Sub Form_Load()
Dim myclass As cdbase

Txtaccno.Locked = True
txtEmail.Locked = True
txtId.Locked = True
txtNames.Locked = True
txtPAddress.Locked = True
txtPhone.Locked = True
txtsubsidy.Locked = True
'txtTCode.Locked = True
txtTown.Locked = True
cboBBranch.Locked = True
cboBName.Locked = True
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
         If Not IsNull(rs.Fields("LName")) Then
         cbolocation.AddItem rs.Fields("LName")
         End If
         .MoveNext
        
        Wend
    
    End With
    
    
    Set rs = CreateObject("adodb.recordset")
    
    rs.Open "SELECT BankName,BranchName FROM d_BANKS", cn
    
    If rs.EOF Then Exit Sub
    
    With rs
        
        While Not .EOF
         If Not IsNull(rs.Fields(0)) Then cboBName.AddItem rs.Fields(0)
         If Not IsNull(rs.Fields(1)) Then cboBBranch.AddItem rs.Fields(1)
         
         .MoveNext
        
        Wend
    
    End With
    
     Set rs = CreateObject("adodb.recordset")
    
    rs.Open "SELECT BName FROM d_Branch", cn
    
    If rs.EOF Then Exit Sub
    
    With rs
        
        While Not .EOF
         
         If Not IsNull(rs.Fields(0)) Then cbobranch.AddItem rs.Fields(0)
         
         .MoveNext
        
        Wend
    
    End With
    'load parent transporters
    Set rst = oSaccoMaster.GetRecordset("select transcode from d_transporters where tt=0")
    While Not rst.EOF
        cboparent.AddItem rst.Fields("transcode")
    rst.MoveNext
    Wend

flat = 0
DTPRegDate = Format(Get_Server_Date, "dd/mm/yyyy")
End Sub

Private Sub Picture5_Click()
Me.MousePointer = vbHourglass
         frmSearchPTransporter.Show vbModal
        txtTCode = sel
        txtTCode_Validate True
        Me.MousePointer = 0
End Sub

Private Sub txtTCode_Validate(Cancel As Boolean)
Dim a As Boolean
Set rs = New ADODB.Recordset
sql = "d_sp_SelectTrans '" & txtTCode & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
If Not IsNull(rs.Fields(0)) Then txtNames = rs.Fields(0)
If Not IsNull(rs.Fields(1)) Then txtId = rs.Fields(1)
If Not IsNull(rs.Fields(2)) Then cbolocation = rs.Fields(2)
If Not IsNull(rs.Fields(3)) Then DTPRegDate = rs.Fields(3)
If Not IsNull(rs.Fields(4)) Then txtEmail = rs.Fields(4)
If Not IsNull(rs.Fields(5)) Then txtPhone = rs.Fields(5)
If Not IsNull(rs.Fields(6)) Then txtTown = rs.Fields(6)
If Not IsNull(rs.Fields(7)) Then txtPAddress = rs.Fields(7)
If Not IsNull(rs.Fields(8)) Then txtsubsidy = rs.Fields(8)
If Not IsNull(rs.Fields(9)) Then Txtaccno = rs.Fields(9)
If Not IsNull(rs.Fields(10)) Then cboBName = rs.Fields(10)
If Not IsNull(rs.Fields(11)) Then cboBBranch = rs.Fields(11)
If Not IsNull(rs.Fields(12)) Then a = rs.Fields(12)
If Not IsNull(rs.Fields(13)) Then cbobranch = rs.Fields(13)
If a = True Then
chkActive = vbChecked
Else
chkActive = vbUnchecked
End If
cmdEdit.Enabled = True
cmdsave.Enabled = True
End If
End Sub
