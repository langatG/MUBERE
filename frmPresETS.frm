VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPresETS 
   BackColor       =   &H00FFC0FF&
   Caption         =   "DEDUCTION SETTINGS"
   ClientHeight    =   3225
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   9750
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPresETS.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3225
   ScaleWidth      =   9750
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmdbranches 
      Height          =   435
      Left            =   7440
      TabIndex        =   19
      Top             =   600
      Width           =   2175
   End
   Begin VB.OptionButton optAllBra 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Caption         =   "Specific Branches"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4560
      TabIndex        =   14
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   375
      Left            =   2640
      TabIndex        =   13
      Top             =   2760
      Width           =   1575
   End
   Begin VB.ComboBox cboDeduct 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      ItemData        =   "frmPresETS.frx":030A
      Left            =   1320
      List            =   "frmPresETS.frx":0314
      TabIndex        =   12
      Top             =   1680
      Width           =   2415
   End
   Begin MSComCtl2.DTPicker dtpSDate 
      Height          =   375
      Left            =   5040
      TabIndex        =   11
      Top             =   1200
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   144703489
      CurrentDate     =   40209
   End
   Begin VB.TextBox txtRate 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   10
      Top             =   2160
      Width           =   2415
   End
   Begin VB.TextBox txtRemarks 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   9
      Top             =   1680
      Width           =   3255
   End
   Begin VB.TextBox txtSNo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   8
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CheckBox chkStopped 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Stopped"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   3960
      TabIndex        =   7
      Top             =   2160
      Width           =   1335
   End
   Begin VB.OptionButton optSpecific 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Caption         =   "Specific Suppliers"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   2895
   End
   Begin VB.OptionButton optAllSup 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Caption         =   "All Suppliers"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   120
      TabIndex        =   15
      Top             =   480
      Width           =   6855
      Begin VB.OptionButton optRate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0FF&
         Caption         =   "Rate per Kg Supplied"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Value           =   -1  'True
         Width           =   2775
      End
      Begin VB.OptionButton optAmnt 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0FF&
         Caption         =   "Fixed amount per month"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   3720
         TabIndex        =   16
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Line Line2 
      X1              =   7215
      X2              =   7215
      Y1              =   0
      Y2              =   1680
   End
   Begin VB.Line Line1 
      X1              =   3495
      X2              =   3480
      Y1              =   0
      Y2              =   1680
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Start Date"
      Height          =   315
      Left            =   3840
      TabIndex        =   6
      Top             =   1200
      Width           =   1110
   End
   Begin VB.Label lblRemarks 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      Height          =   315
      Left            =   3840
      TabIndex        =   5
      Top             =   1680
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Rate/Kg"
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   915
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Deduction"
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   1155
   End
   Begin VB.Label lblSNo 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "SNo"
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   510
   End
End
Attribute VB_Name = "frmPresETS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private Sub cboDeduct_Change()
''If UCase(cboDeduct.Text) = "OTHERS" Then
'lblRemarks.Visible = True
'txtremarks.Visible = True
'txtremarks = ""
''Else
''lblRemarks.Visible = False
''txtremarks.Visible = False
''txtremarks = ""
'End If
'End Sub
Private Sub cboDeduct_Click()
'cboDeduct_Change
If cboDeduct = "BONUS" Then
txtRemarks = "BONUS"
End If
If cboDeduct = "Others" Then
txtRemarks = "Monthly Charges"
End If
End Sub
Private Sub cboDeduct_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Command1_Click()
If (optSpecific.value = True) And Trim(txtSNo = "") Then
 MsgBox "Please enter the supplier number."
    txtSNo.SetFocus
Exit Sub
End If

If (optAllBra.value = True) And Trim(cmdbranches = "") Then
 MsgBox "Please select the Branch Name ."
    cmdbranches.SetFocus
Exit Sub
End If

If Trim(cboDeduct.Text) = "" Then
 MsgBox "Please enter the type of deduction."
    cboDeduct.SetFocus
Exit Sub
End If

If (cboDeduct.Text = "Others") And Trim(txtRemarks = "") Then
 MsgBox "Please enter the Remark or Description."
    txtRemarks.SetFocus
Exit Sub
End If

If Trim(txtRate.Text) = "" Then
 MsgBox "Please enter the rate."
    txtRate.SetFocus
Exit Sub
End If
Dim St As Integer

If chkStopped.value = vbChecked Then
St = 1
Else
St = 0
End If

Dim Chk As Integer
If optAllSup.value = True Then
Chk = 1
Else
Chk = 0
Dim Chkd As String
If optAllBra.value = True Then
Chkd = cmdbranches
Chk = 2
Else
Chk = 0
End If
End If

'd_SP_PreSets
'    ( @SNo  [bigint],
'     @Deduction     [varchar](50),
'     @Remark    [varchar](150),
'     @StartDate     [varchar](50),
'     @Rate  [money],
'     @Stopped   [bit],
'     @AuditId   [varchar](50))
Dim desc As String
desc = cboDeduct.Text


If txtSNo = "" Then
txtSNo = 1
End If

Dim Startdate, Enddate As String
Dim Rated As Integer

If optRate.value = True Then
Rated = 1
Else
Rated = 0
End If
Startdate = DateSerial(year(dtpSDate), month(dtpSDate), 1)
Enddate = DateSerial(year(dtpSDate), month(dtpSDate) + 1, 1 - 1)
'd_SP_PreSets
'    ( @SNo  [bigint],
'     @Deduction     [varchar](50),  @Remark    [varchar](150),     @StartDate     [varchar](50),
'     @Rate  [money],
'     @Stopped   [bit],
'     @AuditId   [varchar](50),
'    @HowMuch    [bigint])
sql = ""
sql = "d_SP_PreSets " & txtSNo & ",'" & cboDeduct & "','" & txtRemarks & "','" & dtpSDate & "'," & txtRate & "," & St & ",'" & User & "'," & Chk & "," & Rated & ",'" & Chkd & "'"
oSaccoMaster.ExecuteThis (sql)

cmdbranches = ""

MsgBox "Records Saved successfully!"
loadbranch
If txtSNo.Visible = True Then
txtSNo = ""
txtSNo.SetFocus
End If

'd_sp_GDedNet @StartDate varchar(10) , @endPeriod varchar(10)   AS

End Sub

Private Sub Command2_Click()
Unload Me
End Sub
Private Sub loadbranch()
    cmdbranches.Clear
    Set rst = New Recordset
    Dim cn As Connection
    Set cn = New ADODB.Connection
    Provider = "MAZIWA"
    'Provider = cnn
    cn.Open Provider, "atm", "atm"
    Set rst = New Recordset
    
    sql = "Select Bname from d_Branch order by Bname asc"
    rst.Open sql, cn, adOpenKeyset, adLockOptimistic
    While Not rst.EOF
    cmdbranches.AddItem rst.Fields(0)
    rst.MoveNext
    Wend
End Sub

Private Sub Form_Load()
dtpSDate = Format(Get_Server_Date, "dd/mm/yyyy")
cmdbranches.Visible = False
'Set cn = CreateObject("adodb.connection")
'
'    cn.Open Provider, "mpesa", "mpesa"
'
'    Set rs = CreateObject("adodb.recordset")
'
'    rs.Open "SELECT Description FROM d_DCodes", cn
'
'    If rs.EOF Then Exit Sub
'
'    With rs
'
'        While Not .EOF
'
'         cboDeduct.AddItem rs.Fields("Description")
'
'         .MoveNext
'
'        Wend
'
'    End With
End Sub

Private Sub optAllBra_Click()
lblSNo.Visible = False
txtSNo.Visible = False
cmdbranches.Visible = True
loadbranch
End Sub

Private Sub optAllSup_Click()
lblSNo.Visible = False
txtSNo.Visible = False

End Sub

Private Sub Option1_Click()

End Sub

Private Sub optAmnt_Click()
Label1 = "Amount"
End Sub

Private Sub optRate_Click()
Label1 = "Rate/Kg"
End Sub

Private Sub optSpecific_Click()
lblSNo.Visible = True
txtSNo.Visible = True
End Sub

Private Sub txtRate_Click()
If Trim(txtRate) = "0.00" Then
txtRate = ""
End If
End Sub

Private Sub txtRate_Validate(Cancel As Boolean)
If Trim(txtRate) = "" Then
txtRate = "0"
End If

txtRate = Format(txtRate, "#0.00")
End Sub
