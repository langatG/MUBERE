VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDebtorsDetails 
   Caption         =   "Debtors Details"
   ClientHeight    =   8925
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7575
   LinkTopic       =   "Form2"
   ScaleHeight     =   8925
   ScaleWidth      =   7575
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFF80&
      Caption         =   "Other Details"
      Height          =   3135
      Left            =   120
      TabIndex        =   35
      Top             =   5280
      Width           =   7455
      Begin VB.CheckBox chkcessapp 
         Caption         =   "Cess Applicable"
         Height          =   255
         Left            =   4680
         TabIndex        =   59
         Top             =   1920
         Width           =   1575
      End
      Begin VB.TextBox txtcessrate 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3240
         TabIndex        =   58
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox txtcessdebit 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1380
         TabIndex        =   53
         Top             =   2280
         Width           =   1440
      End
      Begin VB.CommandButton Command3 
         Caption         =   "..."
         Height          =   285
         Left            =   1080
         TabIndex        =   52
         Top             =   2295
         Width           =   300
      End
      Begin VB.TextBox txtcessdebitdesc 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3015
         TabIndex        =   51
         Top             =   2280
         Width           =   3225
      End
      Begin VB.TextBox txtcesscredit 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1395
         TabIndex        =   50
         Top             =   2760
         Width           =   1440
      End
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         Height          =   285
         Left            =   1080
         TabIndex        =   49
         Top             =   2760
         Width           =   315
      End
      Begin VB.TextBox txtcesscreditdesc 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3030
         TabIndex        =   48
         Top             =   2760
         Width           =   3225
      End
      Begin VB.TextBox txtCrAccName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3030
         TabIndex        =   46
         Top             =   1440
         Width           =   3225
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   285
         Left            =   1440
         TabIndex        =   45
         Top             =   1440
         Width           =   315
      End
      Begin VB.TextBox txtCrAccNo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1755
         TabIndex        =   44
         Top             =   1440
         Width           =   1080
      End
      Begin VB.TextBox lblDrAccName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3015
         TabIndex        =   43
         Top             =   960
         Width           =   3225
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "..."
         Height          =   285
         Left            =   1440
         TabIndex        =   42
         Top             =   975
         Width           =   300
      End
      Begin VB.TextBox txtDrAccNo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1740
         TabIndex        =   41
         Top             =   960
         Width           =   1080
      End
      Begin VB.TextBox txtPrice 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3720
         TabIndex        =   39
         Text            =   "0.00"
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtSubsidy 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   36
         Text            =   "0.00"
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label21 
         Caption         =   "Cess Rate:"
         Height          =   255
         Left            =   2160
         TabIndex        =   57
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label20 
         BackColor       =   &H008080FF&
         Caption         =   "Cess Details and Accounts"
         Height          =   255
         Left            =   120
         TabIndex        =   56
         Top             =   1920
         Width           =   2415
      End
      Begin VB.Label Label19 
         Caption         =   "Cess Debit"
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label18 
         Caption         =   "Cess Credit"
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   2760
         Width           =   855
      End
      Begin VB.Label Label17 
         Caption         =   "Dr Vehicle Accno"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label16 
         Caption         =   "Cr Stock  AccNo"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF80&
         Caption         =   "Price (Per Kg)"
         Height          =   195
         Left            =   2640
         TabIndex        =   38
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF80&
         Caption         =   "Subsidy (Per Kg)"
         Height          =   195
         Left            =   120
         TabIndex        =   37
         Top             =   360
         Width           =   1170
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF80&
      Caption         =   "Bank Details"
      Height          =   1335
      Left            =   120
      TabIndex        =   28
      Top             =   3960
      Width           =   7455
      Begin VB.TextBox txtAccNo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4920
         TabIndex        =   31
         Top             =   720
         Width           =   2415
      End
      Begin VB.ComboBox cboBName 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   30
         Top             =   720
         Width           =   2295
      End
      Begin VB.ComboBox cboBBranch 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2520
         TabIndex        =   29
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF80&
         Caption         =   "Bank Branch"
         Height          =   195
         Left            =   2640
         TabIndex        =   34
         Top             =   360
         Width           =   930
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF80&
         Caption         =   "Bank Name"
         Height          =   195
         Left            =   120
         TabIndex        =   33
         Top             =   360
         Width           =   840
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF80&
         Caption         =   "Account Number"
         Height          =   195
         Left            =   4920
         TabIndex        =   32
         Top             =   360
         Width           =   1200
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFF80&
      Caption         =   "Personal Details"
      Height          =   2175
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   7455
      Begin VB.TextBox txtTCode 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   20
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtNames 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   19
         Top             =   480
         Width           =   3615
      End
      Begin VB.TextBox txtId 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5040
         TabIndex        =   18
         Top             =   480
         Width           =   2295
      End
      Begin VB.ComboBox cboLocation 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1920
         TabIndex        =   17
         Top             =   1080
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
         Left            =   5520
         TabIndex        =   16
         Top             =   1680
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.ComboBox cboBranch 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Top             =   1680
         Width           =   2295
      End
      Begin VB.PictureBox Picture5 
         Height          =   255
         Left            =   960
         Picture         =   "frmDebtorsDetails.frx":0000
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   14
         Top             =   480
         Width           =   255
      End
      Begin MSComCtl2.DTPicker DTPRegDate 
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         _Version        =   393216
         Format          =   121831425
         CurrentDate     =   40096
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF80&
         Caption         =   "DCode"
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF80&
         Caption         =   "Names"
         Height          =   195
         Left            =   1200
         TabIndex        =   26
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF80&
         Caption         =   "Date registered"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   840
         Width           =   1080
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF80&
         Caption         =   "Id Number/Business No"
         Height          =   195
         Left            =   5040
         TabIndex        =   24
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF80&
         Caption         =   "Location"
         Height          =   195
         Left            =   1920
         TabIndex        =   23
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF80&
         Caption         =   "Vehicle No:"
         Height          =   195
         Left            =   480
         TabIndex        =   22
         Top             =   1440
         Width           =   1065
      End
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   375
      Left            =   480
      TabIndex        =   12
      Top             =   8520
      Width           =   735
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   5160
      TabIndex        =   11
      Top             =   8520
      Width           =   855
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   1920
      TabIndex        =   10
      Top             =   8520
      Width           =   735
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   375
      Left            =   1200
      TabIndex        =   9
      Top             =   8520
      Width           =   735
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFF80&
      Caption         =   "Contacts"
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   2280
      Width           =   7455
      Begin VB.TextBox txtPAddress 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox txtPhone 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2880
         TabIndex        =   3
         Top             =   480
         Width           =   3735
      End
      Begin VB.TextBox txtEmail 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   2655
      End
      Begin VB.TextBox txtTown 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2880
         TabIndex        =   1
         Top             =   1200
         Width           =   3735
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF80&
         Caption         =   "Town"
         Height          =   195
         Left            =   2880
         TabIndex        =   8
         Top             =   960
         Width           =   405
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF80&
         Caption         =   "Postal Address"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   1050
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF80&
         Caption         =   "Phone"
         Height          =   195
         Left            =   2880
         TabIndex        =   6
         Top             =   240
         Width           =   465
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF80&
         Caption         =   "E - Mail"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   525
      End
   End
End
Attribute VB_Name = "frmDebtorsDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cessapp As Integer
Dim newa As Integer

Private Sub chkcessapp_Click()
If chkcessapp = vbChecked Then
cessapp = 1
Else
cessapp = 0
End If
End Sub

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdedit_Click()
newa = 0
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
cmdsave.Enabled = True
End Sub

Private Sub cmdNew_Click()
newa = 1
Txtaccno = ""
txtEmail = ""
txtId = ""
txtNames = ""
txtPAddress = ""
txtPhone = ""
txtsubsidy = "0.00"
txtTCode = ""
txtTown = ""
cboBBranch.Text = ""
cboBName.Text = ""
cbolocation.Text = ""
cbobranch.Text = ""
txtPrice = "0.00"

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
cmdEdit.Enabled = False
'cmdSave.Enabled = False
cmdsave.Enabled = True
End Sub

Private Sub cmdsave_Click()
Dim Active As String
On Error GoTo ErrorHandler

If txtTCode = "" Then
MsgBox "Please enter the Debtor code ", vbInformation, "Missing Information"
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
'sql = ""
'sql = "set dateformat dmy SELECT * From d_Debtors where DCode ='" & txtTCode & "'"
'Set rs = oSaccoMaster.GetRecordset(sql)
'If rs.EOF Then
'      Set cn = New ADODB.Connection
'    sql = ""
'    sql = "SET dateformat DMY Update  d_Debtors SET DNmame= '" & txtNames & "',CertNo='" & txtId & "',Locations='" & cboLocation & "',TregDate='" & DTPRegDate & "',email='" & txtEmail & "',Phoneno='" & txtPhone & "',Town='" & txtTown & "',Address='" & txtPAddress & "',price=" & CCur(txtPrice) & ",Active=" & Active & ",AccDr='" & txtDrAccNo & "',AccCr='" & txtCrAccNo & "' where DCode='" & txtTCode & "'"
'    oSaccoMaster.ExecuteThis (sql)
'
'    Else
'     MsgBox "Debtor code already exist, Use a different code ", vbInformation, "Missing Information"
'   Exit Sub
'  Exit Sub
  
 If newa = 1 Then
    Set cn = New ADODB.Connection
    sql = ""
    sql = "d_sp_Debtors '" & txtTCode & "','" & txtNames & "','" & txtId & "','" & cbolocation & "','" & DTPRegDate & "','" & txtEmail & "','" & txtPhone & "','" & txtTown & "','" & txtPAddress & "'," & CCur(txtPrice) & "," & CCur(txtsubsidy) & ",'" & Txtaccno & "','" & cboBName & "'," & Active & ",'" & cboBBranch & "','" & cbobranch & "','" & User & "','" & txtDrAccNo & "','" & txtCrAccNo & "','" & txtcessrate & "','" & txtcessdebit & "','" & txtcesscredit & "'," & cessapp & ""
    oSaccoMaster.ExecuteThis (sql)
   Else
    Set cn = New ADODB.Connection
    sql = ""
    sql = "SET dateformat DMY Update  d_Debtors SET DName= '" & txtNames & "',CertNo='" & txtId & "',Locations='" & cbolocation & "',TregDate='" & DTPRegDate & "',email='" & txtEmail & "',Phoneno='" & txtPhone & "',Town='" & txtTown & "',Address='" & txtPAddress & "',price=" & CCur(txtPrice) & ",Active=" & Active & ",AccDr='" & txtDrAccNo & "',AccCr='" & txtCrAccNo & "' where DCode='" & txtTCode & "'"
    oSaccoMaster.ExecuteThis (sql)
 End If
cmdNew_Click
cmdsave.Enabled = False

MsgBox "Records successively updated."
Exit Sub
ErrorHandler:
MsgBox err.description

End Sub

Private Sub cmdSearch_Click()
frmSearchGLAccounts.Show vbModal, Me
    If Continue Then
        If SearchValue <> "" Then
            txtDrAccNo = SearchValue
            SearchValue = ""
        End If
    End If

End Sub

Private Sub Command1_Click()
frmSearchGLAccounts.Show vbModal, Me
    If Continue Then
        If SearchValue <> "" Then
            txtCrAccNo = SearchValue
            SearchValue = ""
        End If
    End If
End Sub

Private Sub Command2_Click()
frmSearchGLAccounts.Show vbModal, Me
    If Continue Then
        If SearchValue <> "" Then
            txtcesscredit = SearchValue
            SearchValue = ""
        End If
    End If

End Sub

Private Sub Command3_Click()
frmSearchGLAccounts.Show vbModal, Me
    If Continue Then
        If SearchValue <> "" Then
            txtcessdebit = SearchValue
            SearchValue = ""
        End If
    End If

End Sub

Private Sub Form_Load()
Dim myclass As cdbase
cessapp = 0
newa = 0
Txtaccno.Locked = False
txtEmail.Locked = False
txtId.Locked = False
txtNames.Locked = False
txtPAddress.Locked = False
txtPhone.Locked = False
txtsubsidy.Locked = False
'txtTCode.Locked = True
txtTown.Locked = False
cboBBranch.Locked = False
cboBName.Locked = False
cbolocation.Locked = False
cmdEdit.Enabled = False
cmdsave.Enabled = False

    
    Set rst = New Recordset
    'Dim cn As Connection
    Set cn = New ADODB.Connection
    Provider = "MAZIWA"
    'Provider = cnn
    cn.Open Provider, "atm", "atm"
    Set rst = New Recordset
    
    sql = "Select distinct(Locations) from   d_Debtors"
    rst.Open sql, cn, adOpenKeyset, adLockOptimistic
    While Not rst.EOF
    cbobranch.AddItem rst.Fields(0)
    rst.MoveNext
    Wend
    
    
 DTPRegDate = Format(Get_Server_Date, "dd/mm/yyyy")
    
    
    
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


End Sub

Private Sub Picture5_Click()
Me.MousePointer = vbHourglass
         frmSearchDebtors.Show vbModal
        txtTCode = sel
        txtTCode_Validate True
        Me.MousePointer = 0
End Sub



Private Sub Text1_Change()

End Sub

Private Sub txtcesscredit_Change()
 On Error GoTo SysError
    Dim Account As Acc_Details
    Editing = True
    Account = Get_Acc_Details(txtcesscredit, ErrorMessage)
    If Account.ACCNO <> "" Then
        txtcesscreditdesc = Account.AccName
    Else
        If ErrorMessage <> "" Then
            MsgBox ErrorMessage, vbInformation, Me.Caption
            ErrorMessage = ""
        End If
        txtcesscreditdesc = ""
    End If
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub txtcessdebit_Change()
 On Error GoTo SysError
    Dim Account As Acc_Details
    Editing = True
    Account = Get_Acc_Details(txtcessdebit, ErrorMessage)
    If Account.ACCNO <> "" Then
        txtcessdebitdesc = Account.AccName
    Else
        If ErrorMessage <> "" Then
            MsgBox ErrorMessage, vbInformation, Me.Caption
            ErrorMessage = ""
        End If
        txtcessdebitdesc = ""
    End If
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption

End Sub

Private Sub txtCrAccNo_Change()
 On Error GoTo SysError
    Dim Account As Acc_Details
        
        Editing = True
    Account = Get_Acc_Details(txtCrAccNo, ErrorMessage)
    If Account.ACCNO <> "" Then
        txtCrAccName = Account.AccName
    Else
        If ErrorMessage <> "" Then
            MsgBox ErrorMessage, vbInformation, Me.Caption
            ErrorMessage = ""
        End If
        txtCrAccName = ""
    End If
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub txtDrAccNo_Change()
    On Error GoTo SysError
    Dim Account As Acc_Details
    Editing = True
    Account = Get_Acc_Details(txtDrAccNo, ErrorMessage)
    If Account.ACCNO <> "" Then
        lblDrAccName = Account.AccName
    Else
        If ErrorMessage <> "" Then
            MsgBox ErrorMessage, vbInformation, Me.Caption
            ErrorMessage = ""
        End If
        lblDrAccName = ""
    End If
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption

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

Private Sub txtSubsidy_Click()
If Trim(txtsubsidy) = "0.00" Then
txtsubsidy = ""
End If

End Sub

Private Sub txtSubsidy_Validate(Cancel As Boolean)
If Trim(txtsubsidy) = "" Then
txtsubsidy = "0.00"
End If

txtsubsidy = Format(txtsubsidy, "#,##0.00")

End Sub

Private Sub txtTCode_Validate(Cancel As Boolean)
Dim a As Boolean, b As Integer
Set rs = New ADODB.Recordset
sql = "d_sp_Selectdebtors '" & txtTCode & "'"
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
If Not IsNull(rs.Fields(8)) Then txtsubsidy = Format(rs.Fields(8), "#0.00")
If Not IsNull(rs.Fields(9)) Then Txtaccno = rs.Fields(9)
If Not IsNull(rs.Fields(10)) Then cboBName = rs.Fields(10)
If Not IsNull(rs.Fields(11)) Then cboBBranch = rs.Fields(11)
If Not IsNull(rs.Fields(12)) Then a = rs.Fields(12)
If Not IsNull(rs.Fields(13)) Then cbobranch = rs.Fields(13)
If Not IsNull(rs.Fields(14)) Then txtPrice = Format(rs.Fields(14), "#0.00")
If Not IsNull(rs.Fields(15)) Then txtDrAccNo = rs.Fields(15)
If Not IsNull(rs.Fields(16)) Then txtCrAccNo = rs.Fields(16)
If Not IsNull(rs.Fields(17)) Then txtcessrate = rs.Fields(17)
If Not IsNull(rs.Fields(18)) Then txtcessdebit = rs.Fields(18)
If Not IsNull(rs.Fields(19)) Then txtcesscredit = rs.Fields(19)
If Not IsNull(rs.Fields(20)) Then b = rs.Fields(20)
If b = 1 Then
chkcessapp = vbChecked
Else
chkcessapp = vbUnchecked

End If
If a = True Then
chkActive = vbChecked
Else
chkActive = vbUnchecked
End If
cmdEdit.Enabled = True
cmdsave.Enabled = False
End If
End Sub

