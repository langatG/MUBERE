VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmglsetup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GENERAL LEDGER SET UP"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6600
   Icon            =   "frmglsetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   6600
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdclose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5265
      TabIndex        =   4
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4098
      TabIndex        =   3
      Top             =   6600
      Width           =   975
   End
   Begin VB.CommandButton cmdedits 
      Caption         =   "&Edit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2692
      TabIndex        =   2
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton cmdsave 
      Appearance      =   0  'Flat
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1406
      TabIndex        =   1
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton cmdnew 
      Appearance      =   0  'Flat
      Caption         =   "&New "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   6600
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   6495
      Left            =   120
      TabIndex        =   5
      Top             =   30
      Width           =   6255
      Begin VB.CommandButton cmdChange 
         Caption         =   "Change"
         Height          =   375
         Left            =   4560
         TabIndex        =   34
         Top             =   360
         Width           =   975
      End
      Begin VB.CheckBox chkSuspense 
         Caption         =   "Is The Suspense Account"
         Height          =   375
         Left            =   240
         TabIndex        =   33
         Top             =   5520
         Width           =   2535
      End
      Begin VB.CheckBox chkRetainedEarning 
         Caption         =   "Is the Retained Earning Account"
         Height          =   375
         Left            =   240
         TabIndex        =   32
         Top             =   6000
         Width           =   2895
      End
      Begin VB.ComboBox cboSubType 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmglsetup.frx":030A
         Left            =   1860
         List            =   "frmglsetup.frx":031A
         TabIndex        =   30
         Top             =   5040
         Width           =   2175
      End
      Begin VB.PictureBox Picture5 
         Height          =   285
         Left            =   4080
         Picture         =   "frmglsetup.frx":032F
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   29
         Top             =   480
         Width           =   285
      End
      Begin VB.ComboBox cboAccType 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmglsetup.frx":05F1
         Left            =   1860
         List            =   "frmglsetup.frx":05FB
         TabIndex        =   27
         Top             =   3960
         Width           =   2115
      End
      Begin MSComCtl2.DTPicker dtpTransDate 
         Height          =   315
         Left            =   4380
         TabIndex        =   23
         Top             =   4560
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   " dd-MM-yyyy"
         Format          =   107085827
         CurrentDate     =   39532
      End
      Begin VB.TextBox txtOpeningBalance 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4335
         TabIndex        =   21
         Text            =   "0"
         Top             =   5130
         Width           =   1830
      End
      Begin VB.ComboBox cboacccategory 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmglsetup.frx":0610
         Left            =   1800
         List            =   "frmglsetup.frx":0626
         TabIndex        =   19
         Top             =   4560
         Width           =   2175
      End
      Begin VB.ComboBox cbocurrency 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmglsetup.frx":068A
         Left            =   1860
         List            =   "frmglsetup.frx":06A0
         TabIndex        =   17
         Text            =   "KES"
         Top             =   3474
         Width           =   2055
      End
      Begin VB.TextBox txtaccname 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1860
         MaxLength       =   50
         TabIndex        =   10
         Top             =   876
         Width           =   3615
      End
      Begin VB.TextBox txtaccno 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1860
         MaxLength       =   20
         TabIndex        =   9
         Top             =   443
         Width           =   2160
      End
      Begin VB.ComboBox cboaccoounttype 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmglsetup.frx":06C2
         Left            =   1860
         List            =   "frmglsetup.frx":06CF
         TabIndex        =   8
         Top             =   1742
         Width           =   3975
      End
      Begin VB.ComboBox cboaccountgroup 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmglsetup.frx":0707
         Left            =   1860
         List            =   "frmglsetup.frx":0756
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   2610
         Width           =   3975
      End
      Begin VB.ComboBox cbonormalbalance 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmglsetup.frx":0902
         Left            =   1860
         List            =   "frmglsetup.frx":090C
         TabIndex        =   6
         Top             =   3041
         Width           =   3975
      End
      Begin VB.ComboBox cboAccGroup 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmglsetup.frx":091F
         Left            =   1845
         List            =   "frmglsetup.frx":0938
         Sorted          =   -1  'True
         TabIndex        =   25
         Top             =   2175
         Width           =   1815
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Sub Category"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   31
         Top             =   5085
         Width           =   1110
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Account Type"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   28
         Top             =   4005
         Width           =   1170
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Account Sub Group"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   105
         TabIndex        =   26
         Top             =   2660
         Width           =   1620
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Balance As At"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4590
         TabIndex        =   24
         Top             =   4230
         Width           =   1125
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Opening Balance:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4740
         TabIndex        =   22
         Top             =   4875
         Width           =   1425
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Acc Category"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   20
         Top             =   4560
         Width           =   1095
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Currency"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   18
         Top             =   3526
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Gl Account name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   15
         Top             =   928
         Width           =   1410
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "GL Account Number"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   14
         Top             =   495
         Width           =   1665
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Account Type"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   13
         Top             =   1794
         Width           =   1170
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Account Group"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   12
         Top             =   2227
         Width           =   1245
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Normal Balance"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   11
         Top             =   3093
         Width           =   1230
      End
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000008&
      Height          =   4185
      Left            =   120
      TabIndex        =   16
      Top             =   840
      Width           =   6255
   End
End
Attribute VB_Name = "frmglsetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ed As Boolean
Dim sta As Integer
Dim cn As Connection
Dim newg As Integer
Dim myclass As Object
Dim Provider As String

Private Sub Get_Account_Details(ACCNO As String)
    Dim rsAccounts As New Recordset
    On Error GoTo SysError
    Set rsAccounts = oSaccoMaster.GetRecordset("Select * From GLSETUP where " _
    & "AccNo='" & ACCNO & "'")
    With rsAccounts
        If .State = adStateOpen Then
            If Not .EOF Then
                txtaccname = IIf(IsNull(!GlAccName), "", !GlAccName)
                txtaccno = ACCNO
                txtOpeningBalance = Format(IIf(IsNull(!NewGLOpeningBal), 0, !NewGLOpeningBal), Cfmt)
                cboacccategory = IIf(IsNull(!AccCategory), "", !AccCategory)
                cboaccoounttype = IIf(IsNull(!Glacctype), "", !Glacctype)
                cboaccountgroup = IIf(IsNull(!GLAccGroup), "", !GLAccGroup)
                cbocurrency = IIf(IsNull(!Curr), "", !Curr)
                cbonormalbalance = IIf(IsNull(!NormalBal), "", !NormalBal)
                cboAccGroup = IIf(IsNull(!GlAccMainGroup), "", !GlAccMainGroup)
                chkRetainedEarning = IIf(!isREarning = True, vbChecked, vbUnchecked)
                chkSuspense = IIf(!issuspense = True, vbChecked, vbUnchecked)
                cboAccType.Text = IIf(IsNull(!Type), "", !Type)
                cboSubType.Text = IIf(IsNull(!subType), "", !subType)
                dtpTransDate.value = !newglopeningbaldate
            Else
                txtaccname = ""
                txtOpeningBalance = Format(0, Cfmt)
                cboacccategory = ""
                cboaccoounttype = ""
                cboaccountgroup = ""
                cbocurrency = ""
                cbonormalbalance = ""
                cboAccGroup = ""
            End If
        End If
    End With
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub cboAccGroup_Change()
    Select Case cboAccGroup
        Case "ASSETS"
            cbonormalbalance.Text = "Debit"
        Case "LIABILITIES"
            cbonormalbalance.Text = "Credit"
        Case "INCOME"
            cbonormalbalance.Text = "Credit"
        Case "EXPENSES"
            cbonormalbalance.Text = "Debit"
        Case "RETAINED EARNINGS"
            cbonormalbalance.Text = "Credit"
        Case "CAPITAL"
            cbonormalbalance.Text = "Credit"
    End Select
End Sub

Private Sub cboAccGroup_Click()
    cboAccGroup_Change
End Sub



Private Sub chkRetainedEarning_Click()
    With chkSuspense
        If .value = vbChecked Then
            chkSuspense.value = vbUnchecked
        End If
    End With
End Sub

Private Sub chkSuspense_Click()
    With chkSuspense
        If .value = vbChecked Then
            chkRetainedEarning.value = vbUnchecked
        End If
    End With
End Sub

Private Sub cmdCancel_Click()
   sql = "SELECT isnull(availableBalance,0) from cub WHERE Accno='" & txtaccno & "'"
   Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF = True Then
        If rs(0) > 0 Then
            MsgBox "This account have balances and your cannot delete it", vbCritical
            Exit Sub
        End If
    Else
        oSaccoMaster.ExecuteThis ("Delete from glsetup where accno='" & txtaccno & "'")
        If success = True Then
            MsgBox "Delete successfully"
        End If
    End If
End Sub

Private Sub cmdChange_Click()
    Dim oldAccNo As String, NewAccNo As String
    If MsgBox("This activity is irreversible, proceed?", vbQuestion + vbYesNo) = vbNo Then
        Exit Sub
    End If
    
    oldAccNo = txtaccno
    NewAccNo = InputBox("Enter the new AccNo: ", "NEW ACCNO", "")
    
    If NewAccNo = "" Then
        MsgBox "Cant update with an Empty AccNo! ", vbExclamation
        Exit Sub
    End If
    
    'Begin with the setups
    
    'glsetup
     oSaccoMaster.ExecuteThis ("Update glsetup set accno='" & NewAccNo & "' where accno='" & oldAccNo & "' ")
    'Loantypes
    Set rst = oSaccoMaster.GetRecordset("select loanacc from loantype where loanAcc='" & oldAccNo & "'")
    If Not rst.EOF Then
        If Not oSaccoMaster.GetRecordset("Update loantypes set LoanAcc='" & NewAccNo & "' where loanacc='" & oldAccNo & "' ") Then
            MsgBox ErrorMessage
        End If
    End If
    
    Set rst = oSaccoMaster.GetRecordset("select interestAcc from loantype where interestAcc='" & oldAccNo & "'")
    If Not rst.EOF Then
        If Not oSaccoMaster.GetRecordset("Update loantype set interestAcc='" & NewAccNo & "' where interestAcc='" & oldAccNo & "' ") Then
            MsgBox ErrorMessage
        End If
    End If
    
    'sharetype
    
    Set rst = oSaccoMaster.GetRecordset("select sharesacc from sharetype where sharesacc='" & oldAccNo & "'")
    If Not rst.EOF Then
        If Not oSaccoMaster.GetRecordset("Update sharetype set sharesacc='" & NewAccNo & "' where sharesacc='" & oldAccNo & "' ") Then
            MsgBox ErrorMessage
        End If
    End If
    
    'sysparam
    
    Set rst = oSaccoMaster.GetRecordset("select SuspenseAcc from Sysparam where suspenseacc='" & oldAccNo & "'")
    If Not rst.EOF Then
         oSaccoMaster.ExecuteThis ("Update sysparam set SuspenseAcc='" & NewAccNo & "' where SuspenseAcc='" & oldAccNo & "' ")
    End If
    
    Set rst = oSaccoMaster.GetRecordset("select RearningsAcc from Sysparam where RearningsAcc='" & oldAccNo & "'")
    If Not rst.EOF Then
         oSaccoMaster.ExecuteThis ("Update sysparam set RearningsAcc='" & NewAccNo & "' where RearningsAcc='" & oldAccNo & "' ")
    End If
    
    'gltransactions
    
    Set rst = oSaccoMaster.GetRecordset("select draccno from gltransactions where draccno='" & oldAccNo & "'")
    If Not rst.EOF Then
         oSaccoMaster.ExecuteThis ("Update gltransactions set draccno='" & NewAccNo & "' where draccno='" & oldAccNo & "' ")
    End If
    
    Set rst = oSaccoMaster.GetRecordset("select craccno from gltransactions where craccno='" & oldAccNo & "'")
    If Not rst.EOF Then
         oSaccoMaster.ExecuteThis ("Update gltransactions set craccno='" & NewAccNo & "' where craccno='" & oldAccNo & "' ")
    End If
    
    'Banks
    
    Set rst = oSaccoMaster.GetRecordset("select accno from Banks where accno='" & oldAccNo & "'")
    If Not rst.EOF Then
        If Not oSaccoMaster.GetRecordset("Update banks set accno='" & NewAccNo & "' where accno='" & oldAccNo & "' ") Then
            MsgBox ErrorMessage
        End If
    End If
    
    MsgBox "Done! Accno :" & oldAccNo & " Changed to: " & NewAccNo & " Successfully!", vbInformation
    txtAccno_Change
    txtaccno.Text = NewAccNo
    
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdedits_Click()
    newg = 0
    NewRecord = False
    cmdnew.Enabled = True
    cmdsave.Enabled = True
    cmdedits.Enabled = False
End Sub

Private Sub cmdNew_Click()
    On Error GoTo SysError
    
    newg = 1
    txtaccno.Enabled = True
    txtaccname.Enabled = True
    txtOpeningBalance.Enabled = True
    cboacccategory.Enabled = True
    cboaccoounttype.Enabled = True
    cboaccountgroup.Enabled = True
    cbocurrency.Enabled = True
    cbonormalbalance.Enabled = True
    
    
    ClearMe
    NewRecord = True
    cmdnew.Enabled = False
    cmdsave.Enabled = True
    cmdedits.Enabled = True
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
    NewRecord = False
End Sub

Private Sub ClearMe()
    txtaccno = ""
    txtaccname = ""
    txtOpeningBalance = "0.00"
    cboacccategory = ""
    cboaccoounttype = ""
    cboaccountgroup = ""
    cbocurrency = ""
    cbonormalbalance = ""
End Sub

Private Sub cmdsave_Click()
    On Error GoTo Capture
    Dim sta As Integer
    If txtaccno = "" Then
        MsgBox "AccountNo can not be blank"
        Exit Sub
    End If
        
'    sql = "SELECT  GlCode, GlAccName, AccNo, GlAccType, AuditID, AuditDate From GLSETUP where ACCNO= '" & Txtaccno & "'  "
'    Set rs = oSaccoMaster.GetRecordset(sql)
'    If Not rs.EOF Then
'    'Cmbstation = rs!Branch
'    If rs!ACCNO = Txtaccno Then
'    MsgBox "You Need to edit in order to save", vbInformation
'    Exit Sub
'    End If
'    End If
        
    If cboAccType.Text = "" Then
        MsgBox "Account type cannot is not Optional, please enter before you proceed", vbExclamation
        Exit Sub
    ElseIf cboSubType.Text = "" Then
        MsgBox "Account Sub Type cannot is not Optional, please enter before you proceed", vbExclamation
        Exit Sub
    End If
    If Trim$(cbonormalbalance) = "" Then
        MsgBox "Please Indicate the Normal Balance for this Account", vbExclamation, Me.Caption
        Exit Sub
    End If
    If Trim(cboaccoounttype) = "" Then
        MsgBox "Please Indicate the Account Type for this Account", vbExclamation, Me.Caption
        Exit Sub
    End If
    If Trim(cboaccountgroup) = "" Then
        MsgBox "Please Indicate the Account Sub Group for this Account", vbExclamation, Me.Caption
        Exit Sub
    End If
    If Trim(cboAccGroup) = "" Then
        MsgBox "Please Indicate the Account Group for this Account", vbExclamation, Me.Caption
        Exit Sub
    End If

'Dim idno As String, Phone As String, ans As String, NAMES As String
If newg = 1 Then
sql = "SELECT     GlAccName, subtype, AccNo, NormalBal  FROM         GLSETUP  WHERE     (AccNo = '" & txtaccno & "')"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
'check if the idno exists
        sql = "SELECT GlAccName, subtype, AccNo, NormalBal  FROM         GLSETUP  WHERE     (AccNo = '" & txtaccno & "') "
        Set rst = oSaccoMaster.GetRecordset(sql)
        If Not rst.EOF Then
        'ans = MsgBox("The ID no is already available in the system, do you want to proceed with this update", vbYesNo)
        MsgBox "GL AccNo Cannot be duplicate "
        Exit Sub
            'If ans = vbYes Then
            'GoTo next1
            'Else
            'Exit Sub
            'End If
        End If
End If
End If



    Dim rsAccounts As New Recordset
    Set rsAccounts = oSaccoMaster.GetRecordset("Select * From GLSETUP where AccNo='" & txtaccno & "'")
    
   
    With rsAccounts
        If Not .EOF And Not NewRecord Then
            If MsgBox("The Gl Accno is already existing. Did you intend to change its details?", vbQuestion + vbYesNo) = vbNo Then
                Exit Sub
            End If
          
            
        Else
        

            .AddNew
            !ACCNO = txtaccno.Text
        End If
'        !GlCode = IIf(Trim$(txtcode) <> "", txtcode, "")
        !GlAccName = IIf(Trim$(txtaccname) <> "", txtaccname, "")
        !Glacctype = IIf(Trim$(cboaccoounttype) <> "", cboaccoounttype, "")
        !GLAccGroup = IIf(Trim$(cboaccountgroup) <> "", cboaccountgroup, "")
        !GlAccMainGroup = IIf(Trim$(cboAccGroup) <> "", cboAccGroup, "")
        !NormalBal = IIf(Trim$(cbonormalbalance) <> "", cbonormalbalance, "")
'        !GlAccStatus = IIf(Optactive <> True, 0, 1)
        !OpeningBal = IIf(Trim$(txtOpeningBalance) <> "", CDbl(txtOpeningBalance), 0)
        !bal = !CurrentBal
        !transdate = dtpTransDate
        !Curr_Code = cbocurrency.ListIndex
        !AuditOrg = ""
        !auditid = User
        !AuditDate = Get_Server_Date
        !Curr = IIf(Trim$(cbocurrency) <> "", cbocurrency, "")
        !Actuals = 0
        !Budgetted = 0
        !IsSubLedger = 0
        !Type = cboAccType.Text
        !AccCategory = IIf(Trim$(cboacccategory) <> "", cboacccategory, "")
        !NewGLOpeningBal = IIf(Trim$(txtOpeningBalance) <> "", CDbl(txtOpeningBalance), 0)
        !newglopeningbaldate = dtpTransDate
        !subType = cboSubType.Text
        !issuspense = IIf(chkSuspense.value = vbChecked, 1, 0)
        !isREarning = IIf(chkRetainedEarning.value = vbChecked, 1, 0)
        !CurrentBal = 0
        .Update
        
        If chkSuspense.value = vbChecked Then
            SuspenseAcc = txtaccno
             oSaccoMaster.ExecuteThis ("Update glsetup set isSuspense=0 where accno <>'" & SuspenseAcc & "'")
        End If
        If chkRetainedEarning.value = vbChecked Then
            REarningsAcc = txtaccno
             oSaccoMaster.ExecuteThis ("Update glsetup set isREarning=0 where accno <>'" & REarningsAcc & "'")
        End If
        
        MsgBox "Record Saved Successfully", vbInformation, Me.Caption
    End With
    cmdnew.Enabled = True
    cmdsave.Enabled = False
    cmdedits.Enabled = True
    
    
    txtaccno.Enabled = False
    txtaccname.Enabled = False
    txtOpeningBalance.Enabled = False
    cboacccategory.Enabled = False
    cboaccoounttype.Enabled = False
    cboaccountgroup.Enabled = False
    cbocurrency.Enabled = False
    cbonormalbalance.Enabled = False
    
    Exit Sub
 
  Exit Sub
Capture:
  MsgBox IIf(ErrorMessage = "", err.description, ErrorMessage)
    'Form_Load
End Sub

Private Sub Form_Load()
'    Optactive = True
    newg = 0
    Editing = False
    dtpTransDate = Format(Get_Server_Date, " dd-MM-yyyy")
    cmdnew.Enabled = True
    cmdsave.Enabled = False
    cmdedits.Enabled = True
End Sub
'Private Sub OptACTIVE_Click()
'If Optactive = True Then
'sta = 0
'End If
'End Sub

'Private Sub Optinactive_Click()
'
'If Optinactive = True Then
'sta = 1
'End If

'End Sub


Private Sub Picture5_Click()
    On Error GoTo SysError
    frmSearchGLAccounts.Show vbModal, Me
    If Continue Then
        If SearchValue <> "" Then
            txtaccno = SearchValue
        End If
    End If
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub


Private Sub txtaccname_KeyPress(KeyAscii As Integer)
    On Error GoTo errFix
    If KeyAscii <> vbKeyReturn Then 'Catch the Enter key
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
    Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Member Registration"
End Sub

Private Sub txtAccno_Change()
    On Error GoTo SysError
        If Trim$(txtaccno) <> "" Then
            Get_Account_Details txtaccno
            Editing = False
        End If
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub txtAccNo_KeyPress(KeyAscii As Integer)
    On Error GoTo SysError
    If KeyAscii <> vbKeyReturn Then 'Catch the Enter key
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub


Private Sub txtOpeningBalance_KeyPress(KeyAscii As Integer)
    On Error GoTo SysError
    Select Case KeyAscii
        Case 48 To 57
        Case Is = 46
        Case Is = 45
        Case Is = 8
        Case Else
        KeyAscii = 0
    End Select
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub
