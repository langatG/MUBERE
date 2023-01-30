VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAudit 
   Caption         =   "Audit Trail Report"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7335
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAudit.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   7335
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox CrystalReport1 
      Height          =   480
      Left            =   1080
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   10
      Top             =   2280
      Width           =   1200
   End
   Begin VB.ComboBox cboUserID 
      Height          =   315
      Left            =   630
      TabIndex        =   6
      Top             =   1440
      Width           =   2385
   End
   Begin MSComCtl2.DTPicker dtpStartDate 
      Height          =   315
      Left            =   1830
      TabIndex        =   3
      Top             =   585
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   " dd-MM-yyyy"
      Format          =   38862851
      CurrentDate     =   39408
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   405
      Left            =   2265
      TabIndex        =   1
      Top             =   2550
      Width           =   1320
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   405
      Left            =   3630
      TabIndex        =   0
      Top             =   2550
      Width           =   1320
   End
   Begin MSComCtl2.DTPicker dtpFinishDate 
      Height          =   315
      Left            =   3645
      TabIndex        =   5
      Top             =   585
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   " dd-MM-yyyy"
      Format          =   38862851
      CurrentDate     =   39408
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3090
      TabIndex        =   9
      Top             =   1185
      Width           =   915
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "User ID"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   630
      TabIndex        =   8
      Top             =   1185
      Width           =   630
   End
   Begin VB.Label lblUserName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   3060
      TabIndex        =   7
      Top             =   1440
      Width           =   3990
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Finish Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3675
      TabIndex        =   4
      Top             =   345
      Width           =   930
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Start Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1860
      TabIndex        =   2
      Top             =   345
      Width           =   885
   End
End
Attribute VB_Name = "frmAudit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboUserID_Change()
    On Error GoTo SysError
    Dim rsUser As New Recordset
    If Trim$(cboUserID) <> "" Then
        If cboUserID <> "All" Then
            Set rsUser = oSaccoMaster.GetRecordset("Select UserName from USERACCOUNTS " _
            & "Where UserLoginIDs='" & cboUserID & "'")
            With rsUser
                If .State = adStateOpen Then
                    If Not .EOF Then
                        lblUserName = IIf(IsNull(!username), "", !username)
                    Else
                        lblUserName = ""
                    End If
                End If
            End With
        Else
            lblUserName = "All Users"
        End If
    End If
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub cboUserID_Click()
    cboUserID_Change
End Sub

Private Sub cmdcancel_Click()
    Continue = False
    Unload Me
End Sub

Private Sub cmdOk_Click()
 On Error GoTo hell
    Startdate = DTPStartDate
    FinishDate = dtpFinishDate
    AuditUser = cboUserID
    FinishDate = DateAdd("d", 1, FinishDate)
       
        If AuditUser <> "All" Then
     
            STRFORMULA = "{AuditTrans.AUDITTIME}>=#" & Format(Startdate, "dd/mm/yyyy") & "# and {AuditTrans.AUDITTIME}<=#" _
            & Format(FinishDate, "dd/mm/yyyy") & "# and {AuditTrans.AUDITID}='" & Trim(AuditUser) & "'"
 
        Else
            STRFORMULA = "{AuditTrans.AUDITTIME}>=#" & Format(Startdate, "dd/mm/yyyy") & "# and {AuditTrans.AUDITTIME}<=#" _
            & Format(FinishDate, "dd/mm/yyyy") & "#"
        End If
 Set rs = oSaccoMaster.GetRecordset("select reportpath from reportpath")
 If Not rs.EOF Then
 PPP = rs!reportpath
 End If
 reportname = "Audit Trail.rpt"
   
    Show_Sales_Crystal_Report STRFORMULA, reportname, ""   'NB: This function slows down the generation of this Report, by JKChoge'
    STRFORMULA = ""
    AuditUser = ""
    Exit Sub
hell:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub Form_Load()
    On Error GoTo SysError
    Dim rsUser As New Recordset
    dtpFinishDate = Format(Get_Server_Date, " dd-MM-yyyy")
    DTPStartDate = Format(Get_Server_Date, " dd-MM-yyyy")
    cboUserID.AddItem "All"
    lblUserName = "All Users"
    Set rsUser = oSaccoMaster.GetRecordset("Select UserLoginIDs from USERACCOUNTS")
    With rsUser
        If .State = adStateOpen Then
            While Not .EOF
                cboUserID.AddItem IIf(IsNull(!UserLoginIDs), "", !UserLoginIDs)
                .MoveNext
            Wend
        End If
    End With
    cboUserID.ListIndex = 0
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub
