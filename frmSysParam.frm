VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSysParam 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Company Parameters"
   ClientHeight    =   7380
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11325
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSysParam.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7380
   ScaleWidth      =   11325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "THCP"
      Height          =   2295
      Left            =   6360
      TabIndex        =   39
      Top             =   4680
      Width           =   4815
      Begin VB.ComboBox cmdclient 
         Height          =   315
         ItemData        =   "frmSysParam.frx":000C
         Left            =   3720
         List            =   "frmSysParam.frx":000E
         Style           =   1  'Simple Combo
         TabIndex        =   51
         Text            =   "cmdclient"
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox txtbrcode 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3480
         TabIndex        =   50
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtcsms 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2280
         TabIndex        =   47
         Text            =   "15"
         Top             =   1800
         Width           =   855
      End
      Begin VB.CommandButton cmdtchp 
         Caption         =   "Update"
         Height          =   255
         Left            =   3720
         TabIndex        =   46
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox txtpremiumdate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2280
         TabIndex        =   45
         Text            =   "20"
         Top             =   1320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtdebitdate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2280
         TabIndex        =   43
         Text            =   "3"
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtdeductiondate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2280
         TabIndex        =   41
         Text            =   "25"
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label19 
         Caption         =   "BR Code"
         Height          =   255
         Left            =   3480
         TabIndex        =   49
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label16 
         Caption         =   "Warning SMS Date"
         Height          =   375
         Left            =   240
         TabIndex        =   48
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label15 
         Caption         =   "Premium D/I Date"
         Height          =   375
         Left            =   240
         TabIndex        =   44
         Top             =   1320
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label11 
         Caption         =   "Debit Date"
         Height          =   255
         Left            =   240
         TabIndex        =   42
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label10 
         Caption         =   "Deduction Date"
         Height          =   255
         Left            =   240
         TabIndex        =   40
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Other Company Parameters"
      Height          =   4095
      Left            =   6360
      TabIndex        =   29
      Top             =   240
      Width           =   4695
      Begin MSComCtl2.DTPicker DTPTimeSend 
         Height          =   255
         Left            =   1920
         TabIndex        =   38
         Top             =   2880
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   450
         _Version        =   393216
         Format          =   121831426
         CurrentDate     =   40624
      End
      Begin VB.TextBox txtCost 
         Height          =   285
         Left            =   1920
         TabIndex        =   37
         Text            =   "0.00"
         Top             =   2400
         Width           =   2295
      End
      Begin VB.TextBox txtSMSNumber 
         Height          =   285
         Left            =   1920
         MaxLength       =   13
         TabIndex        =   36
         Text            =   "0000000000"
         Top             =   1920
         Width           =   2295
      End
      Begin VB.TextBox txtcompanymotto 
         Appearance      =   0  'Flat
         Height          =   735
         Left            =   240
         TabIndex        =   32
         Text            =   "Text1"
         Top             =   960
         Width           =   4215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Disable Entering of Opening Balances"
         Height          =   375
         Left            =   240
         TabIndex        =   30
         Top             =   3480
         Width           =   3975
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "SMS Send Time : "
         Height          =   195
         Left            =   240
         TabIndex        =   35
         Top             =   2880
         Width           =   1545
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "SMS Cost : "
         Height          =   195
         Left            =   240
         TabIndex        =   34
         Top             =   2400
         Width           =   1020
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "SMS Number : "
         Height          =   195
         Left            =   240
         TabIndex        =   33
         Top             =   1920
         Width           =   1305
      End
      Begin VB.Label Label6 
         Caption         =   "Company Moto"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.TextBox txtWebsite 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   6450
      Width           =   4155
   End
   Begin VB.TextBox txtPhone 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1800
      TabIndex        =   27
      Top             =   5040
      Width           =   4215
   End
   Begin VB.TextBox txtEmail 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1800
      TabIndex        =   26
      Top             =   5520
      Width           =   4215
   End
   Begin VB.TextBox txtFax 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1800
      TabIndex        =   25
      Top             =   5925
      Width           =   4215
   End
   Begin VB.TextBox txtTown 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   1920
      Width           =   4155
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Cancel Process"
      Top             =   6960
      Width           =   855
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   375
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Save Record"
      Top             =   6960
      Width           =   855
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   375
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Edit Record"
      Top             =   6960
      Width           =   855
   End
   Begin VB.CommandButton cmdClose 
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
      Height          =   390
      Left            =   5040
      TabIndex        =   6
      Top             =   6960
      Width           =   1110
   End
   Begin VB.Frame fraBanking 
      Caption         =   "General Company Info"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6795
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   5895
      Begin VB.TextBox txtDivision 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         TabIndex        =   17
         Top             =   3885
         Width           =   4215
      End
      Begin VB.TextBox txtDistrict 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         TabIndex        =   15
         Top             =   3360
         Width           =   4215
      End
      Begin VB.TextBox txtProvince 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         TabIndex        =   14
         Top             =   2880
         Width           =   4215
      End
      Begin VB.TextBox txtCountry 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         TabIndex        =   12
         Top             =   2280
         Width           =   4215
      End
      Begin VB.TextBox txtLocation 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   4410
         Width           =   4155
      End
      Begin VB.TextBox txtCompanyName 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   645
         Width           =   5715
      End
      Begin VB.TextBox txtPostalAddress 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1155
         Width           =   4155
      End
      Begin VB.Label Label5 
         Caption         =   "Website"
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
         Left            =   120
         TabIndex        =   24
         Top             =   6360
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Fax"
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
         Left            =   135
         TabIndex        =   23
         Top             =   5925
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Telephone"
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
         TabIndex        =   22
         Top             =   4920
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "E - Mail"
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
         Left            =   135
         TabIndex        =   21
         Top             =   5430
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Town"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   19
         Top             =   1800
         Width           =   1020
      End
      Begin VB.Label Label21 
         Caption         =   "Division"
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
         TabIndex        =   18
         Top             =   3900
         Width           =   735
      End
      Begin VB.Label Label20 
         Caption         =   "District"
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
         Left            =   105
         TabIndex        =   16
         Top             =   3390
         Width           =   855
      End
      Begin VB.Label Label18 
         Caption         =   "Province"
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
         TabIndex        =   13
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label17 
         Caption         =   "Country"
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
         TabIndex        =   11
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label14 
         Caption         =   "Company Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   120
         TabIndex        =   10
         Top             =   315
         Width           =   1890
      End
      Begin VB.Label Label13 
         Caption         =   "Postal address"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   1140
      End
      Begin VB.Label Label12 
         Caption         =   "Location"
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
         Left            =   120
         TabIndex        =   8
         Top             =   4395
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmSysParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myclass As cdbase

Private Sub cmdcancel_Click()
Form_Load
End Sub

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdedit_Click()
txtCompanyName.Locked = False
txtCountry.Locked = False
txtDistrict.Locked = False
txtDivision.Locked = False
txtEmail.Locked = False
txtFax.Locked = False
txtLocation.Locked = False
txtPhone.Locked = False
txtPostalAddress.Locked = False
txtProvince.Locked = False
txtTown.Locked = False
txtWebsite.Locked = False
txtcompanymotto.Locked = False
cmdupdate.Enabled = True
cmdEdit.Enabled = False

txtSMSNumber.Locked = False

txtCost.Locked = False

End Sub

Private Sub cmdtchp_Click()
'//here if the server is one then run all else run only the report.
sql = "UPDATE    d_company  SET        server=" & cmdclient & ",      ddate=" & txtdebitdate & ",deddate=" & txtdeductiondate & ",pdate=" & txtpremiumdate & ",CSMSDate=" & txtcsms & ",brcode='" & txtbrcode & "'"

oSaccoMaster.ExecuteThis (sql)
MsgBox "Updated successfully"
Exit Sub
End Sub

Private Sub cmdupdate_Click()
On Error GoTo ErrorHandler

Set cn = New ADODB.Connection
sql = ""
sql = "d_sp_UpdateCProfile'" & txtCompanyName & "','" & txtPostalAddress & "','" & txtTown & "','" & txtCountry & "','" & txtProvince & "','" & txtDistrict & "','" & txtDivision & "','" & txtLocation & "','" & txtFax & "','" & txtPhone & "','" & txtEmail & "','" & txtWebsite & "','" & User & "','" & txtcompanymotto & "','" & DTPTimeSend & "'"

oSaccoMaster.ExecuteThis (sql)
Form_Load

MsgBox "Records successively updated."
Exit Sub
ErrorHandler:
MsgBox err.description
End Sub

Private Sub Command1_Click()
'//take of new opening after entering
On Error GoTo ErrorHandler
sql = ""
sql = "update d_company set acc=1"
oSaccoMaster.ExecuteThis (sql)
Exit Sub
ErrorHandler:
MsgBox err.description
End Sub

Private Sub Form_Load()
txtCompanyName.Locked = True
txtCountry.Locked = True
txtDistrict.Locked = True
txtDivision.Locked = True
txtEmail.Locked = True
txtFax.Locked = True
txtLocation.Locked = True
txtPhone.Locked = True
txtPostalAddress.Locked = True
txtProvince.Locked = True
txtTown.Locked = True
txtWebsite.Locked = True
txtcompanymotto.Locked = True
txtcompanymotto.Locked = True
cmdEdit.Enabled = True

cmdupdate.Enabled = False

txtSMSNumber.Locked = True
'DTPTimeSend.Locked = True
txtCost.Enabled = True
loadCompanyParam

End Sub
Public Sub loadCompanyParam()

Set myclass = New cdbase
Set cn = CreateObject("adodb.connection")
Provider = myclass.OpenCon
cn.Open Provider, "atm", "atm"
Set rs = CreateObject("adodb.recordset")
sql = "select * from d_company"
rs.Open sql, cn
If Not rs.EOF Then
txtCompanyName = rs![name]
txtCountry = rs!Country
txtDistrict = rs!District
txtDivision = rs!Division
txtEmail = rs!Email
txtFax = rs!FaxNo
txtLocation = rs!Location
txtPhone = rs!PhoneNo
txtPostalAddress = rs!Adress
txtProvince = rs!Province
txtTown = rs!town
txtWebsite = rs!Website
'txtdebitdate = IIf(IsNull(rs!ddate), "3", rs!ddate)
'txtdeductiondate = IIf(IsNull(rs!deddate), "28", rs!deddate)
'txtpremiumdate = IIf(IsNull(rs!pdate), "20", rs!pdate)
'txtSMSNumber = IIf(IsNull(rs!SMSNo), "", rs!SMSNo)
'txtCost = IIf(IsNull(rs!SMSCost), "", rs!SMSCost)
'txtcsms = IIf(IsNull(rs!csmsdate), "", rs!csmsdate)
'txtbrcode = IIf(IsNull(rs!code), "", rs!code)

'CSMSDate
'DTPTimeSend = IIf(IsNull(rs!SendTime), Format(Get_Server_Date, Time), rs!SendTime)

'txtcompanymotto = IIf(IsNull(rs!motto), "", rs!motto)
'cmdclient.Text = IIf(IsNull(rs!server), 1, rs!server)
End If

End Sub
