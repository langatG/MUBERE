VERSION 5.00
Begin VB.Form frmcurrency1 
   Caption         =   "CURRENCY CODES"
   ClientHeight    =   5565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6000
   Icon            =   "frmcurrency1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5565
   ScaleWidth      =   6000
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtrate 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2160
      TabIndex        =   26
      Top             =   1920
      Width           =   1215
   End
   Begin VB.ComboBox cbonegativedisplay 
      Height          =   315
      ItemData        =   "frmcurrency1.frx":0442
      Left            =   2160
      List            =   "frmcurrency1.frx":044F
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   4440
      Width           =   2175
   End
   Begin VB.ComboBox cbodecimalseparator 
      Height          =   315
      ItemData        =   "frmcurrency1.frx":0477
      Left            =   2160
      List            =   "frmcurrency1.frx":0481
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Top             =   3960
      Width           =   2175
   End
   Begin VB.ComboBox cbothousandseparator 
      Height          =   315
      ItemData        =   "frmcurrency1.frx":0494
      Left            =   2160
      List            =   "frmcurrency1.frx":04A1
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   3480
      Width           =   2175
   End
   Begin VB.ComboBox cbosymbolposition 
      Height          =   315
      ItemData        =   "frmcurrency1.frx":04BB
      Left            =   2160
      List            =   "frmcurrency1.frx":04CB
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   3000
      Width           =   2175
   End
   Begin VB.ComboBox cbodecimalposition 
      Height          =   315
      ItemData        =   "frmcurrency1.frx":051F
      Left            =   2160
      List            =   "frmcurrency1.frx":052F
      TabIndex        =   20
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   3960
      TabIndex        =   19
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton cmddelete 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   2040
      TabIndex        =   18
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   720
      TabIndex        =   17
      Top             =   5040
      Width           =   1095
   End
   Begin VB.TextBox txtcurrencysymbol 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2160
      TabIndex        =   16
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox txtdescription 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2160
      TabIndex        =   15
      Top             =   960
      Width           =   3135
   End
   Begin VB.CommandButton cmdfinder 
      Height          =   375
      Left            =   4200
      Picture         =   "frmcurrency1.frx":053F
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Add New record"
      Top             =   240
      Width           =   375
   End
   Begin VB.CommandButton cmdAdd 
      Height          =   375
      Left            =   4560
      Picture         =   "frmcurrency1.frx":0801
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Add New record"
      Top             =   240
      Width           =   375
   End
   Begin VB.TextBox txtcurrcode 
      Height          =   375
      Left            =   2880
      TabIndex        =   0
      Top             =   240
      Width           =   615
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      ToolTipText     =   "Move to the Last record"
      Top             =   240
      Width           =   375
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      ToolTipText     =   "Move to the Previous record"
      Top             =   240
      Width           =   375
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      Picture         =   "frmcurrency1.frx":0D33
      TabIndex        =   3
      ToolTipText     =   "Move to Last record"
      Top             =   240
      Width           =   375
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      ToolTipText     =   "Move to the Next"
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label9 
      Caption         =   "Rate Aganist Source"
      Height          =   255
      Left            =   240
      TabIndex        =   25
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label8 
      Caption         =   "Negative Display"
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "Decimal Separator"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "Thousand Separator"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "Symbol Position"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Decimal Places"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Symbol"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Description"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Currency Code"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "frmcurrency1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdfinder_Click()
On Error Resume Next
frmsearchcurr.Show vbModal
Dim Y As String
Y = sel
'm = False
If Y <> "" Then
     Dim cn As Connection
    Set cn = New ADODB.Connection
    Provider = SelectedDsn
   cn.Open Provider, "atm", "atm"
sql = ""
'If reportpath = "" Then reportpath = GetSetting("payroll", "AppName", "rptPath", rptPath)
sql = "SELECT     CurrCode, Description, Symbol, rateaganistsource, Decimals, Symbolposition, ThousandSeparator,decimalseparator, NegativeDisplay, auditid, auditdatetime From Curr where currcode='" & Y & "'"
Set rs = New ADODB.Recordset
rs.Open sql, cn
If Not rs.EOF Then

If Not IsNull(rs.Fields(0)) Then txtcurrcode = (rs.Fields(0))
If Not IsNull(rs.Fields(1)) Then txtdescription = (rs.Fields(1))
If Not IsNull(rs.Fields(2)) Then txtcurrencysymbol = (rs.Fields(2))
If Not IsNull(rs.Fields(3)) Then txtRate = (rs.Fields(3))
If Not IsNull(rs.Fields(4)) Then cbodecimalposition = (rs.Fields(4))
If Not IsNull(rs.Fields(5)) Then cbosymbolposition.Text = (rs.Fields(5))
If Not IsNull(rs.Fields(6)) Then cbothousandseparator = Trim(rs.Fields(6))
If Not IsNull(rs.Fields(6)) Then cbodecimalseparator = Trim(rs.Fields(7))
If Not IsNull(rs.Fields(7)) Then cbonegativedisplay.Text = Trim((rs.Fields(8)))
End If


End If
End Sub

Private Sub cmdsave_Click()
On Error GoTo ErrorHandler
Dim myclass As cdbase
Set myclass = New cdbase
Set cn = CreateObject("adodb.connection")
Provider = myclass.OpenCon
cn.Open Provider, "atm", "atm"
sql = ""
sql = "SELECT     CurrCode, Description, Symbol, rateaganistsource, Decimals, Symbolposition, ThousandSeparator, NegativeDisplay, auditid, auditdatetime From Curr where currcode='" & txtcurrcode & "'"
Set rs = New ADODB.Recordset
rs.Open sql, cn, adOpenKeyset, adLockOptimistic
If rs.EOF Then

sql = "INSERT INTO Curr"
   sql = sql & " (CurrCode, Description, Symbol, rateaganistsource, Decimals, Symbolposition, ThousandSeparator,decimalseparator, NegativeDisplay, auditid, auditdatetime)"
 sql = sql & "  VALUES     ('" & txtcurrcode & "', '" & txtdescription & "', '" & txtcurrencysymbol & "', " & txtRate & ", " & cbodecimalposition & ", '" & cbosymbolposition & "', '" & Trim(cbothousandseparator.Text) & "', '" & Trim(cbodecimalseparator.Text) & "','" & cbonegativedisplay & "', '" & User & "', '" & Get_Server_Date & "')"
cn.Execute sql
 Else
 '// update the currency
sql = "Update Curr"
sql = sql & "   SET Description = '" & txtdescription & "', Symbol = '" & txtcurrencysymbol & "', rateaganistsource = " & txtRate & ", Decimals = " & cbodecimalposition & ", Symbolposition = '" & cbosymbolposition & "', ThousandSeparator = '" & Trim(cbothousandseparator.Text) & "',decimalseparator='" & cbodecimalseparator & "',"
sql = sql & " NegativeDisplay = '" & cbonegativedisplay & "', auditid = '" & User & "', auditdatetime = '" & Get_Server_Date & "'"
sql = sql & " WHERE     (CurrCode = '" & txtcurrcode & "')"
cn.Execute sql
 End If
 
 
 
 Exit Sub
ErrorHandler:
 MsgBox err.description
End Sub
