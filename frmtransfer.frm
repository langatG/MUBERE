VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmtransfer 
   Caption         =   "TRANSFERS OF KILOS"
   ClientHeight    =   4920
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10575
   Icon            =   "frmtransfer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   10575
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtkgs 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2640
      TabIndex        =   23
      Top             =   4080
      Width           =   1575
   End
   Begin VB.CommandButton cmdcommit 
      Caption         =   "Commit"
      Height          =   375
      Left            =   5520
      TabIndex        =   21
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      Height          =   375
      Left            =   8040
      TabIndex        =   2
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Receipient"
      Height          =   1575
      Left            =   0
      TabIndex        =   1
      Top             =   2160
      Width           =   10455
      Begin VB.CommandButton Command1 
         Caption         =   "F"
         Height          =   405
         Left            =   3480
         TabIndex        =   17
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtQnty1 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   3240
         TabIndex        =   16
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox txtSNo1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   2400
         TabIndex        =   15
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblNames1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         TabIndex        =   20
         Top             =   360
         Width           =   315
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Quantity Supplied (Kgs)"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Width           =   3015
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Supplier Number"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Donor"
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   10455
      Begin VB.TextBox txtSNo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   2400
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtQnty 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   3240
         TabIndex        =   4
         Top             =   960
         Width           =   2175
      End
      Begin VB.CommandButton cmdfind 
         Caption         =   "F"
         Height          =   405
         Left            =   3480
         TabIndex        =   3
         Top             =   360
         Width           =   375
      End
      Begin MSComCtl2.DTPicker DTPTransferDate 
         Height          =   375
         Left            =   8040
         TabIndex        =   9
         Top             =   960
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         MouseIcon       =   "frmtransfer.frx":0442
         CalendarBackColor=   8454016
         Format          =   16252929
         CurrentDate     =   40095
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Transfer  Date"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5640
         TabIndex        =   10
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Supplier Number"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Quantity Supplied (Kgs)"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   3015
      End
      Begin VB.Label lblNames 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         TabIndex        =   6
         Top             =   360
         Width           =   315
      End
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   3240
      TabIndex        =   11
      Top             =   0
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   16252929
      CurrentDate     =   40157
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   6000
      TabIndex        =   12
      Top             =   0
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   16252929
      CurrentDate     =   40157
   End
   Begin VB.Label Label6 
      Caption         =   "Kilos To Be Transfered"
      Height          =   375
      Left            =   360
      TabIndex        =   22
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Label Label9 
      Caption         =   "Date From"
      Height          =   375
      Left            =   2160
      TabIndex        =   14
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label10 
      Caption         =   "Date To"
      Height          =   255
      Left            =   5040
      TabIndex        =   13
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "frmtransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdcommit_Click()
On Error GoTo errorhandler
Dim Price As Double
Set rs = New ADODB.Recordset
sql = "SELECT Price from d_Price"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
Price = rs!Price
End If
txtQnty = txtQnty
txtQnty1 = txtQnty1
'//check if the kilos is a numeric
If Not IsNumeric(txtkgs) Then
MsgBox "Enter a double or a decimal number", vbCritical
Exit Sub
End If
If CDbl(txtkgs) > CDbl(txtQnty) Then
MsgBox "Total kilo to be transfered will make the donor negative. This is not allowed", vbInformation
Exit Sub
End If
If txtSNo = "" Then
MsgBox "Enter the donor supplier no", vbCritical
Exit Sub
End If
If txtSNo1 = "" Then
MsgBox "Enter the Receipient supplier no", vbCritical
Exit Sub
End If
'//donor
Set cn = New ADODB.Connection
sql = "d_sp_MilkIntake " & txtSNo & ",'" & DTPTransferDate & "'," & txtkgs * -1 & "," & Price & "," & Price * CCur(txtkgs * -1) & ",'" & Time & "','" & User & "','Transfer To " & txtSNo1 & "'"
oSaccoMaster.ExecuteThis (sql)
'//receipient
Set cn = New ADODB.Connection
sql = "d_sp_MilkIntake " & txtSNo1 & ",'" & DTPTransferDate & "'," & txtkgs & "," & Price & "," & Price * CCur(txtkgs) & ",'" & Time & "','" & User & "','Transfer from " & txtSNo & "'"
oSaccoMaster.ExecuteThis (sql)
Exit Sub
errorhandler:
MsgBox err.description
End Sub

Private Sub cmdfind_Click()
  Me.MousePointer = vbHourglass
        frmSearchSupplier.Show vbModal
        txtSNo = sel
        txtSNo_Validate True
        Me.MousePointer = 0
End Sub

Private Sub Command1_Click()
Me.MousePointer = vbHourglass
        frmSearchSupplier.Show vbModal
        txtSNo1 = sel
        txtSNo1_Validate True
        Me.MousePointer = 0
End Sub

Private Sub Form_Load()
dtpFrom = Format(Get_Server_Date, "dd/mm/yyyy")
DTPTransferDate = dtpFrom
dtpFrom = DateSerial(Year(dtpFrom), month(dtpFrom), 1)
dtpTo = DateSerial(Year(dtpFrom), month(dtpFrom) + 1, 1 - 1)
End Sub

Private Sub txtSNo_KeyPress(KeyAscii As Integer)
If (KeyAscii = 8) Or (KeyAscii = 48) Or (KeyAscii = 49) Or (KeyAscii = 50) Or (KeyAscii = 51) Or (KeyAscii = 52) Or (KeyAscii = 53) Or (KeyAscii = 54) Or (KeyAscii = 55) Or (KeyAscii = 56) Or (KeyAscii = 57) Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
MsgBox "Please enter a number "
End If
End Sub



Private Sub txtSNo_Validate(Cancel As Boolean)
On Error GoTo errorhandler
Set rs = New ADODB.Recordset
sql = "d_sp_SelectSuppliers '" & txtSNo & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
If Not IsNull(rs.Fields(2)) Then lblNames.Caption = rs.Fields(2)
'//get total kilos.
Startdate = DateSerial(Year(dtpFrom), month(dtpFrom), 1)
Enddate = DateSerial(Year(dtpTo), month(dtpTo) + 1, 1 - 1)

 Set rs = New ADODB.Recordset
    sql = "d_sp_TotalMonth " & txtSNo & ",'" & dtpFrom & "','" & dtpTo & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
    If Not IsNull(rs.Fields(0)) Then txtQnty = rs.Fields(0)
    Else
    txtQnty = "0.00"
    End If
Else
lblNames.Caption = ""
End If
If rs.RecordCount = 0 Then
lblNames.Caption = ""
End If
Exit Sub
errorhandler:
MsgBox err.description
End Sub
Private Sub txtSNo1_KeyPress(KeyAscii As Integer)
If (KeyAscii = 8) Or (KeyAscii = 48) Or (KeyAscii = 49) Or (KeyAscii = 50) Or (KeyAscii = 51) Or (KeyAscii = 52) Or (KeyAscii = 53) Or (KeyAscii = 54) Or (KeyAscii = 55) Or (KeyAscii = 56) Or (KeyAscii = 57) Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
MsgBox "Please enter a number "
End If
End Sub



Private Sub txtSNo1_Validate(Cancel As Boolean)
On Error GoTo errorhandler
Set rs = New ADODB.Recordset
sql = "d_sp_SelectSuppliers '" & txtSNo1 & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
If Not IsNull(rs.Fields(2)) Then lblNames1.Caption = rs.Fields(2)
'//get total kilos.
Startdate = DateSerial(Year(dtpFrom), month(dtpFrom), 1)
Enddate = DateSerial(Year(dtpTo), month(dtpTo) + 1, 1 - 1)

 Set rs = New ADODB.Recordset
    sql = "d_sp_TotalMonth " & txtSNo1 & ",'" & dtpFrom & "','" & dtpTo & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
    If Not IsNull(rs.Fields(0)) Then txtQnty1 = rs.Fields(0)
    Else
    txtQnty1 = "0.00"
    End If
Else
lblNames1.Caption = ""
End If
If rs.RecordCount = 0 Then
lblNames1.Caption = ""
End If
Exit Sub
errorhandler:
MsgBox err.description
End Sub
