VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmserialization 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "SERIAL NUMBERING"
   ClientHeight    =   4500
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   8295
   Icon            =   "frmserialization.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   8295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdrange 
      Caption         =   "Post Range"
      Height          =   615
      Left            =   7200
      TabIndex        =   29
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton cmdedit 
      Caption         =   "Delete"
      DisabledPicture =   "frmserialization.frx":08CA
      DownPicture     =   "frmserialization.frx":0A14
      Height          =   615
      Left            =   7200
      Picture         =   "frmserialization.frx":0B5E
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      DisabledPicture =   "frmserialization.frx":0E68
      DownPicture     =   "frmserialization.frx":0FB2
      Height          =   615
      Left            =   7200
      Picture         =   "frmserialization.frx":10FC
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "&Save"
      DisabledPicture =   "frmserialization.frx":1246
      DownPicture     =   "frmserialization.frx":1B10
      Height          =   615
      Left            =   7200
      Picture         =   "frmserialization.frx":23DA
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1080
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Height          =   3735
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   6855
      Begin VB.TextBox txtupperrange 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   2400
         TabIndex        =   23
         Top             =   3120
         Width           =   1935
      End
      Begin VB.TextBox txtlowerrange 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   480
         TabIndex        =   22
         Top             =   3120
         Width           =   1815
      End
      Begin VB.CheckBox chkrange 
         Caption         =   "Use serialNo on Range"
         Height          =   195
         Left            =   480
         TabIndex        =   21
         Top             =   2520
         Width           =   2055
      End
      Begin MSComCtl2.DTPicker DTPtransdate 
         Height          =   255
         Left            =   1440
         TabIndex        =   20
         Top             =   1920
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         Format          =   130416641
         CurrentDate     =   38042
      End
      Begin VB.PictureBox Picture2 
         Height          =   300
         Left            =   5160
         Picture         =   "frmserialization.frx":2524
         ScaleHeight     =   240
         ScaleWidth      =   315
         TabIndex        =   18
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox txtserialno 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1320
         TabIndex        =   10
         Top             =   840
         Width           =   3855
      End
      Begin VB.PictureBox Picture1 
         Height          =   300
         Left            =   2640
         Picture         =   "frmserialization.frx":27E6
         ScaleHeight     =   240
         ScaleWidth      =   315
         TabIndex        =   7
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtproductcode 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1320
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         Caption         =   "Use Range"
         Height          =   1335
         Left            =   240
         TabIndex        =   24
         Top             =   2280
         Width           =   6375
         Begin VB.TextBox txtincrement 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5040
            TabIndex        =   31
            Top             =   240
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CheckBox chkuseexisting 
            Caption         =   "Use Existing"
            Height          =   255
            Left            =   2400
            TabIndex        =   30
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox txtrangediff 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   4200
            TabIndex        =   27
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label Label10 
            Caption         =   "Interval Value"
            Height          =   255
            Left            =   3960
            TabIndex        =   32
            Top             =   240
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label lblrangediff 
            Caption         =   "Range difference"
            Height          =   255
            Left            =   4200
            TabIndex        =   28
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label lblupperrange 
            Caption         =   "Upper range"
            Height          =   255
            Left            =   2160
            TabIndex        =   26
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label lbllowerrange 
            Caption         =   "Lower range"
            Height          =   255
            Left            =   240
            TabIndex        =   25
            Top             =   600
            Width           =   975
         End
      End
      Begin VB.Label Label9 
         Caption         =   "Transaction Date"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label lblunserialised 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   4320
         TabIndex        =   16
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Unserialized"
         Height          =   300
         Left            =   3120
         TabIndex        =   15
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Lblserialised 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1320
         TabIndex        =   14
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Serialized"
         Height          =   300
         Left            =   240
         TabIndex        =   13
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Serial No"
         Height          =   300
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblproductname 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   3240
         TabIndex        =   8
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label5 
         Caption         =   "Product Code"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "STOCKS"
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   600
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   1260
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   255
      Left            =   1455
      TabIndex        =   1
      Top             =   240
      Width           =   30
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "Label1"
      Height          =   3855
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   6735
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      TabIndex        =   3
      Top             =   -120
      Width           =   8415
   End
End
Attribute VB_Name = "frmserialization"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public SelectedDsn As String
Dim cn As Connection
Dim edit As Boolean
Dim strName1  As String
Dim Provider As String
Private Sub chkrange_Click()
If chkrange = Unchecked Then
cmdsave.Enabled = True
txtrangediff.Visible = False
txtupperrange.Visible = False
txtlowerrange.Visible = False
lblrangediff.Visible = False
lblupperrange.Visible = False
lbllowerrange.Visible = False
ElseIf chkrange.value = vbChecked Then
cmdsave.Enabled = False
txtrangediff.Visible = True
txtupperrange.Visible = True
txtlowerrange.Visible = True
lblrangediff.Visible = True
lblupperrange.Visible = True
lblupperrange.Visible = True
lbllowerrange.Visible = True
End If
End Sub
Private Sub cmdclose_Click()
If lblunserialised <> 0 Then
MsgBox "You have not serilized all the items, please do so before you exit", vbInformation

Else
Unload Me
End If
End Sub
Private Sub cmdedit_Click()
Dim U, S
Dim serialid
Set cn = CreateObject("adodb.connection")
   Provider = cn
Set cn = New ADODB.Connection
cn.Open Provider, "atm", "atm"
   Dim rst As Recordset
   sql = ""
   sql = "delete from serialno where serialno='" & txtSERIALNO & "'"
   cn.Execute sql
      S = S - 1
      U = U + 1
Dim t As Long
      t = S + U
      If (S = 0 Or S < 0) And U = 0 Then
      MsgBox "No item in the data base", vbInformation, "Serialization"
      Exit Sub
      Else
      If S < 0 Then
      MsgBox "items cannot be zero from the database", vbInformation, "serialization"
      Exit Sub
      Else
sql = ""
sql = "update ag_products set qout=" & t & " , serialized=" & S & " where p_code='" & txtproductcode & "'"
cn.Execute sql
'End If
End If
 If U < 0 Then
     MsgBox "items cannot be zero from the database", vbInformation, "serialization"
      Exit Sub
      
      Else

sql = ""
sql = "update ag_products set qout=" & t & " ,unserialized=" & U & "  where p_code='" & txtproductcode & "'"
cn.Execute sql
'End If
End If

        sql = ""
        sql = "select * from ag_products where p_code='" & txtproductcode & "'"
        Set rs = New ADODB.Recordset
         rs.Open sql, cn, adOpenKeyset, adLockOptimistic
        If Not rs.EOF Then
        If Not IsNull(rs.Fields("serialized")) Then Lblserialised = S
        If Not IsNull(rs.Fields("unserialized")) Then lblunserialised = U
        End If
        MsgBox "Serial Number deleted", vbInformation, "deleting serial no"
     
     If S = 0 Then
      MsgBox "All serialized entries are deleted!", vbInformation, "Serialization"
        
        Exit Sub
      End If
      End If
   On Error Resume Next
   txtproductcode = txtproductcode.Text
   If chkrange.value = vbChecked Then
   
   txtrangediff = ""
   txtlowerrange = ""
   txtupperrange = ""
   txtSERIALNO = ""
   txtproductcode.SetFocus
   End If
End Sub

Private Sub cmdrange_Click()
Dim S
 On Error Resume Next
Dim Z As Double, X As Double, Y As Double, U As Double
Dim I As Double
Dim j As Integer
Dim rst As Recordset

If txtrangediff <= lblunserialised Then
    Lblserialised = Lblserialised
ElseIf Val(txtrangediff) > Val(lblunserialised) Then
MsgBox "Ensure that the total number you are trying to batch serialized should not be greater than the Unserialised value", vbInformation, "Serialization"
'Else: Val (lblunserialised = 0)
Exit Sub
End If
Provider = cn
Set cn = New ADODB.Connection
cn.Open Provider, "atm", "atm"
 If txtrangediff = "" And txtlowerrange = "" And txtupperrange = "" Then
 MsgBox "Enter Ranges", vbInformation, "Using ranges"
 Exit Sub
End If
j = 0
 X = txtlowerrange.Text
Y = txtupperrange.Text
Z = txtrangediff.Text
U = lblunserialised
Dim h As Long
j = 1
I = X
If U = 0 Or U < 0 Then
MsgBox "No item to serialized ", vbInformation, "Serialization"
Exit Sub
End If

Do While j <= Z
If Lblserialised = 0 Then I = X

    If Lblserialised = 0 Then S = Lblserialised + 1 Else S = Lblserialised
      If Lblserialised = txtrangediff.Text Then S = Lblserialised + 1 Else S = Lblserialised + 1
      If Lblserialised = 0 Then U = lblunserialised - 1 Else U = (lblunserialised) - 1
      sql = ""
      sql = "select * from serialno where p_code='" & txtproductcode & "' and serialno='" & I & "'"
      Set rst = New ADODB.Recordset
      rst.Open sql, cn
      If rst.EOF Then
            sql = ""
            sql = "INSERT INTO serialno(serialno,p_code,used)"
            sql = sql & " values('" & I & "','" & txtproductcode & "',0)"
            cn.Execute sql
            sql = ""
            sql = "update ag_products set unserialized=" & U & " , serialized=" & S & ",qout=" & U + S & " where p_code='" & txtproductcode & "'"
            cn.Execute sql
            I = I + 1
            sql = ""
            sql = "select * from ag_products where p_code='" & txtproductcode & "'"
            Set rs = New ADODB.Recordset
            rs.Open sql, cn, adOpenKeyset, adLockOptimistic
            If Not rs.EOF Then
            If Not IsNull(rs.Fields("serialized")) Then Lblserialised = S
            If Not IsNull(rs.Fields("unserialized")) Then lblunserialised = U
            End If
            j = j + 1
            Else
            MsgBox "The serial number  already exist,The System will try moving to the next record", vbInformation, "Serializations"
            Exit Sub
            If j = 1 Then
            Exit Sub
            End If
            End If
Loop
MsgBox "You have successefully seriallized " & txtrangediff & "  ", vbInformation, "Serialization"

Exit Sub
ErrorHandler:
MsgBox err.description
End Sub

Private Sub cmdsave_Click()
Dim U, S
On Error GoTo ErrorHandler
Dim serialid
Set cn = CreateObject("adodb.connection")
    If txtproductcode = "" Then
         MsgBox "Select product Code From The Existing Items", vbInformation, "Serializations"
         ElseIf chkrange.value = vbChecked And txtlowerrange = "" And txtupperrange = "" Then
         MsgBox "Enter Lower range", vbInformation, "Use range"
    Exit Sub
    End If
   Provider = cn
Set cn = New ADODB.Connection
cn.Open Provider, "atm", "atm"
   Dim rst As Recordset
   sql = ""
   sql = "select * from serialno where serialno='" & txtSERIALNO & "' AND P_CODE='" & txtproductcode & "'"
   Set rst = New ADODB.Recordset
   rst.Open sql, cn, adOpenKeyset, adLockOptimistic

   If Not rst.EOF And edit = False Then
   MsgBox "The Serial number you have entered has been used, try a number which has not been used", vbInformation, "Serialization"
   Exit Sub
   End If
   If U = "" Then U = 0
   If S = "" Then S = 0
   U = lblunserialised
   S = Lblserialised
   If S = "" Then S = 0
   If U <= 0 Then
   MsgBox "All the items have been serialized", vbInformation, "Serialization"
   Exit Sub
   End If
   U = U - 1
   S = S + 1
  
cn.CommandTimeout = 32766
cn.Close
Provider = cn
Set cn = New ADODB.Connection
cn.Open Provider, "atm", "atm"

sql = ""
sql = "INSERT INTO serialno(serialno,p_code,used)"
sql = sql & " values('" & txtSERIALNO & "','" & txtproductcode & "',0)"
cn.Execute sql

'// update the ag_products

sql = ""
sql = "update ag_products set unserialized=" & U & " , serialized=" & S & " where p_code='" & txtproductcode & "'"
cn.Execute sql

MsgBox "Records updated Successfully", vbInformation, "Serialization"

    sql = ""
   sql = "select * from ag_products where p_code='" & txtproductcode & "'"
   Set rs = New ADODB.Recordset
   rs.Open sql, cn, adOpenKeyset, adLockOptimistic
   If Not rs.EOF Then
       If Not IsNull(rs.Fields("serialized")) Then Lblserialised = S
       If Not IsNull(rs.Fields("unserialized")) Then lblunserialised = U
   End If
   On Error Resume Next
   txtSERIALNO = ""
   txtSERIALNO.SetFocus
Exit Sub
ErrorHandler:
MsgBox err.description
End Sub

Private Sub Form_Load()
edit = False
txtrangediff.Visible = False
txtupperrange.Visible = False
txtlowerrange.Visible = False
lblrangediff.Visible = False
lblupperrange.Visible = False
cmdsave.Enabled = True
lbllowerrange.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
If lblunserialised <> 0 Then
MsgBox "You have not exhausted the serialization please ensure you clear them before exiting this form", vbInformation
Exit Sub
End If
End Sub

Private Sub Picture1_Click()
Dim U, S
    Dim X
    Dim sal
    Dim procode
    Dim rs As Recordset
Lblserialised = ""
lblunserialised = ""

    frmSearch.Show vbModal
    
    X = sel
    If X <> "" Then
        txtproductcode = X
    End If
    
    
    Set cn = CreateObject("adodb.connection")
    If txtproductcode = "" Then Exit Sub
   Provider = cn
Set cn = New ADODB.Connection
cn.Open Provider, "atm", "atm"

   
   sql = ""
   sql = "select * from ag_products where p_code='" & X & "'"
   Set rs = New ADODB.Recordset
   rs.Open sql, cn, adOpenKeyset, adLockOptimistic
   If Not rs.EOF Then
        lblproductname = rs.Fields("p_name")
       If Not IsNull(rs.Fields("serialized")) Then Lblserialised = rs.Fields("serialized")
       If Not IsNull(rs.Fields("unserialized")) Then lblunserialised = rs.Fields("unserialized")
       If Not IsNull(rs.Fields("date_entered")) Then DTPTransdate = rs.Fields("date_entered")
       S = Lblserialised
       U = lblunserialised
       On Error Resume Next
       txtSERIALNO.SetFocus
   End If
  'End If
End Sub

Private Sub Picture2_Click()
  On Error Resume Next
    strName1 = txtproductcode
        Frmsearchserials.Show vbModal
    Me.Refresh
    If strName1 <> "" Then txtSERIALNO.Text = strName
End Sub

Private Sub txtupperrange_Change()
Dim X, Y, Z
If txtlowerrange = "" Then txtlowerrange = 0
If txtupperrange = "" Then txtupperrange = 0

X = txtlowerrange.Text
Y = txtupperrange.Text

Z = (Y - X) + 1

txtrangediff = Z


End Sub

Private Sub txtupperrange_Click()
txtupperrange_Change
End Sub
