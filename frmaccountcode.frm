VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Begin VB.Form frmaccountcode 
   Caption         =   "ACCOUNT TYPE SET UP"
   ClientHeight    =   5850
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7755
   Icon            =   "frmaccountcode.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5850
   ScaleWidth      =   7755
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   120
      TabIndex        =   8
      Top             =   3120
      Width           =   7455
      Begin VB.TextBox txtcontrolaccount 
         Appearance      =   0  'Flat
         DataField       =   "MinimumBal"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """kshs""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         Height          =   285
         Left            =   2640
         MaxLength       =   20
         TabIndex        =   15
         Top             =   1680
         Width           =   2655
      End
      Begin VB.TextBox txtAccountCode 
         Appearance      =   0  'Flat
         DataField       =   "AccountCode"
         Height          =   285
         Left            =   2640
         TabIndex        =   11
         Top             =   360
         Width           =   2655
      End
      Begin VB.TextBox txtAccountName 
         Appearance      =   0  'Flat
         DataField       =   "AccountName"
         Height          =   285
         Left            =   2640
         TabIndex        =   10
         Top             =   840
         Width           =   4215
      End
      Begin VB.TextBox txtMinimumBalance 
         Appearance      =   0  'Flat
         DataField       =   "MinimumBal"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """kshs""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         Height          =   285
         Left            =   2640
         MaxLength       =   20
         TabIndex        =   9
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Label Label5 
         Caption         =   "Control Account"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Account Code"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Account Name"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Minimum Balance"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   1575
      End
   End
   Begin MSComctlLib.ListView LVWaccountcode 
      Height          =   2895
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   5106
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Close"
      Height          =   375
      Left            =   6480
      Picture         =   "frmaccountcode.frx":030A
      TabIndex        =   5
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H0000FF00&
      Caption         =   "&Save"
      Height          =   375
      Left            =   1320
      MaskColor       =   &H0080FF80&
      Picture         =   "frmaccountcode.frx":074C
      TabIndex        =   4
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton cmdedit 
      Cancel          =   -1  'True
      Caption         =   "&Edit"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton cmdadd 
      BackColor       =   &H0000FF00&
      Caption         =   "&Add"
      Height          =   375
      Left            =   0
      MaskColor       =   &H0080FF80&
      Picture         =   "frmaccountcode.frx":0DB6
      TabIndex        =   2
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H0000FF00&
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5160
      MaskColor       =   &H0080FF80&
      Picture         =   "frmaccountcode.frx":1420
      TabIndex        =   1
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton cmddelete 
      BackColor       =   &H0000FF00&
      Caption         =   "&Delete"
      Default         =   -1  'True
      Height          =   375
      Left            =   3840
      MaskColor       =   &H0080FF80&
      Picture         =   "frmaccountcode.frx":1A8A
      TabIndex        =   0
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000015&
      Height          =   2175
      Left            =   120
      TabIndex        =   7
      Top             =   3120
      Width           =   7575
   End
End
Attribute VB_Name = "frmaccountcode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ed As Boolean
Public Event action(bsaved As Boolean)
Public Sub edit(selected As String)
On Error GoTo ErrorHandler
Dim myclass As cdbase
Set myclass = New cdbase
Set cn = CreateObject("adodb.connection")
Provider = myclass.OpenCon
cn.Open Provider, "atm", "atm"
Set rs = CreateObject("adodb.recordset")
sql = "select * from accountcodes where accountcode='" & selected & "'"
rs.Open sql, cn
If Not rs.EOF Then
txtAccountCode = selected
If Not IsNull(rs!AccountName) Then txtAccountName = rs!AccountName
If Not IsNull(rs!Minimumbal) Then txtMinimumBalance = rs!Minimumbal
If Not IsNull(rs!ACCNO) Then txtcontrolaccount = rs!ACCNO
End If
Exit Sub
ErrorHandler:
MsgBox err.description
End Sub

Private Sub cmdAdd_Click()
Dim I As Object

    For Each I In Controls
    
        If TypeOf I Is TextBox Then I.Text = ""
    
    Next
ed = False
Set I = Nothing

End Sub

Private Sub cmdcancel_Click()
ed = False
Dim I As Object

    For Each I In Controls
    
        If TypeOf I Is TextBox Then I.Text = ""
    
    Next

Set I = Nothing

End Sub

Private Sub cmdclose_Click()
  Unload Me
End Sub

Private Sub cmddelete_Click()
Dim rsr As Object
Dim ans
Dim myclass As cdbase
Set myclass = New cdbase
Set cn = CreateObject("adodb.connection")
Provider = myclass.OpenCon
cn.Open Provider, "atm", "atm"
ans = MsgBox("Are You sure  You Want To Delete", vbYesNo, "Deleting Station")
If ans = vbYes Then
Set rsr = CreateObject("adodb.recordset")
sql = "delete from accountcodes where accountcode='" & txtAccountCode & "'"
rsr.Open sql, cn
myclass.Delete sql
Else
Exit Sub
End If
Form_Load

End Sub

Private Sub cmdedit_Click()
ed = True
edit LVWaccountcode.SelectedItem
End Sub

Private Sub cmdsave_Click()
   On Error GoTo ErrorHandler
    If Not IsNumeric(txtMinimumBalance) Then: MsgBox "Please enter a numeric minimum balance", vbExclamation, "FOSA": Exit Sub
    
    Dim myclass As cdbase
    
    Set myclass = New cdbase
    
    sql = ""
    If ed = True Then ' update a particular account code
    sql = "update accountcodes"
    sql = sql & "  set accountname = '" & txtAccountName
    sql = sql & "', MinimumBal= '" & txtMinimumBalance
    sql = sql & "', accno= '" & txtcontrolaccount
    sql = sql & "'  where Accountcode= '" & txtAccountCode & "'"
    myclass.save sql
    Else 'insert a new a record
    
    sql = "insert into accountcodes(accountcode,accountname,minimumbal,accno)select '" & _
    Trim(txtAccountCode) & "','" & Trim(txtAccountName) & "','" & Trim(txtMinimumBalance) & "','" & Trim(txtcontrolaccount) & "'"
    
    myclass.save sql
    End If
    Set myclass = Nothing
    
   

    Form_Load
    Exit Sub
ErrorHandler:
    MsgBox err.description
End Sub




Private Sub txtAccountCode_KeyPress(KeyAscii As Integer)
If ValidChar(KeyAscii) = False Then KeyAscii = 0
End Sub


Private Sub txtAccountName_KeyPress(KeyAscii As Integer)
If ValidChar(KeyAscii) = False Then KeyAscii = 0
End Sub
Private Sub txtMinimumBalance_KeyPress(KeyAscii As Integer)
If ValidChar(KeyAscii) = False Then KeyAscii = 0
End Sub
Sub LoadAccountCodes()

sql = ""

sql = "select * from AccountCodes"

With LVWaccountcode

    .ListItems.Clear
    .ColumnHeaders.Clear

    .ListItems.Clear
    .ColumnHeaders.Clear

End With





Set rs = CreateObject("Adodb.recordset")

    Set clsClass = New cdbase
    
    Provider = clsClass.OpenCon
    
    Set cn = CreateObject("adodb.connection")
    
   cn.Open Provider, "atm", "atm"

rs.Open sql, cn




    With LVWaccountcode
        .ColumnHeaders.Add , , "Account Code"
        .ColumnHeaders.Add , , "Account Name"
        .ColumnHeaders.Add , , "Minimum Balance"
        .ColumnHeaders.Add , , "Control Account"
    End With
    


While Not rs.EOF
    With LVWaccountcode
        
        Set li = .ListItems.Add(, , rs.Fields("AccountCode"))
        li.ListSubItems.Add , , rs.Fields("AccountName")
        li.ListSubItems.Add , , rs.Fields("MinimumBal")
       If Not IsNull(rs.Fields("accno")) Then li.ListSubItems.Add , , rs.Fields("accno")
        rs.MoveNext
        
        
    End With
Wend

LVWaccountcode.View = lvwReport


Set rs = Nothing
End Sub

Private Sub Form_Load()
LoadAccountCodes
On Error Resume Next
ed = False
Dim I As Object

    For Each I In Controls
    
        If TypeOf I Is TextBox Then I.Text = ""
    
    Next

Set I = Nothing

txtAccountCode.SetFocus

End Sub

