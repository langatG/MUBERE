VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Begin VB.Form frmusergroup 
   Caption         =   "USER GROUP"
   ClientHeight    =   6315
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7800
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmusergroup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   7800
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView Lvwusergroup 
      Height          =   2505
      Left            =   45
      TabIndex        =   17
      Top             =   3300
      Width           =   7650
      _ExtentX        =   13494
      _ExtentY        =   4419
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      MaxLength       =   30
      TabIndex        =   14
      Top             =   660
      Width           =   5655
   End
   Begin VB.Frame Frame1 
      Caption         =   "User Group's Rights"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   60
      TabIndex        =   7
      Top             =   1200
      Width           =   7455
      Begin VB.CheckBox chkFixedAssets 
         Caption         =   "Fixed Assets"
         Height          =   255
         Left            =   4920
         TabIndex        =   20
         Top             =   1200
         Width           =   1545
      End
      Begin VB.CheckBox chkAcc 
         Caption         =   "Accounts"
         Height          =   255
         Left            =   4920
         TabIndex        =   19
         Top             =   720
         Width           =   1545
      End
      Begin VB.CheckBox chkAccountsPayable 
         Caption         =   "Accounts Payable"
         Height          =   195
         Left            =   4920
         TabIndex        =   18
         Top             =   300
         Width           =   1935
      End
      Begin VB.CheckBox chkfile 
         Caption         =   "File module"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   300
         Width           =   1815
      End
      Begin VB.CheckBox chktransactions 
         Caption         =   "Transactions"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   742
         Width           =   1935
      End
      Begin VB.CheckBox chkReports 
         Caption         =   "Reports module"
         Height          =   195
         Left            =   2520
         TabIndex        =   11
         Top             =   300
         Width           =   1575
      End
      Begin VB.CheckBox chkSetup 
         Caption         =   "Set up"
         Height          =   195
         Left            =   2520
         TabIndex        =   10
         Top             =   742
         Width           =   2055
      End
      Begin VB.CheckBox Chkdatabase 
         Caption         =   "Activity"
         Height          =   225
         Left            =   120
         TabIndex        =   9
         Top             =   1125
         Width           =   2295
      End
      Begin VB.CheckBox chkCashBook 
         Caption         =   "Cash Book"
         Height          =   195
         Left            =   2520
         TabIndex        =   8
         Top             =   1125
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   375
      Left            =   1320
      TabIndex        =   6
      Top             =   5910
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   6120
      TabIndex        =   5
      Top             =   5910
      Width           =   1095
   End
   Begin VB.TextBox txtID 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      MaxLength       =   5
      TabIndex        =   4
      Top             =   180
      Width           =   2295
   End
   Begin VB.CommandButton cmdedit 
      Cancel          =   -1  'True
      Caption         =   "&Edit"
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   5910
      Width           =   1095
   End
   Begin VB.CommandButton cmddelete 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   5910
      Width           =   1095
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "&Add"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   5910
      Width           =   1095
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4920
      TabIndex        =   0
      Top             =   5910
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "User Group Name"
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
      Left            =   90
      TabIndex        =   16
      Top             =   660
      Width           =   1470
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Group ID"
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
      Left            =   90
      TabIndex        =   15
      Top             =   195
      Width           =   750
   End
End
Attribute VB_Name = "frmusergroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ed As Boolean
Dim myclass As Object
Dim I As Object
Dim ans
Public Event action(bsaved As Boolean)
Public Sub edit(selected As String)
Dim rsp As Object
Dim cn As Connection
Set myclass = New cdbase
Set cn = CreateObject("adodb.connection")
Provider = myclass.OpenCon

cn.Open Provider, "atm", "atm"
Set rsp = CreateObject("adodb.recordset")
sql = "select * from usergroups where groupid='" & selected & "'"
rsp.Open sql, cn
txtId = selected
If Not rsp.EOF Then
    If Not IsNull(rsp!GroupName) Then txtname = rsp!GroupName
    
    If rsp!CashBook = True Then
    chkCashBook = vbChecked
    Else
    chkCashBook = vbUnchecked
    End If
    
    If rsp!FixedAssets = True Then
    chkFixedAssets = vbChecked
    Else
    chkFixedAssets = vbUnchecked
    End If
    
    If rsp!transactions = True Then
    chktransactions = vbChecked
    Else
    chktransactions = vbUnchecked
    End If
    If rsp!Accounts = True Then
    chkAcc = vbChecked
    Else
    chkAcc = vbUnchecked
    End If
    If rsp!Reports = True Then
    chkReports = vbChecked
    Else
    chkReports = vbUnchecked
    End If
    If rsp!activity = True Then
    Chkdatabase = vbChecked
    Else
    Chkdatabase = vbUnchecked
    End If
    If rsp!Files = True Then
    chkfile = vbChecked
    Else
    chkfile = vbUnchecked
    End If
    If rsp!setup = True Then
    chkSetup = vbChecked
    Else
    chkSetup = vbUnchecked
    End If
End If
End Sub


Private Sub cmdAdd_Click()
 For Each I In Controls
    
        If TypeOf I Is CheckBox Then I.value = 0
        
        If TypeOf I Is TextBox Then I.Text = ""
        
    Next
    txtId.SetFocus
End Sub

Private Sub cmdcancel_Click()
ed = False

 For Each I In Controls
    
        If TypeOf I Is CheckBox Then I.value = 0
        
        If TypeOf I Is TextBox Then I.Text = ""
        
    Next
    
End Sub

Private Sub cmdclose_Click()

    Unload Me
    
End Sub

Private Sub cmddelete_Click()
Dim rsd As Object
Set myclass = New cdbase
Set cn = CreateObject("adodb.connection")
Provider = myclass.OpenCon
cn.Open Provider, "atm", "atm"
ans = MsgBox("Are You sure  You Want To Delete", vbYesNo, "Deleting Station")
If ans = vbYes Then

Set rsd = CreateObject("adodb.recordset")
sql = "delete  from usergroups where groupid='" & txtId & "'"
myclass.Delete sql
Else
Exit Sub
End If
Form_Load

Set rsd = Nothing
    For Each I In Controls
        If TypeOf I Is TextBox Then I.Text = ""
    Next
    Set I = Nothing
End Sub

Private Sub cmdedit_Click()
ed = True
'edit Lvwstation.SelectedItem
edit Lvwusergroup.SelectedItem
End Sub

Private Sub cmdsave_Click()
    Dim rs As Recordset
    Dim cn As Connection
    Set cn = New Connection
    Set rs = CreateObject("adodb.recordset")
    Set cn = CreateObject("adodb.connection")
    Provider = myclass.OpenCon
   cn.Open Provider, "atm", "atm"
    Set rs = oSaccoMaster.GetRecordset("select * from usergroups where groupid='" & txtId & "'")
    '//check if that id exist
    If ed = True Then
        '// update the changes incase you have change the data
        sql = "update usergroups "
        sql = sql & "  set groupid  ='" & txtId
        sql = sql & "', groupName = '" & txtname
        sql = sql & "', files='" & chkfile
        sql = sql & "', transactions ='" & chktransactions
        sql = sql & "',activity='" & Chkdatabase
        sql = sql & "',Reports ='" & chkReports
        sql = sql & "',setup ='" & chkSetup
        sql = sql & "',Accounts ='" & chkAcc
        sql = sql & "',AccountsPay='" & chkAccountsPayable
        sql = sql & "',CashBook='" & chkCashBook
        sql = sql & "',FixedAssets='" & chkFixedAssets
        sql = sql & "'  where groupid ='" & txtId & "'"
        cn.Execute sql
    Else
        '// insert all the new data required
        If Not rs.EOF Then
            MsgBox "GroupID '" & txtId & "' already Exists", vbInformation, "Adding User Group"
            Exit Sub
        End If
        Dim ri As Recordset
        Set ri = New Recordset
        Set cn = New Connection
       cn.Open Provider, "atm", "atm"
        sql = ""
        sql = "INSERT INTO usergroups "
        sql = sql & " (groupid,GroupName,files,Transactions,CashBook,reports,setup,Activity,Accounts,AccountsPay,FixedAssets)"
        sql = sql & " VALUES     ('" & txtId & "','" & txtname & "'," & chkfile & "," & chktransactions & "," & chkCashBook & "," & chkReports & "," & chkSetup & "," & Chkdatabase & "," & chkAcc & "," & chkAccountsPayable & "," & chkFixedAssets & ")"
        cn.Execute sql
    End If
    Form_Load
End Sub

Sub loadUserGroups()
    Dim cn As Connection
    With Lvwusergroup
        .ListItems.Clear
        .ColumnHeaders.Clear
    End With
    Set cn = New Connection
    Set myclass = New cdbase
    Provider = myclass.OpenCon
    Set cn = CreateObject("adodb.connection")
    Set rs = oSaccoMaster.GetRecordset("select * from usergroups")
    With Lvwusergroup
        .ColumnHeaders.Add , , "Group_ID"
        .ColumnHeaders.Add , , "Group Name"
        .ColumnHeaders.Add , , "Files"
        .ColumnHeaders.Add , , "Transactions"
        .ColumnHeaders.Add , , "Cash Book"
        .ColumnHeaders.Add , , "Activity"
        .ColumnHeaders.Add , , "Accounts"
        .ColumnHeaders.Add , , "Report"
        .ColumnHeaders.Add , , "Set up"
        .ColumnHeaders.Add , , "Accounts Payable"
        .ColumnHeaders.Add , , "Fixed Assets"
    End With
    While Not rs.EOF
        With Lvwusergroup
            Set li = .ListItems.Add(, , rs.Fields("GroupID").value)
                li.ListSubItems.Add , , rs.Fields("GroupName")
            If rs.Fields("files") = True Then
                li.ListSubItems.Add , , "Yes"
            Else
                li.ListSubItems.Add , , "No"
            End If
            If rs.Fields("Transactions") = True Then
                li.ListSubItems.Add , , "Yes"
            Else
                li.ListSubItems.Add , , "No"
            End If
            If rs.Fields("CashBook") = True Then
                li.ListSubItems.Add , , "Yes"
            Else
                li.ListSubItems.Add , , "No"
            End If
            If rs.Fields("Activity") = True Then
                li.ListSubItems.Add , , "Yes"
            Else
                li.ListSubItems.Add , , "No"
            End If
            If rs.Fields("Accounts") = True Then
                li.ListSubItems.Add , , "Yes"
            Else
                li.ListSubItems.Add , , "No"
            End If
            If rs.Fields("Reports") = True Then
                li.ListSubItems.Add , , "Yes"
            Else
                li.ListSubItems.Add , , "No"
            End If
            If rs.Fields("setup") = True Then
                li.ListSubItems.Add , , "Yes"
            Else
                li.ListSubItems.Add , , "No"
            End If
            If rs.Fields("AccountsPay") = True Then
                li.ListSubItems.Add , , "Yes"
            Else
                li.ListSubItems.Add , , "No"
            End If
            If rs.Fields("FixedAssets") = True Then
                li.ListSubItems.Add , , "Yes"
            Else
                li.ListSubItems.Add , , "No"
            End If
        End With
        rs.MoveNext
    Wend
    Set rs = Nothing
End Sub

Private Sub Form_Load()
    ed = False
    For Each I In Controls
        If TypeOf I Is CheckBox Then I.value = 0
        If TypeOf I Is TextBox Then I.Text = ""
    Next
    loadUserGroups
    On Error Resume Next
    txtId.SetFocus
End Sub

Private Sub Lvwusergroup_Click()
    On Error GoTo SysError
    If Lvwusergroup.ListItems.Count > 0 Then
        txtId = Lvwusergroup.SelectedItem
        txtname = Lvwusergroup.SelectedItem.SubItems(1)
        chkfile = IIf(Lvwusergroup.SelectedItem.SubItems(2) = "Yes", vbChecked, vbUnchecked)
        chktransactions = IIf(Lvwusergroup.SelectedItem.SubItems(3) = "Yes", vbChecked, vbUnchecked)
        Chkdatabase = IIf(Lvwusergroup.SelectedItem.SubItems(5) = "Yes", vbChecked, vbUnchecked)
        chkCashBook = IIf(Lvwusergroup.SelectedItem.SubItems(4) = "Yes", vbChecked, vbUnchecked)
        chkSetup = IIf(Lvwusergroup.SelectedItem.SubItems(8) = "Yes", vbChecked, vbUnchecked)
        chkAcc = IIf(Lvwusergroup.SelectedItem.SubItems(6) = "Yes", vbChecked, vbUnchecked)
        chkReports = IIf(Lvwusergroup.SelectedItem.SubItems(7) = "Yes", vbChecked, vbUnchecked)
        chkAccountsPayable = IIf(Lvwusergroup.SelectedItem.SubItems(9) = "Yes", vbChecked, vbUnchecked)
        chkFixedAssets = IIf(Lvwusergroup.SelectedItem.SubItems(10) = "Yes", vbChecked, vbUnchecked)
    End If
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub Lvwusergroup_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Lvwusergroup_Click
End Sub
