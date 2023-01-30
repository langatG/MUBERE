VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmsystemuser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SYSTEM USER"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11865
   Icon            =   "frmsystemuser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   11865
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtphone 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5880
      TabIndex        =   29
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Frame Frame7 
      Caption         =   "Signatures"
      Height          =   2775
      Left            =   8280
      TabIndex        =   26
      Top             =   3000
      Width           =   3495
      Begin VB.CommandButton cmdGetPic 
         Caption         =   "Signature"
         Height          =   300
         Left            =   240
         TabIndex        =   28
         Top             =   2280
         Width           =   1000
      End
      Begin VB.CommandButton cmdGetSign 
         Caption         =   "Signature"
         Height          =   300
         Left            =   3600
         TabIndex        =   27
         Top             =   3240
         Width           =   1000
      End
      Begin MSComDlg.CommonDialog DLG 
         Left            =   4680
         Top             =   600
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Image imgSign 
         BorderStyle     =   1  'Fixed Single
         Height          =   855
         Left            =   3480
         Stretch         =   -1  'True
         Top             =   1815
         Width           =   2895
      End
      Begin VB.Image imgPic 
         BorderStyle     =   1  'Fixed Single
         Height          =   1695
         Left            =   240
         Stretch         =   -1  'True
         Top             =   360
         Width           =   3015
      End
   End
   Begin VB.ComboBox cbobranch 
      Height          =   315
      ItemData        =   "frmsystemuser.frx":030A
      Left            =   2040
      List            =   "frmsystemuser.frx":0380
      TabIndex        =   24
      Top             =   5520
      Width           =   2175
   End
   Begin VB.PictureBox Picture3 
      Height          =   255
      Left            =   4800
      Picture         =   "frmsystemuser.frx":04F7
      ScaleHeight     =   195
      ScaleWidth      =   315
      TabIndex        =   21
      Top             =   5160
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSComctlLib.ListView Lvwusers 
      Height          =   2295
      Left            =   120
      TabIndex        =   20
      Top             =   240
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   4048
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.ComboBox cboUser 
      Appearance      =   0  'Flat
      DataField       =   "usergroup"
      Height          =   315
      ItemData        =   "frmsystemuser.frx":07B9
      Left            =   2040
      List            =   "frmsystemuser.frx":07C0
      TabIndex        =   12
      Top             =   4680
      Width           =   2175
   End
   Begin VB.TextBox txtConfirm 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   5880
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   11
      Top             =   3720
      Width           =   2175
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      DataField       =   "Password"
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2040
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   10
      Top             =   3720
      Width           =   2175
   End
   Begin VB.TextBox txtID 
      Appearance      =   0  'Flat
      DataField       =   "userloginid"
      Height          =   285
      Left            =   2040
      MaxLength       =   8
      TabIndex        =   9
      Top             =   2640
      Width           =   2175
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      DataField       =   "UserName"
      Height          =   285
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   8
      Top             =   3120
      Width           =   3135
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   6600
      TabIndex        =   6
      Top             =   6000
      Width           =   1095
   End
   Begin VB.TextBox txtPassExpire 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2040
      MaxLength       =   2
      TabIndex        =   5
      Top             =   4200
      Width           =   2175
   End
   Begin VB.CommandButton cmdedit 
      Cancel          =   -1  'True
      Caption         =   "&Edit"
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5400
      TabIndex        =   3
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton cmddelete 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "&New"
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CheckBox Chksuper 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Super User"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5400
      TabIndex        =   0
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label Label9 
      Caption         =   "Phone No."
      Height          =   255
      Left            =   4440
      TabIndex        =   30
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "Branch"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   5520
      Width           =   855
   End
   Begin VB.Label LBLTELLER 
      Caption         =   "GL Teller"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   5160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblglteller 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2040
      TabIndex        =   22
      Top             =   5160
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label Label5 
      Caption         =   "User Group"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   4680
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Confirm Password"
      Height          =   255
      Left            =   4440
      TabIndex        =   18
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "User Password"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "User Login ID"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "User's Names"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "Password expires after"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   4200
      Width           =   1695
   End
   Begin VB.Label Label7 
      Caption         =   "Days"
      Height          =   255
      Left            =   4440
      TabIndex        =   13
      Top             =   4200
      Width           =   375
   End
End
Attribute VB_Name = "frmsystemuser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event action(bsaved As Boolean)
Dim myclass As cdbase
Dim mvarEnabledStatus As String
Dim ans
'Dim myclass As cdbase
Private strPic As String
Private strsign As String
Dim ed As Boolean
Dim I As Object
Dim gltellerbal As Currency
Dim glnameteller As String
Dim glidnoteller As String
Dim glmemnoteller As String
Dim glpaynoteller As String
Public Sub edit(selected As String)
Dim rsp As Object
Dim myclass As cdbase
    Dim Chksuper1 As Integer
Set myclass = New cdbase
'Set  CreateObject("adodb.connection")
Provider = myclass.OpenCon
cn.Open Provider, "atm", "atm"
Set rsp = CreateObject("adodb.recordset")
sql = "select *  from useraccounts where UserLoginIDs='" & selected & "'"
rsp.Open sql, cn
txtID = selected
If Not rsp.EOF Then
If Not IsNull(rsp!username) Then txtName = rsp!username
On Error Resume Next

If Not IsNull(rsp!usergroup) Then cboUser = rsp!usergroup
If Not IsNull(rsp!passexpire) Then txtPassExpire = rsp!passexpire
If Not IsNull(rsp!DateCreated) Then Date = rsp!DateCreated
If Not IsNull(rsp!Password) Then txtPassword = rsp!Password
If Not IsNull(rsp!Phone) Then txtphone = rsp!Phone
If Not IsNull(rsp!SuperUser) Then Chksuper1 = rsp!SuperUser
If Chksuper1 = 1 Then
Chksuper = vbChecked
Else
Chksuper = vbUnchecked
End If
    End If
End Sub

Private Sub cboUser_Change()
cboUser_Click
End Sub

Private Sub cboUser_Click()
   If cboUser = "TELLER" Then
    LBLTELLER.Visible = True
    lblglteller.Visible = True
    Picture3.Visible = True
    
    Else
    LBLTELLER.Visible = False
    lblglteller.Visible = False
    Picture3.Visible = False
    End If
    Set rs = Nothing
End Sub

Private Sub cboUser_DropDown()
    
    Dim myclass As cdbase
    cboUser.Clear
    
    Set myclass = New cdbase

    Provider = myclass.OpenCon
    
    Set cn = CreateObject("adodb.connection")
    
   cn.Open Provider, "atm", "atm"
    
    Set rs = CreateObject("adodb.recordset")
    
    rs.Open "SELECT GROUPNAME FROM USERGROUPS", cn
    
    If rs.EOF Then Exit Sub
    
    With rs
        
        While Not .EOF
         
         cboUser.AddItem rs.Fields("GROUPNAME")
         
         .MoveNext
        
        Wend
    
    End With
 

    
End Sub

Private Sub cmdAdd_Click()
For Each I In Controls
If TypeOf I Is TextBox Then I.Text = ""
Next
End Sub

Private Sub cmdcancel_Click()
    ed = False
    For Each I In Controls
        
        If TypeOf I Is TextBox Then I.Text = ""
        
    Next
    cboUser.SetFocus
    txtID.SetFocus
End Sub

Private Sub cmdclose_Click()

   Unload Me

End Sub
Private Sub cmddelete_Click()
Dim myclass As Object
Dim rsd As Object
Set myclass = New cdbase
Set cn = CreateObject("adodb.connection")
Provider = myclass.OpenCon
cn.Open Provider, "atm", "atm"
ans = MsgBox("Are You sure  You Want To Delete", vbYesNo, "Deleting a user")
If ans = vbYes Then

Set rsd = CreateObject("adodb.recordset")
sql = "delete  from useraccounts where UserLoginIDs='" & txtID & "'"
myclass.Delete sql
Else
Exit Sub
End If


 Set rsd = Nothing
    For Each I In Controls
        If TypeOf I Is TextBox Then I.Text = ""
    Next
    Set I = Nothing
    cboUser.Clear
    Form_Load
End Sub

Private Sub cmdedit_Click()
ed = True
edit Lvwusers.SelectedItem
End Sub
Sub loadUsers()

With Lvwusers
    .ListItems.Clear
    .ColumnHeaders.Clear
End With

    Set rs = CreateObject("adodb.recordset")
    
        Set clsClass = New cdbase
    
    Provider = clsClass.OpenCon
    
    Set cn = CreateObject("adodb.connection")
    
   cn.Open Provider, "atm", "atm"

    sql = ""
    
    sql = "select * from UserAccounts"

    rs.Open sql, cn
    With Lvwusers
        .ColumnHeaders.Add , , "Login ID"
        .ColumnHeaders.Add , , "User Name"
        .ColumnHeaders.Add , , "User Group"
        .ColumnHeaders.Add , , "Password Expiry(Days)"
        .ColumnHeaders.Add , , "Date Created"
         .ColumnHeaders.Add , , "Phone"
          .ColumnHeaders.Add , , "Branch Code"
    End With
    


While Not rs.EOF
    With Lvwusers
                
        Set li = .ListItems.Add(, , rs.Fields("UserLoginIDs").value)
        li.ListSubItems.Add , , rs.Fields("USerName")
        If rs.Fields("USerGroup") <> "" Then li.ListSubItems.Add , , rs.Fields("USerGroup")
        If rs.Fields("PassExpire") <> "" Then li.ListSubItems.Add , , rs.Fields("PassExpire")
        If rs.Fields("DateCreated") <> "" Then li.ListSubItems.Add , , rs.Fields("DateCreated")
        If rs.Fields("Phone") <> "" Then li.ListSubItems.Add , , rs.Fields("Phone") Else li.ListSubItems.Add , , "Non"
         If rs.Fields("branchcode") <> "" Then li.ListSubItems.Add , , rs.Fields("branchcode")
        rs.MoveNext
        
        
    End With
Wend


rs.Close

Set rs = Nothing

End Sub

Private Sub cmdGetPic_Click()
With DLG
        .Filter = "Bitmap|*.bmp|Jpeg|*.jpg|GIF|*.gif"
        .ShowOpen
        imgPic = LoadPicture(.FileName)
        strPic = .FileName
    End With
End Sub

Private Sub cmdsave_Click()
Dim super As String
Dim myclass As Object
Set myclass = New cdbase
   For Each I In Controls
    If TypeOf I Is TextBox Then
        
        If I.Text = "" Then: MsgBox "Please input all the required information", vbExclamation, "FOSA": I.SetFocus: Exit Sub
    End If
    Next
    Set I = Nothing
    If Chksuper = vbChecked Then
    super = 1
    Else
    super = 0
    End If
    
    
    
    If Len(txtID) < 3 Then
    MsgBox "User ID should be at Least Three Charaters", vbInformation, "Security"
    Exit Sub
    End If
    If Len(txtPassword) < 1 Then
    MsgBox "Password should be at Least Four Charaters", vbInformation, "Security"
    Exit Sub
    End If
    If txtPassword.Text <> txtConfirm.Text Then: MsgBox "Confirmation password must be the same as the password.Please re-enter again", 0 + vbInformation, "Users": txtConfirm.Text = "": txtConfirm.SetFocus: Exit Sub
'Dim Pass As EncryptDecrypt
'Set Pass = New EncryptDecrypt

'txtPassword = Pass.Encrypt(txtPassword)
txtPassword = modsecurity.Encript_String(txtPassword)
    If cboUser = "TELLER" Then
    If lblglteller = "" Then
    MsgBox "Please assign a teller his /her own password", vbInformation, "System Users"
    Exit Sub
    End If
    Else

    End If
    
   If cbobranch = "" Then
     MsgBox "Please select branch", vbInformation
    Exit Sub
    End If
   If txtphone = "" Then
     MsgBox "Please insert Phone number", vbInformation
   Exit Sub
   End If

If ed = True Then
    '//update your table else
sql = "set dateformat dmy update useraccounts "
    sql = sql & " set username  ='" & txtName
    sql = sql & "', password='" & txtPassword
    sql = sql & "', usergroup ='" & cboUser
    sql = sql & "',passexpire='" & txtPassExpire
    sql = sql & "',datecreated ='" & Date
    sql = sql & "',superuser='" & super
    sql = sql & "',AssignGl='" & lblglteller
    sql = sql & "',branchcode='" & cbobranch
    sql = sql & "',Phone='" & txtphone
    sql = sql & "',sign='" & strPic & "'  where UserLoginIDs ='" & txtID & "'"
    
    myclass.save sql
'// insert a new record
Else
   sql = "select * from UserAccounts where UserLoginIDs='" & txtID & "'"
   Set rs = oSaccoMaster.GetRecordset(sql)
   If Not rs.EOF Then
    MsgBox "Please user login ID already used", vbInformation
    txtID.SetFocus
   Exit Sub
   End If

    sql = ""
    sql = "set dateformat dmy insert into useraccounts(username,UserLoginIDs,password,usergroup,passexpire,DateCreated,superuser,AssignGl,branchcode,sign,Phone)"
    sql = sql & "select '" & txtName & "','" & txtID & "','" & txtPassword & "','" & cboUser & "','" & txtPassExpire & "','" & Date & "','" & super & "','" & lblglteller & "','" & cbobranch & "','" & strPic & "','" & txtphone & "'"
    Set myclass = New cdbase
    myclass.save sql
    End If
    Set myclass = Nothing
    loadUsers
   'form_load

End Sub



Private Sub Form_Load()
'On Error Resume Next
loadUsers
ed = False
cboUser.Clear
cbobranch.Clear
      For Each I In Controls
        
        If TypeOf I Is TextBox Then I.Text = ""
        
    Next
  
    Dim myclass As cdbase
    
    Set myclass = New cdbase

    Provider = myclass.OpenCon
    
    Set cn = CreateObject("adodb.connection")
    
   cn.Open Provider, "atm", "atm"
    
    Set rs = CreateObject("adodb.recordset")
    
    rs.Open "SELECT BName FROM d_Branch", cn
    
    If rs.EOF Then Exit Sub
    
    With rs
        
        While Not .EOF
         
         cbobranch.AddItem rs.Fields("BName")
         
         .MoveNext
        
        Wend
    
    End With
    Set I = Nothing
    Chksuper = vbUnchecked
    On Error Resume Next
    txtID.SetFocus

End Sub

Public Property Get ComboEnabled() As Variant

    ComboEnabled = mvarEnabledStatus

End Property

Public Property Let ComboEnabled(ByVal vNewValue As Variant)
    
    mvarEnabledStatus = cboUser.Enabled = vNewValue
    
   ' PropertyChanged ComboEnabled
    
End Property

Private Sub Picture3_Click()
Dim Z, S, U
    Dim rs As Recordset

    frmsearchaccounts.Show vbModal
    Z = strName
    If Z <> "" Then
        lblglteller = Z
        
        End If
        
           Set myclass = New cdbase
    Set cn = CreateObject("adodb.connection")
    Provider = myclass.OpenCon
   cn.Open Provider, "atm", "atm"
sql = "select * from cuB where ACCno='" & lblglteller & "'"
Set rs = New ADODB.Recordset
rs.Open sql, cn
If Not rs.EOF Then
If Not IsNull(rs.Fields("availablebalance")) Then gltellerbal = rs.Fields("availablebalance")
If Not IsNull(rs.Fields("accountname")) Then glnameteller = rs.Fields("accountname")
If Not IsNull(rs.Fields("idno")) Then glidnoteller = rs.Fields("idno")
If Not IsNull(rs.Fields("memberno")) Then glmemnoteller = rs.Fields("memberno")
If Not IsNull(rs.Fields("payno")) Then glpaynoteller = rs.Fields("payno")
End If
End Sub
