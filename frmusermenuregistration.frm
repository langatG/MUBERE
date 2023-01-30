VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Begin VB.Form frmusermenuregistration 
   Caption         =   "USER Menu Assignment"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10845
   LinkTopic       =   "Form1"
   ScaleHeight     =   9180
   ScaleWidth      =   10845
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkselectall 
      Caption         =   "Select All"
      Height          =   255
      Left            =   6720
      TabIndex        =   11
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Remove Previlleges"
      Height          =   375
      Left            =   8760
      TabIndex        =   10
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton cmdprofile 
      Caption         =   "View Profile"
      Height          =   255
      Left            =   4560
      TabIndex        =   8
      Top             =   360
      Width           =   2055
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      Height          =   375
      Left            =   8760
      TabIndex        =   7
      Top             =   2280
      Width           =   1695
   End
   Begin VB.PictureBox Picture5 
      Height          =   255
      Left            =   4200
      Picture         =   "frmusermenuregistration.frx":0000
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   6
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton cmdupdate 
      Caption         =   "Update Previlleges"
      Height          =   375
      Left            =   8760
      TabIndex        =   5
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      DataField       =   "UserName"
      Height          =   285
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   2
      Top             =   840
      Width           =   3615
   End
   Begin VB.TextBox txtID 
      Appearance      =   0  'Flat
      DataField       =   "userloginid"
      Height          =   285
      Left            =   2040
      TabIndex        =   1
      Top             =   360
      Width           =   2175
   End
   Begin MSComctlLib.ListView lvwprevilages 
      Height          =   7095
      Left            =   480
      TabIndex        =   0
      Top             =   1560
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   12515
      View            =   3
      Arrange         =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      OLEDropMode     =   1
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDropMode     =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Menu"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Alias"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Enable"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lvwprofile 
      Height          =   7095
      Left            =   600
      TabIndex        =   9
      Top             =   1560
      Visible         =   0   'False
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   12515
      View            =   3
      Arrange         =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      OLEDropMode     =   1
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDropMode     =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Menu"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Alias"
         Object.Width           =   4410
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "User's Names"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "User Login ID"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "frmusermenuregistration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myclass As cdbase

Private Sub chkselectall_Click()
   On Error GoTo Syserr

    With lvwprevilages
        
        If chkselectall.value = vbChecked Then
                If .ListItems.Count >= 1 Then
                    For I = 1 To .ListItems.Count
                        .ListItems(I).Checked = True
                    Next I
                End If
                'Call RefreshBatch
        Else
            If .ListItems.Count >= 1 Then
                    For I = 1 To .ListItems.Count
                        If .ListItems(I).Checked = True Then
                            .ListItems(I).Checked = False
                        End If
                    Next I
            End If
            'Call RefreshBatch
        End If
    End With
    
    Exit Sub
Syserr:
        MsgBox err.description
End Sub

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdprofile_Click()
lvwprofile.Visible = True
lvwprevilages.Visible = False
Set rs = oSaccoMaster.GetRecordset("SELECT    t.menu, tb.menu  FROM         tbl_usermenus TB INNER JOIN  tbl_menus T ON TB.menu = T.ALIAS WHERE TB.UserLoginIDs='" & txtId & "' order by t.menu")
Dim I As Integer, X As String, Y As String, r As New ADODB.Recordset, Z As String
lvwprofile.ListItems.Clear
With rs
            
                    While Not .EOF
                    Set li = lvwprofile.ListItems.Add(, , IIf(IsNull(rs.Fields(0)), "", rs.Fields(0)))
                    li.SubItems(1) = IIf(IsNull(rs.Fields(1)), "", rs.Fields(1))
                   
                    
                    
                    rs.MoveNext
                    Wend
                    
End With
End Sub

Private Sub cmdupdate_Click()
On Error GoTo ErrorHandler
Set myclass = New cdbase
    Set cn = CreateObject("adodb.connection")
    Provider = myclass.OpenCon
   cn.Open Provider, "atm", "atm"
    Dim bal As Currency
    Dim I As Integer
     For I = 1 To lvwprevilages.ListItems.Count
      If lvwprevilages.ListItems.Item(I).Checked = True Then
      sql = ""
      sql = "select UserLoginID from   tbl_usermenus where UserLoginID='" & txtId & "' and menu='" & lvwprevilages.ListItems(I).ListSubItems(1).Text & "'"
      Set rs = New ADODB.Recordset
      rs.Open sql, cn, adOpenDynamic, adLockOptimistic
      If rs.EOF Then
    sql = ""
    sql = "insert_tbl_usermenus '" & txtId & "','" & lvwprevilages.ListItems(I).ListSubItems(1).Text & "','" & Get_Server_Date & "',1"
    'sql = "set dateformat dmy update customerbalance set unpresented=1,unpdesc='Unpresented' where customerbalanceid ='" & lvememtrans.ListItems(i).ListSubItems(7).Text & "'"
    cn.Execute sql
      End If
      End If
     Next I
     MsgBox "Records successfully Updated"
     Exit Sub
ErrorHandler:

     MsgBox err.description

End Sub

Private Sub Command1_Click()
On Error GoTo ErrorHandler
Set myclass = New cdbase
    Set cn = CreateObject("adodb.connection")
    Provider = myclass.OpenCon
   cn.Open Provider, "atm", "atm"
    Dim bal As Currency
    Dim I As Integer
     For I = 1 To lvwprofile.ListItems.Count
      'If lvwprofile.ListItems.Item(I).Checked = True Then
      sql = ""
      sql = "Delete from tbl_usermenus where UserLoginIDs='" & txtId & "'"
      cn.Execute sql
      'End If
     ' End If
     Next I
     MsgBox "Records successfully Updated"
     Exit Sub
ErrorHandler:

     MsgBox err.description

End Sub

Private Sub Form_Load()
menus
End Sub
Public Sub menus()
'If lvwprevilages.ListItems.Count > 0 Then
       
        lvwprevilages.ListItems.Clear
        Dim rsTrans As New Recordset, DRTotal As Double, CRTotal As Double
        Set rsTrans = oSaccoMaster.GetRecordset("Set Dateformat dmy Select * From   tbl_menus order by menu asc")
        DRTotal = 0
        CRTotal = 0
        
            With rsTrans
            
                    While Not .EOF
                    Set li = lvwprevilages.ListItems.Add(, , IIf(IsNull(!Menu), "", !Menu))
                    li.SubItems(1) = IIf(IsNull(!Alias), "", !Alias)
                     li.SubItems(2) = IIf(IsNull(!Enabled), "", !Enabled)
                    
                    
                    rsTrans.MoveNext
                    Wend
                    
                    End With
                    
 '       End If
End Sub

Private Sub Picture5_Click()
 Me.MousePointer = vbHourglass
        frmsearchsearchusers.Show vbModal
        txtId = sel
        txtID_Validate True
        Me.MousePointer = 0
End Sub



Private Sub txtID_Validate(Cancel As Boolean)
Set rs = oSaccoMaster.GetRecordset("SELECT   USERNAME,UserLoginIDs FROM UserAccounts WHERE UserLoginIDs='" & txtId & "'")
Dim I As Integer, X As String, Y As String, r As New ADODB.Recordset, Z As String
If Not rs.EOF Then
txtId = rs.Fields(1)
txtname = rs.Fields(0)

End If
lvwprevilages.Visible = True


End Sub
