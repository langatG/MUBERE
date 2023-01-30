VERSION 5.00
Begin VB.Form frmSupplier 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "  "
   ClientHeight    =   4350
   ClientLeft      =   210
   ClientTop       =   495
   ClientWidth     =   8055
   ControlBox      =   0   'False
   Icon            =   "frmSupplier.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   8055
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   240
      TabIndex        =   31
      Top             =   3480
      Width           =   6975
      Begin VB.CommandButton cmd7 
         Appearance      =   0  'Flat
         Caption         =   "Close"
         Height          =   375
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   1185
      End
      Begin VB.CommandButton cmd3 
         Appearance      =   0  'Flat
         Caption         =   "&Save"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   1185
      End
      Begin VB.CommandButton cmd2 
         Appearance      =   0  'Flat
         Caption         =   "&Modify"
         Height          =   375
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   1065
      End
      Begin VB.CommandButton cmd1 
         Appearance      =   0  'Flat
         Caption         =   "&Add"
         Height          =   375
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1065
      End
      Begin VB.CommandButton cmd6 
         Appearance      =   0  'Flat
         Caption         =   "&Delete"
         Height          =   375
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   1185
      End
      Begin VB.CommandButton cmd4 
         Appearance      =   0  'Flat
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   1185
      End
   End
   Begin VB.TextBox Text7 
      Enabled         =   0   'False
      Height          =   495
      Left            =   3480
      TabIndex        =   25
      Text            =   "Text7"
      Top             =   8400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      DataField       =   "autono"
      DataSource      =   "Data2"
      Height          =   285
      Left            =   2580
      TabIndex        =   24
      Top             =   7260
      Width           =   1095
   End
   Begin VB.TextBox ctr1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1800
      TabIndex        =   23
      Top             =   7800
      Width           =   735
   End
   Begin VB.TextBox ctr2 
      Height          =   285
      Left            =   2640
      TabIndex        =   22
      Top             =   7800
      Width           =   255
   End
   Begin VB.TextBox ctr3 
      Height          =   285
      Left            =   3000
      TabIndex        =   21
      Top             =   7740
      Width           =   255
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Ctr"
      Top             =   10320
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Select * from Suppliers order by CompanyName"
      Top             =   8160
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Suppliers"
      Top             =   8880
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   2775
      Left            =   120
      TabIndex        =   14
      Top             =   540
      Width           =   7815
      Begin VB.PictureBox Picture5 
         Height          =   255
         Left            =   3840
         Picture         =   "frmSupplier.frx":000C
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   32
         Top             =   480
         Width           =   255
      End
      Begin VB.TextBox Text11 
         Appearance      =   0  'Flat
         DataField       =   "Email"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   5280
         MaxLength       =   20
         TabIndex        =   2
         Top             =   840
         Width           =   2415
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         DataField       =   "ContactPerson"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   3
         Top             =   1320
         Width           =   2415
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         DataField       =   "ContactTitle"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   5280
         MaxLength       =   20
         TabIndex        =   4
         Top             =   1320
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "SupplierID"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   1440
         TabIndex        =   0
         Top             =   420
         Width           =   2415
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         DataField       =   "Fax"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   5280
         MaxLength       =   16
         TabIndex        =   6
         Top             =   1800
         Width           =   2415
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         DataField       =   "Phone"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   1440
         MaxLength       =   16
         TabIndex        =   5
         Top             =   1800
         Width           =   2415
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         DataField       =   "Address"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   1440
         MaxLength       =   40
         TabIndex        =   7
         Top             =   2280
         Width           =   6255
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         DataField       =   "CompanyName"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   1
         Top             =   840
         Width           =   2415
      End
      Begin VB.TextBox Text10 
         DataField       =   "CompanyName"
         DataSource      =   "Data3"
         Enabled         =   0   'False
         Height          =   315
         Left            =   1440
         TabIndex        =   28
         Text            =   "Text10"
         Top             =   840
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "E-mail Address"
         Height          =   375
         Index           =   6
         Left            =   4080
         TabIndex        =   30
         Top             =   840
         Width           =   1395
      End
      Begin VB.Label Label2 
         Caption         =   "Contact Person"
         Height          =   375
         Index           =   4
         Left            =   180
         TabIndex        =   27
         Top             =   1320
         Width           =   1395
      End
      Begin VB.Label Label2 
         Caption         =   "Contact Title"
         Height          =   375
         Index           =   3
         Left            =   4080
         TabIndex        =   26
         Top             =   1320
         Width           =   1395
      End
      Begin VB.Label Label2 
         Caption         =   "Supplier Code"
         Height          =   375
         Index           =   5
         Left            =   180
         TabIndex        =   19
         Top             =   420
         Width           =   1275
      End
      Begin VB.Label Label2 
         Caption         =   "Fax Number"
         Height          =   375
         Index           =   2
         Left            =   4080
         TabIndex        =   18
         Top             =   1800
         Width           =   1155
      End
      Begin VB.Label Label2 
         Caption         =   "Phone Number"
         Height          =   375
         Index           =   1
         Left            =   180
         TabIndex        =   17
         Top             =   1800
         Width           =   1155
      End
      Begin VB.Label Label4 
         Caption         =   "Address"
         Height          =   375
         Left            =   180
         TabIndex        =   16
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Company Name"
         Height          =   375
         Index           =   0
         Left            =   180
         TabIndex        =   15
         Top             =   840
         Width           =   1395
      End
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   30
      Left            =   240
      Top             =   360
      Width           =   10575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      Caption         =   "   Supplies"
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   555
      Index           =   0
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   10815
   End
   Begin VB.Label Label5 
      BackColor       =   &H00404040&
      Caption         =   "Label5"
      Height          =   2655
      Index           =   0
      Left            =   180
      TabIndex        =   20
      Top             =   720
      Width           =   7815
   End
End
Attribute VB_Name = "frmSupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AdoSupplier As ADODB.Recordset
Public SelectedDsn As String
Dim edit As Boolean
Private Sub cmd1_Click()
On Error GoTo AddErr
Set rs = New Recordset
    
    Set cn = New ADODB.Connection
    Dim code
    
       'Dim cn As Connection
    Set cn = New ADODB.Connection
    Provider = "MAZIWA"
   cn.Open Provider, "atm", "atm"
    Set rs = New Recordset
sql = ""
sql = "SELECT     TOP 1 * FROM ag_Supplier1 ORDER BY SupplierID desc"
rs.Open sql, cn

If Not rs.EOF Then
 Text1 = Format$(rs.Fields("supplierid") + 1, "0000")
Else
Text1 = "0000"
 Text1 = Format$(Text1.Text, "0001")
End If
'Text6.Text = Val(Text6.Text) + 1
Text6.Text = Format$(Text1.Text, "0000")
If Text2.Enabled = True Then
MsgBox "Please save first before adding", vbExclamation, "Attention"
Else
ButtonEnabled

'Adodc1.Recordset.AddNew
Text1.Text = Text6.Text
Text2.SetFocus
End If
Exit Sub
AddErr:
MsgBox err.description & "", vbExclamation, "Attention"
End Sub

Private Sub cmd4_Click()
On Error GoTo CancelErr
If Text2.Enabled = True Then
'Adodc1.Recordset.CancelUpdate
ButtonEnabled1
Else
MsgBox "Cancel without record to add/edit.", vbExclamation, "Attention"
End If
If Text7.Enabled = False Then
Text6.Text = Val(Text6.Text) - 1
Else
Text7.Enabled = False
End If
Exit Sub
CancelErr:
 MsgBox err.description & "", vbExclamation, "Attention"
End Sub


Private Sub cmd5_Click()
On Error GoTo SearchErr
Dim a
If comSearch.Text = "Supplier Code" Then
a = Combo1.Text
Adodc1.RecordSource = ("Select * from Suppliers where SupplierID = '" & a & "'")
Adodc1.Refresh
If Text2.Text = "" Then
a = Combo1.Text
MsgBox "Record not found, please try again.", vbExclamation, "Attention"
Adodc1.RecordSource = ("Select * from Suppliers order by SupplierID")
Adodc1.Refresh
End If
Else
If comSearch.Text = "Company Name" Then
a = Combo2.Text
Adodc1.RecordSource = ("Select * from Suppliers where CompanyName = '" & a & "'")
Adodc1.Refresh
If Text2.Text = "" Then
a = Combo2.Text
MsgBox "Record not found, please try again.", vbExclamation, "Attention"
Adodc1.RecordSource = ("Select * from Suppliers order by CompanytName")
Adodc1.Refresh
End If
End If
End If
Exit Sub
SearchErr:
  MsgBox err.description & "", vbExclamation, "Attention"
End Sub



Private Sub Combo1_Click()
On Error GoTo Search1Err
Dim a
If comSearch.Text = "Supplier Code" Then
a = Combo1.Text
Adodc1.RecordSource = ("Select * from Suppliers where SupplierID = '" & a & "'")
Adodc1.Refresh
If Text2.Text = "" Then
a = Combo1.Text
MsgBox "Record not found, please try again.", vbExclamation, "Attention"
Adodc1.RecordSource = ("Select * from Suppliers order by SupplierID")
Adodc1.Refresh
End If
Else
If comSearch.Text = "Company Name" Then
a = Combo2.Text
Adodc1.RecordSource = ("Select * from Suppliers where CompanyName = '" & a & "'")
Adodc1.Refresh
If Text2.Text = "" Then
a = Combo2.Text
MsgBox "Record not found, please try again.", vbExclamation, "Attention"
Adodc1.RecordSource = ("Select * from Suppliers order by CompanytName")
Adodc1.Refresh
End If
End If
End If
Exit Sub
Search1Err:
  MsgBox err.description & "", vbExclamation, "Attention"
End Sub

Private Sub Combo2_Click()
On Error GoTo Search2Err
Dim a
If comSearch.Text = "Supplier Code" Then
a = Combo1.Text
Adodc1.RecordSource = ("Select * from Suppliers where SupplierID = '" & a & "'")
Adodc1.Refresh
If Text2.Text = "" Then
a = Combo1.Text
MsgBox "Record not found, please try again.", vbExclamation, "Attention"
Adodc1.RecordSource = ("Select * from Suppliers order by SupplierID")
Adodc1.Refresh
End If
Else
If comSearch.Text = "Company Name" Then
a = Combo2.Text
Adodc1.RecordSource = ("Select * from Suppliers where CompanyName = '" & a & "'")
Adodc1.Refresh
If Text2.Text = "" Then
a = Combo2.Text
MsgBox "Record not found, please try again.", vbExclamation, "Attention"
Adodc1.RecordSource = ("Select * from Suppliers order by CompanytName")
Adodc1.Refresh
End If
End If
End If
Exit Sub
Search2Err:
  MsgBox err.description & "", vbExclamation, "Attention"
End Sub

Private Sub Command1_Click()
Adodc1.RecordSource = ("Select * from Suppliers order by CompanyName")
Adodc1.Refresh
End Sub

Private Sub comSearch_Change()
If comSearch.Text = "Company Name" Then
Combo1.Visible = False
Combo2.Visible = True

Else
If comSearch.Text = "Supplier Code" Then
Combo1.Visible = True
Combo2.Visible = False

End If
End If
End Sub

Private Sub comSearch_LostFocus()
If comSearch.Text = "Company Name" Then
Combo1.Visible = False
Combo2.Visible = True

Else
If comSearch.Text = "Supplier Code" Then
Combo1.Visible = True
Combo2.Visible = False

End If
End If
End Sub


Private Sub cmd6_Click()
On Error Resume Next
If Text2.Enabled = True Then
Exit Sub
End If
Dim a
If Text2.Text = "" Then
MsgBox "No current record to delete.", vbExclamation, "Attention"
Exit Sub
Else
a = MsgBox("Are you sure you want to delete this record?", vbQuestion + vbYesNo, "Confirm")
If a = vbYes Then
On Error Resume Next
Text6.Text = Val(Text6.Text)
Adodc1.Recordset.Delete
Adodc1.Recordset.MoveFirst
Adodc1.RecordSource = ("Select * from Suppliers order by CompanyName")
Adodc1.Refresh
Else
Cancel = True
Exit Sub
End If
End If
End Sub

Private Sub cmd2_Click()
On Error GoTo ModifyErr
If Text2.Enabled = True Then
MsgBox "Please save first before edit.", vbExclamation, "Attention"
Exit Sub
End If
'// UPDATE THE SUPPLIER LIST
edit = True
'Adodc1.Refresh
ButtonEnabled
Text2.SetFocus
Text7.Enabled = True
Exit Sub
ModifyErr:
 MsgBox err.description & "", vbExclamation, "Attention"
End Sub

Private Sub cmd7_Click()
On Error GoTo CloseErr
Dim a
If Text2.Enabled = True Then
a = MsgBox("Do you want to save the record before you exit?", vbYesNo + vbQuestion, "Confirm")
If a = vbYes Then
cmd3_Click
Exit Sub
Else
cmd4_Click
Unload Me
End If
End If
Unload Me
Exit Sub
CloseErr:
 MsgBox err.description & "", vbExclamation, "Attention"
End Sub


Private Sub Form_Load()
Dim X As Integer
Dim Y As Integer
Dim CTL As Object
Dim Provider As String
Provider = "MAZIWA"
For Each CTL In Me
  If TypeOf CTL Is Adodc Then CTL.ConnectionString = Provider
Next
edit = False
 ' Set db = New Connection
 ' db.CursorLocation = adUseClient
 
    Set rst = New Recordset
   
    Set cn = New ADODB.Connection
    Provider = "MAZIWA"
   cn.Open Provider, "atm", "atm"
    Set rst = New Recordset


  
'  Set AdoSupplier = New Recordset
'      AdoSupplier.Open "select SupplierID,CompanyName from AG_Supplier1 Order by CompanyName", cn, adOpenStatic, adLockOptimistic
'  For x = 1 To AdoSupplier.RecordCount
'''    Combo2.AddItem AdoSupplier.Fields("Companyname")
'    AdoSupplier.MoveNext
'  Next x
'  AdoSupplier.Requery
'  For x = 1 To AdoSupplier.RecordCount
'    Combo1.AddItem AdoSupplier.Fields("SupplierID")
'    AdoSupplier.MoveNext
'  Next x
'If Combo1.ListIndex < 0 Then Combo1.Refresh Else Combo1.ListIndex = 0
'
'If Combo2.ListIndex < 0 Then Combo2.Refresh Else Combo2.ListIndex = 0
'  comSearch.AddItem ("Company Name")
'  comSearch.AddItem ("Supplier Code")
'  comSearch.ListIndex = 0
'   With Adodc1
'  .RecordSource = "Select * from ag_Supplier1 order by CompanyName"
'  .Refresh
'End With
'With Adodc2
'  .RecordSource = "SELECT     * From Ctr"
'  .Refresh
'End With
'With Adodc3
'  .RecordSource = "Select * from Suppliers"
'  .Refresh
'End With
  


End Sub
Private Sub Form_Unload(Cancel As Integer)
Dim a

If Cancel = 0 Then
    If Text2.Enabled = True Then
    a = MsgBox("Do you want to save the record before you exit?", vbYesNo + vbQuestion, "Confirm")
    If a = vbYes Then
            Cancel = 1
            cmd3_Click
            Exit Sub
        Else
        Cancel = 1
        cmd4_Click
        Unload Me
         End If
        End If
    End If
End Sub


Private Sub cmd3_Click()
On Error GoTo SaveErr
Dim a
 
    Set rst = New Recordset
    Dim cn As Connection
    Set cn = New ADODB.Connection
    Provider = "MAZIWA"
   cn.Open Provider, "atm", "atm"
    
If Text2.Enabled = False Then
MsgBox "Please add/edit before saving.", vbExclamation, "Attention"
Exit Sub
End If
If Text7.Enabled = False Then
a = Text2.Text
'Adodc3.RecordSource = ("Select * from Suppliers where CompanyName = '" & A & "'")
'Adodc3.Refresh
If Text10.Text = "" Then
'Adodc3.RecordSource = ("Select * from Suppliers order by CompanyName")
'Adodc3.Refresh
Else

End If
Else
Text7.Enabled = False
End If


If edit = True Then

sql = "Update ag_Supplier1"
sql = sql & "   SET CompanyName = '" & Text2 & "', ContactPerson = '" & Text9 & "', ContactTitle = '" & Text8 & "', Address = '" & Text3 & "', Email = '" & Text11 & "', Phone = '" & Text4 & "',"
sql = sql & "    Fax = '" & Text5 & "'"
sql = sql & "  WHERE     (SupplierID = '" & Text1 & "')"
cn.Execute sql
End If
If Text2.Text = "" Then
MsgBox "Please enter the company name.", vbExclamation, "Attention"
Text2.SetFocus
Exit Sub
ElseIf Text11.Text = "" Then
MsgBox "Please enter the email address.", vbExclamation, "Attention"
Text11.SetFocus
Exit Sub
ElseIf Text9.Text = "" Then
MsgBox "Please enter the contact person.", vbExclamation, "Attention"
Text9.SetFocus
Exit Sub
ElseIf Text8.Text = "" Then
MsgBox "Please enter the contact title.", vbExclamation, "Attention"
Text8.SetFocus
Exit Sub
ElseIf Text4.Text = "" Then
MsgBox "Please enter the phone number or enter (-) if phone number is not available.", vbExclamation, "Attention"
Text4.SetFocus
Exit Sub
ElseIf Text5.Text = "" Then
MsgBox "Please enter the fax number or enter (-) if fax number is not available..", vbExclamation, "Attention"
Text5.SetFocus
Exit Sub
ElseIf Text3.Text = "" Then
MsgBox "Please enter the complete address.", vbExclamation, "Attention"
Text3.SetFocus
Exit Sub
End If
'//insert into suppliers
Dim rssup As New ADODB.Recordset
Set rssup = New ADODB.Recordset
sql = "SELECT  ag_Supplier1.SupplierID, ag_Supplier1.CompanyName, ag_Supplier1.ContactPerson, ag_Supplier1.ContactTitle, ag_Supplier1.Address, ag_Supplier1.Email, ag_Supplier1.Phone, ag_Supplier1.Fax, * FROM  ag_Supplier1 where  ag_Supplier1.supplierid='" & Text1 & "'"
rssup.Open sql, cn, adOpenKeyset, adLockOptimistic
'where suppliers.supplierid='0006';
If rssup.EOF Then
sql = "INSERT INTO   ag_Supplier1(SupplierID, CompanyName, ContactPerson, ContactTitle, Address, Email, Phone, Fax)"
sql = sql & " values('" & Text1 & "','" & Text2 & "','" & Text9 & "','" & Text8 & "','" & Text3 & "','" & Text11 & "','" & Text4 & "','" & Text5 & "');"
cn.Execute sql
End If

'cn.Execute sql
MsgBox "Record successfully saved.", vbInformation, "Save"
'Adodc1.Recordset.Update
'Adodc1.RecordSource = ("Select * from  ag_Supplier1 order by CompanyName")
'Adodc1.Refresh
ButtonEnabled1
Exit Sub
SaveErr:
  MsgBox err.description & "", vbExclamation, "Attention"
End Sub


Private Sub Picture5_Click()
 Me.MousePointer = vbHourglass
        frmSearchVendor.Show vbModal
        Text1 = sel
        Text1_Validate True
        Me.MousePointer = 0
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
If Trim(Text1) = "" Then
Exit Sub
End If
Set rs = oSaccoMaster.GetRecordset("SELECT SupplierID, CompanyName, ContactPerson, ContactTitle, Address, Email, Phone, Fax FROM  ag_Supplier1 where SupplierID = '" & Text1 & "'")

If Not rs.EOF Then
If Not IsNull(rs.Fields(1)) Then Text2 = rs.Fields(1)
If Not IsNull(rs.Fields(2)) Then Text9 = rs.Fields(2)
If Not IsNull(rs.Fields(3)) Then Text8 = rs.Fields(3)
If Not IsNull(rs.Fields(4)) Then Text3 = rs.Fields(4)
If Not IsNull(rs.Fields(5)) Then Text11 = rs.Fields(5)
If Not IsNull(rs.Fields(6)) Then Text4 = rs.Fields(6)
If Not IsNull(rs.Fields(7)) Then Text5 = rs.Fields(7)

End If
End Sub

Private Sub text4_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = 45 Or KeyAscii = vbKeyBack Then
Else
   Beep
   KeyAscii = 0
End If
End Sub

Private Sub text5_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = 45 Or KeyAscii = vbKeyBack Then
Else
   Beep
   KeyAscii = 0
End If
End Sub
Public Sub ButtonEnabled()
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text6.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text8.Enabled = True
Text9.Enabled = True
Text11.Enabled = True
cmd1.Enabled = False
cmd2.Enabled = False
cmd3.Enabled = True
cmd4.Enabled = True
cmd6.Enabled = False
'Combo1.Enabled = False
'Combo2.Enabled = False
'comSearch.Enabled = False
'Command1.Enabled = False
Text1.Locked = True
End Sub

Public Sub ButtonEnabled1()
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text9.Enabled = False
Text8.Enabled = False
Text11.Enabled = False
'Combo1.Enabled = True
'Combo2.Enabled = True
'comSearch.Enabled = True
cmd1.Enabled = True
cmd2.Enabled = True
cmd3.Enabled = False
cmd4.Enabled = False
cmd6.Enabled = True
'Command1.Enabled = True
End Sub

