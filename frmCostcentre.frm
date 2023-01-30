VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmCostcentre 
   BackColor       =   &H80000001&
   Caption         =   "Cost Centres"
   ClientHeight    =   5190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7815
   Icon            =   "frmCostcentre.frx":0000
   LinkTopic       =   "Form2"
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   7815
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fralis 
      Height          =   4455
      Left            =   120
      TabIndex        =   9
      Top             =   600
      Width           =   7575
      Begin VB.Frame fraCcentre 
         Height          =   3615
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Visible         =   0   'False
         Width           =   7335
         Begin VB.Frame fraCcentr 
            Height          =   2775
            Left            =   360
            TabIndex        =   12
            Top             =   240
            Width           =   6375
            Begin VB.TextBox txtComments 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   885
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   8
               Top             =   1440
               Width           =   6135
            End
            Begin VB.TextBox txtDescription 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   315
               Left            =   2040
               TabIndex        =   7
               Top             =   840
               Width           =   4215
            End
            Begin VB.TextBox txtCostcode 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   315
               Left            =   120
               TabIndex        =   6
               Top             =   840
               Width           =   1815
            End
            Begin VB.Label Label3 
               Caption         =   "Comments"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   15
               Top             =   1200
               Width           =   1215
            End
            Begin VB.Label lblDescription 
               Caption         =   "Description"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   2040
               TabIndex        =   14
               Top             =   600
               Width           =   1095
            End
            Begin VB.Label lblCode 
               Caption         =   "Cost Centre Code"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   13
               Top             =   600
               Width           =   1815
            End
         End
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   495
         Left            =   3360
         Picture         =   "frmCostcentre.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Cancel Process"
         Top             =   3840
         Width           =   615
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   495
         Left            =   2640
         Picture         =   "frmCostcentre.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Save Record"
         Top             =   3840
         Width           =   615
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   495
         Left            =   1920
         Picture         =   "frmCostcentre.frx":0646
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Delete Record"
         Top             =   3840
         Width           =   615
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   495
         Left            =   1200
         Picture         =   "frmCostcentre.frx":0748
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Edit Record"
         Top             =   3840
         Width           =   615
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6360
         TabIndex        =   5
         Top             =   3840
         Width           =   975
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&New"
         Height          =   495
         Left            =   480
         Picture         =   "frmCostcentre.frx":084A
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Add New record"
         Top             =   3840
         Width           =   615
      End
      Begin MSComctlLib.ListView lvwcostcent 
         Height          =   3495
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   6165
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   16711680
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   615
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   1085
      ButtonWidth     =   1482
      ButtonHeight    =   1032
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgCTool"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   " Search "
            Key             =   "search"
            Object.ToolTipText     =   "Search for a record"
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Reports"
            ImageIndex      =   2
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "LCost"
                  Text            =   "List of Cost centers"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "DCost"
                  Text            =   "Cost Center Details"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "ECost"
                  Text            =   "Cost-Center's Employees"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList imgCTool 
         Left            =   5520
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCostcentre.frx":0D7C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCostcentre.frx":0E8E
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmCostcentre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim SaveNew As Boolean
Dim CostSet As Boolean
Dim resp As String

Private Sub cmdclose_Click()
    Set CSecurity = Nothing
    Unload Me
End Sub

Private Sub cmddelete_Click()

If CostSet = False Then
    MsgBox "Select the costcenter you would like to delete.", vbInformation
    Exit Sub
End If

If lvwcostcent.SelectedItem = " " Then
    MsgBox "Select the costcenter you would like to delete.", vbInformation
    Exit Sub
End If

resp = MsgBox("Are you sure you want to delete the record.", vbQuestion + vbYesNo)
If resp = vbNo Then
    Exit Sub
End If

CConnect.cnnConnect
sql = "Select * from costcent"
CConnect.Openrs


    With rs
        If .RecordCount > 0 Then
            .MoveFirst
            .Find "CostCode like '*" & lvwcostcent.SelectedItem & "*'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                .Delete
                .Update
            End If
        End If
    End With
    
Set rs = Nothing
Call loadlvweployees

Set cnnPayroll = Nothing

End Sub

Private Sub cmdedit_Click()
'Set rs = New ADODB.Recordset
'Set cn = New ADODB.Connection
'cn.Open MDBase



If lvwcostcent.ListItems.Count < 1 Then
    MsgBox "Select the costcenter you would like to edit.", vbInformation
    Exit Sub
End If

If lvwcostcent.SelectedItem = "" Then
    MsgBox "Select the costcenter you would like to edit.", vbInformation
    Exit Sub
End If

'rs.Open "select * from costcent", cn, adOpenKeyset, adLockOptimistic
Call cleartxt


sql = "Select * from d_CostCent"
Set rs = oSaccoMaster.GetRecordset(sql)

With rs
    If Not rs.EOF Then
        .MoveFirst
        .Find "costcode like '*" & lvwcostcent.SelectedItem & "*'", , adSearchForward, adBookmarkFirst
        
        If Not .EOF Then
        
            txtCostcode = !CostCode & ""
            txtdescription = !description & ""
            txtcomments.Text = !Comments & ""
            
            Call Disablecmd
            cmdsave.Enabled = True
            cmdCancel.Enabled = True
            fraCcentre.Visible = True
            txtCostcode.Enabled = False
            
            SaveNew = False
            
        Else
            Set rs = Nothing
            Set cnnPayroll = Nothing
            
            Exit Sub
        End If
    Else
        Set rs = Nothing
        Set cnnPayroll = Nothing
        
        Exit Sub
    End If
    
End With

Set rs = Nothing
Set cnnPayroll = Nothing
            
End Sub

Private Sub cmdEmployees_Click()

End Sub

Private Sub cmdEmp_Click()

End Sub

Private Sub cmdsave_Click()

If txtCostcode.Text = "" Then
    MsgBox "You have to enter the cost center code.", vbExclamation
    txtCostcode.SetFocus
    Exit Sub
End If

If txtdescription.Text = "" Then
    MsgBox "You have to enter the costcenter name.", vbExclamation
    txtdescription.SetFocus
    Exit Sub
End If


sql = "Select * from d_costcent"
Set rs = oSaccoMaster.GetRecordset(sql)

With rs
    If SaveNew = True Then
        If .RecordCount > 0 Then
            .MoveFirst
            .Find "costcode like '*" & txtCostcode.Text & "*'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                MsgBox "Cost center code already exists.", vbExclamation
                txtCostcode.Text = ""
                txtCostcode.SetFocus
                Exit Sub
            End If
        End If

        .AddNew
        !CostCode = txtCostcode.Text & ""
        !description = txtdescription.Text & ""
        !Comments = txtcomments.Text & ""
        
        
            !UUser = User
            !TTime = Date & " " & Time
      
        
        .Update
        
    Else
        If .RecordCount > 0 Then
            .MoveFirst
            .Find "costcode like '*" & txtCostcode.Text & "*'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                !CostCode = txtCostcode.Text & ""
                !description = txtdescription.Text & ""
                !Comments = txtcomments.Text & ""
                
                
                    !UUser = User
                    !TTime = Get_Server_Date
               
                
                .Update
                
            End If
        End If
    
    End If
    
End With

Set rs = Nothing
Call loadlvweployees
Set cnnPayroll = Nothing

Call Enablecmd
cmdsave.Enabled = False
cmdCancel.Enabled = False
fraCcentre.Visible = False

SaveNew = False

End Sub

Private Sub cmdcancel_Click()
    Call Enablecmd
    cmdsave.Enabled = False
    cmdCancel.Enabled = False
    fraCcentre.Visible = False
    
    SaveNew = False
    
End Sub

Private Sub cmdAdd_Click()
    Call cleartxt
    
    Call Disablecmd
    cmdsave.Enabled = True
    cmdCancel.Enabled = True
    fraCcentre.Visible = True
    txtCostcode.Enabled = True
    On Error Resume Next
    txtCostcode.SetFocus
    
    SaveNew = True
    
End Sub


Private Sub cmdupdate_Click()

End Sub

Private Sub Form_Load()




fraCcentre.Visible = False


With lvwcostcent
    .ColumnHeaders.Add , , "costcode"
    .ColumnHeaders.Add , , "Description", 3000
    .ColumnHeaders.Add , , "Comments", 4000
    .View = lvwReport

End With


Call loadlvweployees

Set cnnPayroll = Nothing

cmdsave.Enabled = False
cmdCancel.Enabled = False



End Sub


Private Sub Form_Resize()
'oSmart.FResize Me, Me.Height, Me.Width

'With frmMain
'        Me.Move .CoolBar1.Width + 230, 1290 ', Abs(.Width - (.CoolBar1.Width + 480))
'    End With

'With frmCostcentre

'Me.ScaleHeight = 6000
'Me.ScaleWidth = 5000
'Me.ScaleTop = 1560
'fraCcentre.Height = Me.ScaleHeight
'fraCcentre.Width = Me.ScaleWidth
'fraCcentre.Top = Me.ScaleTop
'End With
End Sub

Private Sub loadlvweployees()
On Error GoTo ErrorHandler
lvwcostcent.ListItems.Clear
'Dim lis As ListItem
Set rs = New ADODB.Recordset
Set cn = New ADODB.Connection
cn.Open "MAZIWA", "atm"
'rs.Open "select * from costcent", cn, adOpenKeyset, adLockOptimistic
'
CostSet = False

sql = "Select * from d_costcent"
Set rs = oSaccoMaster.GetRecordset(sql)

With rs
    If .RecordCount > 0 Then
        .MoveFirst
        Do While Not rs.EOF
            Set lis = lvwcostcent.ListItems.Add(, , !CostCode & "")
            lis.ListSubItems.Add , , !description & ""
            lis.ListSubItems.Add , , !Comments & ""
            CostSet = True
            
            .MoveNext
          
        Loop
        .MoveFirst
    End If
End With

Set rs = Nothing
Exit Sub
ErrorHandler:
    MsgBox "The Data Base is Either Disconnected or Does Not Exist", vbInformation, "Data Base Connection"
End Sub



Private Sub Disablecmd()
Dim I As Object
    For Each I In Me
        If TypeOf I Is CommandButton Then
            I.Enabled = False
        End If
    Next I
End Sub

Private Sub Enablecmd()
Dim I As Object
    For Each I In Me
        If TypeOf I Is CommandButton Then
            I.Enabled = True
        End If
    Next I
End Sub

Private Sub cleartxt()
Dim I As Object
    For Each I In Me
        If TypeOf I Is TextBox Then
            I.Text = ""
        End If
    Next I
End Sub

Private Sub Form_Unload(Cancel As Integer)
   ' frmMain.lblCostCenters.ForeColor = vbBlue
    Set CSecurity = Nothing
    
End Sub

Private Sub lvwcostcent_DblClick()
frmCostcentre.txtCostcode.Text = lvwcostcent.SelectedItem.Text

CConnect.cnnConnect
sql = "Select * from costcent"
CConnect.Openrs

With rs
    .Find "costcode='" & lvwcostcent.SelectedItem.Text & " '"
    If .EOF Then
        MsgBox "record not found."
    Else
   
        frmCostcentre.txtCostcode.Text = rs!CostCode
        frmCostcentre.txtdescription.Text = rs!description
        frmCostcentre.txtcomments.Text = rs!Comments
        fraCcentre.Visible = True
        cmdsave.Enabled = True
        cmdCancel.Enabled = True
    End If
End With

Set rs = Nothing
Set cnnPayroll = Nothing

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Me.MousePointer = vbHourglass

    Select Case Button.Key
        Case "search"

    End Select

Me.MousePointer = 0
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
On Error GoTo ErrorHandler
Me.MousePointer = vbHourglass

  Select Case ButtonMenu.Key
      
      Case "LCost"
          
          Set a = New Application
          Set r = a.OpenReport(path & "LCost.rpt")
          
          
          r.ReadRecords
          
          With frmReports.CRViewer1
              .ReportSource = r
              .ViewReport
          End With
          
          frmReports.Show vbModal
          
          Set r = Nothing
          
      Case "DCost"
          Set a = New Application
          Set r = a.OpenReport(path & "DCost.rpt")
          
         
          
          r.ReadRecords
          
          With frmReports.CRViewer1
              .ReportSource = r
              .ViewReport
          End With
          
          frmReports.Show vbModal
          
          Set r = Nothing
          
      Case "ECost"
          Set a = New Application
          Set r = a.OpenReport(path & "ECost.rpt")
          
          r.ReadRecords
          
          With frmReports.CRViewer1
              .ReportSource = r
              .ViewReport
          End With
          
          frmReports.Show vbModal
          
          Set r = Nothing
             
  End Select

Me.MousePointer = 0
Exit Sub
ErrorHandler:
MsgBox "Either the report doesnot exist or the report path is wrong", vbInformation, "Report"
End Sub

Private Sub txtComments_KeyPress(KeyAscii As Integer)
    If Len(Trim(txtcomments.Text)) > 50 Then
        Beep
        MsgBox "Can't enter more than 50 characters", vbExclamation
        KeyAscii = 8
    End If
  Select Case KeyAscii
    Case Asc("A") To Asc("Z")
    Case Asc("a") To Asc("z")
    'Case Asc("'")
    Case Asc(" ")
    'Case Asc("vbBack")
    Case Is = 8
    Case Else
        Beep
        KeyAscii = 0
  End Select
End Sub

Private Sub txtCostcode_Change()
    txtCostcode.Text = UCase(txtCostcode.Text)
    txtCostcode.SelStart = Len(txtCostcode.Text)
    
End Sub

Private Sub txtCostcode_KeyPress(KeyAscii As Integer)
    If Len(Trim(txtCostcode.Text)) > 15 Then
        Beep
        MsgBox "Can't enter more than 15 characters", vbExclamation
        KeyAscii = 8
    End If
  Select Case KeyAscii
    Case Asc("A") To Asc("Z")
    Case Asc("a") To Asc("z")
    Case Asc("0") To Asc("9")
    Case Asc("/")
    Case Asc("-")
    Case Is = 8
    Case Else
        Beep
        KeyAscii = 0
  End Select
End Sub

Private Sub txtDescription_KeyPress(KeyAscii As Integer)
    If Len(Trim(txtdescription.Text)) > 40 Then
        Beep
        MsgBox "Can't enter more than 40 characters", vbExclamation
        KeyAscii = 8
    End If
  Select Case KeyAscii
    Case Asc("A") To Asc("Z")
    Case Asc("a") To Asc("z")
    'Case Asc("'")
    Case Asc(" ")
    'Case Asc("vbBack")
    Case Is = 8
    Case Else
        Beep
        KeyAscii = 0
  End Select
End Sub
