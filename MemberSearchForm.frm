VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Begin VB.Form MemberSearchForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Search Form"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7275
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "MemberSearchForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   7275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkshowallmembers 
      Caption         =   "Show All Members"
      Height          =   255
      Left            =   4440
      TabIndex        =   15
      Top             =   600
      Width           =   2055
   End
   Begin VB.ComboBox cboemployername 
      Height          =   315
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton cmddisplay 
      Caption         =   "Display Members"
      Height          =   310
      Left            =   4320
      TabIndex        =   12
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton cmdMemberSearch 
      Height          =   330
      Left            =   4680
      Picture         =   "MemberSearchForm.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1200
      Width           =   360
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   345
      Left            =   5175
      TabIndex        =   9
      Top             =   5430
      Width           =   1050
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select"
      Height          =   345
      Left            =   3945
      TabIndex        =   8
      Top             =   5430
      Width           =   1140
   End
   Begin MSComctlLib.ListView LstSearch 
      Height          =   3045
      Left            =   15
      TabIndex        =   7
      Top             =   2280
      Width           =   7200
      _ExtentX        =   12700
      _ExtentY        =   5371
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.TextBox TxtValue 
      Height          =   315
      Left            =   3150
      TabIndex        =   6
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox TxtRecords 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      Height          =   330
      Left            =   1320
      TabIndex        =   5
      Top             =   5400
      Width           =   915
   End
   Begin VB.ComboBox cboCriteria 
      Height          =   315
      ItemData        =   "MemberSearchForm.frx":010E
      Left            =   1710
      List            =   "MemberSearchForm.frx":011E
      TabIndex        =   3
      Top             =   1200
      Width           =   1305
   End
   Begin VB.ComboBox CboSearchField 
      Height          =   315
      ItemData        =   "MemberSearchForm.frx":0133
      Left            =   45
      List            =   "MemberSearchForm.frx":0135
      TabIndex        =   1
      Top             =   1200
      Width           =   1530
   End
   Begin VB.Label Label5 
      Caption         =   "Employer Name"
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Value"
      Height          =   180
      Left            =   3165
      TabIndex        =   10
      Top             =   990
      Width           =   990
   End
   Begin VB.Label Label3 
      Caption         =   "Records Found"
      Height          =   285
      Left            =   30
      TabIndex        =   4
      Top             =   5475
      Width           =   1275
   End
   Begin VB.Label Label2 
      Caption         =   "Criteria"
      Height          =   270
      Left            =   1740
      TabIndex        =   2
      Top             =   975
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Search Field"
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   975
      Width           =   1350
   End
End
Attribute VB_Name = "MemberSearchForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cboemployername_Change()
cmddisplay_Click
End Sub

Private Sub cboemployername_Click()
cmddisplay_Click
End Sub

Private Sub CboSearchField_Change()
    If cboSearchField.ListIndex > -1 Then
        With lstSearch
            .Sorted = True
            .SortKey = cboSearchField.ListIndex
        End With
    End If
End Sub

Private Sub CboSearchField_Click()
    
    If cboSearchField.ListIndex > -1 Then
        With lstSearch
            .Sorted = True
            .SortKey = cboSearchField.ListIndex
        End With
    End If
    
End Sub

Private Sub CboSearchField_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        cmdSelect.SetFocus
    End If
End Sub


Private Sub cmdcancel_Click()
    Continue = False
    Unload Me
End Sub

Public Sub selected()
On Error GoTo ErrorHandler
    Continue = True
    Select Case SearchForm
        Case " Membership Registration"
        frmMemRegistration.txtMemNum.Text = lstSearch.SelectedItem.Text
        Case "PHOTO MAP"
        frmphotostatus.txtMemberNo.Text = lstSearch.SelectedItem.Text
        Case "SCHEME REGISTRATION"
        frmschemeregistration.txtMemberNo.Text = lstSearch.SelectedItem.Text
        Case " Generate Loan Schedule"
        frmLoanSchedule.txtMemberNo.Text = lstSearch.SelectedItem.Text
        Case " Next Of Kin"
        frmNextOfKin.txtMemNum.Text = lstSearch.SelectedItem.Text
        Case " Member Statement"
        frmMemStatement.txtMemNum.Text = lstSearch.SelectedItem.Text
        Case "Share Contribution"
        frmContributions.txtMemberNo.Text = lstSearch.SelectedItem.Text
        Case "Share Variation"
        frmShareVar.txtMemberNo.Text = lstSearch.SelectedItem.Text
        Case "Christmas Contribution"
        frmChristmusContrib.txtMemberNo.Text = lstSearch.SelectedItem.Text
        Case "Christmus Contrib"
        frmChristmusvar.txtMemberNo.Text = lstSearch.SelectedItem.Text
        Case "Loan Applications"
        frmLoans.txtMemberNo.Text = lstSearch.SelectedItem.Text
        Case "SaccoMaster: Member Savings"
        frmMembersavings.txtMemberNo.Text = lstSearch.SelectedItem.Text
        Case " Loan Balances"
        frmLoanBal.txtMemberNo.Text = lstSearch.SelectedItem.Text
        Case " Benevolent Fund Processing"
        frmBenFund.txtMemberNo.Text = lstSearch.SelectedItem.Text
        Case "Get Member No"
        frmGuaranto.txtGuarNo.Text = lstSearch.SelectedItem.Text
        Case " Withdrawal Form"
        frmWithdrawal.txtMemberNo.Text = lstSearch.SelectedItem.Text
        Case " MEMBER STATEMENTS"
         frmUtilGenMemStatements.txtMemberNo.Text = lstSearch.SelectedItem.Text
        Case "OFFSETTING PROCESS"
        frmoffsettingsharesonloans.txtMemberNo.Text = lstSearch.SelectedItem.Text
        Case " Loan Applications"
        frmLoans.txtMemberNo.Text = lstSearch.SelectedItem.Text
        Case "  Membership Registration"
        frmMemRegistration.txtMemNum.Text = lstSearch.SelectedItem.Text
        Case "GUARANTOR UPDATE"
        frmmemguarantoupdate.txtMemNum.Text = lstSearch.SelectedItem.Text
        Case "BENEFICIARY DETAILS"
        frmbeneficiary.txtMemNum.Text = lstSearch.SelectedItem.Text
        Case "SHARES / LOANS INQUIRY"
        
        frmloansharesinquiry.txtMemberNo.Text = lstSearch.SelectedItem.Text
        Case "NEW LOAN"
        frmnewloanreconciler.txtMemberNo.Text = lstSearch.SelectedItem.Text
        Case "REFUND PROCESSING"
        frmrefunds.txtMemberNo.Text = lstSearch.SelectedItem.Text
        Case "Over deduction"
        frmDeduction.lsvOverdeduction.SelectedItem.Text = lstSearch.SelectedItem.Text
        Case "Resignation"
        If frmResignation.SSTab1.Tab = 0 Then
        
        frmResignation.txtMemberNo.Text = lstSearch.SelectedItem.Text
        ElseIf frmResignation.SSTab1.Tab = 1 Then
        
        frmResignation.txtMemberNo.Text = lstSearch.SelectedItem.Text
        
        ElseIf frmResignation.SSTab1.Tab = 2 Then
        
        frmResignation.txtRMemberno.Text = lstSearch.SelectedItem.Text
        
        End If
       End Select
    Unload Me
    Exit Sub
ErrorHandler:
    MsgBox err.description
End Sub

Private Sub cmddisplay_Click()
'LstSearch.Refresh
 lstSearch.ListItems.Clear
 If chkshowallmembers = vbChecked Then
    Set rst = oSaccoMaster.GetRecordset("select memberno,accno,staffno," _
    & "surname,othernames from MEMBERS  order by MemberNo")
    Else
       Set rst = oSaccoMaster.GetRecordset("select memberno,accno,staffno," _
    & "surname,othernames from MEMBERS  Where CompanyCode='" & cboemployername & "'   order by MemberNo")
    End If
    With rst
        If Not .EOF Then
            While Not .EOF
                Set li = MemberSearchForm.lstSearch.ListItems.Add(, , !memberno)
                li.SubItems(1) = IIf(IsNull(!StaffNo), "", !StaffNo)
                li.SubItems(2) = IIf(IsNull(!ACCNO), "", !ACCNO)
                li.SubItems(3) = IIf(IsNull(!surname), "", !surname)
                li.SubItems(4) = IIf(IsNull(!OtherNames), "", !OtherNames)
                .MoveNext
            Wend
        End If
    End With
    TxtRecords = rst.RecordCount
    cboSearchField.Text = cboSearchField.List(0)
    cboCriteria.Text = cboCriteria.List(3)
'cmdMemberSearch_Click
End Sub

Private Sub cmdMemberSearch_Click()
On Error GoTo ErrorHandler

If cboCriteria.Text = "Like" Then
If chkshowallmembers = vbChecked Then
Set rst = oSaccoMaster.GetRecordset("select * from members where " & cboSearchField.Text & " like '" & txtValue.Text & "%' ")

Else
Set rst = oSaccoMaster.GetRecordset("select * from members where " & cboSearchField.Text & " like '" & txtValue.Text & "%' and CompanyCode='" & cboemployername & "' ")
End If
Else
If chkshowallmembers = vbChecked Then
Set rst = oSaccoMaster.GetRecordset("select * from members where '" & cboSearchField.Text & "' " & cboCriteria.Text & " '" & txtValue.Text & "' ")
Else
Set rst = oSaccoMaster.GetRecordset("select * from members where '" & cboSearchField.Text & "' " & cboCriteria.Text & " '" & txtValue.Text & "' and CompanyCode='" & cboemployername & "' ")
End If
End If
    With rst
        If Not .EOF Then
            If txtValue.Text <> "" Then
'                If cboCriteria.Text = "Like" Then
'                    .Filter = "" & CboSearchField.Text & " like '" & TxtValue.Text & "*'"
'                Else
'                    .Filter = "" & CboSearchField.Text & " " & cboCriteria.Text & " '" & TxtValue.Text & "'"
'                End If
                If Not .EOF Then
                    lstSearch.ListItems.Clear
                    While Not .EOF
                        Set li = lstSearch.ListItems.Add(, , !memberno)
                        li.SubItems(1) = IIf(IsNull(!StaffNo), "", !StaffNo)
                        li.SubItems(2) = IIf(IsNull(!ACCNO), "", !ACCNO)
                        li.SubItems(3) = IIf(IsNull(!surname), "", !surname)
                        li.SubItems(4) = IIf(IsNull(!OtherNames), "", !OtherNames)
                        .MoveNext
                    Wend
                        
                Else
                    lstSearch.ListItems.Clear
                End If
            .Filter = adFilterNone
            End If
        End If
    End With
    If lstSearch.ListItems.Count > 0 Then
        cmdSelect.Enabled = True
    Else
        cmdSelect.Enabled = False
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox err.description
    
    
End Sub

Private Sub cmdSelect_Click()
    selected
End Sub

Private Sub Form_Load()
On Error GoTo ErrorHandler
'//load the employer
Dim myclass As cdbase
Set myclass = New cdbase
Provider = myclass.OpenCon
Set cn = New ADODB.Connection
cn.Open Provider, "atm", "atm"
Set rs = New ADODB.Recordset
sql = ""
sql = "SELECT  distinct   CompanyCode   FROM       members order by companycode"
rs.Open sql, cn
While Not rs.EOF

cboemployername.AddItem rs.Fields(0)
rs.MoveNext
Wend
Set rs = Nothing
Set cn = Nothing
With Me
        .lstSearch.ColumnHeaders.Add , , "MemberNo"
        .lstSearch.ColumnHeaders.Add , , "StaffNo"
        .lstSearch.ColumnHeaders.Add , , "AccNo"
        .lstSearch.ColumnHeaders.Add , , "Surname"
        .lstSearch.ColumnHeaders.Add , , "Othernames"
    End With
    
    I = 1
    Do Until I > MemberSearchForm.lstSearch.ColumnHeaders.Count
        MemberSearchForm.cboSearchField.AddItem MemberSearchForm.lstSearch.ColumnHeaders.Item(I).Text
        I = I + 1
    Loop
Exit Sub
ErrorHandler:
MsgBox err.description
    End Sub

Private Sub lstSearch_DblClick()
    selected
End Sub
Private Sub TxtValue_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdMemberSearch.SetFocus
    End If
End Sub
