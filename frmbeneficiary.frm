VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmbeneficiary 
   Caption         =   "BENEFICIARY DETAILS"
   ClientHeight    =   4815
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8175
   Icon            =   "frmbeneficiary.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4815
   ScaleWidth      =   8175
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Print Register"
      Height          =   375
      Left            =   5520
      TabIndex        =   42
      Top             =   4320
      Width           =   1215
   End
   Begin MSComctlLib.ListView lvwNextOfKin 
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   1200
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   5318
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   16711680
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.ComboBox cboCompany 
      Height          =   315
      Left            =   1080
      Style           =   1  'Simple Combo
      TabIndex        =   39
      Text            =   "cboCompany"
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txtComapny 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2760
      TabIndex        =   38
      Top             =   720
      Width           =   4335
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "<>"
      Height          =   300
      Left            =   2400
      TabIndex        =   37
      Top             =   240
      Width           =   345
   End
   Begin VB.Frame fraNOKDetails 
      Caption         =   "Kin Details"
      Height          =   2895
      Left            =   120
      TabIndex        =   15
      Top             =   1200
      Width           =   7935
      Begin VB.TextBox txtmobileNo 
         Height          =   375
         Left            =   5880
         TabIndex        =   45
         Top             =   2040
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker DTPSignDate 
         Height          =   375
         Left            =   120
         TabIndex        =   41
         Top             =   1440
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   140902401
         CurrentDate     =   40534
      End
      Begin VB.TextBox txtKinNo 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   885
         TabIndex        =   25
         Top             =   360
         Width           =   1305
      End
      Begin VB.TextBox txtKinNames 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2400
         TabIndex        =   24
         Top             =   360
         Width           =   3495
      End
      Begin VB.TextBox txtAddress 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   180
         TabIndex        =   23
         Top             =   885
         Width           =   3300
      End
      Begin VB.TextBox txtHomeTelNo 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   3600
         TabIndex        =   22
         Top             =   885
         Width           =   1890
      End
      Begin VB.TextBox txtOfficeTelNo 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   5520
         TabIndex        =   21
         Top             =   885
         Width           =   1980
      End
      Begin VB.TextBox txtComments 
         ForeColor       =   &H00FF0000&
         Height          =   630
         Left            =   105
         MultiLine       =   -1  'True
         TabIndex        =   20
         Top             =   2040
         Width           =   2730
      End
      Begin VB.TextBox txtIdNo 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1680
         TabIndex        =   19
         Top             =   1440
         Width           =   2130
      End
      Begin VB.ComboBox cboRelationship 
         ForeColor       =   &H00FF0000&
         Height          =   315
         ItemData        =   "frmbeneficiary.frx":030A
         Left            =   6000
         List            =   "frmbeneficiary.frx":0320
         TabIndex        =   18
         Top             =   360
         Width           =   1665
      End
      Begin VB.TextBox txtWitness 
         ForeColor       =   &H00FF0000&
         Height          =   645
         Left            =   2880
         TabIndex        =   17
         Top             =   2040
         Width           =   2925
      End
      Begin VB.TextBox txtPercentage 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   4125
         TabIndex        =   16
         Text            =   "0"
         Top             =   1440
         Width           =   2670
      End
      Begin VB.Label lblmobileno 
         Caption         =   "Mobile No:"
         Height          =   255
         Left            =   5880
         TabIndex        =   44
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   15
         Left            =   5640
         TabIndex        =   43
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Kin No"
         Height          =   225
         Left            =   165
         TabIndex        =   36
         Top             =   360
         Width           =   570
      End
      Begin VB.Label Label7 
         Caption         =   "Kin Names"
         Height          =   195
         Left            =   2460
         TabIndex        =   35
         Top             =   135
         Width           =   2190
      End
      Begin VB.Label Label8 
         Caption         =   "Relationship"
         Height          =   195
         Left            =   6120
         TabIndex        =   34
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "Address"
         Height          =   315
         Left            =   120
         TabIndex        =   33
         Top             =   720
         Width           =   1200
      End
      Begin VB.Label Label10 
         Caption         =   "Mobile No"
         Height          =   195
         Left            =   3600
         TabIndex        =   32
         Top             =   675
         Width           =   1800
      End
      Begin VB.Label Label11 
         Caption         =   "OfficeTel No"
         Height          =   195
         Left            =   5520
         TabIndex        =   31
         Top             =   720
         Width           =   1515
      End
      Begin VB.Label Label12 
         Caption         =   "Sign Date"
         Height          =   225
         Left            =   120
         TabIndex        =   30
         Top             =   1200
         Width           =   1515
      End
      Begin VB.Label Label16 
         Caption         =   "Comments"
         Height          =   180
         Left            =   120
         TabIndex        =   29
         Top             =   1800
         Width           =   1860
      End
      Begin VB.Label Label13 
         Caption         =   "ID No"
         Height          =   225
         Left            =   1860
         TabIndex        =   28
         Top             =   1230
         Width           =   645
      End
      Begin VB.Label label14 
         Caption         =   "Witness"
         Height          =   180
         Left            =   2880
         TabIndex        =   27
         Top             =   1860
         Width           =   1995
      End
      Begin VB.Label Label15 
         Caption         =   "Percentage"
         Height          =   195
         Left            =   4125
         TabIndex        =   26
         Top             =   1245
         Width           =   2115
      End
   End
   Begin VB.TextBox txtMemNum 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   1080
      TabIndex        =   12
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox txtNames 
      Appearance      =   0  'Flat
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   240
      Width           =   4335
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "CLOSE"
      Height          =   495
      Left            =   6720
      TabIndex        =   10
      Top             =   4200
      Width           =   1110
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
      Height          =   495
      Left            =   240
      TabIndex        =   9
      ToolTipText     =   "Move to the Last record"
      Top             =   4200
      Width           =   495
   End
   Begin VB.CommandButton cmdAdd 
      Height          =   495
      Left            =   2880
      Picture         =   "frmbeneficiary.frx":0354
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Add New record"
      Top             =   4200
      Width           =   495
   End
   Begin VB.CommandButton cmdPrev 
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
      Height          =   495
      Left            =   720
      TabIndex        =   7
      ToolTipText     =   "Move to the Previous record"
      Top             =   4200
      Width           =   495
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
      Height          =   495
      Left            =   1680
      Picture         =   "frmbeneficiary.frx":0886
      TabIndex        =   6
      ToolTipText     =   "Move to Last record"
      Top             =   4200
      Width           =   495
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
      Height          =   495
      Left            =   1200
      TabIndex        =   5
      ToolTipText     =   "Move to the Next"
      Top             =   4200
      Width           =   495
   End
   Begin VB.CommandButton cmdEdit 
      Height          =   495
      Left            =   3360
      Picture         =   "frmbeneficiary.frx":0988
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Edit Record"
      Top             =   4200
      Width           =   495
   End
   Begin VB.CommandButton cmdDelete 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      Picture         =   "frmbeneficiary.frx":0A8A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Delete Record"
      Top             =   4200
      Width           =   495
   End
   Begin VB.CommandButton cmdUpdate 
      Height          =   495
      Left            =   4320
      Picture         =   "frmbeneficiary.frx":0B8C
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Save Record"
      Top             =   4200
      Width           =   495
   End
   Begin VB.CommandButton cmdCancel 
      Height          =   495
      Left            =   4800
      Picture         =   "frmbeneficiary.frx":0C8E
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Cancel Process"
      Top             =   4200
      Width           =   495
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "Company"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   120
      TabIndex        =   40
      Top             =   750
      Width           =   945
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      Caption         =   "MemberNo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   135
      TabIndex        =   14
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Names"
      Height          =   255
      Left            =   2760
      TabIndex        =   13
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "frmbeneficiary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim disablemodifying As Boolean
Private Sub cboCompany_Change()
    Set rscompany = oSaccoMaster.GetRecordset("Select companyname  from company where companycode='" & cboCompany & "'")
    With rscompany
        If Not .EOF Then
            txtComapny.Text = !CompanyName
        Else
            txtComapny.Text = ""
        End If
    End With
End Sub

Private Sub cboRelationship_Change()
On Error GoTo errFix
    cboRelationship.Text = UCase(cboRelationship.Text)
    cboRelationship.SelStart = Len(cboRelationship.Text)
      Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Next Of Kin"
End Sub

Private Sub cboRelationship_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
        txtAddress.SetFocus
End If

If Len(cboRelationship.Text) > 14 Then
      Beep
      MsgBox "Can't enter more than 15 characters", vbExclamation
      KeyAscii = 8
  End If

  Select Case KeyAscii
    'Case Asc("vbBack")
    Case Asc("A") To Asc("Z")
    Case Asc("a") To Asc("z")
    Case Asc(" ")
    'Case Asc("'")
    Case Is = 8

    Case Else
    Beep
    KeyAscii = 0
  End Select
  Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Company Setup"
End Sub



Private Sub cmdAdd_Click()
On Error GoTo errFix
    action = "addingRecords"
    Set Rst5 = oSaccoMaster.GetRecordset("SELECT Sum(KIN.Percentage) AS SumOfPercentage FROM KIN WHERE KIN.MEMBERNO= '" & txtMemNum.Text & "'")
    If Not Rst5.EOF Then
        If Not IsNull(Rst5!SumOfPercentage) Then
            If Rst5!SumOfPercentage = 100 Then
                MsgBox "Adjust percentages of other kin to accomodate new kin ,100% ,already allocated", vbInformation, ""
                fraNOKDetails.Visible = True
                'Exit Sub
            End If
        End If
    End If
     Set rst = oSaccoMaster.GetRecordset("select * from members ")
     If rst.RecordCount > 0 Then
     rst.Find ("memberno= '" & txtMemNum.Text & "'")
        If Not rst.EOF Then
           MyBookMark = rst.Bookmark
        End If
     End If
     
    cmdCancel.Enabled = True
    cmdUpdate.Enabled = True
    Cmdedit.Enabled = False
    cmdCancel.Enabled = True
    
    txtKinNo.Text = newKinNo
    
    txtKinNames.Text = ""
    cboRelationship.Text = ""
    txtHomeTelNo.Text = ""
    txtAddress.Text = ""
    txtOfficeTelNo.Text = ""
    DTPSignDate.value = Date
    txtWitness.Text = ""
    txtPercentage.Text = ""
    txtcomments.Text = ""
    txtIdNo.Text = ""
    fraNOKDetails.Enabled = True
    txtKinNo.Visible = True
    txtKinNo.Locked = False
    txtKinNames.Visible = True
    cboRelationship.Visible = True
    txtHomeTelNo.Visible = True
    txtAddress.Visible = True
    txtOfficeTelNo.Visible = True
    DTPSignDate.Visible = True
    txtWitness.Visible = True
    txtPercentage.Visible = True
    txtcomments.Visible = True
    cmddelete.Enabled = False
    lvwNextOfKin.Visible = False
    txtIdNo.Visible = True
    txtKinNo.SetFocus
    Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Company Setup"
End Sub
Private Sub cmdcancel_Click()
On Error GoTo errFix
    txtKinNo.Text = ""
    txtKinNames.Text = ""
    cboRelationship.Text = ""
    txtHomeTelNo.Text = ""
    txtAddress.Text = ""
    txtOfficeTelNo.Text = ""
    DTPSignDate.value = Date
    txtWitness.Text = ""
    txtPercentage.Text = ""
    txtcomments.Text = ""
    txtIdNo.Text = ""
    txtKinNo.Visible = False
    txtKinNames.Visible = False
    cboRelationship.Visible = False
    txtHomeTelNo.Visible = False
    txtAddress.Visible = False
    txtWitness.Visible = False
    txtPercentage.Visible = False
    txtOfficeTelNo.Visible = False
    DTPSignDate.Visible = False
    txtcomments.Visible = False
    cmddelete.Enabled = True
    lvwNextOfKin.Visible = True
    txtIdNo.Visible = False
    cmdAdd.Enabled = True
    Cmdedit.Enabled = True
    cmdCancel.Enabled = False
    cmdUpdate.Enabled = False
    Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Company Setup"
End Sub
Private Sub cmdclose_Click()
    Unload Me
End Sub
Private Sub cmddelete_Click()
On Error GoTo errFix
    If lvwNextOfKin.ListItems.Count > 0 Then
        sel = lvwNextOfKin.SelectedItem
    End If
    Set rst = oSaccoMaster.GetRecordset("select * from kin where kinno= '" & sel & "'")
    If rst.RecordCount > 0 Then
        If MsgBox("Are you sure you want to delete " & rst!KinNames & " ? ", vbYesNo, "Kin deletion") = vbYes Then
            Set rst = oSaccoMaster.GetRecordset("select * from kin where kinno= '" & sel & "'")
            rst.Delete
            rst.Update
            Load_Records
        End If
    End If
     Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Company Setup"
End Sub
Private Sub cmdedit_Click()
On Error GoTo errFix
    If lvwNextOfKin.ListItems.Count < 1 Then
        MsgBox "No selected item to edit", vbInformation, "No Item"
        Exit Sub
    Else
    action = "editingRecords"
    Set rst = oSaccoMaster.GetRecordset("select * from members ")
     If rst.RecordCount > 0 Then
     rst.Find ("memberno= '" & txtMemNum.Text & "'")
        If Not rst.EOF Then
           MyBookMark = rst.Bookmark
        End If
     End If
'    Toolbar1.Buttons("bSearch").Enabled = True
'    Toolbar1.Buttons("bPrint").Enabled = False
    cmdCancel.Enabled = True
    cmdAdd.Enabled = False
    cmdUpdate.Enabled = True
    Cmdedit.Enabled = False
    sel = lvwNextOfKin.SelectedItem
    End If
    If Not sel = "" Then
        
            sel = lvwNextOfKin.SelectedItem
            Set rst = oSaccoMaster.GetRecordset("select * from kin where kinno= '" & sel & "'")
    
            With rst
'                If !kinsigned = "Yes" Then
'                    optKinSigned.Value = vbChecked
'                Else
'                    optKinSigned.Value = vbUnchecked
'                End If
                fraNOKDetails.Enabled = True
                txtKinNo.Locked = True
                txtMemNum.Text = !memberno & ""
                txtAddress.Text = !Address & ""
                txtKinNames.Text = !KinNames & ""
                txtIdNo.Text = !idno & ""
                cboRelationship.Text = !Relationship & ""
                txtHomeTelNo.Text = !HomeTelNo & ""
                txtOfficeTelNo.Text = !OfficeTelNo & ""
                txtKinNo.Text = !kinno & ""
                txtWitness.Text = !Witness & ""
                txtPercentage.Text = !Percentage & ""
                txtcomments.Text = !Comments & ""
                txtKinNo.Visible = True
                txtKinNames.Visible = True
                cboRelationship.Visible = True
                txtHomeTelNo.Visible = True
                txtAddress.Visible = True
                txtOfficeTelNo.Visible = True
                'optKinSigned.Visible = True
                DTPSignDate.Visible = True
                txtWitness.Visible = True
                txtPercentage.Visible = True
                txtcomments.Visible = True
                cmddelete.Enabled = False
                lvwNextOfKin.Visible = False
                lvwNextOfKin.View = lvwReport
                txtIdNo.Visible = True
            End With
     End If
     Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Company Setup"
End Sub

Private Sub cmdFind_Click()
    frmSearchMembers.Show vbModal
    mno = SearchValue
    If mno <> "" Then
        txtMemNum.Text = SearchValue
        mno = txtMemNum
        Set rst = oSaccoMaster.GetRecordset("select companycode,Companyname from company where companycode=(select companycode from members where memberno='" & mno & "')")
        If Not rst.EOF Then
            cboCompany.Text = rst(0)
        End If
    End If

End Sub

Private Sub cmdFirst_Click()
On Error GoTo errFix
    Toolbar1.Buttons("bSearch").Enabled = True
    Toolbar1.Buttons("bPrint").Enabled = True
    Set rst = oSaccoMaster.GetRecordset("select memberno from members order by memberno")
    With rst
        If .RecordCount > 0 Then
            .MoveFirst
            cmdFirst.Enabled = False
            cmdPrev.Enabled = False
            cmdNext.Enabled = True
            cmdLast.Enabled = True
            txtMemNum.Text = rst!memberno & ""
            Load_Records
        End If
    End With
    rst.Close
     Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Company Setup"
End Sub
Private Sub cmdLast_Click()
On Error GoTo errFix
    Toolbar1.Buttons("bSearch").Enabled = True
    Toolbar1.Buttons("bPrint").Enabled = True
    Set rst = oSaccoMaster.GetRecordset("select memberno from members order by memberno")
    With rst
        If .RecordCount > 0 Then
            .MoveLast
            cmdFirst.Enabled = True
            cmdPrev.Enabled = True
            cmdNext.Enabled = False
            cmdLast.Enabled = False
            txtMemNum.Text = rst!memberno & ""
            Load_Records
        End If
    End With
    rst.Close
     Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Company Setup"
End Sub
Private Sub cmdNext_Click()
On Error GoTo errFix
    If action = "addingRecords" Or action = "editingRecords" Then
    Set rst = oSaccoMaster.GetRecordset("select * from members ")
     If rst.RecordCount > 0 Then
     rst.Bookmark = MyBookMark
     txtMemNum.Text = rst!memberno & ""
     End If
    End If
    Toolbar1.Buttons("bSearch").Enabled = True
    Toolbar1.Buttons("bPrint").Enabled = True
    Set rst = oSaccoMaster.GetRecordset("select memberno from members order by memberno")
    With rst
        If .RecordCount > 0 Then
            .MoveFirst
            .Find "memberno= '" & txtMemNum.Text & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                .MoveNext
                If .EOF Then
                    .MoveLast
                    cmdFirst.Enabled = True
                    cmdPrev.Enabled = True
                    cmdNext.Enabled = False
                    cmdLast.Enabled = False
                Else
                    cmdFirst.Enabled = True
                    cmdPrev.Enabled = True
                    cmdNext.Enabled = True
                    cmdLast.Enabled = True
                End If
                txtMemNum.Text = !memberno & ""
                Load_Records
            End If
        End If
    End With
    rst.Close
     Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Company Setup"
End Sub
Private Sub cmdPrev_Click()
On Error GoTo errFix
    If action = "addingRecords" Or action = "editingRecords" Then
    Set rst = oSaccoMaster.GetRecordset("select * from members ")
     If rst.RecordCount > 0 Then
     rst.Bookmark = MyBookMark
     txtMemNum.Text = rst!memberno & ""
     End If
    End If
    Toolbar1.Buttons("bSearch").Enabled = True
    Toolbar1.Buttons("bPrint").Enabled = True
    Set rst = oSaccoMaster.GetRecordset("select memberno from members order by memberno")
    With rst
        If .RecordCount > 0 Then
            .MoveFirst
            .Find "memberno= '" & txtMemNum.Text & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                .MovePrevious
            If .BOF Then
                .MoveFirst
                cmdFirst.Enabled = False
                cmdPrev.Enabled = False
                cmdNext.Enabled = True
                cmdLast.Enabled = True
            Else
                cmdFirst.Enabled = True
                cmdPrev.Enabled = True
                cmdNext.Enabled = True
                cmdLast.Enabled = True
            End If
            txtMemNum.Text = !memberno & ""
            Load_Records
            End If
        End If
        
    End With
    rst.Close
     Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Company Setup"
End Sub
Private Sub CommandButton5_Click()
    Unload Me
End Sub
Private Sub cmdupdate_Click()
On Error GoTo errFix
            While txtKinNo.Text = ""
                MsgBox "Enter kin number", vbInformation, "Kin No"
                txtKinNo.SetFocus
                Exit Sub
            Wend
            
            While txtKinNames.Text = ""
                MsgBox "Enter kin names", vbInformation, "Kin Names"
                txtKinNames.SetFocus
                Exit Sub
            Wend
            
            While cboRelationship.Text = ""
                MsgBox "Select relationship", vbInformation, "Relationship"
                cboRelationship.SetFocus
                Exit Sub
            Wend
            While txtAddress.Text = ""
                MsgBox "Enter address", vbInformation, "Address"
                txtAddress.SetFocus
                Exit Sub
            Wend
            If txtPercentage.Text = "" Then
                If MsgBox("Enter Percentage, Disregard This?", vbYesNo, "Percentage") = vbNo Then
                txtPercentage.SetFocus
                Exit Sub
                End If
            End If
            While DTPSignDate.value = ""
                MsgBox "Enter signdate", vbInformation, "Kin No"
                DTPSignDate.SetFocus
                Exit Sub
            Wend
If action = "addingRecords" Then
        cmdCancel.Enabled = False
        Cmdedit.Enabled = True
        Set rst = oSaccoMaster.GetRecordset("select * from kin where kinno='" & txtKinNo & "'")
        With rst
            .AddNew
            !memberno = txtMemNum.Text & ""
            !kinno = Trim(txtKinNo.Text & "")
            !KinNames = txtKinNames.Text & ""
            !Relationship = cboRelationship.Text & ""
            !Address = txtAddress.Text & ""
            !auditid = User
            !audittime = Format(Now, "DD/MM/YYYY hh:mm:ss")
            !idno = txtIdNo.Text & ""
            !HomeTelNo = txtHomeTelNo.Text & ""
            !OfficeTelNo = txtOfficeTelNo.Text & ""
            !mobileNo = txtMobileNo.Text & ""
            If Not DTPSignDate.value = "" Then
                !SignDate = DTPSignDate.value & ""
            End If
            !Witness = txtWitness.Text & ""
            If txtPercentage.Text = "" Then
            !Percentage = 0
            Else
            !Percentage = txtPercentage.Text
            End If
            !Comments = txtcomments.Text & ""
            .Update
        txtKinNo.Visible = False
        txtKinNames.Visible = False
        cboRelationship.Visible = False
        txtHomeTelNo.Visible = False
        txtAddress.Visible = False
        txtOfficeTelNo.Visible = False
        DTPSignDate.Visible = False
        txtWitness.Visible = False
        txtPercentage.Visible = False
        txtcomments.Visible = False
        cmddelete.Enabled = True
        lvwNextOfKin.Visible = True
        txtIdNo.Visible = False
        Load_Records
        End With
        cmdUpdate.Enabled = False
ElseIf action = "editingRecords" Then
        cmdCancel.Enabled = False
        cmdAdd.Enabled = True
            Set rst = oSaccoMaster.GetRecordset("select * from kin where kinno= '" & sel & "'")
            With rst
            If .RecordCount > 0 Then
                !KinNames = txtKinNames.Text & ""
                !kinno = txtKinNo.Text & ""
                !Address = txtAddress.Text & ""
                !idno = txtIdNo.Text & ""
                !HomeTelNo = txtHomeTelNo.Text & ""
                !auditid = User
                !audittime = Format(Now, "DD/MM/YYYY hh:mm:ss")
                !OfficeTelNo = txtOfficeTelNo.Text & ""
                If Not DTPSignDate.value = "" Then
                    !SignDate = DTPSignDate.value & ""
                End If
                !Witness = txtWitness.Text & ""
                !Percentage = txtPercentage.Text & ""
                !Comments = txtcomments.Text & ""
'                If optKinSigned.Value = vbChecked Then
'                    !kinsigned = "Yes"
'                End If
                .Update
            End If
            txtKinNo.Visible = False
            txtKinNames.Visible = False
            cboRelationship.Visible = False
            txtHomeTelNo.Visible = False
            txtAddress.Visible = False
            txtOfficeTelNo.Visible = False
            'optKinSigned.Visible = False
            DTPSignDate.Visible = False
            txtWitness.Visible = False
            txtPercentage.Visible = False
            txtcomments.Visible = False
            cmddelete.Enabled = True
            lvwNextOfKin.Visible = True
            txtIdNo.Visible = False
            End With
            txtMemNum.Locked = False
            Load_Records
            cmdUpdate.Enabled = False
End If
txtKinNo.Text = newKinNo
action = ""
fraNOKDetails.Enabled = True
 Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Next Of Kin"
End Sub

Private Sub Command1_Click()
reportname = "KinRegisterReport.rpt"
STRFORMULA = ""
Show_Sales_Crystal_Report STRFORMULA, reportname, CompanyName
End Sub

Private Sub DTPSignDate_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    txtPercentage.SetFocus
End If
 Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Next Of Kin"
End Sub

Private Sub DTPSignDate_LostFocus()
On Error GoTo errFix
Set Rst3 = oSaccoMaster.GetRecordset("select ApplicDate from members where memberno= '" & txtMemNum.Text & "'")
If Not Rst3.EOF Then
    If DTPSignDate.value < Rst3!ApplicDate Then
        MsgBox "Sign date should not be earlier than members registration date", vbInformation, "Sign Date"
        DTPSignDate.value = Date
        DTPSignDate.SetFocus
        Exit Sub
    End If
End If
If DTPSignDate.value > Date Then
    MsgBox "Sign date should not be beyond today", vbInformation, "Sign Date"
    DTPSignDate.value = Date
    DTPSignDate.SetFocus
    Exit Sub
End If
 Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Next Of Kin"
End Sub

Function newKinNo()
    Dim jno As String
    Set rst = oSaccoMaster.GetRecordset("Select count(*)+1 from kin where memberno='" & txtMemNum & "'")
    newKinNo = "K" & txtMemNum & "-" & Format(rst(0), "00")
End Function
Private Sub Form_Load()
'On Error GoTo errFix
    PositionForm Me
    txtKinNames.Visible = False
    txtKinNo.Visible = False
    cboRelationship.Visible = False
    txtHomeTelNo.Visible = False
    txtAddress.Visible = False
    txtOfficeTelNo.Visible = False
    DTPSignDate.Visible = False
    txtWitness.Visible = False
    txtPercentage.Visible = False
    fraNOKDetails.Enabled = False
    txtcomments.Visible = False
    cmddelete.Enabled = True
    cmdCancel.Enabled = False
    cmdUpdate.Enabled = False
    lvwNextOfKin.Visible = True
    txtIdNo.Visible = False
    Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Next Of Kin"
End Sub
Public Sub Load_Records()
'On Error GoTo errFix
    cmdCancel.Enabled = False
    If disablemodifying = False Then
    Cmdedit.Enabled = True
    End If
    txtKinNo.Visible = False
    txtKinNames.Visible = False
    cboRelationship.Visible = False
    txtHomeTelNo.Visible = False
    txtAddress.Visible = False
    txtOfficeTelNo.Visible = False
    DTPSignDate.Visible = False
    txtWitness.Visible = False
    txtPercentage.Visible = False
    txtcomments.Visible = False
    If disablemodifying = False Then
    cmddelete.Enabled = True
    End If
    lvwNextOfKin.Visible = True
    txtIdNo.Visible = False
    cmdCancel.Enabled = False
    fraNOKDetails.Enabled = False
    action = ""
    With lvwNextOfKin
        .ColumnHeaders.Clear
        .ListItems.Clear
        .ColumnHeaders.Add 1, "hKinNo", "Kin No", 3000
        .ColumnHeaders.Add 2, "hNextOfKin", "Next Of Kin", 3000
        .ColumnHeaders.Add 3, "hAddress", "Address", 3000
        .ColumnHeaders.Add 4, "hIdNo", "ID No", 3000
        .ColumnHeaders.Add 5, "hRelationship", "Relationship", 3000
        .ColumnHeaders.Add 6, "hHomeTel", "Home Tel No", 3000
        .ColumnHeaders.Add 7, "hSigningDate", "Signing Date", 3000
        .ColumnHeaders.Add 8, "hPercentage", "Percentage", 3000
        .ColumnHeaders.Add 8, "hWitness", "Witness", 3000
    Set rst1 = oSaccoMaster.GetRecordset("select * from kin where memberno= '" & txtMemNum.Text & "'")
    With rst1
        If .RecordCount > 0 Then
            While Not rst1.EOF
                Set li = lvwNextOfKin.ListItems.Add(, , rst1!kinno)
                li.SubItems(1) = rst1!KinNames & ""
                li.SubItems(2) = rst1!Address & ""
                li.SubItems(3) = rst1!idno & ""
                li.SubItems(4) = rst1!Relationship & ""
                li.SubItems(5) = rst1!HomeTelNo & ""
                li.SubItems(6) = rst1!SignDate & ""
                li.SubItems(8) = rst1!Percentage & ""
                li.SubItems(7) = rst1!Witness & ""
                'LI.SubItems(8)=RST
              .MoveNext
            Wend
        Else
            lvwNextOfKin.ListItems.Clear
        End If
    End With
    rst1.Close
    End With
    Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Next Of Kin"
End Sub
Function EncryptPassword()
On Error GoTo errFix
    Dim Pwd As Variant
    Dim Temp As String, PwdChr As Long
    Dim EncryptKey As Long
    Pwd = valToEncrOrDecr
    EncryptKey = Int(Sqr(Len(Pwd) * 81)) + 23
    
    For PwdChr = 1 To Len(Pwd)
        Temp = Temp + Chr(Asc(Mid(Pwd, PwdChr, 1)) Xor EncryptKey)
    Next PwdChr
    
    EncryptPass = Temp
 Exit Function
errFix:
    MsgBox err.description, vbOKOnly, "Next Of Kin"
End Function

Public Sub searchSelect()
On Error GoTo errFix
    txtMemNum.Text = sel
    Load_Records
    frmSearch.Visible = False
     Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Next Of Kin"
End Sub
Public Sub onRefreshOfSearchFrm()
On Error GoTo errFix
    frmSearch.lstSearch.ListItems.Clear
    strSQL = "Select * from members order by memberno"
    Set rst = oSaccoMaster.GetRecordset(strSQL)
    With rst
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                Set li = frmSearch.lstSearch.ListItems.Add(, , !memberno)
                li.SubItems(1) = !StaffNo & ""
                li.SubItems(2) = !surname & ""
                li.SubItems(3) = !OtherNames & ""
                li.SubItems(4) = !idno & ""
                li.SubItems(5) = !employer & ""
                li.SubItems(6) = !companycode & ""
                .MoveNext
            Loop
        End If
        .Close
    End With
    Set rst = Nothing
    frmSearch.txtFrom.Text = ""
    frmSearch.txtTo.Text = ""
    frmSearch.cboCrieria.Text = "="
    frmSearch.cboField.Text = "Member No"
     Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Next Of Kin"
End Sub
Public Sub onFindOfSearchFrmClick()
On Error GoTo errFix
    Dim Find As Integer
     Select Case frmSearch.cboField.Text
         Case "Member No"
             searchField = "memberno"
         Case "Surname"
             searchField = "surname"
         Case "Other Names"
             searchField = "othernames"
         Case "ID No"
             searchField = "idno"
         Case "Staff No"
             searchField = "staffno"
         Case "Employer"
             searchField = "employer"
         Case "Company Code"
             searchField = "companycode"
    End Select
    frmSearch.lstSearch.ListItems.Clear
    If Not frmSearch.cboField.Text = "" Then
        If Not frmSearch.cboCrieria.Text = "" Then
            If Not frmSearch.cboCrieria.Text = "Between" And Not frmSearch.cboCrieria.Text = "Like" Then
                strSQL = "Select * from members where " & searchField & " " & frmSearch.cboCrieria.Text & " '" & frmSearch.txtFrom.Text & "'"
                Set rst = oSaccoMaster.GetRecordset(strSQL)
                
                With rst
                    If .RecordCount > 0 Then
                        .MoveFirst
                        Do While Not .EOF
                            Set li = frmSearch.lstSearch.ListItems.Add(, , !memberno & "")
                            li.ListSubItems.Add , , !surname & ""
                            li.ListSubItems.Add , , !OtherNames & ""
                            .MoveNext
                        Loop
                    End If
                End With
                Set rst = Nothing
            ElseIf frmSearch.cboCrieria.Text = "Like" Then
                sql = "Select * from members order by memberno"
                Set rst = oSaccoMaster.GetRecordset(strSQL)
                With rst
                    If .RecordCount > 0 Then
                        .MoveFirst
                        Do While Not .EOF
                            .Find "" & searchField & " " & frmSearch.cboCrieria.Text & " '" & frmSearch.txtFrom.Text & "%'", , adSearchForward
                            If Not .EOF Then
                                Set li = frmSearch.lstSearch.ListItems.Add(, , !memberno & "")
                                li.ListSubItems.Add , , !surname & ""
                                li.ListSubItems.Add , , !OtherNames & ""
                                .MoveNext
                            End If
                        Loop
                    End If
                End With
                Set rst = Nothing
            Else
                    strSQL = "select * from members where " & searchField & " between'" & txtFrom.Text & "' And '" & txtTo.Text & " '"
                    Set rst = oSaccoMaster.GetRecordset(strSQL)
                With rst
                    If .RecordCount > 0 Then
                        .MoveFirst
                        Do While Not .EOF
                            Set li = frmSearch.lstSearch.ListItems.Add(, , !memberno & "")
                            li.SubItems(1) = !StaffNo & ""
                            li.SubItems(2) = !surname & ""
                            li.SubItems(3) = !OtherNames & ""
                            li.SubItems(4) = !idno & ""
                            li.SubItems(5) = !employer & ""
                            li.SubItems(6) = !companycode & ""
                            .MoveNext
                        Loop
                    End If
                End With
                Set rst = Nothing
            End If
        Else
            MsgBox "Select the search criteria.", vbExclamation
        End If
    Else
        MsgBox "Select the search field.", vbExclamation
    End If
     Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Next Of Kin"
End Sub

Private Sub Form_Unload(Cancel As Integer)
action = ""
End Sub

Private Sub optKinSigned_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    txtWitness.SetFocus
ElseIf KeyAscii = 43 Then
    optKinSigned.value = vbChecked
End If
 Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Next Of Kin"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo errFix
    Select Case Button.Key
    Case "bSearch"
        txtKinNames.Visible = False
        txtKinNo.Visible = False
        cboRelationship.Visible = False
        txtHomeTelNo.Visible = False
        txtAddress.Visible = False
        txtOfficeTelNo.Visible = False
        optKinSigned.Visible = False
        DTPSignDate.Visible = False
        txtWitness.Visible = False
        txtPercentage.Visible = False
        txtcomments.Visible = False
        cmddelete.Enabled = True
        lvwNextOfKin.Visible = True
        txtIdNo.Visible = False
        SearchForm = Me.Caption
        MemberSearchForm.Show vbModal, Me
        
        If Continue = False Then
            Exit Sub
        End If
        Load_Records
    Case "bPrint"
        reportType = "Summary Report"
        Set formCallingRangeSelector = frmNextOfKin
        frmRangeSelection.txtTo.Text = ""
        frmRangeSelection.txtFrom.Text = ""
        frmRangeSelection.Show vbModal
        frmRangeSelection.Caption = "Select Member Range"
        frmRangeSelection.lblFrom.Caption = "From Member No"
        frmRangeSelection.lblTo.Caption = "To Member No"
    End Select
     Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Next Of Kin"
End Sub
Public Sub onOkOfRangeForm()
On Error GoTo errFix
Set rst = oSaccoMaster.GetRecordset("select memberno from members order by memberno")
If Not frmRangeSelection.txtFrom.Text = "" Then
    Set Rst4 = oSaccoMaster.GetRecordset("select memberno from members where memberno= '" & frmRangeSelection.txtFrom.Text & "'")
    If Not Rst4.EOF Then
        rangeFrom = frmRangeSelection.txtFrom.Text
    Else
        MsgBox "Enter an existing No.", vbInformation, "Non-existent No"
        frmRangeSelection.txtFrom.Text = ""
        frmRangeSelection.txtFrom.SetFocus
        Exit Sub
    End If
Else
    If Not rst.EOF Then
        rst.MoveFirst
        rangeFrom = rst!memberno
    End If
End If
If Not frmRangeSelection.txtTo.Text = "" Then
    Set Rst4 = oSaccoMaster.GetRecordset("select memberno from members where memberno= '" & frmRangeSelection.txtTo.Text & "'")
    If Not Rst4.EOF Then
        rangeTo = frmRangeSelection.txtTo.Text
    Else
        MsgBox "Enter an existing Member No.", vbInformation, "Non-existent No"
        frmRangeSelection.txtTo.Text = ""
        frmRangeSelection.txtTo.SetFocus
        Exit Sub
    End If
Else
    
    If Not rst.EOF Then
        rst.MoveLast
        rangeTo = rst!memberno
    End If
End If
If rangeFrom > rangeTo Then
    MsgBox "'To member no should be greater than or equal to 'From memberno'", vbInformation, "Check Range"
    frmRangeSelection.txtTo.SetFocus
    Exit Sub
End If

Set rst = oSaccoMaster.GetRecordset("select memberno from members where memberno= '" & rangeTo & "'")
Set rst1 = oSaccoMaster.GetRecordset("select memberno from members where memberno= '" & rangeFrom & "'")
If rst.RecordCount = 0 Then
    MsgBox "Enter an existing member no for 'From' field", vbInformation, "Member No"
    Exit Sub
End If
If rst1.RecordCount = 0 Then
    MsgBox "Enter an existing member no for 'To' field", vbInformation, "Member No"
    Exit Sub
End If
If rst.RecordCount > 0 And rst1.RecordCount > 0 Then
     frmRangeSelection.Visible = False
        
        Set rst = oSaccoMaster.GetRecordset("select memberno from members")
        If Not rst.EOF Then
            Set a = New CRAXDRT.Application
            Set r = a.OpenReport(App.path & "\Reports\Kin Report.rpt")
            r.RecordSelectionFormula = "{MEMBERS.MemberNo}= '" & rangeFrom & "' to '" & rangeTo & "'"
            r.ReadRecords
            Set Rst5 = oSaccoMaster.GetRecordset("select companyname from sysparam")
            If Not Rst5.EOF Then
                r.ReportTitle = Rst5!CompanyName & ""
            End If
            If chkPrevReport.value = vbUnchecked Then
                r.PrintOut
            Else
                
                With frmReports
                    .Show vbModal, Me
                End With
            End If
            Set r = Nothing
        Else
            MsgBox "No records", vbInformation, "Reports"
        End If
           
 End If
  Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Next Of Kin"
End Sub


Private Sub txtAddress_Change()
On Error GoTo errFix
txtAddress.Text = UCase(txtAddress.Text)
txtAddress.SelStart = Len(txtAddress.Text)
 Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Next Of Kin"
End Sub

Private Sub txtAddress_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If Len(txtAddress.Text) > 49 Then
      Beep
      MsgBox "Can't enter more than 50 characters", vbExclamation
      KeyAscii = 8
End If
If KeyAscii = 13 Then
    txtHomeTelNo.SetFocus
End If
 Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Next Of Kin"
End Sub
Private Sub txtComments_Change()
On Error GoTo errFix
txtcomments.Text = UCase(txtcomments.Text)
txtcomments.SelStart = Len(txtcomments.Text)
 Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Next Of Kin"
End Sub

Private Sub txtComments_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If Len(txtcomments.Text) > 99 Then
    Beep
    MsgBox "Can't enter more than 100 characters", vbExclamation
    KeyAscii = 8
End If

If KeyAscii = 13 Then
    cmdUpdate.SetFocus
End If
 Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Next Of Kin"
End Sub

Private Sub txtHomeTelNo_Change()
On Error GoTo errFix
txtHomeTelNo.Text = UCase(txtHomeTelNo.Text)
txtHomeTelNo.SelStart = Len(txtHomeTelNo.Text)
 Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Next Of Kin"
End Sub

Private Sub txtHomeTelNo_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
   If KeyAscii = 13 Then
        txtOfficeTelNo.SetFocus
    End If
If Len(txtHomeTelNo.Text) > 14 Then
        Beep
        MsgBox "Can't enter more than 15 characters", vbExclamation
        KeyAscii = 8
  End If
  Select Case KeyAscii
    Case Asc("0") To Asc("9")
    Case Asc("-")
    Case Asc("+")
    Case Asc(",")
    Case Asc(" ")
    Case Asc(")")
    Case Asc("(")
    Case Is = 8
    Case Else
        Beep
        KeyAscii = 0
  End Select
   Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Next Of Kin"
End Sub
Private Sub txtIDNo_Change()
On Error GoTo errFix
txtIdNo.Text = UCase(txtIdNo.Text)
txtIdNo.SelStart = Len(txtIdNo.Text)
 Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Next Of Kin"
End Sub

Private Sub txtidno_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If Len(txtIdNo.Text) > 9 Then
      Beep
      MsgBox "Can't enter more than 10 characters", vbExclamation
      KeyAscii = 8
End If
If KeyAscii = 13 Then
    txtPercentage.SetFocus
End If
 Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Next Of Kin"
End Sub

Private Sub txtKinNames_Change()
On Error GoTo errFix
txtKinNames.Text = UCase(txtKinNames.Text)
txtKinNames.SelStart = Len(txtKinNames.Text)
 Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Next Of Kin"
End Sub

Private Sub txtKinNames_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
  If KeyAscii = 13 Then
        cboRelationship.SetFocus
    End If

If Len(txtKinNames.Text) > 49 Then
      Beep
      MsgBox "Can't enter more than 50 characters", vbExclamation
      KeyAscii = 8
End If

  Select Case KeyAscii
    'Case Asc("vbBack")
    Case Asc("A") To Asc("Z")
    Case Asc("a") To Asc("z")
    Case Asc(" ")
    'Case Asc("'")
    Case Is = 8

    Case Else
    Beep
    KeyAscii = 0
  End Select
   Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Next Of Kin"
End Sub

Private Sub txtKinNo_Change()
On Error GoTo errFix
sql = "select * from kin where kinno='" & txtKinNo.Text & "'"
        Set rst = oSaccoMaster.GetRecordset(sql)
        With rst
            If .EOF Then
                Exit Sub
            End If
            
            txtKinNames.Text = !KinNames
            cboRelationship.Text = !Relationship
            txtAddress.Text = !Address
            txtHomeTelNo.Text = !HomeTelNo
            DTPSignDate.value = IIf(IsNull(!SignDate), Get_Server_Date, !SignDate)
            txtPercentage.Text = IIf(IsNull(!Percentage), 0, !Percentage)
            txtcomments.Text = !Comments
            txtWitness.Text = !Witness
            'ListView1.Visible = False
        End With
        Cmdedit.Enabled = True
        
        txtKinNo.Text = UCase(txtKinNo.Text)
        txtKinNo.SelStart = Len(txtKinNo.Text)
 Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Next Of Kin"
End Sub





Private Sub txtKinNo_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If Len(txtKinNo.Text) > 49 Then
        Beep
        MsgBox "Can't enter more than 50 characters", vbExclamation
        KeyAscii = 8
  End If
If KeyAscii = 13 Then
    If action = "addingRecords" Then
        Set rst = oSaccoMaster.GetRecordset("select kinno from kin where kinno= '" & txtKinNo.Text & "' and memberno='" & txtMemNum & "'")
        If Not rst.EOF Then
            MsgBox "Kin with same no exists", vbInformation, "Kin No Exists"
            txtKinNo.Text = ""
            txtKinNo.SetFocus
            Exit Sub
        Else
           
        End If
    End If
     txtKinNames.SetFocus
End If
 Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Next Of Kin"
End Sub

Private Sub txtKinNo_LostFocus()
'On Error GoTo errFix
'If action = "addingRecords" Then
'    Set rst = oSaccoMaster.GetRecordSet("select kinno from kin where kinno= '" & txtKinNo.Text & "'")
'    If Not rst.EOF Then
'        MsgBox "Kin with same no exists", vbInformation, "Kin No Exists"
'         txtKinNo.Text = ""
'         txtKinNo.SetFocus
'        Exit Sub
'    End If
'End If
'txtKinNames.SetFocus
' Exit Sub
'errFix:
'    MsgBox Err.Description, vbOKOnly, "Next Of Kin"
End Sub

Private Sub txtMemNum_Change()
    mysql = ""
    mysql = "select surname,othernames,HomeAddr,companycode  from members  where memberno ='" & txtMemNum & "'"
    Set rs = oSaccoMaster.GetRecordset(mysql)
    If Not rs.EOF Then
        txtNames = rs!surname & "  " & rs!OtherNames
        cboCompany.Text = rs!companycode
        Load_Records
    Else
        txtNames = ""
        Exit Sub
    End If

End Sub

Private Sub txtMemNum_Click()
    txtMemNum_Change
End Sub

Private Sub txtOfficeTelNo_Change()
On Error GoTo errFix
txtOfficeTelNo.Text = UCase(txtOfficeTelNo.Text)
txtOfficeTelNo.SelStart = Len(txtOfficeTelNo.Text)
 Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Next Of Kin"
End Sub

Private Sub txtOfficeTelNo_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
   If KeyAscii = 13 Then
        DTPSignDate.SetFocus
    End If
If Len(txtOfficeTelNo.Text) > 14 Then
        Beep
        MsgBox "Can't enter more than 15 characters", vbExclamation
        KeyAscii = 8
  End If
  Select Case KeyAscii
    Case Asc("0") To Asc("9")
    Case Asc("-")
    Case Asc("+")
    Case Asc(",")
    Case Asc(" ")
    Case Asc(")")
    Case Asc("(")
    Case Is = 8
    Case Else
        Beep
        KeyAscii = 0
  End Select
   Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Next Of Kin"
End Sub

Private Sub txtPercentage_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
   optKinSigned.SetFocus
End If
If Len(txtPercentage.Text) > 3 Then
        Beep
        MsgBox "Can't enter more than 4 characters", vbExclamation
        KeyAscii = 8
  End If
  Select Case KeyAscii
    Case Asc("0") To Asc("9")
    Case 46
    Case Is = 8
    Case Else
        Beep
        KeyAscii = 0
  End Select
   Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Next Of Kin"
    End Sub

Private Sub txtPercentage_LostFocus()
On Error GoTo errFix
    Set rst = oSaccoMaster.GetRecordset("select sum(percentage) as SumOfPercentage from kin where memberno='" & txtMemNum.Text & "'")
    If Not IsNull(rst!SumOfPercentage) Then
        If action = "addingRecords" Then
            If Not txtPercentage.Text = "" Then
                percentageDifference = 100 - rst!SumOfPercentage
                If percentageDifference < txtPercentage.Text Then
                    MsgBox "This kin can have maximum of " & percentageDifference & " %", vbInformation, "Percentage"
                    txtPercentage.Text = ""
                    txtPercentage.SetFocus
                    Exit Sub
                End If
            End If
        ElseIf action = "editingRecords" Then
            Set Rst5 = oSaccoMaster.GetRecordset("select percentage from kin where kinno= '" & txtKinNo.Text & "'")
            If Not Rst5.EOF Then
                percentageDifference = 100 - (rst!SumOfPercentage - Rst5!Percentage)
            End If
            If Not txtPercentage.Text = "" Then
                If percentageDifference < txtPercentage.Text Then
                    MsgBox "This kin can have maximum of " & percentageDifference & " %", vbInformation, "Percentage"
                    txtPercentage.Text = ""
                    txtPercentage.SetFocus
                    Exit Sub
                End If
            End If
        End If
    End If
 Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Next Of Kin"
End Sub

Private Sub txtWitness_Change()
On Error GoTo errFix
txtWitness.Text = UCase(txtWitness.Text)
txtWitness.SelStart = Len(txtWitness.Text)
 Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Next Of Kin"
End Sub

Private Sub txtWitness_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
 If KeyAscii = 13 Then
        txtcomments.SetFocus
    End If

If Len(txtWitness.Text) > 49 Then
      Beep
      MsgBox "Can't enter more than 50 characters", vbExclamation
      KeyAscii = 8
  End If

  Select Case KeyAscii
    'Case Asc("vbBack")
    Case Asc("A") To Asc("Z")
    Case Asc("a") To Asc("z")
    Case Asc(" ")
    'Case Asc("'")
    Case Is = 8

    Case Else
    Beep
    KeyAscii = 0
  End Select
   Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Next Of Kin"
End Sub

