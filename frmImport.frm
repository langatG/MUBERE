VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmImport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Import Records"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6555
   Icon            =   "frmImport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5400
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   22
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImport.frx":030A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2520
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Height          =   345
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   945
      Width           =   375
   End
   Begin VB.TextBox txtErrorLog 
      Height          =   3195
      Left            =   2760
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   1560
      Width           =   3645
   End
   Begin VB.ListBox lvwImportFields 
      Height          =   3210
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   5
      Top             =   1560
      Width           =   2565
   End
   Begin VB.CheckBox chkImportAllFields 
      Caption         =   "Import All Fields"
      Height          =   240
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   1830
   End
   Begin VB.TextBox txtFileName 
      Height          =   315
      Left            =   150
      TabIndex        =   1
      Top             =   960
      Width           =   4455
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   1323
      ButtonWidth     =   1138
      ButtonHeight    =   1164
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Import"
            Key             =   "bImport"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "Error Log"
      Height          =   195
      Left            =   2760
      TabIndex        =   4
      Top             =   1320
      Width           =   2475
   End
   Begin VB.Label Label1 
      Caption         =   "File Name"
      Height          =   180
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   2145
   End
End
Attribute VB_Name = "frmImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strFileName As String

Private Sub Command1_Click()
    On Error GoTo errFix
    CommonDialog1.ShowOpen
    strFileName = CommonDialog1.FileName
    frmImport.txtFileName.Text = strFileName
    Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Import"
End Sub



Private Sub lvwImportFields_ItemCheck(Item As Integer)
On Error GoTo errFix
Select Case lvwImportFields.List(Item)
    Case "Member No"
    If lvwImportFields.selected(Item) = False Then
        frmImport.txtErrorLog.Text = ""
        frmImport.txtErrorLog.Text = "Member No has to be imported"
        lvwImportFields.selected(Item) = True
    End If
    Case "Surname"
    If lvwImportFields.selected(Item) = False Then
        frmImport.txtErrorLog.Text = ""
        frmImport.txtErrorLog.Text = "Surname has to be imported"
        lvwImportFields.selected(Item) = True
    End If

    Case "OtherNames"
    If lvwImportFields.selected(Item) = False Then
        frmImport.txtErrorLog.Text = ""
        frmImport.txtErrorLog.Text = "Other Names have to be imported"
        lvwImportFields.selected(Item) = True
    End If
    Case "AsAtDate"
    If lvwImportFields.selected(Item) = False Then
        frmImport.txtErrorLog.Text = ""
        frmImport.txtErrorLog.Text = "As At Date has to be imported"
        lvwImportFields.selected(Item) = True
    End If
 
    Case "InitShares"
    If lvwImportFields.selected(Item) = False Then
        frmImport.txtErrorLog.Text = ""
        frmImport.txtErrorLog.Text = "Initial shares has to be imported"
        lvwImportFields.selected(Item) = True
    End If
    Case "ID No"
    If lvwImportFields.selected(Item) = False Then
        frmImport.txtErrorLog.Text = ""
        frmImport.txtErrorLog.Text = "ID No has to be imported"
        lvwImportFields.selected(Item) = True
    End If
    Case "Init Monthly Contr"
    If lvwImportFields.selected(Item) = False Then
        frmImport.txtErrorLog.Text = ""
        frmImport.txtErrorLog.Text = "Initial Monthly Contr has to be imported"
        lvwImportFields.selected(Item) = True
    End If
    Case Else
    frmImport.txtErrorLog.Text = ""
    End Select
    Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Import"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo errFix
Select Case Button.Key
    Case "bImport"
    Set rstRecordsImported = Read_Excel(txtFileName.Text)
    If Not IsEmpty(rstRecordsImported) Then
    Call formCallingImport.onImpBtnOfImportFrmClick
    Else
    txtErrorLog.Text = "No records in your excel sheet"
    End If
End Select
Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Import"
End Sub
Public Function Read_Excel(ByVal sFile As String) As ADODB.Recordset
     On Error GoTo fix_err
     Dim rs As ADODB.Recordset
     Set rs = New ADODB.Recordset
     Dim sconn As String
     rs.CursorLocation = adUseClient
     rs.CursorType = adOpenKeyset
     rs.LockType = adLockBatchOptimistic
     sconn = "DRIVER=Microsoft Excel Driver (*.xls);" & "DBQ=" & sFile
     rs.Open "SELECT * FROM [sheet1$]", sconn
     If Not rs.EOF Then
     Set Read_Excel = rs
     Else
        Exit Function
     End If
     Set rs = Nothing
     Exit Function
fix_err:
    txtErrorLog.Text = err.description + " " + err.Source
     Debug.Print err.description & " " & _
                 err.Source, vbCritical, "Import"
     err.Clear
End Function

