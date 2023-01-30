VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmchartsofaccounts 
   Caption         =   "GL Chart Of Accounts"
   ClientHeight    =   7425
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11760
   Icon            =   "frmchartsofaccounts.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   11760
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdprintledgers 
      Caption         =   "Print "
      Height          =   375
      Left            =   7200
      TabIndex        =   13
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      Height          =   375
      Left            =   10680
      TabIndex        =   12
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CommandButton cmdhistory 
      Caption         =   "History"
      Height          =   375
      Left            =   9240
      TabIndex        =   11
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CommandButton cmdbudget 
      Caption         =   "Budget"
      Height          =   375
      Left            =   5640
      TabIndex        =   10
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CommandButton cmdcomparison 
      Caption         =   "Comparison"
      Height          =   375
      Left            =   4320
      TabIndex        =   9
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CommandButton cmdtransactions 
      Caption         =   "Transactions"
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton cmdopen 
      Caption         =   "Open"
      Height          =   375
      Left            =   960
      TabIndex        =   7
      Top             =   6960
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker transdate 
      Height          =   315
      Left            =   1680
      TabIndex        =   6
      Top             =   555
      Width           =   1710
      _ExtentX        =   3016
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   " dd-MM-yyyy"
      Format          =   179634179
      CurrentDate     =   39030
   End
   Begin VB.PictureBox Picture5 
      Height          =   285
      Left            =   3390
      Picture         =   "frmchartsofaccounts.frx":0442
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   3
      Top             =   120
      Width           =   285
   End
   Begin VB.TextBox txtaccno 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1695
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin MSComctlLib.ListView Lvwchartsofaccounts 
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   9975
      View            =   3
      SortOrder       =   -1  'True
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Baltic"
         Size            =   8.25
         Charset         =   186
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList ImgSorted 
      Left            =   4440
      Top             =   480
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
            Picture         =   "frmchartsofaccounts.frx":0704
            Key             =   "up"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmchartsofaccounts.frx":0C9E
            Key             =   "down"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Period Ending"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   495
      TabIndex        =   5
      Top             =   600
      Width           =   1170
   End
   Begin VB.Label lblname 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3660
      TabIndex        =   4
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Starting Account"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1425
   End
End
Attribute VB_Name = "frmchartsofaccounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdbudget_Click()
frmbudgetting.Show vbModal
End Sub

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdopen_Click()
On Error GoTo ErrorHandler
 Txtaccno.Text = Lvwchartsofaccounts.SelectedItem.Text
 Set rs = New ADODB.Recordset
 
 Set clsClass = New cdbase
    
    Provider = clsClass.OpenCon
    
    Set cn = CreateObject("adodb.connection")
    
    cn.Open Provider, "atm", "atm"
        sql = ""
   sql = "select * from glsetup where accno='" & Txtaccno & "'"
   Set rs = New ADODB.Recordset
   rs.Open sql, cn, adOpenKeyset, adLockOptimistic
   If Not rs.EOF Then
    If Not IsNull(rs.Fields("glaccname")) Then lblname = rs.Fields("glaccname")
    End If
        Lvwchartsofaccounts_DblClick
        Exit Sub
ErrorHandler:
        MsgBox err.description
End Sub

Private Sub cmdprintledgers_Click()
On Error Resume Next
    reportname = "glsetaccounts.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, CompanyName
End Sub

Private Sub Form_Load()
With Lvwchartsofaccounts
    
        .ListItems.Clear
        
        .ColumnHeaders.Clear
    
    End With

    Set rs = CreateObject("adodb.recordset")
    
    sql = "Select * from glsetup order by GLCODE,GLACCGROUP "
    
    Set rs = CreateObject("adodb.recordset")
    
    Set clsClass = New cdbase
    
    Provider = clsClass.OpenCon
    
    Set cn = CreateObject("adodb.connection")
    
    cn.Open Provider, "atm", "atm"
    
    rs.Open sql, cn
    Dim l As Integer
    
    With Lvwchartsofaccounts
        
        .ColumnHeaders.Add , , "Account Number"
        .ColumnHeaders.Add , , "Description", 3000
        .ColumnHeaders.Add , , "Type"
        .ColumnHeaders.Add , , "Normal Balance"
        .ColumnHeaders.Add , , "Account Group"
        .ColumnHeaders.Add , , "Status"
        .ColumnHeaders.Add , , "currency"
        .ColumnHeaders.Add , , "Is subledger"
          
        
        While Not rs.EOF
        
            Set li = .ListItems.Add(, , Trim(rs.Fields("accno")))
           If Not IsNull(rs.Fields("glaccName")) Then li.ListSubItems.Add , , Trim(rs.Fields("glaccName"))
           If Not IsNull(rs.Fields("glacctype")) Then li.ListSubItems.Add , , Trim(rs.Fields("glacctype"))
           If Not IsNull(rs.Fields("Normalbal")) Then li.ListSubItems.Add , , Trim(rs.Fields("Normalbal"))
           If Not IsNull(rs.Fields("Glaccgroup")) Then li.ListSubItems.Add , , Trim(rs.Fields("Glaccgroup"))
           'If Not IsNull(rs.Fields("status")) Then li.ListSubItems.Add , , Trim(rs.Fields("status"))
           If Not IsNull(rs.Fields("curr")) Then li.ListSubItems.Add , , Trim(rs.Fields("curr"))
           If rs.Fields("issubledger") = True Then
           l = 0
           Else
           l = 1
           End If
           If Not IsNull(rs.Fields("issubledger")) Then li.ListSubItems.Add , , l
            
            rs.MoveNext
        
        Wend
        
    End With
    
    rs.Close
    
    Set rs = Nothing
End Sub

Private Sub Lvwchartsofaccounts_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With Lvwchartsofaccounts
        .Sorted = True
        .SortKey = ColumnHeader.SubItemIndex
        If .SortOrder = lvwAscending Then
            .SortOrder = lvwDescending
        Else
            .SortOrder = lvwAscending
        End If
    End With
End Sub
Private Sub ClearHeaderIcons(CurrentHeader As Integer)
    Dim I As Integer
    For I = 1 To Lvwchartsofaccounts.ColumnHeaders.Count
        If Lvwchartsofaccounts.ColumnHeaders(I).Index <> CurrentHeader Then
            Lvwchartsofaccounts.ColumnHeaders(I).Icon = Empty
        End If
    Next
End Sub
Private Sub Lvwchartsofaccounts_DblClick()
On Error GoTo ErrorHandler
 If Lvwchartsofaccounts.ListItems.Count > 0 Then
    
        Txtaccno.Text = Lvwchartsofaccounts.SelectedItem.Text
        
        frmglsetup.Txtaccno = Txtaccno
        Set rs = New ADODB.Recordset
        sql = ""
   sql = "select * from glsetup where accno='" & Txtaccno & "'"
   Set rs = New ADODB.Recordset
   rs.Open sql, cn, adOpenKeyset, adLockOptimistic
   If Not rs.EOF Then
        
           If Not IsNull(rs.Fields("glaccname")) Then frmglsetup.txtAccName = rs.Fields("glaccname")
           If Not IsNull(rs.Fields("accno")) Then frmglsetup.Txtaccno = rs.Fields("accno")
           If Not IsNull(rs.Fields("glacctype")) Then frmglsetup.cboaccoounttype = rs.Fields("glacctype")
           If Not IsNull(rs.Fields("glaccgroup")) Then frmglsetup.cboaccountgroup = rs.Fields("glaccgroup")
           If Not IsNull(rs.Fields("normalbal")) Then frmglsetup.cbonormalbalance = rs.Fields("normalbal")
           If Not IsNull(rs.Fields("curr")) Then frmglsetup.cbocurrency = rs.Fields("curr")
           If Not IsNull(rs.Fields("acccategory")) Then frmglsetup.cboacccategory = rs.Fields("acccategory")
          
          
           If Not IsNull(rs.Fields("glaccname")) Then lblname = rs.Fields("glaccname")
          frmglsetup.Show vbModal
          
          Else
          MsgBox "The General Ledger Account Does not exist or was not opened properly", vbCritical
Exit Sub
          End If
         
'        txtamount.text = lsvDeduction.SelectedItem.ListSubItems(1).text
'        txtAmountInterest.text = lsvDeduction.SelectedItem.ListSubItems(2).text
'        txtDeductionCode.text = lsvDeduction.SelectedItem.ListSubItems(3).text
'        txtVoucherNo.text = lsvDeduction.SelectedItem.ListSubItems(4).text
'Else

        
        
    End If
    Exit Sub
ErrorHandler:
    MsgBox err.description
End Sub

Private Sub Picture5_Click()
    On Error GoTo SysError
    frmSearchGLAccounts.Show vbModal, Me
    If Continue Then
        If SearchValue <> "" Then
            Txtaccno = SearchValue
        End If
    End If
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub txtAccno_Change()
    On Error GoTo SysError
    Dim Accts As Account_Details
    Accts = Get_Account_Details(Txtaccno, "BOSA", ErrorMessage)
    If Accts.AccountNo <> "" Then
        lblname = Accts.AccountName
    Else
        lblname = ""
        If ErrorMessage <> "" Then
            MsgBox ErrorMessage, vbInformation, Me.Caption
            ErrorMessage = ""
        End If
    End If
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub txtAccNo_KeyPress(KeyAscii As Integer)
    KeyAscii = To_Upper_Case(KeyAscii)
End Sub
