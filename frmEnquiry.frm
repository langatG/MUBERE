VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmEnquery 
   Caption         =   "Farmers Details "
   ClientHeight    =   10515
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14355
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H0000FFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   10515
   ScaleWidth      =   14355
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdclose 
      Caption         =   "close"
      Height          =   375
      Left            =   11760
      TabIndex        =   53
      Top             =   3360
      Width           =   855
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1575
      Left            =   600
      TabIndex        =   52
      Top             =   3240
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   2778
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777088
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Quantity"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Amount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Balance"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdCarryfor 
      Caption         =   "Carryforword List"
      Height          =   615
      Left            =   12360
      TabIndex        =   51
      Top             =   9720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmddeductsta 
      Caption         =   "Suppliers Deductions List"
      Height          =   615
      Left            =   10200
      TabIndex        =   50
      Top             =   9720
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Statement"
      Height          =   615
      Left            =   8520
      TabIndex        =   49
      Top             =   9720
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Unfreeze supplier"
      Height          =   615
      Left            =   2160
      TabIndex        =   48
      Top             =   9720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Freeze supplier"
      Height          =   645
      Left            =   480
      TabIndex        =   47
      Top             =   9720
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SHARE STATEMENT"
      Height          =   615
      Left            =   6120
      TabIndex        =   46
      Top             =   9720
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox txtreg 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   12480
      TabIndex        =   44
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox TXTshares 
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   12600
      Locked          =   -1  'True
      TabIndex        =   42
      Top             =   1440
      Width           =   1695
   End
   Begin VB.TextBox txtcanno 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4680
      TabIndex        =   41
      Top             =   1320
      Width           =   2415
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "&Export To Excel"
      Height          =   585
      Left            =   4080
      TabIndex        =   39
      Top             =   9720
      Width           =   1980
   End
   Begin VB.TextBox TXTIDNO 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   840
      TabIndex        =   27
      Top             =   1200
      Width           =   2175
   End
   Begin VB.TextBox txtLocation 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   720
      TabIndex        =   24
      Top             =   480
      Width           =   1575
   End
   Begin VB.PictureBox Picture5 
      Height          =   255
      Left            =   1920
      Picture         =   "frmEnquiry.frx":0000
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   23
      Top             =   120
      Width           =   255
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   20
      Top             =   1560
      Width           =   7455
      Begin VB.CommandButton cmdShow 
         Caption         =   "Show"
         Default         =   -1  'True
         Height          =   375
         Left            =   6000
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   375
         Left            =   1320
         TabIndex        =   2
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   143261697
         CurrentDate     =   40157
      End
      Begin MSComCtl2.DTPicker dtpTo 
         Height          =   375
         Left            =   4080
         TabIndex        =   3
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   143261697
         CurrentDate     =   40157
      End
      Begin VB.Label Label10 
         Caption         =   "Date To"
         Height          =   255
         Left            =   3120
         TabIndex        =   22
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "Date From"
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.TextBox txtAccNo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9000
      TabIndex        =   19
      Top             =   960
      Width           =   3015
   End
   Begin VB.TextBox txtBBranch 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4200
      TabIndex        =   17
      Top             =   960
      Width           =   2175
   End
   Begin VB.TextBox txtSNo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtTransport 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9000
      TabIndex        =   14
      Top             =   120
      Width           =   3015
   End
   Begin VB.TextBox txtBox 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3120
      TabIndex        =   13
      Top             =   480
      Width           =   4575
   End
   Begin VB.TextBox txtTelNo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9000
      TabIndex        =   12
      Top             =   480
      Width           =   3015
   End
   Begin VB.TextBox txtBank 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   840
      TabIndex        =   11
      Top             =   840
      Width           =   2175
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3120
      TabIndex        =   10
      Top             =   120
      Width           =   4575
   End
   Begin MSComctlLib.ListView lvwEnguery 
      Height          =   6615
      Left            =   120
      TabIndex        =   0
      Top             =   2400
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   11668
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Milk Intake (Kgs)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "CR"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "DR"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Balance"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6720
      Top             =   9720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label22 
      Caption         =   "Registration Fee"
      Height          =   255
      Left            =   12480
      TabIndex        =   45
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label21 
      Caption         =   "TOTAL SHARES"
      Height          =   615
      Left            =   12480
      TabIndex        =   43
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label20 
      Caption         =   "Can Number"
      Height          =   255
      Left            =   3240
      TabIndex        =   40
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label19 
      Caption         =   "Deductions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7680
      TabIndex        =   38
      Top             =   9240
      Width           =   855
   End
   Begin VB.Label Label18 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8640
      TabIndex        =   37
      Top             =   9240
      Width           =   2055
   End
   Begin VB.Label Label17 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5280
      TabIndex        =   36
      Top             =   9240
      Width           =   1935
   End
   Begin VB.Label Label16 
      Caption         =   "Gross"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   35
      Top             =   9240
      Width           =   735
   End
   Begin VB.Label Label13 
      Caption         =   "Total Kgs"
      Height          =   375
      Left            =   240
      TabIndex        =   34
      Top             =   9240
      Width           =   1335
   End
   Begin VB.Label lblTKgs 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0.00"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1560
      TabIndex        =   33
      Top             =   9240
      Width           =   2415
   End
   Begin VB.Label Label15 
      Caption         =   "Deductions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   -3000
      TabIndex        =   32
      Top             =   10080
      Width           =   855
   End
   Begin VB.Label lblDeductions 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   -2040
      TabIndex        =   31
      Top             =   10080
      Width           =   2055
   End
   Begin VB.Label lblGross 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   -5400
      TabIndex        =   30
      Top             =   10080
      Width           =   1935
   End
   Begin VB.Label Label14 
      Caption         =   "Gross"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   -6240
      TabIndex        =   29
      Top             =   10080
      Width           =   735
   End
   Begin VB.Label Label12 
      Caption         =   "ID:"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label lblNPay 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   615
      Left            =   7920
      TabIndex        =   26
      Top             =   1440
      Width           =   4095
   End
   Begin VB.Label Label11 
      Caption         =   "Loc :"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Account Number :"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7080
      TabIndex        =   18
      Top             =   960
      Width           =   1800
   End
   Begin VB.Label Label7 
      Caption         =   "Branch :"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   16
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "M/No :"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "Transport :"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      TabIndex        =   9
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Telephone :"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      TabIndex        =   8
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Box :"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   7
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Name :"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   6
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Bank :"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "frmEnquery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCarryfor_Click()
reportname = "carryforward.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub cmdclose_Click()
ListView1.Visible = False
cmdclose.Visible = False
End Sub

Private Sub cmddeductsta_Click()
reportname = "SUPPLIERS DEDUCTIONS.rpt"
Show_Sales_Crystal_Report "", reportname, ""
End Sub

Private Sub cmdExport_Click()
  On Error GoTo SsyError
  Dim sno As String
  Dim NAMES As String
  sno = txtSNo
  NAMES = txtname
    Dim MyFso As New FileSystemObject, strData As String, MFile As TextStream, _
    FileName As String, I As Long, li As ListItem
    If lvwEnguery.ListItems.Count > 0 Then
        With CommonDialog1
            .Filter = "Comma Seperated Values|*.csv"
            .FileName = "STATEMENT"
            .ShowSave
            
            If .FileName <> "" Then
                FileName = .FileName
            End If
            .FileName = ""
        End With
        Set MFile = MyFso.OpenTextFile(FileName, ForWriting, True)
        strData = "Supplier No  :" & sno
        MFile.WriteLine strData
        strData = "Supplier Name:" & NAMES
        MFile.WriteLine strData
        strData = "---------------------------------------------"
        MFile.WriteLine strData
        strData = ""
        'strData = "Period for" - "& dtpFrom &" & "to" & "& dtpto &"
        strData = "Transdate    ,Description         ,Intake     ,DEBIT,CREDIT,Balance"
        MFile.WriteLine strData
        strData = ""
        For I = 1 To lvwEnguery.ListItems.Count
            Set li = lvwEnguery.ListItems(I)
            strData = li & "," & li.SubItems(1) & "," & CDbl(li.SubItems(2)) & "," & CDbl(li.SubItems(3)) _
            & "," & (li.SubItems(4)) & "," & li.SubItems(5)
            MFile.WriteLine strData
            strData = ""
        Next I
    Else
        MsgBox "There are no records to be exported", vbInformation, Me.Caption
    End If
    MsgBox "Items Successfully Imported Into CSV file", vbOKOnly
    Exit Sub
SsyError:
    MsgBox err.description, vbInformation, Me.Caption

End Sub

Private Sub cmdShow_Click()
txtSNo_Validate True
End Sub

Private Sub Command1_Click()
frmLedgerFees.Show vbModal
End Sub

Private Sub Command2_Click()
'check the user
    sql = "SELECT     UserLoginIDs, SUPERUSER,username From UserAccounts where UserLoginIDs='" & User & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
    If rs!UserLoginIDs <> "JEFF" Then
'    optDelete.Visible = False
    MsgBox "You are not allowed to Freeze Suppliers", vbInformation
    Exit Sub
    End If
    End If


sql = ""
sql = "update d_suppliers set freezed='1' where sno='" & txtSNo & "'"
oSaccoMaster.ExecuteThis (sql)
MsgBox "Farmer Freezed Succesfully"
End Sub

Private Sub Command3_Click()
'check the user
    sql = "SELECT     UserLoginIDs, SUPERUSER,username From UserAccounts where UserLoginIDs='" & User & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
    If rs!UserLoginIDs <> "JEFF" Then
'    optDelete.Visible = False
    MsgBox "You are not allowed to UnFreeze Suppliers", vbInformation
    Exit Sub
    End If
    End If


sql = ""
sql = "update d_suppliers set freezed='0' where sno='" & txtSNo & "'"
oSaccoMaster.ExecuteThis (sql)
MsgBox "Farmer unFreezed Succesfully"
End Sub

Private Sub Command4_Click()
frmSupplierStmt.Show vbModal
End Sub

Private Sub Form_Activate()
txtSNo.SetFocus
End Sub

Private Sub Form_Load()
DTPfrom = Format(Get_Server_Date, "dd/mm/yyyy")
DTPfrom = DateSerial(year(DTPfrom), month(DTPfrom), 1)
DTPto = DateSerial(year(DTPfrom), month(DTPfrom) + 1, 1 - 1)
ListView1.Visible = False
cmdclose.Visible = False
 
End Sub

Private Sub lvwEnguery_DblClick()
'On Error GoTo SsyError
'Dim bal As Double
'Dim descre, vaf, daf As String
''''check the decription of deduction IF its agrovet
'descre = lvwEnguery.SelectedItem.SubItems(1)
'    If descre <> "Agrovet" Then
'    Set rsoutle = oSaccoMaster.GetRecordset("set dateformat dmy SELECT * From d_supplier_deduc WHERE SNo=" & txtSNo & " and Description='Agrovet' and (Remarks like 'C/F%' or Remarks = 'Brought Forward') and Date_Deduc ='" & lvwEnguery.SelectedItem & "' ")
'    If rsoutle.EOF Then
'     Exit Sub
'     End If
'    End If
''IsNumeric
'sql = ""
'Set rsoutle = oSaccoMaster.GetRecordset("set dateformat dmy SELECT * From d_supplier_deduc WHERE SNo=" & txtSNo & " and Description='Agrovet' and (Remarks like 'C/F%' or Remarks = 'Brought Forward') and Date_Deduc >='" & dtpFrom & "' and Date_Deduc<='" & dtpTo & "' ")
'If Not rsoutle.EOF Then
'    ListView1.ListItems.Clear
'    vaf = "Agrovet Brought Forward"
'    daf = "Agrovet C/F"
'
'    bal = 0
'    ListView1.Visible = True
'    cmdclose.Visible = True
'
'If descre = vaf Then
'    Startdate = DateSerial(year(dtpFrom), month(dtpFrom) - 1, 1)
'    Enddate = DateSerial(year(Startdate), month(Startdate) + 1, 1 - 1)
'    Set rs = oSaccoMaster.GetRecordset("SELECT T_Date,Remarks,Qua, Amount From ag_Receipts WHERE S_No=" & txtSNo & " and T_Date>='" & Startdate & "' and T_Date<='" & Enddate & "' ORDER BY T_Date")
'    With rs
'    While Not rs.EOF
'
'       Set li = ListView1.ListItems.Add(, , IIf(IsNull(!T_Date), "", !T_Date))
'       li.SubItems(1) = IIf(IsNull(!Remarks), "", !Remarks)
'       li.SubItems(2) = IIf(IsNull(!Qua), "", !Qua)
'       li.SubItems(3) = IIf(IsNull(!amount), 0, !amount)
'       bal = Format(bal + li.SubItems(3), "#,##0.00")
'       li.SubItems(4) = bal
'       .MoveNext
'
'    Wend
'    End With
'        bal = li.SubItems(4)
'       Set li = ListView1.ListItems.Add(, , Enddate)
'       li.SubItems(1) = "Amount Deducted"
'       li.SubItems(2) = lvwEnguery.SelectedItem.SubItems(2)
'       'li.SubItems(3) = lvwEnguery.SelectedItem.SubItems(4)
'       bal = Format(bal - lvwEnguery.SelectedItem.SubItems(4), "#,##0.00")
'       li.SubItems(3) = bal
'       li.SubItems(4) = lvwEnguery.SelectedItem.SubItems(4)
'
'       ''''''calc carry forward
'       Set li = ListView1.ListItems.Add(, , lvwEnguery.SelectedItem)
'       li.SubItems(1) = lvwEnguery.SelectedItem.SubItems(1)
'       li.SubItems(2) = lvwEnguery.SelectedItem.SubItems(2)
'       li.SubItems(3) = lvwEnguery.SelectedItem.SubItems(4)
'       'Bal = Format(Bal - lvwEnguery.SelectedItem.SubItems(4), "#,##0.00")
'       'li.SubItems(3) = Bal
'       li.SubItems(4) = lvwEnguery.SelectedItem.SubItems(2)
'  Else
'      Startdate = DateSerial(year(dtpFrom), month(dtpFrom), 1)
'      Enddate = DateSerial(year(Startdate), month(Startdate) + 1, 1 - 1)
'        Set rs = oSaccoMaster.GetRecordset("SELECT T_Date,Remarks,Qua, Amount From ag_Receipts WHERE S_No=" & txtSNo & " and T_Date>='" & Startdate & "' and T_Date<='" & Enddate & "' ORDER BY T_Date")
'        With rs
'        While Not rs.EOF
'
'           Set li = ListView1.ListItems.Add(, , IIf(IsNull(!T_Date), "", !T_Date))
'           li.SubItems(1) = IIf(IsNull(!Remarks), "", !Remarks)
'           li.SubItems(2) = IIf(IsNull(!Qua), "", !Qua)
'           li.SubItems(3) = IIf(IsNull(!amount), 0, !amount)
'           bal = Format(bal + li.SubItems(3), "#,##0.00")
'           li.SubItems(4) = bal
'           .MoveNext
'        Wend
'        End With
'            bal = li.SubItems(4)
'           Set li = ListView1.ListItems.Add(, , Enddate)
'           li.SubItems(1) = "Amount Recovered"
'           li.SubItems(2) = lvwEnguery.SelectedItem.SubItems(2)
'           'li.SubItems(3) = lvwEnguery.SelectedItem.SubItems(4)
'              'Dim I As Integer
'              Dim ass As String
'              Dim sString As String
'              sString = lvwEnguery.SelectedItem.SubItems(1)
'                For I = 1 To Len(sString)
'                    If Mid(sString, I, 1) Like "[0-9]" Then
'                        ass = ass + (Mid(sString, I, 1))
'                    End If
'                Next I
'           bal = Format(bal - CCur(ass), "#,##0.00")
'           li.SubItems(3) = bal
'           li.SubItems(4) = ass
'
'           ''''''calc carry forward
'           Set li = ListView1.ListItems.Add(, , lvwEnguery.SelectedItem)
'           li.SubItems(1) = lvwEnguery.SelectedItem.SubItems(1)
'           li.SubItems(2) = lvwEnguery.SelectedItem.SubItems(2)
'           li.SubItems(3) = ass
'           'Bal = Format(Bal - lvwEnguery.SelectedItem.SubItems(4), "#,##0.00")
'           'li.SubItems(3) = Bal
'           li.SubItems(4) = lvwEnguery.SelectedItem.SubItems(2)
'    End If
'End If
'Exit Sub
'SsyError:
'    MsgBox err.description, vbInformation, Me.Caption
End Sub
Private Sub Picture5_Click()
    
Me.MousePointer = vbHourglass
        frmSearchSupplier.Show vbModal
        txtSNo = sel
        txtSNo_Validate True
        Me.MousePointer = 0
End Sub

Private Sub txtSNo_KeyPress(KeyAscii As Integer)
If (KeyAscii = 8) Or (KeyAscii = 48) Or (KeyAscii = 49) Or (KeyAscii = 50) Or (KeyAscii = 51) Or (KeyAscii = 52) Or (KeyAscii = 53) Or (KeyAscii = 54) Or (KeyAscii = 55) Or (KeyAscii = 56) Or (KeyAscii = 57) Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
MsgBox "Please enter a number "
End If
End Sub

Private Sub txtSNo_Validate(Cancel As Boolean)
On Error GoTo errmsg
Dim a, t As Boolean
If Trim(txtSNo) = "" Then
            txtSNo.SetFocus
        Exit Sub
    End If

txtname = ""
Txtaccno = ""
txtBank = ""
txtbbranch = ""
txtBox = ""
txtidno = ""
txtLocation = ""
txtTelNo = ""
txtTransport = ""
lvwEnguery.ListItems.Clear
lblNPay = "0.00"
    
Set rs = New ADODB.Recordset
sql = "d_sp_SupplierEnquiry " & txtSNo & ""
Set rs = oSaccoMaster.GetRecordset(sql)
If rs.EOF Then
    MsgBox "There is no supplier with number " & txtSNo
    Exit Sub
End If
If Not rs.EOF Then
' [Names], AccNo, Bcode, BBranch, Location, PhoneNo, Address + ' ' + Town AS ADDRESS
'FROM         d_Suppliers WHERE SNo= @SNo
If Not IsNull(rs.Fields(0)) Then txtname = rs.Fields(0)
If Not IsNull(rs.Fields(1)) Then Txtaccno = rs.Fields(1)
If Not IsNull(rs.Fields(2)) Then txtBank = rs.Fields(2)
If Not IsNull(rs.Fields(3)) Then txtbbranch = rs.Fields(3)
If Not IsNull(rs.Fields(4)) Then txtLocation = rs.Fields(4)
If Not IsNull(rs.Fields(5)) Then txtTelNo = rs.Fields(5)
If Not IsNull(rs.Fields(6)) Then txtBox = rs.Fields(6)
If Not IsNull(rs.Fields(7)) Then txtidno = rs.Fields(7)
If Not IsNull(rs.Fields(8)) Then txtcanno = rs.Fields(8)


 
Set rs = New ADODB.Recordset
sql = "d_sp_TransName " & txtSNo & ""
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
If Not IsNull(rs.Fields(0)) Then txtTransport = rs.Fields(0)
End If
'Shares
Dim rstsr As New Recordset
Dim rsts As New Recordset
Dim rss As New Recordset
Dim shareamt As Double
Set rsts = oSaccoMaster.GetRecordset("SELECT    SUM(Amount) AS amtt From d_sconribution WHERE     (transdescription LIKE '%shares%') AND (SNo = '" & txtSNo & "')")
If Not rsts.EOF Then
shareamt = IIf(IsNull(rsts!amtt), 0, rsts!amtt)
End If
Set rss = oSaccoMaster.GetRecordset("SELECT    SUM(Amount) AS amt From d_supplier_deduc WHERE     (Description LIKE '%shares%') AND (SNo = '" & txtSNo & "')")
If Not rss.EOF Then
TXTshares = IIf(IsNull(rss!amt), 0, rss!amt) + shareamt
End If
Set rstsr = oSaccoMaster.GetRecordset("SELECT    Amount From d_Registration WHERE     (transdescription LIKE '%reg%') AND (SNo = '" & txtSNo & "')")
If Not rstsr.EOF Then
txtreg = IIf(IsNull(rstsr!amount), 0, rstsr!amount)
End If
LoadData

  'SNo
End If
Exit Sub
errmsg:
MsgBox txtname & " did not supply milk between " & DTPfrom & " and " & DTPto

End Sub
Private Sub LoadData()
Dim bal As Double, rss As New Recordset, amt As Double, rsts As New Recordset, rstsr As New Recordset, shareamt As Double
Dim mid As Date
mid = DateSerial(year(DTPto), month(DTPto), Day(DTPto) + 1)
Set rs = oSaccoMaster.GetRecordset("d_sp_SupNet " & txtSNo & ",'" & DTPfrom & "','" & mid & "', 0")

If Not IsNull(rs.Fields(0)) Then
lblTKgs = rs.Fields(0)
Else
lblTKgs = "0.00"
End If

If Not IsNull(rs.Fields(1)) Then
Label17 = rs.Fields(1)
Else
Label17 = "0.00"
End If

Set rs = oSaccoMaster.GetRecordset("d_sp_SupNet " & txtSNo & ",'" & DTPfrom & "','" & mid & "', 1")
If rs.RecordCount > 0 Then
If Not IsNull(rs.Fields(0)) Then
Label18 = rs.Fields(0)
Else
Label18 = "0.00"
End If
End If

bal = 0
lvwEnguery.ListItems.Clear
oSaccoMaster.ExecuteThis ("DELETE FROM d_tmpEnquery")

oSaccoMaster.ExecuteThis ("d_sp_UpdatetmpEnquery " & txtSNo & ",'" & DTPfrom & "','" & mid & "'")
oSaccoMaster.ExecuteThis ("d_sp_UpdatetmpEnqueryDed " & txtSNo & ",'" & DTPfrom & "','" & mid & "'")
Dim Descrption

Set rs = oSaccoMaster.GetRecordset("SELECT TransDate, Description,Intake,CR,DR From d_tmpEnquery WHERE SNo=" & txtSNo & " ORDER BY TransDate")
With rs
While Not rs.EOF
   Descrption = !description
   If Trim(!description) = "HShares" Then
   Descrption = "Shares"
   End If
   
   If Trim(!description) = "TMShares" Then
   Descrption = "Registration"
   End If
   
   Set li = lvwEnguery.ListItems.Add(, , IIf(IsNull(!transdate), "", !transdate))
   li.SubItems(1) = IIf(IsNull(!description), "", Descrption)
   li.SubItems(2) = IIf(IsNull(!Intake), "", !Intake)
   li.SubItems(3) = IIf(IsNull(!cr), 0, !cr)
   li.SubItems(4) = IIf(IsNull(!dr), 0, !dr)
   bal = Format(bal + li.SubItems(3) - li.SubItems(4), "#,##0.00")
   li.SubItems(5) = bal
   .MoveNext

Wend
End With
lblNPay = "Net Pay :" & Format(bal, "#,##0.00")
'Set rsts = oSaccoMaster.GetRecordset("SELECT    SUM(Amount) AS amtt From d_sconribution WHERE     (transdescription LIKE '%shares%') AND (SNo = '" & txtSNo & "')")
'If Not rsts.EOF Then
'shareamt = IIf(IsNull(rsts!amtt), 0, rsts!amtt)
'End If
'Set rss = oSaccoMaster.GetRecordset("SELECT    SUM(Amount) AS amt From d_supplier_deduc WHERE     (Description LIKE '%shares%') AND (SNo = '" & txtSNo & "')")
'If Not rss.EOF Then
'TXTshares = IIf(IsNull(rss!amt), 0, rss!amt) + shareamt
'End If
'Set rstsr = oSaccoMaster.GetRecordset("SELECT    Amount From d_Registration WHERE     (transdescription LIKE '%reg%') AND (SNo = '" & txtSNo & "')")
'If Not rstsr.EOF Then
'txtreg = IIf(IsNull(rstsr!amount), 0, rstsr!amount)
'End If
End Sub
