VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmglinquiry 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "GENERAL LEDGER INQUIRY"
   ClientHeight    =   9945
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14340
   Icon            =   "frmglinquiry.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9945
   ScaleWidth      =   14340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Caption         =   "Print Data"
      Height          =   675
      Left            =   1725
      TabIndex        =   16
      Top             =   9225
      Visible         =   0   'False
      Width           =   8445
      Begin MSComCtl2.DTPicker dtpEnddate 
         Height          =   255
         Left            =   3960
         TabIndex        =   22
         Top             =   225
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   450
         _Version        =   393216
         Format          =   97583105
         CurrentDate     =   39568
      End
      Begin MSComCtl2.DTPicker dtpstartdate 
         Height          =   330
         Left            =   1275
         TabIndex        =   20
         Top             =   210
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   582
         _Version        =   393216
         Format          =   97583105
         CurrentDate     =   39568
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         Height          =   405
         Left            =   6765
         TabIndex        =   17
         Top             =   180
         Width           =   1290
      End
      Begin VB.Label Label18 
         Caption         =   "End Date"
         Height          =   315
         Left            =   3135
         TabIndex        =   21
         Top             =   210
         Width           =   1215
      End
      Begin VB.Label Label17 
         Caption         =   "Start Date"
         Height          =   375
         Left            =   285
         TabIndex        =   19
         Top             =   225
         Width           =   1425
      End
   End
   Begin VB.Frame fraTrans 
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Caption         =   "Transaction Details"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   480
      TabIndex        =   11
      Top             =   4560
      Visible         =   0   'False
      Width           =   10935
      Begin VB.CommandButton cmdTrans 
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7020
         TabIndex        =   12
         Top             =   2565
         Width           =   1080
      End
      Begin MSComctlLib.ListView lvwTrans 
         Height          =   2325
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   10740
         _ExtentX        =   18944
         _ExtentY        =   4101
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483624
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Account Name"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "AccNo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Debit Amount"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Credit Amount"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Description"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.Label lblCredit 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5445
         TabIndex        =   15
         Top             =   2430
         Width           =   1440
      End
      Begin VB.Label lblDebit 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3975
         TabIndex        =   14
         Top             =   2430
         Width           =   1440
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF80&
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14055
      Begin VB.CommandButton cmdSearch 
         Caption         =   "..."
         Height          =   285
         Left            =   1680
         TabIndex        =   24
         Top             =   240
         Width           =   300
      End
      Begin MSComCtl2.DTPicker dtpFromdate 
         Height          =   375
         Left            =   9720
         TabIndex        =   23
         Top             =   600
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         Format          =   97583105
         CurrentDate     =   40236
      End
      Begin VB.CommandButton cmdprintstatement 
         Caption         =   "Print Statement"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6120
         MaskColor       =   &H0080FF80&
         TabIndex        =   4
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox Txtaccno 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         MaxLength       =   15
         TabIndex        =   3
         Top             =   240
         Width           =   2055
      End
      Begin VB.ComboBox Cbodetail 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmglinquiry.frx":0442
         Left            =   120
         List            =   "frmglinquiry.frx":0449
         TabIndex        =   2
         Text            =   "Account Number"
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdrefresh 
         Caption         =   "Refresh"
         Height          =   315
         Left            =   4920
         TabIndex        =   1
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Opening Balance Date"
         Height          =   255
         Left            =   7920
         TabIndex        =   25
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label lblname 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   4080
         TabIndex        =   8
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label lblaccname 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   4080
         TabIndex        =   7
         Top             =   600
         Width           =   3735
      End
      Begin VB.Label Label6 
         Caption         =   "Current Balance"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label lblavail 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2040
         TabIndex        =   5
         Top             =   600
         Width           =   2055
      End
   End
   Begin MSComctlLib.ListView lvememtrans 
      Height          =   6630
      Left            =   0
      TabIndex        =   9
      ToolTipText     =   "Shows actual /available balances for the period specified"
      Top             =   1680
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   11695
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   12648384
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label Label16 
      Caption         =   "Label16"
      Height          =   495
      Left            =   6570
      TabIndex        =   18
      Top             =   4725
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000015&
      Caption         =   "Label9"
      Height          =   1815
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   11655
   End
End
Attribute VB_Name = "frmglinquiry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Dim LV As ListItem

Private Sub cmdprintstatement_Click()
'new transactions list
' STRFORMULA = "{d_LPO.Vendor}='" & rs.Fields(0) & "' and {d_Requisition.Status}='Ordered'"
STRFORMULA = "{GLTRANSACTIONS2.accno}='" & txtaccno & "'"
reportname = "specificledgers.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, title
End Sub

Private Sub cmdSearch_Click()
frmSearchGLAccounts.Show vbModal, Me
    If Continue Then
        If SearchValue <> "" Then
            txtaccno = SearchValue
            SearchValue = ""
        End If
    End If

End Sub

Private Sub Form_Load()
dtpFromdate = "30/12/2009"
With lvememtrans
        .ColumnHeaders.Clear
        .ListItems.Clear
    End With
    With lvememtrans
         .ColumnHeaders.Clear
         .ColumnHeaders.Add 1, , "Trans_Date"
         .ColumnHeaders.Add 2, , "V_No"
         .ColumnHeaders.Add 3, , "Dr", , lvwColumnRight
         .ColumnHeaders.Add 4, , "Cr", , lvwColumnRight
         .ColumnHeaders.Add 5, , "Avai_Bal", , lvwColumnRight
         .ColumnHeaders.Add 6, , "Description"
        
       
    End With
    
    lvememtrans.View = lvwReport

mysql = "delete  from GLTRANSACTIONS2"
oSaccoMaster.ExecuteThis (mysql)
 MousePointer = vbHourglass
 
 lblavail = getGlCurrentBalance(txtaccno)
 
' '// Get Opening Balances
'mysql = ""
'mysql = "Get_OpeningBalances '30/12/2009'"
'oSaccoMaster.ExecuteThis (mysql)
'
''//Get Non-Member Transactions
'mysql = ""
'mysql = "Get_Non_member_Transaction '30/12/2009','" & Format(Date, "dd/MM/yyyy") & "'"
'oSaccoMaster.ExecuteThis (mysql)
MousePointer = vbNormal
sql = ""
End Sub

Private Sub txtAccno_Change()
Dim AvailableBal As Currency
Dim Amount As Currency
sql = "SELECT     glaccname,mheader    FROM         GLSETUP  where accno='" & txtaccno & "'"
Set Rst = oSaccoMaster.GetRecordset(sql)
If Not Rst.EOF Then
If Not IsNull(Rst.Fields(0)) Then lblname = Rst.Fields(0)
If Not IsNull(Rst.Fields(1)) Then lblaccname = Rst.Fields(1)
End If
lvememtrans.ListItems.Clear
AvailableBal = 0
Dim ssql As String
    ssql = "SELECT     *   FROM         GLTRANSACTIONS2 where accno= '" & txtaccno & "' order by TRANSDATE ASC,id ASC"
    Set rs = cn.Execute(ssql)
    Do While Not rs.EOF

    With lvememtrans
      If rs!transdate <> "" Then
       Set LV = .ListItems.Add(, , rs!transdate)
                 If Not IsNull(rs!chequeno) Then
                  LV.ListSubItems.Add , , rs!chequeno
                Else
                  LV.ListSubItems.Add , , "Not*in*File"
                End If
                Amount = rs!Amount
                 If Check_if_asset(txtaccno) = True Then
        If rs!Amount <> "" Then
          If UCase(Trim(rs!transtype)) = "DR" Then
            LV.ListSubItems.Add , , Format(rs!Amount, "###,###,###.00")
            LV.ListSubItems.Add , , Format(0, "0.00")
            AvailableBal = AvailableBal + Amount
          Else
            LV.ListSubItems.Add , , Format(0, "0.00")
            LV.ListSubItems.Add , , Format(rs!Amount, "###,###,###.00")
            AvailableBal = AvailableBal - Amount
          End If
        Else
             rs!Amount = 0
        End If
        Else
        If rs!Amount <> "" Then
          If UCase(Trim(rs!transtype)) = "DR" Then
            LV.ListSubItems.Add , , Format(rs!Amount, "###,###,###.00")
            LV.ListSubItems.Add , , Format(0, "0.00")
            AvailableBal = AvailableBal - Amount
          Else
            LV.ListSubItems.Add , , Format(0, "0.00")
            LV.ListSubItems.Add , , Format(rs!Amount, "###,###,###.00")
            AvailableBal = AvailableBal + Amount
          End If
        Else
             rs!Amount = 0
        End If
        End If
        
        If Not IsNull(rs!available) Then
            If Check_if_asset(txtaccno) = True Then
             LV.ListSubItems.Add , , Format(Format((AvailableBal), "###,###,###.00"), "###,###,###.00")
            Else
            LV.ListSubItems.Add , , Format((AvailableBal), "###,###,###.00")
            End If
        Else
             LV.ListSubItems.Add , , "0.00"
        End If
   
         If rs!TransDescription <> "" Then
              LV.ListSubItems.Add , , rs!TransDescription
         Else
              LV.ListSubItems.Add , , "No Desc"
         End If
    
      
      End If
    End With


    rs.MoveNext
    Loop
lblavail = AvailableBal
End Sub
Private Sub Load_data()
'Dim RsRecords As New ADODB.Recordset
'Dim NormBal As String
'Dim OpeningBal As Double, Closingbalance As Double
'Dim available As Double
'    Set rs = oSaccoMaster.GetRecordset("set dateformat dmy Select * from GLSETUP where Accno='" & Trim(Txtaccno) & "' AND NewGLOpeningBalDate>='" & dtpFromdate & "'")
'    Closingbalance = 0
'    available = 0
'    If Not rs.EOF Then
'        NormBal = rs!NormalBal
'        transdate = rs!NewGLOpeningBalDate
'    End If
''Get Opening Balance
'mysql = "SET DateFormat DMY select amount as Amount,available as available,Accno, Transtype,month(TransDate) as TransMonth," _
'& " Year(TransDate) as TransYear from GLTRANSACTIONS2 WHERE Accno ='" & Trim(Txtaccno) & "' and transdescription='BAL B/F' order by month(TransDate) asc" ' and TransDate>='" & dtpFromdate & "' and TransDate<='" _
'& " and transdescription='BAL B/F' order by month(TransDate) desc"
'
'Set RsRecords = Nothing
'Set RsRecords = oSaccoMaster.GetRecordset(mysql)
'
'                           If Not RsRecords.EOF Then
'                         '  lsvTrans.ListItems.Clear
'                           Do While Not RsRecords.EOF
'
'                            Set li = lsvTrans.ListItems.Add(, , transdate) 'DateSerial(RsRecords!TransYear, RsRecords!TransMonth + 1, 1 - 1))
'
'                                li.ListSubItems.Add , , ""
'
'                                If RsRecords!transtype = "DR" Then
'                                    li.ListSubItems.Add , , Format(RsRecords!amount, "###,###,###.00")
'                                    li.ListSubItems.Add , , "0"
'                                 Closingbalance = Closingbalance + (IIf(IsNull(RsRecords!amount), 0, RsRecords!amount) * -1)
'
'                                         strValue = "Opening Balance"
'                                  available = RsRecords!amount * (-1)
'                                Else
'                                    li.ListSubItems.Add , , "0"
'                                    li.ListSubItems.Add , , Format(RsRecords!amount, "###,###,###.00")
'                                 Closingbalance = Closingbalance + IIf(IsNull(RsRecords!amount), 0, RsRecords!amount)
'
'                                strValue = "Opening Balance"
'                               available = RsRecords!amount
'                                End If
'
'                                li.ListSubItems.Add , , Format(available, "###,###,###,###.00")
'                                li.ListSubItems.Add , , strValue
'
'                           RsRecords.MoveNext
'                           Loop
'
'                           Else
'                           lsvTrans.ListItems.Clear
'
'                           End If
'
'
'
''Get Other Transactions
'    mysql = "SET DateFormat DMY select sum(amount) as Amount,sum(available) as available,Accno, Transtype,month(TransDate) as TransMonth," _
'            & " Year(TransDate) as TransYear from GLTRANSACTIONS2 WHERE Accno ='" & Trim(Txtaccno) & "' and TransDate>='" & dtpFromdate & "' and TransDate<='" _
'            & dtpTodate & "' and transdescription<>'BAL B/F' group by Accno,Transtype,month(TransDate),Year(TransDate) order by month(TransDate) asc"
'
'Set RsRecords = Nothing
'Set RsRecords = oSaccoMaster.GetRecordset(mysql)
'
'                           If Not RsRecords.EOF Then
'
'                           Do While Not RsRecords.EOF
'
'                            Set li = lsvTrans.ListItems.Add(, , DateSerial(RsRecords!TransYear, RsRecords!TransMonth + 1, 1 - 1))
'                                'li.ListSubItems.Add , , RsRecords!chequeno & ""
'                                li.ListSubItems.Add , , ""
'
'                                If RsRecords!transtype = "DR" Then
'                                    li.ListSubItems.Add , , Format(RsRecords!amount, "###,###,###.00")
'                                    li.ListSubItems.Add , , "0"
'                                 Closingbalance = Closingbalance + (IIf(IsNull(RsRecords!amount), 0, RsRecords!amount) * -1)
'                                 available = available + (IIf(IsNull(RsRecords!amount), 0, RsRecords!amount) * -1)
'
'                                strValue = "Receipts"
'                                Else
'                                    li.ListSubItems.Add , , "0"
'                                    li.ListSubItems.Add , , Format(RsRecords!amount, "###,###,###.00")
'                                 Closingbalance = Closingbalance + IIf(IsNull(RsRecords!amount), 0, RsRecords!amount)
'                                 available = available + IIf(IsNull(RsRecords!amount), 0, RsRecords!amount)
'                                        strValue = "Payments"
'
'                                End If
'
'                                li.ListSubItems.Add , , Format(available, "###,###,###,###.00")
'                                'li.ListSubItems.Add , , RsRecords!TransDescription & ""
'                                li.ListSubItems.Add , , MonthName(RsRecords!TransMonth) & " " & strValue
'
'                           RsRecords.MoveNext
'                           Loop
'                        lblCurrentbalance.Caption = Format(Closingbalance, "###,###,###.00")
'                           Else
'                          ' lsvTrans.ListItems.Clear
'
'                           End If
End Sub

Private Function Check_if_asset(Accno As String) As Boolean
    Dim RsRecords As New ADODB.Recordset
    
    mysql = ""
    mysql = "select glaccmaingroup from glsetup  where  accno ='" & Accno & "'"
    
    Set RsRecords = oSaccoMaster.GetRecordset(mysql)
    
    If Not RsRecords.EOF Then
        If (LCase(RsRecords!GLAccMainGroup) = LCase("Current Assets") Or LCase(RsRecords!GLAccMainGroup) = LCase("Fixed Assets")) Then
            Check_if_asset = True
        Else
            Check_if_asset = False
        End If
    Else
        Check_if_asset = False
    End If
End Function

