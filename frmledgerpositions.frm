VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmledgerpositions 
   Caption         =   "LEDGER BALANCES "
   ClientHeight    =   8235
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11160
   Icon            =   "frmledgerpositions.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   11160
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture4 
      Height          =   285
      Left            =   3480
      Picture         =   "frmledgerpositions.frx":0442
      ScaleHeight     =   225
      ScaleWidth      =   195
      TabIndex        =   3
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox txtaccno 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7335
      Left            =   0
      TabIndex        =   0
      Top             =   1320
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   12938
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Ledger Positions"
      TabPicture(0)   =   "frmledgerpositions.frx":0704
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label26"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbluncleared"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblbookbalance"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblavail"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblaccname"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblname"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Periodic Balances"
      TabPicture(1)   =   "frmledgerpositions.frx":0720
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(1)=   "lvememtrans"
      Tab(1).Control(2)=   "txttransdate"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "History"
      TabPicture(2)   =   "frmledgerpositions.frx":073C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin MSComCtl2.DTPicker txttransdate 
         Height          =   255
         Left            =   -73680
         TabIndex        =   12
         Top             =   480
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   450
         _Version        =   393216
         Format          =   130416641
         CurrentDate     =   39029
      End
      Begin MSComctlLib.ListView lvememtrans 
         Height          =   4935
         Left            =   -74880
         TabIndex        =   4
         Top             =   1440
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   8705
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         OLEDragMode     =   1
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         OLEDragMode     =   1
         NumItems        =   0
      End
      Begin VB.Label Label2 
         Caption         =   "From Date"
         Height          =   255
         Left            =   -74760
         TabIndex        =   11
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblname 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   3600
         TabIndex        =   10
         Top             =   960
         Width           =   3135
      End
      Begin VB.Label lblaccname 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   3600
         TabIndex        =   9
         Top             =   1320
         Width           =   3135
      End
      Begin VB.Label lblavail 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   720
         TabIndex        =   8
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label lblbookbalance 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   720
         TabIndex        =   7
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Label lbluncleared 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5040
         TabIndex        =   6
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label26 
         Caption         =   "Uncleared Effects"
         Height          =   255
         Left            =   3600
         TabIndex        =   5
         Top             =   1680
         Width           =   1455
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Account Number"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmledgerpositions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myclass As cdbase
Public Event CloseControl(bExit As Boolean)
Dim rsd As Object
Dim AccName As String
Dim custno As String
Dim desc As String
Dim lblacname As String
Dim withcharges As Currency
Dim totalcharges As Currency
Dim charge1 As Currency
Dim charge2 As Currency
Dim charge3 As Currency
Dim charge4 As Currency
Dim minBal As Currency
Dim AVAIL1 As Currency
Dim custno1 As String
Dim idno1 As String
Dim payno1 As String
Dim name1 As String
Dim teller As String
Dim accname1 As String
Dim AVAIL2 As Currency
Dim custno2 As String
Dim idno2 As String
Dim payno2 As String
Dim name2 As String
Dim accname2 As String
Dim glnamE As String 'FOR CONTRA
Dim glidno As String 'FOR CONTRA
Dim glmemno As String 'FOR CONTRA
Dim glpayno As String 'FOR CONTRA
Dim bookba As Currency
Dim bookba1 As Currency
Dim bookba2 As Currency
Dim bookba3 As Currency
Dim glcomm As String 'FOR CONTRA
Dim glaccno As String
Dim authorisecomm As Currency
Dim glnamecom As String 'FOR COMMISSION
Dim glcommemno As String 'FOR COMMISSION
Dim glcomidno As String 'FOR COMMISSION
Dim glcompayno As String 'FOR COMMISSION
Dim glcommission As String
Dim glnamestamp As String
Dim glidnostamp As String
Dim glpaynostamp As String
Dim glmemnostamp As String
Dim glnameteller As String
Dim glcombal As Currency
Dim gltellerbal As Currency
Dim glstampbal As Currency
Dim glcbocharge1accno As String
Dim glcbocharge1idno As String
Dim glcbocharge1memberno As String
Dim glcbocharge1payno As String
Dim glcbocharge1boobal As Currency
Dim glcbocharge1name As String
Dim glcbocharge2accno As String
Dim glcbocharge2idno As String
Dim glcbocharge2memberno As String
Dim glcbocharge2payno As String
Dim glcbocharge2boobal As Currency
Dim glcbocharge2name As String
Dim glcbocharge3accno As String
Dim glcbocharge3idno As String
Dim glcbocharge3memberno As String
Dim glcbocharge3payno As String
Dim glcbocharge3boobal As Currency
Dim glcbocharge3name As String
Dim glcbocharge4accno As String
Dim glcbocharge4idno As String
Dim glcbocharge4memberno As String
Dim glcbocharge4payno As String
Dim glcbocharge4boobal As Currency
Dim glcbocharge4name As String

Private Sub Picture4_Click()
Me.MousePointer = vbHourglass
         frmsearchacc.Show vbModal
        Txtaccno = sel
        txtAccNo_Validate True
        Me.MousePointer = 0
End Sub

Private Sub txtAccno_Change()
Dim myrec1 As Object
Dim rss As Object
Dim amt As Long
Dim rsCODE As Recordset
Set myclass = New cdbase
Set cn = CreateObject("adodb.connection")
Provider = myclass.OpenCon
cn.Open Provider, "atm", "atm"
Dim rsun As Recordset
Dim uncleared As Currency
'// check first if it has a control account on the accountcodes
Dim rsacccode As Recordset
sql = ""
sql = "SELECT     *  FROM         AccountCodes where accno='" & Txtaccno & "'"
Set rsacccode = New ADODB.Recordset
rsacccode.Open sql, cn, adOpenKeyset, adLockOptimistic
If Not rsacccode.EOF Then
'// if it exist then do the sum in the general ledger
'get_sum_savings stored procedures.
Dim rsr As Recordset
Set rsr = New ADODB.Recordset
sql = ""
sql = "select accno from cub where accno='" & Txtaccno & "'"
rsr.Open sql, cn, adOpenKeyset, adLockOptimistic
If Not rsr.EOF Then
'//get the balance now
Dim sssql As String
Dim tmount1 As Currency
Dim rspro As Recordset
Set rspro = New ADODB.Recordset
sssql = "get_sum_savings '" & Txtaccno & "' "
Set rspro = cn.Execute(sssql)
Dim tamount As Currency
tamount = Round_Of_Two_Decimal(rspro.Fields(0))
Dim ss As String
Dim rsp As Recordset
Set rsp = New ADODB.Recordset
ss = "get_sum_saving_avail '" & Txtaccno & "'"
Set rsp = cn.Execute(ss)
tmount1 = Round_Of_Two_Decimal(rsp.Fields(0))

End If

End If
'// GET TOTAL IF IS SUBLEDGER ACCOUNT
Dim rssubledger As New ADODB.Recordset
Dim Ramount As Currency

sql = ""
sql = "SELECT     *  FROM         cub where accno='" & Txtaccno & "' and  (hassubledgers = 1)"
Set rssubledger = New ADODB.Recordset
rssubledger.Open sql, cn, adOpenKeyset, adLockOptimistic

If Not rssubledger.EOF Then
'// total all the subledgers
If Not IsNull(rssubledger.Fields("availablebalance")) Then Ramount = rssubledger.Fields("availablebalance")

'// get the sum of the others
Dim rssum As Recordset
Dim t As Currency
Dim tota As Currency
sql = ""
sql = "SELECT     sum(availablebalance)as att  FROM         cub where main_accno='" & Txtaccno & "'"
Set rssum = New ADODB.Recordset
rssum.Open sql, cn, adOpenKeyset, adLockOptimistic
If Not rssum.EOF Then
If Not IsNull(rssum.Fields(0)) Then tota = rssum.Fields(0)
'// get the total and posted on the screen.
t = tota + Ramount
tamount = Round_Of_Two_Decimal(t)

End If

End If

'// check if there exist uncleared cheques
sql = "SELECT     SUM(Amount) AS unclearedamnt FROM         CustomerBalance  WHERE     (AccNO = '" & Txtaccno & "') AND (TransDescription LIKE 'Cheque Dep(uncleared)%')"
Set rsun = New ADODB.Recordset
rsun.Open sql, cn
If Not rsun.EOF Then
If Not IsNull(rsun.Fields("unclearedamnt")) Then lbluncleared = rsun.Fields("unclearedamnt") Else uncleared = 0
lbluncleared = Format(lbluncleared, "###,###,###.00")
End If


 Set rs = CreateObject("adodb.recordset")
    
    sql = "SELECT *  FROM CustomerBalance where accno='" & Txtaccno & _
     "' ORDER BY TRANSDATE,customerbalanceid ASC"
     rs.Open sql, cn
    rs.Close

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Set myrec1 = CreateObject("adodb.recordset")
    sql = "SELECT top 1 * from cub where accno='" & Txtaccno & "' "
     myrec1.Open sql, cn
     If myrec1.EOF Then
     lblname = ""
     lblaccname = ""
   
     lblavail = 0#
     lvememtrans.Visible = False
     'MsgBox "Check if Member  Exist OR Check if the account is valid?? ", vbInformation, "Transactional details"
     Exit Sub
     Else
        lvememtrans.Visible = True
        If Not IsNull(myrec1!name) Then lblname = myrec1!name
        If Not IsNull(myrec1!AccountName) Then lblaccname = myrec1!AccountName
        Set rsCODE = CreateObject("ADODB.Recordset")
    
         rsCODE.Open "SELECT * from AccountCodes WHERE AccountName='" & lblaccname & "'", cn
        
        If rsCODE.EOF Then
        MsgBox "Try eiditing the accounttype. The account type you have does not exist in our records ", vbCritical, "Transactions"
             'Exit Sub
        Else
            lblacname = rsCODE!AccountName
        
            minBal = rsCODE!Minimumbal
        End If
        
        'rebuild_accno Txtaccno
        Dim rsproc As Recordset
     Set rsproc = New ADODB.Recordset
     'Dim sssql As String
     sssql = "proc_rebuild '" & Txtaccno & "'"
     If tamount = 0 Then
        'If Not IsNull(myrec1!accNo) Then lblaccno = myrec1!accNo
        If Not IsNull(myrec1!availablebalance) Then lblavail = Format(myrec1!availablebalance, "#,###,###.00") Else lblavail = 0#
        If Not IsNull(myrec1!availablebalance) Then lblbookbalance = Format(myrec1!availablebalance - minBal, "#,###,###.00") Else lblbookbalance = 0#
  
         If lbluncleared = "" Then lbluncleared = 0
         lblavail = CCur(lblavail) + CCur(lbluncleared)
         lblavail = Format(lblavail, "###,###,###.00")
         Else
         lblavail = Format(tamount, "###,###,###.00")
         lblbookbalance = Format(tmount1, "###,###,###.00")
         End If
        'If Not IsNull(myrec1!memberno) Then lblgno = myrec1!memberno
     End If
     
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++


    Dim LV As ListItem

    lvememtrans.ListItems.Clear
    rs.Open

    Do While Not rs.EOF

    With lvememtrans
      If rs!transdate <> "" Then
      Set LV = .ListItems.Add(, , rs!transdate)
        If rs!transDescription <> "" Then
              LV.ListSubItems.Add , , rs!transDescription
         Else
              LV.ListSubItems.Add , , "No Desc"
         End If
        If rs!amount <> "" Then
          If UCase(Trim(rs!transtype)) = "DR" Then
            
            LV.ListSubItems.Add , , Format(rs!amount, "###,###,###.00")
            LV.ListSubItems.Add , , Format(0, "0.00")
           ' lvememtrans.ListItems.Add , , RS!amount = lvwColumnRight

          Else
          ' lvememtrans.ListItems.item(3).Left
            LV.ListSubItems.Add , , Format(0, "0.00")
            LV.ListSubItems.Add , , Format(rs!amount, "###,###,###.00")
           ' lvememtrans.ListItems.Add , , RS!amount = lvwColumnRight


          End If
        Else
             rs!amount = 0
        End If
        
        If Not IsNull(rs!availablebalance) Then
             LV.ListSubItems.Add , , Format(rs!availablebalance, "###,###,###.00")
        Else
             LV.ListSubItems.Add , , "0.00"
        End If
        If Not IsNull(rs!Commission) Then
            LV.ListSubItems.Add , , rs!Commission
        Else
           LV.ListSubItems.Add , , "0.00"
        End If
        
          If Not IsNull(rs!vno) Then
                  LV.ListSubItems.Add , , rs!vno
                Else
                  LV.ListSubItems.Add , , "DNN"
           End If
      LV.ListSubItems.Item(3).Bold = True
      
      End If
    End With


    rs.MoveNext
    Loop

    rs.Filter = 0
    rs.Close

End Sub

Private Sub txtAccNo_Validate(Cancel As Boolean)
Dim myrec1 As Object
Dim rss As Object
Dim amt As Currency
Set myclass = New cdbase
Set cn = CreateObject("adodb.connection")
Provider = myclass.OpenCon
cn.Open Provider, "atm", "atm"
Dim MYRE As Recordset
    Set MYRE = CreateObject("adodb.recordset")
    sql = "SELECT top 1 * from cub where accno='" & Txtaccno & "' "
     MYRE.Open sql, cn
     If Txtaccno <> "" Then
     If MYRE.EOF Then
      MsgBox "The account does not exist Please Seek assistance from the customer services", vbInformation, "Transactions"
     Exit Sub
     End If
     End If
End Sub
