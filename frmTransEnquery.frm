VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTransEnquery 
   Caption         =   " "
   ClientHeight    =   10050
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   14670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10050
   ScaleWidth      =   14670
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Unfreeze transporter"
      Height          =   375
      Left            =   3000
      TabIndex        =   32
      Top             =   9240
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Freeze transporter"
      Height          =   375
      Left            =   360
      TabIndex        =   31
      Top             =   9240
      Width           =   2055
   End
   Begin VB.TextBox txtcanno 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4080
      TabIndex        =   30
      Top             =   1080
      Width           =   2415
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3360
      TabIndex        =   16
      Top             =   0
      Width           =   4575
   End
   Begin VB.TextBox txtBank 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   720
      TabIndex        =   15
      Top             =   720
      Width           =   2175
   End
   Begin VB.TextBox txtTelNo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9240
      TabIndex        =   14
      Top             =   360
      Width           =   3015
   End
   Begin VB.TextBox txtBox 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3360
      TabIndex        =   13
      Top             =   360
      Width           =   4575
   End
   Begin VB.TextBox txtTransport 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9240
      TabIndex        =   12
      Top             =   0
      Width           =   3015
   End
   Begin VB.TextBox txtTCode 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   720
      TabIndex        =   11
      Top             =   0
      Width           =   1335
   End
   Begin VB.TextBox txtBBranch 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4080
      TabIndex        =   10
      Top             =   720
      Width           =   2175
   End
   Begin VB.TextBox txtAccNo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9240
      TabIndex        =   9
      Top             =   840
      Width           =   3015
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   0
      TabIndex        =   3
      Top             =   1440
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
         TabIndex        =   5
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   119930881
         CurrentDate     =   40157
      End
      Begin MSComCtl2.DTPicker dtpTo 
         Height          =   375
         Left            =   4080
         TabIndex        =   6
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   119930881
         CurrentDate     =   40157
      End
      Begin VB.Label Label9 
         Caption         =   "Date From"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Date To"
         Height          =   255
         Left            =   3120
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.PictureBox Picture5 
      Height          =   255
      Left            =   2160
      Picture         =   "frmTransEnquery.frx":0000
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   2
      Top             =   0
      Width           =   255
   End
   Begin VB.TextBox txtLocation 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   720
      TabIndex        =   1
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox TXTIDNO 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   720
      TabIndex        =   0
      Top             =   1080
      Width           =   2175
   End
   Begin MSComctlLib.ListView lvwEnguery 
      Height          =   6615
      Left            =   0
      TabIndex        =   17
      Top             =   2280
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description/SNo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "CR"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "DR"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Balance"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label20 
      Caption         =   "Can Number"
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
      Left            =   3000
      TabIndex        =   29
      Top             =   1080
      Width           =   1575
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
      Left            =   0
      TabIndex        =   28
      Top             =   720
      Width           =   1215
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
      Left            =   2520
      TabIndex        =   27
      Top             =   0
      Width           =   735
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
      Left            =   2640
      TabIndex        =   26
      Top             =   360
      Width           =   495
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
      Left            =   8040
      TabIndex        =   25
      Top             =   360
      Width           =   1215
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
      Left            =   8040
      TabIndex        =   24
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "Code"
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
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   615
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
      Left            =   3120
      TabIndex        =   22
      Top             =   720
      Width           =   855
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
      Left            =   7320
      TabIndex        =   21
      Top             =   840
      Width           =   1800
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
      Left            =   0
      TabIndex        =   20
      Top             =   360
      Width           =   615
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
      Left            =   8160
      TabIndex        =   19
      Top             =   1320
      Width           =   4095
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
      Left            =   0
      TabIndex        =   18
      Top             =   1080
      Width           =   495
   End
End
Attribute VB_Name = "frmTransEnquery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdShow_Click()
txtTCode_Validate True
End Sub

Private Sub Command1_Click()
'check the user
    sql = "SELECT     UserLoginIDs, SUPERUSER,username From UserAccounts where UserLoginIDs='" & User & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
    If rs!UserLoginIDs <> "JEFF" Then
'    optDelete.Visible = False
    MsgBox "You are not allowed to Freeze transporters", vbInformation
    Exit Sub
    End If
    End If
sql = ""
sql = "update d_Transporters set freezed='1' where transcode='" & txtTCode & "'"
oSaccoMaster.ExecuteThis (sql)
MsgBox "Transporter Freezed Succesfully"
End Sub

Private Sub Command2_Click()
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
sql = "update d_Transporters set freezed='0' where TransCode='" & txtTCode & "'"
oSaccoMaster.ExecuteThis (sql)
MsgBox "Transporter unFreezed Succesfully"
End Sub

Private Sub Form_Load()
dtpFrom = Format(Get_Server_Date, "dd/mm/yyyy")
dtpFrom = DateSerial(Year(dtpFrom), month(dtpFrom), 1)
dtpTo = DateSerial(Year(dtpFrom), month(dtpFrom) + 1, 1 - 1)
WindowState = vbMaximized
End Sub
Private Sub updatetransporter()
    '''UPDATE FOR THOSE HAVING TWO PRICES ON THE SAME MONTHE
        sql = "set dateformat dmy SELECT Sno  FROM d_Transport  WHERE Trans_Code='" & txtTCode & "' AND (Active = 0) AND DateInactivate BETWEEN '" & Format(dtpFrom, "dd/mm/yyyy") & "' AND '" & Format(dtpTo, "dd/mm/yyyy") & "' order by sno asc"
        Set rschecking = oSaccoMaster.GetRecordset(sql)
        While Not rschecking.EOF
        DoEvents
        Dim kgamount, qnt, qntamount, qntkgs As Double

        kgamount = 0
        qnt = 0
        qntamount = 0
        qntkgs = 0
            sql = "set dateformat dmy select Sno,Rate,startdate from d_Transport  WHERE Sno='" & rschecking.Fields(0) & "' and Trans_Code='" & txtTCode & "' AND (Active = 1) AND startdate BETWEEN '" & Format(dtpFrom, "dd/mm/yyyy") & "' AND '" & Format(dtpTo, "dd/mm/yyyy") & "'"
            Set rsconfirm = oSaccoMaster.GetRecordset(sql)
            If Not rsconfirm.EOF Then
              'sql = "SET DATEFORMAT DMY SELECT isnull(SUM(QSupplied),0) AS Kgs from d_Milkintake WHERE Sno='" & rsconfirm!sno & "' and TransDate BETWEEN '" & Format(rsconfirm!Startdate, "dd/mm/yyyy") & "' AND '" & Format(Enddate, "dd/mm/yyyy") & "'"
              Set rsp = oSaccoMaster.GetRecordset("d_sp_UpdateTkgs '" & rsconfirm!sno & "','" & Format(rsconfirm!Startdate, "dd/mm/yyyy") & "', '" & Format(dtpTo, "dd/mm/yyyy") & "'")
              If rsp!kgs > 0 Then
                 kgamount = rsp!kgs * rsconfirm!rate
                sql = "SET DATEFORMAT DMY SELECT QNTY, Amount from d_TransDetailed where Sno='" & rsconfirm!sno & "' and EndPeriod BETWEEN '" & Format(dtpFrom, "dd/mm/yyyy") & "' AND '" & Format(dtpTo, "dd/mm/yyyy") & "'"
                Set rsz = oSaccoMaster.GetRecordset(sql)
                If Not rsz.EOF Then
                   qnt = rsz!qnty + rsp!kgs
                   qntamount = rsz!amount + kgamount
                   qntkgs = 0
                  oSaccoMaster.ExecuteThis ("d_sp_UpdateT '" & qnt & "','" & qntamount & "','" & rsconfirm!sno & "','" & Format(dtpFrom, "dd/mm/yyyy") & "','" & Format(dtpTo, "dd/mm/yyyy") & "'")
                End If
              End If
            End If
        rschecking.MoveNext
        Wend
    '''END
End Sub
Private Sub LoadData()
Dim I As Long
Dim bal As Double
bal = 0
I = 0
lvwEnguery.ListItems.Clear
oSaccoMaster.ExecuteThis ("d_sp_TransUpdate '" & dtpFrom & "','" & dtpTo & "','" & User & "','" & txtTCode & "'")
updatetransporter
oSaccoMaster.ExecuteThis ("d_sp_UpdateTranstmpEnquery '" & txtTCode & "','" & dtpTo & "'")
oSaccoMaster.ExecuteThis ("d_sp_UpdateTranstmpEnqueryDed '" & txtTCode & "','" & dtpFrom & "','" & dtpTo & "'")

'Set rs = oSaccoMaster.GetRecordset("SELECT TransDate, SNo, CR, DR, Bal FROM d_tmpTransEnquery WHERE Code ='" & txtTCode & "' ORDER BY TransDate")
Set rs = oSaccoMaster.GetRecordset("SELECT TransDate, SNo, CR, DR, Bal FROM d_tmpTransEnquery WHERE Code ='" & txtTCode & "' ORDER BY CASE WHEN SNO LIKE '%[^0-9]%' THEN 9E99 ELSE CAST(SNO AS INTEGER) END, SNO")
'CASE IsNumeric(sno) WHEN 1 THEN Replicate(Char(0), 100 - Len(sno)) + sno ELSE sno END
'case when isnumeric(your_column) = 1 then your_column else 999999999 end,
'your_colum
'CASE WHEN col LIKE '%[^0-9]%' THEN 9E99 ELSE CAST(col AS INTEGER) END, col
With rs
While Not rs.EOF
DoEvents
   '//check if it is a flat rate case
   Dim rst As New ADODB.Recordset, rate As Double, samson As Integer, amount As Double, L3 As Double
   
   Set rst = oSaccoMaster.GetRecordset("SELECT     *  FROM         d_Transporters   WHERE     (transcode = '" & txtTCode & "') and isfrate=1")
   If rst.EOF Then
   Set li = lvwEnguery.ListItems.Add(, , IIf(IsNull(!transdate), "", !transdate))
   li.SubItems(1) = IIf(IsNull(!sno), "", !sno)
   li.SubItems(2) = IIf(IsNull(!cr), 0, !cr)
   li.SubItems(3) = IIf(IsNull(!dr), 0, !dr)
   bal = bal + li.SubItems(2) - li.SubItems(3)
   li.SubItems(4) = bal
   Else
    Set li = lvwEnguery.ListItems.Add(, , dtpFrom)
    rate = IIf(IsNull(rst.Fields("rate")), 1, rst.Fields("rate"))
    If I = 0 Then
    li.SubItems(1) = "Flat Rate"
     End If
   samson = Days_In_Month(month(dtpTo), month(dtpTo))
   amount = samson * rate
   If I = 0 Then
   li.SubItems(2) = amount
   End If
   li.SubItems(3) = IIf(IsNull(!dr), 0, !dr)
   L3 = li.SubItems(3)
   If I > 1 Then
   If L3 > 0 Then
   li.SubItems(1) = IIf(IsNull(!sno), "", !sno)
   bal = bal - li.SubItems(3)
   End If
   Else
   If I = 0 Then
   bal = bal + li.SubItems(2) - li.SubItems(3)
   End If
   End If
   li.SubItems(4) = bal
   End If
   I = I + 1
   .MoveNext

Wend
End With
lblNPay = "Net Pay :" & Format(bal, "#,##0.00")
End Sub

Private Sub Picture5_Click()
Me.MousePointer = vbHourglass
         frmSearchPTransporter.Show vbModal
        txtTCode = sel
        txtTCode_Validate True
        Me.MousePointer = 0
End Sub

Private Sub txtTCode_Validate(Cancel As Boolean)



Set rs = New ADODB.Recordset
sql = "d_sp_TransEnquiry  '" & txtTCode & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
' SELECT     TransName, CertNo, Locations, Phoneno, Address + ' ' + Town AS Expr1, Bcode, BBranch, Accno
'From dbo.d_Transporters
If Not IsNull(rs.Fields(0)) Then txtName = rs.Fields(0)
If Not IsNull(rs.Fields(1)) Then TXTIDNO = rs.Fields(1)
If Not IsNull(rs.Fields(2)) Then txtLocation = rs.Fields(2)
If Not IsNull(rs.Fields(3)) Then txtTelNo = rs.Fields(3)
If Not IsNull(rs.Fields(4)) Then txtBox = rs.Fields(4)
If Not IsNull(rs.Fields(5)) Then txtBank = rs.Fields(5)
If Not IsNull(rs.Fields(6)) Then txtBBranch = rs.Fields(6)
If Not IsNull(rs.Fields(7)) Then txtAccNo = rs.Fields(7)
If Not IsNull(rs.Fields(8)) Then txtcanno = rs.Fields(8)
End If

LoadData
End Sub
