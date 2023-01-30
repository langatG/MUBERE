VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmassetsinquiry 
   BackColor       =   &H00C0C000&
   Caption         =   "AC-Assets Inquiry"
   ClientHeight    =   9300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8985
   Icon            =   "frmassetsinquiry.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   9300
   ScaleWidth      =   8985
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Print Register"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   18
      Top             =   8760
      Width           =   1695
   End
   Begin VB.CommandButton cmdprint 
      Caption         =   " Depreciation Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7680
      TabIndex        =   17
      Top             =   8520
      Width           =   1215
   End
   Begin VB.CommandButton cmdinqure 
      Caption         =   "Assets Registration"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   8400
      Width           =   2055
   End
   Begin VB.CommandButton cmddepre 
      Caption         =   "Process Depreciation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   5640
      TabIndex        =   8
      Top             =   7920
      Width           =   2055
   End
   Begin VB.CommandButton cmdload 
      Caption         =   "Load"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2400
      TabIndex        =   7
      Top             =   8400
      Width           =   855
   End
   Begin VB.CommandButton cmdfinder 
      Height          =   285
      Left            =   3720
      Picture         =   "frmassetsinquiry.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Add New record"
      Top             =   240
      Width           =   375
   End
   Begin VB.TextBox txtSERIALNO 
      Appearance      =   0  'Flat
      DataField       =   "assetserialno"
      Height          =   285
      Index           =   1
      Left            =   1980
      TabIndex        =   5
      Top             =   555
      Width           =   3975
   End
   Begin VB.TextBox txtASSETSNAME 
      Appearance      =   0  'Flat
      DataField       =   "assetsname"
      Height          =   285
      Index           =   1
      Left            =   1980
      TabIndex        =   1
      Top             =   870
      Width           =   3975
   End
   Begin VB.TextBox txtASSETSNO 
      DataField       =   "assetsno"
      Height          =   285
      Index           =   1
      Left            =   1980
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   4440
      TabIndex        =   9
      Top             =   8520
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "yyyy"
      Format          =   143982595
      CurrentDate     =   40748
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   3480
      TabIndex        =   10
      Top             =   8520
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "MMM"
      Format          =   143982595
      CurrentDate     =   40748
   End
   Begin MSComctlLib.ListView lvwasset 
      Height          =   6375
      Left            =   120
      TabIndex        =   13
      Top             =   1320
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   11245
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Asset Code"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Asset Name"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Depreciation Value"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "GL Account No"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "DeprecAcc"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "PDate"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Purchase Amount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Current Value"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "SerialNo"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   7920
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      Caption         =   "Total Current Value"
      Height          =   375
      Left            =   6840
      TabIndex        =   16
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label lbltotal 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   6720
      TabIndex        =   15
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Month"
      Height          =   255
      Left            =   3480
      TabIndex        =   12
      Top             =   8280
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Year"
      Height          =   255
      Left            =   4440
      TabIndex        =   11
      Top             =   8280
      Width           =   735
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Asset Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   390
      TabIndex        =   4
      Top             =   900
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "RegNo/Serial No.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   390
      TabIndex        =   3
      Top             =   570
      Width           =   1590
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Asset No:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   390
      TabIndex        =   2
      Top             =   240
      Width           =   840
   End
End
Attribute VB_Name = "frmassetsinquiry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ContraAcc As String
Private Sub cmddepre_Click()
Dim rsdepv As New Recordset, depam As Double, deprate As Double, rsCredits As New Recordset, findate As Date, asv As Double, rsDebits As New Recordset, ACCNO As String, Startdate As Date, Debit As Double, Credit As Double
Dim mMonth As Double, yYear As Double

    Startdate = DateSerial(Year(DTPicker2), month(DTPicker1), 1)
    findate = DateSerial(Year(DTPicker2), month(DTPicker1) + 1, 1 - 1)
    
    sql = ""
sql = "select top 1 * from depreciation order by yyear desc,mmonth desc"
Set rsdepv = oSaccoMaster.GetRecordset(sql)
If Not rsdepv.EOF Then
  mMonth = rsdepv.Fields("mmonth")
  yYear = rsdepv.Fields("yyear")
    If Year(DTPicker2) = yYear And month(DTPicker1) > mMonth + 1 Then
      MsgBox "Process  Depreciation for -" & MonthName(mMonth + 1) & " first", vbInformation
    Exit Sub
  End If
End If

sql = ""
sql = "select  * from depreciation where mmonth=" & month(findate) & " and yyear=" & Year(findate) & ""
Set rsdepv = oSaccoMaster.GetRecordset(sql)
If Not rsdepv.EOF Then
    MsgBox "Depreciation for the period you choose has been processed", vbInformation
Else
    Dim ProcessDepre As New ADODB.Connection
    Set ProcessDepre = New ADODB.Connection
    ProcessDepre.Open "MAZIWA"
    ProcessDepre.BeginTrans
    
    On Error GoTo TransError
       
    Startdate = DateSerial(Year(DTPicker2), month(DTPicker1), 1)
    findate = DateSerial(Year(DTPicker2), month(DTPicker1) + 1, 1 - 1)
  
     
    If Date < findate Then MsgBox "You have not reached End month", vbInformation: Exit Sub
    
    If lvwasset.ListItems.Count > 0 Then
        ProgressBar1.Visible = True
        ProgressBar1.max = lvwasset.ListItems.Count
    Else
        MsgBox "Please Load Any Unposted Products", vbInformation, Me.Caption
        Exit Sub
    End If
    
    If MsgBox("Do You want to want to Process  Depreciation?", vbQuestion + vbYesNo, _
       Me.Caption) = vbNo Then
        Exit Sub
    End If
    I = 0
    
    'GetTransactionNo
    For I = 1 To lvwasset.ListItems.Count
        ACCNO = lvwasset.ListItems(I).SubItems(3)
        asv = lvwasset.ListItems(I).SubItems(7)
        ContraAcc = lvwasset.ListItems(I).SubItems(4)
        ProgressBar1.value = I
        ' ****** This Would Have worked well if every Fixed Asset had its own AccNo ****** '''
        ' ****** Therefore we base on  every Fixed Asset Current Value ****** '''

        If asv > 0 Then
         deprate = lvwasset.ListItems(I).SubItems(2)
         depam = (asv * (deprate / 100 / 12))
         depam = asv * lvwasset.ListItems(I).SubItems(2) / (12 * 100)
         depam = Round(depam, 2)
        End If
        

        If depam > 0 Then
              ' NewTransaction depam, findate, "Fixed Assets Depreciation "
              NewTransaction depam, findate, "Fixed Assets Depreciation"
                
                sql = ""
                sql = "INSERT INTO Depreciation  (AssetCode, mmonth, yyear, DepreciationAmt, uuser)"
                sql = sql & " VALUES     ('" & lvwasset.ListItems(I).Text & "', " & month(findate) & ", " & Year(findate) & ", " & depam & ", '" & User & "')"
                oSaccoMaster.Execute (sql)
                
                '*** Update Fixed Asset current value ***********
                sql = "Update assets_register Set CurrentValue= CurrentValue-" & depam & "  where AssetCode='" & lvwasset.ListItems(I).Text & "'"
                oSaccoMaster.Execute (sql)
   
            If Not Save_GLTRANSACTION(findate, CDbl(depam), ContraAcc, ACCNO, "Depreciation", ACCNO, User, "", "Fixed Asset Depreciation", 1, 1, lvwasset.ListItems(I).Text, transactionNo, "Maziwa") Then
                GoTo TransError
            End If
            
        End If
        depam = 0
        Debit = 0
        Credit = 0
        ContraAcc = ""
    Next I
    
    ProcessDepre.CommitTrans
    
    MsgBox "Depreciation processing  for the period Complete", vbInformation
    Form_Load
End If
    Exit Sub
TransError:
    feesTrans.RollbackTrans
    MsgBox IIf(ErrorMessage = "", err.description, ErrorMessage)
Capture:
    MsgBox IIf(ErrorMessage = "", err.description, ErrorMessage)
End Sub



Private Sub cmdinqure_Click()
 'frmassetregistration.Show vbModal
 frmAssetMaster.Show vbModal, Me
End Sub



Private Sub cmdprint_Click()
  On Error Resume Next
     STRFORMULA = ""
    reportname = "Depreciation.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, title
End Sub

Private Sub Command1_Click()
frmassetregistration.Show vbModal
End Sub

Private Sub Command2_Click()
    reportname = "AssetRegi.rpt"
    STRFORMULA = ""
    Show_Sales_Crystal_Report STRFORMULA, reportname, CompanyName
End Sub

Public Sub Form_Load()
   DTPicker1 = Date
   DTPicker2 = Date
Dim rsassets As New Recordset
Dim TotalCurr As Double
TotalCurr = 0
lvwasset.ListItems.Clear
sql = ""
sql = "select * from  assets order by AssetsNo asc"
Set rsassets = oSaccoMaster.GetRecordset(sql)
With rsassets
    While Not .EOF
        Set li = lvwasset.ListItems.Add(, , IIf(IsNull(!AssetsNo), "", !AssetsNo))
        li.SubItems(1) = IIf(IsNull(!AssetsName), "", !AssetsName)
        li.SubItems(2) = IIf(IsNull(!depreciation), 0, !depreciation)
        li.SubItems(3) = IIf(IsNull(!ACCNO), "", !ACCNO)
        li.SubItems(4) = IIf(IsNull(!ACCNO), "", !ACCNO)
        li.SubItems(5) = IIf(IsNull(!transdate), "", !transdate)
        li.SubItems(6) = IIf(IsNull(!PurchasePrice), "", !PurchasePrice)
        li.SubItems(7) = IIf(IsNull(!CurrentValue), "", !CurrentValue)
        li.SubItems(8) = IIf(IsNull(!AssetserialNo), "", !AssetserialNo)
        TotalCurr = TotalCurr + li.SubItems(7)
    .MoveNext
    Wend
End With
lbltotal = Format(TotalCurr, Cfmt)
End Sub

Private Sub lvwasset_DblClick()
frmAssetMaster.txtASSETSNO(1) = lvwasset.SelectedItem.Text
frmAssetMaster.txtASSETSNAME(1) = lvwasset.SelectedItem.ListSubItems(1)
frmAssetMaster.txtDEPRECIATION(1) = lvwasset.SelectedItem.ListSubItems(2)
frmAssetMaster.txtCrAccNo = lvwasset.SelectedItem.ListSubItems(3)
frmAssetMaster.txtDrAccNo = lvwasset.SelectedItem.ListSubItems(4)
frmAssetMaster.txtdatebought.value = lvwasset.SelectedItem.ListSubItems(5)
frmAssetMaster.txtPURCHASE(1) = lvwasset.SelectedItem.ListSubItems(6)
frmAssetMaster.txtCURRENTVALUE = lvwasset.SelectedItem.ListSubItems(7)
frmAssetMaster.txtSERIALNO(1) = IIf(IsNull(lvwasset.SelectedItem.ListSubItems(8)), 0, lvwasset.SelectedItem.ListSubItems(8))
frmAssetMaster.Show vbModal

End Sub


