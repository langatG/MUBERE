VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmAssetMaster 
   BackColor       =   &H00808000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Asset Master"
   ClientHeight    =   7770
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   7500
   Icon            =   "frmAssetMaster.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   7500
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkaddglsetup 
      Caption         =   "Add to GL"
      Height          =   255
      Left            =   5280
      TabIndex        =   53
      Top             =   4800
      Width           =   1455
   End
   Begin VB.TextBox txtCrAccName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1950
      TabIndex        =   48
      Top             =   5595
      Width           =   3225
   End
   Begin VB.TextBox lblDrAccName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1950
      TabIndex        =   47
      Top             =   5025
      Width           =   3225
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "..."
      Height          =   285
      Left            =   15
      TabIndex        =   46
      Top             =   5040
      Width           =   300
   End
   Begin VB.TextBox txtDrAccNo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   330
      TabIndex        =   45
      Top             =   5025
      Width           =   1440
   End
   Begin VB.CommandButton Command3 
      Caption         =   "..."
      Height          =   285
      Left            =   0
      TabIndex        =   44
      Top             =   5595
      Width           =   315
   End
   Begin VB.TextBox txtCrAccNo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   315
      TabIndex        =   43
      Top             =   5595
      Width           =   1440
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4575
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   7215
      Begin VB.ComboBox cboassettype 
         Height          =   315
         Left            =   2040
         TabIndex        =   54
         Top             =   1080
         Width           =   2415
      End
      Begin VB.CommandButton cmdfinder 
         Height          =   285
         Left            =   3840
         Picture         =   "frmAssetMaster.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Add New record"
         Top             =   120
         Width           =   375
      End
      Begin MSComCtl2.DTPicker txtdatebought 
         Height          =   375
         Left            =   1980
         TabIndex        =   3
         ToolTipText     =   "date when it was bought"
         Top             =   1440
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   145096705
         CurrentDate     =   38860
      End
      Begin MSComCtl2.DTPicker DTPtransdate 
         Height          =   375
         Left            =   1980
         TabIndex        =   8
         Top             =   3240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   145031169
         CurrentDate     =   37982
      End
      Begin VB.TextBox txtCURRENTVALUE 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1980
         TabIndex        =   7
         Top             =   2880
         Width           =   1815
      End
      Begin VB.TextBox txtUNITNO 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1980
         TabIndex        =   4
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox txtASSETSNO 
         Appearance      =   0  'Flat
         DataField       =   "assetsno"
         Height          =   285
         Index           =   1
         Left            =   1980
         TabIndex        =   0
         Top             =   120
         Width           =   1935
      End
      Begin VB.TextBox txtSERIALNO 
         Appearance      =   0  'Flat
         DataField       =   "assetserialno"
         Height          =   285
         Index           =   1
         Left            =   1980
         TabIndex        =   1
         Top             =   435
         Width           =   3135
      End
      Begin VB.CommandButton Command2 
         Height          =   375
         Left            =   6000
         Picture         =   "frmAssetMaster.frx":0704
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Assets"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6000
         TabIndex        =   21
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txtASSETSNAME 
         Appearance      =   0  'Flat
         DataField       =   "assetsname"
         Height          =   285
         Index           =   1
         Left            =   1980
         TabIndex        =   2
         Top             =   750
         Width           =   5175
      End
      Begin VB.TextBox txtDEPRECIATION 
         Appearance      =   0  'Flat
         DataField       =   "depreciation"
         Height          =   285
         Index           =   1
         Left            =   1980
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   2610
         Width           =   1815
      End
      Begin VB.TextBox txtPURCHASE 
         Appearance      =   0  'Flat
         DataField       =   "purchaseprice"
         Height          =   285
         Index           =   1
         Left            =   1980
         TabIndex        =   5
         Top             =   2295
         Width           =   1815
      End
      Begin VB.TextBox txtNOTES 
         Appearance      =   0  'Flat
         DataField       =   "notes"
         Height          =   285
         Index           =   1
         Left            =   1980
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   3615
         Width           =   5055
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   2820
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   20
         Top             =   1980
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Transaction Date:"
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
         Left            =   120
         TabIndex        =   41
         Top             =   3240
         Width           =   1575
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
         Left            =   870
         TabIndex        =   35
         Top             =   120
         Width           =   840
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
         Left            =   120
         TabIndex        =   34
         Top             =   450
         Width           =   1590
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Asset Type:"
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
         Index           =   2
         Left            =   690
         TabIndex        =   33
         Top             =   1110
         Width           =   1020
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
         Left            =   630
         TabIndex        =   32
         Top             =   780
         Width           =   1080
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Depreciation %:"
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
         Index           =   4
         Left            =   300
         TabIndex        =   31
         Top             =   2580
         Width           =   1350
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Purchase Price:"
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
         Index           =   5
         Left            =   345
         TabIndex        =   30
         Top             =   2250
         Width           =   1365
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Location:"
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
         Index           =   6
         Left            =   780
         TabIndex        =   29
         Top             =   3615
         Width           =   810
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Current Value:"
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
         Index           =   7
         Left            =   480
         TabIndex        =   28
         Top             =   2910
         Width           =   1230
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Unit No:"
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
         Index           =   8
         Left            =   990
         TabIndex        =   27
         Top             =   1920
         Width           =   720
      End
      Begin VB.Label Label1 
         Height          =   255
         Left            =   3960
         TabIndex        =   26
         Top             =   1080
         Width           =   2535
      End
      Begin VB.Label Label2 
         Height          =   255
         Left            =   3420
         TabIndex        =   25
         Top             =   1920
         Width           =   2535
      End
      Begin VB.Label Label3 
         Height          =   255
         Left            =   3870
         TabIndex        =   24
         Top             =   1530
         Width           =   2535
      End
      Begin VB.Label lblLabels 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "Date Bought:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   9
         Left            =   540
         TabIndex        =   23
         Top             =   1500
         Width           =   1140
      End
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   0
      ScaleHeight     =   540
      ScaleWidth      =   7500
      TabIndex        =   13
      Top             =   6690
      Width           =   7500
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   300
         Left            =   1213
         TabIndex        =   11
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   300
         Left            =   59
         TabIndex        =   10
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   300
         Left            =   4675
         TabIndex        =   18
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   300
         Left            =   3521
         TabIndex        =   17
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   300
         Left            =   2367
         TabIndex        =   16
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   300
         Left            =   1213
         TabIndex        =   15
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   300
         Left            =   59
         TabIndex        =   14
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.PictureBox picStatBox 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   0
      ScaleHeight     =   540
      ScaleWidth      =   7500
      TabIndex        =   12
      Top             =   7230
      Width           =   7500
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   615
         Picture         =   "frmAssetMaster.frx":0886
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   960
         Picture         =   "frmAssetMaster.frx":0BC8
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4665
         Picture         =   "frmAssetMaster.frx":0F0A
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   5010
         Picture         =   "frmAssetMaster.frx":124C
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1305
         TabIndex        =   40
         Top             =   0
         Width           =   3360
      End
   End
   Begin MSComctlLib.ListView lvwAccName 
      Height          =   1350
      Left            =   7590
      TabIndex        =   49
      Top             =   4965
      Visible         =   0   'False
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   2381
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "AccName"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "AccNo"
         Object.Width           =   18
      EndProperty
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "AccName"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1965
      TabIndex        =   52
      Top             =   4800
      Width           =   1635
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Dr AccNo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   360
      TabIndex        =   51
      Top             =   4815
      Width           =   765
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Cr AccNo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   360
      TabIndex        =   50
      Top             =   5370
      Width           =   765
   End
End
Attribute VB_Name = "frmAssetMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rst As Recordset

Private Sub cmdfinder_Click()
On Error GoTo ErrorHandler
frmsearchassets.Show vbModal
Dim Y As String
Y = sel
'm = False
If Y <> "" Then
     Dim cn As Connection
    Set cn = New ADODB.Connection
    
    cn.Open frmODBCLogon.cboDSNList, "atm"
sql = ""
'If reportpath = "" Then reportpath = GetSetting("payroll", "AppName", "rptPath", rptPath)
sql = "SELECT     AssetsNo, AssetserialNo, AssetsName, AssetType, datebought, UnitNo, PurchasePrice, Depreciation, Currentvalue, notes, Transdate FROM         assets where assetsno='" & Y & "' order by assetsno"
Set rs = New ADODB.Recordset
rs.Open sql, cn
If Not rs.EOF Then

If Not IsNull(rs.Fields(0)) Then txtASSETSNO(1) = (rs.Fields(0))
If Not IsNull(rs.Fields(1)) Then txtSERIALNO(1) = (rs.Fields(1))
If Not IsNull(rs.Fields(2)) Then txtASSETSNAME(1) = (rs.Fields(2))
If Not IsNull(rs.Fields(3)) Then cboassettype.Text = (rs.Fields(3))
If Not IsNull(rs.Fields(4)) Then txtdatebought = (rs.Fields(4))
If Not IsNull(rs.Fields(5)) Then txtUNITNO = (rs.Fields(5))
If Not IsNull(rs.Fields(6)) Then txtPURCHASE(1) = (rs.Fields(6))
If Not IsNull(rs.Fields(7)) Then txtDEPRECIATION(1) = (rs.Fields(7))
If Not IsNull(rs.Fields(8)) Then txtCURRENTVALUE = (rs.Fields(8))
If Not IsNull(rs.Fields(9)) Then txtNOTES(1) = (rs.Fields(9))
If Not IsNull(rs.Fields(10)) Then DTPTransdate = (rs.Fields(10))

'Call cboname_p

End If
End If
Exit Sub
ErrorHandler:
MsgBox err.description
End Sub

Private Sub cmdsearch_Click()
frmSearchGLAccounts.Show vbModal, Me
    If Continue Then
        If SearchValue <> "" Then
            txtDrAccNo = SearchValue
            SearchValue = ""
        End If
    End If
End Sub

Private Sub Command1_Click()
    frmAssets.Show vbModal
End Sub



Private Sub Command2_Click()
Dim errormsg As String
Dim path As String
path = Get_Path(errormsg)
  Set a = New CRAXDRT.Application
          Set r = a.OpenReport(path & "Assetsregister.rpt")
          r.ReadRecords
          
          With frmReports.CRViewer1
              .ReportSource = r
              .ViewReport
          End With
          
          frmReports.Show vbModal
          
          Set r = Nothing
End Sub

Private Sub Command3_Click()
frmSearchGLAccounts.Show vbModal, Me
    If Continue Then
        If SearchValue <> "" Then
            txtCrAccNo = SearchValue
            SearchValue = ""
        End If
    End If
End Sub



Private Sub Form_Load()
    On Error GoTo 10
    Set rst = New Recordset
      Dim cn As Connection
    Set cn = New ADODB.Connection
    
    cn.Open frmODBCLogon.cboDSNList, "atm"
    rst.Open "select * from assetS order by ASSETSNO", cn, adOpenKeyset, adLockOptimistic
    Dim oText As TextBox
   
   
   '// populate the asset type in the TXTASSETTYPE
   
   
   Dim rsg As Recordset
   sql = ""
   sql = "select Assetname from assetcode"
   Set rsg = New ADODB.Recordset
   rsg.Open sql, cn, adOpenKeyset, adLockOptimistic
   cboassettype.Clear
   While Not rsg.EOF
   cboassettype.AddItem rsg.Fields(0)
   rsg.MoveNext
   Wend
   txtdatebought = Format(Get_Server_Date, "dd/mm/yyyy")
   DTPTransdate = Format(Get_Server_Date, "dd/mm/yyyy")
    Exit Sub
10:    MsgBox err.description
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'  If mbEditFlag Or mbAddNewFlag Then Exit Sub

'  Select Case KeyCode
'    Case vbKeyEscape
'      cmdClose_Click
'    Case vbKeyEnd
'      cmdLast_Click
'    Case vbKeyHome
'      cmdFirst_Click
'    Case vbKeyUp, vbKeyPageUp
'      If Shift = vbCtrlMask Then
'        cmdFirst_Click
'      Else
'        cmdPrevious_Click
'      End If
'    Case vbKeyDown, vbKeyPageDown
'      If Shift = vbCtrlMask Then
'        cmdLast_Click
'      Else
'        cmdNext_Click
'      End If
'  End Select
'DTPtransdate.value = Format(Get_Server_Date, "DD/MM/YYYY")
'txtdatebought = Format(Get_Server_Date, "DD/MM/YYYY")
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub



Private Sub cmdAdd_Click()
  On Error GoTo AddErr
  With rst
   'If Not (.BOF And .EOF) Then
     ' mvBookMark = .Bookmark
   ' End If
'    .AddNew
    lblStatus.Caption = "Add record"
    ' True
    SetButtons False
  End With

  Exit Sub
AddErr:
  MsgBox err.description
End Sub

Private Sub cmddelete_Click()
  On Error GoTo DeleteErr
  MsgBox "Once an asset has been registered , you cannot delete it."
  Exit Sub
DeleteErr:
  MsgBox err.description
End Sub

Private Sub cmdrefresh_Click()
  'This is only needed for multi user apps
   On Error GoTo RefreshErr
  If txtASSETSNO(1) = "" Then
  Else
 
  rst.MoveLast
  End If
  Exit Sub
RefreshErr:
  MsgBox err.description
End Sub

Private Sub cmdedit_Click()
  On Error GoTo EditErr

  lblStatus.Caption = "Edit record"
  ' True
  SetButtons False
  Exit Sub

EditErr:
  MsgBox err.description
End Sub
Private Sub cmdcancel_Click()
  On Error Resume Next

  SetButtons True
  ' False
  ' False
'  rst.CancelUpdate
  '
'    rst.MoveFirst
  'End If
  ' False

End Sub

Private Sub cmdupdate_Click()
  On Error GoTo UpdateErr
    Set rst = New Recordset
      Dim cn As Connection
    Set cn = New ADODB.Connection
    
  cn.Open frmODBCLogon.cboDSNList, "atm"
  
  Dim transdate As Date, amount As Double, DRaccno As String, Craccno As String, DocumentNo As String, _
        TransSource As String, User1 As String, ErrorMessage As String, transDescription As String, CashBook As Long, doc_posted As Integer, chequeno As String
  
  If txtASSETSNO(1) = "" Or txtASSETSNAME(1) = "" Or cboassettype = "" Or txtCURRENTVALUE = "" Or txtDEPRECIATION(1) = "" Or txtNOTES(1) = "" Or txtPURCHASE(1) = "" Or txtSERIALNO(1) = "" Or txtUNITNO = "" Then
  MsgBox "Please Enter The Required Details", vbInformation, "ASSETS REGISTER"
  Exit Sub
  Else
  '//SELECT TO CHECK IF IT IS AVAILABLE
  sql = ""
  sql = "select assetsno from assets where assetsno='" & txtASSETSNO(1) & "'"
  Set rst = oSaccoMaster.GetRecordset(sql)
  If rst.EOF Then
  sql = "SET DATEFORMAT DMY insert into assets (assetsno,assetserialno,assetsname,assettype,Accno,datebought,unitno,purchaseprice,depreciation,currentvalue,notes,transdate) "
  sql = sql & " values('" & txtASSETSNO(1) & "','" & txtSERIALNO(1) & "','" & txtASSETSNAME(1) & "','" & cboassettype & "','" & txtDrAccNo & "','" & txtdatebought & "'," & txtUNITNO & "," & CDbl(txtPURCHASE(1)) & "," & txtDEPRECIATION(1) & "," & CDbl(txtCURRENTVALUE) & ",'" & txtNOTES(1) & "','" & DTPTransdate & "')"
  cn.Execute sql
  Else
  sql = ""
  sql = "SET DATEFORMAT DMY UPDATE    assets SET   assetserialno='" & txtSERIALNO(1) & "', assetsname='" & txtASSETSNAME(1) & "',Accno='" & txtDrAccNo & "',purchaseprice=" & CDbl(txtPURCHASE(1)) & ",datebought='" & txtdatebought & "',currentvalue=" & CDbl(txtCURRENTVALUE) & ",notes='" & txtNOTES(1) & "' where  assetsno='" & txtASSETSNO(1) & "'  "
  oSaccoMaster.ExecuteThis (sql)
  End If
  End If
  
  If chkaddglsetup = vbChecked Then
  transdate = DTPTransdate
  amount = txtCURRENTVALUE
  DRaccno = Trim(txtDrAccNo)
  Craccno = Trim(txtCrAccNo)
  DocumentNo = Trim(txtSERIALNO(1))
        TransSource = txtASSETSNAME(1)
        User1 = User
        transDescription = txtASSETSNAME(1)
        CashBook = 1
        doc_posted = 1
        chequeno = cboassettype.Text
  If Not Save_GLTRANSACTION(transdate, amount, DRaccno, Craccno, DocumentNo, _
        TransSource, User1, ErrorMessage, transDescription, CashBook, doc_posted, chequeno, TransNo) Then
            If ErrorMessage <> "" Then
                MsgBox ErrorMessage, vbInformation, Me.Caption
                ErrorMessage = ""
            End If
        End If
  End If
  MsgBox "Update completed successfully", vbInformation
  txtASSETSNAME(1) = ""
  txtASSETSNO(1) = ""
  
  txtCURRENTVALUE = ""
  txtDEPRECIATION(1) = ""
  txtNOTES(1) = ""
 
  txtPURCHASE(1) = ""
  txtSERIALNO(1) = ""
  txtUNITNO = ""
  SetButtons True
 
  Call cmdNext_Click
  Exit Sub
UpdateErr:
  MsgBox err.description
End Sub

Private Sub cmdclose_Click()
  Unload Me
End Sub

Private Sub cmdFirst_Click()
  On Error GoTo GoFirstError
  rst.MoveFirst
  Exit Sub
GoFirstError:
  MsgBox err.description
End Sub

Private Sub cmdLast_Click()
  On Error GoTo GoLastError

  rst.MoveLast
  Exit Sub

GoLastError:
  MsgBox err.description
End Sub

Private Sub cmdNext_Click()
  On Error GoTo GoNextError
    Beep
  Exit Sub
GoNextError:
  MsgBox err.description
End Sub

Private Sub cmdPrevious_Click()
  On Error GoTo GoPrevError

 
  Exit Sub

GoPrevError:
  MsgBox err.description
End Sub

Private Sub SetButtons(bVal As Boolean)
  cmdAdd.Visible = bVal
  cmdEdit.Visible = bVal
  cmdupdate.Visible = Not bVal
  cmdCancel.Visible = Not bVal
  cmdDelete.Visible = bVal
  cmdclose.Visible = bVal
  cmdRefresh.Visible = bVal
  cmdNext.Enabled = bVal
  cmdFirst.Enabled = bVal
  cmdLast.Enabled = bVal
  cmdPrevious.Enabled = bVal
End Sub

Private Sub Picture1_Click()
    On Error Resume Next
'    frmSelAsset.Show vbModal
    'If strName <> "" Then txtFields(2) = strName
End Sub

Private Sub Picture2_Click()
    On Error Resume Next
    'strName = txtFields(9)
   ' frmSelPropertyUnit.Show vbModal
    'If strName <> "" Then txtFields(8) = strName
End Sub

Private Sub Picture3_Click()
    On Error Resume Next
   ' frmSelProperty.Show vbModal
    If strName <> "" Then
       ' txtFields(9) = strName: txtFields(8) = ""
    End If
End Sub

Private Sub txtASSETSNO_Change(Index As Integer)
 On Error Resume Next
      Dim cn As Connection
    Set cn = New ADODB.Connection
    
    cn.Open frmODBCLogon.cboDSNList, "atm"
    Dim X As Recordset
    'If Index = 2 Then
        Label1 = ""
        Set X = New Recordset
        X.Open "select * from assetCODE where assetCODE='" & txtASSETSNO(1) & "'", cn, adOpenStatic, adLockOptimistic
        If X.RecordCount > 0 Then Label1 = X("aSSETname")
        X.Close
End Sub

'Private Sub txtFields_Change(Index As Integer)
'    On Error Resume Next
'      Dim cn As Connection
'    Set cn = New ADODB.Connection
'
'    cn.Open modCommon.pConnection
'    Dim x As Recordset
'    'If Index = 2 Then
'        Label1 = ""
'        Set x = New Recordset
'        x.Open "select * from assetCODE where assetCODE='" & txtFields(3) & "'", cn, adOpenStatic, adLockOptimistic
'        If x.RecordCount > 0 Then Label1 = x("aSSETname")
'        x.Close
'    'End If
'    If Index = 9 Then
'        Label3 = ""
'        Set x = New Recordset
'        'x.Open "select * from [property master] where [property id]='" & txtFields(9) & "'", cn, adOpenStatic, adLockOptimistic
'        'If x.RecordCount > 0 Then Label3 = x("property name")
'        'x.Close
'    End If
'    Label2 = txtFields(8)
'    If Index = 5 Or Index = 4 Then
'        If IsNumeric(txtFields(Index)) Then
'            If CCur(txtFields(4)) > 100 Or CCur(txtFields(4)) < 0 Then txtFields(4) = 0
'            txtFields(4) = Format(txtFields(4), "##0.00")
'            txtFields(5) = Format(txtFields(5), "#,##0.00")
'        End If
'    End If
'End Sub
Private Sub txtASSETSTYPE_Change(Index As Integer)

End Sub

Private Sub TXTASSETTYPE_Change()
Dim rsg As Recordset
  Dim cn As Connection
    Set cn = New ADODB.Connection
    cn.Open frmODBCLogon.cboDSNList, "atm"
Dim r As Double
   sql = ""
   sql = "select rate from assetcode WHERE ASSETname='" & cboassettype & "'"
   Set rsg = New ADODB.Recordset
   rsg.Open sql, cn, adOpenKeyset, adLockOptimistic
   
 If Not rsg.EOF Then
   If Not IsNull(rsg.Fields(0)) Then r = rsg.Fields(0)
   txtDEPRECIATION(1) = r
 End If
  
End Sub

Private Sub TXTASSETTYPE_Click()
TXTASSETTYPE_Change
End Sub

Private Sub txtCrAccNo_Change()
On Error GoTo SysError
    Dim Account As Acc_Details
        
        Editing = True
    Account = Get_Acc_Details(txtCrAccNo, ErrorMessage)
    If Account.ACCNO <> "" Then
        txtCrAccName = Account.AccName
    Else
        If ErrorMessage <> "" Then
            MsgBox ErrorMessage, vbInformation, Me.Caption
            ErrorMessage = ""
        End If
        txtCrAccName = ""
    End If
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub txtDrAccNo_Change()
   On Error GoTo SysError
    Dim Account As Acc_Details
    Editing = True
    Account = Get_Acc_Details(txtDrAccNo, ErrorMessage)
    If Account.ACCNO <> "" Then
        lblDrAccName = Account.AccName
    Else
        If ErrorMessage <> "" Then
            MsgBox ErrorMessage, vbInformation, Me.Caption
            ErrorMessage = ""
        End If
        lblDrAccName = ""
    End If
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub
