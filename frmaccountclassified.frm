VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmaccountclassified 
   Caption         =   "ACCOUNTS BY CATEGORY"
   ClientHeight    =   7980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9315
   LinkTopic       =   "Form1"
   ScaleHeight     =   7980
   ScaleWidth      =   9315
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Sales/Income"
      Height          =   6855
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   8895
      Begin MSComctlLib.ListView lvwincome 
         Height          =   6375
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   11245
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "AccountNo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Amount"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.ComboBox cbosubheader 
      Height          =   315
      Left            =   5040
      TabIndex        =   3
      Top             =   360
      Width           =   4095
   End
   Begin VB.ComboBox cboaccountheader 
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Sub Account Header"
      Height          =   255
      Left            =   3360
      TabIndex        =   2
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Account Header"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "frmaccountclassified"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cboaccountheader_Change()
 
'On Error GoTo SysError


   

        Dim a As Currency

        lvwincome.ListItems.Clear
        Dim rsTrans As New Recordset, DRTotal As Double, CRTotal As Double
        Set rsTrans = oSaccoMaster.GetRecordset("SELECT     TBBALANCE.ACCNO,TBBALANCE.ACCNAME,TBBALANCE.AMOUNT  FROM         TBBALANCE TBBALANCE INNER JOIN  GLSETUP GLSETUP ON TBBALANCE.AccNo = GLSETUP.AccNo WHERE     (glsetup.header='" & cboaccountheader & "')")
        DRTotal = 0
       
        CRTotal = 0
        With lvwincome
            
                While Not rsTrans.EOF
                    Set li = lvwincome.ListItems.Add(, , IIf(IsNull(rsTrans!accno), "", rsTrans!accno))
                    
                     li.SubItems(1) = IIf(IsNull(rsTrans!AccName), "", rsTrans!AccName)
                     li.SubItems(2) = IIf(IsNull(Format(rsTrans!AMOUNT, "###,###,###.0#")), 0#, (Format(rsTrans!AMOUNT, "###,###,###.0#")))
                     CRTotal = li.SubItems(2) + CRTotal
                    rsTrans.MoveNext
                Wend
            
        End With
End Sub

Private Sub cboaccountheader_Click()
cboaccountheader_Change
End Sub

Private Sub Form_Load()
 Dim mClass As New GeneralLedger
    Dim Account As Account_Details
        Set rs = oSaccoMaster.GetRecordset("select distinct NewGLOpeningBalDate from GLSETUP")
    If Not rs.EOF Then

    End If
Set rs = oSaccoMaster.GetRecordset("SELECT HName FROM  d_Headers ORDER BY Hname")
    While Not rs.EOF
    
    If Not IsNull(rs.Fields(0)) Then cboaccountheader.AddItem (rs.Fields(0))

    rs.MoveNext
    Wend
    
    Set rs = oSaccoMaster.GetRecordset("SELECT MName FROM   d_MainAccount ORDER BY Mname")
    While Not rs.EOF
    
    If Not IsNull(rs.Fields(0)) Then cbosubheader.AddItem (rs.Fields(0))

    rs.MoveNext
    Wend

mysql = "Get_OpeningBalances '30/12/2009'"
oSaccoMaster.ExecuteThis (mysql)

    If Not mClass.generate_trialbalance("30/12/2009", Date, "MAZIWA", ErrorMessage) Then
        If ErrorMessage <> "" Then
            MsgBox ErrorMessage, vbInformation, Me.Caption
            ErrorMessage = ""
           ' Label1.Visible = False
        Else
            'Label1.Visible = False
        End If
    End If
End Sub
