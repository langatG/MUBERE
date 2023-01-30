VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Begin VB.Form frmmilk 
   Caption         =   "Income Statements"
   ClientHeight    =   9525
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   ScaleHeight     =   9525
   ScaleWidth      =   9330
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      Height          =   375
      Left            =   7560
      TabIndex        =   15
      Top             =   8640
      Width           =   1575
   End
   Begin VB.CommandButton cmdprint 
      Caption         =   "Print Profit and Loss"
      Height          =   375
      Left            =   7560
      TabIndex        =   14
      Top             =   9120
      Width           =   1575
   End
   Begin VB.Frame Frame3 
      Caption         =   "Expenses"
      Height          =   2535
      Left            =   240
      TabIndex        =   2
      Top             =   6000
      Width           =   8775
      Begin MSComctlLib.ListView lvwexpensens 
         Height          =   2055
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   3625
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
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
   Begin VB.Frame Frame2 
      Caption         =   "Production/Carriages"
      Height          =   2415
      Left            =   240
      TabIndex        =   1
      Top             =   3120
      Width           =   8775
      Begin MSComctlLib.ListView lvwproductioncosts 
         Height          =   2055
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   3625
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
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
   Begin VB.Frame Frame1 
      Caption         =   "Sales/Income"
      Height          =   2535
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   8775
      Begin MSComctlLib.ListView lvwincome 
         Height          =   2055
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   3625
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
   Begin VB.Label lbloverallexpenses 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5400
      TabIndex        =   13
      Top             =   9120
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "Overall Perfomance"
      Height          =   255
      Left            =   3120
      TabIndex        =   12
      Top             =   9120
      Width           =   1695
   End
   Begin VB.Label lblexpenses 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5400
      TabIndex        =   11
      Top             =   8640
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Total Expenses"
      Height          =   255
      Left            =   3120
      TabIndex        =   10
      Top             =   8640
      Width           =   1455
   End
   Begin VB.Label lblproductionscosts 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5160
      TabIndex        =   9
      Top             =   5640
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Total Production Costs"
      Height          =   255
      Left            =   2880
      TabIndex        =   8
      Top             =   5640
      Width           =   1935
   End
   Begin VB.Label lbltotalincome 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5160
      TabIndex        =   5
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Total Income"
      Height          =   255
      Left            =   2880
      TabIndex        =   4
      Top             =   2760
      Width           =   1935
   End
End
Attribute VB_Name = "frmmilk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdPrint_Click()
Set rs = oSaccoMaster.GetRecordset("SELECT     TOP 1 Name FROM         d_company")
If Not rs.EOF Then
CompanyName = IIf(IsNull(rs.Fields(0)), "", rs.Fields(0))
End If
'If frmAccounts.chkall = vbChecked Then
' reportname = "kimincomeandexpenditure.rpt"
'
' Show_Sales_Crystal_Report "", reportname, CompanyName
' Else
'
' reportname = "kimincomeandexpenditure.rpt"
' STRFORMULA = "{TBBALANCE.AccType} = '" & frmAccounts.cboacccategory & "'"
' Show_Sales_Crystal_Report STRFORMULA, reportname, CompanyName
' End If

End Sub

Private Sub Form_Load()
    On Error GoTo SysError
'On Error GoTo SysError


   

        Dim a As Currency

        lvwincome.ListItems.Clear
        Dim rsTrans As New Recordset, DRTotal As Double, CRTotal As Double
        Set rsTrans = oSaccoMaster.GetRecordset("SELECT     TBBALANCE.ACCNO,TBBALANCE.ACCNAME,TBBALANCE.AMOUNT  FROM         TBBALANCE TBBALANCE INNER JOIN  GLSETUP GLSETUP ON TBBALANCE.AccNo = GLSETUP.AccNo WHERE     (glsetup.glaccmaingroup='INCOME')")
        DRTotal = 0
       
        CRTotal = 0
        With lvwincome
            
                While Not rsTrans.EOF
                    Set li = lvwincome.ListItems.Add(, , IIf(IsNull(rsTrans!ACCNO), "", rsTrans!ACCNO))
                    
                     li.SubItems(1) = IIf(IsNull(rsTrans!AccName), "", rsTrans!AccName)
                     li.SubItems(2) = IIf(IsNull(Format(rsTrans!amount, "###,###,###.0#")), 0#, (Format(rsTrans!amount, "###,###,###.0#")))
                     CRTotal = li.SubItems(2) + CRTotal
                    rsTrans.MoveNext
                Wend
            
        End With
        lbltotalincome = Format(CRTotal, "###,###,###.0#")
        
         Set rsTrans = oSaccoMaster.GetRecordset("SELECT     TBBALANCE.ACCNO,TBBALANCE.ACCNAME,TBBALANCE.AMOUNT  FROM         TBBALANCE TBBALANCE INNER JOIN  GLSETUP GLSETUP ON TBBALANCE.AccNo = GLSETUP.AccNo WHERE     (glsetup.glaccmaingroup='PRODUCTION COST')")
        DRTotal = 0
       
        CRTotal = 0
        With lvwproductioncosts
            
                While Not rsTrans.EOF
                    Set li = lvwproductioncosts.ListItems.Add(, , IIf(IsNull(rsTrans!ACCNO), "", rsTrans!ACCNO))
                    
                     li.SubItems(1) = IIf(IsNull(rsTrans!AccName), "", rsTrans!AccName)
                     li.SubItems(2) = IIf(IsNull(Format(rsTrans!amount, "###,###,###.0#")), 0#, (Format(rsTrans!amount, "###,###,###.0#")))
                     CRTotal = li.SubItems(2) + CRTotal
                    rsTrans.MoveNext
                Wend
            
        End With
        lblproductionscosts = Format(CRTotal, "###,###,###.0#")
      
    Set rsTrans = oSaccoMaster.GetRecordset("SELECT     TBBALANCE.ACCNO,TBBALANCE.ACCNAME,TBBALANCE.AMOUNT  FROM         TBBALANCE TBBALANCE INNER JOIN  GLSETUP GLSETUP ON TBBALANCE.AccNo = GLSETUP.AccNo WHERE     (glsetup.glaccmaingroup='EXPENSES')")
        DRTotal = 0
       
        CRTotal = 0
        With lvwexpensens
            
                While Not rsTrans.EOF
                    Set li = lvwexpensens.ListItems.Add(, , IIf(IsNull(rsTrans!ACCNO), "", rsTrans!ACCNO))
                    
                     li.SubItems(1) = IIf(IsNull(rsTrans!AccName), "", rsTrans!AccName)
                     li.SubItems(2) = IIf(IsNull(Format(rsTrans!amount, "###,###,###.0#")), "", (Format(rsTrans!amount, "###,###,###.0#")))
                     CRTotal = li.SubItems(2) + CRTotal
                    rsTrans.MoveNext
                Wend
            
        End With
        lblexpenses = Format(CRTotal, "###,###,###.0#")
        
        lbloverallexpenses = CDbl(lbltotalincome) - CDbl(lblproductionscosts) - CDbl(lblexpenses)
        lbloverallexpenses = Format(lbloverallexpenses, "###,###,###.0#")

    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption

End Sub
