VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmtrialbalance 
   AutoRedraw      =   -1  'True
   Caption         =   "TRIAL BALANCE GENERATION"
   ClientHeight    =   9075
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11505
   DrawStyle       =   6  'Inside Solid
   Icon            =   "frmtrialbalance.frx":0000
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   14689.2
   ScaleMode       =   0  'User
   ScaleWidth      =   11907.27
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Trial Balance"
      Height          =   7935
      Left            =   120
      TabIndex        =   53
      Top             =   600
      Width           =   11175
      Begin MSComctlLib.ListView lvwincome 
         Height          =   7455
         Left            =   120
         TabIndex        =   54
         Top             =   360
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   13150
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
         NumItems        =   4
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
            Text            =   "DR"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "CR"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin MSComCtl2.DTPicker DTPGENE 
      Height          =   255
      Left            =   11760
      TabIndex        =   52
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      _Version        =   393216
      Format          =   65011713
      CurrentDate     =   38451
   End
   Begin VB.TextBox TT11 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9000
      TabIndex        =   49
      Text            =   "66009999"
      Top             =   13440
      Width           =   1575
   End
   Begin VB.TextBox T11 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6960
      TabIndex        =   48
      Text            =   "66000000"
      Top             =   13440
      Width           =   1575
   End
   Begin VB.ComboBox Combo11 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      TabIndex        =   47
      Top             =   13440
      Width           =   1815
   End
   Begin VB.TextBox TT10 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9000
      TabIndex        =   44
      Text            =   "65009999"
      Top             =   12960
      Width           =   1575
   End
   Begin VB.TextBox T10 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6960
      TabIndex        =   43
      Text            =   "65000000"
      Top             =   12960
      Width           =   1575
   End
   Begin VB.ComboBox Combo10 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      TabIndex        =   42
      Top             =   12960
      Width           =   1815
   End
   Begin VB.TextBox TT9 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9000
      TabIndex        =   39
      Text            =   "63009999"
      Top             =   12480
      Width           =   1575
   End
   Begin VB.TextBox T9 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6960
      TabIndex        =   38
      Text            =   "63000000"
      Top             =   12480
      Width           =   1575
   End
   Begin VB.ComboBox Combo9 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      TabIndex        =   37
      Top             =   12480
      Width           =   1815
   End
   Begin VB.TextBox TT8 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9000
      TabIndex        =   34
      Text            =   "62009999"
      Top             =   12000
      Width           =   1575
   End
   Begin VB.TextBox T8 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6960
      TabIndex        =   33
      Text            =   "62000000"
      Top             =   12000
      Width           =   1575
   End
   Begin VB.ComboBox Combo8 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      TabIndex        =   32
      Top             =   12000
      Width           =   1815
   End
   Begin VB.ComboBox Combo7 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      TabIndex        =   29
      Top             =   11520
      Width           =   1815
   End
   Begin VB.TextBox T7 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6960
      TabIndex        =   28
      Text            =   "61000000"
      Top             =   11520
      Width           =   1575
   End
   Begin VB.TextBox TT7 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9000
      TabIndex        =   27
      Text            =   "61009999"
      Top             =   11520
      Width           =   1575
   End
   Begin VB.TextBox TT6 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9000
      TabIndex        =   26
      Text            =   "59009999"
      Top             =   11040
      Width           =   1575
   End
   Begin VB.TextBox T6 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6960
      TabIndex        =   25
      Text            =   "59000000"
      Top             =   11040
      Width           =   1575
   End
   Begin VB.TextBox TT5 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9000
      TabIndex        =   24
      Text            =   "56009999"
      Top             =   10560
      Width           =   1575
   End
   Begin VB.TextBox T5 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6960
      TabIndex        =   23
      Text            =   "56000000"
      Top             =   10560
      Width           =   1575
   End
   Begin VB.TextBox TT4 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9000
      TabIndex        =   22
      Text            =   "55009999"
      Top             =   10080
      Width           =   1575
   End
   Begin VB.TextBox T4 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6960
      TabIndex        =   21
      Text            =   "55000000"
      Top             =   10080
      Width           =   1575
   End
   Begin VB.TextBox TT3 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9000
      TabIndex        =   20
      Text            =   "53009999"
      Top             =   9600
      Width           =   1575
   End
   Begin VB.TextBox T3 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6960
      TabIndex        =   19
      Text            =   "53000000"
      Top             =   9600
      Width           =   1575
   End
   Begin VB.TextBox TT2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9000
      TabIndex        =   18
      Text            =   "52009999"
      Top             =   9120
      Width           =   1575
   End
   Begin VB.TextBox T2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6960
      TabIndex        =   17
      Text            =   "52000000"
      Top             =   9120
      Width           =   1575
   End
   Begin VB.ComboBox Combo6 
      Height          =   315
      ItemData        =   "frmtrialbalance.frx":08CA
      Left            =   120
      List            =   "frmtrialbalance.frx":08CC
      TabIndex        =   10
      Top             =   11040
      Width           =   1815
   End
   Begin VB.ComboBox Combo5 
      Height          =   315
      Left            =   120
      TabIndex        =   8
      Top             =   10560
      Width           =   1815
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Top             =   10080
      Width           =   1815
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   9600
      Width           =   1815
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   9120
      Width           =   1815
   End
   Begin MSComCtl2.DTPicker transdate 
      Height          =   300
      Left            =   3240
      TabIndex        =   1
      Top             =   195
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   393216
      CalendarBackColor=   8421376
      CalendarForeColor=   65535
      CalendarTitleForeColor=   8438015
      CalendarTrailingForeColor=   8421631
      Format          =   65011713
      CurrentDate     =   38199
   End
   Begin VB.Label lbltotaldebits 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   8160
      TabIndex        =   58
      Top             =   8640
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Total Debits"
      Height          =   255
      Left            =   6120
      TabIndex        =   57
      Top             =   8640
      Width           =   1815
   End
   Begin VB.Label lbltotalcredits 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2640
      TabIndex        =   56
      Top             =   8640
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Total Credits"
      Height          =   255
      Left            =   600
      TabIndex        =   55
      Top             =   8640
      Width           =   1815
   End
   Begin VB.Label d11 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4320
      TabIndex        =   51
      Top             =   13440
      Width           =   2415
   End
   Begin VB.Label L11 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2040
      TabIndex        =   50
      Top             =   13440
      Width           =   2175
   End
   Begin VB.Label d10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4320
      TabIndex        =   46
      Top             =   12960
      Width           =   2415
   End
   Begin VB.Label L10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2040
      TabIndex        =   45
      Top             =   12960
      Width           =   2175
   End
   Begin VB.Label d9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4320
      TabIndex        =   41
      Top             =   12480
      Width           =   2415
   End
   Begin VB.Label L9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2040
      TabIndex        =   40
      Top             =   12480
      Width           =   2175
   End
   Begin VB.Label d8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4320
      TabIndex        =   36
      Top             =   12000
      Width           =   2415
   End
   Begin VB.Label L8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2040
      TabIndex        =   35
      Top             =   12000
      Width           =   2175
   End
   Begin VB.Label L7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2040
      TabIndex        =   31
      Top             =   11520
      Width           =   2175
   End
   Begin VB.Label d7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4320
      TabIndex        =   30
      Top             =   11520
      Width           =   2415
   End
   Begin VB.Label d3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4320
      TabIndex        =   16
      Top             =   9600
      Width           =   2415
   End
   Begin VB.Label d4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4320
      TabIndex        =   15
      Top             =   10080
      Width           =   2415
   End
   Begin VB.Label d5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4320
      TabIndex        =   14
      Top             =   10560
      Width           =   2415
   End
   Begin VB.Label d6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4320
      TabIndex        =   13
      Top             =   11040
      Width           =   2415
   End
   Begin VB.Label d2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4320
      TabIndex        =   12
      Top             =   9120
      Width           =   2415
   End
   Begin VB.Label L6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2040
      TabIndex        =   11
      Top             =   11040
      Width           =   2175
   End
   Begin VB.Label L5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2040
      TabIndex        =   9
      Top             =   10560
      Width           =   2175
   End
   Begin VB.Label L4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2040
      TabIndex        =   7
      Top             =   10080
      Width           =   2175
   End
   Begin VB.Label L3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2040
      TabIndex        =   5
      Top             =   9600
      Width           =   2175
   End
   Begin VB.Label L2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2040
      TabIndex        =   3
      Top             =   9120
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "TransDate:"
      Height          =   255
      Left            =   1920
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "frmtrialbalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdprint_Click()

 reportname = "kimincomeandexpenditure.rpt"
 
 Show_Sales_Crystal_Report "", reportname, "& Company_Name &"

End Sub

Private Sub Form_Load()
    On Error GoTo SysError
'On Error GoTo SysError


   

        Dim a As Currency

        lvwincome.ListItems.Clear
        Dim rsTrans As New Recordset, DRTotal As Double, CRTotal As Double
        Set rsTrans = oSaccoMaster.GetRecordset("SELECT     TBBALANCE.ACCNO,TBBALANCE.ACCNAME,TBBALANCE.AMOUNT,TBBALANCE.TRANSTYPE  FROM         TBBALANCE TBBALANCE INNER JOIN  GLSETUP GLSETUP ON TBBALANCE.AccNo = GLSETUP.AccNo")
        DRTotal = 0
       
        CRTotal = 0
        With lvwincome
            
                While Not rsTrans.EOF
                    Set li = lvwincome.ListItems.Add(, , IIf(IsNull(rsTrans!AccNo), "", rsTrans!AccNo))
                    
                     li.SubItems(1) = IIf(IsNull(rsTrans!AccName), "", rsTrans!AccName)
                     'li.SubItems(2) = IIf(IsNull(Format(rsTrans!amount, "###,###,###.0#")), 0#, (Format(rsTrans!amount, "###,###,###.0#")))
                          If UCase(Trim(rsTrans!transtype)) = "DR" Then
                                li.ListSubItems.Add , , Format(rsTrans!amount, "###,###,###.00")
                                li.ListSubItems.Add , , Format(0, "0.00")
                                 DRTotal = rsTrans!amount + DRTotal
                                
                            Else
                                    li.ListSubItems.Add , , Format(0, "0.00")
                                     li.ListSubItems.Add , , Format(rsTrans!amount, "###,###,###.00")
                                CRTotal = rsTrans!amount + CRTotal
                            End If
     
                    rsTrans.MoveNext
                Wend
            
        End With
        lbltotalcredits = Format(CRTotal, "###,###,###.0#")
        lbltotaldebits = Format(DRTotal, "###,###,###.0#")
        
        
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption

End Sub
