VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmbudgetting 
   Caption         =   "BUDGETTING"
   ClientHeight    =   6420
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9195
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmbudgetting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   9195
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   330
      Top             =   5910
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox cboyear 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmbudgetting.frx":0442
      Left            =   3000
      List            =   "frmbudgetting.frx":0464
      TabIndex        =   9
      Top             =   30
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   7905
      TabIndex        =   1
      Top             =   5970
      Width           =   1095
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5415
      Left            =   150
      TabIndex        =   0
      Top             =   480
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   9551
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "BUDGETS"
      TabPicture(0)   =   "frmbudgetting.frx":04A4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label18"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblactuals"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "error"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "LBLGLCONTRA"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Picture1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdupdate"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "dtpBudgetYear"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cboPeriod"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lvwBudget"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "glName1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Frame1"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lvwAccounts"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "COMPARISON"
      TabPicture(1)   =   "frmbudgetting.frx":04C0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(1)=   "Lvwcategory"
      Tab(1).Control(2)=   "dtpBudgetPeriod"
      Tab(1).Control(3)=   "cmdLoad"
      Tab(1).Control(4)=   "cmdExport"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "PLANT STATISTICS"
      TabPicture(2)   =   "frmbudgetting.frx":04DC
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label19"
      Tab(2).Control(1)=   "Label17"
      Tab(2).Control(2)=   "lblfemales"
      Tab(2).Control(3)=   "Label16"
      Tab(2).Control(4)=   "lblmales"
      Tab(2).Control(5)=   "Label12"
      Tab(2).Control(6)=   "Label14"
      Tab(2).Control(7)=   "lbltotalshares"
      Tab(2).Control(8)=   "Label15"
      Tab(2).Control(9)=   "lbltotalloanportfolio"
      Tab(2).Control(10)=   "lblnoofmembers"
      Tab(2).Control(11)=   "Label11"
      Tab(2).Control(12)=   "Label13"
      Tab(2).Control(13)=   "lowestloanee"
      Tab(2).Control(14)=   "memberno4"
      Tab(2).Control(15)=   "Label7"
      Tab(2).Control(16)=   "Label8"
      Tab(2).Control(17)=   "memberno3"
      Tab(2).Control(18)=   "memberno1"
      Tab(2).Control(19)=   "Label10"
      Tab(2).Control(20)=   "Label9"
      Tab(2).Control(21)=   "lowestsaver"
      Tab(2).Control(22)=   "memberno"
      Tab(2).Control(23)=   "lblmemberno"
      Tab(2).Control(24)=   "highestloanee"
      Tab(2).Control(25)=   "Label6"
      Tab(2).Control(26)=   "lblhighestsaver"
      Tab(2).Control(27)=   "Label5"
      Tab(2).ControlCount=   28
      Begin VB.CommandButton cmdExport 
         Caption         =   "&Export"
         Height          =   345
         Left            =   -72075
         TabIndex        =   42
         Top             =   675
         Width           =   1245
      End
      Begin VB.CommandButton cmdLoad 
         Caption         =   "&Load"
         Height          =   345
         Left            =   -73410
         TabIndex        =   41
         Top             =   675
         Width           =   1245
      End
      Begin MSComCtl2.DTPicker dtpBudgetPeriod 
         Height          =   315
         Left            =   -74880
         TabIndex        =   39
         Top             =   690
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   " MMM yyyy"
         Format          =   158466051
         CurrentDate     =   39533
      End
      Begin MSComctlLib.ListView lvwAccounts 
         Height          =   1260
         Left            =   3600
         TabIndex        =   32
         Top             =   -480
         Visible         =   0   'False
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   2223
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "AccNo"
            Object.Width           =   18
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "AccName"
            Object.Width           =   10583
         EndProperty
      End
      Begin VB.Frame Frame1 
         Caption         =   "Budget Method"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1680
         Left            =   5070
         TabIndex        =   33
         Top             =   1530
         Width           =   3720
         Begin VB.ComboBox txtmonthd 
            Height          =   315
            ItemData        =   "frmbudgetting.frx":04F8
            Left            =   2160
            List            =   "frmbudgetting.frx":0502
            TabIndex        =   55
            Top             =   480
            Width           =   1095
         End
         Begin VB.OptionButton optFixed 
            Caption         =   "Fixed Amount"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   105
            TabIndex        =   38
            Top             =   675
            Value           =   -1  'True
            Width           =   1965
         End
         Begin VB.OptionButton optSpread 
            Caption         =   "Spread Amount"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   90
            TabIndex        =   37
            Top             =   330
            Width           =   1890
         End
         Begin VB.TextBox txtbudgettedamount 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1590
            TabIndex        =   34
            Top             =   1260
            Width           =   1695
         End
         Begin VB.Label Label20 
            Caption         =   "Budget per month?"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2040
            TabIndex        =   56
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label lblAmount 
            Alignment       =   1  'Right Justify
            Caption         =   "Spread Amount"
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
            Left            =   45
            TabIndex        =   36
            Top             =   1290
            Width           =   1455
         End
      End
      Begin VB.TextBox glName1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3600
         TabIndex        =   31
         Top             =   600
         Width           =   4965
      End
      Begin MSComctlLib.ListView lvwBudget 
         Height          =   3165
         Left            =   120
         TabIndex        =   30
         Top             =   1620
         Width           =   4830
         _ExtentX        =   8520
         _ExtentY        =   5583
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Period"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "End Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Budgetted Amount"
            Object.Width           =   3881
         EndProperty
      End
      Begin VB.ComboBox cboPeriod 
         Height          =   315
         ItemData        =   "frmbudgetting.frx":050F
         Left            =   5505
         List            =   "frmbudgetting.frx":0519
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   1155
         Visible         =   0   'False
         Width           =   1845
      End
      Begin MSComCtl2.DTPicker dtpBudgetYear 
         Height          =   315
         Left            =   1230
         TabIndex        =   5
         Top             =   1170
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy"
         Format          =   158466051
         CurrentDate     =   39173
      End
      Begin VB.CommandButton cmdupdate 
         Appearance      =   0  'Flat
         Caption         =   "Update"
         Height          =   375
         Left            =   4995
         TabIndex        =   35
         Top             =   4410
         Width           =   1215
      End
      Begin VB.PictureBox Picture1 
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
         Left            =   1560
         Picture         =   "frmbudgetting.frx":052E
         ScaleHeight     =   195
         ScaleWidth      =   210
         TabIndex        =   3
         Top             =   600
         Width           =   270
      End
      Begin VB.TextBox LBLGLCONTRA 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1860
         TabIndex        =   2
         Top             =   600
         Width           =   1695
      End
      Begin MSComctlLib.ListView Lvwcategory 
         Height          =   4215
         Left            =   -74880
         TabIndex        =   12
         Top             =   1125
         Width           =   8670
         _ExtentX        =   15293
         _ExtentY        =   7435
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "AccNo"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "AccName"
            Object.Width           =   4763
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Budgetted Amount"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Actual Amount"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Variance"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Percentage"
            Object.Width           =   3175
         EndProperty
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   -68010
         TabIndex        =   54
         Top             =   4395
         Width           =   1215
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "No Of Members With Loans"
         Height          =   195
         Left            =   -70110
         TabIndex        =   53
         Top             =   4425
         Width           =   1950
      End
      Begin VB.Label lblfemales 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   -72855
         TabIndex        =   52
         Top             =   4335
         Width           =   1395
      End
      Begin VB.Label Label16 
         Caption         =   "Female"
         Height          =   255
         Left            =   -74715
         TabIndex        =   51
         Top             =   4320
         Width           =   780
      End
      Begin VB.Label lblmales 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   -72855
         TabIndex        =   50
         Top             =   3810
         Width           =   1395
      End
      Begin VB.Label Label12 
         Caption         =   "Males"
         Height          =   255
         Left            =   -74715
         TabIndex        =   49
         Top             =   3810
         Width           =   780
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Total Shares"
         Height          =   195
         Left            =   -70125
         TabIndex        =   48
         Top             =   3990
         Width           =   900
      End
      Begin VB.Label lbltotalshares 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   -68700
         TabIndex        =   47
         Top             =   4005
         Width           =   1890
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Total Loan Porfolio"
         Height          =   195
         Left            =   -70200
         TabIndex        =   46
         Top             =   3405
         Width           =   1335
      End
      Begin VB.Label lbltotalloanportfolio 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   -68700
         TabIndex        =   45
         Top             =   3420
         Width           =   1890
      End
      Begin VB.Label lblnoofmembers 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   -72840
         TabIndex        =   44
         Top             =   3375
         Width           =   2055
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "No Of Members"
         Height          =   195
         Left            =   -74730
         TabIndex        =   43
         Top             =   3375
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Budget Period"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74865
         TabIndex        =   40
         Top             =   450
         Width           =   1185
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Lowest Loanee"
         Height          =   195
         Left            =   -74760
         TabIndex        =   28
         Top             =   2760
         Width           =   1080
      End
      Begin VB.Label lowestloanee 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   -72840
         TabIndex        =   27
         Top             =   2715
         Width           =   2055
      End
      Begin VB.Label memberno4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   -69000
         TabIndex        =   26
         Top             =   2715
         Width           =   2175
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Member No."
         Height          =   195
         Left            =   -70200
         TabIndex        =   25
         Top             =   2760
         Width           =   870
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Member No."
         Height          =   195
         Left            =   -70200
         TabIndex        =   24
         Top             =   2160
         Width           =   870
      End
      Begin VB.Label memberno3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   -69000
         TabIndex        =   23
         Top             =   2115
         Width           =   2175
      End
      Begin VB.Label memberno1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   -69000
         TabIndex        =   22
         Top             =   1515
         Width           =   2175
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Member No."
         Height          =   195
         Left            =   -70200
         TabIndex        =   21
         Top             =   1560
         Width           =   870
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Lowest Saver"
         Height          =   195
         Left            =   -74760
         TabIndex        =   20
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label lowestsaver 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   -72840
         TabIndex        =   19
         Top             =   1515
         Width           =   2055
      End
      Begin VB.Label memberno 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   -69000
         TabIndex        =   18
         Top             =   915
         Width           =   2175
      End
      Begin VB.Label lblmemberno 
         AutoSize        =   -1  'True
         Caption         =   "Member No."
         Height          =   195
         Left            =   -70200
         TabIndex        =   17
         Top             =   960
         Width           =   870
      End
      Begin VB.Label highestloanee 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   -72840
         TabIndex        =   16
         Top             =   2115
         Width           =   2055
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Highest Loanee"
         Height          =   195
         Left            =   -74760
         TabIndex        =   15
         Top             =   2160
         Width           =   1110
      End
      Begin VB.Label lblhighestsaver 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   -72840
         TabIndex        =   14
         Top             =   915
         Width           =   2055
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Highest Saver"
         Height          =   195
         Left            =   -74760
         TabIndex        =   13
         Top             =   945
         Width           =   1005
      End
      Begin VB.Label error 
         BackColor       =   &H008080FF&
         Caption         =   "The Budget For Year Choosen Has Not Been Set"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   735
         Left            =   5175
         TabIndex        =   11
         Top             =   645
         Visible         =   0   'False
         Width           =   3615
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblactuals 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6285
         TabIndex        =   8
         Top             =   4020
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Actuals"
         Height          =   255
         Left            =   5700
         TabIndex        =   7
         Top             =   4065
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Budget Year"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   135
         TabIndex        =   6
         Top             =   1215
         Width           =   1035
      End
      Begin VB.Label Label18 
         Caption         =   "GL Contra Account"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1575
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Budget Year"
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
      Left            =   1395
      TabIndex        =   10
      Top             =   75
      Visible         =   0   'False
      Width           =   1500
   End
End
Attribute VB_Name = "frmbudgetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myclass As cdbase
Dim ulmt As Integer
Dim number As Integer
Dim principal As Currency
Dim ref As Integer
Dim interest1 As Currency
Dim currentbalance As Currency
Dim remi As Double
Dim expectedamount1 As Currency, expected2 As Currency, expected3 As Currency, expected4 As Currency, expected5 As Currency
Dim remainder1 As Double, remainder2 As Double, remainder3 As Double, remainder4 As Double, remainder5 As Double
Dim remainder6 As Double, remainder7 As Double

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
Dim loan
Public maxRec As Long
Public myLevel As Integer

Private Type accoInfo
    ACCNO As String
    custName As String
    custBal As Currency
    AccName As String
    custno As String
    pic As String
    sign As String
End Type
Private Type faInfo
    minCall As Currency
    minFixed As Currency
End Type
Private Type shareinfo
    memberno As String
    totalshares As Currency
    
End Type
Private Type loansinfo
    MemNo As String
    Loanno As String
    LoanAmount As Currency
    repayperiod As Integer
End Type
Private Type sinfo
   meberno As String
   transdate As Date
   totalshares As Currency
End Type
Private Type tellerInfo
    tellerName As String
    tellerCubicle  As String
    tellerCurrBal As Currency
    tellerMaxBal As Currency
    tellerReplenish As Currency
    amtPayManager As Currency
    amtPaySuper As Currency
    amtPayTeller As Currency
    ttype As Byte '0 for not known, 1 for manager, 2 for Super, 3 for Teller
End Type



Private Type saInfo
    withLmt As Currency
    withInt As Integer
    FOSATarriffGuide As Currency
    lessThanWithIntCharge As Currency
    minBal As Currency
    lessThanMinBalCharge As Currency
    withCharge As Currency
    bankerscheque As Currency
    group3 As Currency
    individual3 As Currency
    over30 As String
    amoutover As Currency
    mobile As Currency
    intonauthorisedod As Integer
    intonunauthorisedod As Integer
    intonclearedchqs As Integer
    intonloans As Integer
    staffcode As String
    STAMPDUTY As Currency
End Type

Private Type transInfo
    accType As Byte   'the Type of account
                            '            1 is fixed call account,
                            '            2 is fixed term account,
                            '            3 is normal account,
    transAmt As Currency
    Month As Byte
    desc As String
    tdate As Date
    openFee As Currency
    ACCNO As String
    idno As String
    PAYNO As String
    custno As String
    availbal As Currency
    AccName As String
    dateValid As Boolean
    fTrans As Boolean
    lastWithDate As Date
    vno As String
End Type
Private Type loanbalinfo
 loan As String
 memberno As String
 balance As String
 repayrate As Integer ' principal
 lastdate As Date
 interest As Double '% percenatage
 repaymethod As String ' either stl,rbal ,amrt
 repayperiod   As Integer 'period in months
End Type
Private loanbal As loanbalinfo
Private Loans As loansinfo
Private tos As sinfo
Private share As shareinfo
Private sa As saInfo
Private fa As faInfo
Private tInfo As transInfo
Private accData() As accoInfo
Private transtype As Byte '1 for deposit, withdrawal for 2,3 for loan repayment ,4 for share contribution

Private Sub cboyear_Change()
'    With Lvwcategory
'        .ListItems.Clear
'        .Columnheaders.Clear
'    End With
'    Set rs = CreateObject("adodb.recordset")
'    sql = "Select Accno, mmonth, yyear, Actual, Budgetted, Variance from budgets where yyear=" & cboyear & " order by accno"
'    Set rs = CreateObject("adodb.recordset")
'    Set clsClass = New cdbase
'    Provider = clsClass.OpenCon
'    Set cn = CreateObject("adodb.connection")
'   cn.Open Provider, "atm","atm"
'    rs.Open sql, cn
'    With Lvwcategory
'        .Columnheaders.Add , , "Account Number"
'        .Columnheaders.Add , , "Year"
'        .Columnheaders.Add , , "Actuals"
'        .Columnheaders.Add , , "Budgetted"
'        .Columnheaders.Add , , "Variance"
'        While Not rs.EOF
'            Set Li = .ListItems.Add(, , Trim(rs.Fields("accno")))
'            Li.ListSubItems.Add , , Trim(rs.Fields("YYear"))
'            Li.ListSubItems.Add , , Trim(rs.Fields("Actual"))
'            Li.ListSubItems.Add , , Trim(rs.Fields("Budgetted"))
'            Li.ListSubItems.Add , , Trim(rs.Fields("Variance"))
'            rs.MoveNext
'        Wend
'    End With
'    rs.Close
'    Set rs = Nothing
'    Lvwcategory.View = lvwReport
End Sub

Private Sub cboyear_Click()
cboyear_Change
End Sub

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub Save_My_Budget()
    On Error GoTo SysError
    Dim I As Long, rsBudget As New Recordset
    If Trim$(LBLGLCONTRA) = "" Then
        MsgBox "Please supply the Account No", vbInformation, Me.Caption
        LBLGLCONTRA.SetFocus
        SendKeys "{Home}+{End}"
        Exit Sub
    End If
    Set rsBudget = oSaccoMaster.GetRecordset("Select AccNo From GLSETUP where AccNo='" _
    & LBLGLCONTRA & "'")
    With rsBudget
        If .State = adStateOpen Then
            If Not .EOF Then
                Set rsBudget = oSaccoMaster.GetRecordset("Delete From BUDGETS where AccNo='" & LBLGLCONTRA & "' and yYear=" & Year(dtpBudgetYear))
              If txtmonthd = "No" Then
                For I = 1 To 12
                    Set li = lvwBudget.ListItems(1)
                    If Not Save_The_Budget(LBLGLCONTRA, I, Year(dtpBudgetYear), _
                    CDbl(li.SubItems(2)), ErrorMessage) Then
                        If ErrorMessage <> "" Then
                            MsgBox ErrorMessage, vbInformation, Me.Caption
                            ErrorMessage = ""
                        End If
                    End If
                Next
              Else
               For I = 1 To 1
                    Set li = lvwBudget.ListItems(1)
                    If Not Save_The_Budget(LBLGLCONTRA, lvwBudget.ListItems(1), Year(dtpBudgetYear), _
                    CDbl(li.SubItems(2)), ErrorMessage) Then
                        If ErrorMessage <> "" Then
                            MsgBox ErrorMessage, vbInformation, Me.Caption
                            ErrorMessage = ""
                        End If
                    End If
                Next
              End If
            Else
                MsgBox "Account No " & LBLGLCONTRA & " not found in the Chart Of Accounts", vbInformation, Me.Caption
                LBLGLCONTRA.SetFocus
                SendKeys "{Home}+{End}"
                Exit Sub
            End If
        End If
    End With
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub cmdExport_Click()
    On Error GoTo SysError
    Dim FileName As String, MFSO As New FileSystemObject, BudgetFile As TextStream, _
    strData As String
    If Lvwcategory.ListItems.Count = 0 Then
        MsgBox "There are no records to be Exported", vbInformation, Me.Caption
        Exit Sub
    End If
    With CommonDialog1
        .Filter = "Excel Files|*.csv"
        .ShowSave
        If .FileName <> "" Then
            FileName = .FileName
        End If
    End With
    If FileName <> "" Then
        If MFSO.FileExists(FileName) Then
            If MsgBox("The specified file exists. Do you want to overwrite it?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
                Exit Sub
            End If
        End If
        Set BudgetFile = MFSO.CreateTextFile(FileName, True)
        strData = ",Budgetted Vs Actual Amounts for " & Format(dtpBudgetPeriod, "MMMM yyyy")
        BudgetFile.WriteLine strData
        strData = ""
        BudgetFile.WriteLine strData
        strData = "AccNo,Account Name,Budgetted Amount,Actual Amount,Variance,Percentage Variance"
        BudgetFile.WriteLine strData
        strData = ""
        For I = 1 To Lvwcategory.ListItems.Count
            Set li = Lvwcategory.ListItems(I)
            strData = li & "," & li.SubItems(1) & "," & CDbl(li.SubItems(2)) & "," & CDbl(li.SubItems(3)) _
            & "," & CDbl(li.SubItems(4)) & "," & CDbl(li.SubItems(5))
            BudgetFile.WriteLine strData
            strData = ""
        Next I
        Set MFSO = Nothing
        BudgetFile.Close
        Set BudgetFile = Nothing
        MsgBox "Transfer Completed Successfully", vbInformation, Me.Caption
    Else
        MsgBox "There is no specified file to save", vbExclamation, Me.Caption
        Exit Sub
    End If
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub cmdLoad_Click()
    On Error GoTo SysError
    Dim rsBudgets As New Recordset, rsActual As New Recordset, sAccNo As String, PAmount As Double
    Lvwcategory.ListItems.Clear
    Set rsBudgets = oSaccoMaster.GetRecordset("Set DateFormat DMY Select B.*,A.GLAccName From " _
    & "BUDGETS B inner join GLSETUP A on B.AccNo=A.AccNo where " _
    & "mMonth=" & Month(dtpBudgetPeriod) & " and yYear=" & Year(dtpBudgetPeriod) & " order by " _
    & "B.AccNo")
    With rsBudgets
        While Not .EOF
            sAccNo = IIf(IsNull(!ACCNO), "", !ACCNO)
            sAccNo = Trim$(sAccNo)
            Set li = Lvwcategory.ListItems.Add(, , IIf(IsNull(!ACCNO), "", !ACCNO))
            li.SubItems(1) = IIf(IsNull(!GlAccName), "", !GlAccName)
            Set rsActual = oSaccoMaster.GetRecordset("Get_Actual_Amount '" & sAccNo & "'," _
            & Month(dtpBudgetPeriod) & "," & Year(dtpBudgetPeriod))
            With rsActual
                If .State = adStateOpen Then
                    If Not .EOF Then
                        li.SubItems(3) = Format(IIf(IsNull(!amount), 0, !amount), Cfmt)
                    Else
                        li.SubItems(3) = "0.00"
                    End If
                Else
                    li.SubItems(3) = "0.00"
                End If
            End With
            li.SubItems(2) = Format(IIf(IsNull(!Budgetted), 0, !Budgetted), Cfmt)
            'Li.SubItems(3) = "0.00"
            li.SubItems(4) = Format(CDbl(li.SubItems(2)) - CDbl(li.SubItems(3)), Cfmt)
            If CDbl(li.SubItems(4)) >= 0 Then
                li.SubItems(5) = Format((CDbl(li.SubItems(4)) / CDbl(li.SubItems(2))) * 100, Cfmt)
            Else
                li.SubItems(5) = Format((CDbl(li.SubItems(3)) / CDbl(li.SubItems(2))) * 100, Cfmt)
            End If
            .MoveNext
        Wend
    End With
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub cmdupdate_Click()
    'SELECT     Accno, mmonth, yyear, Actual, Budgetted, Variance  FROM         budgets
    Save_My_Budget
    MsgBox "Budget Updated Successfully", vbInformation, Me.Caption
    LBLGLCONTRA = ""
    glnamE1 = ""
    txtbudgettedamount = 0
    lvwBudget.ListItems.Clear
    glnamE1.SetFocus
    'SendKeys "{Home}+{End}"
    Exit Sub
    Set myclass = New cdbase
    Set cn = CreateObject("adodb.connection")
    Provider = myclass.OpenCon
   cn.Open Provider, "atm", "atm"
    '//get the last items
    sql = ""
    sql = "select accno from budgets where accno='" & LBLGLCONTRA & "' and yyear=" & CBOYEAR & ""
    Set rs = New ADODB.Recordset
    rs.Open sql, cn, adOpenKeyset, adLockOptimistic
    If rs.EOF Then
        sql = ""
        sql = "set dateformat dmy insert into budgets(Accno, mmonth, yyear, Actual, Budgetted, Variance) values('" & LBLGLCONTRA & "'," & Month(transdate) & "," & CBOYEAR & "," & lblactuals & "," & txtbudgettedamount & "," & CCur(CCur(txtbudgettedamount) - CCur(lblactuals)) & ")"
        cn.Execute sql
        MsgBox "You have successfully added the records", vbInformation
        Exit Sub
    Else
        sql = ""
        sql = " set dateformat dmy UPDATE    budgets  SET actual=" & lblactuals & " ,variance=" & CCur(lblactuals) - txtbudgettedamount & " where accno='" & LBLGLCONTRA & "' and yyear=" & CBOYEAR & ""
        cn.Execute sql
        MsgBox "You have successfully updated the account", vbInformation
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    transdate = Format(Get_Server_Date, "dd/mm/yyyy")
    CBOYEAR = Year(transdate)
    dtpBudgetYear = Get_Server_Date
    dtpBudgetPeriod = Format(DateSerial(Year(dtpBudgetYear), Month(dtpBudgetYear) + 1, 1 - 1), "MM-yyyy")
    '//update the actual balances
    Set myclass = New cdbase
    Set cn = CreateObject("adodb.connection")
    Provider = myclass.OpenCon
   cn.Open Provider, "atm", "atm"
'    sql = "select accno from budgets where yyear=" & cboyear & ""
'    Set rs = New ADODB.Recordset
'    rs.Open sql, cn, adOpenKeyset, adLockOptimistic
'    If Not rs.EOF Then
'    While Not rs.EOF
'    Dim rst As New ADODB.Recordset
'    Dim Bal As Currency
'    '// get the latest balance from cub
'    sql = ""
'    sql = "select availablebalance from cub where accno ='" & rs.Fields(0) & "'"
'    Set rst = New ADODB.Recordset
'    rst.Open sql, cn, adOpenKeyset, adLockOptimistic
'    If Not rst.EOF Then
'    If Not IsNull(rst.Fields(0)) Then Bal = rst.Fields(0) Else Bal = 0
'    '//get the budget for the year
'    Dim rsb As Recordset
'    Set rsb = New ADODB.Recordset
'    Dim b As Currency
'    Dim varian As Currency
'    cboPeriod.ListIndex = 0
'
'    sql = ""
'    sql = "select budgetted from budgets where accno='" & rs.Fields(0) & "' and yyear=" & cboyear & ""
'    rsb.Open sql, cn, adOpenKeyset, adLockOptimistic
'    If Not rsb.EOF Then
'    If Not IsNull(rsb.Fields(0)) Then b = rsb.Fields(0) Else b = 0
'    End If
'    '// get the variances
'    varian = b - Bal
'    '//update the budget stuff
'    sql = ""
'    sql = "set dateformat dmy update budgets set actual=" & Bal & ",variance=" & varian & " where accno='" & rs.Fields(0) & "' and yyear=" & cboyear & ""
'    cn.Execute sql
'
'    End If
'    rs.MoveNext
'    Wend
'    End If
'    '//get the statistics right
    Dim RsShares As Recordset
    Set RsShares = New ADODB.Recordset
    sql = ""
    sql = "SELECT     TOP 1 *  FROM         SHARES   ORDER BY TotalShares DESC"
    'RsShares.Open sql, cn, adOpenKeyset, adLockOptimistic
    'If Not RsShares.EOF Then
     '   lblhighestsaver = Format(IIf(IsNull(RsShares!totalshares), 0, RsShares!totalshares), Cfmt)
        'If Not IsNull(rsShares.Fields("totalshares")) Then lblhighestsaver = rsShares.Fields("totalshares")
        'If Not IsNull(rsShares.Fields("memberno")) Then memberno = rsShares.Fields("memberno")
      '  memberno = IIf(IsNull(RsShares!memberno), "", RsShares!memberno)
    'End If
    '//lowest savers
    'Dim rsshares As Recordset
    'Set RsShares = New ADODB.Recordset
    'sql = ""
    'sql = "SELECT     TOP 1 *  FROM         SHARES   ORDER BY TotalShares asc"
    'RsShares.Open sql, cn, adOpenKeyset, adLockOptimistic
    'If Not RsShares.EOF Then
     '   lowestsaver = Format(IIf(IsNull(RsShares!totalshares), 0, RsShares!totalshares), Cfmt)
        'If Not IsNull(rsShares.Fields("totalshares")) Then lowestsaver = rsShares.Fields("totalshares")
        'If Not IsNull(rsShares.Fields("memberno")) Then memberno1 = rsShares.Fields("memberno")
      '  memberno1 = IIf(IsNull(RsShares!memberno), "", RsShares!memberno)
    'End If
    
    '//SELECT     SUM(Balance) AS a, MemberNo  FROM         LOANBAL  GROUP BY MemberNo   ORDER BY a DESC
'    Dim RsLoans As Recordset
'    Set RsLoans = New ADODB.Recordset
'    sql = ""
'    sql = "SELECT     SUM(Balance) AS a, MemberNo  FROM         LOANBAL  GROUP BY MemberNo   ORDER BY a DESC"
'    RsLoans.Open sql, cn, adOpenKeyset, adLockOptimistic
'    If Not RsLoans.EOF Then
'        highestloanee = Format(IIf(IsNull(RsLoans!A), 0, RsLoans!A), Cfmt)
'        'If Not IsNull(rsLoans.Fields(0)) Then highestloanee = rsLoans.Fields(0) Else highestloanee = 0
'        'If Not IsNull(rsLoans.Fields(1)) Then memberno3 = rsLoans.Fields(1)
'        memberno3 = IIf(IsNull(RsLoans!memberno), "", RsLoans!memberno)
'    End If
    '//lowest nominee.
    'Dim rsloans As Recordset
'    Set RsLoans = New ADODB.Recordset
'    sql = ""
'    sql = "SELECT     SUM(Balance) AS a, MemberNo  FROM         LOANBAL  GROUP BY MemberNo   ORDER BY a asc"
'    RsLoans.Open sql, cn, adOpenKeyset, adLockOptimistic
'    If Not RsLoans.EOF Then
'        lowestloanee = Format(IIf(IsNull(RsLoans!A), 0, RsLoans!A), Cfmt)
'        'If Not IsNull(rsLoans.Fields(0)) Then lowestloanee = rsLoans.Fields(0)
'        'If Not IsNull(rsLoans.Fields(1)) Then memberno4 = rsLoans.Fields(1)
'        memberno4 = IIf(IsNull(RsLoans!memberno), "", RsLoans!memberno)
'    End If
    '//membership
    
End Sub

Private Sub Get_Budget(ACCNO As String, mYear As Long)
    On Error GoTo SysError
    Dim rsBudget As New Recordset, mMonth As Long, BudgettedAmount As Double
    lvwBudget.ListItems.Clear
    Set rsBudget = oSaccoMaster.GetRecordset("Select * From BUDGETS where AccNo='" & ACCNO _
    & "' and yYear=" & mYear & " order by mMonth")
    With rsBudget
        If .State = adStateOpen Then
            While Not .EOF
                I = I + 1
                mMonth = IIf(IsNull(!mMonth), 1, !mMonth)
                Set li = lvwBudget.ListItems.Add(, , I)
                Enddate = Format(DateSerial(mYear, mMonth + 1, 1 - 1), "dd-MM-yyyy")
                li.SubItems(1) = Enddate
                li.SubItems(2) = Format(IIf(IsNull(!Budgetted), 0, !Budgetted), Cfmt)
                BudgettedAmount = BudgettedAmount + CDbl(li.SubItems(2))
                .MoveNext
            Wend
        End If
        txtbudgettedamount = Format(Round(BudgettedAmount, 0), Cfmt)
    End With
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub glName1_Change()
    On Error GoTo SysError
    Dim rsAccounts As New Recordset
    lvwAccounts.ListItems.Clear
    If Trim$(glnamE1) <> "" Then
        If Not Editing Then
            Set rsAccounts = oSaccoMaster.GetRecordset("Select AccNo,GLAccName From GLSETUP" _
            & " where GLAccName like '%" & glnamE1 & "%'")
            With rsAccounts
                If .State = adStateOpen Then
                    If Not .EOF Then
                        lvwAccounts.Visible = True
                    Else
                        lvwAccounts.Visible = False
                    End If
                    If .RecordCount = 1 Then
                        LBLGLCONTRA = IIf(IsNull(!ACCNO), "", !ACCNO)
                        glnamE1 = IIf(IsNull(!GlAccName), "", !GlAccName)
                        lvwAccounts.Visible = False
                        Get_Budget LBLGLCONTRA, dtpBudgetYear
                        txtbudgettedamount.SetFocus
                        SendKeys "{Home}+{End}"
                        Exit Sub
                    End If
                    While Not .EOF
                        Set li = lvwAccounts.ListItems.Add(, , IIf(IsNull(!ACCNO), "", !ACCNO))
                        li.SubItems(1) = IIf(IsNull(!GlAccName), "", !GlAccName)
                        .MoveNext
                    Wend
                End If
            End With
        End If
    Else
        lvwAccounts.Visible = False
    End If
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub glName1_KeyPress(KeyAscii As Integer)
    On Error GoTo errFix
    If KeyAscii <> vbKeyReturn Then 'Catch the Enter key
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
    Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Member Registration"
End Sub

Private Sub LBLGLCONTRA_Change()
    On Error GoTo SysError
    Dim rsBudget As New Recordset
    If Trim$(LBLGLCONTRA) <> "" Then
        Set rsBudget = oSaccoMaster.GetRecordset("Select * From GLSETUP where AccNo='" & LBLGLCONTRA & "'")
        With rsBudget
            If .State = adStateOpen Then
                If Not .EOF Then
                    Editing = True
                    glnamE1 = IIf(IsNull(!GlAccName), "", !GlAccName)
                    Get_Budget LBLGLCONTRA, Year(dtpBudgetYear)
                    Editing = False
                End If
            End If
        End With
    Else
        glnamE1 = ""
        lvwBudget.ListItems.Clear
    End If
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub lvwAccounts_DblClick()
    On Error GoTo SysError
    Editing = True
    If lvwAccounts.ListItems.Count > 0 Then
        LBLGLCONTRA = lvwAccounts.SelectedItem
        'glName1 = lvwAccounts.SelectedItem.SubItems(1)
        lvwAccounts.Visible = False
        txtbudgettedamount.SetFocus
        SendKeys "{Home}+{End}"
    End If
    Editing = False
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub lvwBudget_DblClick()
    On Error GoTo SysError
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub optFixed_Click()
    lblAmount = "Fixed Amount"
End Sub

Private Sub optSpread_Click()
    lblAmount = "Spread Amount"
End Sub

Private Sub Picture1_Click()
    Dim Z, S, U
    error.Visible = False
    Dim rs As Recordset
     frmSearchGLAccounts.Show vbModal, Me
    If Continue Then
        If SearchValue <> "" Then
            LBLGLCONTRA = SearchValue
            SearchValue = ""
        End If
    End If

    Z = strName
    If Z <> "" Then
        LBLGLCONTRA = Z
        glcomm = Z
        End If
    Set myclass = New cdbase
    Set cn = CreateObject("adodb.connection")
    Provider = myclass.OpenCon
   cn.Open Provider, "atm", "atm"
    sql = "select * from cuB where ACCNO='" & glcomm & "'"
    Set rs = New ADODB.Recordset
    rs.Open sql, cn
    If Not rs.EOF Then
    If Not IsNull(rs.Fields("accno")) Then glaccno = rs.Fields("accno")
    If Not IsNull(rs.Fields("accountname")) Then glnamE1 = rs.Fields("name")
    If Not IsNull(rs.Fields("idno")) Then glidno = rs.Fields("idno")
    If Not IsNull(rs.Fields("memberno")) Then glmemno = rs.Fields("memberno")
    If Not IsNull(rs.Fields("payno")) Then glpayno = rs.Fields("payno")
    If Not IsNull(rs.Fields("availablebalance")) Then lblactuals = rs.Fields("availablebalance") Else lblactuals = 0
    End If
'glPremium = Scheme_GL_Field(AccountCode, "glPremium")
'bookba = cub_balance(glaccno)
'///get the other amount
sql = ""
Dim St As Recordset
sql = ""
    sql = "select Budgetted from budgets where accno='" & LBLGLCONTRA & "'  and yyear=" & CBOYEAR & " order by yyear desc"
    Set St = New ADODB.Recordset
    Set St = oSaccoMaster.GetRecordset(sql)
    If Not St.EOF Then
    If Not IsNull(St.Fields("Budgetted")) Then txtbudgettedamount = St.Fields("Budgetted")
    Else
    error.Visible = True
    txtbudgettedamount = 0
    End If

End Sub

Private Sub txtbudgettedamount_KeyPress(KeyAscii As Integer)
    On Error GoTo SysError
    Select Case KeyAscii
        Case 48 To 57
        Case Is = 46
        Case Is = 8
        Case Else
        KeyAscii = 0
    End Select
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub


Private Sub txtbudgettedamount_LostFocus()
    On Error GoTo SysError
    Dim BudgettedAmount As Double, I As Long, myDate As Date
    lvwBudget.ListItems.Clear
    If Trim$(txtbudgettedamount) <> "" Then
        BudgettedAmount = CDbl(txtbudgettedamount)
    Else
        txtbudgettedamount = 0
    End If
        If txtmonthd = "" Then
     MsgBox "select if its for single month or the whole year", vbInformation, Me.Caption
     txtmonthd.SetFocus
     Exit Sub
    End If
    If txtmonthd = "No" Then
        If optSpread.value = True Then
            If BudgettedAmount > 0 Then
                BudgettedAmount = BudgettedAmount / 12
            End If
        End If
        For I = 1 To 12
            myDate = Format(DateSerial(Year(dtpBudgetYear), I + 1, 1 - 1), "dd-MM-yyyy")
            lvwBudget.ListItems.Add , , I
            lvwBudget.ListItems(I).SubItems(1) = myDate
            lvwBudget.ListItems(I).SubItems(2) = Format(BudgettedAmount, Cfmt)
        Next I
    Else
    Dim m As Integer
        m = Format(DateSerial(Year(dtpBudgetYear), Month(dtpBudgetYear) + 1, 1 - 1), "MM")
          For I = 1 To 1
            myDate = Format(DateSerial(Year(dtpBudgetYear), Month(dtpBudgetYear) + 1, 1 - 1), "dd-MM-yyyy")
            lvwBudget.ListItems.Add , , m
            lvwBudget.ListItems(I).SubItems(1) = myDate
            lvwBudget.ListItems(I).SubItems(2) = Format(BudgettedAmount, Cfmt)
         Next I
    End If
    txtbudgettedamount = Format(txtbudgettedamount, Cfmt)
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub
