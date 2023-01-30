VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmOutletsale 
   BackColor       =   &H00FFFF00&
   Caption         =   "OUTLET STOCK AND SALES"
   ClientHeight    =   9480
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11400
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9480
   ScaleWidth      =   11400
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10200
      TabIndex        =   31
      Top             =   6240
      Width           =   975
   End
   Begin MSComCtl2.DTPicker txtdateenterered 
      Height          =   375
      Left            =   9480
      TabIndex        =   28
      Top             =   720
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   174850049
      CurrentDate     =   38814
   End
   Begin VB.ComboBox cbobranch 
      Height          =   315
      ItemData        =   "frmOutletsale.frx":0000
      Left            =   8640
      List            =   "frmOutletsale.frx":0002
      TabIndex        =   45
      Top             =   1800
      Width           =   2535
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9495
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   16748
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      BackColor       =   16777088
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "STOCK TO OUTLET"
      TabPicture(0)   =   "frmOutletsale.frx":0004
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "OUTLET SALES"
      TabPicture(1)   =   "frmOutletsale.frx":0020
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Vehicle Milk Dispatch"
      TabPicture(2)   =   "frmOutletsale.frx":003C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame7"
      Tab(2).Control(1)=   "Frame6"
      Tab(2).Control(2)=   "Frame5"
      Tab(2).Control(3)=   "ListView200"
      Tab(2).ControlCount=   4
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Statements"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   -67680
         TabIndex        =   109
         Top             =   3240
         Width           =   3855
         Begin VB.CommandButton cmddailys 
            Caption         =   "Daily Sales Report"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2640
            TabIndex        =   114
            Top             =   1560
            Width           =   1095
         End
         Begin VB.CommandButton cmdincomdairy 
            Caption         =   "Dairy Income Statement"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   113
            Top             =   1560
            Width           =   2415
         End
         Begin VB.CommandButton cmdParchase 
            Caption         =   "Process Sales statement"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   112
            Top             =   240
            Width           =   2775
         End
         Begin VB.CommandButton cmdpareport 
            Caption         =   "Purchase Report"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   111
            Top             =   960
            Width           =   1575
         End
         Begin VB.CommandButton cmdSalessta 
            Caption         =   "Sales Report"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1800
            TabIndex        =   110
            Top             =   960
            Width           =   1335
         End
         Begin VB.Line Line1 
            BorderWidth     =   3
            X1              =   120
            X2              =   3720
            Y1              =   840
            Y2              =   840
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "This Month Milk Dispatch"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3615
         Left            =   -74760
         TabIndex        =   87
         Top             =   5400
         Width           =   9735
         Begin MSComctlLib.ListView ListView30 
            Height          =   3135
            Left            =   120
            TabIndex        =   89
            Top             =   360
            Width           =   9495
            _ExtentX        =   16748
            _ExtentY        =   5530
            View            =   3
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            AllowReorder    =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            HotTracking     =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   65280
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Vehicle Milk Dispatch Details"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   2775
         Left            =   -74880
         TabIndex        =   78
         Top             =   480
         Width           =   11055
         Begin VB.CommandButton cmdvehreport 
            Caption         =   "Dispatch Report"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   8880
            TabIndex        =   99
            Top             =   1800
            Width           =   1935
         End
         Begin VB.CommandButton cmdnext1 
            Caption         =   "Next"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   5760
            TabIndex        =   91
            Top             =   1800
            Width           =   1335
         End
         Begin VB.CommandButton cmdremov 
            Caption         =   "Remove"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   4440
            TabIndex        =   90
            Top             =   1800
            Width           =   1215
         End
         Begin VB.CommandButton cmdMilkVeh 
            Caption         =   "Save"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   7320
            TabIndex        =   88
            Top             =   1800
            Width           =   1455
         End
         Begin VB.TextBox txtMilkd 
            Height          =   495
            Left            =   2280
            TabIndex        =   85
            Top             =   1200
            Width           =   2055
         End
         Begin VB.ComboBox cbov 
            Height          =   315
            Left            =   2280
            TabIndex        =   84
            Top             =   600
            Width           =   3375
         End
         Begin MSComctlLib.ProgressBar prgStatus 
            Height          =   255
            Left            =   7200
            TabIndex        =   117
            Top             =   2520
            Visible         =   0   'False
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   1
            Appearance      =   0
            Scrolling       =   1
         End
         Begin VB.Label txtlbl 
            BackColor       =   &H00FFFF80&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   7560
            TabIndex        =   105
            Top             =   2760
            Width           =   3375
         End
         Begin VB.Label Label36 
            BackColor       =   &H00FFFFC0&
            Height          =   495
            Left            =   2280
            TabIndex        =   86
            Top             =   1800
            Width           =   2055
         End
         Begin VB.Label Label35 
            Caption         =   "Date."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8040
            TabIndex        =   83
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label34 
            Caption         =   "Today's(Kgs)."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   82
            Top             =   1920
            Width           =   1815
         End
         Begin VB.Label Label33 
            Caption         =   "Quantity(Kgs)."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   81
            Top             =   1320
            Width           =   1815
         End
         Begin VB.Label Label32 
            Caption         =   "Vehicle No."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   80
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label Label31 
            Caption         =   "Outlet Name"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8520
            TabIndex        =   79
            Top             =   960
            Width           =   1935
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFC0&
         Height          =   8535
         Left            =   -74880
         TabIndex        =   1
         Top             =   480
         Width           =   11055
         Begin VB.ComboBox cbotxttyp 
            Height          =   315
            ItemData        =   "frmOutletsale.frx":0058
            Left            =   5640
            List            =   "frmOutletsale.frx":006B
            TabIndex        =   118
            Text            =   "Combo1"
            Top             =   720
            Width           =   735
         End
         Begin VB.CheckBox chkyoghurt 
            Caption         =   "is it Yoghurt?"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   5640
            TabIndex        =   116
            Top             =   480
            Width           =   1575
         End
         Begin VB.TextBox txttyp 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   285
            Left            =   6000
            TabIndex        =   115
            Text            =   "0"
            Top             =   720
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CommandButton cmdvehiclenew 
            Caption         =   "New Vehicle"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   7800
            TabIndex        =   107
            Top             =   5760
            Width           =   1575
         End
         Begin VB.CommandButton cmdedit 
            Caption         =   "Edit"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1800
            TabIndex        =   106
            Top             =   5760
            Width           =   1215
         End
         Begin VB.CheckBox chkrepeat 
            Caption         =   "Remove the kgs"
            Height          =   255
            Left            =   6120
            TabIndex        =   103
            Top             =   1680
            Width           =   2295
         End
         Begin VB.ComboBox txtpname 
            Height          =   315
            Left            =   1680
            TabIndex        =   100
            Top             =   720
            Width           =   3495
         End
         Begin VB.ComboBox cbovb 
            Height          =   315
            Left            =   3960
            TabIndex        =   97
            Top             =   3720
            Width           =   2775
         End
         Begin VB.CheckBox chkcustomer 
            Caption         =   "Supply by a Vehicle?"
            Height          =   195
            Left            =   1800
            TabIndex        =   96
            Top             =   3480
            Width           =   2055
         End
         Begin VB.CommandButton cmddisreport 
            Caption         =   "Outlet Milk  Report"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   4800
            TabIndex        =   93
            Top             =   5760
            Width           =   1335
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
            Height          =   315
            Left            =   3360
            TabIndex        =   70
            Top             =   4800
            Visible         =   0   'False
            Width           =   6825
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
            Height          =   315
            Left            =   3360
            TabIndex        =   69
            Top             =   4440
            Visible         =   0   'False
            Width           =   6825
         End
         Begin VB.CommandButton Command1 
            Caption         =   "New Outlet"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   6480
            TabIndex        =   51
            Top             =   5760
            Width           =   1215
         End
         Begin VB.CommandButton cmdnew 
            Caption         =   "New"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   480
            TabIndex        =   16
            Top             =   5760
            Width           =   1095
         End
         Begin VB.CommandButton mm 
            Caption         =   "&Save"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3240
            TabIndex        =   15
            Top             =   5760
            Width           =   1095
         End
         Begin VB.TextBox txtpprice 
            Appearance      =   0  'Flat
            Height          =   405
            Left            =   3120
            TabIndex        =   11
            Top             =   2280
            Width           =   1095
         End
         Begin VB.TextBox txtsellingprice 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   3120
            TabIndex        =   9
            Top             =   2760
            Width           =   1095
         End
         Begin VB.TextBox txtquantity 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   1680
            TabIndex        =   7
            Top             =   1440
            Width           =   1935
         End
         Begin VB.TextBox txtbalance 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   375
            Left            =   3720
            TabIndex        =   6
            Top             =   1440
            Width           =   1575
         End
         Begin VB.TextBox txtpcode 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   1680
            TabIndex        =   3
            Top             =   240
            Width           =   1935
         End
         Begin VB.PictureBox Picture2 
            BorderStyle     =   0  'None
            DrawStyle       =   2  'Dot
            DrawWidth       =   17015
            Height          =   360
            Left            =   3600
            Picture         =   "frmOutletsale.frx":0091
            ScaleHeight     =   360
            ScaleWidth      =   240
            TabIndex        =   2
            Top             =   240
            Width           =   240
         End
         Begin VB.Frame Frame2 
            Caption         =   "SET THE PRICE"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   120
            TabIndex        =   12
            Top             =   2040
            Width           =   10335
            Begin VB.Label Label10 
               Caption         =   "Retail Price "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   240
               TabIndex        =   14
               Top             =   840
               Width           =   1215
            End
            Begin VB.Label Label9 
               Caption         =   "Wholesale Price "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   375
               Left            =   240
               TabIndex        =   13
               Top             =   360
               Width           =   1815
            End
         End
         Begin MSComctlLib.ListView lvWBranch2 
            Height          =   2055
            Left            =   120
            TabIndex        =   61
            Top             =   6360
            Width           =   10815
            _ExtentX        =   19076
            _ExtentY        =   3625
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin VB.Frame Frame4 
            Caption         =   "Outlet ledgers"
            Height          =   855
            Left            =   120
            TabIndex        =   71
            Top             =   4200
            Visible         =   0   'False
            Width           =   10335
            Begin VB.CommandButton Command2 
               Caption         =   "..."
               Height          =   285
               Left            =   1680
               TabIndex        =   75
               Top             =   600
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
               Height          =   315
               Left            =   1995
               TabIndex        =   74
               Top             =   600
               Width           =   1080
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
               Height          =   315
               Left            =   1980
               TabIndex        =   73
               Top             =   240
               Width           =   1080
            End
            Begin VB.CommandButton cmdSearch 
               Caption         =   "..."
               Height          =   285
               Left            =   1680
               TabIndex        =   72
               Top             =   240
               Width           =   300
            End
            Begin VB.Label Label30 
               Caption         =   "Dr Stock"
               BeginProperty Font 
                  Name            =   "Century"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   77
               Top             =   360
               Width           =   1335
            End
            Begin VB.Label Label29 
               Caption         =   "Cr Sales"
               BeginProperty Font 
                  Name            =   "Century"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   76
               Top             =   720
               Width           =   1215
            End
         End
         Begin VB.Label Label39 
            Caption         =   "Select Vehicle"
            Height          =   255
            Left            =   4200
            TabIndex        =   98
            Top             =   3480
            Width           =   2175
         End
         Begin VB.Label Label19 
            Caption         =   "OUTLET NAME"
            Height          =   255
            Left            =   5880
            TabIndex        =   47
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label Label6 
            Caption         =   "Date Entered"
            Height          =   255
            Index           =   0
            Left            =   7920
            TabIndex        =   29
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label3 
            Caption         =   "Stock Balance"
            Height          =   255
            Left            =   3840
            TabIndex        =   10
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label Label4 
            Caption         =   "Quantity"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   1560
            Width           =   1215
         End
         Begin VB.Label Label2 
            Caption         =   "Product Name"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Product Code"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFC0&
         Height          =   8895
         Left            =   120
         TabIndex        =   17
         Top             =   480
         Width           =   11055
         Begin VB.CommandButton cmdnewvehiclz 
            Caption         =   "New Vehicle"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   8160
            TabIndex        =   108
            Top             =   5760
            Width           =   1455
         End
         Begin VB.CommandButton cmdshort 
            Caption         =   "Expenses form"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   6120
            TabIndex        =   104
            Top             =   5760
            Width           =   1815
         End
         Begin VB.CommandButton cmddela 
            Caption         =   "Delete"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   4560
            TabIndex        =   102
            Top             =   5760
            Width           =   1335
         End
         Begin VB.CommandButton cmdoutsalesre 
            Caption         =   "Sales Report"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2400
            TabIndex        =   101
            Top             =   5760
            Width           =   1935
         End
         Begin VB.TextBox TXTCHANGE 
            Height          =   375
            Left            =   9360
            TabIndex        =   68
            Text            =   "0"
            Top             =   5040
            Width           =   1335
         End
         Begin VB.TextBox TXTTOTAL 
            Enabled         =   0   'False
            Height          =   375
            Left            =   9360
            TabIndex        =   65
            Text            =   "0"
            Top             =   3600
            Width           =   1335
         End
         Begin VB.CommandButton cmdremove 
            Caption         =   "Remove"
            Height          =   375
            Left            =   5400
            TabIndex        =   64
            Top             =   3480
            Width           =   1095
         End
         Begin VB.CommandButton cmdnextitem 
            Caption         =   "Next item"
            Default         =   -1  'True
            Height          =   375
            Left            =   7320
            TabIndex        =   62
            Top             =   3480
            Width           =   1095
         End
         Begin VB.TextBox txtpaybill 
            Height          =   285
            Left            =   1200
            TabIndex        =   54
            Top             =   3600
            Width           =   3495
         End
         Begin VB.TextBox txtAmount 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   405
            Left            =   9360
            TabIndex        =   53
            Top             =   4320
            Width           =   1335
         End
         Begin VB.TextBox Text2 
            ForeColor       =   &H8000000D&
            Height          =   405
            Left            =   1680
            TabIndex        =   48
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Frame fra1 
            BorderStyle     =   0  'None
            Height          =   975
            Left            =   1440
            TabIndex        =   34
            Top             =   2280
            Width           =   8175
            Begin VB.TextBox txtcracc 
               Height          =   375
               Left            =   1680
               TabIndex        =   38
               Top             =   600
               Width           =   6375
            End
            Begin VB.TextBox txtdracc 
               Height          =   375
               Left            =   1680
               TabIndex        =   37
               Top             =   120
               Width           =   6375
            End
            Begin VB.PictureBox Picture3 
               Height          =   255
               Left            =   1320
               Picture         =   "frmOutletsale.frx":0213
               ScaleHeight     =   195
               ScaleWidth      =   195
               TabIndex        =   36
               Top             =   720
               Width           =   255
            End
            Begin VB.PictureBox Picture1 
               Height          =   255
               Left            =   1320
               Picture         =   "frmOutletsale.frx":0ADD
               ScaleHeight     =   195
               ScaleWidth      =   195
               TabIndex        =   35
               Top             =   240
               Width           =   255
            End
            Begin VB.Label lblcracc 
               BackColor       =   &H8000000E&
               Height          =   375
               Left            =   120
               TabIndex        =   40
               Top             =   600
               Width           =   1095
            End
            Begin VB.Label lbldracc 
               BackColor       =   &H8000000E&
               Height          =   375
               Left            =   120
               TabIndex        =   39
               Top             =   120
               Width           =   1095
            End
         End
         Begin VB.PictureBox Picture4 
            BorderStyle     =   0  'None
            DrawStyle       =   2  'Dot
            DrawWidth       =   17015
            Height          =   360
            Left            =   3240
            Picture         =   "frmOutletsale.frx":13A7
            ScaleHeight     =   360
            ScaleWidth      =   240
            TabIndex        =   27
            Top             =   360
            Width           =   240
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   1680
            TabIndex        =   25
            Top             =   360
            Width           =   1575
         End
         Begin VB.CheckBox chkretail 
            Caption         =   "Retail Sale"
            BeginProperty Font 
               Name            =   "Britannic Bold"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3480
            TabIndex        =   23
            Top             =   1560
            Width           =   1335
         End
         Begin VB.CheckBox chkWhole 
            Caption         =   "Wholesale"
            BeginProperty Font 
               Name            =   "Britannic Bold"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   1680
            TabIndex        =   22
            Top             =   1560
            Width           =   1215
         End
         Begin VB.ComboBox cboproductname1 
            Height          =   315
            Left            =   1680
            TabIndex        =   20
            Top             =   720
            Width           =   4455
         End
         Begin VB.CommandButton cmdsave1 
            Caption         =   "Save"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1200
            TabIndex        =   19
            Top             =   5760
            Width           =   975
         End
         Begin VB.CommandButton cmdnew2 
            Caption         =   "New"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   18
            Top             =   5760
            Width           =   855
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   2415
            Left            =   120
            TabIndex        =   60
            Top             =   6360
            Width           =   10815
            _ExtentX        =   19076
            _ExtentY        =   4260
            View            =   3
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            AllowReorder    =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            HotTracking     =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   65280
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin MSComctlLib.ListView Lvwitems 
            Height          =   1695
            Left            =   120
            TabIndex        =   63
            Top             =   3960
            Width           =   9015
            _ExtentX        =   15901
            _ExtentY        =   2990
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
            MousePointer    =   4
            MouseIcon       =   "frmOutletsale.frx":1529
            NumItems        =   6
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "ITEM"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Name"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "QNTY"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "AMOUNT"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Remarks"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Mpesa/Cash"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label Label38 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   4320
            TabIndex        =   95
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label37 
            Caption         =   "Receive:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3360
            TabIndex        =   94
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label28 
            Caption         =   "Change"
            Height          =   255
            Left            =   9360
            TabIndex        =   67
            Top             =   4800
            Width           =   855
         End
         Begin VB.Label Label27 
            Caption         =   "Amount"
            Height          =   255
            Left            =   9360
            TabIndex        =   66
            Top             =   3360
            Width           =   855
         End
         Begin VB.Label Label24 
            Caption         =   "Cr."
            Height          =   255
            Left            =   8520
            TabIndex        =   57
            Top             =   1920
            Width           =   495
         End
         Begin VB.Label Label26 
            Height          =   255
            Left            =   9000
            TabIndex        =   59
            Top             =   1920
            Width           =   1215
         End
         Begin VB.Label Label25 
            Height          =   255
            Left            =   9000
            TabIndex        =   58
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label Label23 
            Caption         =   "Dr."
            Height          =   255
            Left            =   8520
            TabIndex        =   56
            Top             =   1680
            Width           =   495
         End
         Begin VB.Label Label22 
            Caption         =   "Paybill Code"
            Height          =   255
            Left            =   1680
            TabIndex        =   55
            Top             =   3360
            Width           =   1455
         End
         Begin VB.Label Label21 
            Caption         =   "Paid"
            Height          =   255
            Left            =   9360
            TabIndex        =   52
            Top             =   4080
            Width           =   855
         End
         Begin VB.Label Label20 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   6720
            TabIndex        =   50
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label Label12 
            Caption         =   "Quantity:"
            Height          =   375
            Left            =   120
            TabIndex        =   49
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label Label14 
            Caption         =   "OUTLET NAME"
            Height          =   255
            Left            =   8520
            TabIndex        =   46
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label18 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   7680
            TabIndex        =   44
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label17 
            Caption         =   "Quantity Bal"
            Height          =   255
            Left            =   6600
            TabIndex        =   43
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Height          =   375
            Left            =   3720
            TabIndex        =   42
            Top             =   1920
            Width           =   855
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Height          =   375
            Left            =   1800
            TabIndex        =   41
            Top             =   1920
            Width           =   855
         End
         Begin VB.Label Label8 
            Caption         =   "Cr Outlet Stock"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   2880
            Width           =   1335
         End
         Begin VB.Label Label13 
            Caption         =   "Dr Bank"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   2400
            Width           =   1335
         End
         Begin VB.Label Label6 
            Caption         =   "Date Entered"
            Height          =   255
            Index           =   1
            Left            =   7920
            TabIndex        =   30
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label11 
            Caption         =   "Product Code:"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label7 
            Caption         =   "Amount:"
            Height          =   375
            Left            =   5760
            TabIndex        =   24
            Top             =   1200
            Width           =   855
         End
         Begin VB.Label Label5 
            Caption         =   "Product Name:"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   720
            Width           =   1335
         End
      End
      Begin MSComctlLib.ListView ListView200 
         Height          =   2175
         Left            =   -74760
         TabIndex        =   92
         Top             =   3240
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   3836
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
         MousePointer    =   4
         MouseIcon       =   "frmOutletsale.frx":168B
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Vehicle No."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Outlet Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Quantity"
            Object.Width           =   2540
         EndProperty
      End
   End
End
Attribute VB_Name = "frmOutletsale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As String
Dim k As Integer
Dim nr As Integer

Private Sub cbobranch_Click()
    Set rs = New Recordset
    'Dim cn As Connection
    Set cn = New ADODB.Connection
    Provider = "MAZIWA"
    'provider = cn
   cn.Open Provider, "atm", "atm"
    Set rs = New Recordset
    sql = "Select BName1,Dr, Cr from d_Outletbranch where BName1='" & cbobranch & "'"
    rs.Open sql, cn, adOpenKeyset, adLockOptimistic
    While Not rs.EOF
    'cbobranch.AddItem rs.Fields(0)
    Label25 = rs.Fields(1)
    lblcracc = rs.Fields(1)
    Label26 = rs.Fields(2)
    lbldracc = "1003"
    rs.MoveNext
    Wend
   Label23.Visible = True
   Label24.Visible = True
   Label25.Visible = True
   Label26.Visible = True
   branames
   bra
End Sub
Private Sub bra()
    txtpname.Clear
    Set rst = New Recordset
    Dim cn As Connection
    Set cn = New ADODB.Connection
    Provider = "MAZIWA"
    'Provider = cnn
    cn.Open Provider, "atm", "atm"
    Set rst = New Recordset
    sql = "Select distinct(p_name) from d_Outlet where Branch='" & cbobranch & "' ORDER BY p_name"
    rst.Open sql, cn, adOpenKeyset, adLockOptimistic
    While Not rst.EOF
    txtpname.AddItem rst.Fields(0)
    rst.MoveNext
    Wend

End Sub
Private Sub branames()
    cboproductname1.Clear
    Set rst = New Recordset
    Dim cn As Connection
    Set cn = New ADODB.Connection
    Provider = "MAZIWA"
    'Provider = cnn
    cn.Open Provider, "atm", "atm"
    Set rst = New Recordset
    sql = "Select distinct(p_name) from d_Outlet where Branch='" & cbobranch & "' ORDER BY p_name"
    rst.Open sql, cn, adOpenKeyset, adLockOptimistic
    While Not rst.EOF
    cboproductname1.AddItem rst.Fields(0)
    rst.MoveNext
    Wend

End Sub
Private Sub cboproductname1_Click()
If cbobranch = "" Then
MsgBox "Please select the Outlet", vbInformation
Exit Sub
End If
Startdate = DateSerial(year(txtdateenterered), month(txtdateenterered), 1)
Enddate = DateSerial(year(txtdateenterered), month(txtdateenterered) + 1, 1 - 1)

Set rstf = New ADODB.Recordset

Provider = "MAZIWA"
Set cn = New ADODB.Connection
cn.Open Provider, "atm", "atm"
Set rst = New ADODB.Recordset
rst.Open sql, cn
'If rs.EOF Then
Set rst = oSaccoMaster.GetRecordset("set dateformat dmy select p_code, p_name,Wprice, Rprice,Qout from d_Outlet where p_name ='" & cboproductname1 & "' and branch ='" & cbobranch & "' and Date_Entered ='" & txtdateenterered & "'")
If Not rst.EOF Then
Text1.Text = rst.Fields("p_code")
cboproductname1 = rst.Fields("p_name")
Label15 = rst.Fields("Wprice")
Label16 = rst.Fields("Rprice")
Set rstf = oSaccoMaster.GetRecordset("set dateformat dmy select isnull(sum(Qout),0) from d_Outlet where p_name ='" & cboproductname1 & "' and branch ='" & cbobranch & "' and Date_Entered>='" & Startdate & "'and Date_Entered<='" & Enddate & "'")
Label18 = rstf.Fields(0)
'txtsel
Else
Set rst2 = New ADODB.Recordset
Set rst2 = oSaccoMaster.GetRecordset("set dateformat dmy select p_code, p_name,Wprice, Rprice,Qout from d_Outlet where p_name ='" & cboproductname1 & "' and branch ='" & cbobranch & "' ")
 Text1.Text = rst2.Fields("p_code")
 Set rstf5 = New ADODB.Recordset
 Set rstf5 = oSaccoMaster.GetRecordset("set dateformat dmy select isnull(sum(Qout),0) from d_Outlet where p_name ='" & cboproductname1 & "' and branch ='" & cbobranch & "' and Date_Entered>='" & Startdate & "'and Date_Entered<='" & Enddate & "'")
 Label18 = rstf5.Fields(0)
 Label38 = 0
  MsgBox "Please you have not enter any milk for this product today", vbInformation
End If
'chkretail.value = 1
kgs
'kgs2
Text2.SetFocus
End Sub
Private Sub kgs()
Provider = "MAZIWA"
Set cn = New ADODB.Connection
cn.Open Provider, "atm", "atm"
Set rst = New ADODB.Recordset
rst.Open sql, cn
'If rs.EOF Then
Set rst = oSaccoMaster.GetRecordset("select Quantity from d_Outletstock where p_name ='" & cboproductname1 & "' and OutletName ='" & cbobranch & "' and Date_Entered='" & txtdateenterered & "'")
If Not rst.EOF Then
Label38 = rst.Fields("Quantity")
'txtsel
End If
End Sub
Private Sub kgs2()
Provider = "MAZIWA"
Set cn = New ADODB.Connection
cn.Open Provider, "atm", "atm"
Set rst = New ADODB.Recordset
rst.Open sql, cn
'If rs.EOF Then
Set rst = oSaccoMaster.GetRecordset("select Quantity from d_Outletstock where p_name ='" & cboproductname1 & "' and OutletName ='" & cbobranch & "' and Date_Entered='" & txtdateenterered & "'")
If Not rst.EOF Then
Label38 = rst.Fields("Quantity")
'txtsel
End If
End Sub
Private Sub cbov_Click()
 'loaddispmilk
' NAMES1
txtMilkd.SetFocus
End Sub
Private Sub NAMES1()
'Private Sub SSTab1_DblClick()
    cbov.Clear
    Set rst = New Recordset
    Dim cn As Connection
    Set cn = New ADODB.Connection
    Provider = "MAZIWA"
    'Provider = cnn
    cn.Open Provider, "atm", "atm"
    Set rst = New Recordset
    sql = "Select distinct(Locations) from   d_Debtors order by Locations"
    'Select distinct(Locations) from   d_Debtors
    rst.Open sql, cn, adOpenKeyset, adLockOptimistic
    While Not rst.EOF
    cbov.AddItem rst.Fields(0)
    rst.MoveNext
    Wend
End Sub
Private Sub NAMES3()
'Private Sub SSTab1_DblClick()
    cbovb.Clear
    Set rst = New Recordset
    Dim cn As Connection
    Set cn = New ADODB.Connection
    Provider = "MAZIWA"
    'Provider = cnn
    cn.Open Provider, "atm", "atm"
    Set rst = New Recordset
    sql = "Select distinct(Locations) from   d_Debtors order by Locations"
    'Select distinct(Locations) from   d_Debtors
    rst.Open sql, cn, adOpenKeyset, adLockOptimistic
    While Not rst.EOF
    cbovb.AddItem rst.Fields(0)
    rst.MoveNext
    Wend
End Sub
Public Sub loaddispmilk()
     
    With ListView30
    
        .ListItems.Clear
        
        .ColumnHeaders.Clear
    
    End With

    Set rs2 = CreateObject("adodb.recordset")
    
    sql = "set dateformat dmy Select * from  d_OutletDispatch where Date='" & txtdateenterered.value & "'"
    
    Set rs2 = CreateObject("adodb.recordset")
    
    Set clsClass = New cdbase
    
    Provider = clsClass.OpenCon
    
    Set cn = CreateObject("adodb.connection")
    
   cn.Open Provider, "atm", "atm"
    
    rs2.Open sql, cn
    
    With ListView30
    ', , ,
        
        .ColumnHeaders.Add , , "Date"
        .ColumnHeaders.Add , , "Vehicle"
        .ColumnHeaders.Add , , "OutletName"
        .ColumnHeaders.Add , , "Quantity"
        While Not rs2.EOF
        'Code, Name, Date, Quantity, Price, Amount, APaid
            Set li = .ListItems.Add(, , Trim(rs2.Fields("Date")))
            li.ListSubItems.Add , , Trim(rs2.Fields("Vehicle"))
            li.ListSubItems.Add , , Trim(rs2.Fields("OutletName"))
            li.ListSubItems.Add , , Trim(rs2.Fields("Quantity"))
            rs2.MoveNext
        
        Wend
        
    End With
    
    rs2.Close
    
    Set rs2 = Nothing
    
ListView30.View = lvwReport

End Sub

Private Sub chkcustomer_Click()

If chkcustomer = 1 Then
cbovb.Visible = True
Label39.Visible = True
Else
cbovb.Visible = False
Label39.Visible = False
End If
End Sub

Private Sub chkrepeat_Click()
If chkrepeat = 1 Then
'chkrepeat.value = True
Else
chkrepeat = 0
End If
End Sub

Private Sub chkretail_Click()
If chkretail = 1 Then
chkWhole.Visible = False
Label15.Visible = False
If chkretail = 1 Then
If Text2 = "" Then
MsgBox "Please insert quantity", vbInformation
Exit Sub
End If

Label20 = Label16 * Text2
'Label20 = Label20 * 10
a = "Retail sales"
End If
Else
chkWhole.Visible = True
Label15.Visible = True
Label20 = ""
End If
End Sub

Private Sub chkWhole_Click()
If chkWhole = 1 Then
chkretail.Visible = False
Label16.Visible = False
If chkWhole = 1 Then

If Text2 = "" Then
MsgBox "Please insert quantity", vbInformation
Exit Sub
End If

Label20 = Label15 * Text2
a = "Whole sales"
End If
Else
chkretail.Visible = True
Label16.Visible = True
Label20 = ""
End If
End Sub

Private Sub chkyoghurt_Click()
If chkyoghurt = 1 Then
cbotxttyp.Visible = True
'Label39.Visible = True
Else
cbotxttyp.Visible = False
'Label39.Visible = False
End If
End Sub

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmddailys_Click()
Dim ans As String
ans = MsgBox("Do you Want a Report as per price??", vbYesNo)
If ans = vbYes Then
reportname = "SALES PER DAY.rpt"
Else
reportname = "dailysales.rpt"
End If
Show_Sales_Crystal_Report "", reportname, ""
End Sub

Private Sub cmddela_Click()
On Error GoTo ErrorHandler
  
  sql = ""
  sql = "set dateformat dmy delete from d_OutletSales where PCode='" & Text1 & "' and PName='" & cboproductname1 & "' and Date='" & txtdateenterered.value & "' and Description='" & a & "'"
  cn.Execute sql
  '//XXXXXXXXXXXXXXX
    '********** credit agent ledger sale and stock
    sql = ""
    sql = "set dateformat dmy delete from gltransactions where transdate='" & txtdateenterered.value & "' and source='" & Text2 & "' and transdescript ='SALES ON- " & "" & cbobranch & "'"
    oSaccoMaster.ExecuteThis (sql)
    '********** credit agent ledger and bank
''    sql = ""
''    sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & txtdateenterered & "'," & Lvwitems.SelectedItem.SubItems(3) & ",'" & lbldracc & "','" & Label25 & "','','" & Lvwitems.SelectedItem.SubItems(2) & "' ,'SALES ON- " & "" & cbobranch & "','" & User & "',0,0)"
''    oSaccoMaster.ExecuteThis (sql)
'XXXXXXXXXXXXXXXXXXXXXX
  
 '********************************************************************end

Dim rsinstock As Recordset
sql = ""
sql = "select p_code, p_name, Date_Entered,Qout from d_Outlet where p_code= '" & Text1 & "' and p_name= '" & cboproductname1 & "' AND  Branch='" & cbobranch & "' and Date_Entered='" & txtdateenterered & "'"
Set rsinstock = oSaccoMaster.GetRecordset(sql)
'Dim Qout As Integer
'Qout = rsinstock!Qout
sql = "set dateformat DMY Update d_Outlet SET Qout =" & rsinstock.Fields("Qout") + Text2 & " WHERE p_code= '" & Text1 & "' and Branch='" & cbobranch & "' and Date_Entered='" & txtdateenterered & "'"
cn.Execute sql
'*****************************************************Qout =" & rsinstock2.Fields(0) - Text2 & "
MsgBox "Records Deleted succesfully."
loadBranchesTypes
SSTab1_DblClick
SSTab2_DblClick
chkretail.value = vbUnchecked
chkWhole.value = vbUnchecked
Text1.Text = ""
cboproductname1.Text = ""
Label20 = ""
cbobranch.Text = ""
Label18 = ""
TXTCHANGE.Text = ""
TXTTOTAL.Text = ""
k = 0
'txtCustName.Text = ""
Lvwitems.ListItems.Clear
'Text2.Text = ""
'txtAmount.Text = ""
txtpaybill = "CASH PAYMENT"
cmdnew2.Enabled = True
cmdsave1.Enabled = False
cmdnew2_Click

Exit Sub
ErrorHandler:
MsgBox err.description

End Sub

Private Sub cmddisreport_Click()
reportname = "Outletdispatch Report.rpt"
Show_Sales_Crystal_Report "", reportname, ""
End Sub

Private Sub cmdedit_Click()
nr = 0
txtsellingprice.Locked = False
txtpprice.Locked = False
txtquantity.Locked = False
txtpname.Locked = False
txtbalance.Locked = False
End Sub

Private Sub cmdincomdairy_Click()
On Error GoTo ErrorHandler
Dim rst, rstg, rsa As Recordset
Startdate = DateSerial(year(txtdateenterered), month(txtdateenterered), 1)
Enddate = DateSerial(year(txtdateenterered), month(txtdateenterered) + 1, 1 - 1)
sql = ""
sql = "set dateformat dmy delete from d_incomestate where Date >= '" & Startdate & "' And Date<='" & Enddate & "'"
cn.Execute sql
     
     sql = ""
     sql = "set dateformat dmy Select distinct(TransDate) from d_Milkintake WHERE TransDate >='" & Startdate & "' And TransDate<='" & Enddate & "' order by TransDate asc"
     Set rstg = cn.Execute(sql)
     While Not rstg.EOF
      sql = ""
      sql = "set dateformat dmy Select isnull(sum(PAmount),0) from d_Milkintake WHERE TransDate ='" & rstg.Fields(0) & "'"
  '  sql = "set dateformat dmy SELECT d.DispQnty,m.DName, d.Price, d.DispQnty,d.DCode FROM d_MilkControl AS d INNER JOIN d_Debtors AS m ON d.DCode = m.DCode WHERE " & C & " and DispDate between " & Startdate & " And " & Enddate & """"
      Set rst = cn.Execute(sql)
      If Not rst.EOF Then
       sql = ""
       sql = "set dateformat dmy Select isnull(sum(Amount),0) from d_Debtorsparchases WHERE Date ='" & rstg.Fields(0) & "'"
       Set rsa = cn.Execute(sql)
        If Not rsa.EOF Then
         sql = ""
         sql = "set dateformat dmy insert into  d_incomestate(Date, Sales, Purchases,Diff)"
         sql = sql & "  values('" & rstg.Fields(0) & "','" & rsa.Fields(0) & "'," & rst.Fields(0) & ",'" & rsa.Fields(0) - rst.Fields(0) & "')"
         cn.Execute sql
        End If
      End If
       rstg.MoveNext
      Wend
reportname = "Incomestatement.rpt"
Show_Sales_Crystal_Report STRFORMULA, reportname, ""
Exit Sub
ErrorHandler:
MsgBox err.description
End Sub

Private Sub cmdMilkVeh_Click()
On Error GoTo ErrorHandler
'******************check if aready disptch********************

    sql = ""
    sql = "SET dateformat dmy SELECT * FROM  d_OutletDispatch WHERE Date = '" & txtdateenterered & "'and Vehicle = '" & cbov & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
    MsgBox "You have already dispatch for this outlet on that day", vbInformation
    txtMilkd.SetFocus
    Exit Sub
    End If
'*******************end ************************
'cbobranch = "Null"

'**********************8list view
  Dim j As Integer
   If ListView200.ListItems.Count = 0 Then
     MsgBox "There are no items sold."
   Exit Sub
   End If
   j = 1
   
   Dim total, bam As Currency
   total = 0
   Do While Not j > (ListView200.ListItems.Count)
     ListView200.ListItems.Item(j).selected = True
     total = total + CCur(ListView200.SelectedItem.SubItems(3))
     j = j + 1
   Loop
      '*************************** check if amount paid is less expected

'// check if they are in stock.
For j = 1 To ListView200.ListItems.Count
 ListView200.ListItems.Item(j).selected = True
 
  Dim rsinstock As Recordset
  sql = ""
  sql = "select * from d_OutletDispatch where Vehicle= '" & ListView200.SelectedItem & "' AND OutletName='" & ListView200.SelectedItem.SubItems(2) & "' AND Date ='" & txtdateenterered.value & "'"
   Set rsinstock = oSaccoMaster.GetRecordset(sql)
  If rsinstock.EOF Then
   '// insert into d_outletsales
  sql = ""
  sql = "set dateformat dmy insert into  d_OutletDispatch(Date, Vehicle, OutletName, Quantity)"
  sql = sql & "  values('" & ListView200.SelectedItem.SubItems(1) & "','" & ListView200.SelectedItem & "','" & ListView200.SelectedItem.SubItems(2) & "','" & ListView200.SelectedItem.SubItems(3) & "')"
  cn.Execute sql
    '//XXXXXXXXXXXXXXX
    Else
   sql = ""
   sql = "set dateformat DMY Update d_OutletDispatch SET Vehicle= '" & ListView200.SelectedItem & "',OutletName='" & ListView200.SelectedItem.SubItems(2) & "', Date='" & ListView200.SelectedItem.SubItems(1) & "',Quantity='" & ListView200.SelectedItem.SubItems(3) & "' WHERE Vehicle= '" & ListView200.SelectedItem & "' and OutletName='" & ListView200.SelectedItem.SubItems(2) & "' and Date='" & ListView200.SelectedItem.SubItems(1) & "'"
   cn.Execute sql
 '********************************************************************end

End If
Next j
txtMilkd = ""
cbov = ""
cbobranch = ""
ListView200.ListItems.Clear
loaddispmilk

Exit Sub
ErrorHandler:
MsgBox err.description

End Sub

Private Sub cmdNew_Click()
sql = ""
sql = "select count(p_code) from d_Outlet"
Set rs = oSaccoMaster.GetRecordset(sql)
'Set rs = oSaccoMaster.GetRecordset("d_sp_PNO")
If Not rs.EOF Then
txtpcode = rs.Fields(0) + 1
Else
txtpcode = 1
Exit Sub
End If
nr = 1
chkyoghurt.value = vbUnchecked
'txtpassit = ""
txtsellingprice = ""
txtpprice = ""
txtquantity = ""
'cbosupplier = ""
txtpname = ""
txtbalance = ""
'txtserialno = ""
mm.Enabled = True
End Sub

Private Sub txtpcodeO_Change()
'//TWNG001
Provider = "MAZIWA"
Set cn = New ADODB.Connection
cn.Open Provider, "atm", "atm"
'If reportpath = "" Then reportpath = GetSetting("payroll", "AppName", "rptPath", rptPath)
sql = "select p_code, p_name, Qin, Qout, o_bal, Wprice, Rprice from d_Outlet where p_code='" & Y & "'AND Branch='" & cbobranch & "'"
'sql = "select P_CODE,P_NAME,S_NO,QOUT,supplierid,pprice, sprice,QIN from ag_products where p_code='" & txtpcode & "'AND Branch='" & cbobranch & "' "
Set rs = New ADODB.Recordset
rs.Open sql, cn
If Not rs.EOF Then
 txtpcode = (rs.Fields(0))
 If Not IsNull(rs.Fields(1)) Then txtpname = (rs.Fields(1))
 If Not IsNull(rs.Fields(3)) Then txtbalance = (rs.Fields(3))
'If Not IsNull(rs.Fields(4)) Then cbosupplier = (rs.Fields(4))
 If Not IsNull(rs.Fields(5)) Then txtpprice = (rs.Fields(5))
 If Not IsNull(rs.Fields(6)) Then txtsellingprice = (rs.Fields(6))
'If Not IsNull(rs.Fields(7)) Then txtreceived = (rs.Fields(7))
If txtbalance <= 0 Then
MsgBox "Warning:Your stock is below zero please reorder", vbInformation
Else

End If
End If


'// check with serial no if it exist
End Sub

Private Sub Combo1_Change()

End Sub

Private Sub cmdnew2_Click()
k = 0
 Text1.Text = ""
 cboproductname1.Text = ""
 Label20 = ""
 cbobranch.Text = ""
chkretail.value = vbuncheck
chkWhole.value = vbuncheck
cmdsave1.Enabled = True
SSTab1_DblClick
SSTab2_DblClick
End Sub

Private Sub cmdnewvehiclz_Click()
frmVehicleReg.Show vbModal
End Sub

Private Sub cmdnext1_Click()
Dim cash As Integer
Dim total As Double
Dim j, Coun As Integer
j = 1
cbobranch = "Null"
'Check if same item is in the list
   Do While Not j > (Coun)
         ListView200.ListItems.Item(j).selected = True
            
    If ListView200.SelectedItem = Text1 Then
        'Text2 = (CCur(Text2) + CCur(Lvwitems.SelectedItem.ListSubItems(2)))
        ListView200.ListItems.Remove (ListView200.SelectedItem.Index)
                        
        Set li = ListView200.ListItems.Add(, , cbov)
                        li.SubItems(1) = txtdateenterered & ""
                        li.SubItems(2) = cbobranch & ""
                        li.SubItems(3) = txtMilkd & ""
'                        li.SubItems(1) = cboproductname1 & ""
'                        li.SubItems(2) = Text2 & ""
'                        li.SubItems(3) = Label20 & ""
'                        'li.SubItems(4) = CCur(txtprice) * CCur(txtquantity) & ""
'                        li.SubItems(4) = a & ""
'                        li.SubItems(5) = txtpaybill & ""
'                        'Total = CCur(Total + li.SubItems(4))
'                        TXTTOTAL = total
                                                
        j = Coun + 1
        
 '       Label18 = CCur(Label18) - CCur(Text2)

        cbobranch = ""
        txtMilkd = ""
        cbobranch.SetFocus
        'cboproductname1.SetFocus
        Exit Sub
    
  '   lvwItems.ListItems.Item(J).selected = True
   End If
   j = j + 1
    Loop
    
     If j > 1 Then
   
    Set li = ListView200.ListItems.Add(, , cbov)
                        li.SubItems(1) = txtdateenterered & ""
                        li.SubItems(2) = cbobranch & ""
                        li.SubItems(3) = txtMilkd & ""
'                        li.SubItems(2) = Text2 & ""
'                        li.SubItems(3) = Label20 & ""
'                        'li.SubItems(4) = CCur(txtprice) * (CCur(txtquantity)) & ""
'                        li.SubItems(4) = a & ""
'                        li.SubItems(5) = txtpaybill & ""
'                        'Total = CCur(Total + li.SubItems(4))
'                        TXTTOTAL = total
'
         cbobranch = ""
        txtMilkd = ""
        cbobranch.SetFocus
        'lblbalance = CCur(lblbalance) - CCur(txtquantity)
'        cboproductname1 = ""
'        Text2 = ""
'        'txtserialno = ""
'        cboproductname1.SetFocus
        Exit Sub
    End If
     If Coun = 0 Then
     Set li = ListView200.ListItems.Add(, , cbov)
                        li.SubItems(1) = txtdateenterered & ""
                        li.SubItems(2) = cbobranch & ""
                        li.SubItems(3) = txtMilkd & ""
'                        li.SubItems(1) = cboproductname1 & ""
'                        li.SubItems(2) = Text2 & ""
'                        li.SubItems(3) = Label20 & ""
'                        'li.SubItems(4) = CCur(txtprice) * (CCur(txtquantity)) & ""
'                        li.SubItems(4) = a & ""
 '                       li.SubItems(5) = txtpaybill & ""
                        'Total = CCur(Total + li.SubItems(4))
   '                     TXTTOTAL = total
    End If


Do While Not j > (ListView200.ListItems.Count)
 ListView200.ListItems.Item(j).selected = True
 total = total + CCur(ListView200.SelectedItem.SubItems(3))
 'TXTTOTAL = total
j = j + 1
Loop
 txtMilkd = ""
 cbov = ""


End Sub

Private Sub cmdnextitem_Click()
 On Error GoTo ErrorHandler

Set rst20 = New ADODB.Recordset
Set rst20 = oSaccoMaster.GetRecordset("set dateformat dmy select isnull(sum(Qout),0) from d_Outlet where p_name ='" & cboproductname1 & "' and branch ='" & cbobranch & "' and Date_Entered ='" & txtdateenterered & "'")
oSaccoMaster.ExecuteThis (sql)
If rst20.Fields(0) <= 0 Then
 MsgBox "Sorry Stock for this date " & txtdateenterered & " is Zero", vbInformation
Exit Sub
End If
''end
Provider = "maziwa"
Set cn = New ADODB.Connection
cn.Open Provider, "atm", "atm"
'// check if they are in stock.
Dim rsinstock, rsinstock2 As Recordset
sql = ""
sql = "set dateformat dmy select isnull(sum(Qout),0) from d_Outlet where p_code='" & Text1 & "' and Branch like'" & cbobranch & "%' and Date_Entered>='" & Startdate & "'and Date_Entered<='" & Enddate & "'"
Set rsinstock = New ADODB.Recordset
rsinstock.Open sql, cn
'// check the stock if it is less than zero
k = 0
txtamount.Text = ""

'''''' '''convert youghart to litres
sql = ""
sql = "set dateformat dmy select p_code,Qout,type from d_Outlet where p_code='" & Text1 & "' and Branch like'" & cbobranch & "%' AND Date_Entered='" & txtdateenterered.value & "'"
Set rsk4 = New ADODB.Recordset
Set rsk4 = oSaccoMaster.GetRecordset(sql)
 '''convert youghart to litres
 If rsk4.Fields(2) = 1 Then
 Text2 = Text2 * 0.1
  ElseIf rsk4.Fields(2) = 2 Then
 Text2 = Text2 * 0.15
  ElseIf rsk4.Fields(2) = 3 Then
 Text2 = Text2 * 0.25
  ElseIf rsk4.Fields(2) = 4 Then
 Text2 = Text2 * 0.5
 End If
''''''' end yoghurt

If rsinstock.Fields(0) < CCur(Text2) Then
MsgBox "Sorry Stock cannot be less, please re-stock before your proceed", vbInformation
Exit Sub
End If

If rsinstock.Fields(0) < 0 Then
MsgBox "Sorry Stock is Zero please re-stock before your proceed", vbInformation
Exit Sub
End If

If txtpaybill = "" Then
MsgBox "Please enter the PAYBILL CODE or If CASH", vbInformation
txtpaybill.SetFocus
Exit Sub
End If

'''convert youghart to pices
 If rsk4.Fields(2) = 1 Then
 Text2 = Text2 / 0.1
  ElseIf rsk4.Fields(2) = 2 Then
 Text2 = Text2 / 0.15
  ElseIf rsk4.Fields(2) = 3 Then
 Text2 = Text2 / 0.25
  ElseIf rsk4.Fields(2) = 4 Then
 Text2 = Text2 / 0.5
 End If
''''''' end yoghurt

Dim cash As Integer
Dim total As Double
Dim j, Coun As Integer
j = 1

'Check if same item is in the list
   Do While Not j > (Coun)
         Lvwitems.ListItems.Item(j).selected = True
            
    If Lvwitems.SelectedItem = Text1 Then
        Text2 = (CCur(Text2) + CCur(Lvwitems.SelectedItem.ListSubItems(2)))
        Lvwitems.ListItems.Remove (Lvwitems.SelectedItem.Index)
                        
        Set li = Lvwitems.ListItems.Add(, , Text1)
                        li.SubItems(1) = cboproductname1 & ""
                        li.SubItems(2) = Text2 & ""
                        li.SubItems(3) = Label20 & ""
                        'li.SubItems(4) = CCur(txtprice) * CCur(txtquantity) & ""
                        li.SubItems(4) = a & ""
                        li.SubItems(5) = txtpaybill & ""
                        'Total = CCur(Total + li.SubItems(4))
                        TXTTOTAL = total
                                                
        j = Coun + 1
        
        Label18 = CCur(Label18) - CCur(Text2)

        cboproductname1 = ""
        Text2 = ""
       ' txtserialno = ""
        cboproductname1.SetFocus
        Exit Sub
    
  '   lvwItems.ListItems.Item(J).selected = True
   End If
   j = j + 1
    Loop
    
     If j > 1 Then
   
    Set li = Lvwitems.ListItems.Add(, , Text1)
                        li.SubItems(1) = cboproductname1 & ""
                        li.SubItems(2) = Text2 & ""
                        li.SubItems(3) = Label20 & ""
                        'li.SubItems(4) = CCur(txtprice) * (CCur(txtquantity)) & ""
                        li.SubItems(4) = a & ""
                        li.SubItems(5) = txtpaybill & ""
                        'Total = CCur(Total + li.SubItems(4))
                        TXTTOTAL = total
                        
        'lblbalance = CCur(lblbalance) - CCur(txtquantity)
        cboproductname1 = ""
        Text2 = ""
        'txtserialno = ""
        cboproductname1.SetFocus
        Exit Sub
    End If
     If Coun = 0 Then
     Set li = Lvwitems.ListItems.Add(, , Text1)
                        li.SubItems(1) = cboproductname1 & ""
                        li.SubItems(2) = Text2 & ""
                        li.SubItems(3) = Label20 & ""
                        'li.SubItems(4) = CCur(txtprice) * (CCur(txtquantity)) & ""
                        li.SubItems(4) = a & ""
                        li.SubItems(5) = txtpaybill & ""
                        'Total = CCur(Total + li.SubItems(4))
                        TXTTOTAL = total
    End If

'lblbalance = CCur(lblbalance) - CCur(txtquantity)
TXTTOTAL = 0
'Coun = Lvwitems.ListItems.Count
'For j = 1 To Lvwitems.ListItems.Count
'    Total = CCur(Total + li.SubItems(4))
'    txttotal = Total
'
'Next j
Do While Not j > (Lvwitems.ListItems.Count)
'For j = 1 To Lvwitems.ListItems.Count
 Lvwitems.ListItems.Item(j).selected = True
 total = total + CCur(Lvwitems.SelectedItem.SubItems(3))
 TXTTOTAL = total
j = j + 1
Loop

 '''convert youghart to litres
 If rsk4.Fields(2) = 1 Then
 Text2 = Text2 * 0.1
  ElseIf rsk4.Fields(2) = 2 Then
 Text2 = Text2 * 0.15
  ElseIf rsk4.Fields(2) = 3 Then
 Text2 = Text2 * 0.25
  ElseIf rsk4.Fields(2) = 4 Then
 Text2 = Text2 * 0.5
 End If
''''''' end yoghurt
'************************update  stockbalance*************
sql = ""
sql = "set dateformat dmy select Qout from d_Outlet where p_code='" & Text1 & "' and Branch='" & cbobranch & "' and Date_Entered='" & txtdateenterered & "'"
Set rsinstock2 = New ADODB.Recordset
rsinstock2.Open sql, cn
 sql = ""
 sql = "set dateformat DMY Update d_Outlet SET Qout =" & rsinstock2.Fields(0) - Text2 & " WHERE p_code= '" & Text1 & "' and Date_Entered='" & txtdateenterered & "' and branch='" & cbobranch & "'"
 cn.Execute sql
'************************end ******************************


txtpaybill = "CASH PAYMENT"
Text2 = "0"
cboproductname1 = ""
Text1 = ""
Label18 = ""
Label38 = ""
chkretail.value = vbUnchecked
chkWhole.value = vbUnchecked
txtamount.SetFocus
Exit Sub
ErrorHandler:
MsgBox err.description
End Sub

Private Sub cmdoutsalesre_Click()
reportname = "Outlet Report.rpt"
Show_Sales_Crystal_Report "", reportname, ""
End Sub

Private Sub cmdParchase_Click()
On Error GoTo ErrorHandler
Startdate = DateSerial(year(txtdateenterered), month(txtdateenterered), 1)
Enddate = DateSerial(year(txtdateenterered), month(txtdateenterered) + 1, 1 - 1)
sql = ""
sql = "set dateformat dmy delete from d_Debtorsparchases where Date >= '" & Startdate & "' And Date<='" & Enddate & "'"
cn.Execute sql
prgStatus.Visible = True
txtlbl.Visible = 1
If txtlbl.Visible = True Then
 txtlbl = "Please wait as it precess"
End If
      ' MsgBox "Please wait as it precess  "
sql = ""
sql = "set dateformat dmy Select count(distinct(DCode)) as j  from   d_MilkControl where DispDate >= '" & Startdate & "' And DispDate<='" & Enddate & "' "
Set rs = cn.Execute(sql)
'Set rs = oSaccoMaster.GetRecordset(sql)
Dim a As Double
a = rs.Fields(0)
j = rs.Fields(0)
prgStatus.max = 100
prgStatus.Min = 0
I = 0
'baet
 sql = ""
 sql = "set dateformat dmy Select distinct(DCode) from   d_MilkControl where DispDate >= '" & Startdate & "' And DispDate<='" & Enddate & "' order by DCode asc"
 Set rsd = cn.Execute(sql)
  While Not rsd.EOF
  Do While Not j = 0
I = I + 1
prgStatus = Round((I / a) * 100, 0)
  If Not rsd.EOF Then
  'C = "CB003A22"
  C = rsd.Fields(0)

'     If C = "CB003A22" Then
'        MsgBox "Warning:Please wait " & C & "", vbInformation
'     End If
 '  Label40 = "Please wait as it precess"

   sql = ""
   sql = "set dateformat dmy Select distinct(Price) from   d_MilkControl where DCode='" & C & "'and DispDate >= '" & Startdate & "' And DispDate<='" & Enddate & "' "
   Set rs = cn.Execute(sql)
    sql = ""
    sql = "set dateformat dmy Select count(distinct(Price)) from   d_MilkControl where DCode='" & C & "'and DispDate >= '" & Startdate & "' And DispDate<='" & Enddate & "' "
    Set rsg = cn.Execute(sql)
     'g = rsg.Fields(0)
     'q = 1
   'For q = 1 To g
   Do While Not rs.EOF
    If Not rs.EOF Then
     k = rs.Fields(0)
     sql = ""
     sql = "set dateformat dmy Select distinct(DispDate) from   d_MilkControl WHERE DCode= '" & C & "'and Price='" & k & "' and DispDate >='" & Startdate & "' And DispDate<='" & Enddate & "'"
     Set rstg = cn.Execute(sql)
     While Not rstg.EOF
      sql = ""
      sql = "set dateformat dmy Select sum(DispQnty) from   d_MilkControl WHERE DCode= '" & C & "'and Price='" & k & "' and DispDate ='" & rstg.Fields(0) & "'"
  '  sql = "set dateformat dmy SELECT d.DispQnty,m.DName, d.Price, d.DispQnty,d.DCode FROM d_MilkControl AS d INNER JOIN d_Debtors AS m ON d.DCode = m.DCode WHERE " & C & " and DispDate between " & Startdate & " And " & Enddate & """"
      Set rst = cn.Execute(sql)
      If Not rst.EOF Then
        sql = ""
        sql = "Select distinct(DName) from   d_Debtors where DCode='" & C & "' "
        Set rsa = cn.Execute(sql)
        If Not rsa.EOF Then
         p = rsa.Fields(0)
           sql = ""
           sql = "Select distinct(Locations) from   d_Debtors where DCode='" & C & "' "
           Set rsv = cn.Execute(sql)
         sql = ""
         sql = "set dateformat dmy insert into  d_Debtorsparchases(Debtor, Name, Kgs, Price,Amount,Description,Branch,Date)"
         sql = sql & "  values('" & C & "','" & p & "'," & rst.Fields(0) & ",'" & k & "','" & rst.Fields(0) * k & "','OUTLET SALES','" & rsv.Fields(0) & "','" & rstg.Fields(0) & "')"
         cn.Execute sql
        End If
      End If
       rstg.MoveNext
      Wend
    End If
  k = rs.MoveNext
  Loop
  'Next q
    Else
    'baet
    'kamorok
  txtlbl.Visible = False
  
  
  MsgBox "Completed succesfully ", vbInformation
    Exit Sub
    End If

  j = j - 1

  'MsgBox "Warning:Please wait " & j & "", vbInformation
 rsd.MoveNext

Loop
Wend
  'baet
  'kamorok
  txtlbl.Visible = False
MsgBox "Completed succesfully ", vbInformation
Exit Sub
ErrorHandler:
MsgBox err.description
End Sub
Private Sub kamorok()
sql = ""
sql = "set dateformat dmy Select count(distinct(OutletName)) as j  from   d_OutletSales where Date >= '" & Startdate & "' And Date<='" & Enddate & "' "
Set rs = cn.Execute(sql)
'Set rs = oSaccoMaster.GetRecordset(sql)
a = rs.Fields(0)
Dim t As String
sql = ""
sql = "set dateformat dmy Select distinct(OutletName) as j  from   d_OutletSales where Date >= '" & Startdate & "' And Date<='" & Enddate & "'ORDER BY OutletName ASC"
Set rss = cn.Execute(sql)
'Set rs = oSaccoMaster.GetRecordset(sql)
'baet
 'sql = ""
 'sql = "set dateformat dmy Select distinct(PCode) from   d_OutletSales where Date >= '" & Startdate & "' And Date<='" & Enddate & "'"
 'Set rsd = cn.Execute(sql)
 'While Not rsd.EOF
  Do While Not a <= 0
   t = rss.Fields(0)
  sql = ""
  sql = "set dateformat dmy Select distinct(PCode) from   d_OutletSales where OutletName='" & t & "'AND Date >= '" & Startdate & "' And Date<='" & Enddate & "'"
  Set rsd = cn.Execute(sql)
  Do While Not rsd.EOF
  If Not rsd.EOF Then
  'C = "12"
  C = rsd.Fields(0)
   
   sql = ""
   sql = "set dateformat dmy Select Wprice, Rprice,type from d_Outlet where p_code='" & C & "' and Branch='" & t & "'"
   Set rs = cn.Execute(sql)
     Set rsr = New Recordset
    sql = ""
    sql = "set dateformat dmy Select count(distinct(Description))as r from   d_OutletSales where PCode='" & C & "'and OutletName='" & t & "' and Date >= '" & Startdate & "' And Date<='" & Enddate & "' "
    Set rsr = oSaccoMaster.GetRecordset(sql)
    If rsr.Fields(0) > 0 Then
     Dim r As Integer
      r = rsr.Fields(0)
      o = 1
     End If
      sql = ""
      sql = "set dateformat dmy Select distinct(Description) from d_OutletSales where PCode='" & C & "'and OutletName='" & t & "' and Date >= '" & Startdate & "' And Date<='" & Enddate & "' order by Description desc "
      Set rsq = cn.Execute(sql)
   For o = 1 To r
    If Not rs.EOF Then

      S = rsq.Fields(0)
      If S = "Whole sales" Then
        k = rs.Fields(0)
      Else
        k = rs.Fields(1)
      End If
      '////////checking per day then to sum
    sql = ""
    sql = "set dateformat dmy Select distinct(Date) from d_OutletSales WHERE PCode='" & C & "'and OutletName='" & t & "' and Description='" & rsq.Fields(0) & "' and Date >='" & Startdate & "' And Date<='" & Enddate & "'"
    Set rstgsk = cn.Execute(sql)
    While Not rstgsk.EOF
    
     sql = ""
     sql = "set dateformat dmy Select sum(Quant) from   d_OutletSales WHERE PCode= '" & C & "'and OutletName='" & t & "' and Description='" & rsq.Fields(0) & "' and Date ='" & rstgsk.Fields(0) & "'"
  '  sql = "set dateformat dmy SELECT d.DispQnty,m.DName, d.Price, d.DispQnty,d.DCode FROM d_MilkControl AS d INNER JOIN d_Debtors AS m ON d.DCode = m.DCode WHERE " & C & " and DispDate between " & Startdate & " And " & Enddate & """"
     Set rst = cn.Execute(sql)
      If Not rst.EOF Then
        sql = ""
        sql = "set dateformat dmy Select distinct(PName) from   d_OutletSales where PCode='" & C & "' and OutletName='" & t & "' and Date >= '" & Startdate & "' And Date<='" & Enddate & "' "
        Set rsa = cn.Execute(sql)
        If Not rsa.EOF Then
         p = rsa.Fields(0)
         Dim D As Double
         sql = ""
         sql = "set dateformat dmy Select distinct(type) from d_Outlet where p_code='" & C & "' and Branch='" & t & "' and Date_Entered >= '" & Startdate & "' And Date_Entered<='" & Enddate & "'"
         Set rsoutle = cn.Execute(sql)
         If rsoutle.Fields(0) > 0 Then
          D = rst.Fields(0)
         '''convert youghart to pices
             If rsoutle.Fields(0) = 1 Then
             D = D / 0.1
              ElseIf rsoutle.Fields(0) = 2 Then
             D = D / 0.15
              ElseIf rsoutle.Fields(0) = 3 Then
             D = D / 0.25
              ElseIf rsoutle.Fields(0) = 4 Then
             D = D / 0.5
             End If
            ''''''' end yoghurt
           sql = ""
           sql = "set dateformat dmy insert into  d_Debtorsparchases(Debtor, Name, Kgs, Price,Amount,Description,Branch,Date)"
           sql = sql & "  values('" & C & "','" & p & "'," & rst.Fields(0) & ",'" & k & "','" & D * k & "','" & S & "','" & t & "','" & rstgsk.Fields(0) & "')"
           cn.Execute sql
           GoTo Mwiraria
         End If
         sql = ""
         sql = "set dateformat dmy insert into  d_Debtorsparchases(Debtor, Name, Kgs, Price,Amount,Description,Branch,Date)"
         sql = sql & "  values('" & C & "','" & p & "'," & rst.Fields(0) & ",'" & k & "','" & rst.Fields(0) * k & "','" & S & "','" & t & "','" & rstgsk.Fields(0) & "')"
         cn.Execute sql
        End If
      End If
Mwiraria:
      rstgsk.MoveNext
     Wend
    End If
  S = rsq.MoveNext
  Next o
    Else
 
    Exit Sub
    End If
  
  rsd.MoveNext
 Loop
  rss.MoveNext
  a = a - 1
  'MsgBox "Warning:Please wait " & j & "", vbInformation
 
Loop
Exit Sub
'Wend

End Sub

Private Sub baet()
sql = ""
sql = "SET dateformat dmy Select count(distinct(Price)) as j  from d_Outsalesb where Description='Sales from siche'and Date >= '" & Startdate & "' And Date<='" & Enddate & "'"
Set rsq = cn.Execute(sql)
j = rsq.Fields(0)
 sql = ""
 sql = "set dateformat dmy Select distinct(Price) from d_Outsalesb where Description='Sales from siche'and Date >= '" & Startdate & "' And Date<='" & Enddate & "'"
 Set rsd = cn.Execute(sql)
  While Not rsd.EOF
  Do While Not j <= 0
  If Not rsd.EOF Then
   C = rsd.Fields(0)
    sql = ""
    sql = "set dateformat dmy Select distinct(Date) from d_Outsalesb WHERE Price='" & C & "'and Date >= '" & Startdate & "' And Date<='" & Enddate & "' and Description='Sales from siche'"
    Set rstgs = cn.Execute(sql)
     While Not rstgs.EOF
    sql = ""
    sql = "set dateformat dmy Select distinct(Name) from   d_Outsalesb where Price='" & C & "'and Date = '" & rstgs.Fields(0) & "' and Description='Sales from siche' "
    Set rsy = cn.Execute(sql)
       
       sql = ""
       sql = "set dateformat dmy Select count(distinct(Name)) as n  from d_Outsalesb where Price='" & C & "' and Date = '" & rstgs.Fields(0) & "' and Description='Sales from siche'"
       Set rsc = cn.Execute(sql)
      Do While Not rsy.EOF
       If Not rsy.EOF Then
       If Not rsc.EOF Then
       
       
     sql = ""
     sql = "set dateformat dmy Select sum(Quantity) from   d_Outsalesb WHERE Price= '" & C & "'and Name='" & rsy.Fields(0) & "' and Date ='" & rstgs.Fields(0) & "' and Description='Sales from siche'"
     Set rst = cn.Execute(sql)
      If Not rst.EOF Then
        sql = ""
         sql = "Select distinct(Code) from d_Outsalesb where Name='" & rsy.Fields(0) & "' and Description='Sales from siche'"
         Set rsa = cn.Execute(sql)
        If Not rsa.EOF Then
         p = rsa.Fields(0)
         sql = ""
         sql = "set dateformat dmy insert into  d_Debtorsparchases(Debtor, Name, Kgs, Price,Amount,Description,Branch,Date)"
         sql = sql & "  values('" & p & "','" & rsy.Fields(0) & "'," & rst.Fields(0) & ",'" & C & "','" & rst.Fields(0) * C & "','Sales from siche','FACTORY','" & rstgs.Fields(0) & "')"
         cn.Execute sql
        End If
      End If
    End If
    rsy.MoveNext
    Else
    rsc.MoveNext
    End If
    Loop
    
    rstgs.MoveNext
  Wend
  End If
  j = j - 1

 rsd.MoveNext
Loop
Wend
End Sub


Private Sub cmdpareport_Click()
reportname = "MILK PURCHASE REPORT.rpt"
Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub cmdRemove_Click()
On Error GoTo ErrorHandler
Dim total As Double
Dim j, Coun As Integer

'''''' '''convert youghart to litres
sql = ""
sql = "set dateformat dmy select p_code,Qout,type from d_Outlet where p_code='" & Lvwitems.SelectedItem & "' and Branch='" & cbobranch & "' and Date_Entered='" & txtdateenterered & "'"
Set rsk4 = New ADODB.Recordset
Set rsk4 = oSaccoMaster.GetRecordset(sql)
 '''convert youghart to litres
 If rsk4.Fields(2) = 1 Then
 Lvwitems.SelectedItem.SubItems(2) = Lvwitems.SelectedItem.SubItems(2) / 0.1
  ElseIf rsk4.Fields(2) = 2 Then
 Lvwitems.SelectedItem.SubItems(2) = Lvwitems.SelectedItem.SubItems(2) / 0.15
  ElseIf rsk4.Fields(2) = 3 Then
 Lvwitems.SelectedItem.SubItems(2) = Lvwitems.SelectedItem.SubItems(2) / 0.25
  ElseIf rsk4.Fields(2) = 4 Then
 Lvwitems.SelectedItem.SubItems(2) = Lvwitems.SelectedItem.SubItems(2) / 0.5
 End If
''''''' end yoghurt

'************************update  stockbalance*************
Dim rsinstock As Recordset
sql = ""
sql = "set dateformat dmy select p_code,Qout from d_Outlet where p_code='" & Lvwitems.SelectedItem & "' and Branch='" & cbobranch & "' and Date_Entered='" & txtdateenterered & "'"
Set rsinstock = cn.Execute(sql)
'Set rsa = cn.Execute(sql)
If Not rsinstock.EOF Then
 sql = ""
 sql = "set dateformat DMY Update d_Outlet SET Qout =" & rsinstock.Fields(1) + Lvwitems.SelectedItem.SubItems(2) & " WHERE p_code= '" & Lvwitems.SelectedItem & "' and branch='" & cbobranch & "' and Date_Entered='" & txtdateenterered & "'"
 cn.Execute sql
'************************end ******************************
j = 1
'On Error GoTo ErrorHandler
TXTTOTAL = 0
'If Lvwitems.ListItems.Count > 0 Then
''Total = CCur(txttotal - li.SubItems(4))
Lvwitems.ListItems.Remove (Lvwitems.SelectedItem.Index)  '// removes the selected item

Do While Not j > (Lvwitems.ListItems.Count)
'For j = 1 To Lvwitems.ListItems.Count
 Lvwitems.ListItems.Item(j).selected = True
 total = total + CCur(Lvwitems.SelectedItem.SubItems(3))
 TXTTOTAL = total
j = j + 1
Loop

End If
Exit Sub
ErrorHandler:
MsgBox err.description
End Sub

Private Sub cmdReportoutlet_Click()
reportname = "Outlet Report.rpt"
Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub cmdSalessta_Click()

'Show_Sales_Crystal_Report STRFORMULA, reportname, ""

Dim ans As String
ans = MsgBox("Do you Want a Report as per price??", vbYesNo)
If ans = vbYes Then
 'reportname = "d_Dailysummary2.rpt"
reportname = "MILK SALES REPORT1.rpt"
 Else
 'reportname = "d_Dailysummary.rpt"
reportname = "MILK SALES REPORT.rpt"
 End If
 Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub cmdsave1_Click()
On Error GoTo ErrorHandler
'If txtpaybill = "" Then
'MsgBox "Please enter the PAYBILL CODE", vbInformation
'txtpaybill.SetFocus
'Exit Sub
'End If
'If cbobranch = "" Then
'MsgBox "Please select branch", vbInformation
'Exit Sub
'End If

If txtamount = "" Then
MsgBox "Amount paid needed", vbInformation
Exit Sub
End If
'''''' '''convert youghart to litres
sql = ""
sql = "set dateformat dmy select p_code,Qout,type from d_Outlet where p_code='" & Lvwitems.SelectedItem & "' and Branch='" & cbobranch & "' AND Date_Entered='" & txtdateenterered.value & "'"
Set rsk4 = New ADODB.Recordset
Set rsk4 = oSaccoMaster.GetRecordset(sql)
 '''convert youghart to litres
 If rsk4.Fields(2) = 1 Then
 Lvwitems.SelectedItem.SubItems(2) = Lvwitems.SelectedItem.SubItems(2) * 0.1
  ElseIf rsk4.Fields(2) = 2 Then
 Lvwitems.SelectedItem.SubItems(2) = Lvwitems.SelectedItem.SubItems(2) * 0.15
  ElseIf rsk4.Fields(2) = 3 Then
 Lvwitems.SelectedItem.SubItems(2) = Lvwitems.SelectedItem.SubItems(2) * 0.25
  ElseIf rsk4.Fields(2) = 4 Then
 Lvwitems.SelectedItem.SubItems(2) = Lvwitems.SelectedItem.SubItems(2) * 0.5
 End If
''''''' end yoghurt

'**********************8list view
  Dim j As Integer
   If Lvwitems.ListItems.Count = 0 Then
     MsgBox "There are no items sold."
   Exit Sub
   End If
   j = 1

Do While Not j > (Lvwitems.ListItems.Count)
 Lvwitems.ListItems.Item(j).selected = True
  Z = Z + CCur(Lvwitems.SelectedItem.SubItems(2))
j = j + 1
Loop
 
   Dim total, bam As Currency
   total = 0
   Do While Not j > (Lvwitems.ListItems.Count)
     Lvwitems.ListItems.Item(j).selected = True
     total = total + CCur(Lvwitems.SelectedItem.SubItems(3))
     j = j + 1
   Loop
   
      '*************************** check if amount paid is less expected
    If TXTCHANGE < 0 Then
'       If MsgBox("Insufficient Amount Received, Do you want to continue with that -ve? ", vbYesNo) = vbYes Then
'            lblCheckOff_Click
'            lblCheckOff.value = True
'            optCash.value = False
          ' Exit Sub
         bam = TXTCHANGE / (j - 1)
        Else
         '  Exit Sub
         'End If
          bam = TXTCHANGE / (j - 1)
    End If

'// check if they are in stock.
For j = 1 To Lvwitems.ListItems.Count
 Lvwitems.ListItems.Item(j).selected = True
 ''' check if the product if for this branch
 sql = ""
 sql = "select * from d_Outlet where p_code='" & Lvwitems.SelectedItem & "' and Branch='" & cbobranch & "'"
 Set rsi20 = New ADODB.Recordset
 rsi20.Open sql, cn
 If rsi20.EOF Then
 MsgBox "Sorry this product " & Lvwitems.SelectedItem.SubItems(1) & " is not for this Branch, please select the correct branch to proceed", vbInformation
 Exit Sub
 End If
 '''''end
'       sql = "set dateformat dmy insert into  d_Outsalesb(Code,Name,Date,Quantity,Price,Amount,APaid,Description)"
'       sql = sql & "  values('" & Lvwitems.SelectedItem & "','" & Lvwitems.SelectedItem.SubItems(1) & "','" & txtdateenterered.value & "','" & Lvwitems.SelectedItem.SubItems(2) & "'," & Lvwitems.SelectedItem.SubItems(3) & "," & Lvwitems.SelectedItem.SubItems(4) & "," & Lvwitems.SelectedItem.SubItems(4) + bam & ",'" & Lvwitems.SelectedItem.SubItems(5) & "')"
'        cn.Execute sql
   '// insert into d_outletsales
 If TXTCHANGE < 0 Then
  sql = ""
  sql = "set dateformat dmy insert into  d_OutletSales(PCode, PName, Date,AuditDate,Quant,Amount,Paid,AuditId, Description, OutletName,Mpesa)"
  sql = sql & "  values('" & Lvwitems.SelectedItem & "','" & Lvwitems.SelectedItem.SubItems(1) & "','" & txtdateenterered.value & "','" & Now & "'," & Lvwitems.SelectedItem.SubItems(2) & "," & Lvwitems.SelectedItem.SubItems(3) & "," & Lvwitems.SelectedItem.SubItems(3) + bam & ",'" & User & "','" & Lvwitems.SelectedItem.SubItems(4) & "','" & cbobranch & "','" & Lvwitems.SelectedItem.SubItems(5) & "')"
  cn.Execute sql
    '//XXXXXXXXXXXXXXX
    '********** credit agent ledger sale and stock
    sql = ""
    sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & txtdateenterered & "'," & Lvwitems.SelectedItem.SubItems(3) & ",'" & Label25 & "','" & Label26 & "','','" & Lvwitems.SelectedItem.SubItems(2) & "' ,'SALES ON- " & "" & cbobranch & "','" & User & "',0,0)"
    oSaccoMaster.ExecuteThis (sql)
    '********** credit agent ledger and bank
    sql = ""
    sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & txtdateenterered & "'," & Lvwitems.SelectedItem.SubItems(3) & ",'" & lbldracc & "','" & Label25 & "','','" & Lvwitems.SelectedItem.SubItems(2) & "' ,'SALES ON- " & "" & cbobranch & "','" & User & "',0,0)"
    oSaccoMaster.ExecuteThis (sql)
'XXXXXXXXXXXXXXXXXXXXXX
  
 Else
  sql = ""
  sql = "set dateformat dmy insert into  d_OutletSales(PCode, PName, Date,AuditDate,Quant,Amount,Paid,AuditId, Description, OutletName,Mpesa)"
  sql = sql & "  values('" & Lvwitems.SelectedItem & "','" & Lvwitems.SelectedItem.SubItems(1) & "','" & txtdateenterered.value & "','" & Now & "'," & Lvwitems.SelectedItem.SubItems(2) & "," & Lvwitems.SelectedItem.SubItems(3) & "," & Lvwitems.SelectedItem.SubItems(3) + bam & ",'" & User & "','" & Lvwitems.SelectedItem.SubItems(4) & "','" & cbobranch & "','" & Lvwitems.SelectedItem.SubItems(5) & "')"
  cn.Execute sql
  '//XXXXXXXXXXXXXXX
    '********** credit agent ledger sale and stock
    sql = ""
    sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & txtdateenterered & "'," & Lvwitems.SelectedItem.SubItems(3) & ",'" & Label25 & "','" & Label26 & "','','" & Lvwitems.SelectedItem.SubItems(2) & "' ,'SALES ON- " & "" & cbobranch & "','" & User & "',0,0)"
    oSaccoMaster.ExecuteThis (sql)
    '********** credit agent ledger and bank
    sql = ""
    sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & txtdateenterered & "'," & Lvwitems.SelectedItem.SubItems(3) & ",'" & lbldracc & "','" & Label25 & "','','" & Lvwitems.SelectedItem.SubItems(2) & "' ,'SALES ON- " & "" & cbobranch & "','" & User & "',0,0)"
    oSaccoMaster.ExecuteThis (sql)
'XXXXXXXXXXXXXXXXXXXXXX
  
 End If
 '********************************************************************end

'Dim rsinstock As Recordset
'sql = ""
'sql = "select p_code, p_name, Date_Entered,Qout from d_Outlet where p_code= '" & Lvwitems.SelectedItem & "' AND  branch='" & cbobranch & "'"
'Set rsinstock = oSaccoMaster.GetRecordset(sql)
''Dim Qout As Integer
''Qout = rsinstock!Qout
'sql = "set dateformat DMY Update d_Outlet SET Qout =" & rsinstock.Fields("Qout") - Lvwitems.SelectedItem.SubItems(2) & " WHERE p_code= '" & Lvwitems.SelectedItem & "' and branch='" & cbobranch & "'"
'cn.Execute sql
''*****************************************************
 
Next j

MsgBox "Records succesfully saved."
loadBranchesTypes
SSTab1_DblClick
SSTab2_DblClick
chkretail.value = vbUnchecked
chkWhole.value = vbUnchecked
Text1.Text = ""
cboproductname1.Text = ""

cbobranch.Text = ""
Label18 = ""
TXTCHANGE.Text = ""
TXTTOTAL.Text = ""
k = 0
'txtCustName.Text = ""
Lvwitems.ListItems.Clear
'Text2.Text = ""
'txtAmount.Text = ""
txtpaybill = "CASH PAYMENT"
cmdnew2.Enabled = True
cmdsave1.Enabled = False
cmdnew2_Click

Exit Sub
ErrorHandler:
MsgBox err.description
End Sub

Private Sub cmdSearch_Click()
frmSearchGLAccounts.Show vbModal, Me
    If Continue Then
        If SearchValue <> "" Then
            txtDrAccNo = SearchValue
            SearchValue = ""
        End If
    End If
End Sub

Private Sub cmdshort_Click()
frmNominals.Show vbModal
End Sub

Private Sub cmdvehiclenew_Click()
frmVehicleReg.Show vbModal
End Sub

Private Sub cmdvehreport_Click()
reportname = "OutletVehicledispatch Report.rpt"
Show_Sales_Crystal_Report "", reportname, ""
End Sub

Private Sub Command1_Click()
frmOutletreg.Show vbModal
SSTab1_DblClick
'Form_Load
End Sub

Private Sub Command2_Click()
frmSearchGLAccounts.Show vbModal, Me
    If Continue Then
        If SearchValue <> "" Then
            txtCrAccNo = SearchValue
            SearchValue = ""
        End If
    End If
End Sub

Private Sub cmdremov_Click()

Dim total As Double
Dim j, Coun As Integer
j = 1
On Error GoTo ErrorHandler
TXTTOTAL = 0
'If Lvwitems.ListItems.Count > 0 Then
''Total = CCur(txttotal - li.SubItems(4))
ListView200.ListItems.Remove (ListView200.SelectedItem.Index)  '// removes the selected item

Do While Not j > (ListView200.ListItems.Count)
'For j = 1 To Lvwitems.ListItems.Count
 ListView200.ListItems.Item(j).selected = True
 total = total + CCur(ListView200.SelectedItem)
 'TXTTOTAL = total
j = j + 1
Loop

'End If
Exit Sub
ErrorHandler:
MsgBox err.description

End Sub
Private Sub Form_Load()
txtdateenterered = Format(Get_Server_Date, "dd/mm/yyyy")
txtdateenterered = Format(Get_Server_Date, "dd/mm/yyyy")
'txtdateenterered = DTPMilkDate
   Label23.Visible = False
   Label24.Visible = False
   Label25.Visible = False
   Label26.Visible = False
   cbovb.Visible = False
   prgStatus.Visible = False
   Label39.Visible = False
   chkcustomer.value = 0
   chkrepeat = 0
   k = 0
   nr = 1
   txtpaybill = "CASH PAYMENT"
   txtlbl.Visible = False
SSTab1_DblClick
SSTab2_DblClick
NAMES1
NAMES3
txtdateenterered_Click
cmdsave1.Enabled = False
End Sub

Private Sub lblcracc_Change()
    Set rst = oSaccoMaster.GetRecordset("select glaccname from glsetup where accno='" & lblcracc & "'")
    If Not rst.EOF Then
    txtcracc = rst.Fields("glaccname")
    End If
End Sub

Private Sub lbldracc_Change()
    Set rst = oSaccoMaster.GetRecordset("select glaccname from glsetup where accno='" & lbldracc & "'")
    If Not rst.EOF Then
    txtdracc = rst.Fields("glaccname")
    End If
End Sub

Public Sub loadBranchesTypes()
    
    With ListView1
    
        .ListItems.Clear
        
        .ColumnHeaders.Clear
    
    End With

    Set rs2 = CreateObject("adodb.recordset")
    
    sql = "set dateformat dmy Select PName, Date, Quant, Amount, Paid,Description,Mpesa,OutletName from d_OutletSales where Date='" & txtdateenterered & "'"
'    sql = ""
'    sql = "set dateformat dmy SELECT d.RefNo,m.DName, d.DispDate, d.DispQnty,d.Amount,d.PaidAmount FROM d_MilkControl AS d INNER JOIN d_Debtors AS m ON d.DCode = m.DCode WHERE     (DispDate = '" & txtdateenterered & "') and vehicleno='" & cboVehicle & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    
    Set rs2 = CreateObject("adodb.recordset")
    
    Set clsClass = New cdbase
    
    Provider = clsClass.OpenCon
    
    Set cn = CreateObject("adodb.connection")
    
   cn.Open Provider, "atm", "atm"
    
    rs2.Open sql, cn
    
    With ListView1
        
        .ColumnHeaders.Add , , "Date"
        .ColumnHeaders.Add , , "Product"
        .ColumnHeaders.Add , , "Quantity"
        .ColumnHeaders.Add , , "Amount"
        .ColumnHeaders.Add , , "Paid"
        .ColumnHeaders.Add , , "Type"
        .ColumnHeaders.Add , , "Mpesa"
        .ColumnHeaders.Add , , "Outlet"
        While Not rs2.EOF
        
            Set li = .ListItems.Add(, , Trim(rs2.Fields("Date")))
            
            li.ListSubItems.Add , , Trim(rs2.Fields("PName"))
            li.ListSubItems.Add , , Trim(rs2.Fields("Quant"))
            li.ListSubItems.Add , , Trim(rs2.Fields("Amount"))
            li.ListSubItems.Add , , Trim(rs2.Fields("Paid"))
            li.ListSubItems.Add , , Trim(rs2.Fields("Description"))
            li.ListSubItems.Add , , Trim(rs2.Fields("Mpesa"))
            li.ListSubItems.Add , , Trim(rs2.Fields("OutletName"))
            rs2.MoveNext
        
        Wend
        
    End With
    
    rs2.Close
    
    Set rs2 = Nothing
    
ListView1.View = lvwReport

End Sub

Private Sub lvWBranch1_BeforeLabelEdit(Cancel As Integer)

End Sub
Public Sub loadpro()
    
    With lvWBranch2
    
        .ListItems.Clear
        
        .ColumnHeaders.Clear
    
    End With

    Set rs2 = CreateObject("adodb.recordset")
    
    sql = "set dateformat dmy Select p_name, Date_Entered, Qin, Wprice, Rprice, Branch from d_Outlet where Date_Entered='" & txtdateenterered & "'"
    
    Set rs2 = CreateObject("adodb.recordset")
    
    Set clsClass = New cdbase
    
    Provider = clsClass.OpenCon
    
    Set cn = CreateObject("adodb.connection")
    
   cn.Open Provider, "atm", "atm"
    
    rs2.Open sql, cn
    
    With lvWBranch2
        
        .ColumnHeaders.Add , , "Date"
        .ColumnHeaders.Add , , "Product"
        .ColumnHeaders.Add , , "Quantity"
        .ColumnHeaders.Add , , "W.Price"
        .ColumnHeaders.Add , , "R.Price"
        .ColumnHeaders.Add , , "Branch"
        While Not rs2.EOF
        
            Set li = .ListItems.Add(, , Trim(rs2.Fields("Date_Entered")))
            
            li.ListSubItems.Add , , Trim(rs2.Fields("p_name"))
            li.ListSubItems.Add , , Trim(rs2.Fields("Qin"))
            li.ListSubItems.Add , , Trim(rs2.Fields("Wprice"))
            li.ListSubItems.Add , , Trim(rs2.Fields("Rprice"))
            li.ListSubItems.Add , , Trim(rs2.Fields("Branch"))
            rs2.MoveNext
        
        Wend
        
    End With
    
    rs2.Close
    
    Set rs2 = Nothing
    
lvWBranch2.View = lvwReport

End Sub
Private Sub listview1_DblClick()
chkWhole = 0
chkretail = 0
k = 1
txtamount.Text = ""
cbobranch = ListView1.SelectedItem.SubItems(7)
txtdateenterered = ListView1.SelectedItem
txtpaybill = ListView1.SelectedItem.SubItems(6)
cboproductname1 = ListView1.SelectedItem.SubItems(1)
cboproductname1_Click
Text2 = ListView1.SelectedItem.SubItems(2)
If ListView1.SelectedItem.SubItems(5) = "Whole sales" Then
 chkWhole = 1
Else
 chkretail = 1
End If
txtamount.Text = ListView1.SelectedItem.SubItems(3)
End Sub

Private Sub ListView30_DblClick()
cbovh = ListView30.SelectedItem
cbovh_Validate True
End Sub

Private Sub mm_Click()
On Error GoTo ErrorHandler
Dim w As Integer
 If cbobranch = "" Then
  MsgBox "Please select branch", vbInformation
 Exit Sub
 End If
 
 If txtpname = "" Then
  MsgBox "Please select the product", vbInformation
 Exit Sub
 End If
'''''' '''convert youghart to litres
If chkyoghurt = 1 Then
  If cbotxttyp = "" Then
   MsgBox "Please select the Yoghurt Quantity", vbInformation
  Exit Sub
  End If
 End If
If chkyoghurt = 1 Then
 If cbotxttyp = "100ml" Then
 txtquantity = txtquantity * 0.1
 w = 1
 ElseIf cbotxttyp = "150ml" Then
 txtquantity = txtquantity * 0.15
 w = 2
  ElseIf cbotxttyp = "250ml" Then
 txtquantity = txtquantity * 0.25
 w = 3
  ElseIf cbotxttyp = "500ml" Then
 txtquantity = txtquantity * 0.5
 w = 4
 End If
End If
'''''''''' end yoghurt
If nr = 1 Then
 If chkrepeat = 0 Then
'******************check if aready disptch********************
    sql = ""
    sql = "SET dateformat dmy     SELECT * FROM  d_Outlet  WHERE     Date_Entered = '" & txtdateenterered & "'and p_code = '" & txtpcode & "' and Qin = '" & txtquantity & "' and Branch = '" & cbobranch & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
    MsgBox "You have already dispatch for that day", vbInformation
    txtquantity.SetFocus
    Exit Sub
    End If
'*******************end ************************
 Else
  txtquantity = txtquantity * -1
 End If
 If Trim(txtquantity) = "" Then
  MsgBox "Quantity cannot be Zero", vbInformation
 Exit Sub
 End If

'End If
 If Not IsNumeric(txtquantity) Then
  MsgBox "Enter values please", vbCritical
  txtquantity = ""
  txtquantity.SetFocus
 Exit Sub
 End If

 If Trim(txtbalance) = "" Then txtbalance = 0
  Provider = "MAZIWA"
  Set cn = New ADODB.Connection
 cn.Open Provider, "atm", "atm"
  sql = ""
'If reportpath = "" Then reportpath = GetSetting("payroll", "AppName", "rptPath", rptPath)
  sql = "SET dateformat dmy select P_CODE,qout,Qin from d_Outlet where p_code='" & txtpcode & "'and Date_Entered='" & txtdateenterered.value & "' AND Branch ='" & cbobranch & "'"
  Set rs = New ADODB.Recordset
  rs.Open sql, cn
 If rs.EOF Then
'// insert into ag_products
  If txtSERIALNO = "" Then txtSERIALNO = 0
  Dim vet As Double
  sql = ""
  sql = "set dateformat dmy insert into  d_Outlet( p_code, p_name, Date_Entered, Qin, Qout, o_bal, user_id, Wprice, Rprice, Branch,type)"
  sql = sql & "  values('" & txtpcode.Text & "','" & txtpname.Text & "','" & txtdateenterered.value & "'," & txtquantity.Text & "," & txtquantity.Text & "," & txtquantity.Text & ",'" & User & "'," & txtpprice & "," & txtsellingprice & ",'" & cbobranch & "','" & w & "')"
'  sql = sql & "  values('" & txtpcode.Text & "','" & txtpname.Text & "','" & txtdateenterered.value & "'," & txtquantity.Text & "," & txtbalance.Text + txtquantity.Text & "," & txtquantity.Text & ",'" & User & "'," & txtpprice & "," & txtsellingprice & ",'" & cbobranch & "')"
  cn.Execute sql

  Else
'Dim D As Double
'If Not IsNull(rs.Fields(2)) Then D = rs.Fields(2)
  sql = "set dateformat DMY update d_Outlet set p_name='" & txtpname & "',Qin=" & rs.Fields("Qin") + txtquantity.Text & ",Qout=" & txtquantity.Text + rs.Fields("qout") & ",o_bal=" & txtquantity.Text + rs.Fields("qout") & ",Date_Entered='" & txtdateenterered.value & "',user_id='" & User & "',Wprice=" & txtpprice & ",Rprice=" & txtsellingprice & ",type='" & w & "' where p_code='" & txtpcode.Text & "' and branch='" & cbobranch & "'and Date_Entered='" & txtdateenterered.value & "'"
  cn.Execute sql

  End If
'd_Outletstock
  sql = "set dateformat dmy select Date_Entered, p_name, Quantity, OutletName from d_Outletstock where Date_Entered='" & txtdateenterered.value & "' and p_name='" & txtpname & "'AND OutletName ='" & cbobranch & "'"
  Set rs = New ADODB.Recordset
  rs.Open sql, cn
 If rs.EOF Then
  sql = ""
  sql = "set dateformat dmy insert into  d_Outletstock(Date_Entered,p_name, Quantity, OutletName)"
  sql = sql & "  values('" & txtdateenterered.value & "','" & txtpname & "'," & txtquantity.Text & ",'" & cbobranch & "')"
  cn.Execute sql
 Else
  sql = "set dateformat DMY update d_Outletstock set Quantity =" & rs.Fields("Quantity") + txtquantity.Text & " where p_name='" & txtpname & "'and Date_Entered='" & txtdateenterered.value & "' and OutletName='" & cbobranch & "'"
  cn.Execute sql
 End If

''''* to vehicle table
  If chkcustomer = 1 Then
   Provider = "MAZIWA"
   Set cn = New ADODB.Connection
  cn.Open Provider, "atm", "atm"
   'Set rs = New ADODB.Recordset
   sql = "set dateformat dmy select Vehicle, Date, Kgs, Customer from d_OutletVehicle where Vehicle ='" & cbovb & "' and Date='" & txtdateenterered.value & "'"
   Set rst = New ADODB.Recordset
   rst.Open sql, cn
    If rst.EOF Then
    sql = ""
    sql = "set dateformat dmy insert into  d_OutletVehicle(Vehicle, Date, Kgs, Customer)"
    sql = sql & "  values('" & cbovb & "','" & txtdateenterered.value & "'," & txtquantity.Text & ",'" & cbobranch & "')"
    cn.Execute sql
   Else
    sql = ""
    sql = "set dateformat DMY update d_OutletVehicle set Kgs=" & txtquantity.Text + rst.Fields("Kgs") & " where Vehicle ='" & cbovb & "' and Date='" & txtdateenterered.value & "'"
    cn.Execute sql
   End If
End If
Else
  sql = ""
  sql = "set dateformat DMY update d_Outlet set p_name='" & txtpname & "',user_id='" & User & "',Wprice=" & txtpprice & ",Rprice=" & txtsellingprice & ",type='" & w & "' where p_code='" & txtpcode.Text & "' and branch='" & cbobranch & "'"
  cn.Execute sql
End If

'txtpassit = ""
txtsellingprice = ""
txtpprice = ""
txtquantity = ""
chkrepeat.value = vbUnchecked
chkyoghurt.value = vbUnchecked
txtpname = ""
txtbalance = ""
MsgBox "Record Saved Successfully"
cmdnew.Visible = True
loadpro
'SSTab1_DblClick
SSTab2_DblClick
'Form_Load
Exit Sub
ErrorHandler:
MsgBox err.description

End Sub

Private Sub Picture1_Click()
Me.MousePointer = vbHourglass
        frmSearchGLAccounts.Show vbModal, Me
    If Continue Then
        If SearchValue <> "" Then
            lbldracc = SearchValue
            SearchValue = ""
        End If
    End If
    Me.MousePointer = 0
End Sub

Private Sub Picture2_Click()
If cbobranch = "" Then
MsgBox "Please select branch", vbInformation
Exit Sub
End If

frmSearch1.Show vbModal
'frmSearch.Show vbModal
Dim Y As String
Y = sel

If Y <> "" Then

Provider = "MAZIWA"

Set cn = New ADODB.Connection
cn.Open Provider, "atm", "atm"
sql = "select p_code, p_name, Qin, Qout, o_bal, Wprice, Rprice from d_Outlet where p_code='" & Y & "'AND Branch='" & cbobranch & "'"
'If reportpath = "" Then reportpath = GetSetting("payroll", "AppName", "rptPath", rptPath)
'sql = "select P_CODE,P_NAME,S_NO,QOUT,supplierID,pprice,sprice,QIN from ag_products where p_code='" & Y & "'AND Branch='" & cbobranch & "'"
Set rs = New ADODB.Recordset
rs.Open sql, cn
If Not rs.EOF Then
If Not IsNull(rs.Fields(0)) Then txtpcode = (rs.Fields(0))
If Not IsNull(rs.Fields(1)) Then txtpname = (rs.Fields(1))
'If Not IsNull(rs.Fields(4)) Then cbosupplier = (rs.Fields(4))
If Not IsNull(rs.Fields(5)) Then txtpprice = (rs.Fields(5))
If Not IsNull(rs.Fields(6)) Then txtsellingprice = (rs.Fields(6))
'If Not IsNull(rs.Fields(7)) Then txtreceived = (rs.Fields(7))
If Not IsNull(rs.Fields(3)) Then txtbalance = (rs.Fields(3))

If txtbalance <= 0 Then
MsgBox "Your stock is below zero please reorder", vbInformation
End If
'// check with serial no if it exist


End If
End If
End Sub

Private Sub Picture3_Click()
Me.MousePointer = vbHourglass
        frmSearchGLAccounts.Show vbModal, Me
    If Continue Then
        If SearchValue <> "" Then
            lblcracc = SearchValue
            SearchValue = ""
        End If
    End If
    Me.MousePointer = 0
End Sub

Private Sub Picture4_Click()
If cbobranch = "" Then
MsgBox "Please select branch", vbInformation
Exit Sub
End If

frmSearch1.Show vbModal
'frmSearch.Show vbModal
Dim Y As String
Y = sel

If Y <> "" Then

Provider = "MAZIWA"

Set cn = New ADODB.Connection
cn.Open Provider, "atm", "atm"
sql = "select p_code, p_name, Qin, Qout, o_bal, Wprice, Rprice from d_Outlet where p_code='" & Y & "'AND Branch='" & cbobranch & "'"
'If reportpath = "" Then reportpath = GetSetting("payroll", "AppName", "rptPath", rptPath)
'sql = "select P_CODE,P_NAME,S_NO,QOUT,supplierID,pprice,sprice,QIN from ag_products where p_code='" & Y & "'AND Branch='" & cbobranch & "'"
Set rs = New ADODB.Recordset
rs.Open sql, cn
If Not rs.EOF Then
If Not IsNull(rs.Fields(0)) Then Text1 = (rs.Fields(0))
If Not IsNull(rs.Fields(1)) Then cboproductname1 = (rs.Fields(1))
'If Not IsNull(rs.Fields(4)) Then cbosupplier = (rs.Fields(4))
If Not IsNull(rs.Fields(5)) Then Label15 = (rs.Fields(5))
If Not IsNull(rs.Fields(6)) Then Label16 = (rs.Fields(6))
'If Not IsNull(rs.Fields(7)) Then txtreceived = (rs.Fields(7))
If Not IsNull(rs.Fields(3)) Then Label18 = (rs.Fields(3))

If Label18 <= 0 Then
MsgBox "Your stock is below zero please reorder", vbInformation
End If
'// check with serial no if it exist


End If
End If
End Sub

Private Sub Braj_DblClick()
    cbobranch.Clear
    Set rst = New Recordset
    Dim cn As Connection
    Set cn = New ADODB.Connection
    Provider = "MAZIWA"
    'Provider = cnn
    cn.Open Provider, "atm", "atm"
    Set rst = New Recordset
    sql = "Select distinct(BName1) from   d_Outletbranch order by BName1"
    rst.Open sql, cn, adOpenKeyset, adLockOptimistic
    While Not rst.EOF
    cbobranch.AddItem rst.Fields(0)
    rst.MoveNext
    Wend

End Sub
Private Sub SSTab2_DblClick()
    
    cboproductname1.Clear
    Set rs = New Recordset
    Dim cn As Connection
    Set cn = New ADODB.Connection
    Provider = "MAZIWA"
    'Provider = cnn
    cn.Open Provider, "atm", "atm"
    Set rs = New Recordset
    'Dim cn As Connection
    Set cn = New ADODB.Connection
    Provider = "MAZIWA"
    'provider = cn
   cn.Open Provider, "atm", "atm"
    Set rs = New Recordset
    sql = "Select p_name from d_Outlet ORDER BY p_name"
    rs.Open sql, cn, adOpenKeyset, adLockOptimistic
    While Not rs.EOF
    cboproductname1.AddItem rs.Fields(0)
    rs.MoveNext
    Wend
End Sub
Private Sub SSTab1_DblClick()
    cbobranch.Clear
    cboproductname1.Clear
    Set rst = New Recordset
    Dim cn As Connection
    Set cn = New ADODB.Connection
    Provider = "MAZIWA"
    'Provider = cnn
    cn.Open Provider, "atm", "atm"
    Set rst = New Recordset

    sql = "Select BName1 from   d_Outletbranch ORDER BY BName1"
    rst.Open sql, cn, adOpenKeyset, adLockOptimistic
    While Not rst.EOF
    cbobranch.AddItem rst.Fields(0)
    rst.MoveNext
    Wend
    bra
    chkyoghurt.value = vbChecked
    chkyoghurt.value = vbUnchecked
'    Set rs = New Recordset
'    'Dim cn As Connection
'    Set cn = New ADODB.Connection
'    Provider = "MAZIWA"
'    'provider = cn
'   cn.Open Provider, "atm","atm"
'    Set rs = New Recordset
'    sql = "Select p_name from d_Outlet"
'    rs.Open sql, cn, adOpenKeyset, adLockOptimistic
'    While Not rs.EOF
'    cboproductname1.AddItem rs.Fields(0)
'    rs.MoveNext
'    Wend
End Sub
Private Sub cbovh_Validate(Cancel As Boolean)
Dim a As Boolean, b As Integer
Set rs = New ADODB.Recordset
sql = ""
sql = "set dateformat dmy select * from d_OutletDispatch where Vehicle= '" & ListView30.SelectedItem.SubItems(1) & "' AND OutletName='" & ListView30.SelectedItem.SubItems(2) & "' AND Date ='" & txtdateenterered.value & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
If Not IsNull(rs.Fields(0)) Then txtdateenterered = rs.Fields(0)
If Not IsNull(rs.Fields(1)) Then cbov = rs.Fields(1)
If Not IsNull(rs.Fields(2)) Then cbobranch = rs.Fields(2)
If Not IsNull(rs.Fields(3)) Then txtMilkd = rs.Fields(3)
End If
End Sub

Private Sub Text1_Change()
If KeyAscii = 13 Then
Provider = "MAZIWA"
Set cn = New ADODB.Connection
cn.Open Provider, "atm", "atm"
'If reportpath = "" Then reportpath = GetSetting("payroll", "AppName", "rptPath", rptPath)
sql = "select p_code, p_name, Qin, Qout, o_bal, Wprice, Rprice from d_Outlet where p_code='" & Text1 & "'AND Branch like'" & cbobranch & "' "
'sql = "select P_CODE,P_NAME,S_NO,QOUT,supplierid,pprice,sprice from ag_products where p_code='" & txtpcode & "'"
Set rs = New ADODB.Recordset
rs.Open sql, cn
If Not rs.EOF Then
If Not IsNull(rs.Fields(5)) Then Text1 = (rs.Fields(0))
If Not IsNull(rs.Fields(1)) Then cboproductname1 = (rs.Fields(1))
If Not IsNull(rs.Fields(5)) Then Label15 = (rs.Fields(5))
If Not IsNull(rs.Fields(5)) Then Label16 = (rs.Fields(6))
If Not IsNull(rs.Fields(6)) Then Label18 = (rs.Fields(3))

End If
End If
End Sub

Private Sub Text2_Change()
On Error GoTo ErrorHandler
chkretail.value = 0
Provider = "maziwa"
Set cn = New ADODB.Connection
cn.Open Provider, "atm", "atm"
'// check if they are in stock.
Dim rsinstock As Recordset
sql = ""
sql = "set dateformat dmy select p_code,Qout from d_Outlet where p_code='" & Text1 & "' and Branch like'" & cbobranch & "%' AND Date_Entered='" & txtdateenterered.value & "'"
'sql = "set dateformat dmy select isnull(sum(Qout),0) from d_Outlet where p_code='" & Text1 & "' and Branch like'" & cbobranch & "%' and Date_Entered>='" & Startdate & "'and Date_Entered<='" & Enddate & "'"
Set rsinstock = New ADODB.Recordset
rsinstock.Open sql, cn
'// check the stock if it is less than zero

If rsinstock.Fields(1) < 0 Then
MsgBox "Sorry Stock is Zero please re-stock before your proceed", vbInformation
Exit Sub
End If
'cmdnextitem.SetFocus
chkretail.value = vbChecked
Exit Sub
ErrorHandler:
MsgBox err.description
End Sub

Private Sub txtAmount_Change()
On Error Resume Next
'If k = 0 Then
' j = 1
' Do While Not j > (Lvwitems.ListItems.Count)
'  Lvwitems.ListItems.Item(j).selected = True
'  Z = Z + CCur(Lvwitems.SelectedItem.SubItems(2))
'  j = j + 1
' Loop
'
'sql = ""
'sql = "select p_code,Qout from d_Outlet where p_code='" & Lvwitems.SelectedItem() & "' and Branch='" & cbobranch & "'"
'Set rsinstock = New ADODB.Recordset
'rsinstock.Open sql, cn
'If rsinstock.Fields(1) < Z Then
'MsgBox "Sorry Stock cannot be less, please re-stock before your proceed", vbInformation
''n = 0
'Exit Sub
'End If
'End If
TXTCHANGE = txtamount - TXTTOTAL
    Exit Sub
SysError:
MsgBox err.description, vbInformation, Me.Caption
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

Private Sub txtdateenterered_Change()
fra1.Visible = True
loadBranchesTypes
loadpro
milkout
loaddispmilk
End Sub
Private Sub txtdateenterered_Click()
fra1.Visible = True
loadBranchesTypes
loadpro
milkout
loaddispmilk
End Sub
Private Sub milkout()
  sql = ""
  sql = "set dateformat dmy select sum(QSupplied)from d_Milkintake where TransDate='" & txtdateenterered & "'"
  Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
    If Not IsNull(rs.Fields(0)) Then Label36.Caption = rs.Fields(0)
    Else
    Label36.Caption = "0"
    End If
End Sub

Private Sub txtdateenterered_KeyPress(KeyAscii As Integer)
fra1.Visible = True
End Sub

Private Sub txtdateenterered_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
fra1.Visible = True
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

Private Sub txtMilkd_Change()
If cbov = "" Then
MsgBox "Please select Vehicle Number", vbInformation
Exit Sub
End If
'txtMilkd.SetFocus
End Sub

'Private Sub txtdateenterered_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
'fra1.Visible = True
'End Sub

Private Sub txtpcode_Change()
Startdate = DateSerial(year(txtdateenterered), month(txtdateenterered), 1)
Enddate = DateSerial(year(txtdateenterered), month(txtdateenterered) + 1, 1 - 1)

Provider = "MAZIWA"
Set cn = New ADODB.Connection
cn.Open Provider, "atm", "atm"
'If reportpath = "" Then reportpath = GetSetting("payroll", "AppName", "rptPath", rptPath)
sql = "select p_code, p_name, Qin, Qout, o_bal, Wprice, Rprice from d_Outlet where p_code='" & txtpcode & "'AND Branch='" & cbobranch & "'"
'sql = "select P_CODE,P_NAME,S_NO,QOUT,supplierid,pprice, sprice,QIN from ag_products where p_code='" & txtpcode & "'AND Branch='" & cbobranch & "' "
Set rs = New ADODB.Recordset
rs.Open sql, cn
If Not rs.EOF Then
 txtpcode = (rs.Fields(0))
 If Not IsNull(rs.Fields(1)) Then txtpname = (rs.Fields(1))
 'Dim sim As Double
 Set rst = oSaccoMaster.GetRecordset("set dateformat dmy select sum(Qout) from d_Outlet where p_code='" & txtpcode & "'AND Branch='" & cbobranch & "'AND Date_Entered>='" & Startdate & "'AND Date_Entered<='" & Enddate & "'")
 If Not IsNull(rst.Fields(0)) Then txtbalance = (rst.Fields(0))
'If Not IsNull(rs.Fields(4)) Then cbosupplier = (rs.Fields(4))
 If Not IsNull(rs.Fields(5)) Then txtpprice = (rs.Fields(5))
 If Not IsNull(rs.Fields(6)) Then txtsellingprice = (rs.Fields(6))
 End If
 If KeyAscii = 13 Then
txtpcodeO_Change
'txtpcode11_KeyPress
Else
Exit Sub
End If
End Sub

Private Sub txtpname_Click()
If cbobranch = "" Then
MsgBox "Please select the Outlet", vbInformation
Exit Sub
End If

Provider = "MAZIWA"
Set cn = New ADODB.Connection
cn.Open Provider, "atm", "atm"
Set rst = New ADODB.Recordset
rst.Open sql, cn
'If rs.EOF Then
Set rst = oSaccoMaster.GetRecordset("select p_code from d_Outlet where p_name ='" & txtpname & "' and branch='" & cbobranch & "'")
If Not rst.EOF Then
txtpcode = rst.Fields("p_code")
'Y = txtpcode.Text
txtpcode_Change
'txtpcodeO_Change
''txtsel
End If
End Sub

Private Sub txtquantity_Change()
'******************check if aready disptch********************
   If chkrepeat = 0 Then
    sql = ""
    sql = "SET dateformat dmy     SELECT * FROM  d_Outlet  WHERE     Date_Entered = '" & txtdateenterered & "'and p_code = '" & txtpcode & "' and Qin = '" & txtquantity & "'and Branch = '" & cbobranch & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
      If MsgBox("You have already dispatch for that day, Do you want to continue receivinmg? ", vbYesNo) = vbYes Then
      Else
       Exit Sub
      End If
    'MsgBox "You have already dispatch for that day", vbInformation
    'txtquantity.SetFocus
    'Exit Sub
    End If
 End If
'*******************end ************************
End Sub

Private Sub txttotal_Change()
On Error Resume Next
TXTCHANGE = txtamount - TXTTOTAL
End Sub
