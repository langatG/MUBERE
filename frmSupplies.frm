VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSupplies 
   BackColor       =   &H80000013&
   Caption         =   "Supplies Details"
   ClientHeight    =   9360
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9360
   ScaleWidth      =   15120
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   7680
      TabIndex        =   51
      Top             =   9000
      Width           =   855
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   6600
      TabIndex        =   22
      Top             =   9000
      Width           =   855
   End
   Begin VB.CheckBox chkActive 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Caption         =   "Active"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   10200
      TabIndex        =   17
      Top             =   3600
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.TextBox txtNames 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8040
      TabIndex        =   9
      Top             =   1200
      Width           =   4935
   End
   Begin VB.ComboBox cboLocation 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmSupplies.frx":0000
      Left            =   2520
      List            =   "frmSupplies.frx":0002
      TabIndex        =   4
      Top             =   2880
      Width           =   3255
   End
   Begin VB.ComboBox cboDistrict 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmSupplies.frx":0004
      Left            =   2520
      List            =   "frmSupplies.frx":0006
      TabIndex        =   3
      Top             =   2280
      Width           =   3255
   End
   Begin VB.TextBox txtIdNumber 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2520
      MaxLength       =   8
      TabIndex        =   2
      Top             =   1680
      Width           =   3255
   End
   Begin VB.TextBox txtEMail 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8040
      TabIndex        =   10
      Top             =   1680
      Width           =   4935
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   375
      Left            =   5040
      TabIndex        =   50
      Top             =   9000
      Width           =   735
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   375
      Left            =   5760
      TabIndex        =   49
      Top             =   9000
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Caption         =   "Bank Details"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   120
      TabIndex        =   31
      Top             =   6240
      Width           =   11655
      Begin VB.TextBox txtcomments 
         Appearance      =   0  'Flat
         Height          =   1215
         Left            =   7680
         MultiLine       =   -1  'True
         TabIndex        =   85
         Top             =   1200
         Width           =   2535
      End
      Begin VB.ComboBox cboBBranch 
         Height          =   315
         ItemData        =   "frmSupplies.frx":0008
         Left            =   2400
         List            =   "frmSupplies.frx":000A
         TabIndex        =   21
         Top             =   1920
         Width           =   4695
      End
      Begin VB.ComboBox cboBankName 
         Height          =   315
         ItemData        =   "frmSupplies.frx":000C
         Left            =   2400
         List            =   "frmSupplies.frx":000E
         TabIndex        =   19
         Top             =   600
         Width           =   4695
      End
      Begin VB.TextBox txtAccNumber 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         TabIndex        =   20
         Top             =   1200
         Width           =   4695
      End
      Begin VB.Label Label37 
         Caption         =   "Remarks"
         Height          =   255
         Left            =   8880
         TabIndex        =   86
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label34 
         BackColor       =   &H80000000&
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   135
         Left            =   1680
         TabIndex        =   79
         Top             =   600
         Width           =   135
      End
      Begin VB.Label Label33 
         BackColor       =   &H80000000&
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   135
         Left            =   2280
         TabIndex        =   78
         Top             =   1200
         Width           =   135
      End
      Begin VB.Label lblLoan 
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "No Loan"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   495
         Left            =   7440
         TabIndex        =   55
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Account Number :"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   34
         Top             =   1080
         Width           =   2235
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Bank Name :"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   33
         Top             =   480
         Width           =   1530
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Bank Branch :"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   32
         Top             =   1920
         Width           =   1665
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Suppliers Details"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   15495
      Begin VB.TextBox txtrefee 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   11760
         TabIndex        =   98
         Top             =   3960
         Width           =   1095
      End
      Begin VB.TextBox txtbonus 
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
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   8880
         TabIndex        =   96
         Top             =   4320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CheckBox chkbonus 
         Caption         =   "Bonus"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6600
         TabIndex        =   94
         Top             =   4440
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.ComboBox cbocontcode 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmSupplies.frx":0010
         Left            =   8160
         List            =   "frmSupplies.frx":0020
         TabIndex        =   92
         Top             =   3960
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5640
         Picture         =   "frmSupplies.frx":0047
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   87
         Top             =   4200
         Width           =   255
      End
      Begin VB.Frame Frame3 
         Caption         =   "TCHP DETAILS"
         Height          =   1455
         Left            =   120
         TabIndex        =   62
         Top             =   4680
         Visible         =   0   'False
         Width           =   14775
         Begin VB.CheckBox chkinoutpatient 
            Caption         =   "COMPREHENSIVE"
            Height          =   255
            Left            =   2520
            TabIndex        =   90
            Top             =   1080
            Width           =   1815
         End
         Begin VB.CheckBox chkoutpatient 
            Caption         =   "BASIC"
            Height          =   255
            Left            =   480
            TabIndex        =   89
            Top             =   1080
            Width           =   1455
         End
         Begin VB.CheckBox chkinpatient 
            Caption         =   "OTHERS"
            Height          =   255
            Left            =   5040
            TabIndex        =   88
            Top             =   1080
            Width           =   1215
         End
         Begin VB.ComboBox txtthcppremium 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2760
            Locked          =   -1  'True
            TabIndex        =   84
            Top             =   240
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.ComboBox cbostatus 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmSupplies.frx":0309
            Left            =   10800
            List            =   "frmSupplies.frx":031F
            Locked          =   -1  'True
            TabIndex        =   82
            Text            =   "New"
            Top             =   720
            Width           =   1935
         End
         Begin VB.CommandButton cmdtchpp 
            Caption         =   "Change TCHP Premium"
            Height          =   255
            Left            =   4680
            TabIndex        =   81
            Top             =   720
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.TextBox txtaartkno 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   10800
            TabIndex        =   74
            Top             =   360
            Width           =   1935
         End
         Begin MSComCtl2.DTPicker DTPend 
            Height          =   255
            Left            =   8400
            TabIndex        =   72
            Top             =   720
            Visible         =   0   'False
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   450
            _Version        =   393216
            Format          =   110493697
            CurrentDate     =   40916
         End
         Begin MSComCtl2.DTPicker DTPstart 
            Height          =   255
            Left            =   8520
            TabIndex        =   71
            Top             =   720
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            Format          =   110493697
            CurrentDate     =   40916
         End
         Begin VB.TextBox txtapremium 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1920
            TabIndex        =   66
            Top             =   720
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.CheckBox chkthcp 
            Caption         =   "TCHP"
            Height          =   255
            Left            =   120
            TabIndex        =   64
            Top             =   360
            Width           =   735
         End
         Begin VB.CheckBox chkthcpactive 
            Caption         =   "TCHP Active"
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
            Left            =   4680
            TabIndex        =   63
            Top             =   360
            Visible         =   0   'False
            Width           =   1695
         End
         Begin MSComCtl2.DTPicker dtptmd 
            Height          =   255
            Left            =   8400
            TabIndex        =   68
            Top             =   360
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   450
            _Version        =   393216
            CalendarBackColor=   49152
            Format          =   110493697
            CurrentDate     =   40095
         End
         Begin VB.Label Label36 
            Caption         =   "Status"
            Height          =   255
            Left            =   9960
            TabIndex        =   83
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label31 
            BackColor       =   &H80000000&
            Caption         =   "*"
            ForeColor       =   &H000000FF&
            Height          =   135
            Left            =   8040
            TabIndex        =   76
            Top             =   360
            Width           =   135
         End
         Begin VB.Label Label20 
            Caption         =   "AAR TK#"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   9960
            TabIndex        =   75
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label30 
            Caption         =   "End Date"
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
            Left            =   7320
            TabIndex        =   73
            Top             =   720
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Label29 
            Caption         =   "Start Date"
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
            Left            =   7560
            TabIndex        =   70
            Top             =   720
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Label21 
            Caption         =   "TCHP Member Date"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6360
            TabIndex        =   69
            Top             =   360
            Width           =   2055
         End
         Begin VB.Label Label22 
            Caption         =   "Annual TCHP Premium"
            Height          =   255
            Left            =   120
            TabIndex        =   67
            Top             =   720
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.Label thcppremiumlabel 
            Caption         =   "Monthly TCHP Premium"
            Height          =   255
            Left            =   960
            TabIndex        =   65
            Top             =   360
            Visible         =   0   'False
            Width           =   1695
         End
      End
      Begin VB.TextBox txtRate 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Left            =   11280
         TabIndex        =   15
         Top             =   3000
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.PictureBox Picture5 
         Height          =   255
         Left            =   5640
         Picture         =   "frmSupplies.frx":0357
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   53
         Top             =   840
         Width           =   255
      End
      Begin VB.ComboBox cbobrnch 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmSupplies.frx":0619
         Left            =   2400
         List            =   "frmSupplies.frx":061B
         TabIndex        =   7
         Top             =   4200
         Width           =   3255
      End
      Begin VB.CheckBox chkTrader 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "Trader"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   11520
         TabIndex        =   18
         Top             =   3480
         Width           =   1215
      End
      Begin VB.ComboBox cboTCode 
         Height          =   315
         ItemData        =   "frmSupplies.frx":061D
         Left            =   8160
         List            =   "frmSupplies.frx":0627
         TabIndex        =   14
         Top             =   3000
         Width           =   1815
      End
      Begin VB.PictureBox picMemSign 
         Height          =   1095
         Left            =   13080
         ScaleHeight     =   1035
         ScaleWidth      =   1755
         TabIndex        =   45
         Top             =   3000
         Width           =   1815
      End
      Begin VB.CommandButton cmdShowSignature 
         Height          =   375
         Left            =   13080
         Picture         =   "frmSupplies.frx":063C
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Add Signature"
         Top             =   4200
         Width           =   375
      End
      Begin VB.CommandButton cmdDeleteSignature 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   14520
         Picture         =   "frmSupplies.frx":0B6E
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Delete Signature"
         Top             =   4080
         Width           =   375
      End
      Begin VB.CommandButton txtClosePic 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   14400
         Picture         =   "frmSupplies.frx":0C70
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Delete Photo"
         Top             =   2520
         Width           =   375
      End
      Begin VB.CommandButton cmdOpenPic 
         Height          =   375
         Left            =   13320
         Picture         =   "frmSupplies.frx":0D72
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Add Photo"
         Top             =   2520
         Width           =   375
      End
      Begin VB.PictureBox picMemPhoto 
         Height          =   1860
         Left            =   12960
         ScaleHeight     =   1800
         ScaleWidth      =   1755
         TabIndex        =   40
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtTown 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         TabIndex        =   6
         Top             =   3840
         Width           =   3255
      End
      Begin VB.TextBox txtVillage 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7920
         TabIndex        =   13
         Top             =   2640
         Width           =   4935
      End
      Begin VB.TextBox txtDivision 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7920
         TabIndex        =   12
         Top             =   2280
         Width           =   4935
      End
      Begin VB.TextBox txtPNumber 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7920
         MaxLength       =   12
         TabIndex        =   11
         Top             =   1920
         Width           =   4935
      End
      Begin VB.TextBox txtSNo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         MaxLength       =   8
         TabIndex        =   8
         Top             =   840
         Width           =   3255
      End
      Begin VB.ComboBox cboType 
         Height          =   315
         ItemData        =   "frmSupplies.frx":12A4
         Left            =   7920
         List            =   "frmSupplies.frx":12B1
         TabIndex        =   16
         Top             =   3480
         Width           =   1935
      End
      Begin VB.TextBox txtPAddress 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         TabIndex        =   5
         Top             =   3360
         Width           =   3255
      End
      Begin MSComCtl2.DTPicker DTPRegDate 
         Height          =   255
         Left            =   7680
         TabIndex        =   38
         Top             =   600
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   450
         _Version        =   393216
         CalendarBackColor=   49152
         Format          =   110493697
         CurrentDate     =   40095
      End
      Begin VB.Label Label39 
         Caption         =   "Registration Fee"
         Enabled         =   0   'False
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
         Left            =   10080
         TabIndex        =   97
         Top             =   4080
         Width           =   1695
      End
      Begin VB.Label Label19 
         Caption         =   "Rate Per kg :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7800
         TabIndex        =   95
         Top             =   4440
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label27 
         BackColor       =   &H80000000&
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   135
         Index           =   1
         Left            =   8040
         TabIndex        =   93
         Top             =   3120
         Width           =   135
      End
      Begin VB.Label Label38 
         Caption         =   "Payment Frequency"
         Enabled         =   0   'False
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
         Left            =   6120
         TabIndex        =   91
         Top             =   3960
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label Label35 
         BackColor       =   &H80000000&
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   135
         Left            =   1200
         TabIndex        =   80
         Top             =   2880
         Width           =   135
      End
      Begin VB.Label Label32 
         BackColor       =   &H80000000&
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   135
         Left            =   7680
         TabIndex        =   77
         Top             =   1920
         Width           =   135
      End
      Begin VB.Label Label28 
         BackColor       =   &H80000000&
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   135
         Left            =   7560
         TabIndex        =   61
         Top             =   600
         Width           =   135
      End
      Begin VB.Label Label27 
         BackColor       =   &H80000000&
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   135
         Index           =   0
         Left            =   7200
         TabIndex        =   60
         Top             =   3480
         Width           =   135
      End
      Begin VB.Label Label26 
         BackColor       =   &H80000000&
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   135
         Left            =   7320
         TabIndex        =   59
         Top             =   1080
         Width           =   135
      End
      Begin VB.Label Label25 
         BackColor       =   &H80000000&
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   135
         Left            =   960
         TabIndex        =   58
         Top             =   4200
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label Label24 
         BackColor       =   &H80000000&
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   135
         Left            =   1560
         TabIndex        =   57
         Top             =   1680
         Width           =   135
      End
      Begin VB.Label Label23 
         BackColor       =   &H80000000&
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   135
         Left            =   2280
         TabIndex        =   56
         Top             =   840
         Width           =   135
      End
      Begin VB.Label lblRatw 
         AutoSize        =   -1  'True
         Caption         =   "Rate Per Kg "
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
         Left            =   10080
         TabIndex        =   54
         Top             =   3000
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.Label lblbranch 
         AutoSize        =   -1  'True
         Caption         =   "Branch"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   52
         Top             =   4080
         Width           =   840
      End
      Begin VB.Label Label18 
         Caption         =   "Gender :"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6360
         TabIndex        =   48
         Top             =   3360
         Width           =   855
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Signature"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   13440
         TabIndex        =   47
         Top             =   4200
         Width           =   1035
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Suppliers Photo"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   13080
         TabIndex        =   46
         Top             =   240
         Width           =   1800
      End
      Begin VB.Label Label13 
         Caption         =   "Town"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   39
         Top             =   3720
         Width           =   855
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Postal address"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   37
         Top             =   3240
         Width           =   1665
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Reg Date :"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   6360
         TabIndex        =   36
         Top             =   480
         Width           =   1230
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Payment Frequency:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   6000
         TabIndex        =   35
         Top             =   3000
         Width           =   1965
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "District"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   2160
         Width           =   1020
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Village"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6360
         TabIndex        =   29
         Top             =   2520
         Width           =   870
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Division"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6360
         TabIndex        =   28
         Top             =   2160
         Width           =   1050
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Phone No."
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6360
         TabIndex        =   27
         Top             =   1800
         Width           =   1290
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "E - Mail :"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6360
         TabIndex        =   26
         Top             =   1440
         Width           =   1125
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Names :"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6360
         TabIndex        =   25
         Top             =   960
         Width           =   960
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Location"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   2760
         Width           =   1035
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Id Number"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   1560
         Width           =   1350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Supplier Number "
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   2175
      End
   End
   Begin VB.Image Image1 
      Height          =   1140
      Left            =   11880
      Picture         =   "frmSupplies.frx":12CB
      Top             =   6960
      Width           =   1890
   End
End
Attribute VB_Name = "frmSupplies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim thcp As Integer
Dim thcpactive As Integer
Dim newb As Integer
Dim types As String

Private Sub cboBankName_Change()
Dim ACCNO As String
Dim supno As String
'If cboBankName = "FSA" Then
'    txtAccNumber.Enabled = False
'    supno = txtSNo
'    supno = Format(supno, "00000")
'    ACCNO = "010100S" & supno & "00"
'    txtAccNumber = ACCNO
'Else
    'txtAccNumber.Enabled = True
'End If
End Sub

Private Sub cboBankName_Click()
'cboBankName_Change
End Sub

Private Sub cboBankName_KeyPress(KeyAscii As Integer)
'MsgBox "Please select do not edit."
'KeyAscii = 0
Dim ACCNO As String
'If cboBankName = "FSA" Then
'    ACCNO = "010100S" & txtSNo & "00"
'    txtAccNumber = ACCNO
'
'End If

End Sub

Private Sub cbobrnch_Change1()
Dim rsr As New ADODB.Recordset
Dim rsg As New ADODB.Recordset
Dim gts As New ADODB.Recordset
Dim I As Object
Dim Mylength As Integer
Dim prefix As String
sql = ""
sql = "SELECT     brcode  FROM         d_company"
Set gts = oSaccoMaster.GetRecordset(sql)
If Not gts.EOF Then
prefix = Trim(gts.Fields(0))
Else
prefix = "KIP"
End If
mysql = ""
mysql = "select GenerateReceiptno from param"

Set rsg = oSaccoMaster.GetRecordset(mysql)
If Not rsg.EOF Then
    ''''check check
    If rsg!GenerateReceiptno = True Then
    
        mysql = ""
        mysql = "select * from sno where receiptno like '" & prefix & "%' order by Receipthnoid desc"
        
        Set rsr = oSaccoMaster.GetRecordset(mysql)
        
        If Not rsr.EOF Then
            Mylength = CInt(mid(rsr!ReceiptNo, 4, 8))
            Mylength = Mylength + 1
            txtSNo = Padding(Mylength)
            txtSNo = prefix & txtSNo
        Else
            Mylength = 1
            txtSNo = prefix & Padding(Mylength)
            
        End If
Else
    ''//receiptno  will be keyed in
End If
End If
End Sub

Private Sub cbobrnch_Validate(Cancel As Boolean)
'cbobrnch_Change

End Sub
Private Sub cbobrnch_Click()
'cbobrnch_Change

End Sub
Private Sub all()
    cbobrnch.Clear
    Set rs = New Recordset
    'Dim cn As Connection
    Set cn = New ADODB.Connection
    Provider = "MAZIWA"
    'provider = cn
   cn.Open Provider, "atm", "atm"
    Set rs = New Recordset
    sql = "Select distinct(BName) from d_Branch order by BName"
    rs.Open sql, cn, adOpenKeyset, adLockOptimistic
    While Not rs.EOF
    cbobrnch.AddItem rs.Fields(0)
    rs.MoveNext
    Wend
End Sub
Private Sub chkbonus_Click()
If chkbonus = vbChecked Then
txtbonus.Visible = True
Label19.Visible = True
Else
txtbonus.Visible = False
Label19.Visible = False
End If
End Sub

Private Sub chkinoutpatient_Click()
If chkthcp = vbChecked Then
chkinpatient = vbUnchecked
chkoutpatient = vbUnchecked
End If
txtthcppremium.Clear
Set rs = CreateObject("adodb.recordset")
    If chkinoutpatient = vbChecked Then
    sql = "SELECT RATE FROM Tchp_Rate where type='COMPREHENSIVE'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    
    If rs.EOF Then Exit Sub
    
    With rs
        
        While Not .EOF
         
         txtthcppremium.AddItem rs.Fields(0)
         
         
         .MoveNext
        
        Wend
    
    End With
    Else
      sql = "SELECT RATE FROM Tchp_Rate where type<>'COMPREHENSIVE'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    
    If rs.EOF Then Exit Sub
    
    With rs
        
        While Not .EOF
         
         txtthcppremium.AddItem rs.Fields(0)
         
         
         .MoveNext
        
        Wend
    
    End With
    End If

End Sub

Private Sub chkinoutpatient_Validate(Cancel As Boolean)
chkinoutpatient_Click
End Sub

Private Sub chkinpatient_Click()
If chkinpatient = vbChecked Then
chkinoutpatient = vbUnchecked
chkoutpatient = vbUnchecked
End If

Set rs = CreateObject("adodb.recordset")
txtthcppremium.Clear
    If chkinpatient = vbChecked Then
    sql = "SELECT RATE FROM Tchp_Rate where type='OTHERS'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    
    If rs.EOF Then Exit Sub
    
    With rs
        
        While Not .EOF
         
         txtthcppremium.AddItem rs.Fields(0)
         
         
         .MoveNext
        
        Wend
    
    End With
    Else
      sql = "SELECT RATE FROM Tchp_Rate where type<>'OTHERS'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    
    If rs.EOF Then Exit Sub
    
    With rs
        
        While Not .EOF
         
         txtthcppremium.AddItem rs.Fields(0)
         
         
         .MoveNext
        
        Wend
    
    End With
    End If
End Sub

Private Sub chkinpatient_Validate(Cancel As Boolean)
chkinpatient_Click
End Sub

Private Sub chkoutpatient_Click()
Set rs = CreateObject("adodb.recordset")
If chkthcp = vbChecked Then
chkinoutpatient = vbUnchecked
chkinpatient = vbUnchecked
End If
txtthcppremium.Clear
    If chkoutpatient = vbChecked Then
    sql = "SELECT RATE FROM Tchp_Rate where type='BASIC'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    
    If rs.EOF Then Exit Sub
    
    With rs
        
        While Not .EOF
         
         txtthcppremium.AddItem rs.Fields(0)
         
         
         .MoveNext
        
        Wend
    
    End With
    Else
      sql = "SELECT RATE FROM Tchp_Rate where type<>'BASIC'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    
    If rs.EOF Then Exit Sub
    
    With rs
        
        While Not .EOF
         
         txtthcppremium.AddItem rs.Fields(0)
         
         
         .MoveNext
        
        Wend
    
    End With
    End If
End Sub

Private Sub chkoutpatient_Validate(Cancel As Boolean)
chkoutpatient_Click
End Sub

Private Sub chkthcp_Click()
If chkthcp = vbChecked Then
thcppremiumlabel.Visible = True
txtthcppremium.Visible = True
thcpactive = 1
thcp = 1

Else
thcppremiumlabel.Visible = False
txtthcppremium.Visible = False
thcpactive = 0
thcp = 0

End If
End Sub

Private Sub chkthcp_Validate(Cancel As Boolean)
'//get power to untick the checkbox
If newb = 0 Then
sql = ""
sql = "SELECT     *  FROM         d_Suppliers WHERE     (thcpactive = 1) AND (SNo = '" & txtSNo & "')"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
'//check the powers to unchange it now.
sql = "SELECT     *   FROM         UserAccounts where userloginid='" & User & "' and superuser=1"
Set rst = oSaccoMaster.GetRecordset(sql)
If Not rst.EOF Then
'//is allowed to uncheck
Else
'//not allowed to uncheck
MsgBox "You are not allowed to uncheck this status, please seek more advice from your supervisor", vbInformation
chkthcp = vbChecked
Exit Sub
End If
End If
End If
End Sub

Private Sub chkthcpactive_Click()
If chkthcpactive = vbChecked Then
thcpactive = 1
Else
thcpactive = 0
End If
End Sub

Private Sub cmdclose_Click()
Unload Me
End Sub


Private Sub cmdedit_Click()
newb = 0
txtSNo.Locked = False
txtEMail.Locked = False
txtIdNumber.Locked = False
txtNames.Locked = False
txtPAddress.Locked = False
txtPNumber.Locked = False
txtAccNumber.Locked = False
txtDivision.Locked = False
txtVillage.Locked = False
txtTown.Locked = False
txtaartkno.Locked = False
txtthcppremium.Locked = False
cboBBranch.Locked = False
cboBankName.Locked = False
cboLocation.Locked = False
cbobrnch.Locked = False
cboDistrict.Locked = False
cboTCode.Locked = False
cboType.Locked = False
'cmdEdit.Enabled = False
'cmdSave.Enabled = False
cmdSave.Enabled = True
cmdEdit.Enabled = False
cmdNew.Enabled = False

txtSNo_Validate True
End Sub

Private Sub cmdNew_Click()
Set rs = oSaccoMaster.GetRecordset("d_sp_SNO")
If Not rs.EOF Then
'txtSNo = rs.Fields(0) + 1
'Else
'txtSNo = 1
End If
'cbobrnch_Change1
newb = 1
txtEMail = ""
txtIdNumber = ""
txtNames = ""
txtPAddress = ""
txtPNumber = ""
txtAccNumber = ""
txtDivision = ""
txtVillage = ""
txtTown = ""
'txtthcppremium.Clear
cboBBranch.Text = ""
cboBankName.Text = ""
cboLocation.Text = ""
'cbobrnch.Text = ""
cboDistrict.Text = ""
cboTCode.Text = ""
cboType.Text = ""
txtthcppremium.Locked = False
txtSNo.Locked = False
txtEMail.Locked = False
txtIdNumber.Locked = False
txtNames.Locked = False
txtPAddress.Locked = False
txtPNumber.Locked = False
txtAccNumber.Locked = False
txtDivision.Locked = False
txtVillage.Locked = False
txtTown.Locked = False

cboBBranch.Locked = False
cboBankName.Locked = False
cboLocation.Locked = False
cbobrnch.Locked = False
cboDistrict.Locked = False
cboTCode.Locked = False
cboType.Locked = False
'cmdEdit.Enabled = False
'cmdSave.Enabled = False
cmdSave.Enabled = True
cmdNew.Enabled = False
chkthcp = vbUnchecked

End Sub

Private Sub cmdsave_Click()
Dim Active, Trader As String
On Error GoTo ErrorHandler
If chkinoutpatient = vbChecked Then
types = "COMPREHENSIVE"
End If
If chkinpatient = vbChecked Then
types = "OTHERS"
End If
If chkoutpatient = vbChecked Then
types = "BASIC"
End If
If txtSNo = "" Then
MsgBox "Please enter the supplier number ", vbInformation, "Missing Information"
txtSNo.SetFocus
Exit Sub
End If
'If txtrefee = "0" Then
'MsgBox "Please enter the Members Registration fee", vbInformation, "Missing Information"
'txtrefee.SetFocus
'Exit Sub
'End If
If chkbonus = vbChecked Then
If txtbonus <= 0 Then
MsgBox "Please enter the rate for Bonus", vbInformation, "Missing Information"
txtbonus.SetFocus
Exit Sub
End If
End If

'If Len(txtSNo) < 8 Then 'kpk00011
'MsgBox "SNO number length cannot be less than eight", vbInformation
'Exit Sub
'End If

If Len(Trim$(txtIdNumber)) < 7 Then
MsgBox "IDNO number length cannot be less than seven digits", vbInformation
Exit Sub
End If
If cboTCode = "" Then
MsgBox "Please enter the Payment Frequency  ", vbInformation, "Missing Information"
cboTCode.SetFocus
'Exit Sub
Exit Sub
End If
 
'If Not IsNumeric(txtPNumber) Then
' MsgBox "Please enter Valid Phone Number"
' Exit Sub
' End If
' If Not IsNumeric(txtAccNumber) Then
' MsgBox "Please enter Valid Account Number"
' Exit Sub
' End If
 
If cboLocation = "" Then
MsgBox "Please enter the location ", vbInformation, "Missing Information"
Exit Sub
Exit Sub
End If

If txtNames = "" Then
MsgBox "Please enter the supplier name ", vbInformation, "Missing Information"
txtSNo.SetFocus
Exit Sub
End If

If cbobrnch = "" Then
 MsgBox "Please select the branch name ", vbInformation, "Missing Information"
 cbobrnch.SetFocus
 Exit Sub
End If

If txtIdNumber = "" Then
MsgBox "Please enter the Identity number ", vbInformation, "Missing Information"
txtSNo.SetFocus
Exit Sub
End If

If cboType = "" Then
MsgBox "Please select Type", vbInformation, "Missing Information"
txtSNo.SetFocus
Exit Sub
End If
If Trim(txtRate) = "" Then
txtRate = "0"
End If
'If Trim(cbobrnch) = "" Then
'MsgBox "Please enter the branch"
'cbobrnch.SetFocus
'End If
'If txtPNumber = "" Then
'MsgBox "The phoneNo is a mandatory field., please fill it before you proceed", vbInformation
'Exit Sub
'End If
If txtAccNumber = "" Then
MsgBox "The Bank Details are mandatory field., please fill it before you proceed", vbInformation
'Exit Sub
End If

If cboBankName = "" Then
MsgBox "The bank name is a mandatory field., please fill it before you proceed", vbInformation
'Exit Sub
End If
'chkthcp

If chkActive.value = vbChecked Then
    Active = "1"
Else
    Active = "0"
End If

If chkTrader.value = vbChecked Then
    Trader = "1"
Else
    Trader = "0"
End If
If chkthcp = vbChecked Then
If txtaartkno = "" Then
'MsgBox "AAR NO cannot be blank if the member is a tchp member", vbInformation
'Exit Sub
End If
End If

If chkthcp = vbChecked Then
If txtthcppremium = "" Then
MsgBox "premium amount cannot be blank if the member is a tchp member", vbInformation
Exit Sub
End If
End If

If chkthcp = vbChecked Then
If txtthcppremium <= 0 Then
MsgBox "premium amount cannot be zero if the member is a tchp member", vbInformation
Exit Sub
End If
End If

'If Len(txtPNumber) < 10 Then
'MsgBox "phone number length cannot be less than ten", vbInformation
'Exit Sub
'End If
If txtthcppremium = "" Then txtthcppremium = 0


'dtptmd = ""



If txtaartkno = "" Then txtaartkno = "NA"
'//Save to transport table
'sql = "set dateformat dmy select trans_code,sno,active from d_transport where  sno=" & txtSNo & ""
'Set rs = oSaccoMaster.GetRecordset(sql)
'If Not Trim(cboTCode) = "" Then
'If rs.EOF Then
'sql = "d_sp_TransAssign '" & cboTCode & "'," & txtSNo & "," & txtrate & ",'" & DTPRegDate & "','" & User & "'"
'oSaccoMaster.ExecuteThis (sql)
'End If
'End If
'cboTCode = " "
'//check if the id not is the already available

Dim idno As String, Phone As String, ans As String, NAMES As String
If newb = 1 Then
sql = "SELECT     IdNo, AccNo, SNo, PhoneNo  FROM         d_Suppliers  WHERE     (SNo = '" & txtSNo & "')"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
'check if the idno exists
        sql = "SELECT     IdNo, AccNo, SNo, PhoneNo  FROM         d_Suppliers  WHERE     (idno = '" & txtIdNumber & "') "
        Set rst = oSaccoMaster.GetRecordset(sql)
        If Not rst.EOF Then
        'ans = MsgBox("The ID no is already available in the system, do you want to proceed with this update", vbYesNo)
        MsgBox "ID No Cannot be duplicate "
        Exit Sub
            'If ans = vbYes Then
            'GoTo next1
            'Else
            'Exit Sub
            'End If
        End If
        
        
next1:
        sql = "SELECT     IdNo, AccNo, SNo, PhoneNo  FROM         d_Suppliers  WHERE     (phoneno = '" & txtPNumber & "')"
        Set rst = oSaccoMaster.GetRecordset(sql)
        If Not rst.EOF Then
        ans = MsgBox("The Phone no is already available in the system, do you want to proceed with this update", vbYesNo)
            If ans = vbYes Then
            GoTo next2
            Else
            Exit Sub
            End If
        End If
next2:
        sql = "SELECT     IdNo, AccNo, SNo, PhoneNo  FROM         d_Suppliers  WHERE     (accno = '" & txtAccNumber & "')"
        Set rst = oSaccoMaster.GetRecordset(sql)
        If Not rst.EOF Then
        ans = MsgBox("The Accno no is already available in the system, do you want to proceed with this update", vbYesNo)
            If ans = vbYes Then
            GoTo next3
            Else
            Exit Sub
            End If
        End If
next3:
       sql = "SELECT     IdNo, AccNo, SNo, PhoneNo,NAMES  FROM         d_Suppliers  WHERE     (NAMES = '" & txtNames & "')"
        Set rst = oSaccoMaster.GetRecordset(sql)
        If Not rst.EOF Then
        ans = MsgBox("The Names is already available in the system, do you want to proceed with this update", vbYesNo)
            If ans = vbYes Then
            GoTo next4
            Else
            Exit Sub
            End If
        End If
next4:
'check if the phone exists with someone else
Else
      sql = "SELECT     IdNo, AccNo, SNo,names, PhoneNo  FROM         d_Suppliers  WHERE     (idno = '" & txtIdNumber & "') "
        Set rst = oSaccoMaster.GetRecordset(sql)
        If Not rst.EOF Then
         MsgBox "The ID no is already available to '" & rst!NAMES & "' supplyno '" & rst!sno & "' in the system, You cannot proceed any further", vbInformation
            'If ans = vbYes Then
            'GoTo next5
            'Else
            Exit Sub
            'End If
        End If
next5:
sql = "SELECT     IdNo, AccNo, SNo, PhoneNo  FROM         d_Suppliers  WHERE     (phoneno = '" & txtPNumber & "') "
        Set rst = oSaccoMaster.GetRecordset(sql)
        If Not rst.EOF Then
        ans = MsgBox("The Phone no is already available in the system, do you want to proceed with this update", vbYesNo)
            If ans = vbYes Then
            GoTo next6
            Else
            Exit Sub
            End If
        End If
next6:
        sql = "SELECT     IdNo, AccNo, SNo, PhoneNo,NAMES  FROM         d_Suppliers  WHERE     (accno = '" & txtAccNumber & "') "
        Set rst = oSaccoMaster.GetRecordset(sql)
        If Not rst.EOF Then
        ans = MsgBox("The Accno no is already available in the system, do you want to proceed with this update", vbYesNo)
            If ans = vbYes Then
            Dim n As String, idnooo As String, sno78 As String
            n = rst.Fields(4)
            idnooo = rst.Fields(0)
            sno78 = rst.Fields(2)
            'If txtcomments = "" Then
            'MsgBox "Please enter comments before you proceed", vbInformation
            'txtcomments = "The account is being used by " & n & "and IDno: " & idnooo & " And Supplier No: " & sno78 & ""
             'Exit Sub
           ' Else
            GoTo next7
            'End If
            
            Else
            Exit Sub
            End If
        End If
next7:

       sql = "SELECT     IdNo, AccNo, SNo, PhoneNo,NAMES  FROM         d_Suppliers  WHERE     (NAMES = '" & txtNames & "')"
        Set rst = oSaccoMaster.GetRecordset(sql)
        If Not rst.EOF Then
        ans = MsgBox("The Names is already available in the system, do you want to proceed with this update", vbYesNo)
            If ans = vbYes Then
            GoTo next8
            Else
            Exit Sub
            End If
        End If
next8:
End If
End If

'Save to suppliers table
Set cn = New ADODB.Connection
sql = "d_sp_Suppliers " & txtSNo & ",'" & DTPRegDate & "','" & txtIdNumber & "','" & txtNames & "','" & txtAccNumber & "','" & cboBankName & "','" & cboBBranch & "','" & cboType & "','" & txtVillage & "','" & cboLocation & "','" & txtDivision & "','" & cboDistrict & "'," & Trader & "," & Active & ",0,'" & cbobrnch & "','" & txtPNumber & "','" & txtPAddress & "','" & txtTown & "','" & txtEMail & "','" & cboTCode & "','" & "Sign" & "','" & "Photo" & "','" & User & "','" & cbocontcode & "','" & txtaartkno & "','" & dtptmd & "'," & thcp & "," & thcpactive & "," & txtthcppremium
'sql = "d_sp_Suppliers " & txtSNo & ",'" & DTPregdate & "','" & txtIdNumber & "','" & txtNames & "','" & txtAccNumber & "','" & cboBankName & "','" & cboBBranch & "','" & cboType & "','" & txtVillage & "','" & cboLocation & "','" & txtDivision & "','" & cboDistrict & "'," & Trader & "," & Active & ",'" & cbobrnch & "','" & txtPNumber & "','" & txtPAddress & "','" & txtTown & "','" & txtEMail & "','" & cboTCode & "','" & "Sign" & "','" & "Photo" & "','" & User & "','" & cbocontcode & "','" & txtaartkno & "','" & dtptmd & "'," & thcp & "," & thcpactive & "," & txtthcppremium

oSaccoMaster.ExecuteThis (sql)
'update uncheck tchp member
If txtthcppremium = 0 Then
txtthcppremium = 0
sql = ""
sql = "update d_Suppliers set TYPES=null,tmd=null,status1=0 where sno='" & txtSNo & "'"
oSaccoMaster.ExecuteThis (sql)

End If
'//save to payroll bank details

'//save to the next sno
Dim jk As New ADODB.Recordset
sql = "select receiptno from sno where receiptno='" & txtSNo & "'"
Set jk = oSaccoMaster.GetRecordset(sql)
If jk.EOF Then
sql = ""
sql = "insert into sno(receiptno,auditid) values('" & txtSNo & "','" & User & "')"
oSaccoMaster.ExecuteThis (sql)
End If
Dim ds As Date
ds = Get_Server_Date
ds = Format(ds, "dd/mm/yyyy")
'//check if it exist in the duration tables
'SELECT     sno, dthcps, dthcpd, durations
If txtaartkno = "NA" Then txtaartkno = ""
'From tchp_durations
'SELECT     sno, aarno, startdate, enddate, mpremium, premium, tchpactive, Tmdate, auditid, auditdate
'From tchp_members
'If chkthcp = vbChecked Then
'chkthcp = 1
If txtthcppremium > 0 Then
Set rs = oSaccoMaster.GetRecordset("select sno from tchp_members where sno='" & txtSNo & "'")
If rs.EOF Then
If txtapremium = "" Then txtapremium = 0
sql = ""
sql = "SET DATEFORMAT DMY INSERT INTO tchp_members"
sql = sql & "           (sno, aarno, startdate, enddate, mpremium, premium, tchpactive, Tmdate, auditid, auditdate,statusr)"
sql = sql & "  VALUES     ('" & txtSNo & "','" & txtaartkno & "','" & DTPstart & "','" & DTPend & "'," & txtthcppremium & "," & txtapremium & "," & thcp & ",'" & dtptmd & "','" & User & "','" & Get_Server_Date & "','" & cbostatus & "')"
oSaccoMaster.ExecuteThis (sql)
    If cbostatus <> "Pending" Then
        sql = ""
        sql = "set dateformat dmy INSERT INTO tchp_trxs"
        sql = sql & "     (sno,transdate, description, Debits, CreditsD, CreditsC, Balance, auditid)"
        sql = sql & " VALUES     ('" & txtSNo & "','" & dtptmd & "','Debit'," & txtthcppremium & ",0,0," & txtthcppremium & ",'" & User & "')"
        oSaccoMaster.ExecuteThis (sql)
    End If
sql = ""
sql = "set dateformat dmy INSERT INTO tchp_durations"
sql = sql & "     (sno, dthcps,status)"
sql = sql & " VALUES     ('" & txtSNo & "','" & Format(Get_Server_Date, "dd/mm/yyyy") & "'," & thcp & ")"
oSaccoMaster.ExecuteThis (sql)

Else
'//calculate the durations here
Dim durations As Integer, ldate As Date, cudate As Date
cudate = Format(Get_Server_Date, "dd/mm/yyyy")
'get the last date of signing
If thcp = 1 Then
Set rs = oSaccoMaster.GetRecordset("SELECT TOP 1 dthcpd  FROM  tchp_durations  where sno='" & txtSNo & "' ORDER BY dthcpd DESC, id DESC")
If Not rs.EOF Then
ldate = IIf(IsNull(rs.Fields(0)), Date, rs.Fields(0))
If ldate = "01/01/1900" Then ldate = Date
durations = DateDiff("d", ldate, cudate)
End If

sql = "set dateformat dmy INSERT INTO tchp_durations"
sql = sql & "     (sno, dthcps,status,durations)"
sql = sql & " VALUES     ('" & txtSNo & "','" & Format(Get_Server_Date, "dd/mm/yyyy") & "'," & thcp & "," & durations & ")"
oSaccoMaster.ExecuteThis (sql)
Else
Set rs = oSaccoMaster.GetRecordset("SELECT     TOP 1 dthcps  FROM         tchp_durations  where sno='" & txtSNo & "' ORDER BY dthcps DESC, id DESC")
If Not rs.EOF Then
ldate = rs.Fields(0)
If ldate = "01/01/1900" Then ldate = Date
durations = DateDiff("d", ldate, cudate)
End If
If chkthcp = vbChecked Then
thcp = 1
Else
thcp = 0
End If

sql = "set dateformat dmy INSERT INTO tchp_durations"
sql = sql & "     (sno, dthcpd,status,durations)"
sql = sql & " VALUES     ('" & txtSNo & "','" & Format(Get_Server_Date, "dd/mm/yyyy") & "'," & thcp & "," & durations & ")"
oSaccoMaster.ExecuteThis (sql)
End If
'// update the lastdate on the screen
'SELECT     thcpactive
'From d_Suppliers
'ORDER BY id DESC
'SELECT     status, status2, status3, status4, status5, status6 FROM         tchp_members
'SELECT     status, status2, status3, status4, status5, status6  FROM         d_Suppliers

'update status pending members to blank if unchecked

Dim MyStatus As String
Set rst = Nothing
Dim rsbonus As New Recordset
sql = "select statusr from tchp_members where sno='" & txtSNo & "'"
Set rst = oSaccoMaster.GetRecordset(sql)
    If Not rst.EOF Then
       MyStatus = rst.Fields(0)
        If MyStatus = "Pending" Then
        sql = ""
        sql = "update tchp_members set statusr='' ,tchpactive=" & thcp & ",MPREMIUM=" & CDbl(txtthcppremium) & ",lastdate='" & cudate & "' where sno='" & txtSNo & "'"
        oSaccoMaster.ExecuteThis (sql)
        Else
        sql = ""
        sql = "set dateformat dmy update tchp_members set lastdate='" & cudate & "',tchpactive=" & thcp & ",statusr='" & cbostatus & "',MPREMIUM=" & CDbl(txtthcppremium) & " where sno='" & txtSNo & "'"
        oSaccoMaster.ExecuteThis (sql)
        End If
    End If

sql = ""
sql = "update d_Suppliers set thcpactive=" & thcp & " where sno='" & txtSNo & "'"
oSaccoMaster.ExecuteThis (sql)

End If
End If
'//check if it has the
'End If
sql = ""
sql = ""
'cmdNew_Click
'sql = ""
'sql = "set dateformat dmy update tchp_members set status=0, status2=0, status3=0, status4=0, status5=0, status6=0 where sno='" & txtSNo & "'"
'oSaccoMaster.ExecuteThis (sql)
sql = ""
sql = "update d_Suppliers set status=0, status2=0, status3=0, status4=0, status5=0,Approval = 1, status6=0 where sno='" & txtSNo & "'"
oSaccoMaster.ExecuteThis (sql)
'//update d_payroll for changes to reflect in statement printing
sql = ""
sql = "set dateformat dmy update d_payroll set bank='" & cboBankName & "',AccountNumber='" & txtAccNumber & "',BBranch='" & cboBBranch & "',IdNo='" & txtIdNumber & "' where sno='" & txtSNo & "'"
oSaccoMaster.ExecuteThis (sql)

If chkinoutpatient = vbChecked Then
types = "COMPREHENSIVE"
sql = ""
sql = "update d_Suppliers set TYPES='" & types & "' where sno='" & txtSNo & "'"
oSaccoMaster.ExecuteThis (sql)
End If
If chkinpatient = vbChecked Then
types = "OTHERS"
sql = ""
sql = "update d_Suppliers set TYPES='" & types & "' where sno='" & txtSNo & "'"
oSaccoMaster.ExecuteThis (sql)
End If
If chkoutpatient = vbChecked Then
types = "BASIC"
sql = ""
sql = "update d_Suppliers set TYPES='" & types & "' where sno='" & txtSNo & "'"
oSaccoMaster.ExecuteThis (sql)
End If
''''save the registration amount
Dim rstsr As New Recordset
Dim rsts2 As New Recordset
Dim rsts As New Recordset
Dim tot As Double
Set rsts = oSaccoMaster.GetRecordset("SELECT  isnull(sum(Amount),0) From d_Registration WHERE     (transdescription LIKE '%reg%') AND (SNo = '" & txtSNo & "')")
Set rstsr = oSaccoMaster.GetRecordset("SELECT    Amount From d_Registration WHERE     (transdescription LIKE '%reg%') AND (SNo = '" & txtSNo & "')")
If rstsr.EOF Then
sql = ""
sql = "set dateformat dmy insert into d_Registration (sno, transdate, amount, bal, transdescription, auditid, mno, toledgers, datepostedtoledger, userledger)values('" & txtSNo & "','" & DTPRegDate & "','" & txtrefee & "','" & rsts.Fields(0) & "','registration fee','" & User & "','" & txtSNo & "','1','" & DTPRegDate & "','1')"
oSaccoMaster.ExecuteThis (sql)
Else
sql = ""
sql = "update d_Registration set amount='" & txtrefee & "',bal='" & rsts.Fields(0) & "' where sno='" & txtSNo & "' and transdescription LIKE '%reg%'"
oSaccoMaster.ExecuteThis (sql)
End If
'''''check on gltransaction
Set rsts2 = oSaccoMaster.GetRecordset("SELECT Amount From GLTRANSACTIONS WHERE     (TransDescript LIKE '%reg%') AND (Source = '" & txtSNo & "')")
If rsts2.EOF Then
 sql = ""
 '('" & txtSNo & "','" & DTPRegDate & "','" & txtrefee & "','" & rsts.Fields(0) & "','registration fee','" & User & "','" & txtSNo & "','1','" & DTPRegDate & "','1')
 sql = "set dateformat dmy insert into GLTRANSACTIONS (TransDate, Amount, DrAccNo, CrAccNo, DocumentNo, Source, TransDescript, AuditID, Cash, doc_posted)values('" & DTPRegDate & "','" & txtrefee & "','999999','1','" & txtSNo & "','" & txtSNo & "','registration fee','" & User & "','0','0')"
 oSaccoMaster.ExecuteThis (sql)
Else
 sql = ""
 sql = "update GLTRANSACTIONS set Amount='" & txtrefee & "' where Source='" & txtSNo & "' and TransDescript LIKE '%reg%'"
 oSaccoMaster.ExecuteThis (sql)
End If
'''''end
''''end
'TransDate, Amount, DrAccNo, CrAccNo, DocumentNo, Source, TransDescript, AuditTime, AuditID, Cash, doc_posted


cmdSave.Enabled = False
'//DO THE TYPES

MsgBox "Records successfully updated."
sql = ""
sql = "select * from d_presets where remark = 'Monthly Charges' AND sno='" & txtSNo & "'"
Set rsbonus = oSaccoMaster.GetRecordset(sql)
If rsbonus.EOF Then
 sql = ""
 sql = "set dateformat dmy insert into d_presets (SNo, Deduction, Remark, StartDate, Rate, Stopped,AuditId, Rated) values('" & txtSNo & "','Others','Monthly Charges','" & DTPRegDate & "','50','0','" & User & "','0')"
 oSaccoMaster.ExecuteThis (sql)
End If
'''''''''update run status
'sql = ""
'sql = "update d_Suppliers set Run=0 where sno='" & txtSNo & "'"
'oSaccoMaster.ExecuteThis (sql)
'''''''''end
Form_Load
chkinoutpatient = vbUnchecked
chkinpatient = vbUnchecked
chkoutpatient = vbUnchecked
chkbonus = vbUnchecked
Exit Sub
ErrorHandler:
MsgBox err.description

End Sub

Private Sub cmdtchpp_Click()
frmtchppremiumchange.Show vbModal, Me
End Sub

Private Sub DTPstart_Validate(Cancel As Boolean)
DTPstart = DateSerial(Year(DTPstart), month(DTPstart), 1)
DTPend = DateAdd("m", 18, DTPstart)
DTPend = DateSerial(Year(DTPend), month(DTPend) + 1, 1 - 1)
End Sub



Private Sub dtptmd_Change()
'dtptmd = Format(Get_Server_Date, "dd/mm/yyyy")
dtptmd_Validate 0
End Sub

Private Sub dtptmd_Validate(Cancel As Boolean)
Dim tdate As Date
tdate = Format(Get_Server_Date, "dd/mm/yyyy")
If dtptmd > tdate Then
MsgBox "TCHP Tanykina member Date cannot be a future date", vbCritical

dtptmd = Format(Get_Server_Date, "dd/mm/yyyy")
Exit Sub
End If
End Sub
''******************
Private Sub mini_click()
    cbobrnch.Clear
    Set rs = New Recordset
    'Dim cn As Connection
    Set cn = New ADODB.Connection
    Provider = "MAZIWA"
    'provider = cn
   cn.Open Provider, "atm", "atm"
    Set rs = New Recordset
    sql = "Select distinct(BName) from d_Branch order by BName"
    rs.Open sql, cn, adOpenKeyset, adLockOptimistic
    While Not rs.EOF
    cbobrnch.AddItem rs.Fields(0)
      
    rs.MoveNext
    Wend
    '*************
End Sub
Private Sub Form_Load()
Dim br As String
Dim rsu As New ADODB.Recordset, rsc As New Recordset

DTPRegDate = Format(Get_Server_Date, "dd/mm/yyyy")
newb = 0
txtSNo = ""
txtEMail = ""
txtIdNumber = ""
txtNames = ""
txtPAddress = ""
txtPNumber = ""
txtAccNumber = ""
txtDivision = ""
txtVillage = ""
txtTown = ""
txtrefee = "0"

cboBBranch.Text = ""
cboBankName.Text = ""
cboLocation.Text = ""
'cbobrnch.Text = ""
cboDistrict.Text = ""
cboTCode.Text = ""
cboType.Text = ""
chkthcp = vbUnchecked
chkthcpactive = vbUnchecked
txtthcppremium.Clear
txtaartkno = ""
cmdtchpp.Visible = True
txtSNo.Locked = True
txtEMail.Locked = True
txtIdNumber.Locked = True
txtNames.Locked = True
txtPAddress.Locked = True
txtPNumber.Locked = True
txtAccNumber.Locked = True
txtDivision.Locked = True
txtVillage.Locked = True
txtTown.Locked = True
txtthcppremium.Locked = True
cboBBranch.Locked = True
cboBankName.Locked = True
cboLocation.Locked = True
'cbobrnch.Locked = False
cboDistrict.Locked = True
cboTCode.Locked = True
cboType.Locked = True

cmdNew.Enabled = True
cmdEdit.Enabled = False
cmdSave.Enabled = False
mini_click
 all2
 all
'check the user
sql = "SELECT     UserLoginIDs, UserGroup, SUPERUSER From UserAccounts where UserLoginIDs='" & User & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
If rs!SuperUser <> "1" Then
Frame2.Visible = False
End If
End If


'txtthcppremium

 Set rs = CreateObject("adodb.recordset")
    
'    sql = "SELECT RATE FROM Tchp_Rate"
'    Set rs = oSaccoMaster.GetRecordset(sql)
'
'    If rs.EOF Then
    'MsgBox " Set TCHP Rates First"
    
'    Exit Sub
'    End If
    With rs
        
'        While Not .EOF
'
'         txtthcppremium.AddItem rs.Fields(0)
'
'
'         .MoveNext
'
'        Wend
    
    End With



Dim myclass As cdbase
    'Set cn = New ADODB.Connection
    
    Set cn = CreateObject("adodb.connection")
    
    cn.Open Provider, "atm", "atm"
    
    Set rs = CreateObject("adodb.recordset")
    rs.Open "SELECT BankName,BranchName FROM d_BANKS", cn
    
    If rs.EOF Then Exit Sub
    
    With rs
        
        While Not .EOF
         
         cboBankName.AddItem rs.Fields(0)
         cboBBranch.AddItem rs.Fields(1)
         
         .MoveNext
        
        Wend
    
    End With
    
     Set rs = CreateObject("adodb.recordset")
    
    rs.Open "SELECT LName FROM d_Location", cn
    
    If rs.EOF Then Exit Sub
    
    With rs
        
        While Not .EOF
         If Not IsNull(rs.Fields("LName")) Then
         cboLocation.AddItem rs.Fields("LName")
         End If
         
         .MoveNext
        
        Wend
    
    End With
    
    
     Set rs = CreateObject("adodb.recordset")
    cbobrnch.Clear
    rs.Open "SELECT BName FROM d_Branch", cn

    If rs.EOF Then Exit Sub

    With rs

        While Not .EOF

         cbobrnch.AddItem rs.Fields(0)


         .MoveNext

        Wend

    End With
     Set rsc = CreateObject("adodb.recordset")

    rsc.Open "SELECT ContCode, ContName FROM d_CType", cn

    If rsc.EOF Then Exit Sub

    With rsc

        While Not .EOF

         cbocontcode.AddItem rsc.Fields(1)
        'txtcanname.AddItem rsc.Fields(1)

         .MoveNext

        Wend

    End With
sql = "SELECT     name    FROM         d_company"
Set rsu = oSaccoMaster.GetRecordset(sql)
If Not rsu.EOF Then
br = Trim(rsu.Fields(0))
Else
br = "A"
End If
If br = "A" Then cbobrnch = "SOITARAN FCS"
If br = "B" Then cbobrnch = "AGROVET"
'If br = "C" Then cbobrnch = "KOISOLIK"
'If br = "D" Then cbobrnch = "KORMAET"
'If br = "E" Then cbobrnch = "ITIGO"
'If br = "F" Then cbobrnch = "CHEMUSWO"
'If br = "G" Then cbobrnch = "KAMUNGEI"
'If br = "H" Then cbobrnch = "TALAI"
'If br = "I" Then cbobrnch = "SIRONO"
'If br = "J" Then cbobrnch = "KIMNGERU"
'If br = "K" Then cbobrnch = "KAPSIRIA"
'If BR = "KPK" Then cbobrnch = "KAPKOROS"
'If BR = "KPT" Then cbobrnch = "KAPTEL"
'If BR = "NDP" Then cbobrnch = "NDAPTABWA"

DTPRegDate = Format(Get_Server_Date, "dd/mm/yyyy")

dtptmd = DTPRegDate
If Day(Get_Server_Date) >= 29 And Day(Get_Server_Date) <= 31 Then
cbostatus = "Pending"
End If
If Day(Get_Server_Date) >= 1 And Day(Get_Server_Date) <= 2 Then
cbostatus = "Pending"
End If
DTPstart = DateSerial(Year(DTPstart), month(DTPstart), 1)
DTPend = DateAdd("m", 18, DTPstart)
DTPend = DateSerial(Year(DTPend), month(DTPend) + 1, 1 - 1)
    
    Set rs = CreateObject("adodb.recordset")
    
    rs.Open "SELECT LName FROM d_Location", cn
    
    If rs.EOF Then Exit Sub
    
    With rs
        
        While Not .EOF
         If Not IsNull(rs.Fields("LName")) Then
         cboLocation.AddItem rs.Fields("LName")
         End If
         
         .MoveNext
        
        Wend
    
    End With
    
    Set rs = CreateObject("adodb.recordset")
    cboDistrict.Clear
    rs.Open "SELECT DName FROM d_Districts", cn
    
    If rs.EOF Then Exit Sub
    
    With rs
        
        While Not .EOF
         If Not IsNull(rs.Fields("DName")) Then
         
         cboDistrict.AddItem rs.Fields("DName")
         
         End If
         .MoveNext
        
        Wend
    
    End With
    
    Set rs = CreateObject("adodb.recordset")
    
    rs.Open "SELECT TransCode FROM d_Transporters", cn
    
    If rs.EOF Then Exit Sub
    
    With rs
        
        While Not .EOF
         If Not IsNull(rs.Fields("TransCode")) Then
         cboTCode.AddItem rs.Fields("TransCode")
         End If
         
         .MoveNext
        
        Wend
    End With
'cboType
      Set rs = CreateObject("adodb.recordset")
    
    rs.Open "SELECT BName FROM d_Type", cn
    
    If rs.EOF Then Exit Sub
    cboType.Clear
    
    With rs
        
        While Not .EOF
         
         cboType.AddItem rs.Fields(0)
         
         
         .MoveNext
        
        Wend
    
    End With


End Sub
Private Sub all2()
    cboDistrict.Clear
    Set rs = New Recordset
    'Dim cn As Connection
    Set cn = New ADODB.Connection
    Provider = "MAZIWA"
    'provider = cn
   cn.Open Provider, "atm", "atm"
    Set rs = New Recordset
    sql = "Select distinct(DName) FROM d_Districts order by DName"
    rs.Open sql, cn, adOpenKeyset, adLockOptimistic
    While Not rs.EOF
    cboDistrict.AddItem rs.Fields(0)
      
    rs.MoveNext
    Wend
End Sub
Private Sub Picture5_Click()
        Me.MousePointer = vbHourglass
        frmSearchSupplier.Show vbModal
        txtSNo = sel
        txtSNo_Validate True
        Me.MousePointer = 0
End Sub

Private Sub txtSNo_Change()
Dim MyType As String
Set rst = New ADODB.Recordset
'txtSNo = 1020
sql = " select type from d_suppliers where sno='" & txtSNo & "'"
Set rst = oSaccoMaster.GetRecordset(sql)
If Not rst.EOF Then
MyType = IIf(IsNull(rst.Fields(0)), "", rst.Fields(0))
If MyType = "BASIC" Then
chkoutpatient = vbChecked
ElseIf MyType = "COMPREHENSIVE" Then
chkinoutpatient = vbChecked
ElseIf MyType = "OTHERS" Then
chkinpatient = vbChecked
Else
chkoutpatient = vbUnchecked
chkinoutpatient = vbUnchecked
chkinpatient = vbUnchecked
End If

End If


End Sub

Private Sub txtSNo_Validate(Cancel As Boolean)
If Trim(txtSNo) = "" Then
Exit Sub
End If
txtaartkno = ""
txtapremium = ""
Dim mthcp As Integer, thcpactive As Integer, thcppremium   As Double
'SELECT      RegDate, IdNo, Names, Accno,BCode,BBranch,Type, Village,Location,Division,District,Trader, Active,branch,PhoneNo,address,town, Email,TransCode,[Sign],Photo,SCODE,loan,aarno, tmd, mthcp, thcpactive, thcppremium
'FROM         d_Suppliers where SNo=@SNo
Dim a, t, D As Boolean
txtEMail = ""
txtIdNumber = ""
txtNames = ""
txtPAddress = ""
txtPNumber = ""
txtAccNumber = ""
txtDivision = ""
txtVillage = ""
txtTown = ""
txtrefee = "0"

cboBBranch.Text = ""
cboBankName.Text = ""
cboLocation.Text = ""
'cbobrnch.Text = ""
cboDistrict.Text = ""
cboTCode.Text = ""
cboType.Text = ""
chkthcp = vbUnchecked
chkthcpactive = vbUnchecked
txtthcppremium.Clear
txtaartkno = ""

Dim rsf As New Recordset
Set rsf = New ADODB.Recordset
sql = "d_sp_SelectSuppliers '" & txtSNo & "'"
Set rsf = oSaccoMaster.GetRecordset(sql)
If Not rsf.EOF Then
If Not IsNull(rsf.Fields(0)) Then DTPRegDate = rsf.Fields(0)
If Not IsNull(rsf.Fields(1)) Then txtIdNumber = rsf.Fields(1)
If Not IsNull(rsf.Fields(2)) Then txtNames = rsf.Fields(2)
If Not IsNull(rsf.Fields(3)) Then txtAccNumber = rsf.Fields(3)
If Not IsNull(rsf.Fields(4)) Then cboBankName = rsf.Fields(4)
If Not IsNull(rsf.Fields(5)) Then cboBBranch = rsf.Fields(5)
If Not IsNull(rsf.Fields(6)) Then cboType = rsf.Fields(6)
If Not IsNull(rsf.Fields(7)) Then txtVillage = rsf.Fields(7)
If Not IsNull(rsf.Fields(8)) Then cboLocation = rsf.Fields(8)
If Not IsNull(rsf.Fields(9)) Then txtDivision = rsf.Fields(9)
If Not IsNull(rsf.Fields(10)) Then cboDistrict = rsf.Fields(10)
If Not IsNull(rsf.Fields(11)) Then t = rsf.Fields(11)
If Not IsNull(rsf.Fields(12)) Then a = rsf.Fields(12)
If Not IsNull(rsf.Fields(13)) Then cbobrnch = rsf.Fields(13)
If Not IsNull(rsf.Fields(14)) Then txtPNumber = rsf.Fields(14)
If Not IsNull(rsf.Fields(15)) Then txtPAddress = rsf.Fields(15)
If Not IsNull(rsf.Fields(16)) Then txtTown = rsf.Fields(16)
If Not IsNull(rsf.Fields(17)) Then txtEMail = rsf.Fields(17)
If Not IsNull(rsf.Fields(18)) Then cboTCode = rsf.Fields(18)

'If Not IsNull(rs.Fields(24)) Then dtptmd = rs.Fields(24)
'If Not IsNull(rs.Fields(25)) Then mthcp = rs.Fields(25)
'If Not IsNull(rs.Fields(26)) Then thcpactive = rs.Fields(26)
'If Not IsNull(rs.Fields(27)) Then txtthcppremium = rs.Fields(27)
Dim rstsr As New Recordset
Set rstsr = oSaccoMaster.GetRecordset("SELECT Amount From d_Registration WHERE     (transdescription LIKE '%reg%') AND (SNo = '" & txtSNo & "')")
If Not rstsr.EOF Then
txtrefee = IIf(IsNull(rstsr!amount), 0, rstsr!amount)
End If
''''Set rsk = New ADODB.Recordset
''''sql = "select Stopped,Rate from d_PreSets where Deduction = 'BONUS' AND SNo='" & txtSNo & "'"
''''Set rsk = oSaccoMaster.GetRecordset(sql)
''''If Not rsk.EOF Then
'''' If rsk.Fields(0) = False Then
''''  chkbonus = vbChecked
''''  txtbonus = rsk.Fields(1)
'''' Else
''''  chkbonus = vbUnchecked
''''  txtbonus = "0"
'''' End If
''''End If

If a = True Then
chkActive = vbChecked
Else
chkActive = vbUnchecked
End If

If t = True Then
chkTrader = vbChecked
Else
chkTrader = vbUnchecked
End If
If mthcp = 1 Then
chkthcp = vbChecked
txtthcppremium.Visible = True


Else
chkthcp = vbUnchecked
txtthcppremium.Visible = False
End If
If thcpactive = 1 Then
chkthcpactive = vbChecked
'//GET THE ANNUAL PREMIUM FROM TCHP MEMBER TABLE
'tchp_tchpmember
Set rst = New ADODB.Recordset
sql = "tchp_tchpmember '" & txtSNo & "'"
Set rst = oSaccoMaster.GetRecordset(sql)
If Not rst.EOF Then
txtapremium = rst.Fields(3)
cbostatus = rst.Fields(6)
End If
Else
'sql = "tchp_tchpmember '" & txtSNo & "'"
'Set Rst = oSaccoMaster.GetRecordset(sql)
'
'chkthcpactive = vbUnchecked
End If
cmdEdit.Enabled = True
cmdSave.Enabled = True


cboBankName.Locked = False
txtAccNumber.Locked = False
cboBBranch.Locked = False
lblLoan = "No Loan"

If rsf.Fields(22) = True Then
lblLoan = "Has Loan"
cboBankName.Locked = True
txtAccNumber.Locked = True
cboBBranch.Locked = True
End If
End If

End Sub

Private Sub txtthcppremium_Change()
txtthcppremium.Locked = True
End Sub

Private Sub txtthcppremium_Validate(Cancel As Boolean)
If txtthcppremium = "" Then txtthcppremium = 0
txtapremium = CDbl(txtthcppremium) * 18
txtthcppremium.Locked = True
End Sub
