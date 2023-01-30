VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMilkTests 
   BackColor       =   &H00FF00C0&
   Caption         =   "."
   ClientHeight    =   9075
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8340
   Icon            =   "frmMilkTests.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9075
   ScaleWidth      =   8340
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdShowReport 
      Caption         =   "Show Report"
      Height          =   375
      Left            =   240
      TabIndex        =   53
      Top             =   8640
      Width           =   1095
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FF60C0&
      Caption         =   "Reasons For Rejection ; "
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   43
      Top             =   6240
      Width           =   8055
      Begin MSComctlLib.ListView lvwReasons 
         Height          =   1455
         Left            =   240
         TabIndex        =   47
         Top             =   360
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   2566
         View            =   3
         LabelEdit       =   1
         SortOrder       =   -1  'True
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         MousePointer    =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman Baltic"
            Size            =   8.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Type"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "REASONS"
            Object.Width           =   8890
         EndProperty
      End
   End
   Begin VB.CheckBox chkPrint 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      Caption         =   "Print Receipt to The Farmer"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2760
      TabIndex        =   40
      Top             =   8280
      Width           =   2295
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   6840
      TabIndex        =   28
      Top             =   8640
      Width           =   855
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   375
      Left            =   3000
      TabIndex        =   27
      Top             =   8640
      Width           =   855
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   4680
      TabIndex        =   25
      Top             =   8640
      Width           =   855
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   375
      Left            =   2160
      TabIndex        =   24
      Top             =   8640
      Width           =   855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FF60C0&
      Caption         =   "Tests"
      Height          =   2655
      Left            =   120
      TabIndex        =   2
      Top             =   3600
      Width           =   8055
      Begin VB.TextBox txtPCheck 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5760
         TabIndex        =   52
         Top             =   1680
         Width           =   2175
      End
      Begin VB.OptionButton optNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Caption         =   "No"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1440
         TabIndex        =   51
         Top             =   1800
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton optYes 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Caption         =   "Yes"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2040
         TabIndex        =   50
         Top             =   1800
         Width           =   615
      End
      Begin VB.TextBox txtDram 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5760
         TabIndex        =   45
         Top             =   2160
         Width           =   2175
      End
      Begin VB.TextBox txtPC 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1440
         TabIndex        =   44
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FF60C0&
         Caption         =   "Other Tests"
         Height          =   1095
         Left            =   4200
         TabIndex        =   30
         Top             =   360
         Width           =   3135
         Begin VB.ComboBox cboRez 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "frmMilkTests.frx":030A
            Left            =   1440
            List            =   "frmMilkTests.frx":0323
            TabIndex        =   32
            Top             =   360
            Width           =   855
         End
         Begin VB.ComboBox cboLact 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "frmMilkTests.frx":033C
            Left            =   1440
            List            =   "frmMilkTests.frx":03A6
            TabIndex        =   31
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF60C0&
            Caption         =   "Rezasurin :"
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
            Left            =   120
            TabIndex        =   34
            Top             =   360
            Width           =   1080
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF60C0&
            Caption         =   "Lactometer :"
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
            Left            =   120
            TabIndex        =   33
            Top             =   720
            Width           =   1230
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FF60C0&
         Caption         =   "Organoleptic Tests"
         Height          =   1095
         Left            =   2160
         TabIndex        =   15
         Top             =   360
         Width           =   2055
         Begin VB.OptionButton optBad 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF60C0&
            Caption         =   "Bad"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   19
            Top             =   600
            Width           =   1095
         End
         Begin VB.OptionButton optGood 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF60C0&
            Caption         =   "Good"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   18
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FF60C0&
         Caption         =   "Alcohol Tests"
         Height          =   1095
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   2055
         Begin VB.OptionButton optNegative 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF60C0&
            Caption         =   "Negative"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   480
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.OptionButton optPositive 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF60C0&
            Caption         =   "Positive"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF60C0&
         Caption         =   "Pota Check :"
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
         Left            =   4440
         TabIndex        =   49
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF60C0&
         Caption         =   "Clot on Boil :"
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
         Left            =   120
         TabIndex        =   48
         Top             =   1680
         Width           =   1320
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF60C0&
         Caption         =   "Draminiski :"
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
         Left            =   4440
         TabIndex        =   42
         Top             =   2160
         Width           =   1245
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF60C0&
         Caption         =   "Plate Count :"
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
         Left            =   120
         TabIndex        =   41
         Top             =   2160
         Width           =   1260
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF60C0&
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8055
      Begin VB.PictureBox Picture3 
         Height          =   285
         Left            =   2400
         Picture         =   "frmMilkTests.frx":0498
         ScaleHeight     =   225
         ScaleWidth      =   195
         TabIndex        =   46
         Top             =   720
         Width           =   255
      End
      Begin VB.TextBox txtCCpacity 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2160
         TabIndex        =   38
         Top             =   1800
         Width           =   855
      End
      Begin VB.ComboBox cboTTransporter 
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         ItemData        =   "frmMilkTests.frx":075A
         Left            =   2280
         List            =   "frmMilkTests.frx":0767
         TabIndex        =   37
         Text            =   "Individual Farmer"
         Top             =   3000
         Width           =   2535
      End
      Begin VB.TextBox txtApproxRejected 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2760
         TabIndex        =   29
         Top             =   1200
         Width           =   855
      End
      Begin VB.ComboBox cboTransportMode 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1920
         TabIndex        =   23
         Top             =   2400
         Width           =   2895
      End
      Begin VB.ComboBox cboContainerType 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmMilkTests.frx":0798
         Left            =   5040
         List            =   "frmMilkTests.frx":079A
         TabIndex        =   22
         Top             =   1800
         Width           =   2895
      End
      Begin VB.TextBox txtTotalDelivered 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5880
         TabIndex        =   21
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtSNo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1800
         TabIndex        =   20
         Top             =   720
         Width           =   615
      End
      Begin MSComCtl2.DTPicker DTPTimeIn 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "HH:mm:ss"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   4
         EndProperty
         Height          =   375
         Left            =   6120
         TabIndex        =   12
         Top             =   2400
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   120127490
         CurrentDate     =   40095
      End
      Begin MSComCtl2.DTPicker DTPRejDate 
         Height          =   375
         Left            =   6360
         TabIndex        =   1
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   120127489
         CurrentDate     =   40095
      End
      Begin MSComCtl2.DTPicker DTPTimeOut 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "HH:mm:ss"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   4
         EndProperty
         Height          =   375
         Left            =   6120
         TabIndex        =   13
         Top             =   2880
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   120127490
         CurrentDate     =   40095
      End
      Begin VB.Label lblTransporter 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF00C0&
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2400
         TabIndex        =   39
         Top             =   840
         Width           =   60
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FF60C0&
         Caption         =   "Container Capacity :"
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
         Left            =   120
         TabIndex        =   36
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FF60C0&
         Caption         =   "Type of Transporter :"
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
         Left            =   120
         TabIndex        =   35
         Top             =   3000
         Width           =   2175
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF60C0&
         Caption         =   "Rejection Date :"
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
         Left            =   4680
         TabIndex        =   11
         Top             =   240
         Width           =   1560
      End
      Begin VB.Label lblNames 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF00C0&
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2400
         TabIndex        =   10
         Top             =   600
         Width           =   60
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF60C0&
         Caption         =   "Time Out :"
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
         Left            =   5040
         TabIndex        =   9
         Top             =   2880
         Width           =   1080
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF60C0&
         Caption         =   "Time In :"
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
         Left            =   5040
         TabIndex        =   8
         Top             =   2400
         Width           =   900
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF60C0&
         Caption         =   "Total Kgs Delivered"
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
         Left            =   3840
         TabIndex        =   7
         Top             =   1200
         Width           =   1995
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF60C0&
         Caption         =   "Transport Mode :"
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
         Left            =   120
         TabIndex        =   6
         Top             =   2400
         Width           =   1710
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF60C0&
         Caption         =   "Container Type :"
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
         Left            =   3240
         TabIndex        =   5
         Top             =   1800
         Width           =   1635
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF60C0&
         Caption         =   "Approximate Kgs Rejected"
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
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   2595
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF60C0&
         Caption         =   "Supplier Number"
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
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   3840
      TabIndex        =   26
      Top             =   8640
      Width           =   855
   End
End
Attribute VB_Name = "frmMilkTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Bbit As Integer
Dim En As Integer

Private Sub cboLact_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cboLact_Validate(Cancel As Boolean)
On Error GoTo Shika
En = 1
Dim Criteria As String
Dim dvalue As Integer
Set rs = oSaccoMaster.GetRecordset("d_sp_TestMilk 'Lact'")
Criteria = Trim(rs.Fields("Criteria"))
dvalue = CCur(rs.Fields("dValue"))
If Trim(cboLact) = "" Then
Exit Sub
End If
sql = "SELECT Criteria,Type,Reasons FROM d_M_QSettings WHERE Type LIKE 'Lact%' AND dValue " & Criteria & " " & cboLact & ""
Set rs = oSaccoMaster.GetRecordset(sql)

If Not rs.EOF Then
    While lvwReasons.ListItems.Count > En - 1
    lvwReasons.ListItems.Item(En).selected = True
     If lvwReasons.SelectedItem = "Lactometer" Then
     Exit Sub
     End If
     En = En + 1
     Wend
Else
    While lvwReasons.ListItems.Count > En - 1
    lvwReasons.ListItems.Item(En).selected = True
     If lvwReasons.SelectedItem = "Lactometer" Then
     lvwReasons.ListItems.Remove (lvwReasons.SelectedItem.index)
     End If
     En = En + 1
     Wend
End If

If Not rs.EOF Then
      Set li = lvwReasons.ListItems.Add(, , rs.Fields("Type"))
                        li.SubItems(1) = rs.Fields("Reasons") & ""
    End If

Exit Sub

Shika:
MsgBox err.description
End Sub

Private Sub cboRez_KeyPress(KeyAscii As Integer)
If (KeyAscii = 8) Then
        KeyAscii = KeyAscii
        Exit Sub
End If

If Len(cboRez) < 1 Then
    KeyAscii = KeyAscii
Else
    KeyAscii = 0
        MsgBox "Please enter an integer between 0 and 6", vbInformation, "INVALID ENTRY"
    Exit Sub
End If

If (KeyAscii = 48) Or (KeyAscii = 49) Or (KeyAscii = 50) Or (KeyAscii = 51) Or (KeyAscii = 52) Or (KeyAscii = 53) Or (KeyAscii = 54) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
        MsgBox "Please enter an integer between 0 and 6", vbInformation, "INVALID ENTRY"
    End If
End Sub


Private Sub cboRez_Validate(Cancel As Boolean)
On Error GoTo Shika
En = 1
Dim Criteria, E, g, l
Dim dvalue As Integer
Set rs = oSaccoMaster.GetRecordset("d_sp_TestMilk 'Rez'")
Criteria = Trim(rs.Fields("Criteria"))
dvalue = CCur(rs.Fields("dValue"))
If Trim(cboRez) = "" Then
Exit Sub
End If
sql = "SELECT Criteria,Type,Reasons FROM d_M_QSettings WHERE Type LIKE 'Rez%' AND dValue " & Criteria & " " & cboRez & ""
Set rs = oSaccoMaster.GetRecordset(sql)

If Not rs.EOF Then
    While lvwReasons.ListItems.Count > En - 1
    lvwReasons.ListItems.Item(En).selected = True
     If lvwReasons.SelectedItem = "Rezasurin" Then
     Exit Sub
     End If
     En = En + 1
     Wend
Else
    While lvwReasons.ListItems.Count > En - 1
    lvwReasons.ListItems.Item(En).selected = True
     If lvwReasons.SelectedItem = "Rezasurin" Then
     lvwReasons.ListItems.Remove (lvwReasons.SelectedItem.index)
     End If
     En = En + 1
     Wend
End If

If Not rs.EOF Then
      Set li = lvwReasons.ListItems.Add(, , rs.Fields("Type"))
                        li.SubItems(1) = rs.Fields("Reasons") & ""
    End If

Exit Sub

Shika:
MsgBox err.description

End Sub

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdEdit_Click()
cmdedit.Enabled = False
cmdsave.Enabled = True

Bbit = 0
End Sub

Private Sub cmdNew_Click()
cmdnew.Enabled = False
cmdsave.Enabled = True

Bbit = 1
End Sub

Private Sub cmdsave_Click()
Dim Reasons As String
En = 1
    If txtSNo = "" Then
        MsgBox "Please enter supplier number"
            txtSNo.SetFocus
        Exit Sub
    End If
    
    If txtApproxRejected = "" Then
        MsgBox "Please enter the approximate kgs rejected"
            txtApproxRejected.SetFocus
        Exit Sub
    End If
    
    If txtTotalDelivered = "" Then
        MsgBox "Please enter the total kgs delivered"
            txtTotalDelivered.SetFocus
        Exit Sub
    End If
    
    If cboContainerType = "<Select Container>" Then
        MsgBox "Please select container type."
            cboContainerType.SetFocus
        Exit Sub
    End If
    
    If cboTransportMode = "<Select Transport Mode>" Then
        MsgBox "Please Select Transport Mode."
            cboTransportMode.SetFocus
        Exit Sub
    End If
    
      
    If CCur(txtApproxRejected) > CCur(txtTotalDelivered) Then
        MsgBox "Quantity rejected cannot be greater than the quantity delivered.", vbCritical, "ILLOGICAL"
            txtTotalDelivered.SetFocus
        Exit Sub
    End If
    
   Dim Alcohol As String
   Dim Organo As String
   
   If optBad.value = True Then
    Organo = "BAD"
   Else
    Organo = "GOOD"
   End If
   
   If optNegative.value = True Then
    Alcohol = "NEGATIVE"
   Else
    Alcohol = "POSITIVE"
   End If
   
 Reasons = ""
 
 While lvwReasons.ListItems.Count > En - 1
 lvwReasons.ListItems.Item(En).selected = True
 Reasons = Reasons & lvwReasons.SelectedItem.ListSubItems(1) & vbNewLine
 En = En + 1
 Wend
 
If txtPC = "" Then

txtPC = "0"
End If

If txtPCheck = "" Then
txtPCheck = "0"
End If

If txtDram = "" Then
txtDram = "0"
End If

sql = "d_sp_MilkTests " & txtSNo & ",'" & DTPRejDate & "'," & txtApproxRejected & "," & txtTotalDelivered & ",'" & cboContainerType & "','" & cboTransportMode
sql = sql & "','" & Organo & "'," & cboRez & "," & cboLact & "," & txtPC & "," & txtDram & ",'" & Alcohol & "','" & DTPTimeIn & "','" & DTPTimeOut & "','" & Reasons
sql = sql & "','" & cboTTransporter & "'," & txtCCpacity & ",'" & User & "'," & Bbit & "," & txtPCheck & ""
    
oSaccoMaster.ExecuteThis (sql)
'-- d_sp_MilkTests @Sno bigint,@RejDate varchar(10),@ApproxKgs float,@DeKgs float,@Conttype varchar(50),@TransMode varchar(50),@Organoleptic varchar(20),@Rez int,@Lact float,@PlateCount float,@Draminisk float,@Alcohol varchar(20),@TimeIn varchar(20),@TimeOut varchar(20),@Remarks varchar(155),@Transporter varchar(35),@CCapacity float,@auditid varchar(35),@Bit bit,@PCheck float
If chkPrint = vbChecked Then
    If Trim(Reasons) = "" Then
        MsgBox "Data saved but receipt cannot be printed." & vbNewLine & "There is no reasons for rejection.", vbInformation, "MILK TEST RECEIPT"
            lvwReasons.SetFocus
            Form_Load
        Exit Sub
    End If
    
'/*Print out
 Dim fso, chkPrinter, txtfile
    Dim ttt
     Dim escFeedAndCut As String
     Dim escNewLine As String
     Dim escUnerLineON As String
     Dim escUnerLineOnX2 As String
     Dim escUnerLineOff As String
     Dim escBoldOn As String
     Dim escBoldOff As String
     Dim escNegativeOn As String
     Dim escNegativeOff As String
     Dim esc8CpiOn As String
     Dim esc8CPiOff As String
     Dim esc16Cpi As String
     Dim esc20Cpi As String
     Dim escAlignLeft As String
     Dim escAlignCenter As String
     Dim escAlignRight As String
    
     
        escNewLine = Chr(10) '//New Line (LF Line Feed)
        escUnerLineON = Chr(27) + Chr(45) + Chr(1) '//Unerline on
        escUnerLineOnX2 = Chr(27) + Chr(45) + Chr(1) '//Unerline on X2
        escUnerLineOff = Chr(27) + Chr(45) + Chr(0) '//unerline off
        escBoldOn = Chr(27) + Chr(69) + Chr(1) '//Bold on
        escBoldOff = Chr(27) + Chr(69) + Chr(0) '//Bold off
        escNegativeOn = Chr(29) + Chr(66) + Chr(1) '//White on Black on
        escNegativeOff = Chr(29) + Chr(66) + Chr(0) '//white on
        esc8CpiOn = Chr(29) + Chr(33) + Chr(16) '//Font Size X2 on
        esc8CPiOff = Chr(29) + Chr(33) + Chr(0) '//Font size X2 off
        esc16Cpi = Chr(27) + Chr(77) + Chr(48) '//Font A - Normal Size
        esc20Cpi = Chr(27) + Chr(77) + Chr(49) '//Font B - Small Font
        escAlignLeft = Chr(27) + Chr(97) + Chr(48) '//Align text to the left
        escAlignCenter = Chr(27) + Chr(97) + Chr(49) '//Align text to the center
        escAlignRight = Chr(27) + Chr(97) + Chr(50) '//Align text to the right
        escFeedAndCut = Chr(29) + Chr(86) + Chr(65) '//Partial cut and feed
       'cdgPrint.Filter = "*.csv|*.txt"
        'cdgPrint.ShowSave
        ttt = "LPT1"
        'ttt = "D:\PROJECTS\FOSA\DAILY" & Date & ""
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set chkPrinter = fso.GetFile(ttt)
       
        
        Set txtfile = fso.CreateTextFile(ttt, True)
    txtfile.Write escAlignCenter
    txtfile.WriteLine "" & cname & ""
    txtfile.WriteLine "MILK REJECTION REPORT"
    txtfile.Write escAlignLeft
    txtfile.WriteLine "---------------------------------------"
    txtfile.WriteLine "SNo :" & txtSNo
    txtfile.WriteLine "Name :" & lblNames
    txtfile.WriteLine "Quantity Delivered :" & Format(txtTotalDelivered, "#,##0.00") & " Kgs"
    txtfile.WriteLine "Quantity Rejected :" & Format(txtApproxRejected, "#,##0.00") & " Kgs"
    txtfile.WriteLine "---------------------------------------"
    txtfile.WriteLine "Your milk has been rejected due to;-"
    txtfile.WriteLine Reasons
    txtfile.WriteLine "---------------------------------------"
    txtfile.WriteLine " For Enquiries contact the quality staff"
    txtfile.WriteLine "            " & cname & ""
    txtfile.WriteLine "            QUALITY PERSONELL"
    txtfile.WriteLine "        Date :" & Format(Get_Server_Date, "dd/mm/yyyy HH:MM:SS AM/PM")
    txtfile.WriteLine "         " & motto & ""
    txtfile.WriteLine "---------------------------------------"
    txtfile.WriteLine escFeedAndCut
    
txtfile.Close
End If
Form_Load

MsgBox "Records saved successifully!", vbDefaultButton1, Me.Caption

End Sub

Private Sub cmdShowReport_Click()
reportname = "d_MilkQuality.rpt"
'    ReportTitle = "TO :" & UCase(cboBank) & " ; " & vbNewLine & " Please pay the following farmers the amount indicated: (Our Ref is SNo)"
    '{d_Payroll.NPay} > 0 and {d_Payroll.Bank} <> '' and month({d_Payroll.EndofPeriod})= month(30/09/2010)  AND year({d_Payroll.EndofPeriod}) = Year(30/09/2010)
'    STRFORMULA = "{d_Payroll.NPay} > 0 and {d_Payroll.Bank} = '" & cboBank & "' and month({d_Payroll.EndofPeriod})=" & month(dtpEndPeriod) & " AND year({d_Payroll.EndofPeriod}) =" & Year(dtpEndPeriod)
    Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub Form_Load()
lvwReasons.ListItems.Clear

txtSNo = ""
txtApproxRejected = ""
txtTotalDelivered = ""
txtCCpacity = ""
txtPC = ""
txtDram = ""

cmdnew.Enabled = True
cmdedit.Enabled = False
cmdsave.Enabled = False
cmdDelete.Enabled = False


optNegative.value = True
optGood.value = True


DTPRejDate = Format(Get_Server_Date, "dd/mm/yyyy")
DTPRejDate.MaxDate = Format(Get_Server_Date, "dd/mm/yyyy")

DTPTimeIn.value = "9:00:00"
DTPTimeOut.value = Time()

cboContainerType.Clear
cboTransportMode.Clear

Set rs = oSaccoMaster.GetRecordset("d_sp_ContName")
    While Not rs.EOF
        cboContainerType.AddItem (rs.Fields(0))
        rs.MoveNext
    Wend
    
   ' d_sp_SelTransMode
    Set rs = oSaccoMaster.GetRecordset("d_sp_SelTransMode")
    While Not rs.EOF
        cboTransportMode.AddItem (rs.Fields(0))
        rs.MoveNext
    Wend
    cboContainerType = "<Select Container>"
    cboTransportMode = "<Select Transport Mode>"
End Sub

Private Sub optBad_Validate(Cancel As Boolean)
On Error GoTo Shika
En = 1
Set rst = oSaccoMaster.GetRecordset("d_sp_TestMilk 'Organ'")

    
    While lvwReasons.ListItems.Count > En - 1
    lvwReasons.ListItems.Item(En).selected = True
     If lvwReasons.SelectedItem = "Organoleptic" Then
     Exit Sub
     End If
     En = En + 1
     Wend
     
    If Not rst.EOF Then
      Set li = lvwReasons.ListItems.Add(, , rst.Fields("Type"))
                        li.SubItems(1) = rst.Fields("Reasons") & ""
    End If
Exit Sub

Shika:
MsgBox err.description
End Sub

Private Sub optGood_Validate(Cancel As Boolean)
On Error GoTo Shika
En = 1
    While lvwReasons.ListItems.Count > En - 1
    lvwReasons.ListItems.Item(En).selected = True
     If lvwReasons.SelectedItem = "Organoleptic" Then
      lvwReasons.ListItems.Remove (lvwReasons.SelectedItem.index)
     Exit Sub
     End If
     En = En + 1
     Wend
    Exit Sub
    
Shika:
 MsgBox err.description
End Sub

Private Sub Option1_Click()

End Sub

Private Sub optNegative_Validate(Cancel As Boolean)
 En = 1
    While lvwReasons.ListItems.Count > En - 1
    lvwReasons.ListItems.Item(En).selected = True
     If lvwReasons.SelectedItem = "Alcohol" Then
     lvwReasons.ListItems.Remove (lvwReasons.SelectedItem.index)
     End If
     En = En + 1
     Wend
     
End Sub

Private Sub optNo_Validate(Cancel As Boolean)
     While lvwReasons.ListItems.Count > En - 1
    lvwReasons.ListItems.Item(En).selected = True
     If lvwReasons.SelectedItem = "Clot on Boil" Then
     lvwReasons.ListItems.Remove (lvwReasons.SelectedItem.index)
     End If
     En = En + 1
     Wend

End Sub

Private Sub optPositive_Validate(Cancel As Boolean)
'--d_sp_TestMilk @Type varchar(55) AS
En = 1
Set rs = oSaccoMaster.GetRecordset("d_sp_TestMilk 'ALco'")

    
     While lvwReasons.ListItems.Count > En - 1
     lvwReasons.ListItems.Item(En).selected = True
     If lvwReasons.SelectedItem = "Alcohol" Then
     Exit Sub
     End If
     
     En = En + 1
     Wend
     
    If Not rs.EOF Then
      Set li = lvwReasons.ListItems.Add(, , rs.Fields("Type"))
                        li.SubItems(1) = rs.Fields("Reasons") & ""
    End If

    
End Sub

Private Sub optYes_Validate(Cancel As Boolean)
En = 1
    While lvwReasons.ListItems.Count > En - 1
    lvwReasons.ListItems.Item(En).selected = True
     If lvwReasons.SelectedItem = "Clot on Boil" Then
     Exit Sub
     End If
     En = En + 1
     Wend
     


      Set li = lvwReasons.ListItems.Add(, , "Clot on Boil")
                        li.SubItems(1) = "Not good for processing" & ""
     
     
End Sub

Private Sub Picture3_Click()
        Me.MousePointer = vbHourglass
        frmSearchMilkTests.Show vbModal
        If Not sel = "" Then
        cmdedit.Enabled = True
        cmdnew.Enabled = False
        End If
         
        txtSNo = sel
        DTPRejDate = frmSearchMilkTests.lstSearch.SelectedItem.ListSubItems(1)
        txtApproxRejected = frmSearchMilkTests.lstSearch.SelectedItem.ListSubItems(2)
        txtTotalDelivered = frmSearchMilkTests.lstSearch.SelectedItem.ListSubItems(3)
        'cboTTransporter = frmSearchMilkTests.lstSearch.SelectedItem.ListSubItems(5)
        cboContainerType = frmSearchMilkTests.lstSearch.SelectedItem.ListSubItems(6)
        cboTransportMode = frmSearchMilkTests.lstSearch.SelectedItem.ListSubItems(7)
        
        If UCase(frmSearchMilkTests.lstSearch.SelectedItem.ListSubItems(8)) = "BAD" Then
        optBad = True
        Else
        optGood = True
        End If
        
        cboRez = frmSearchMilkTests.lstSearch.SelectedItem.ListSubItems(9)
        cboLact = frmSearchMilkTests.lstSearch.SelectedItem.ListSubItems(10)
        txtPC = frmSearchMilkTests.lstSearch.SelectedItem.ListSubItems(11)
        
        If UCase(frmSearchMilkTests.lstSearch.SelectedItem.ListSubItems(12)) = "POSITIVE" Then
        optPositive = True
        Else
        optNegative = True
        End If
        
        DTPTimeIn = frmSearchMilkTests.lstSearch.SelectedItem.ListSubItems(13)
        DTPTimeOut = frmSearchMilkTests.lstSearch.SelectedItem.ListSubItems(14)
        txtPCheck = frmSearchMilkTests.lstSearch.SelectedItem.ListSubItems(15)
        txtDram = frmSearchMilkTests.lstSearch.SelectedItem.ListSubItems(16)
        
        txtSNo_Validate True
        Me.MousePointer = 0
End Sub

Private Sub txtApproxRejected_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 46) Or (KeyAscii = 8) Or (KeyAscii = 48) Or (KeyAscii = 49) Or (KeyAscii = 50) Or (KeyAscii = 51) Or (KeyAscii = 52) Or (KeyAscii = 53) Or (KeyAscii = 54) Or (KeyAscii = 55) Or (KeyAscii = 56) Or (KeyAscii = 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
        MsgBox "Please enter a number "
    End If
End Sub

Private Sub txtCCpacity_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 46) Or (KeyAscii = 8) Or (KeyAscii = 48) Or (KeyAscii = 49) Or (KeyAscii = 50) Or (KeyAscii = 51) Or (KeyAscii = 52) Or (KeyAscii = 53) Or (KeyAscii = 54) Or (KeyAscii = 55) Or (KeyAscii = 56) Or (KeyAscii = 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
        MsgBox "Please enter a number "
    End If
End Sub

Private Sub txtDram_KeyPress(KeyAscii As Integer)
If (KeyAscii = 46) Or (KeyAscii = 8) Or (KeyAscii = 48) Or (KeyAscii = 49) Or (KeyAscii = 50) Or (KeyAscii = 51) Or (KeyAscii = 52) Or (KeyAscii = 53) Or (KeyAscii = 54) Or (KeyAscii = 55) Or (KeyAscii = 56) Or (KeyAscii = 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
        MsgBox "Please enter a number "
    End If
End Sub

Private Sub txtPC_KeyPress(KeyAscii As Integer)
If (KeyAscii = 46) Or (KeyAscii = 8) Or (KeyAscii = 48) Or (KeyAscii = 49) Or (KeyAscii = 50) Or (KeyAscii = 51) Or (KeyAscii = 52) Or (KeyAscii = 53) Or (KeyAscii = 54) Or (KeyAscii = 55) Or (KeyAscii = 56) Or (KeyAscii = 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
        MsgBox "Please enter a number "
    End If
End Sub

Private Sub txtPCheck_KeyPress(KeyAscii As Integer)
If (KeyAscii = 46) Or (KeyAscii = 8) Or (KeyAscii = 48) Or (KeyAscii = 49) Or (KeyAscii = 50) Or (KeyAscii = 51) Or (KeyAscii = 52) Or (KeyAscii = 53) Or (KeyAscii = 54) Or (KeyAscii = 55) Or (KeyAscii = 56) Or (KeyAscii = 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
        MsgBox "Please enter a number "
    End If
End Sub

Private Sub txtSNo_KeyPress(KeyAscii As Integer)
If (KeyAscii = 8) Or (KeyAscii = 48) Or (KeyAscii = 49) Or (KeyAscii = 50) Or (KeyAscii = 51) Or (KeyAscii = 52) Or (KeyAscii = 53) Or (KeyAscii = 54) Or (KeyAscii = 55) Or (KeyAscii = 56) Or (KeyAscii = 57) Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
MsgBox "Please enter a number "
End If
End Sub

Private Sub txtSNo_Validate(Cancel As Boolean)

If Trim(txtSNo) = "" Then
    lblNames = ""
        lblTransporter = ""
    Picture3.Visible = True
Exit Sub
End If

Set rs = oSaccoMaster.GetRecordset("Select [Names] FROM d_Suppliers WHERE SNo=" & txtSNo)
    If Not IsNull(rs.Fields(0)) Then
        lblNames = rs.Fields(0)
        Picture3.Visible = False
    Else
        lblNames = ""
        Picture3.Visible = True
    End If
    
Set rs = oSaccoMaster.GetRecordset("d_sp_TransName " & txtSNo & "")
      
    If Not IsNull(rs.Fields("TransName")) Then
        If Not rs.EOF Then
            lblTransporter = "Transporter : " & rs.Fields("TransName")
            'Picture3.Visible = False
        
    Else
        lblTransporter = ""
        'Picture3.Visible = True
        End If
    End If

End Sub

Private Sub txtTotalDelivered_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 46) Or (KeyAscii = 8) Or (KeyAscii = 48) Or (KeyAscii = 49) Or (KeyAscii = 50) Or (KeyAscii = 51) Or (KeyAscii = 52) Or (KeyAscii = 53) Or (KeyAscii = 54) Or (KeyAscii = 55) Or (KeyAscii = 56) Or (KeyAscii = 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
        MsgBox "Please enter a number "
    End If
End Sub
