VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmdebtor1 
   Caption         =   "DEBTORS SALES AND REGISTRY"
   ClientHeight    =   8655
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11160
   LinkTopic       =   "Form1"
   ScaleHeight     =   8655
   ScaleWidth      =   11160
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      Height          =   495
      Left            =   10080
      TabIndex        =   26
      Top             =   8040
      Width           =   855
   End
   Begin MSComCtl2.DTPicker txtdateenterered 
      Height          =   375
      Left            =   9240
      TabIndex        =   20
      Top             =   840
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   129761281
      CurrentDate     =   38814
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   15266
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   32768
      ForeColor       =   192
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "DEBTORS REGISTRATION"
      TabPicture(0)   =   "Form4.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdSave"
      Tab(0).Control(1)=   "cmdEdit"
      Tab(0).Control(2)=   "cmdNew"
      Tab(0).Control(3)=   "Frame1"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "POINT OF SALES "
      TabPicture(1)   =   "Form4.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label7"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label3"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label4"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label5"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label9"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label6"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "txtdcode"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label11"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label12"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label13"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "lbldrstock"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "lblcrvehicle"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Label16"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Label10"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "ListView1"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "cmdnewsearch"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Picture1"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "txtRefNo"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "cboVehicle"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "cboNames"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "txtDispatch"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "txtamountp"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "txtIntake"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "chkpai"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "txtamount"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "cmdnew3"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "cmdsave3"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "cmdstatement"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "txtremarks"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "fra1"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).ControlCount=   32
      Begin VB.Frame fra1 
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   1800
         TabIndex        =   65
         Top             =   3360
         Width           =   4695
         Begin VB.PictureBox Picture2 
            Height          =   255
            Left            =   1320
            Picture         =   "Form4.frx":0038
            ScaleHeight     =   195
            ScaleWidth      =   195
            TabIndex        =   69
            Top             =   240
            Width           =   255
         End
         Begin VB.PictureBox Picture3 
            Height          =   255
            Left            =   1320
            Picture         =   "Form4.frx":0902
            ScaleHeight     =   195
            ScaleWidth      =   195
            TabIndex        =   68
            Top             =   720
            Width           =   255
         End
         Begin VB.TextBox txtdracc 
            Height          =   375
            Left            =   1680
            TabIndex        =   67
            Top             =   120
            Width           =   2895
         End
         Begin VB.TextBox txtcracc 
            Height          =   375
            Left            =   1680
            TabIndex        =   66
            Top             =   600
            Width           =   2895
         End
         Begin VB.Label lbldracc 
            BackColor       =   &H8000000E&
            Height          =   375
            Left            =   120
            TabIndex        =   71
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label lblcracc 
            BackColor       =   &H8000000E&
            Height          =   375
            Left            =   120
            TabIndex        =   70
            Top             =   600
            Width           =   1095
         End
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   375
         Left            =   -71520
         TabIndex        =   59
         Top             =   8160
         Width           =   735
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit"
         Height          =   375
         Left            =   -73080
         TabIndex        =   58
         Top             =   8160
         Width           =   735
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "New"
         Height          =   375
         Left            =   -74760
         TabIndex        =   57
         Top             =   8160
         Width           =   735
      End
      Begin VB.Frame Frame1 
         Caption         =   "Debtors Details"
         BeginProperty Font 
            Name            =   "News701 BT"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7695
         Left            =   -74880
         TabIndex        =   35
         Top             =   360
         Width           =   10935
         Begin VB.TextBox txtVehicle 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5040
            TabIndex        =   74
            Top             =   2160
            Width           =   2895
         End
         Begin VB.CheckBox chkActive 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Active"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8760
            TabIndex        =   73
            Top             =   1440
            Value           =   1  'Checked
            Width           =   975
         End
         Begin VB.CommandButton cmdSearch 
            Caption         =   "..."
            Height          =   285
            Left            =   1560
            TabIndex        =   72
            Top             =   3240
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
            Height          =   315
            Left            =   1860
            TabIndex        =   64
            Top             =   3240
            Width           =   1080
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
            Left            =   3135
            TabIndex        =   63
            Top             =   3240
            Width           =   3225
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
            Left            =   1875
            TabIndex        =   62
            Top             =   3600
            Width           =   1080
         End
         Begin VB.CommandButton Command1 
            Caption         =   "..."
            Height          =   285
            Left            =   1560
            TabIndex        =   61
            Top             =   3600
            Width           =   315
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
            Left            =   3150
            TabIndex        =   60
            Top             =   3600
            Width           =   3225
         End
         Begin VB.TextBox txtEmail 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   960
            TabIndex        =   56
            Top             =   1440
            Width           =   2175
         End
         Begin VB.TextBox txtPrice 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   960
            TabIndex        =   55
            Text            =   "0.00"
            Top             =   2160
            Width           =   1095
         End
         Begin VB.TextBox txtPAddress 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5040
            TabIndex        =   54
            Top             =   1320
            Width           =   2895
         End
         Begin VB.TextBox txtTown 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5040
            TabIndex        =   53
            Top             =   1680
            Width           =   2895
         End
         Begin VB.TextBox txtPhone 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5040
            TabIndex        =   52
            Top             =   840
            Width           =   2895
         End
         Begin VB.TextBox txtId 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   960
            TabIndex        =   51
            Top             =   840
            Width           =   2295
         End
         Begin VB.TextBox txtNames 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4320
            TabIndex        =   49
            Top             =   360
            Width           =   3615
         End
         Begin VB.PictureBox Picture5 
            Height          =   255
            Left            =   3000
            Picture         =   "Form4.frx":11CC
            ScaleHeight     =   195
            ScaleWidth      =   195
            TabIndex        =   48
            Top             =   360
            Width           =   255
         End
         Begin VB.TextBox txtTCode 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1680
            TabIndex        =   47
            Top             =   360
            Width           =   1335
         End
         Begin VB.Frame Frame2 
            Caption         =   "GL Ledgers"
            BeginProperty Font 
               Name            =   "News706 BT"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1575
            Left            =   120
            TabIndex        =   44
            Top             =   2760
            Width           =   10335
            Begin VB.Label Label26 
               Caption         =   "Cr Vehicle"
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
               TabIndex        =   46
               Top             =   840
               Width           =   1215
            End
            Begin VB.Label Label25 
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
               TabIndex        =   45
               Top             =   480
               Width           =   1335
            End
         End
         Begin MSComctlLib.ListView ListView8 
            Height          =   3135
            Left            =   120
            TabIndex        =   76
            Top             =   4560
            Width           =   10695
            _ExtentX        =   18865
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
         Begin VB.Label Label24 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Vehicle No:"
            BeginProperty Font 
               Name            =   "Century"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   3720
            TabIndex        =   75
            Top             =   2160
            Width           =   1185
         End
         Begin VB.Label Label23 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Date:"
            BeginProperty Font 
               Name            =   "Century"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   8280
            TabIndex        =   50
            Top             =   600
            Width           =   555
         End
         Begin VB.Label Label22 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Town:"
            BeginProperty Font 
               Name            =   "Century"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   4080
            TabIndex        =   43
            Top             =   1680
            Width           =   630
         End
         Begin VB.Label Label21 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Price:"
            BeginProperty Font 
               Name            =   "Century"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   240
            TabIndex        =   42
            Top             =   2160
            Width           =   615
         End
         Begin VB.Label Label20 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Email:"
            BeginProperty Font 
               Name            =   "Century"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   240
            TabIndex        =   41
            Top             =   1440
            Width           =   690
         End
         Begin VB.Label Label19 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Phone No:"
            BeginProperty Font 
               Name            =   "Century"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   3480
            TabIndex        =   40
            Top             =   840
            Width           =   1065
         End
         Begin VB.Label Label18 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Name:"
            BeginProperty Font 
               Name            =   "Century"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   3480
            TabIndex        =   39
            Top             =   360
            Width           =   675
         End
         Begin VB.Label Label17 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Postal Address:"
            BeginProperty Font 
               Name            =   "Century"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   3360
            TabIndex        =   38
            Top             =   1320
            Width           =   1605
         End
         Begin VB.Label Label15 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "ID No:"
            BeginProperty Font 
               Name            =   "Century"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   120
            TabIndex        =   37
            Top             =   840
            Width           =   690
         End
         Begin VB.Label Label14 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Debtors Code:"
            BeginProperty Font 
               Name            =   "Century"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   120
            TabIndex        =   36
            Top             =   360
            Width           =   1425
         End
      End
      Begin VB.TextBox txtremarks 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   7920
         TabIndex        =   34
         Top             =   3720
         Width           =   2655
      End
      Begin VB.CommandButton cmdstatement 
         Caption         =   "Debtors Statements"
         Height          =   495
         Left            =   4680
         TabIndex        =   32
         Top             =   8040
         Width           =   1935
      End
      Begin VB.CommandButton cmdsave3 
         Caption         =   "Save"
         Height          =   495
         Left            =   3240
         TabIndex        =   31
         Top             =   8040
         Width           =   975
      End
      Begin VB.CommandButton cmdnew3 
         Caption         =   "New"
         Height          =   495
         Left            =   1800
         TabIndex        =   30
         Top             =   8040
         Width           =   1095
      End
      Begin VB.TextBox txtamount 
         Alignment       =   1  'Right Justify
         Height          =   495
         Left            =   8160
         TabIndex        =   28
         Top             =   3000
         Width           =   1095
      End
      Begin VB.CheckBox chkpai 
         Caption         =   "Make Payments"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   4080
         TabIndex        =   27
         Top             =   2880
         Width           =   2535
      End
      Begin VB.TextBox txtIntake 
         Height          =   375
         Left            =   1800
         TabIndex        =   15
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox txtamountp 
         Height          =   375
         Left            =   5400
         TabIndex        =   11
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox txtDispatch 
         Height          =   375
         Left            =   1800
         TabIndex        =   10
         Top             =   2280
         Width           =   1335
      End
      Begin VB.ComboBox cboNames 
         Height          =   315
         Left            =   1800
         TabIndex        =   7
         Top             =   1680
         Width           =   2055
      End
      Begin VB.ComboBox cboVehicle 
         Height          =   315
         Left            =   1800
         TabIndex        =   6
         Top             =   1200
         Width           =   2055
      End
      Begin VB.TextBox txtRefNo 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   1800
         TabIndex        =   3
         Top             =   600
         Width           =   2055
      End
      Begin VB.PictureBox Picture1 
         Height          =   285
         Left            =   3840
         Picture         =   "Form4.frx":148E
         ScaleHeight     =   225
         ScaleWidth      =   195
         TabIndex        =   2
         Top             =   600
         Width           =   255
      End
      Begin VB.CommandButton cmdnewsearch 
         Caption         =   "New "
         Height          =   285
         Left            =   4080
         TabIndex        =   1
         Top             =   600
         Width           =   615
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3135
         Left            =   240
         TabIndex        =   5
         Top             =   4920
         Width           =   10695
         _ExtentX        =   18865
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
      Begin VB.Label Label10 
         Caption         =   "Remarks."
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6720
         TabIndex        =   33
         Top             =   3720
         Width           =   1455
      End
      Begin VB.Label Label16 
         Caption         =   "Paid Amount:"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6720
         TabIndex        =   29
         Top             =   3120
         Width           =   1335
      End
      Begin VB.Label lblcrvehicle 
         BackColor       =   &H00FFFFC0&
         Height          =   255
         Left            =   6360
         TabIndex        =   25
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label lbldrstock 
         BackColor       =   &H00FFFFC0&
         Height          =   255
         Left            =   6360
         TabIndex        =   24
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label13 
         Caption         =   "Cr Vehicle"
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
         Left            =   5040
         TabIndex        =   23
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label12 
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
         Left            =   5040
         TabIndex        =   22
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Date."
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   8160
         TabIndex        =   21
         Top             =   840
         Width           =   555
      End
      Begin VB.Label txtdcode 
         BackColor       =   &H00FFFFC0&
         Height          =   255
         Left            =   3840
         TabIndex        =   19
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Dr Vehicle"
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
         Left            =   240
         TabIndex        =   17
         Top             =   3480
         Width           =   1335
      End
      Begin VB.Label Label9 
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
         Left            =   240
         TabIndex        =   16
         Top             =   4080
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Intake."
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Amount Payable."
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         TabIndex        =   13
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Dispatch."
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Debtor Name."
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Vehicle No."
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Reference No."
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1425
      End
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      Height          =   495
      Left            =   5520
      TabIndex        =   18
      Top             =   4440
      Width           =   1215
   End
End
Attribute VB_Name = "frmdebtor1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Price As Currency
Dim newa As Integer

Private Sub cboNames_Click()
'NAMES
If cboVehicle = "" Then
    MsgBox "Please select the Vehicle Number."
        cboVehicle.SetFocus
    Exit Sub
End If

Provider = "MAZIWA"
Set cn = New ADODB.Connection
cn.Open Provider, "bi"
Set rst = New ADODB.Recordset
rst.Open sql, cn
'If rs.EOF Then
Set rst = oSaccoMaster.GetRecordset("select DCode,AccDr, AccCr from d_Debtors where DName ='" & cboNames & "'")
If Not rst.EOF Then
txtdcode = rst.Fields("DCode")
lbldrstock = rst.Fields("AccDr")
lblcrvehicle = rst.Fields("AccCr")
End If
Debtorsgl
 'loadoutsale
End Sub

Private Sub cboVehicle_Click()
NAMES
loadBranchesTypes
'cboNames_Click
'SSTab1_DblClick
    'cboVehicle.Clear
End Sub

Private Sub chkpai_Click()
If chkpai = 1 Then
'chkpai.value = 1
''    'cboVehicle.Clear
''    Set rst = New Recordset
''    Dim cn As Connection
''    Set cn = New ADODB.Connection
''    Provider = "MAZIWA"
''    'Provider = cnn
''    cn.Open Provider
''    Set rst = New Recordset
''    sql = "Select  AccCr from  d_Debtors where DCode='" & txtdcode & "'"
''    'Select distinct(Locations) from   d_Debtors
''    rst.Open sql, cn, adOpenKeyset, adLockOptimistic
''    While Not rst.EOF
''    lbldracc.AddItem rst.Fields(0)
''    rst.MoveNext
''    Wend
'
''fra1.Visible = True
''Label13.Visible = True
''Label8.Visible = True
Else
''fra1.Visible = False
''Label13.Visible = False
''Label8.Visible = False
chkpai.value = 0
End If


End Sub

Private Sub cmdedit_Click()
newa = 0
txtVehicle.Locked = False
txtEmail.Locked = False
txtId.Locked = False
txtNames.Locked = False
txtPAddress.Locked = False
txtPhone.Locked = False
'txtsubsidy.Locked = False
txtTCode.Locked = False
txtTown.Locked = False
'cboBBranch.Locked = False
'cboBName.Locked = False
'cbolocation.Locked = False
cmdSave.Enabled = True
End Sub
Private Sub cmdNew_Click()
newa = 1
txtVehicle = ""
txtEmail = ""
txtId = ""
txtNames = ""
txtPAddress = ""
txtPhone = ""
txtCrAccName = ""
txtTCode = ""
txtTown = ""
txtDrAccNo = ""
txtCrAccNo = ""
lblDrAccName = ""
'cbobranch.Text = ""
txtPrice = "0.00"

txtVehicle.Locked = False
txtEmail.Locked = False
txtId.Locked = False
txtNames.Locked = False
txtPAddress.Locked = False
txtPhone.Locked = False
'txtsubsidy.Locked = False
txtTCode.Locked = False
txtTown.Locked = False
'cboBBranch.Locked = False
'cboBName.Locked = False
'cbolocation.Locked = False
cmdEdit.Enabled = False
'cmdSave.Enabled = False
cmdSave.Enabled = True
End Sub

Private Sub cmdnew3_Click()
    txtDispatch.Locked = False
    txtIntake.Locked = True
    txtDispatch = ""
    cboNames = ""
    cboVehicle = ""
    txtamountp = ""
    txtamount = ""
    chkpai.value = vbUnchecked
    cmdnew3.Enabled = False
    cmdsave3.Enabled = True
    cmdnewsearch_Click
End Sub
Public Sub loadReg()
    
    With ListView8
    
        .ListItems.Clear
        
        .ColumnHeaders.Clear
    
    End With

    Set rs2 = CreateObject("adodb.recordset")
    
    sql = "Select DCode,DName,Locations,price,AccDr,AccCr from d_Debtors order by DName"
'    sql = ""
'    sql = "set dateformat dmy SELECT d.RefNo,m.DName, d.DispDate, d.DispQnty,d.Amount,d.PaidAmount FROM d_MilkControl AS d INNER JOIN d_Debtors AS m ON d.DCode = m.DCode WHERE     (DispDate = '" & txtdateenterered & "') and vehicleno='" & cboVehicle & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    
    Set rs2 = CreateObject("adodb.recordset")
    
    Set clsClass = New cdbase
    
    Provider = clsClass.OpenCon
    
    Set cn = CreateObject("adodb.connection")
    
    cn.Open Provider, "bi"
    
    rs2.Open sql, cn
    
    With ListView8
        
        .ColumnHeaders.Add , , "Debtor Code"
        .ColumnHeaders.Add , , "Name"
        .ColumnHeaders.Add , , "Vehicle"
        .ColumnHeaders.Add , , "Price"
        .ColumnHeaders.Add , , "Dr"
        .ColumnHeaders.Add , , "Cr"
        While Not rs2.EOF
        
            Set li = .ListItems.Add(, , Trim(rs2.Fields("DCode")))
            
            li.ListSubItems.Add , , Trim(rs2.Fields("DName"))
            li.ListSubItems.Add , , Trim(rs2.Fields("Locations"))
            li.ListSubItems.Add , , Trim(rs2.Fields("price"))
            li.ListSubItems.Add , , Trim(rs2.Fields("AccDr"))
            li.ListSubItems.Add , , Trim(rs2.Fields("AccCr"))
            rs2.MoveNext
        
        Wend
        
    End With
    
    rs2.Close
    
    Set rs2 = Nothing
    
ListView8.View = lvwReport

End Sub

Private Sub cmdnewsearch_Click()
Dim rsr As New ADODB.Recordset
Dim rsg As New ADODB.Recordset
Dim I As Object
Dim Mylength As Integer
'//if this record is new then look for receipts no

''//clear all textboxes
'mysql = ""
'mysql = "set dateformat dmy select GenerateReceiptno from param"
sql = ""
sql = "set dateformat dmy select GenerateReceiptno from param"
Set rsg = oSaccoMaster.GetRecordset(sql)
If Not rsg.EOF Then
    ''''check check
    If rsg!GenerateReceiptno = True Then
    
        sql = ""
        sql = "select * from Receiptno where receiptno like 'RF-%' order by Receipthnoid desc"
        
        Set rsr = oSaccoMaster.GetRecordset(sql)
        
        If Not rsr.EOF Then
            Mylength = CInt(Mid(rsr!ReceiptNo, 5, 10))
            Mylength = Mylength + 1
            txtRefNo = Padding(Mylength)
            txtRefNo = "RF-" & txtRefNo
        Else
            Mylength = 1
            txtRefNo = "RF-" & Padding(Mylength)
            
        End If
Else
    ''//receiptno  will be keyed in
End If
End If
End Sub

Private Sub cmdsave_Click()
Dim Active As String
On Error GoTo ErrorHandler

If txtTCode = "" Then
MsgBox "Please enter the Debtor code ", vbInformation, "Missing Information"
txtTCode.SetFocus
Exit Sub
End If

If chkActive.value = vbChecked Then
    Active = "1"
Else
    Active = "0"
End If
'sql = ""
'sql = "set dateformat dmy SELECT * From d_Debtors where DCode ='" & txtTCode & "'"
'Set rs = oSaccoMaster.GetRecordset(sql)
'If rs.EOF Then
'      Set cn = New ADODB.Connection
'    sql = ""
'    sql = "SET dateformat DMY Update  d_Debtors SET DNmame= '" & txtNames & "',CertNo='" & txtId & "',Locations='" & cboLocation & "',TregDate='" & DTPRegDate & "',email='" & txtEmail & "',Phoneno='" & txtPhone & "',Town='" & txtTown & "',Address='" & txtPAddress & "',price=" & CCur(txtPrice) & ",Active=" & Active & ",AccDr='" & txtDrAccNo & "',AccCr='" & txtCrAccNo & "' where DCode='" & txtTCode & "'"
'    oSaccoMaster.ExecuteThis (sql)
'
'    Else
'     MsgBox "Debtor code already exist, Use a different code ", vbInformation, "Missing Information"
'   Exit Sub
'  Exit Sub
  
 If newa = 1 Then
    Set cn = New ADODB.Connection
    sql = ""
    sql = "d_sp_Debtors '" & txtTCode & "','" & txtNames & "','" & txtId & "','" & txtVehicle & "','" & txtdateenterered & "','" & txtEmail & "','" & txtPhone & "','" & txtTown & "','" & txtPAddress & "'," & CCur(txtPrice) & "," & CCur(txtSubsidy) & ",'" & txtAccNo & "','" & cboBName & "'," & Active & ",'" & cboBBranch & "','" & cboBranch & "','" & User & "','" & txtDrAccNo & "','" & txtCrAccNo & "','" & txtcessrate & "','" & txtcessdebit & "','" & txtcesscredit & "','" & cessapp & "'"
    oSaccoMaster.ExecuteThis (sql)
   Else
    Set cn = New ADODB.Connection
    sql = ""
    sql = "SET dateformat DMY Update  d_Debtors SET DName= '" & txtNames & "',CertNo='" & txtId & "',Locations='" & txtVehicle & "',TregDate='" & txtdateenterered & "',email='" & txtEmail & "',Phoneno='" & txtPhone & "',Town='" & txtTown & "',Address='" & txtPAddress & "',price=" & CCur(txtPrice) & ",Active=" & Active & ",AccDr='" & txtDrAccNo & "',AccCr='" & txtCrAccNo & "' where DCode='" & txtTCode & "'"
    oSaccoMaster.ExecuteThis (sql)
 End If
cmdNew_Click
cmdSave.Enabled = False
MsgBox "Records successively updated."
loadReg
Exit Sub
ErrorHandler:
MsgBox err.description

End Sub

Private Sub cmdsave3_Click()
  If txtdcode = "" Then
   MsgBox "Debtors code cannot be blank; input an existing one", vbCritical
   cboNames.SetFocus
  Exit Sub
  End If
  
  If txtamount > 0 Then
   If txtremarks = "" Then
     MsgBox "Please enter the Remarks if Cash or Paybill."
     txtremarks.SetFocus
    Exit Sub
   End If
  End If

If chkpai = 0 Then
 If txtdcode = "" Then
  MsgBox "Debtors code cannot be blank; input an existing one", vbCritical
  cboNames.SetFocus
 Exit Sub
 End If
 If txtDispatch = "" Then
    MsgBox "Please enter the dispatch quantity."
        txtDispatch.SetFocus
    Exit Sub
 End If

 If txtIntake = "" Then
    MsgBox "Please enter the intake quantity."
        txtIntake.SetFocus
    Exit Sub
 End If

 If txtRefNo = "" Then
    MsgBox "Please enter the reference number."
        txtRefNo.SetFocus
    Exit Sub
 End If
'//check if the dispatch is greater than the dipping
  If CDbl(txtIntake) < CDbl(txtDispatch) Then 'raiise an alarm
   MsgBox "You cannot take more than what you have in the tank", vbCritical
  Exit Sub
  End If
  Dim Debit As String
  Dim Credit As String

    sql = ""
    sql = "SET dateformat dmy SELECT * FROM  d_MilkControl  WHERE     DispDate = '" & txtdateenterered & "' and DispQnty = '" & txtDispatch & "'and dcode = '" & txtdcode & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
      MsgBox "You have already dispatch for that day", vbInformation
    Exit Sub
    End If
 Dim y As String
 y = cboNames
 Debit = lbldrstock

 Credit = lblcrvehicle

    If Not Save_GLTRANSACTION(Format(txtdateenterered, "dd/mm/yyyy"), (CCur(Price) * CCur(txtDispatch)), Debit, Credit, y, txtRefNo, User, ErrorMessage, "Milk Sales", 1, 1, txtRefNo, transactionNo, "", "", 0) Then
            If ErrorMessage <> "" Then
                MsgBox ErrorMessage, vbInformation, Me.Caption
                ErrorMessage = ""
            End If
    End If
    
'    If capp = 1 Then
'
'    If Not Save_GLTRANSACTION(Format(txtdateenterered, "dd/mm/yyyy"), (CCur(crate) * CCur(txtDispatch)), cessdr, cesscr, "Cess Deductions ", txtRefNo, User, ErrorMessage, "Cess Deductions", 1, 1, txtRefNo, transactionNo, "", "", 0) Then
'            If ErrorMessage <> "" Then
'                MsgBox ErrorMessage, vbInformation, Me.Caption
'                ErrorMessage = ""
'            End If
'    End If
    
'    End If
        
'd_sp_MilkControl @DispDate char(10), @DipsQnty float,@DipQnty float,@InQnty float,@VarQnty float,@Price char(10),@RefNo varchar(35),@CreditAcc varchar(35),@DebitAcc varchar(35),@AuditID varchar (50)
''Set rs = New ADODB.Recordset
''sql = "d_sp_MilkControl  '" & txtdateenterered & "'," & txtDispatch & ",'0'," & txtIntake & ",'0'," & Price & ",'" & txtRefNo & "','" & Credit & "','" & Debit & "','" & User & "','" & txtdcode & "','" & cboVehicle & "','" & txtamountp & "','" & txtamount & "'"
''oSaccoMaster.ExecuteThis (sql)
Else
' If txtdcode = "" Then
'  MsgBox "Debtors code cannot be blank; input an existing one", vbCritical
'  cboNames.SetFocus
' Exit Sub
' End If
'
'  If txtremarks = "" Then
'    MsgBox "Please enter the Remarks if Cash or Paybill."
'        txtremarks.SetFocus
'    Exit Sub
' End If
End If
    
    ''...................insert the amount to debtor if available................................
       Dim Amount1 As Integer
       Set rs = New ADODB.Recordset
       sql = ""
       sql = "SET dateformat dmy Select Amount,PaidAmount  from d_MilkControl  where DCode ='" & txtdcode & "' and DispDate='" & txtdateenterered.value & "'"
       Set rs = oSaccoMaster.GetRecordset(sql)

       If rs.EOF Then
         sql = "d_sp_MilkControl  '" & txtdateenterered & "'," & txtDispatch & ",'0'," & txtIntake & ",'0'," & Price & ",'" & txtRefNo & "','" & Credit & "','" & Debit & "','" & User & "','" & txtdcode & "','" & cboVehicle & "','" & txtamountp & "','" & txtamount & "'"
         oSaccoMaster.ExecuteThis (sql)
'        sql = ""
'        sql = "set dateformat dmy insert into  d_MilkControl(Amount) values('" & CDbl(txtDistributed(Index)) & "') where DCode ='" & txtTCode & "' and DispDate='" & DTPTransdate(Index) & "'"
'        oSaccoMaster.ExecuteThis (sql)
       Else
         sql = ""
         sql = "set dateformat DMY update d_MilkControl set PaidAmount=" & rs.Fields("PaidAmount") + txtamount & " where DCode ='" & txtdcode & "' and DispDate='" & txtdateenterered.value & "' "
         oSaccoMaster.ExecuteThis (sql)
       End If
     'Else
     'End If

    '''..................end of debtor...........................................................
'******************* *********insert to gl
'txtamount = 0
  If txtamount > 0 Then
   If txtremarks = "" Then
     MsgBox "Please enter the Remarks if Cash or Paybill."
     txtremarks.SetFocus
    Exit Sub
   End If
  Dim E As String
   E = txtremarks
    sql = ""
    sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,AuditTime,auditid,cash,doc_posted) values('" & txtdateenterered & "'," & txtamount & ",'" & lbldracc & "','" & lblcracc & "','" & cboNames & "','' ,'" & E & "-MILK PAYMENTS','" & Now & "','" & User & "',0,0)"
    oSaccoMaster.ExecuteThis (sql)
  Else
   'Exit Sub
  End If
'****************************end of gl
'End If

'//subtract from the dispatch table
'    sql = ""
'    sql = "SET      dateformat dmy     SELECT     ID, Intake,transdate     FROM         d_dispatch    WHERE     transdate = '" & txtdateenterered & "'"
'    Set rs = oSaccoMaster.GetRecordset(sql)
'    If rs.EOF Then
'        sql = ""
'        sql = "set dateformat dmy INSERT INTO d_dispatch (Transdate, descrip, Intake, dipping, dispatch, auditid, auditdate)values ('" & txtdateenterered.value & "','Dispatch','0','0'," & CDbl(txtDispatch) & ",'" & User & "','" & Get_Server_Date & "')"
'        oSaccoMaster.ExecuteThis (sql)
'      Else
'        sql = ""
'        sql = "set dateformat dmy UPDATE    d_dispatch  SET   dipping ='0',dispatch=" & txtDispatch & "  WHERE     (Transdate = '" & txtdateenterered & "')"
'        oSaccoMaster.ExecuteThis (sql)
'     End If

         Dim DName As String
          Set rs = New ADODB.Recordset
          sql = "SELECT DName from d_Debtors where DCode='" & txtdcode & "'"
          Set rs = oSaccoMaster.GetRecordset(sql)
          If Not rs.EOF Then
          DName = rs!DName
          End If


'..............END OF  DAILY INTAKE INSERT FOR DEBTORS ONLY.........................
mysql = "set dateformat dmy Insert into Receiptno(Receiptno,Auditdate,auditid)values('" & txtRefNo & "','" & Format(Get_Server_Date, "dd/MM/yyyy") & "','" & User & "')"
oSaccoMaster.ExecuteThis (mysql)
If chkPrint = vbChecked Then
    
If chkPrint = vbChecked Then
    
'/*Print out
 Dim fso, chkPrinter, txtFile
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
        Dim PORT As String
   '     PORT = ports
        'ttt = "LPT1" 'LPT1
        ttt = ports
        'ttt = "D:\PROJECTS\FOSA\DAILY" & Date & ""
        Set fso = CreateObject("Scripting.FileSystemObject")
        'Set chkPrinter = fso.GetFile(ttt)
        
    Set txtFile = fso.CreateTextFile(ttt, True)
    txtFile.WriteLine "         " & cname & ""
    txtFile.WriteLine "         Address :" & paddress & ""
    txtFile.WriteLine "         Phone :" & Phone & ""
    txtFile.WriteLine "         Email :" & Email & ""
    'txtfile.WriteLine " " & txtSNo
    
    txtFile.WriteLine "          Delivery Note"
    txtFile.WriteLine "**********************************************"
        
    Set rs2 = New ADODB.Recordset
    sql = "d_sp_ReceiptNumber"
    Set rs2 = oSaccoMaster.GetRecordset(sql)
    
    Dim RNumber As String
    'RNumber = rs2.Fields(0)
    If Not IsNull(rs2.Fields(0)) Then RNumber = rs2.Fields(0)
    'Else
    'RNumber = "0"
    'End If
    
    txtFile.WriteLine "CsNO :" & txtRefNo
    txtFile.WriteLine "To :" & lblDebtors
   txtFile.WriteLine " *********************************************************************"
    txtFile.WriteLine "DESCRIPTION " & vbTab & "" & vbTab & "value"
    sql = "SELECT     d.DCode, d.DName, SUM(m.DispQnty) AS quantity FROM         d_Debtors AS d INNER JOIN d_MilkControl AS m ON d.DCode = m.dcode WHERE     (DispDate = '" & txtdateenterered & "') GROUP BY d.DCode, d.DName"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then

    txtFile.WriteLine "Milk supplied :" & vbTab & "" & vbTab & " " & rs!Quantity & ""
    txtFile.WriteLine "Amount Payable :" & vbTab & "  " & txtamountp
    txtFile.WriteLine "Receipt Number :" & vbTab & "  " & txtRefNo
    txtFile.WriteLine "Dispatched by :" & vbTab & " " & username & ""
    
    txtFile.WriteLine "---------------------------------------"
    End If
'    txtFile.WriteLine "Receipt Number :" & RNumber
'    txtFile.WriteLine "TRANSPORTER :" & TRANSPORTER
    txtFile.WriteLine "Vehicle No :" & cboVehicle
    txtFile.WriteLine "Received by :" & txtreceiveby
    txtFile.WriteLine "---------------------------------------"
    txtFile.WriteLine "     Date :" & Format(txtdateenterered, "dd/mm/yyyy") & " ,Time : " & Format(Time, "hh:mm:ss AM/PM")
    txtFile.WriteLine "" & motto & ""
    txtFile.WriteLine "---------------------------------------"
    'If chkComment.value = vbChecked Then
        'txtFile.WriteLine txtComment
        txtFile.WriteLine "---------------------------------------"
        txtFile.WriteLine "********POWERED BY EASYMA***************"
    'End If
    txtFile.WriteLine escFeedAndCut
    
 txtFile.Close
 Reset
End If
End If

MsgBox "Records saved successifully."
txtdcode = ""
txtDispatch = ""
txtIntake = ""
txtamount = ""
txtRefNo = ""
txtamountp = ""
txtremarks = ""
txtdracc = ""
txtcracc = ""
lbldracc = ""
lblcracc = ""
    'ListView2.Visible = False
    chkpai.value = vbUnchecked
    'chkPay.value = vbUnchecked
    cmdnew3.Enabled = True
    cmdsave3.Enabled = True
   ' cmdEdit.Enabled = False
    SSTab1_DblClick
    loadBranchesTypes
    cmdnewsearch_Click
    
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

Private Sub cmdstatement_Click()
'milkstatement

   'reportname = "milkstatement.rpt"
    reportname = "d_DebtorsInvoice.rpt"
    Show_Sales_Crystal_Report "", reportname, ""
    'Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub
Public Sub loadBranchesTypes()
    
    With ListView1
    
        .ListItems.Clear
        
        .ColumnHeaders.Clear
    
    End With

    Set rs2 = CreateObject("adodb.recordset")
    
    'sql = "Select RefNo,DispDate, DispQnty, Amount, PaidAmount from d_MilkControl where DispDate='" & txtdateenterered & "'"
    sql = ""
    sql = "set dateformat dmy SELECT d.RefNo,m.DName, d.DispDate, d.DispQnty,d.Amount,d.PaidAmount FROM d_MilkControl AS d INNER JOIN d_Debtors AS m ON d.DCode = m.DCode WHERE     (DispDate = '" & txtdateenterered & "') and vehicleno='" & cboVehicle & "' order by RefNo desc"
    Set rs = oSaccoMaster.GetRecordset(sql)
    
    Set rs2 = CreateObject("adodb.recordset")
    
    Set clsClass = New cdbase
    
    Provider = clsClass.OpenCon
    
    Set cn = CreateObject("adodb.connection")
    
    cn.Open Provider, "bi"
    
    rs2.Open sql, cn
    
    With ListView1
        
        .ColumnHeaders.Add , , "Receipt"
        .ColumnHeaders.Add , , "Name"
        .ColumnHeaders.Add , , "Date"
        .ColumnHeaders.Add , , "Kgs"
        .ColumnHeaders.Add , , "Amount"
        .ColumnHeaders.Add , , "Paid Amount"
        While Not rs2.EOF
        
            Set li = .ListItems.Add(, , Trim(rs2.Fields("RefNo")))
            
            li.ListSubItems.Add , , Trim(rs2.Fields("DName"))
            li.ListSubItems.Add , , Trim(rs2.Fields("DispDate"))
            li.ListSubItems.Add , , Trim(rs2.Fields("DispQnty"))
            li.ListSubItems.Add , , Trim(rs2.Fields("Amount"))
            li.ListSubItems.Add , , Trim(rs2.Fields("PaidAmount"))
            rs2.MoveNext
        
        Wend
        
    End With
    
    rs2.Close
    
    Set rs2 = Nothing
    
ListView1.View = lvwReport

End Sub



Private Sub Command1_Click()
frmSearchGLAccounts.Show vbModal, Me
    If Continue Then
        If SearchValue <> "" Then
            txtCrAccNo = SearchValue
            SearchValue = ""
        End If
    End If
End Sub

Private Sub Form_Load()
txtdateenterered = Format(Get_Server_Date, "dd/mm/yyyy")
txtdateenterered = Format(Get_Server_Date, "dd/mm/yyyy")
SSTab1_DblClick
loadBranchesTypes
NAMES
loadReg
txtdateenterered_Click
End Sub

Private Sub ListView8_dbclick()

End Sub



Private Sub ListView8_DblClick()
txtTCode.Text = ListView8.SelectedItem
txtTCode_Validate True
End Sub

Private Sub Picture1_Click()
Me.MousePointer = vbHourglass
         frmSearchMilkControl.Show vbModal
        txtRefNo = sel
        txtRefNo_Validate True
        Me.MousePointer = 0
End Sub

Private Sub Picture2_Click()
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

Private Sub Picture5_Click()
Me.MousePointer = vbHourglass
         frmSearchDebtors.Show vbModal
        txtTCode = sel
        txtTCode_Validate True
        Me.MousePointer = 0
End Sub

Private Sub txtPrice_Change()
If Trim(txtPrice) = "0.00" Then
txtPrice = ""
End If
End Sub

Private Sub txtTCode_Validate(Cancel As Boolean)
Dim a As Boolean, b As Integer
Set rs = New ADODB.Recordset
sql = "d_sp_Selectdebtors '" & txtTCode & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
If Not IsNull(rs.Fields(0)) Then txtNames = rs.Fields(0)
If Not IsNull(rs.Fields(1)) Then txtId = rs.Fields(1)
'If Not IsNull(rs.Fields(2)) Then cbolocation = rs.Fields(2)
If Not IsNull(rs.Fields(3)) Then txtdateenterered = rs.Fields(3)
If Not IsNull(rs.Fields(4)) Then txtEmail = rs.Fields(4)
If Not IsNull(rs.Fields(5)) Then txtPhone = rs.Fields(5)
If Not IsNull(rs.Fields(6)) Then txtTown = rs.Fields(6)
If Not IsNull(rs.Fields(7)) Then txtPAddress = rs.Fields(7)
'If Not IsNull(rs.Fields(8)) Then txtsubsidy = Format(rs.Fields(8), "#0.00")
If Not IsNull(rs.Fields(9)) Then txtVehicle = rs.Fields(2)
'If Not IsNull(rs.Fields(10)) Then cboBName = rs.Fields(10)
'If Not IsNull(rs.Fields(11)) Then cboBBranch = rs.Fields(11)
If Not IsNull(rs.Fields(12)) Then a = rs.Fields(12)
'If Not IsNull(rs.Fields(13)) Then cboBranch = rs.Fields(13)
If Not IsNull(rs.Fields(14)) Then txtPrice = Format(rs.Fields(14), "#0.00")
If Not IsNull(rs.Fields(15)) Then txtDrAccNo = rs.Fields(15)
If Not IsNull(rs.Fields(16)) Then txtCrAccNo = rs.Fields(16)
'If Not IsNull(rs.Fields(17)) Then txtcessrate = rs.Fields(17)
'If Not IsNull(rs.Fields(18)) Then txtcessdebit = rs.Fields(18)
'If Not IsNull(rs.Fields(19)) Then txtcesscredit = rs.Fields(19)
'If Not IsNull(rs.Fields(20)) Then b = rs.Fields(20)
'If b = 1 Then
'chkcessapp = vbChecked
'Else
'chkcessapp = vbUnchecked
'End If
If a = True Then
chkActive = vbChecked
Else
chkActive = vbUnchecked
End If
cmdEdit.Enabled = True
cmdSave.Enabled = False
End If
End Sub

Private Sub SSTab1_DblClick()
'Private Sub SSTab1_DblClick()
    cboVehicle.Clear
    Set rst = New Recordset
    Dim cn As Connection
    Set cn = New ADODB.Connection
    Provider = "MAZIWA"
    'Provider = cnn
    cn.Open Provider
    Set rst = New Recordset
    sql = "Select distinct(Locations) from   d_Debtors order by Locations"
    'Select distinct(Locations) from   d_Debtors
    rst.Open sql, cn, adOpenKeyset, adLockOptimistic
    While Not rst.EOF
    cboVehicle.AddItem rst.Fields(0)
    rst.MoveNext
    Wend
NAMES
'End Sub
End Sub
Private Sub NAMES()
'Private Sub SSTab1_DblClick()
    cboNames.Clear
    Set rst = New Recordset
    Dim cn As Connection
    Set cn = New ADODB.Connection
    Provider = "MAZIWA"
    'Provider = cnn
    cn.Open Provider
    Set rst = New Recordset
    sql = "Select DName from   d_Debtors where Locations='" & cboVehicle & "' order by DName"
    'Select distinct(Locations) from   d_Debtors
    rst.Open sql, cn, adOpenKeyset, adLockOptimistic
    While Not rst.EOF
    cboNames.AddItem rst.Fields(0)
    rst.MoveNext
    Wend
End Sub

Private Sub Text3_Change()

End Sub

Private Sub txtAmount_Change()
If txtamount = "" Then
 txtamount = 0
 Exit Sub
End If
End Sub

Private Sub txtdateenterered_Click()
    Set rs = New ADODB.Recordset
    sql = "d_sp_DailyTotal '" & txtdateenterered & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not IsNull(rs.Fields(0)) Then
    txtIntake = Format(rs.Fields(0), "#0.00")
    Else
    txtIntake = "0"
    End If
    loadBranchesTypes
    loadReg
End Sub

Private Sub Debtorsgl()
   lblcracc = "A004"
    Set rsd = New Recordset
    Dim cn As Connection
    Set cn = New ADODB.Connection
    Provider = "MAZIWA"
    'Provider = cnn
    cn.Open Provider
    Set rsd = New Recordset
    sql = "Select  AccCr from  d_Debtors where DCode='" & txtdcode & "'"
    'Select distinct(Locations) from   d_Debtors
    rsd.Open sql, cn, adOpenKeyset, adLockOptimistic
    While Not rsd.EOF
    lbldracc = rsd.Fields(0)
    rsd.MoveNext
    Wend
    End Sub

Private Sub txtRefNo_Validate(Cancel As Boolean)
On Error GoTo ErrorHandler
'SELECT TOP 300 DispDate,RefNo,dcode,DispQnty,Price,InQnty,Variance FROM dbo.d_MilkControl"
If Trim(txtRefNo) = "" Then
Exit Sub
End If
 Set rs = oSaccoMaster.GetRecordset("SELECT DispDate,dcode,DispQnty,Price,InQnty,sum(Variance) FROM d_MilkControl WHERE RefNo = '" & txtRefNo & "'")
 
'txtdcode_Validate True
Exit Sub
ErrorHandler:
MsgBox err.description
End Sub
Private Sub txtdcode_Validate(Cancel As Boolean)
Set rs = oSaccoMaster.GetRecordset("SELECT dname,Price,accdr,acccr,drcess,crcess,capp,crate FROM d_Debtors WHERE DCode = '" & txtdcode & "'")
If rs.RecordCount > 0 Then
If Not IsNull(rs.Fields(1)) Then Price = rs.Fields(1)
If Not IsNull(rs.Fields(0)) Then lblDebtors = rs.Fields(0)
If Not IsNull(rs.Fields(2)) Then Label10 = rs.Fields(2)
If Not IsNull(rs.Fields(3)) Then Label11 = rs.Fields(3)
If Not IsNull(rs.Fields(4)) Then cessdr = rs.Fields(4)
If Not IsNull(rs.Fields(5)) Then cesscr = rs.Fields(5)
If Not IsNull(rs.Fields(6)) Then capp = Abs(rs.Fields(6))
If Not IsNull(rs.Fields(7)) Then crate = rs.Fields(7)
txtamountp = txtDispatch * rs.Fields(1)
Else
'lblDebtors = ""
End If
End Sub
Private Sub txtDispatch_Change()
'txtDipping = txtDispatch
If txtIntake = "" Then
  MsgBox "Please Intake the milk."
 Exit Sub
End If

If txtDispatch = "" Then
txtDispatch = "0"
End If

'**************PRICE***************'
Set rs = oSaccoMaster.GetRecordset("SELECT dname,Price,accdr,acccr,drcess,crcess,capp,crate FROM d_Debtors WHERE DCode = '" & txtdcode & "'")
If rs.RecordCount > 0 Then
If Not IsNull(rs.Fields(1)) Then Price = rs.Fields(1)
'If Not IsNull(rs.Fields(0)) Then lblDebtors = rs.Fields(0)
'If Not IsNull(rs.Fields(2)) Then Label10 = rs.Fields(2)
'If Not IsNull(rs.Fields(3)) Then Label11 = rs.Fields(3)
'If Not IsNull(rs.Fields(4)) Then cessdr = rs.Fields(4)
'If Not IsNull(rs.Fields(5)) Then cesscr = rs.Fields(5)
'If Not IsNull(rs.Fields(6)) Then capp = Abs(rs.Fields(6))
'If Not IsNull(rs.Fields(7)) Then crate = rs.Fields(7)
txtamountp = txtDispatch * rs.Fields(1)

Else
'lblDebtors = ""
End If

'****************END********************'

End Sub

Private Sub txtRefNo_Change()
On Error GoTo ErrorHandler
'SELECT TOP 300 DispDate,RefNo,dcode,DispQnty,Price,InQnty,Variance FROM dbo.d_MilkControl"
If Trim(txtRefNo) = "" Then
Exit Sub
End If
' Set rs = oSaccoMaster.GetRecordset("SELECT DispDate,dcode,DispQnty,Price,InQnty,sum(Variance) FROM d_MilkControl WHERE RefNo = '" & txtRefNo & "'")
'
' If rs.RecordCount > 0 Then
'    DTPDispatchDate = rs.Fields(0)
'    txtDispatch = rs.Fields(2)
'    'txtDipping = txtDispatch
'    txtIntake = rs.Fields(4)
'    'txtVariance = rs.Fields(5)
'    Label10 = rs.Fields(1)
'
'    cmdEdit.Enabled = True
'Else
'    cmdEdit.Enabled = False
'
'End If
'txtdcode_Validate True
Exit Sub
ErrorHandler:
MsgBox err.description
End Sub
