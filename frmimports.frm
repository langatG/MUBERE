VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmimports 
   Caption         =   "IMPORT OTHER BRANCH DETAILS"
   ClientHeight    =   8505
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11520
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmimports.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   11520
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab EDMS 
      Height          =   6735
      Left            =   120
      TabIndex        =   11
      Top             =   1680
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   11880
      _Version        =   393216
      Tabs            =   10
      Tab             =   3
      TabsPerRow      =   10
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "MILK INTAKE"
      TabPicture(0)   =   "frmimports.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1(0)"
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(5)=   "Label6"
      Tab(0).Control(6)=   "Label7"
      Tab(0).Control(7)=   "Label8"
      Tab(0).Control(8)=   "Label9"
      Tab(0).Control(9)=   "Label10"
      Tab(0).Control(10)=   "Label11"
      Tab(0).Control(11)=   "Label14"
      Tab(0).Control(12)=   "txtsno"
      Tab(0).Control(13)=   "txttransdate"
      Tab(0).Control(14)=   "txtqsupplied"
      Tab(0).Control(15)=   "txtppu"
      Tab(0).Control(16)=   "txtpamount"
      Tab(0).Control(17)=   "txttranstime"
      Tab(0).Control(18)=   "txtauditid"
      Tab(0).Control(19)=   "txtauditdatetime"
      Tab(0).Control(20)=   "txtpaid"
      Tab(0).Control(21)=   "txtLR"
      Tab(0).Control(22)=   "txtremark"
      Tab(0).Control(23)=   "Frame1"
      Tab(0).Control(24)=   "txtbr"
      Tab(0).ControlCount=   25
      TabCaption(1)   =   "SUPPLIERS"
      TabPicture(1)   =   "frmimports.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1(4)"
      Tab(1).Control(1)=   "Label36"
      Tab(1).Control(2)=   "Label37"
      Tab(1).Control(3)=   "Label38"
      Tab(1).Control(4)=   "Label39"
      Tab(1).Control(5)=   "Label40"
      Tab(1).Control(6)=   "Label41"
      Tab(1).Control(7)=   "Label42"
      Tab(1).Control(8)=   "Label43"
      Tab(1).Control(9)=   "Label44"
      Tab(1).Control(10)=   "Label45"
      Tab(1).Control(11)=   "Label46"
      Tab(1).Control(12)=   "Label1(5)"
      Tab(1).Control(13)=   "Label47"
      Tab(1).Control(14)=   "Label48"
      Tab(1).Control(15)=   "Label49"
      Tab(1).Control(16)=   "Label50"
      Tab(1).Control(17)=   "Label51"
      Tab(1).Control(18)=   "Label52"
      Tab(1).Control(19)=   "Label53"
      Tab(1).Control(20)=   "Label54"
      Tab(1).Control(21)=   "Label55"
      Tab(1).Control(22)=   "Label56"
      Tab(1).Control(23)=   "Label57"
      Tab(1).Control(24)=   "Label1(6)"
      Tab(1).Control(25)=   "Label58"
      Tab(1).Control(26)=   "Label59"
      Tab(1).Control(27)=   "Label60"
      Tab(1).Control(28)=   "Label61"
      Tab(1).Control(29)=   "Label62"
      Tab(1).Control(30)=   "txtdivision"
      Tab(1).Control(31)=   "txtdistrict"
      Tab(1).Control(32)=   "txtlocations"
      Tab(1).Control(33)=   "txtvillage"
      Tab(1).Control(34)=   "txttype"
      Tab(1).Control(35)=   "txtbbranch"
      Tab(1).Control(36)=   "txtbcode"
      Tab(1).Control(37)=   "txtaccno"
      Tab(1).Control(38)=   "txtnames"
      Tab(1).Control(39)=   "txtidno"
      Tab(1).Control(40)=   "txtregdate"
      Tab(1).Control(41)=   "txtsno1"
      Tab(1).Control(42)=   "txtscode"
      Tab(1).Control(43)=   "txtloan"
      Tab(1).Control(44)=   "txtauditdatetime1"
      Tab(1).Control(45)=   "txtauditid1"
      Tab(1).Control(46)=   "txttranscode"
      Tab(1).Control(47)=   "txtemail"
      Tab(1).Control(48)=   "txttown"
      Tab(1).Control(49)=   "txtaddress"
      Tab(1).Control(50)=   "txtphone"
      Tab(1).Control(51)=   "txtbranch"
      Tab(1).Control(52)=   "txtactive"
      Tab(1).Control(53)=   "txttrader"
      Tab(1).Control(54)=   "TXTBR1"
      Tab(1).Control(55)=   "txthast"
      Tab(1).Control(56)=   "txtrate"
      Tab(1).Control(57)=   "txtfrate"
      Tab(1).Control(58)=   "txtisfrare"
      Tab(1).Control(59)=   "txtcompare"
      Tab(1).Control(60)=   "Frame2"
      Tab(1).ControlCount=   61
      TabCaption(2)   =   "Suppliers Deductions"
      TabPicture(2)   =   "frmimports.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label12"
      Tab(2).Control(1)=   "Label13"
      Tab(2).Control(2)=   "Label15"
      Tab(2).Control(3)=   "Label16"
      Tab(2).Control(4)=   "Label17"
      Tab(2).Control(5)=   "Label18"
      Tab(2).Control(6)=   "Label19"
      Tab(2).Control(7)=   "Label21"
      Tab(2).Control(8)=   "Label22"
      Tab(2).Control(9)=   "Label23"
      Tab(2).Control(10)=   "Label24"
      Tab(2).Control(11)=   "Label1(2)"
      Tab(2).Control(12)=   "Frame3"
      Tab(2).Control(13)=   "txtremarks"
      Tab(2).Control(14)=   "txtyear"
      Tab(2).Control(15)=   "txtauditdate3"
      Tab(2).Control(16)=   "txtauditid3"
      Tab(2).Control(17)=   "txtenddate"
      Tab(2).Control(18)=   "txtstartdate"
      Tab(2).Control(19)=   "txtperiod"
      Tab(2).Control(20)=   "txtamout"
      Tab(2).Control(21)=   "txtdescription"
      Tab(2).Control(22)=   "txtdatededuc"
      Tab(2).Control(23)=   "txtsno3"
      Tab(2).Control(24)=   "txtbr3"
      Tab(2).ControlCount=   25
      TabCaption(3)   =   "Transport Deductions"
      TabPicture(3)   =   "frmimports.frx":035E
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Label1(3)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label25"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Label26"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Label27"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Label28"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "Label29"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "Label30"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "Label31"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "Label32"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "Label33"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "Label34"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).Control(11)=   "Label35"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).Control(12)=   "txtbr4"
      Tab(3).Control(12).Enabled=   0   'False
      Tab(3).Control(13)=   "txtsno4"
      Tab(3).Control(13).Enabled=   0   'False
      Tab(3).Control(14)=   "txtdatededuc2"
      Tab(3).Control(14).Enabled=   0   'False
      Tab(3).Control(15)=   "txtdescription1"
      Tab(3).Control(15).Enabled=   0   'False
      Tab(3).Control(16)=   "txtamount1"
      Tab(3).Control(16).Enabled=   0   'False
      Tab(3).Control(17)=   "txtperiod1"
      Tab(3).Control(17).Enabled=   0   'False
      Tab(3).Control(18)=   "txtstartdate1"
      Tab(3).Control(18).Enabled=   0   'False
      Tab(3).Control(19)=   "txtenddate1"
      Tab(3).Control(19).Enabled=   0   'False
      Tab(3).Control(20)=   "txtauditid4"
      Tab(3).Control(20).Enabled=   0   'False
      Tab(3).Control(21)=   "txtauditdate4"
      Tab(3).Control(21).Enabled=   0   'False
      Tab(3).Control(22)=   "txtyear1"
      Tab(3).Control(22).Enabled=   0   'False
      Tab(3).Control(23)=   "txttrate"
      Tab(3).Control(23).Enabled=   0   'False
      Tab(3).Control(24)=   "Frame4"
      Tab(3).Control(24).Enabled=   0   'False
      Tab(3).ControlCount=   25
      TabCaption(4)   =   "TRANSPORT DETAILS"
      TabPicture(4)   =   "frmimports.frx":037A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label63"
      Tab(4).Control(1)=   "Label64"
      Tab(4).Control(2)=   "Label65"
      Tab(4).Control(3)=   "Label66"
      Tab(4).Control(4)=   "Label67"
      Tab(4).Control(5)=   "Label68"
      Tab(4).Control(6)=   "Label69"
      Tab(4).Control(7)=   "Label70"
      Tab(4).Control(8)=   "Label71"
      Tab(4).Control(9)=   "Label72"
      Tab(4).Control(10)=   "Label73"
      Tab(4).Control(11)=   "Label76"
      Tab(4).Control(12)=   "Label79"
      Tab(4).Control(13)=   "Label80"
      Tab(4).Control(14)=   "Label81"
      Tab(4).Control(15)=   "Label82"
      Tab(4).Control(16)=   "Label83"
      Tab(4).Control(17)=   "Label84"
      Tab(4).Control(18)=   "Label1(7)"
      Tab(4).Control(19)=   "Label74"
      Tab(4).Control(20)=   "txtisfrate5"
      Tab(4).Control(21)=   "txtrate5"
      Tab(4).Control(22)=   "txtbr5"
      Tab(4).Control(23)=   "txtactive5"
      Tab(4).Control(24)=   "txtsubsidy5"
      Tab(4).Control(25)=   "txtphoneno5"
      Tab(4).Control(26)=   "txtaddress5"
      Tab(4).Control(27)=   "txttown5"
      Tab(4).Control(28)=   "txtemail5"
      Tab(4).Control(29)=   "txtauditid5"
      Tab(4).Control(30)=   "txtauditdatetime5"
      Tab(4).Control(31)=   "txttranscode5"
      Tab(4).Control(32)=   "txtregdate5"
      Tab(4).Control(33)=   "txtcertno5"
      Tab(4).Control(34)=   "txtname5"
      Tab(4).Control(35)=   "txtaccno5"
      Tab(4).Control(36)=   "txtbcode5"
      Tab(4).Control(37)=   "txtbbranch5"
      Tab(4).Control(38)=   "txtlocation5"
      Tab(4).Control(39)=   "txttbranch5"
      Tab(4).Control(40)=   "Frame5"
      Tab(4).ControlCount=   41
      TabCaption(5)   =   "EDMS DATA"
      TabPicture(5)   =   "frmimports.frx":0396
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "cmdupdate"
      Tab(5).Control(1)=   "transdate"
      Tab(5).ControlCount=   2
      TabCaption(6)   =   "TRANSPORT ASSIGNMENT"
      TabPicture(6)   =   "frmimports.frx":03B2
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Label77"
      Tab(6).Control(1)=   "Isfrate"
      Tab(6).Control(2)=   "Label86"
      Tab(6).Control(3)=   "Label87"
      Tab(6).Control(4)=   "Label88"
      Tab(6).Control(5)=   "Label89"
      Tab(6).Control(6)=   "Label90"
      Tab(6).Control(7)=   "Label91"
      Tab(6).Control(8)=   "Label92"
      Tab(6).Control(9)=   "Label1(8)"
      Tab(6).Control(10)=   "txtbr7"
      Tab(6).Control(11)=   "Frame6"
      Tab(6).Control(12)=   "txtisfrate7"
      Tab(6).Control(13)=   "txtauditdatetime7"
      Tab(6).Control(14)=   "txtauditid7"
      Tab(6).Control(15)=   "txtdateinactive7"
      Tab(6).Control(16)=   "txtacitve7"
      Tab(6).Control(17)=   "txtstartdate7"
      Tab(6).Control(18)=   "txtrate7"
      Tab(6).Control(19)=   "txtsno7"
      Tab(6).Control(20)=   "txttranscode7"
      Tab(6).ControlCount=   21
      TabCaption(7)   =   "Tab 7"
      TabPicture(7)   =   "frmimports.frx":03CE
      Tab(7).ControlEnabled=   0   'False
      Tab(7).ControlCount=   0
      TabCaption(8)   =   "Tab 8"
      Tab(8).ControlEnabled=   0   'False
      Tab(8).ControlCount=   0
      TabCaption(9)   =   "Tab 9"
      Tab(9).ControlEnabled=   0   'False
      Tab(9).ControlCount=   0
      Begin VB.TextBox txttranscode7 
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
         Left            =   -73185
         TabIndex        =   219
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox txtsno7 
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
         Left            =   -73185
         TabIndex        =   218
         Top             =   1140
         Width           =   2055
      End
      Begin VB.TextBox txtrate7 
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
         Left            =   -73185
         TabIndex        =   217
         Top             =   1425
         Width           =   2055
      End
      Begin VB.TextBox txtstartdate7 
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
         Left            =   -73200
         TabIndex        =   216
         Top             =   1725
         Width           =   2055
      End
      Begin VB.TextBox txtacitve7 
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
         Left            =   -73185
         TabIndex        =   215
         Top             =   2010
         Width           =   2055
      End
      Begin VB.TextBox txtdateinactive7 
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
         Left            =   -73185
         TabIndex        =   214
         Top             =   2310
         Width           =   2055
      End
      Begin VB.TextBox txtauditid7 
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
         Left            =   -73185
         TabIndex        =   213
         Top             =   2595
         Width           =   2055
      End
      Begin VB.TextBox txtauditdatetime7 
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
         Left            =   -73185
         TabIndex        =   212
         Top             =   2895
         Width           =   2055
      End
      Begin VB.TextBox txtisfrate7 
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
         Left            =   -73185
         TabIndex        =   211
         Top             =   3180
         Width           =   2055
      End
      Begin VB.Frame Frame6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   -70440
         TabIndex        =   207
         Top             =   600
         Width           =   1695
         Begin VB.CommandButton Command9 
            BackColor       =   &H80000001&
            Caption         =   "&Post"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   165
            TabIndex        =   210
            Top             =   645
            Width           =   1410
         End
         Begin VB.CommandButton Command8 
            BackColor       =   &H80000001&
            Caption         =   "&Close"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   165
            TabIndex        =   209
            Top             =   1080
            Width           =   1410
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H80000001&
            Caption         =   "Import "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   165
            TabIndex        =   208
            Top             =   225
            Width           =   1410
         End
      End
      Begin VB.TextBox txtbr7 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -73200
         TabIndex        =   206
         Top             =   3480
         Width           =   2055
      End
      Begin MSComCtl2.DTPicker transdate 
         Height          =   375
         Left            =   -72360
         TabIndex        =   205
         Top             =   840
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         Format          =   120193025
         CurrentDate     =   40823
      End
      Begin VB.CommandButton cmdupdate 
         Caption         =   "Update Easyma"
         Height          =   375
         Left            =   -74160
         TabIndex        =   204
         Top             =   840
         Width           =   1455
      End
      Begin VB.Frame Frame5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   -67200
         TabIndex        =   200
         Top             =   840
         Width           =   1695
         Begin VB.CommandButton cmdposttransportersdetails 
            BackColor       =   &H80000001&
            Caption         =   "&Post"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   165
            TabIndex        =   203
            Top             =   645
            Width           =   1410
         End
         Begin VB.CommandButton Command7 
            BackColor       =   &H80000001&
            Caption         =   "&Close"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   165
            TabIndex        =   202
            Top             =   1080
            Width           =   1410
         End
         Begin VB.CommandButton cmdimporttransportdetails 
            BackColor       =   &H80000001&
            Caption         =   "Import "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   165
            TabIndex        =   201
            Top             =   225
            Width           =   1410
         End
      End
      Begin VB.TextBox txttbranch5 
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
         Left            =   -69555
         TabIndex        =   198
         Top             =   1080
         Width           =   2055
      End
      Begin VB.TextBox txtlocation5 
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
         Left            =   -73275
         TabIndex        =   178
         Top             =   2280
         Width           =   2055
      End
      Begin VB.TextBox txtbbranch5 
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
         Left            =   -73275
         TabIndex        =   177
         Top             =   5595
         Width           =   2055
      End
      Begin VB.TextBox txtbcode5 
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
         Left            =   -73275
         TabIndex        =   176
         Top             =   5310
         Width           =   2055
      End
      Begin VB.TextBox txtaccno5 
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
         Left            =   -73275
         TabIndex        =   175
         Top             =   5010
         Width           =   2055
      End
      Begin VB.TextBox txtname5 
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
         Left            =   -73290
         TabIndex        =   174
         Top             =   1920
         Width           =   2055
      End
      Begin VB.TextBox txtcertno5 
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
         Left            =   -73275
         TabIndex        =   173
         Top             =   1665
         Width           =   2055
      End
      Begin VB.TextBox txtregdate5 
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
         Left            =   -73275
         TabIndex        =   172
         Top             =   1380
         Width           =   2055
      End
      Begin VB.TextBox txttranscode5 
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
         Left            =   -73275
         TabIndex        =   171
         Top             =   1080
         Width           =   2055
      End
      Begin VB.TextBox txtauditdatetime5 
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
         Left            =   -69570
         TabIndex        =   170
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox txtauditid5 
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
         Left            =   -69555
         TabIndex        =   169
         Top             =   1380
         Width           =   2055
      End
      Begin VB.TextBox txtemail5 
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
         Left            =   -73320
         TabIndex        =   168
         Top             =   3075
         Width           =   2055
      End
      Begin VB.TextBox txttown5 
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
         Left            =   -73275
         TabIndex        =   167
         Top             =   3870
         Width           =   2055
      End
      Begin VB.TextBox txtaddress5 
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
         Left            =   -73275
         TabIndex        =   166
         Top             =   4290
         Width           =   2055
      End
      Begin VB.TextBox txtphoneno5 
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
         Left            =   -73290
         TabIndex        =   165
         Top             =   3405
         Width           =   2055
      End
      Begin VB.TextBox txtsubsidy5 
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
         Left            =   -73275
         TabIndex        =   164
         Top             =   4665
         Width           =   2055
      End
      Begin VB.TextBox txtactive5 
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
         Left            =   -73275
         TabIndex        =   163
         Top             =   2700
         Width           =   2055
      End
      Begin VB.TextBox txtbr5 
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
         Left            =   -69555
         TabIndex        =   162
         Top             =   3270
         Width           =   2055
      End
      Begin VB.TextBox txtrate5 
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
         Left            =   -69570
         TabIndex        =   161
         Top             =   2685
         Width           =   2055
      End
      Begin VB.TextBox txtisfrate5 
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
         Left            =   -69555
         TabIndex        =   160
         Top             =   2100
         Width           =   2055
      End
      Begin VB.Frame Frame4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   6360
         TabIndex        =   144
         Top             =   840
         Width           =   1695
         Begin VB.CommandButton cmdimporttransdeduc 
            BackColor       =   &H80000001&
            Caption         =   "Import "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   165
            TabIndex        =   147
            Top             =   225
            Width           =   1410
         End
         Begin VB.CommandButton Command6 
            BackColor       =   &H80000001&
            Caption         =   "&Close"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   165
            TabIndex        =   146
            Top             =   1080
            Width           =   1410
         End
         Begin VB.CommandButton cmdposttransdeductions 
            BackColor       =   &H80000001&
            Caption         =   "&Post"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   165
            TabIndex        =   145
            Top             =   645
            Width           =   1410
         End
      End
      Begin VB.TextBox txttrate 
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
         Left            =   2055
         TabIndex        =   143
         Top             =   3885
         Width           =   2055
      End
      Begin VB.TextBox txtyear1 
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
         Left            =   2055
         TabIndex        =   142
         Top             =   3600
         Width           =   2055
      End
      Begin VB.TextBox txtauditdate4 
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
         Left            =   2040
         TabIndex        =   141
         Top             =   3300
         Width           =   2055
      End
      Begin VB.TextBox txtauditid4 
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
         Left            =   2055
         TabIndex        =   140
         Top             =   3015
         Width           =   2055
      End
      Begin VB.TextBox txtenddate1 
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
         Left            =   2055
         TabIndex        =   139
         Top             =   2715
         Width           =   2055
      End
      Begin VB.TextBox txtstartdate1 
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
         Left            =   2055
         TabIndex        =   138
         Top             =   2430
         Width           =   2055
      End
      Begin VB.TextBox txtperiod1 
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
         Left            =   2055
         TabIndex        =   137
         Top             =   2130
         Width           =   2055
      End
      Begin VB.TextBox txtamount1 
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
         Left            =   2040
         TabIndex        =   136
         Top             =   1845
         Width           =   2055
      End
      Begin VB.TextBox txtdescription1 
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
         Left            =   2055
         TabIndex        =   135
         Top             =   1545
         Width           =   3015
      End
      Begin VB.TextBox txtdatededuc2 
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
         Left            =   2040
         TabIndex        =   134
         Top             =   1260
         Width           =   2055
      End
      Begin VB.TextBox txtsno4 
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
         Left            =   2055
         TabIndex        =   133
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox txtbr4 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2040
         TabIndex        =   132
         Top             =   4200
         Width           =   2055
      End
      Begin VB.TextBox txtbr 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -73320
         TabIndex        =   131
         Top             =   3960
         Width           =   2055
      End
      Begin VB.TextBox txtbr3 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -73320
         TabIndex        =   130
         Top             =   3960
         Width           =   2055
      End
      Begin VB.TextBox txtsno3 
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
         Left            =   -73305
         TabIndex        =   117
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox txtdatededuc 
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
         Left            =   -73320
         TabIndex        =   116
         Top             =   1020
         Width           =   2055
      End
      Begin VB.TextBox txtdescription 
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
         Left            =   -73305
         TabIndex        =   115
         Top             =   1305
         Width           =   3015
      End
      Begin VB.TextBox txtamout 
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
         Left            =   -73320
         TabIndex        =   114
         Top             =   1605
         Width           =   2055
      End
      Begin VB.TextBox txtperiod 
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
         Left            =   -73305
         TabIndex        =   113
         Top             =   1890
         Width           =   2055
      End
      Begin VB.TextBox txtstartdate 
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
         Left            =   -73305
         TabIndex        =   112
         Top             =   2190
         Width           =   2055
      End
      Begin VB.TextBox txtenddate 
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
         Left            =   -73305
         TabIndex        =   111
         Top             =   2475
         Width           =   2055
      End
      Begin VB.TextBox txtauditid3 
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
         Left            =   -73305
         TabIndex        =   110
         Top             =   2775
         Width           =   2055
      End
      Begin VB.TextBox txtauditdate3 
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
         Left            =   -73305
         TabIndex        =   109
         Top             =   3060
         Width           =   2055
      End
      Begin VB.TextBox txtyear 
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
         Left            =   -73305
         TabIndex        =   108
         Top             =   3360
         Width           =   2055
      End
      Begin VB.TextBox txtremarks 
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
         Left            =   -73305
         TabIndex        =   107
         Top             =   3645
         Width           =   2055
      End
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   -69000
         TabIndex        =   103
         Top             =   600
         Width           =   1695
         Begin VB.CommandButton Command5 
            BackColor       =   &H80000001&
            Caption         =   "&Post"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   165
            TabIndex        =   106
            Top             =   645
            Width           =   1410
         End
         Begin VB.CommandButton Command4 
            BackColor       =   &H80000001&
            Caption         =   "&Close"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   165
            TabIndex        =   105
            Top             =   1080
            Width           =   1410
         End
         Begin VB.CommandButton cmdimportsupplierdeduct 
            BackColor       =   &H80000001&
            Caption         =   "Import "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   165
            TabIndex        =   104
            Top             =   225
            Width           =   1410
         End
      End
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   -65880
         TabIndex        =   99
         Top             =   4200
         Width           =   1695
         Begin VB.CommandButton cmdpostsuppliers 
            BackColor       =   &H80000001&
            Caption         =   "&Post"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   165
            TabIndex        =   102
            Top             =   645
            Width           =   1410
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H80000001&
            Caption         =   "&Close"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   165
            TabIndex        =   101
            Top             =   1080
            Width           =   1410
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H80000001&
            Caption         =   "Import "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   165
            TabIndex        =   100
            Top             =   225
            Width           =   1410
         End
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   -70560
         TabIndex        =   95
         Top             =   480
         Width           =   1695
         Begin VB.CommandButton cmdimport 
            BackColor       =   &H80000001&
            Caption         =   "Import "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   165
            TabIndex        =   98
            Top             =   225
            Width           =   1410
         End
         Begin VB.CommandButton cmdclose 
            BackColor       =   &H80000001&
            Caption         =   "&Close"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   165
            TabIndex        =   97
            Top             =   1080
            Width           =   1410
         End
         Begin VB.CommandButton cmdpost 
            BackColor       =   &H80000001&
            Caption         =   "&Post"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   165
            TabIndex        =   96
            Top             =   645
            Width           =   1410
         End
      End
      Begin VB.TextBox txtcompare 
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
         Left            =   -65865
         TabIndex        =   88
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox txtisfrare 
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
         Left            =   -65865
         TabIndex        =   87
         Top             =   900
         Width           =   2055
      End
      Begin VB.TextBox txtfrate 
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
         Left            =   -65865
         TabIndex        =   86
         Top             =   1185
         Width           =   2055
      End
      Begin VB.TextBox txtrate 
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
         Left            =   -65880
         TabIndex        =   85
         Top             =   1485
         Width           =   2055
      End
      Begin VB.TextBox txthast 
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
         Left            =   -65865
         TabIndex        =   84
         Top             =   1770
         Width           =   2055
      End
      Begin VB.TextBox TXTBR1 
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
         Left            =   -65865
         TabIndex        =   83
         Top             =   2070
         Width           =   2055
      End
      Begin VB.TextBox txttrader 
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
         Left            =   -69600
         TabIndex        =   70
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox txtactive 
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
         Left            =   -69585
         TabIndex        =   69
         Top             =   900
         Width           =   2055
      End
      Begin VB.TextBox txtbranch 
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
         Left            =   -69585
         TabIndex        =   68
         Top             =   1185
         Width           =   2055
      End
      Begin VB.TextBox txtphone 
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
         Left            =   -69600
         TabIndex        =   67
         Top             =   1485
         Width           =   2055
      End
      Begin VB.TextBox txtaddress 
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
         Left            =   -69585
         TabIndex        =   66
         Top             =   1770
         Width           =   2055
      End
      Begin VB.TextBox txttown 
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
         Left            =   -69585
         TabIndex        =   65
         Top             =   2070
         Width           =   2055
      End
      Begin VB.TextBox txtemail 
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
         Left            =   -69585
         TabIndex        =   64
         Top             =   2355
         Width           =   2055
      End
      Begin VB.TextBox txttranscode 
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
         Left            =   -69585
         TabIndex        =   63
         Top             =   2655
         Width           =   2055
      End
      Begin VB.TextBox txtauditid1 
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
         Left            =   -69585
         TabIndex        =   62
         Top             =   2940
         Width           =   2055
      End
      Begin VB.TextBox txtauditdatetime1 
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
         Left            =   -69600
         TabIndex        =   61
         Top             =   3240
         Width           =   2055
      End
      Begin VB.TextBox txtloan 
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
         Left            =   -69585
         TabIndex        =   60
         Top             =   3825
         Width           =   2055
      End
      Begin VB.TextBox txtscode 
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
         Left            =   -69585
         TabIndex        =   59
         Top             =   3525
         Width           =   2055
      End
      Begin VB.TextBox txtsno1 
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
         Left            =   -73305
         TabIndex        =   46
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox txtregdate 
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
         Left            =   -73305
         TabIndex        =   45
         Top             =   900
         Width           =   2055
      End
      Begin VB.TextBox txtidno 
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
         Left            =   -73305
         TabIndex        =   44
         Top             =   1185
         Width           =   2055
      End
      Begin VB.TextBox txtnames 
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
         Left            =   -73320
         TabIndex        =   43
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox txtaccno 
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
         Left            =   -73305
         TabIndex        =   42
         Top             =   1770
         Width           =   2055
      End
      Begin VB.TextBox txtbcode 
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
         Left            =   -73305
         TabIndex        =   41
         Top             =   2070
         Width           =   2055
      End
      Begin VB.TextBox txtbbranch 
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
         Left            =   -73305
         TabIndex        =   40
         Top             =   2355
         Width           =   2055
      End
      Begin VB.TextBox txttype 
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
         Left            =   -73305
         TabIndex        =   39
         Top             =   2655
         Width           =   2055
      End
      Begin VB.TextBox txtvillage 
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
         Left            =   -73305
         TabIndex        =   38
         Top             =   2940
         Width           =   2055
      End
      Begin VB.TextBox txtlocations 
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
         Left            =   -73305
         TabIndex        =   37
         Top             =   3240
         Width           =   2055
      End
      Begin VB.TextBox txtdistrict 
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
         Left            =   -73305
         TabIndex        =   36
         Top             =   3825
         Width           =   2055
      End
      Begin VB.TextBox txtdivision 
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
         Left            =   -73305
         TabIndex        =   35
         Top             =   3525
         Width           =   2055
      End
      Begin VB.TextBox txtremark 
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
         Left            =   -73305
         TabIndex        =   22
         Top             =   3645
         Width           =   2055
      End
      Begin VB.TextBox txtLR 
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
         Left            =   -73305
         TabIndex        =   21
         Top             =   3360
         Width           =   2055
      End
      Begin VB.TextBox txtpaid 
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
         Left            =   -73305
         TabIndex        =   20
         Top             =   3060
         Width           =   2055
      End
      Begin VB.TextBox txtauditdatetime 
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
         Left            =   -73305
         TabIndex        =   19
         Top             =   2775
         Width           =   2055
      End
      Begin VB.TextBox txtauditid 
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
         Left            =   -73305
         TabIndex        =   18
         Top             =   2475
         Width           =   2055
      End
      Begin VB.TextBox txttranstime 
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
         Left            =   -73305
         TabIndex        =   17
         Top             =   2190
         Width           =   2055
      End
      Begin VB.TextBox txtpamount 
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
         Left            =   -73305
         TabIndex        =   16
         Top             =   1890
         Width           =   2055
      End
      Begin VB.TextBox txtppu 
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
         Left            =   -73320
         TabIndex        =   15
         Top             =   1605
         Width           =   2055
      End
      Begin VB.TextBox txtqsupplied 
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
         Left            =   -73305
         TabIndex        =   14
         Top             =   1305
         Width           =   2055
      End
      Begin VB.TextBox txttransdate 
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
         Left            =   -73305
         TabIndex        =   13
         Top             =   1020
         Width           =   2055
      End
      Begin VB.TextBox txtsno 
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
         Left            =   -73305
         TabIndex        =   12
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Trans Code"
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
         Index           =   8
         Left            =   -74295
         TabIndex        =   229
         Top             =   840
         Width           =   960
      End
      Begin VB.Label Label92 
         AutoSize        =   -1  'True
         Caption         =   "S No."
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
         Left            =   -73695
         TabIndex        =   228
         Top             =   1140
         Width           =   405
      End
      Begin VB.Label Label91 
         AutoSize        =   -1  'True
         Caption         =   "Rate"
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
         Left            =   -73695
         TabIndex        =   227
         Top             =   1425
         Width           =   360
      End
      Begin VB.Label Label90 
         AutoSize        =   -1  'True
         Caption         =   "Start Date"
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
         Left            =   -74055
         TabIndex        =   226
         Top             =   1725
         Width           =   795
      End
      Begin VB.Label Label89 
         AutoSize        =   -1  'True
         Caption         =   "Active"
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
         Left            =   -73815
         TabIndex        =   225
         Top             =   2010
         Width           =   510
      End
      Begin VB.Label Label88 
         AutoSize        =   -1  'True
         Caption         =   "Date Inactive"
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
         Left            =   -74295
         TabIndex        =   224
         Top             =   2310
         Width           =   1035
      End
      Begin VB.Label Label87 
         AutoSize        =   -1  'True
         Caption         =   "Audit Id"
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
         Left            =   -73920
         TabIndex        =   223
         Top             =   2595
         Width           =   630
      End
      Begin VB.Label Label86 
         AutoSize        =   -1  'True
         Caption         =   "Audit Date Time"
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
         Left            =   -74535
         TabIndex        =   222
         Top             =   2880
         Width           =   1305
      End
      Begin VB.Label Isfrate 
         AutoSize        =   -1  'True
         Caption         =   "Isfrate"
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
         Left            =   -73815
         TabIndex        =   221
         Top             =   3180
         Width           =   540
      End
      Begin VB.Label Label77 
         AutoSize        =   -1  'True
         Caption         =   "BR"
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
         Left            =   -73455
         TabIndex        =   220
         Top             =   3465
         Width           =   210
      End
      Begin VB.Label Label74 
         AutoSize        =   -1  'True
         Caption         =   "T Branch"
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
         Left            =   -70320
         TabIndex        =   199
         Top             =   1080
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Trans Code"
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
         Index           =   7
         Left            =   -74265
         TabIndex        =   197
         Top             =   1080
         Width           =   960
      End
      Begin VB.Label Label84 
         AutoSize        =   -1  'True
         Caption         =   "Reg Date"
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
         Left            =   -74025
         TabIndex        =   196
         Top             =   1380
         Width           =   720
      End
      Begin VB.Label Label83 
         AutoSize        =   -1  'True
         Caption         =   "Cert No."
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
         Left            =   -74025
         TabIndex        =   195
         Top             =   1665
         Width           =   660
      End
      Begin VB.Label Label82 
         AutoSize        =   -1  'True
         Caption         =   "Names"
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
         Left            =   -74025
         TabIndex        =   194
         Top             =   1965
         Width           =   570
      End
      Begin VB.Label Label81 
         AutoSize        =   -1  'True
         Caption         =   "Accno"
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
         Left            =   -73905
         TabIndex        =   193
         Top             =   5040
         Width           =   510
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "B Code"
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
         Left            =   -73905
         TabIndex        =   192
         Top             =   5310
         Width           =   585
      End
      Begin VB.Label Label79 
         AutoSize        =   -1  'True
         Caption         =   "BBranch"
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
         Left            =   -74025
         TabIndex        =   191
         Top             =   5595
         Width           =   675
      End
      Begin VB.Label Label76 
         AutoSize        =   -1  'True
         Caption         =   "Locations"
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
         Left            =   -74265
         TabIndex        =   190
         Top             =   2280
         Width           =   810
      End
      Begin VB.Label Label73 
         AutoSize        =   -1  'True
         Caption         =   "Active"
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
         Left            =   -73860
         TabIndex        =   189
         Top             =   2700
         Width           =   510
      End
      Begin VB.Label Label72 
         AutoSize        =   -1  'True
         Caption         =   "Subsidy"
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
         Left            =   -73980
         TabIndex        =   188
         Top             =   4665
         Width           =   660
      End
      Begin VB.Label Label71 
         AutoSize        =   -1  'True
         Caption         =   "Phone"
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
         Left            =   -73890
         TabIndex        =   187
         Top             =   3480
         Width           =   525
      End
      Begin VB.Label Label70 
         AutoSize        =   -1  'True
         Caption         =   "Address"
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
         Left            =   -74070
         TabIndex        =   186
         Top             =   4290
         Width           =   720
      End
      Begin VB.Label Label69 
         AutoSize        =   -1  'True
         Caption         =   "Town"
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
         Left            =   -73920
         TabIndex        =   185
         Top             =   3870
         Width           =   465
      End
      Begin VB.Label Label68 
         AutoSize        =   -1  'True
         Caption         =   "Email"
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
         Left            =   -73845
         TabIndex        =   184
         Top             =   3075
         Width           =   435
      End
      Begin VB.Label Label67 
         AutoSize        =   -1  'True
         Caption         =   "Audit Id"
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
         Left            =   -70320
         TabIndex        =   183
         Top             =   1380
         Width           =   630
      End
      Begin VB.Label Label66 
         AutoSize        =   -1  'True
         Caption         =   "Audit Date Time"
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
         Left            =   -70920
         TabIndex        =   182
         Top             =   1680
         Width           =   1305
      End
      Begin VB.Label Label65 
         AutoSize        =   -1  'True
         Caption         =   "isfrate"
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
         Left            =   -70200
         TabIndex        =   181
         Top             =   2100
         Width           =   540
      End
      Begin VB.Label Label64 
         AutoSize        =   -1  'True
         Caption         =   "Rate"
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
         Left            =   -70080
         TabIndex        =   180
         Top             =   2685
         Width           =   360
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "BR"
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
         Left            =   -69840
         TabIndex        =   179
         Top             =   3270
         Width           =   210
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   "Rate"
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
         Left            =   1560
         TabIndex        =   159
         Top             =   3885
         Width           =   360
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   "BR"
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
         Left            =   1680
         TabIndex        =   158
         Top             =   4185
         Width           =   210
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "yyear"
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
         Left            =   1560
         TabIndex        =   157
         Top             =   3600
         Width           =   450
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "Audit Date Time"
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
         Left            =   600
         TabIndex        =   156
         Top             =   3300
         Width           =   1305
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "Audit ID"
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
         Left            =   1320
         TabIndex        =   155
         Top             =   3015
         Width           =   630
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "End Date"
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
         Left            =   1200
         TabIndex        =   154
         Top             =   2715
         Width           =   705
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "StartDate"
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
         Left            =   1200
         TabIndex        =   153
         Top             =   2430
         Width           =   750
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "period"
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
         Left            =   1320
         TabIndex        =   152
         Top             =   2130
         Width           =   540
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Amount"
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
         Left            =   1200
         TabIndex        =   151
         Top             =   1845
         Width           =   660
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Description"
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
         Left            =   1080
         TabIndex        =   150
         Top             =   1545
         Width           =   945
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Date Deduc"
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
         Left            =   1080
         TabIndex        =   149
         Top             =   1260
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Trans Code"
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
         Index           =   3
         Left            =   1080
         TabIndex        =   148
         Top             =   960
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Supplier Number"
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
         Index           =   2
         Left            =   -74760
         TabIndex        =   129
         Top             =   720
         Width           =   1395
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Date Deduc"
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
         Left            =   -74280
         TabIndex        =   128
         Top             =   1020
         Width           =   915
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Description"
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
         Left            =   -74280
         TabIndex        =   127
         Top             =   1305
         Width           =   945
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Amount"
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
         Left            =   -74040
         TabIndex        =   126
         Top             =   1605
         Width           =   660
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "period"
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
         Left            =   -73920
         TabIndex        =   125
         Top             =   1890
         Width           =   540
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "StartDate"
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
         Left            =   -74160
         TabIndex        =   124
         Top             =   2190
         Width           =   750
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "End Date"
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
         Left            =   -74040
         TabIndex        =   123
         Top             =   2475
         Width           =   705
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Audit ID"
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
         Left            =   -74040
         TabIndex        =   122
         Top             =   2775
         Width           =   630
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Audit Date Time"
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
         Left            =   -74640
         TabIndex        =   121
         Top             =   3060
         Width           =   1305
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "yyear"
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
         Left            =   -73800
         TabIndex        =   120
         Top             =   3360
         Width           =   450
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "BR"
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
         Left            =   -73680
         TabIndex        =   119
         Top             =   3945
         Width           =   210
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Remarks"
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
         Left            =   -74160
         TabIndex        =   118
         Top             =   3645
         Width           =   750
      End
      Begin VB.Label Label62 
         AutoSize        =   -1  'True
         Caption         =   "BR"
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
         Left            =   -66150
         TabIndex        =   94
         Top             =   2040
         Width           =   210
      End
      Begin VB.Label Label61 
         AutoSize        =   -1  'True
         Caption         =   "hast"
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
         Left            =   -66300
         TabIndex        =   93
         Top             =   1770
         Width           =   360
      End
      Begin VB.Label Label60 
         AutoSize        =   -1  'True
         Caption         =   "Rate"
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
         Left            =   -66360
         TabIndex        =   92
         Top             =   1485
         Width           =   360
      End
      Begin VB.Label Label59 
         AutoSize        =   -1  'True
         Caption         =   "frate"
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
         Left            =   -66330
         TabIndex        =   91
         Top             =   1185
         Width           =   390
      End
      Begin VB.Label Label58 
         AutoSize        =   -1  'True
         Caption         =   "isfrate"
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
         Left            =   -66450
         TabIndex        =   90
         Top             =   900
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Compare"
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
         Index           =   6
         Left            =   -66720
         TabIndex        =   89
         Top             =   600
         Width           =   765
      End
      Begin VB.Label Label57 
         AutoSize        =   -1  'True
         Caption         =   "Scode"
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
         Left            =   -70290
         TabIndex        =   82
         Top             =   3525
         Width           =   510
      End
      Begin VB.Label Label56 
         AutoSize        =   -1  'True
         Caption         =   "Loan"
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
         Left            =   -70215
         TabIndex        =   81
         Top             =   3825
         Width           =   405
      End
      Begin VB.Label Label55 
         AutoSize        =   -1  'True
         Caption         =   "Audit Date Time"
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
         Left            =   -70935
         TabIndex        =   80
         Top             =   3240
         Width           =   1305
      End
      Begin VB.Label Label54 
         AutoSize        =   -1  'True
         Caption         =   "Audit Id"
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
         Left            =   -70350
         TabIndex        =   79
         Top             =   2940
         Width           =   630
      End
      Begin VB.Label Label53 
         AutoSize        =   -1  'True
         Caption         =   "Trans Code"
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
         Left            =   -70590
         TabIndex        =   78
         Top             =   2640
         Width           =   960
      End
      Begin VB.Label Label52 
         AutoSize        =   -1  'True
         Caption         =   "Email"
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
         Left            =   -70155
         TabIndex        =   77
         Top             =   2355
         Width           =   435
      End
      Begin VB.Label Label51 
         AutoSize        =   -1  'True
         Caption         =   "Town"
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
         Left            =   -70230
         TabIndex        =   76
         Top             =   2070
         Width           =   465
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         Caption         =   "Address"
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
         Left            =   -70380
         TabIndex        =   75
         Top             =   1770
         Width           =   720
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         Caption         =   "Phone"
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
         Left            =   -70320
         TabIndex        =   74
         Top             =   1485
         Width           =   525
      End
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         Caption         =   "Branch"
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
         Left            =   -70290
         TabIndex        =   73
         Top             =   1185
         Width           =   570
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         Caption         =   "Active"
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
         Left            =   -70290
         TabIndex        =   72
         Top             =   900
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Trader"
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
         Index           =   5
         Left            =   -70320
         TabIndex        =   71
         Top             =   600
         Width           =   555
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         Caption         =   "Division"
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
         Left            =   -74130
         TabIndex        =   58
         Top             =   3525
         Width           =   645
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         Caption         =   "District"
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
         Left            =   -74175
         TabIndex        =   57
         Top             =   3825
         Width           =   585
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         Caption         =   "Locations"
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
         Left            =   -74175
         TabIndex        =   56
         Top             =   3240
         Width           =   810
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         Caption         =   "Village"
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
         Left            =   -74070
         TabIndex        =   55
         Top             =   2940
         Width           =   555
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   "Type"
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
         Left            =   -73950
         TabIndex        =   54
         Top             =   2640
         Width           =   405
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "BBranch"
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
         Left            =   -74115
         TabIndex        =   53
         Top             =   2355
         Width           =   675
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         Caption         =   "B Code"
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
         Left            =   -74190
         TabIndex        =   52
         Top             =   2070
         Width           =   585
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         Caption         =   "Accno"
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
         Left            =   -74100
         TabIndex        =   51
         Top             =   1770
         Width           =   510
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         Caption         =   "Names"
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
         Left            =   -74160
         TabIndex        =   50
         Top             =   1485
         Width           =   570
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         Caption         =   "ID No."
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
         Left            =   -74010
         TabIndex        =   49
         Top             =   1185
         Width           =   450
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         Caption         =   "Reg Date"
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
         Left            =   -74370
         TabIndex        =   48
         Top             =   900
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Supplier Number"
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
         Index           =   4
         Left            =   -74880
         TabIndex        =   47
         Top             =   600
         Width           =   1395
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Remark"
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
         Left            =   -74130
         TabIndex        =   34
         Top             =   3645
         Width           =   645
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "BR"
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
         Left            =   -73695
         TabIndex        =   33
         Top             =   3945
         Width           =   210
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "LR"
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
         Left            =   -73695
         TabIndex        =   32
         Top             =   3360
         Width           =   210
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Paid"
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
         Left            =   -73830
         TabIndex        =   31
         Top             =   3060
         Width           =   345
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Audit Date Time"
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
         Left            =   -74790
         TabIndex        =   30
         Top             =   2775
         Width           =   1305
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Audit Id"
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
         Left            =   -74115
         TabIndex        =   29
         Top             =   2475
         Width           =   630
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Trans Time"
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
         Left            =   -74430
         TabIndex        =   28
         Top             =   2190
         Width           =   945
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Payable Amount"
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
         Left            =   -74820
         TabIndex        =   27
         Top             =   1890
         Width           =   1335
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "PPU"
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
         Left            =   -73800
         TabIndex        =   26
         Top             =   1605
         Width           =   315
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Q supplied"
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
         Left            =   -74370
         TabIndex        =   25
         Top             =   1305
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Trans Date"
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
         Left            =   -74370
         TabIndex        =   24
         Top             =   1020
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Supplier Number"
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
         Index           =   0
         Left            =   -74880
         TabIndex        =   23
         Top             =   720
         Width           =   1395
      End
   End
   Begin VB.ComboBox cbobranch 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8040
      TabIndex        =   8
      Top             =   360
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmimports.frx":03EA
      Left            =   1515
      List            =   "frmimports.frx":03F4
      TabIndex        =   3
      Text            =   "Comma Delimited Text File"
      Top             =   120
      Width           =   6015
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
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
      Left            =   1200
      Picture         =   "frmimports.frx":042C
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   2
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox txtImportedFile 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   5940
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   1560
      TabIndex        =   0
      Top             =   840
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   600
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblbranch 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   8040
      TabIndex        =   10
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "Branch Code"
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
      Left            =   8055
      TabIndex        =   9
      Top             =   120
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "File Format:"
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
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Import File:"
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
      Index           =   27
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   930
   End
   Begin VB.Label lblRecords 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1320
      TabIndex        =   5
      Top             =   1305
      Width           =   6375
   End
   Begin VB.Label lblProgress 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   615
   End
   Begin VB.Menu mnuImport 
      Caption         =   ""
      Begin VB.Menu mnuFosa 
         Caption         =   "Fosa Transactions"
      End
      Begin VB.Menu mnuBosa 
         Caption         =   "Bosa Transactions"
      End
   End
End
Attribute VB_Name = "frmimports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objDs As Clsimportpayments
Dim objDs1 As CLsimportpayments1
Dim objds3 As clsimportsuppliersdeduc
Dim objds4 As clsimporttransdeductions
Dim objds5 As clsimporttransportersdetails
Dim objds6 As clsimportstransportassignment
'Dim colBind As BindingCollection
Dim myclass As Object
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim br As String
Dim bookba As Currency

Private Sub cbobranch_Change()
    Set myclass = New cdbase
    
    Provider = myclass.OpenCon
    
    Set cn = CreateObject("adodb.connection")
    
   cn.Open Provider, "atm", "atm", ""
    
    Set rs = CreateObject("ADODB.Recordset")
    
    rs.Open "SELECT bname from d_branch where bcode='" & cbobranch & "'", cn
        
        If rs.EOF Then Exit Sub
        
        lblbranch = rs!Bname
End Sub

Private Sub cbobranch_Click()
cbobranch_Change
End Sub

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdImport_Click()
    On Error GoTo 10
    Import_Fosa_Trans
    Exit Sub
10:    MsgBox err.description
End Sub

Private Sub Import_Bosa_Trans()
    On Error GoTo SysError
    If Trim$(txtImportedFile) = "" Then
        MsgBox "Please select the file to Import", vbInformation, Me.Caption
        Exit Sub
    End If
    Dim Bosa_FSO As New FileSystemObject, LoanFIle As TextStream, strData As String, _
    I As Long, Loanno As String, memberno As String, _
    datereceived As Date, paymentno As Long, amount As Double, principal As Double, interest As _
    Double, IntrCharged As Double, IntrOwed As Double, loanbalance As Double, ReceiptNo As _
    String, Remarks As String, auditid As String, transby As String, MyRecs As Long
    Set LoanFIle = Bosa_FSO.OpenTextFile(txtImportedFile, ForReading, False)
    Do Until LoanFIle.AtEndOfStream
        MyRecs = MyRecs + 1
        strData = LoanFIle.ReadLine
    Loop
    strData = ""
    LoanFIle.Close
    ProgressBar1.max = MyRecs
    MyRecs = 0
    Set LoanFIle = Bosa_FSO.OpenTextFile(txtImportedFile, ForReading, False)
    Do Until LoanFIle.AtEndOfStream
        strData = LoanFIle.ReadLine
        MyRecs = MyRecs + 1
        ProgressBar1.value = MyRecs
        DoEvents
        Do Until InStr(1, strData, ",", vbTextCompare) < 0
            I = I + 1
            If InStr(1, strData, ",", vbTextCompare) = 0 Then
                transby = strData
                strData = ""
                I = 0
                
                Exit Do
            Else
                Select Case I
                    Case 1 'LoanNo
                    Loanno = Left(strData, InStr(1, strData, ",", vbTextCompare) - 1)
                    Case 2 'MemberNo
                    memberno = Left(strData, InStr(1, strData, ",", vbTextCompare) - 1)
                    Case 3 'DateReceived
                    datereceived = Left(strData, InStr(1, strData, ",", vbTextCompare) - 1)
                    Case 4 'PaymentNo
                    paymentno = Left(strData, InStr(1, strData, ",", vbTextCompare) - 1)
                    Case 5 'Amount
                    amount = Left(strData, InStr(1, strData, ",", vbTextCompare) - 1)
                    Case 6 'Principal
                    principal = Left(strData, InStr(1, strData, ",", vbTextCompare) - 1)
                    Case 7 'Interest
                    interest = Left(strData, InStr(1, strData, ",", vbTextCompare) - 1)
                    Case 8 'IntrCharged
                    IntrCharged = Left(strData, InStr(1, strData, ",", vbTextCompare) - 1)
                    Case 9 'IntrOwed
                    IntrOwed = Left(strData, InStr(1, strData, ",", vbTextCompare) - 1)
                    Case 10 'LoanBalance
                    loanbalance = Left(strData, InStr(1, strData, ",", vbTextCompare) - 1)
                    Case 11 'ReceiptNo
                    ReceiptNo = Left(strData, InStr(1, strData, ",", vbTextCompare) - 1)
                    Case 12 'Remarks
                    Remarks = Left(strData, InStr(1, strData, ",", vbTextCompare) - 1)
                    Case 13 'AuditID
                    auditid = Left(strData, InStr(1, strData, ",", vbTextCompare) - 1)
                End Select
                strData = Right(strData, Len(strData) - InStr(1, strData, ",", vbTextCompare))
            End If
        Loop
    Loop
    MsgBox "Importing Bosa Transactions Completed Successfully", vbInformation, Me.Caption
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub Import_Fosa_Trans()
    On Error GoTo SysError
    DelimiterConstant = 0
    If Trim(txtImportedFile) = "" Then
        MsgBox "Please select the file to import", vbInformation, Me.Caption
        Exit Sub
    End If
    If UCase("Tab Delimited Text File") = UCase(Combo1) Then
        DelimiterConstant = 9
    End If
    If UCase("Comma Delimited Text File") = UCase(Combo1) Then
        DelimiterConstant = 44
    End If
    If DelimiterConstant = 0 Then
        MsgBox "The selected File Format Not Supported. Try selecting again .", vbExclamation
        Exit Sub
    End If
    Call BindText
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub
Sub BindText_transportsdetails()
    Set objds5 = New clsimporttransportersdetails
    txttranscode5.Text = objds5.rs("txttranscode5")
    txtname5.Text = objds5.rs("txtname5")
    txtcertno5.Text = objds5.rs("txtcertno5")
    txtlocation5.Text = objds5.rs("txtlocation5")
    txtregdate5.Text = objds5.rs("txtregdate5")
    txtemail5.Text = objds5.rs("txtemail5")
    txtphoneno5.Text = objds5.rs("txtphoneno5")
    txttown5.Text = objds5.rs("txttown5")
    txtaddress5.Text = objds5.rs("txtaddress5")
    txtsubsidy5.Text = objds5.rs("txtsubsidy5")
    txtaccno5.Text = objds5.rs("txtaccno5")
    txtbcode5.Text = objds5.rs("txtbcode5")
    txtbbranch5.Text = objds5.rs("txtbbranch5")
    txtactive5.Text = objds5.rs("txtactive5")
    txttbranch5.Text = objds5.rs("txttbranch5")
    txtauditid5.Text = objds5.rs("txtauditid5")
    txtauditdatetime5.Text = objds5.rs("txtauditdatetime5")
    txtisfrate5.Text = objds5.rs("txtisfrate5")
    txtrate5.Text = objds5.rs("txtrate5")
    txtbr5.Text = objds5.rs("txtbr5")
    
    
    
    
    MsgBox "HAS SUCCESSFULLY IMPORTED RECORDS", vbInformation
End Sub
Sub BindText_suppliers()
    Set objDs1 = New CLsimportpayments1
    txtsno1.Text = objDs1.rs("txtsno1")
    txtregdate.Text = objDs1.rs("txtregdate")
    txtIdNo.Text = objDs1.rs("txtidno")
    txtNames.Text = objDs1.rs("txtnames")
    Txtaccno.Text = objDs1.rs("txtaccno")
    txtBCode.Text = objDs1.rs("txtbcode")
    txtbbranch.Text = objDs1.rs("txtbbranch")
    txtType.Text = objDs1.rs("txttype")
    txtVillage.Text = objDs1.rs("txtvillage")
    txtlocations.Text = objDs1.rs("txtlocations")
    txtDivision.Text = objDs1.rs("txtdivision")
    txtDistrict.Text = objDs1.rs("txtdistrict")
    txttrader.Text = objDs1.rs("txttrader")
    txtactive.Text = objDs1.rs("txtactive")
    txtbranch.Text = objDs1.rs("txtbranch")
    txtPhone.Text = objDs1.rs("txtphone")
    txtAddress.Text = objDs1.rs("txtaddress")
    txtTown.Text = objDs1.rs("txttown")
    txtEmail.Text = objDs1.rs("txtemail")
    txttranscode.Text = objDs1.rs("txttranscode")
    txtauditid1.Text = objDs1.rs("txtauditid1")
    txtauditdatetime1.Text = objDs1.rs("txtauditdatetime1")
    txtscode.Text = objDs1.rs("txtscode")
    txtloan.Text = objDs1.rs("txtloan")
    txtcompare.Text = objDs1.rs("txtcompare")
    txtisfrare.Text = objDs1.rs("txtisfrare")
    txtfrate.Text = objDs1.rs("txtfrate")
    txtrate.Text = objDs1.rs("txtrate")
    txthast.Text = objDs1.rs("txthast")
    TXTBR1.Text = objDs1.rs("txtbr")
    
    
    
    MsgBox "HAS SUCCESSFULLY IMPORTED RECORDS", vbInformation
End Sub
Sub BindText_trans_Deduc()
    Set objds4 = New clsimporttransdeductions
    txtsno4.Text = objds4.rs("txtsno4")
    txtdatededuc2.Text = objds4.rs("txtdatededuc2")
    txtdescription1.Text = objds4.rs("txtdescription1")
    txtamount1.Text = objds4.rs("txtamount1")
    txtperiod1.Text = objds4.rs("txtperiod1")
    txtstartdate1.Text = objds4.rs("txtstartdate1")
    txtenddate1.Text = objds4.rs("txtenddate1")
    txtauditid4.Text = objds4.rs("txtauditdate4")
    txtauditdate4.Text = objds4.rs("txtauditdate4")
    txtyear1.Text = objds4.rs("txtyear1")
    txttrate.Text = objds4.rs("txttrate")
    txtbr4.Text = objds4.rs("txtbr4")
    
    MsgBox "HAS SUCCESSFULLY IMPORTED RECORDS", vbInformation
End Sub
Sub BindText_Supp_Deduc()
    Set objds3 = New clsimportsuppliersdeduc
    txtsno3.Text = objds3.rs("txtsno3")
    txtdatededuc.Text = objds3.rs("txtdatededuc")
    txtdescription.Text = objds3.rs("txtdescription")
    txtamout.Text = objds3.rs("txtamout")
    txtperiod.Text = objds3.rs("txtperiod")
    txtstartdate.Text = objds3.rs("txtstartdate")
    txtenddate.Text = objds3.rs("txtenddate")
    txtauditid3.Text = objds3.rs("txtauditdate3")
    txtauditdate3.Text = objds3.rs("txtauditdate3")
    txtyear.Text = objds3.rs("txtyear")
    txtremarks.Text = objds3.rs("txtremarks")
    txtbr3.Text = objds3.rs("txtbr3")
    
    MsgBox "HAS SUCCESSFULLY IMPORTED RECORDS", vbInformation
End Sub
Sub BindText()
    Set objDs = New Clsimportpayments
    txtSNo.Text = objDs.rs("txtsno")
    txttransdate.Text = objDs.rs("txttransdate")
    txtqsupplied.Text = objDs.rs("txtqsupplied")
    txtppu.Text = objDs.rs("txtppu")
    txtpamount.Text = objDs.rs("txtpamount")
    txttranstime.Text = objDs.rs("txttranstime")
    txtauditid.Text = objDs.rs("txtauditid")
    txtauditdatetime.Text = objDs.rs("txtauditdatetime")
    txtpaid.Text = objDs.rs("txtpaid")
    txtLR.Text = objDs.rs("txtLR")
    txtremark.Text = objDs.rs("txtremark")
    txtbr.Text = objDs.rs("txtbr")
    
    MsgBox "HAS SUCCESSFULLY IMPORTED RECORDS", vbInformation
End Sub

Private Sub cmdimportsupplierdeduct_Click()
 On Error GoTo SysError
    DelimiterConstant = 0
    If Trim(txtImportedFile) = "" Then
        MsgBox "Please select the file to import", vbInformation, Me.Caption
        Exit Sub
    End If
    If UCase("Tab Delimited Text File") = UCase(Combo1) Then
        DelimiterConstant = 9
    End If
    If UCase("Comma Delimited Text File") = UCase(Combo1) Then
        DelimiterConstant = 44
    End If
    If DelimiterConstant = 0 Then
        MsgBox "The selected File Format Not Supported. Try selecting again .", vbExclamation
        Exit Sub
    End If
    Call BindText_Supp_Deduc
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub
Sub BindText_Transportassignment()
    Set objds6 = New clsimportstransportassignment
    txttranscode7.Text = objds6.rs("txttranscode7")
    txtsno7.Text = objds6.rs("txtsno7")
    txtrate7.Text = objds6.rs("txtrate7")
    txtstartdate7.Text = objds6.rs("txtstartdate7")
    txtacitve7.Text = objds6.rs("txtacitve7")
    txtdateinactive7.Text = objds6.rs("txtdateinactive7")
    txtauditid7.Text = objds6.rs("txtauditid7")
    txtauditdatetime7.Text = objds6.rs("txtauditdatetime7")
    txtisfrate7.Text = objds6.rs("txtisfrate7")
    txtbr7.Text = objds6.rs("txtbr7")
    
    
    MsgBox "HAS SUCCESSFULLY IMPORTED RECORDS", vbInformation
End Sub
Private Sub cmdimporttransdeduc_Click()
On Error GoTo SysError
    DelimiterConstant = 0
    If Trim(txtImportedFile) = "" Then
        MsgBox "Please select the file to import", vbInformation, Me.Caption
        Exit Sub
    End If
    If UCase("Tab Delimited Text File") = UCase(Combo1) Then
        DelimiterConstant = 9
    End If
    If UCase("Comma Delimited Text File") = UCase(Combo1) Then
        DelimiterConstant = 44
    End If
    If DelimiterConstant = 0 Then
        MsgBox "The selected File Format Not Supported. Try selecting again .", vbExclamation
        Exit Sub
    End If
    Call BindText_trans_Deduc
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub cmdimporttransportdetails_Click()
On Error GoTo SysError
    DelimiterConstant = 0
    If Trim(txtImportedFile) = "" Then
        MsgBox "Please select the file to import", vbInformation, Me.Caption
        Exit Sub
    End If
    If UCase("Tab Delimited Text File") = UCase(Combo1) Then
        DelimiterConstant = 9
    End If
    If UCase("Comma Delimited Text File") = UCase(Combo1) Then
        DelimiterConstant = 44
    End If
    If DelimiterConstant = 0 Then
        MsgBox "The selected File Format Not Supported. Try selecting again .", vbExclamation
        Exit Sub
    End If
    Call BindText_transportsdetails
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub cmdPost_Click()
    On Error GoTo h
    Dim MFSO As New FileSystemObject, strData As String, ImportFile As TextStream, _
    SigI1D As String, Sig2ID As String, Sig3ID As String, Sig4ID As String, _
    SigName1 As String, SigName2 As String, SigName3 As String, SigName4 As String, _
    Nomi1ID As String, Nomi2ID As String, Nomi3ID As String, NomiName As String, _
    Nomi2Name As String, Nomi3Name As String, ACCNO As String, I As Long, Nominees As Boolean
    Set myclass = New cdbase
    Provider = myclass.OpenCon
    Set cn = CreateObject("ADODB.Connection")
    cn.Open Provider, "atm", "atm", ""
    Dim Has_Transactions As Boolean
    If Trim$(txtImportedFile) = "" Then
        MsgBox "Please select the file to import from", vbInformation, Me.Caption
        Exit Sub
    End If
    If Not MFSO.FileExists(txtImportedFile) Then
        MsgBox "The selected file does not exist." & vbCrLf _
        & "Please confirm the path.", vbInformation, Me.Caption
        Exit Sub
    End If
    If cbobranch = "" Then
        MsgBox "Please enter the branch code before you proceed", vbInformation, "Posting The daily branch updates"
        Exit Sub
    End If
    I = 0
    While objDs.rs.EOF = False And Trim(txtImportedFile) <> ""
        Me.Refresh
        DoEvents
        ProgressBar1.value = objDs.rs.AbsolutePosition
        lblProgress = CSng(objDs.rs.AbsolutePosition) / CSng(objDs.rs.RecordCount) * 100 & " %"
        lblRecords = "Completed Importing " & CSng(objDs.rs.AbsolutePosition) & " Records"
        Dim gename As String
        txtSNo.Text = Trim(objDs.rs("txtsno"))
        txttransdate.Text = Trim(objDs.rs("txttransdate"))
        txtqsupplied.Text = objDs.rs("txtqsupplied")
        txtppu.Text = Trim(objDs.rs("txtppu"))
        txtpamount.Text = objDs.rs("txtpamount")
        txttranstime.Text = Trim(objDs.rs("txttranstime"))
        txtauditid.Text = Trim(objDs.rs("txtauditid"))
        txtauditdatetime.Text = Trim(objDs.rs("txtauditdatetime"))
        txtpaid.Text = Trim(objDs.rs("txtpaid"))
        txtLR.Text = Trim(objDs.rs("txtLR"))
        txtremark.Text = Trim(objDs.rs("txtremark"))
        txtbr.Text = objDs.rs("txtbr")
        
        '// check if the payrollno do exist
        sql = ""
        Dim rscheck As Recordset
        Set rscheck = New ADODB.Recordset
        sql = "select br from  d_company where br='" & Trim(txtbr.Text) & "'"
        Set rscheck = oSaccoMaster.GetRecordset(sql)
        
        If Not rscheck.EOF Then
        If Trim(txtbr) = Trim(rscheck.Fields(0)) Then
        MsgBox "You are trying to run for the same branch it details. Pole", vbInformation
        
        Exit Sub
        GoTo kapsiriat
        End If
        End If
        sql = ""
        sql = "set dateformat dmy SELECT     *   FROM         d_Milkintake where sno='" & txtSNo & "' and transdate='" & txttransdate & "' and qsupplied=" & txtqsupplied & " and auditid='" & txtauditid & "' and transtime='" & txttranstime & "' and br='" & Trim(txtbr) & "'"
        Set rs = oSaccoMaster.GetRecordset(sql)
        If rs.EOF Then
        '//save it here
            Set cn = New ADODB.Connection
                        Set cn = New ADODB.Connection
            sql = "set dateformat dmy INSERT INTO d_Milkintake"
            sql = sql & "      (SNo, TransDate, QSupplied, PPU, PAmount, TransTime, AuditId, auditdatetime, Paid, LR, Remarks, BR)"
            sql = sql & " VALUES     (" & txtSNo & ",'" & txttransdate & "'," & txtqsupplied & "," & txtppu & "," & txtpamount & ",'" & txttranstime & "','" & txtauditid & "','" & txtauditdatetime & "'," & txtpaid & "," & txtLR & ",'" & txtremark & "','" & Trim(txtbr) & "')"
            oSaccoMaster.ExecuteThis (sql)
            
        End If
        
        
        I = I + 1
        ProgressBar1.max = objDs.rs.RecordCount
        objDs.rs.MoveNext
    Wend
    'objDs.rs.MoveFirst
    
    Nominees = False
    MsgBox "Records were posted successfully.", vbInformation
    ProgressBar1.value = 0
    Label2 = ""
    lblRecords = "Completed Importing " & objDs.rs.RecordCount & " Records"
kapsiriat:
    Exit Sub
h:
    MsgBox err.description
    Nominees = False
End Sub

Sub Read_Files()
   Dim fso As New FileSystemObject, txtFile, _
      fil1 As File, ts As TextStream
    fso.CopyFile "c:\testfile.txt", True
    Dim S As String
    MsgBox "reading file"
    ' Write a line.
    Set fil1 = fso.GetFile("c:\testfile.txt")
    Set ts = fil1.OpenAsTextStream(ForReading)
    'ts.Read fil1
    'ts.Column
    ts.Close
    ' Read the contents of the file.
    Set ts = fil1.OpenAsTextStream(ForReading)
    'For ts = 1 To ts.ReadAll
   ' s = ts.ReadLine
    'MsgBox s
    
   ' Next ts
    'ts.Close

End Sub

Private Sub cmdpostsuppliers_Click()
    On Error GoTo h
    Dim MFSO As New FileSystemObject, strData As String, ImportFile As TextStream, _
    SigI1D As String, Sig2ID As String, Sig3ID As String, Sig4ID As String, _
    SigName1 As String, SigName2 As String, SigName3 As String, SigName4 As String, _
    Nomi1ID As String, Nomi2ID As String, Nomi3ID As String, NomiName As String, _
    Nomi2Name As String, Nomi3Name As String, ACCNO As String, I As Long, Nominees As Boolean
    Set myclass = New cdbase
    Provider = myclass.OpenCon
    Set cn = CreateObject("ADODB.Connection")
    cn.Open Provider, "atm", "atm"
    Dim Has_Transactions As Boolean
    If Trim$(txtImportedFile) = "" Then
        MsgBox "Please select the file to import from", vbInformation, Me.Caption
        Exit Sub
    End If
    If Not MFSO.FileExists(txtImportedFile) Then
        MsgBox "The selected file does not exist." & vbCrLf _
        & "Please confirm the path.", vbInformation, Me.Caption
        Exit Sub
    End If
    If cbobranch = "" Then
        MsgBox "Please enter the branch code before you proceed", vbInformation, "Posting The daily branch updates"
        Exit Sub
    End If
    I = 0
    While objDs1.rs.EOF = False And Trim(txtImportedFile) <> ""
        Me.Refresh
        DoEvents
        ProgressBar1.value = objDs1.rs.AbsolutePosition
        lblProgress = CSng(objDs1.rs.AbsolutePosition) / CSng(objDs1.rs.RecordCount) * 100 & " %"
        lblRecords = "Completed Importing " & CSng(objDs1.rs.AbsolutePosition) & " Records"
        Dim gename As String
        txtsno1.Text = objDs1.rs("txtsno1")
        txtregdate.Text = objDs1.rs("txtregdate")
        txtIdNo.Text = objDs1.rs("txtidno")
        txtNames.Text = objDs1.rs("txtnames")
        Txtaccno.Text = objDs1.rs("txtaccno")
        txtBCode.Text = objDs1.rs("txtbcode")
        txtbbranch.Text = objDs1.rs("txtbbranch")
        txtType.Text = objDs1.rs("txttype")
        txtVillage.Text = objDs1.rs("txtvillage")
        txtlocations.Text = objDs1.rs("txtlocations")
        txtDivision.Text = objDs1.rs("txtdivision")
        txtDistrict.Text = objDs1.rs("txtdistrict")
        txttrader.Text = objDs1.rs("txttrader")
        txtactive.Text = objDs1.rs("txtactive")
        txtbranch.Text = objDs1.rs("txtbranch")
        txtPhone.Text = objDs1.rs("txtphone")
        txtAddress.Text = objDs1.rs("txtaddress")
        txtTown.Text = objDs1.rs("txttown")
        txtEmail.Text = objDs1.rs("txtemail")
        txttranscode.Text = objDs1.rs("txttranscode")
        txtauditid1.Text = objDs1.rs("txtauditid1")
        txtauditdatetime1.Text = objDs1.rs("txtauditdatetime1")
        txtscode.Text = objDs1.rs("txtscode")
        txtloan.Text = objDs1.rs("txtloan")
        txtcompare.Text = objDs1.rs("txtcompare")
        txtisfrare.Text = objDs1.rs("txtisfrare")
        txtfrate.Text = objDs1.rs("txtfrate")
        txtrate.Text = objDs1.rs("txtrate")
        txthast.Text = objDs1.rs("txthast")
        TXTBR1.Text = objDs1.rs("txtbr")
        
        '// check if the payrollno do exist
        sql = ""
        Dim rscheck As Recordset
        Set rscheck = New ADODB.Recordset
        sql = "select br from  d_company where br='" & Trim(txtbr.Text) & "'"
        Set rscheck = oSaccoMaster.GetRecordset(sql)
        
        If Not rscheck.EOF Then
        If Trim(txtbr) = Trim(rscheck.Fields(0)) Then
        MsgBox "You are trying to run for the same branch it details. Pole", vbInformation
        
        Exit Sub
        GoTo kapsiriat
        End If
        End If
        sql = ""
        sql = "set dateformat dmy SELECT     *   FROM         d_Suppliers where sno=" & txtsno1 & ""
        Set rs = oSaccoMaster.GetRecordset(sql)
        If rs.EOF Then
        '//save it here
            Set cn = New ADODB.Connection
           Set cn = New ADODB.Connection
           Dim AC As Integer, tr As Integer
           If txtactive = True Then AC = 1 Else AC = 0
           If txttrader = False Then tr = 0 Else tr = 1
            sql = "set dateformat dmy INSERT INTO d_Suppliers"
            sql = sql & " (SNo, Regdate, IdNo, [Names], AccNo, Bcode, BBranch, Type, Village, Location, Division, District, Trader, active, Branch, PhoneNo, Address, Town,"
            sql = sql & " Email, TransCode, AuditId, auditdatetime, scode, Loan, Compare, isfrate, frate,rate, hast, BR)"
            sql = sql & " VALUES     (" & txtsno1 & ",'" & txtregdate & "','" & txtIdNo & "','" & txtNames & "','" & Txtaccno & "','" & txtBCode & "',"
            sql = sql & " '" & txtbbranch & "','" & txtType & "','" & txtVillage & "','" & txtlocations & "','" & txtDivision & "','" & txtDistrict & "',"
            sql = sql & " " & tr & "," & AC & ",'" & txtbranch & "','" & txtPhone & "','" & txtAddress & "','" & txtTown & "','" & txtEmail & "','" & txttranscode & "','" & txtauditid1 & "','" & txtauditdatetime1 & "','" & txtscode & "'," & txtloan & ",'" & txtcompare & "','" & txtisfrare & "','" & txtfrate & "','" & txtrate & "'," & txthast & ",'" & TXTBR1 & "')"
            oSaccoMaster.ExecuteThis (sql)
           
        End If
        
        
        I = I + 1
        ProgressBar1.max = objDs1.rs.RecordCount
        objDs1.rs.MoveNext
    Wend
    'objDs.rs.MoveFirst
    
    Nominees = False
    MsgBox "Records were posted successfully.", vbInformation
    ProgressBar1.value = 0
    Label2 = ""
    lblRecords = "Completed Importing " & objDs1.rs.RecordCount & " Records"
kapsiriat:
    Exit Sub
h:
    MsgBox err.description
    Nominees = False
End Sub

Private Sub cmdposttransdeductions_Click()
    On Error GoTo h
    Dim MFSO As New FileSystemObject, strData As String, ImportFile As TextStream, _
    SigI1D As String, Sig2ID As String, Sig3ID As String, Sig4ID As String, _
    SigName1 As String, SigName2 As String, SigName3 As String, SigName4 As String, _
    Nomi1ID As String, Nomi2ID As String, Nomi3ID As String, NomiName As String, _
    Nomi2Name As String, Nomi3Name As String, ACCNO As String, I As Long, Nominees As Boolean
    Set myclass = New cdbase
    Provider = myclass.OpenCon
    Set cn = CreateObject("ADODB.Connection")
    cn.Open Provider, "atm", "atm"
    Dim Has_Transactions As Boolean
    If Trim$(txtImportedFile) = "" Then
        MsgBox "Please select the file to import from", vbInformation, Me.Caption
        Exit Sub
    End If
    If Not MFSO.FileExists(txtImportedFile) Then
        MsgBox "The selected file does not exist." & vbCrLf _
        & "Please confirm the path.", vbInformation, Me.Caption
        Exit Sub
    End If
    If cbobranch = "" Then
        MsgBox "Please enter the branch code before you proceed", vbInformation, "Posting The daily branch updates"
        Exit Sub
    End If
    I = 0
    While objds4.rs.EOF = False And Trim(txtImportedFile) <> ""
        Me.Refresh
        DoEvents
        ProgressBar1.value = objds4.rs.AbsolutePosition
        lblProgress = CSng(objds4.rs.AbsolutePosition) / CSng(objds4.rs.RecordCount) * 100 & " %"
        lblRecords = "Completed Importing " & CSng(objds4.rs.AbsolutePosition) & " Records"
        Dim gename As String
        txtsno4.Text = objds4.rs("txtsno4")
        txtdatededuc2.Text = objds4.rs("txtdatededuc2")
        txtdescription1.Text = objds4.rs("txtdescription1")
        txtamount1.Text = objds4.rs("txtamount1")
        txtperiod1.Text = objds4.rs("txtperiod1")
        txtstartdate1.Text = objds4.rs("txtstartdate1")
        txtenddate1.Text = objds4.rs("txtenddate1")
        txtauditid4.Text = objds4.rs("txtauditdate4")
        txtauditdate4.Text = objds4.rs("txtauditdate4")
        txtyear1.Text = objds4.rs("txtyear1")
        txttrate.Text = objds4.rs("txttrate")
        txtbr4.Text = objds4.rs("txtbr4")
            
        '// check if the payrollno do exist
        sql = ""
        Dim rscheck As Recordset
        Set rscheck = New ADODB.Recordset
        sql = "select br from  d_company where br='" & Trim(txtbr.Text) & "'"
        Set rscheck = oSaccoMaster.GetRecordset(sql)
        
        If Not rscheck.EOF Then
        If Trim(txtbr) = Trim(rscheck.Fields(0)) Then
        MsgBox "You are trying to run for the same branch it details. Pole", vbInformation
        
        Exit Sub
        GoTo kapsiriat
        End If
        End If
        sql = ""
        sql = "set dateformat dmy SELECT     *   FROM         d_Transport_Deduc where TransCode='" & txtsno4 & "' and tdate_deduc='" & txtdatededuc2 & "'  and auditid='" & txtauditid4 & "'  and br='" & Trim(txtbr4) & "'"
        Set rs = oSaccoMaster.GetRecordset(sql)
        If rs.EOF Then
        '//save it here
        Set cn = New ADODB.Connection
        
        sql = "set dateformat dmy INSERT INTO d_Transport_Deduc"
        sql = sql & "      ( TransCode, TDate_Deduc, Description, Amount, Period, StartDate, EndDate, auditid, auditdatetime, yyear, RATE, BR)"
        sql = sql & " VALUES     ('" & txtsno4 & "','" & txtdatededuc2 & "','" & txtdescription1 & "'," & txtamount1 & ",'" & txtperiod1 & "','" & txtstartdate1 & "','" & txtenddate1 & "','" & txtauditid4 & "','" & txtauditdate4 & "'," & txtyear1 & ",'" & txttrate & "','" & Trim(txtbr4) & "')"
        oSaccoMaster.ExecuteThis (sql)

        End If
        
        
        I = I + 1
        ProgressBar1.max = objds4.rs.RecordCount
        objds4.rs.MoveNext
    Wend
    'objDs.rs.MoveFirst
    
    Nominees = False
    MsgBox "Records were posted successfully.", vbInformation
    ProgressBar1.value = 0
    Label2 = ""
    lblRecords = "Completed Importing " & objds4.rs.RecordCount & " Records"
kapsiriat:
    Exit Sub
h:
    MsgBox err.description
    Nominees = False

End Sub

Private Sub cmdposttransportersdetails_Click()
    On Error GoTo h
    Dim MFSO As New FileSystemObject, strData As String, ImportFile As TextStream, _
    SigI1D As String, Sig2ID As String, Sig3ID As String, Sig4ID As String, _
    SigName1 As String, SigName2 As String, SigName3 As String, SigName4 As String, _
    Nomi1ID As String, Nomi2ID As String, Nomi3ID As String, NomiName As String, _
    Nomi2Name As String, Nomi3Name As String, ACCNO As String, I As Long, Nominees As Boolean
    Set myclass = New cdbase
    Provider = myclass.OpenCon
    Set cn = CreateObject("ADODB.Connection")
    cn.Open Provider, "atm", "atm"
    Dim Has_Transactions As Boolean
    If Trim$(txtImportedFile) = "" Then
        MsgBox "Please select the file to import from", vbInformation, Me.Caption
        Exit Sub
    End If
    If Not MFSO.FileExists(txtImportedFile) Then
        MsgBox "The selected file does not exist." & vbCrLf _
        & "Please confirm the path.", vbInformation, Me.Caption
        Exit Sub
    End If
    If cbobranch = "" Then
        MsgBox "Please enter the branch code before you proceed", vbInformation, "Posting The daily branch updates"
        Exit Sub
    End If
    I = 0
    While objds5.rs.EOF = False And Trim(txtImportedFile) <> ""
        Me.Refresh
        DoEvents
        ProgressBar1.value = objds5.rs.AbsolutePosition
        lblProgress = CSng(objds5.rs.AbsolutePosition) / CSng(objds5.rs.RecordCount) * 100 & " %"
        lblRecords = "Completed Importing " & CSng(objds5.rs.AbsolutePosition) & " Records"
        Dim gename As String
        
        txttranscode5.Text = objds5.rs("txttranscode5")
        txtname5.Text = objds5.rs("txtname5")
        txtcertno5.Text = objds5.rs("txtcertno5")
        txtlocation5.Text = objds5.rs("txtlocation5")
        txtregdate5.Text = objds5.rs("txtregdate5")
        txtemail5.Text = objds5.rs("txtemail5")
        txtphoneno5.Text = objds5.rs("txtphoneno5")
        txttown5.Text = objds5.rs("txttown5")
        txtaddress5.Text = objds5.rs("txtaddress5")
        txtsubsidy5.Text = objds5.rs("txtsubsidy5")
        txtaccno5.Text = objds5.rs("txtaccno5")
        txtbcode5.Text = objds5.rs("txtbcode5")
        txtbbranch5.Text = objds5.rs("txtbbranch5")
        txtactive5.Text = objds5.rs("txtactive5")
        txttbranch5.Text = objds5.rs("txttbranch5")
        txtauditid5.Text = objds5.rs("txtauditid5")
        txtauditdatetime5.Text = objds5.rs("txtauditdatetime5")
        txtisfrate5.Text = objds5.rs("txtisfrate5")
        txtrate5.Text = objds5.rs("txtrate5")
        txtbr5.Text = objds5.rs("txtbr5")
    
    
        
        '// check if the payrollno do exist
        sql = ""
        Dim rscheck As Recordset
        Set rscheck = New ADODB.Recordset
        sql = "select br from  d_company where br='" & Trim(txtbr.Text) & "'"
        Set rscheck = oSaccoMaster.GetRecordset(sql)
        
        If Not rscheck.EOF Then
        If Trim(txtbr) = Trim(rscheck.Fields(0)) Then
        MsgBox "You are trying to run for the same branch it details. Pole", vbInformation
        
        Exit Sub
        GoTo kapsiriat
        End If
        End If
        sql = ""
        sql = "set dateformat dmy SELECT     *   FROM         d_Transporters where transcode='" & txttranscode5 & "' "
        Set rs = oSaccoMaster.GetRecordset(sql)
        If rs.EOF Then
        '//save it here
            Set cn = New ADODB.Connection
           Set cn = New ADODB.Connection
           Dim AC As Integer, tr As Integer
           If txtactive5 = True Then AC = 1 Else AC = 0
           
            sql = "set dateformat dmy INSERT INTO d_Transporters"
            sql = sql & "      (TransCode, TransName, CertNo, Locations, TregDate, email, Phoneno, Town, Address, Subsidy, Accno, Bcode, BBranch, Active, TBranch, auditid,"
            sql = sql & "     auditdatetime, isfrate, rate, BR)"
            sql = sql & "  VALUES     ('" & txttranscode5 & "','" & txtname5 & "','" & txtcertno5 & "','" & txtlocation5 & "','" & txtregdate5 & "','" & txtemail5 & "',"
            sql = sql & " '" & txtphoneno5 & "','" & txttown5 & "','" & txtaddress5 & "'," & txtsubsidy5 & ",'" & txtaccno5 & "','" & txtbcode5 & "','" & txtbbranch5 & "',"
            sql = sql & " " & AC & ",'" & txttbranch5 & "','" & txtauditid5 & "','" & txtauditdatetime5 & "','" & txtfrate & "'," & txtrate5 & ",'" & txtbr5 & "')"
            oSaccoMaster.ExecuteThis (sql)
           
        End If
        
        
        I = I + 1
        ProgressBar1.max = objds5.rs.RecordCount
        objds5.rs.MoveNext
    Wend
    'objDs.rs.MoveFirst
    
    Nominees = False
    MsgBox "Records were posted successfully.", vbInformation
    ProgressBar1.value = 0
    Label2 = ""
    lblRecords = "Completed Importing " & objds5.rs.RecordCount & " Records"
kapsiriat:
    Exit Sub
h:
    MsgBox err.description
    Nominees = False

End Sub

Private Sub cmdupdate_Click()
Dim prov4 As String
Dim EDMS As New ADODB.Connection
prov4 = "EDMS"
Set EDMS = New ADODB.Connection
EDMS.Open prov4
Dim txtSNo As Long, txttransdate As Date, txtqsupplied As Double
Dim txtppu As Double, txtpamount As Double, txttranstime As String
Dim txtauditid As String, txtauditdatetime As Date, txtpaid As Integer
Dim txtLR As Integer, txtremark As String, txtbr As String, tripno As Integer
'required to go on
'SELECT     SNo, TransDate, QSupplied, PPU, PAmount, TransTime, AuditId, auditdatetime
'From d_Milkintake
'ORDER BY SNo DESC

'///SELECT     MemNo, WDate, WTime, WKgs, Operator  From MilkTransaction
sql = ""
sql = "set dateformat dmy SELECT     MemNo, WDate, WTime, WKgs, Operator,tripno,receiptno  From MilkTransaction where wdate='" & transdate & "' order by 1"
Set rst = New ADODB.Recordset
rst.Open sql, EDMS, adOpenKeyset, adLockOptimistic
While Not rst.EOF
txtSNo = rst.Fields(0)
txttransdate = rst.Fields(1)
txtqsupplied = rst.Fields(3)
txttranstime = rst.Fields(2)
txtauditid = rst.Fields(4)
tripno = rst.Fields(5)
txtremark = "Intake"
txtbr = rst.Fields(6)
Set rs = New ADODB.Recordset
sql = "SELECT Price from d_Price"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
txtppu = rs!Price
txtpamount = CCur(txtppu) * CDbl(txtqsupplied)
End If
  sql = ""
        sql = "set dateformat dmy SELECT     *   FROM         d_Milkintake where sno='" & txtSNo & "' and transdate='" & txttransdate & "' and qsupplied=" & txtqsupplied & " and auditid='" & txtauditid & "' and transtime='" & txttranstime & "' and tripno=" & tripno & " and br='" & txtbr & "'"
        Set rs = oSaccoMaster.GetRecordset(sql)
        If rs.EOF Then
        '//save it here
            Set cn = New ADODB.Connection
                        Set cn = New ADODB.Connection
            sql = "set dateformat dmy INSERT INTO d_Milkintake"
            sql = sql & "      (SNo, TransDate, QSupplied, PPU, PAmount, TransTime, AuditId, auditdatetime, Paid, LR, Remarks, BR,tripno)"
            sql = sql & " VALUES     (" & txtSNo & ",'" & txttransdate & "'," & txtqsupplied & "," & txtppu & "," & txtpamount & ",'" & txttranstime & "','" & txtauditid & "','" & Get_Server_Date & "',0,0,'" & txtremark & "','" & Trim(txtbr) & "'," & tripno & ")"
            oSaccoMaster.ExecuteThis (sql)
            
        End If
        frmimports.Caption = txtSNo
        rst.MoveNext
Wend
MsgBox "Records successfully imported"
End Sub

Private Sub Command1_Click()
On Error GoTo SysError
    DelimiterConstant = 0
    If Trim(txtImportedFile) = "" Then
        MsgBox "Please select the file to import", vbInformation, Me.Caption
        Exit Sub
    End If
    If UCase("Tab Delimited Text File") = UCase(Combo1) Then
        DelimiterConstant = 9
    End If
    If UCase("Comma Delimited Text File") = UCase(Combo1) Then
        DelimiterConstant = 44
    End If
    If DelimiterConstant = 0 Then
        MsgBox "The selected File Format Not Supported. Try selecting again .", vbExclamation
        Exit Sub
    End If
    Call BindText_suppliers
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
On Error GoTo SysError
    DelimiterConstant = 0
    If Trim(txtImportedFile) = "" Then
        MsgBox "Please select the file to import", vbInformation, Me.Caption
        Exit Sub
    End If
    If UCase("Tab Delimited Text File") = UCase(Combo1) Then
        DelimiterConstant = 9
    End If
    If UCase("Comma Delimited Text File") = UCase(Combo1) Then
        DelimiterConstant = 44
    End If
    If DelimiterConstant = 0 Then
        MsgBox "The selected File Format Not Supported. Try selecting again .", vbExclamation
        Exit Sub
    End If
    Call BindText_Transportassignment
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Command5_Click()
    On Error GoTo h
    Dim MFSO As New FileSystemObject, strData As String, ImportFile As TextStream, _
    SigI1D As String, Sig2ID As String, Sig3ID As String, Sig4ID As String, _
    SigName1 As String, SigName2 As String, SigName3 As String, SigName4 As String, _
    Nomi1ID As String, Nomi2ID As String, Nomi3ID As String, NomiName As String, _
    Nomi2Name As String, Nomi3Name As String, ACCNO As String, I As Long, Nominees As Boolean
    Set myclass = New cdbase
    Provider = myclass.OpenCon
    Set cn = CreateObject("ADODB.Connection")
   cn.Open Provider, "atm", "atm"
    Dim Has_Transactions As Boolean
    If Trim$(txtImportedFile) = "" Then
        MsgBox "Please select the file to import from", vbInformation, Me.Caption
        Exit Sub
    End If
    If Not MFSO.FileExists(txtImportedFile) Then
        MsgBox "The selected file does not exist." & vbCrLf _
        & "Please confirm the path.", vbInformation, Me.Caption
        Exit Sub
    End If
    If cbobranch = "" Then
        MsgBox "Please enter the branch code before you proceed", vbInformation, "Posting The daily branch updates"
        Exit Sub
    End If
    I = 0
    While objds3.rs.EOF = False And Trim(txtImportedFile) <> ""
        Me.Refresh
        DoEvents
        'ProgressBar1.value = objds3.rs.AbsolutePosition
        lblProgress = CSng(objds3.rs.AbsolutePosition) / CSng(objds3.rs.RecordCount) * 100 & " %"
        lblRecords = "Completed Importing " & CSng(objds3.rs.AbsolutePosition) & " Records"
        Dim gename As String
        txtsno3.Text = objds3.rs("txtsno3")
        txtdatededuc.Text = objds3.rs("txtdatededuc")
        txtdescription.Text = objds3.rs("txtdescription")
        txtamout.Text = objds3.rs("txtamout")
        txtperiod.Text = objds3.rs("txtperiod")
        txtstartdate.Text = objds3.rs("txtstartdate")
        txtenddate.Text = objds3.rs("txtenddate")
        txtauditid3.Text = objds3.rs("txtauditdate3")
        txtauditdate3.Text = objds3.rs("txtauditdate3")
        txtyear.Text = objds3.rs("txtyear")
        txtremarks.Text = objds3.rs("txtremarks")
        txtbr3.Text = objds3.rs("txtbr3")
            
        '// check if the payrollno do exist
        sql = ""
        Dim rscheck As Recordset
        Set rscheck = New ADODB.Recordset
        sql = "select br from  d_company where br='" & Trim(txtbr.Text) & "'"
        Set rscheck = oSaccoMaster.GetRecordset(sql)
        
        If Not rscheck.EOF Then
        If Trim(txtbr) = Trim(rscheck.Fields(0)) Then
        MsgBox "You are trying to run for the same branch it details. Pole", vbInformation
        
        Exit Sub
        GoTo kapsiriat
        End If
        End If
        sql = ""
        sql = "set dateformat dmy SELECT     *   FROM         d_supplier_deduc where sno='" & Trim(txtsno3) & "' and date_deduc='" & txtdatededuc & "'  and auditid='" & Trim(txtauditid3) & "' and amount=" & txtamout & "  and br='" & Trim(txtbr3) & "'"
        Set rs = oSaccoMaster.GetRecordset(sql)
        If rs.EOF Then
        '//save it here
        Set cn = New ADODB.Connection
        
        sql = "set dateformat dmy INSERT INTO d_supplier_deduc"
        sql = sql & "      ( SNo, Date_Deduc, Description, Amount, Period, StartDate, EndDate, auditid, auditdatetime, yyear, Remarks, BR)"
        sql = sql & " VALUES     (" & txtsno3 & ",'" & txtdatededuc & "','" & txtdescription & "'," & txtamout & ",'" & txtperiod & "','" & txtstartdate & "','" & txtenddate & "','" & txtauditid3 & "','" & txtauditdate3 & "'," & txtyear & ",'" & txtremarks & "','" & Trim(txtbr3) & "')"
        oSaccoMaster.ExecuteThis (sql)

        End If
        
        
        I = I + 1
        ProgressBar1.max = objds3.rs.RecordCount
        objds3.rs.MoveNext
    Wend
    'objDs.rs.MoveFirst
    
    Nominees = False
    MsgBox "Records were posted successfully.", vbInformation
    ProgressBar1.value = 0
    Label2 = ""
    lblRecords = "Completed Importing " & objds3.rs.RecordCount & " Records"
kapsiriat:
    Exit Sub
h:
    MsgBox err.description
    Nominees = False

End Sub

Private Sub Command6_Click()
Unload Me
End Sub

Private Sub Command7_Click()
Unload Me
End Sub

Private Sub Command8_Click()
Unload Me
End Sub

Private Sub Command9_Click()
    On Error GoTo h
    Dim MFSO As New FileSystemObject, strData As String, ImportFile As TextStream, _
    SigI1D As String, Sig2ID As String, Sig3ID As String, Sig4ID As String, _
    SigName1 As String, SigName2 As String, SigName3 As String, SigName4 As String, _
    Nomi1ID As String, Nomi2ID As String, Nomi3ID As String, NomiName As String, _
    Nomi2Name As String, Nomi3Name As String, ACCNO As String, I As Long, Nominees As Boolean
    Set myclass = New cdbase
    Provider = myclass.OpenCon
    Set cn = CreateObject("ADODB.Connection")
   cn.Open Provider, "atm", "atm"
    Dim Has_Transactions As Boolean
    If Trim$(txtImportedFile) = "" Then
        MsgBox "Please select the file to import from", vbInformation, Me.Caption
        Exit Sub
    End If
    If Not MFSO.FileExists(txtImportedFile) Then
        MsgBox "The selected file does not exist." & vbCrLf _
        & "Please confirm the path.", vbInformation, Me.Caption
        Exit Sub
    End If
    If cbobranch = "" Then
        MsgBox "Please enter the branch code before you proceed", vbInformation, "Posting The daily branch updates"
        Exit Sub
    End If
    I = 0
    While objds6.rs.EOF = False And Trim(txtImportedFile) <> ""
        Me.Refresh
        DoEvents
        ProgressBar1.value = objds6.rs.AbsolutePosition
        lblProgress = CSng(objds6.rs.AbsolutePosition) / CSng(objds6.rs.RecordCount) * 100 & " %"
        lblRecords = "Completed Importing " & CSng(objds6.rs.AbsolutePosition) & " Records"
        Dim gename As String
        txttranscode7.Text = objds6.rs("txttranscode7")
        txtsno7.Text = objds6.rs("txtsno7")
        txtrate7.Text = objds6.rs("txtrate7")
        txtstartdate7.Text = objds6.rs("txtstartdate7")
        txtacitve7.Text = objds6.rs("txtacitve7")
        txtdateinactive7.Text = objds6.rs("txtdateinactive7")
        txtauditid7.Text = objds6.rs("txtauditid7")
        txtauditdatetime7.Text = objds6.rs("txtauditdatetime7")
        txtisfrate7.Text = objds6.rs("txtisfrate7")
        txtbr7.Text = objds6.rs("txtbr7")
       
            
        '// check if the payrollno do exist
        sql = ""
        Dim rscheck As Recordset
        Set rscheck = New ADODB.Recordset
        sql = "select br from  d_company where br='" & Trim(txtbr7.Text) & "'"
        Set rscheck = oSaccoMaster.GetRecordset(sql)
        
        If Not rscheck.EOF Then
        If Trim(txtbr7) = Trim(rscheck.Fields(0)) Then
        MsgBox "You are trying to run for the same branch it details. Pole", vbInformation
        
        'Exit Sub
        GoTo kapsiriat
        End If
        End If
        sql = ""
        sql = "set dateformat dmy SELECT     *   FROM         d_Transport where sno='" & Trim(txtsno7) & "' and startdate='" & txtstartdate7 & "'  and auditid='" & Trim(txtauditid7) & "' and Trans_Code='" & Trim(txttranscode7) & "' and  br='" & Trim(txtbr7) & "'"
        Set rs = oSaccoMaster.GetRecordset(sql)
        If rs.EOF Then
        '//save it here
        Set cn = New ADODB.Connection
        If txtacitve7 = True Then txtacitve7 = 1 Else txtacitve7 = 0
        sql = "set dateformat dmy INSERT INTO d_Transport"
        sql = sql & "      (trans_code, sno, rate, startdate, active, dateinactivate, auditid, auditdatetime, isfrate, br)"
        sql = sql & " VALUES     ('" & Trim(txttranscode7) & "'," & txtsno7 & "," & txtrate7 & ",'" & txtstartdate7 & "'," & txtacitve7 & ",'" & txtdateinactive7 & "','" & txtauditid7 & "','" & txtauditdatetime7 & "','" & txtisfrate7 & "','" & txtbr7 & "')"
        oSaccoMaster.ExecuteThis (sql)

        End If
        
        
        I = I + 1
        ProgressBar1.max = objds6.rs.RecordCount
        objds6.rs.MoveNext
    Wend
    'objDs.rs.MoveFirst
    
    Nominees = False
    MsgBox "Records were posted successfully.", vbInformation
    ProgressBar1.value = 0
    Label2 = ""
    lblRecords = "Completed Importing " & objds6.rs.RecordCount & " Records"
kapsiriat:
    Exit Sub
h:
    MsgBox err.description
    Nominees = False


End Sub

Private Sub Form_Load()
Set myclass = New cdbase

    Provider = myclass.OpenCon

    Set cn = CreateObject("adodb.connection")

   cn.Open Provider, "atm", "atm"

    Set rs = CreateObject("adodb.recordset")
    
    rs.Open "SELECT BCODE FROM d_Branch", cn
    
    If rs.EOF Then Exit Sub
    
    With rs
        
        While Not .EOF
         
         cbobranch.AddItem rs.Fields(0)
         
         
         .MoveNext
        
        Wend
    
    End With
'Set rs = CreateObject("ADODB.Recordset")
    
'    sql = "SELECT     TOP 1 BR  FROM         d_company"
'    rs.Open sql, cn
'   If Not rs.EOF Then
'        cbobranch.AddItem Trim(rs!br)
'        Else
'        cbobranch = "A"
'   End If
   transdate = Format(Get_Server_Date, "dd/mm/yyyy")
End Sub

Private Sub mnuBosa_Click()
    Import_Bosa_Trans
End Sub

Private Sub mnuFosa_Click()
    Import_Fosa_Trans
End Sub

Private Sub Picture1_Click()
    On Error GoTo 10
    ' Set filters.
    CommonDialog1.Filter = "All Files (*.*)|*.*|Text _Files (*.txt)|*.txt|Excel _CSV(*.CSV)|*.CSV|Batch Files (*.bat)|*.bat"   ' Specify default filter.
    CommonDialog1.FilterIndex = 2   ' Display the Open dialog box.
    CommonDialog1.ShowOpen    ' Call the open file procedure.
    txtImportedFile = CommonDialog1.FileName
    strFileName = txtImportedFile

    Exit Sub
10:    MsgBox err.description
End Sub


