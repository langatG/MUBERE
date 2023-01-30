VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.MDIForm MainForm 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0FFC0&
   Caption         =   "EASYMA"
   ClientHeight    =   9915
   ClientLeft      =   -5910
   ClientTop       =   -3735
   ClientWidth     =   18960
   Icon            =   "MainForm.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MainForm.frx":164A
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   13770
      Left            =   0
      Picture         =   "MainForm.frx":3A06F
      ScaleHeight     =   14158.81
      ScaleMode       =   0  'User
      ScaleWidth      =   18960
      TabIndex        =   0
      Top             =   0
      Width           =   18960
      Begin VB.CheckBox Check1 
         Caption         =   "Show Samli Details"
         BeginProperty Font 
            Name            =   "Bodoni MT"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   7800
         TabIndex        =   18
         Top             =   0
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Frame Frame1 
         Caption         =   "Siche Details"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   9975
         Left            =   17520
         TabIndex        =   5
         Top             =   0
         Width           =   2895
         Begin VB.Label Label13 
            Caption         =   "Label13"
            ForeColor       =   &H008080FF&
            Height          =   375
            Left            =   1800
            TabIndex        =   17
            Top             =   3360
            Width           =   735
         End
         Begin VB.Label Label12 
            Caption         =   "Label12"
            ForeColor       =   &H008080FF&
            Height          =   375
            Left            =   1440
            TabIndex        =   16
            Top             =   2760
            Width           =   975
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   1920
            TabIndex        =   15
            Top             =   1920
            Width           =   735
         End
         Begin VB.Label Label10 
            Caption         =   "Label10"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   1680
            TabIndex        =   14
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Label9"
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   1800
            TabIndex        =   13
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label8 
            Caption         =   "Label8"
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   1920
            TabIndex        =   12
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label7 
            Caption         =   "Today's Kgs"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   11
            Top             =   2880
            Width           =   1815
         End
         Begin VB.Image Image1 
            Height          =   1965
            Left            =   -360
            Picture         =   "MainForm.frx":6C4A8
            Top             =   5520
            Width           =   3630
         End
         Begin VB.Line Line2 
            X1              =   120
            X2              =   4080
            Y1              =   3840
            Y2              =   3840
         End
         Begin VB.Label Label6 
            Caption         =   "This Month  kgs"
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
            TabIndex        =   10
            Top             =   3360
            Width           =   1815
         End
         Begin VB.Line Line1 
            X1              =   120
            X2              =   4080
            Y1              =   2520
            Y2              =   2520
         End
         Begin VB.Label Label5 
            Caption         =   "This Month Active Suppliers"
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
            TabIndex        =   9
            Top             =   1920
            Width           =   1815
         End
         Begin VB.Label Label4 
            Caption         =   "Male Suppliers"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   8
            Top             =   1440
            Width           =   1815
         End
         Begin VB.Label Label3 
            Caption         =   "Female Suppliers"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   7
            Top             =   960
            Width           =   1815
         End
         Begin VB.Label Label2 
            Caption         =   "Today's Suppliers"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   6
            Top             =   480
            Width           =   1815
         End
      End
      Begin MSComctlLib.StatusBar StatusBar2 
         Height          =   855
         Index           =   0
         Left            =   0
         TabIndex        =   3
         Top             =   11040
         Width           =   23895
         _ExtentX        =   42148
         _ExtentY        =   1508
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   9
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Bevel           =   2
               Object.Width           =   4304
               MinWidth        =   4304
               Text            =   "USER : Birgen Gideon K."
               TextSave        =   "USER : Birgen Gideon K."
               Object.ToolTipText     =   "EASYMA User"
            EndProperty
            BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   4304
               MinWidth        =   4304
            EndProperty
            BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   7832
               MinWidth        =   7832
            EndProperty
            BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   5292
               MinWidth        =   5292
               Text            =   "DATE : 07/12/2009"
               TextSave        =   "DATE : 07/12/2009"
               Object.ToolTipText     =   "Today's Date"
            EndProperty
            BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Style           =   5
               Object.Width           =   1764
               MinWidth        =   1764
               TextSave        =   "09:39 PM"
            EndProperty
            BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Style           =   1
               Enabled         =   0   'False
               Object.Width           =   4057
               MinWidth        =   4057
               TextSave        =   "CAPS"
            EndProperty
            BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Style           =   2
               Enabled         =   0   'False
               TextSave        =   "NUM"
            EndProperty
            BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   0
         Top             =   0
      End
      Begin MSComctlLib.StatusBar StatusBar1 
         Height          =   615
         Left            =   240
         TabIndex        =   4
         Top             =   3720
         Visible         =   0   'False
         Width           =   12975
         _ExtentX        =   22886
         _ExtentY        =   1085
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   6
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Text            =   "user"
               TextSave        =   "user"
            EndProperty
            BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   630
         Left            =   -360
         TabIndex        =   19
         Top             =   11280
         Width           =   20250
         _ExtentX        =   35719
         _ExtentY        =   1111
         ButtonWidth     =   609
         ButtonHeight    =   953
         Appearance      =   1
         _Version        =   393216
         Begin MSComCtl2.DTPicker DTPPeriod 
            Height          =   375
            Left            =   16080
            TabIndex        =   20
            Top             =   120
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   114688001
            UpDown          =   -1  'True
            CurrentDate     =   40095
         End
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "MAZIWA SYSTEM"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   495
         Left            =   600
         TabIndex        =   2
         Top             =   6720
         Width           =   7335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "EASYMA"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   495
         Left            =   480
         TabIndex        =   1
         Top             =   6840
         Width           =   5055
      End
   End
   Begin VB.Menu mnuFiles 
      Caption         =   "Files"
      Begin VB.Menu mnuEnquiry 
         Caption         =   "Supplier Enquiry"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuspaces 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFsaaccountinquiry 
         Caption         =   "Fsa account inquiry"
      End
      Begin VB.Menu mnuTransporterEnquiry 
         Caption         =   "Transporter Enquiry"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuLogOff 
         Caption         =   "Log off"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnulogoutt 
         Caption         =   "Logout User"
      End
   End
   Begin VB.Menu mnutransactions 
      Caption         =   "Transactions"
      Begin VB.Menu mnuFarmers 
         Caption         =   "Farmers"
         Begin VB.Menu mnuRegistration 
            Caption         =   "Registration"
         End
         Begin VB.Menu mnuStaffregistration 
            Caption         =   "Staff registration"
         End
         Begin VB.Menu mnuDeductionassignment 
            Caption         =   "Deduction assignment"
         End
         Begin VB.Menu mnuTransportassignment 
            Caption         =   "Transport assignment"
         End
         Begin VB.Menu mnusupap 
            Caption         =   "Approve suppliers"
         End
      End
      Begin VB.Menu mnuTransporters 
         Caption         =   "Transporters"
         Begin VB.Menu mnuregistertransporter 
            Caption         =   "Registrations"
         End
         Begin VB.Menu mnutransportdeductionsassignment 
            Caption         =   "Transport Deductions  Assignment"
         End
      End
      Begin VB.Menu mnuAgrovet 
         Caption         =   "Agrovet-Store"
         Begin VB.Menu mnustock 
            Caption         =   "Stock"
         End
         Begin VB.Menu mnusales 
            Caption         =   "Sales"
         End
         Begin VB.Menu mnusalesreturn 
            Caption         =   "Sales Return"
         End
         Begin VB.Menu mnuDrawnstock 
            Caption         =   "Drawnstock"
         End
      End
      Begin VB.Menu mnubonusste 
         Caption         =   "Farmers Default Deduction"
         Begin VB.Menu mnuDeductionSettings 
            Caption         =   "Monthly Deduction Settings"
         End
         Begin VB.Menu mnubprocess 
            Caption         =   "Monthly Deduction Processing"
         End
      End
      Begin VB.Menu mnustandingorders 
         Caption         =   "Standing Orders "
      End
   End
   Begin VB.Menu mnucashbook1 
      Caption         =   "Cash Book"
      Visible         =   0   'False
      Begin VB.Menu mnupayment 
         Caption         =   "Payment Requisition"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuapprovepayment 
         Caption         =   "Approve Payment Requisitions"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnutransactionslistings 
         Caption         =   "Transactions listings"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnucashpayments 
         Caption         =   "Cash Payments"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuvoidcheque 
         Caption         =   "Void Cheque"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuRecievePayment 
         Caption         =   "Receive Payment"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuspecialpurchacepayment 
         Caption         =   "Special purchace payments"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuinvoicepayment 
         Caption         =   "Invoice payment"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnucashReciepts 
         Caption         =   "Cash Reciepts"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnupettycash 
         Caption         =   "Petty Cash"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuAccountpayable 
      Caption         =   "Accounts Payable"
      Visible         =   0   'False
      Begin VB.Menu mnupurchase 
         Caption         =   "Purchase"
         Begin VB.Menu mnucreaterequisition 
            Caption         =   "Create requisition"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuapproverequisition 
            Caption         =   "Approve requisition"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuapprovedrequisition 
            Caption         =   "Approved requisition"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuraisepurchaseorder 
            Caption         =   "Raise Purchase Order"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuapprovepurchaseorders 
            Caption         =   "Approve Purchase Orders"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mnuinventory 
         Caption         =   "Inventory"
         Begin VB.Menu mnurecievegoods 
            Caption         =   "Recieve Goods"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuviewgoodsrecived 
            Caption         =   "View Goods Recieved"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuissueinventory 
            Caption         =   "Issue Inventory"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnupostinvetory 
            Caption         =   "Post Inventory Issue"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuinterstoretransfer 
            Caption         =   "Interstore Transfer"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuviewinterstorefransfer 
            Caption         =   "View Interstore Transfer"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnustocktaking 
            Caption         =   "Stock Taking"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mnuinvooices 
         Caption         =   "Invoices"
         Begin VB.Menu mnucreateinvoice 
            Caption         =   "Create Invoice"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnucreditdebitmemos 
            Caption         =   "Receive Invoice"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnurecieveutiltybills 
            Caption         =   "Recieve Utility Bills"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuviewsupplierinvoice 
            Caption         =   "View Supplier Invoices"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuliabilities 
            Caption         =   "Liabilities"
            Enabled         =   0   'False
         End
      End
   End
   Begin VB.Menu mnuAccountsRecievable 
      Caption         =   "Accounts Recievable"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnumilkhandling 
         Caption         =   "Milk Handling"
         Begin VB.Menu mnumilkcollection 
            Caption         =   "Milk Collection"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuviewmilkcollection 
            Caption         =   "View Milk Collection"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnurejectedmilk 
            Caption         =   "Rejected Milk"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuviewrejectedmilk 
            Caption         =   "View Rejected milk"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnumilkdipach 
            Caption         =   "Milk Dispach"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuviewmilkdispatches 
            Caption         =   "View Milk Dispatches"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnumilksales1 
            Caption         =   "Milk Sales"
            Enabled         =   0   'False
            Index           =   0
         End
         Begin VB.Menu mnuviewmilksales 
            Caption         =   "View Milk Sales"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuadjustmentcategories 
            Caption         =   "Adjustment Categories"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnufarmeradjustment 
            Caption         =   "Farmer Adjustments"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnumidmonthprocessing 
            Caption         =   "Mid Month processing"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnupostmidmonthpayment 
            Caption         =   "Post Mid month Payments"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnurollbackmidmonthpayment 
            Caption         =   "Roll Back Mid Month Payments"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuprintmidmonthschedule 
            Caption         =   "Print Mid Month Schedule"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuprocessmilkpayment 
            Caption         =   "Process Milk Payments"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnurollbackmilkpayments 
            Caption         =   "Roll Back Milk Payments"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnupostmilkpayments 
            Caption         =   "Post milk Payments"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mnufarmeradjustment2 
         Caption         =   "Farmers Adjustments"
         Enabled         =   0   'False
         Begin VB.Menu mnuaddDuductioncaategory 
            Caption         =   "Add Duduction Category"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuadjustmentscategories 
            Caption         =   "Adjustments categories"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnufarmersadjustments 
            Caption         =   "Farmers Aadjustments"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuviewfarmersadjustments 
            Caption         =   "View Farmers Adjustments"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuadvance 
            Caption         =   "Advance"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuviewadvance 
            Caption         =   "View Advance"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mnuinvoices 
         Caption         =   "Invoices"
         Begin VB.Menu mnumilksalesinvoices 
            Caption         =   "Milk sale Invoices"
         End
         Begin VB.Menu mnuviewmilksalesinvoices 
            Caption         =   "View Milk Sales Invoices"
         End
         Begin VB.Menu mnucustomersInvoices 
            Caption         =   "Customers Invoices"
         End
         Begin VB.Menu mnuviewcustomersinvoices 
            Caption         =   "View Customers Invoices"
         End
      End
      Begin VB.Menu mnucreditissue 
         Caption         =   "Credit Issue"
      End
      Begin VB.Menu mnuviewMemberIssue 
         Caption         =   "View Member Issue"
      End
      Begin VB.Menu mnuviewstaff 
         Caption         =   "View Staff"
      End
      Begin VB.Menu mnucashsales 
         Caption         =   "Cash Sales"
      End
      Begin VB.Menu mnuviewcashsales 
         Caption         =   "View Cash Sales"
      End
   End
   Begin VB.Menu mnuassets 
      Caption         =   "Fixed Assets "
      Begin VB.Menu mnucategories1 
         Caption         =   "Asset Categories"
      End
      Begin VB.Menu mnuassetregistration 
         Caption         =   "Asset Registration"
      End
      Begin VB.Menu mnuassetinquiry 
         Caption         =   "Asset Inquiry"
      End
      Begin VB.Menu mnudepreciation 
         Caption         =   "Depreciation And Valuation"
      End
      Begin VB.Menu mnuassetdisposal 
         Caption         =   "Asset Disposal"
      End
   End
   Begin VB.Menu mnuActivities 
      Caption         =   "Activities"
      Begin VB.Menu mnuMilkIntake 
         Caption         =   "Milk Intake"
      End
      Begin VB.Menu mnueasysacco 
         Caption         =   "Staff Management"
      End
      Begin VB.Menu mnucomplaindesk 
         Caption         =   "Complain Desk"
      End
      Begin VB.Menu mnubranchin 
         Caption         =   "Branch Intake"
      End
      Begin VB.Menu mnudispatchbmik 
         Caption         =   "Milk Dispatch Form"
      End
      Begin VB.Menu mnuSMS 
         Caption         =   "Send SMS"
      End
      Begin VB.Menu mnudebtorssale 
         Caption         =   "Debtors Registration and sales"
      End
      Begin VB.Menu mnuAccountspayable 
         Caption         =   "Accountspayable"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnubackup 
         Caption         =   "BACKUP DATABASE"
      End
      Begin VB.Menu mnupaybill 
         Caption         =   "B2C Payment Report"
      End
   End
   Begin VB.Menu mnuaccounts 
      Caption         =   "Accounts"
      Begin VB.Menu mnuGLSetup 
         Caption         =   "GL Set up"
      End
      Begin VB.Menu mnuChartsofaccounts 
         Caption         =   "Charts of accounts"
      End
      Begin VB.Menu mnuaccountsclassifed 
         Caption         =   "Accounts Classified"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnubookings 
         Caption         =   "Bookings"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu Gtransactions 
         Caption         =   "GL Transaction"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnufundsources 
         Caption         =   "Fund sources"
      End
      Begin VB.Menu mnunominal 
         Caption         =   "Receipt/payments"
      End
      Begin VB.Menu mnujournals 
         Caption         =   "Journal postings"
      End
      Begin VB.Menu mnuPostings 
         Caption         =   "Non Member Transactions"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnubudgetting 
         Caption         =   "Budgettings"
      End
      Begin VB.Menu mnuglinquiry 
         Caption         =   "GL Inquiry"
      End
      Begin VB.Menu mnumanagementreports 
         Caption         =   "Management Reports"
      End
      Begin VB.Menu mnujournaltypes 
         Caption         =   "Journal Types"
      End
      Begin VB.Menu mnubankrecon 
         Caption         =   "Bank Recon"
      End
      Begin VB.Menu mnucashbook 
         Caption         =   "Cash Book"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuglposting 
         Caption         =   "GL Posting"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuviewfarmerpayment 
         Caption         =   "GL Farmers Postings"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuperiods 
         Caption         =   "periods"
      End
      Begin VB.Menu mnuposting 
         Caption         =   "Staff Mpesa Posting"
      End
   End
   Begin VB.Menu mnuSetUp 
      Caption         =   "Set up"
      Begin VB.Menu mnuPricing 
         Caption         =   "Pricing"
         Begin VB.Menu mnuBPrice 
            Caption         =   "Buying Price"
         End
         Begin VB.Menu mnuSPrice 
            Caption         =   "Sms Setting"
         End
      End
      Begin VB.Menu mnuDeductions 
         Caption         =   "Deductions"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuMilkTestsParam 
         Caption         =   "Milk Tests Param"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuParameters 
         Caption         =   "Parameters"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSecurity 
         Caption         =   "Security"
         Begin VB.Menu mnuUserGroups 
            Caption         =   "User Groups"
         End
         Begin VB.Menu mnuUsers 
            Caption         =   "Users"
         End
         Begin VB.Menu mnuusermenus 
            Caption         =   "User Menus"
         End
         Begin VB.Menu mnuuserprevilleges 
            Caption         =   "User Previlleges"
         End
      End
      Begin VB.Menu mnucategories 
         Caption         =   "Categories"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuBanks 
         Caption         =   "Banks"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuAccountSetup 
         Caption         =   "Account Setup"
         Begin VB.Menu mnuAccountHeaders 
            Caption         =   "Account Headers"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuMainAccounts 
            Caption         =   "Main Accounts"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuaccountperiod 
            Caption         =   "Account Period"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mnuaddvehicle 
         Caption         =   "Add Vehicle"
      End
      Begin VB.Menu mnulocation 
         Caption         =   "Location"
      End
      Begin VB.Menu mnudistrict 
         Caption         =   "District"
      End
      Begin VB.Menu mnuBranch 
         Caption         =   "Branch"
      End
      Begin VB.Menu mnuProcessPayroll 
         Caption         =   "Process Payroll"
      End
      Begin VB.Menu mnuLoanSettings 
         Caption         =   "Loan Settings"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMaximumShares 
         Caption         =   "Maximum Shares"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnureportpath 
         Caption         =   "Report Path"
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "Reports"
      Begin VB.Menu mnupayrolls 
         Caption         =   "Payrolls Reports"
         Begin VB.Menu mnusuppliespayroll 
            Caption         =   "Suppliers Payroll"
         End
         Begin VB.Menu mnuTPayroll 
            Caption         =   "Transporters Payroll"
         End
         Begin VB.Menu mnucarryf 
            Caption         =   "CarryForward Report"
         End
      End
      Begin VB.Menu mnusuppliersR 
         Caption         =   "Suppliers Reports"
         Begin VB.Menu mnuSuppliersReg 
            Caption         =   "Suppliers Register"
         End
         Begin VB.Menu mnusuppliersdeductions 
            Caption         =   "Suppliers deductions"
         End
         Begin VB.Menu mnuSuppliersStatement 
            Caption         =   "Supplier's Statement"
         End
         Begin VB.Menu mnuactivesup 
            Caption         =   "Active Suppliers"
         End
      End
      Begin VB.Menu mnumilkin 
         Caption         =   "Milk Intake Reports"
         Begin VB.Menu mnudailySummary 
            Caption         =   "Daily Summary"
         End
         Begin VB.Menu mnuMilkIntakeSummary 
            Caption         =   "Today's Intake Summary"
         End
         Begin VB.Menu mnuBranchintakeanalysis 
            Caption         =   "Branch intake analysis"
         End
      End
      Begin VB.Menu mnutransrop 
         Caption         =   "Transporters Reports"
         Begin VB.Menu mnuTransportersDailyintake 
            Caption         =   "Transporters Dailyintake"
         End
         Begin VB.Menu mnuTransDetailed 
            Caption         =   "Transporters Detailed report"
         End
         Begin VB.Menu mnuTransportersStatement 
            Caption         =   "Transporter's Statement"
         End
         Begin VB.Menu mnutransporterperiodicreport 
            Caption         =   "Transporter Periodic Report"
         End
         Begin VB.Menu mnutransdeducreport 
            Caption         =   "Transporters Deduction Report"
         End
      End
      Begin VB.Menu mnudebtor 
         Caption         =   "Debtors Reports"
         Begin VB.Menu mnudebtorslist 
            Caption         =   "Debtors List"
         End
         Begin VB.Menu mnudebsta 
            Caption         =   "Debtors Statement"
         End
         Begin VB.Menu mnukiarie 
            Caption         =   "Kiarie Report"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuvehicleexp 
            Caption         =   "Vehicle Expense"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuoutsalesr 
         Caption         =   "Outlet sales Reports"
         Visible         =   0   'False
         Begin VB.Menu mnuoutletm 
            Caption         =   "Outlet milk"
         End
         Begin VB.Menu mnuoutsale 
            Caption         =   "Outlet Sales"
         End
         Begin VB.Menu mnuoutletdis 
            Caption         =   "Outlet Dispatch"
         End
      End
      Begin VB.Menu mnupurchaseR 
         Caption         =   "Purchases and Sales Reports"
         Begin VB.Menu mnupurchaseRep 
            Caption         =   "Purchases"
         End
         Begin VB.Menu mnusalesR 
            Caption         =   "Sales"
         End
         Begin VB.Menu mnudairyi 
            Caption         =   "Dairy Income Statement"
         End
         Begin VB.Menu mnudailys 
            Caption         =   "Daily Sales"
         End
      End
      Begin VB.Menu mnuBranchm 
         Caption         =   "Branch Milk Reports"
         Begin VB.Menu mnuBranchr 
            Caption         =   "Branch"
         End
         Begin VB.Menu mnuvehiclere 
            Caption         =   "Vehicle"
         End
      End
      Begin VB.Menu mnuplantsales 
         Caption         =   "Plant Sales & Boilings reports"
         Visible         =   0   'False
         Begin VB.Menu mnucastreg 
            Caption         =   "Customers List"
         End
         Begin VB.Menu mnuplansare 
            Caption         =   "Plant Sales"
         End
         Begin VB.Menu mnumsalesre 
            Caption         =   "Monthly Sales"
         End
         Begin VB.Menu mnuindividualre 
            Caption         =   "Individual Report"
         End
         Begin VB.Menu mnugemijare 
            Caption         =   "Gemija Report"
         End
      End
      Begin VB.Menu mnuAgrovetRepo 
         Caption         =   "Agrovet Reports"
         Begin VB.Menu mnucashr 
            Caption         =   "Cash"
         End
         Begin VB.Menu mnumpesar 
            Caption         =   "Mpesa"
         End
         Begin VB.Menu mnucheckoff 
            Caption         =   "Check Off"
         End
         Begin VB.Menu mnustaffre 
            Caption         =   "Staff"
         End
         Begin VB.Menu mnuallsales 
            Caption         =   "All Sales"
         End
         Begin VB.Menu mnustockBal 
            Caption         =   "Stock Balance"
         End
         Begin VB.Menu mnucashstaffre 
            Caption         =   "Cash Staff "
         End
         Begin VB.Menu mnusalesanare 
            Caption         =   "Sales Analysis"
         End
         Begin VB.Menu mnudispatchre 
            Caption         =   "Dispatch"
         End
         Begin VB.Menu mnuMonthlysalpre 
            Caption         =   "Monthly Sales Per Farmer"
            Visible         =   0   'False
         End
         Begin VB.Menu mnustockre 
            Caption         =   "Stock Receive"
         End
      End
      Begin VB.Menu mnuSpecificDed 
         Caption         =   "Specific Deduction Report"
      End
      Begin VB.Menu mnuGLReports 
         Caption         =   "GL Reports"
         Begin VB.Menu mnuGLTransactions 
            Caption         =   "GL Transactions"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuchartofaccount 
            Caption         =   "Chart Of Accounts"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuviewincomestatement 
            Caption         =   "Income Statement"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mnuaudittrail 
         Caption         =   "Audit Trail- CB/Postings"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuUsersSummary 
         Caption         =   "Users Summary"
      End
      Begin VB.Menu mnuIntakeAudit 
         Caption         =   "Intake Audit"
      End
      Begin VB.Menu mnutrendanalysis 
         Caption         =   "Trend Analysis"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ddate As Integer
Dim deddate As Integer
Dim tdate As Date
Dim csmsdate As Integer
Dim hhhh As Integer

Private Sub Check1_Click()
If Check1.value = 1 Then
Frame1.Visible = True
Else
Frame1.Visible = False
End If
End Sub

Private Sub Gtransactions_Click()
frmglinquiry.Show vbModal
End Sub

Private Sub MDIForm_Activate()
'Dim rmenu As String
'sql = ""
'sql = "SELECT   menu  FROM  tbl_usermenus where UserLoginIDs='" & User & "' ORDER BY Menu"
'Set rs = oSaccoMaster.GetRecordset(sql)
'If Not rs.EOF Then
'While Not rs.EOF
' rmenu = rs.Fields(0)
' 'rmenu = "mnuFiles"
'MainForm.Controls(rmenu).Enabled = True
'''MainForm.mnuFiles.Enabled = True
'
'rs.MoveNext
'Wend
'End If
End Sub

Private Sub MDIForm_Load()

'Dim ddate As Integer
'Dim deddate As Integer
'Dim tdate As Date
On Error Resume Next
Dim myclass As cdbase
Set myclass = New cdbase
Set cn = CreateObject("adodb.connection")
Provider = myclass.OpenCon
cn.Open Provider
Set rs = CreateObject("adodb.recordset")
sql = "select * from d_company"
rs.Open sql, cn
If Not rs.EOF Then
With StatusBar2(0).Panels
    .Item(1).Text = "USER : " & username
    .Item(2).Text = "DATE : " & Format(Get_Server_Date, "dd/mm/yyyy")
    .Item(3).Text = "COMPANY NAME:" & rs![name]
    .Item(4).Text = "TOWN:" & rs!town
    .Item(5).Text = "ADRESS:" & rs!Adress
    .Item(6).Text = "Tell:" & rs!PhoneNo
End With
If Not IsNull(rs![name]) Then cname = rs![name]
If Not IsNull(rs!Adress) Then paddress = rs!Adress
If Not IsNull(rs!town) Then town = rs!town
If Not IsNull(rs!motto) Then motto = rs!motto
If Not IsNull(rs!Email) Then Email = rs!Email
If Not IsNull(rs!SMSNo) Then CPhone = rs!SMSNo
If Not IsNull(rs!PhoneNo) Then Phone = rs!PhoneNo
If Not IsNull(rs!ddate) Then ddate = rs!ddate
If Not IsNull(rs!deddate) Then deddate = rs!deddate
If Not IsNull(rs!csmsdate) Then csmsdate = rs!csmsdate
If Not IsNull(rs!server) Then sserver = rs!server
'sserver
'csmsdate
End If

tdate = Format(Get_Server_Date, "dd/mm/yyyy")
dismenu

'oSaccoMaster.ExecuteThis (sql)
'oSaccoMaster.ExecuteThis ("update LOGINS set LogedOut=No where LogedOut = 'Yes' and UserLoginIDs='" & User & "'")

'If UserLoginIDs <> "" Then
'sql = ""
'sql = "select * from deduction where sno='" & Cells(i, 1) & "'"
'sql = "update LOGINS set LogedOut='No' where LogedOut = 'yes' and UserLoginIDs='" & User & "'"
'cnb.Execute (sql)
 'End If




Dim rmenu As String
'Dim rmenu As String
Dim x1 As String
Dim Rs1 As New ADODB.Recordset

sql = "select alias from tbl_menus order by id"
Set rs = oSaccoMaster.GetRecordset(sql)
While Not rs.EOF
x1 = rs.Fields(0)
'mnuenquiry
Dim x2 As String
sql = "select enable from tbl_usermenus where UserLoginID='" & User & "' and [menu]='" & x1 & "'"
Set Rs1 = oSaccoMaster.GetRecordset(sql)
If Not Rs1.EOF Then
'x2 = Rs1.Fields(0)

MainForm.Controls(x1).Enabled = True
Else
MainForm.Controls(x1).Enabled = False
End If

rs.MoveNext
Wend
'End If
'//call the class for the menus to be enables
'sql = ""
'sql = "SELECT   menu  FROM  tbl_usermenus where UserLoginIDs='" & User & "' ORDER BY Menu"
'Set rs = oSaccoMaster.GetRecordset(sql)
'If Not rs.EOF Then
'While Not rs.EOF
'
' rmenu = rs.Fields(0)
' DoEvents
'MainForm.Controls(rmenu).Enabled = True
''MainForm.mnuFiles.Enabled = True
'
'rs.MoveNext
'Wend
mnuFiles.Enabled = True
mnutransactions.Enabled = True
mnucashbook1.Enabled = True
mnuAccountpayable.Enabled = True
mnuAccountsRecievable.Enabled = True
mnuassets.Enabled = True
mnuActivities.Enabled = True
mnuaccounts.Enabled = True
mnuSetUp.Enabled = True
mnuReports.Enabled = True

'End If
'details
DTPPeriod = Format(Get_Server_Date, "mmm/yyyy")
Frame1.Visible = False
'details
End Sub
Public Sub details()
MainForm.Caption = "EasyMa " & "--(Milk Intake Form)"
DTPPeriod = Date
Startdate = DateSerial(Year(DTPPeriod), month(DTPPeriod), 1)
Enddate = DateSerial(Year(DTPPeriod), month(DTPPeriod) + 1, 1 - 1)
Dim rsst, rss, rst, rsg, rs, rsh As Recordset
'''''todays no of suppliers
sql = ""
sql = "set dateformat DMY select count(distinct(SNo)) from  d_Milkintake where TransDate='" & DTPPeriod & "' "
Set rsst = New ADODB.Recordset
Set rsst = oSaccoMaster.GetRecordset(sql)
If Not rsst.EOF Then
 Label8 = rsst.Fields(0)
Else
 Label8 = "0"
End If
''''active suppliers month
 sql = ""
 sql = "set dateformat DMY select count(distinct(SNo)) from  d_Milkintake where TransDate>='" & Startdate & "' and TransDate<='" & Enddate & "' "
 Set rss = New ADODB.Recordset
 Set rss = oSaccoMaster.GetRecordset(sql)
 If Not rss.EOF Then
 Label11 = rss.Fields(0)
 Else
 Label11 = "0"
 End If
 '''''todays kgs
sql = ""
sql = "set dateformat DMY select isnull(sum(QSupplied),0) from  d_Milkintake where TransDate='" & DTPPeriod & "' "
Set rst = New ADODB.Recordset
Set rst = oSaccoMaster.GetRecordset(sql)
If Not rst.EOF Then
 Label12 = rst.Fields(0)
End If
 '''Month kgs
 sql = ""
 sql = "set dateformat DMY select isnull(sum(QSupplied),0) from  d_Milkintake where TransDate>='" & Startdate & "' and TransDate<='" & Enddate & "' "
 Set rsg = New ADODB.Recordset
 Set rsg = oSaccoMaster.GetRecordset(sql)
 If Not rsg.EOF Then
 Label13 = rsg.Fields(0)
 End If
 '''''''female count
sql = ""
'sql = "set dateformat dmy SELECT d.DCode, d.DName, m.DispQnty,m.DispDate FROM  d_Debtors AS d INNER JOIN d_MilkControl AS m ON d.DCode = m.dcode WHERE     (DispDate = '" & DTPDispatchDate & "') and status=0"
sql = "set dateformat dmy SELECT isnull(count(distinct(d.SNo)),0) FROM  d_Suppliers AS d INNER JOIN d_Milkintake AS m ON d.SNo = m.SNo WHERE TransDate>='" & Startdate & "' and TransDate<='" & Enddate & "' and Type='FEMALE'"
 Set rs = New ADODB.Recordset
 Set rs = oSaccoMaster.GetRecordset(sql)
 If Not rs.EOF Then
  Label9 = rs.Fields(0)
 End If
 ''''''''male count
 sql = ""
 sql = "set dateformat dmy SELECT isnull(count(distinct(d.SNo)),0) FROM  d_Suppliers AS d INNER JOIN d_Milkintake AS m ON d.SNo = m.SNo WHERE TransDate>='" & Startdate & "' and TransDate<='" & Enddate & "' and Type='MALE'"
 Set rsh = New ADODB.Recordset
 Set rsh = oSaccoMaster.GetRecordset(sql)
 If Not rsh.EOF Then
  Label10 = rsh.Fields(0)
 End If
'''''end of all
End Sub
Private Sub dismenu()
'On Error Resume Next
Dim I As Control
Dim intIncrement As Integer

For Each I In Controls
If TypeOf I Is Menu Then
I.Enabled = False
End If
Next
'

'
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub

Private Sub mnuAccountHeaders_Click()
frmHeaders.Show vbModal
End Sub

Private Sub mnuaccountperiod_Click()
frmPeriods.Show vbModal
End Sub

Private Sub mnuaccountsclassifed_Click()
frmaccountclassified.Show vbModal
End Sub

Private Sub mnuAccountspayable_Click()
frmaccountspayable.Show vbModal
End Sub

Private Sub mnuactivesup_Click()
 reportname = "ACTIVE SUPPLIERS.rpt"
 Show_Sales_Crystal_Report "", reportname, ""
End Sub

Private Sub mnuaddDuductioncaategory_Click()
frmAddfarmersadjustmentscategory.Show vbModal
End Sub

Private Sub mnuaddvehicle_Click()
frmVehicleReg.Show vbModal
End Sub

Private Sub mnuadjustmentcategories_Click()
frmfarmeradjustmentcategory.Show vbModal
End Sub
Private Sub mnuadjustmentscategories_Click()
frmfarmersadjustmentscategorylisting.Show vbModal
End Sub
Private Sub mnuadvance_Click()
frmmembersadvancepayment.Show vbModal
End Sub

Private Sub mnuagrcashsales_Click()
    reportname = "CASH agrovet sales.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub mnuAgro_Click()

End Sub

Private Sub mnuagroversuppliers_Click()
frmSupplier.Show vbModal
End Sub

Private Sub mnuai_Click()
frmAI.Show vbModal, Me
End Sub

Private Sub mnuallsalea_Click()
    reportname = "all agrovet sales.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub mnuannualinnJJJJJ_Click()
reportname = "MilkintakeMonthly.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub mnuapprove_Click()
'check the user
'sql = "SELECT     UserLoginIDs,levels, UserGroup, SUPERUSER From UserAccounts where UserLoginIDs='" & User & "'"
'Set rs = oSaccoMaster.GetRecordset(sql)
'If Not rs.EOF Then
'If rs!usergroup <> "Manager" Then
'MsgBox "Manager only Allowed", vbInformation
'Exit Sub
'End If
'End If
'
'
'Frmapproval.Show vbModal, Me
End Sub

Private Sub mnuallsales_Click()
    reportname = "all agrovet sales.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub mnuapprovedrequisition_Click()
frmapprovedrequisitions.Show vbModal
End Sub
Private Sub mnuapprovepayment_Click()
frmPayApproval.Show vbModal
End Sub
Private Sub mnuapprovepurchaseorders_Click()
frmpurchaselist.Show vbModal
End Sub
Private Sub mnuapproverequisition_Click()
frmnewrequisitions.Show vbModal
End Sub

Private Sub mnuargive_Click()
frmargive.Show vbModal, Me
End Sub

Private Sub mnuassetdisposal_Click()
frmassetsdisposals.Show , Me
End Sub

Private Sub mnuassetinquiry_Click()
''frmassetsinquiry.Show vbModal, Me
frmassetsinquiry.Show vbModal, Me
End Sub

Private Sub mnuassetregistration_Click()
frmAssetMaster.Show vbModal, Me
End Sub

Private Sub mnuaudittrail_Click()
'//audittrailcbpostings
reportname = "audittrailcbpostings.rpt"
Show_Sales_Crystal_Report "", reportname, ""

End Sub

Private Sub mnubackup_Click()
frmSQLSRVBackup.Show vbModal
End Sub

Private Sub mnuBankPayments_Click()
frmPayBanks.Show vbModal
End Sub

Private Sub mnubankrecon_Click()
frmBankRec.Show vbModal
End Sub

Private Sub mnuBanks_Click()
frmBankSetup.Show vbModal
End Sub

Private Sub mnuBonustatement_Click()
reportname = "Bonus statement.rpt"
 
 Show_Sales_Crystal_Report "", reportname, ""
End Sub

Private Sub mnuBonusstatement_Click()
reportname = "Bonus statement.rpt"
 
 Show_Sales_Crystal_Report "", reportname, ""
End Sub

Private Sub mnubookings_Click()
frmentrypostings.Show vbModal
End Sub

Private Sub mnuBPrice_Click()
sql = "SELECT     UserLoginIDs,levels, UserGroup, SUPERUSER From UserAccounts where UserLoginIDs='" & User & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
'If rs!Levels <> "Manager" Then
'MsgBox "Access Denied", vbInformation
'Exit Sub
'End If
End If
frmPricing.Show vbModal
End Sub

Private Sub mnubprocess_Click()
frmbonusprocess.Show vbModal
End Sub

Private Sub mnuBranch_Click()
'frmBranch.Show vbModal
frmBranch.Show vbModal
End Sub

Private Sub mnubranchin_Click()
frmbranchmilk.Show vbModal
End Sub

Private Sub mnuBranchintakeanalysis_Click()
reportname = "milkintake analysis.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub mnubranchs_Click()
reportname = "Branch sales.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub mnubranchm_Click()
'frmbranchmilk.Show vbModal
End Sub

Private Sub mnuBranchr_Click()
Dim ans As String
ans = MsgBox("Do you Want Report per Branch ?", vbYesNo)
If ans = vbYes Then
 reportname = "d_BranchInvoice.rpt"
Else
reportname = "d_BranchInvoice2.rpt"
End If
Show_Sales_Crystal_Report "", reportname, ""
End Sub

Private Sub mnubudgetting_Click()
frmbudgetting.Show vbModal
End Sub
Private Sub mnucashaccounts_Click()
frmcashaccounts.Show vbModal
End Sub

Private Sub mnuCanregistration_Click()
frmContainer.Show vbModal
End Sub

Private Sub mnuCarryforward_Click()
reportname = "carryforward.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub mnucarryf_Click()
reportname = "carryforward.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub mnucashbook_Click()
frmCashBookTransaction.Show vbModal
End Sub

Private Sub mnucashpay_Click()
frmcashpay.Show vbModal
End Sub

Private Sub mnucashpayments_Click()
frmmainpaymentaccount.Show vbModal
End Sub

Private Sub mnucashr_Click()
    reportname = "CASH agrovet sales.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub mnucashReciepts_Click()
frmcashreciept.Show vbModal
End Sub

Private Sub mnucashstaffre_Click()
    reportname = "Agrovet staffc.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub mnucastreg_Click()
reportname = "Plan list Report.rpt"
Show_Sales_Crystal_Report "", reportname, ""
End Sub

Private Sub mnucategories1_Click()
frmAssets.Show vbModal, Me
End Sub

Private Sub mnuchartofaccount_Click()
reportname = "chartsofaccounts.rpt"
 
 Show_Sales_Crystal_Report "", reportname, ""
End Sub

Private Sub mnuChartsofaccounts_Click()
frmchartsofaccounts.Show vbModal
End Sub

Private Sub mnucheckoffs_Click()
    reportname = "CHECK OFF agrovet sales.rpt"
    Show_Sales_Crystal_Report "", reportname, ""
End Sub

Private Sub mnucheckoff_Click()
    reportname = "CHECK OFF agrovet sales.rpt"
    Show_Sales_Crystal_Report "", reportname, ""
End Sub

Private Sub mnucomplaindesk_Click()
' sql = "SELECT     UserLoginIDs, UserGroup,levels SUPERUSER From UserAccounts where UserLoginIDs='" & User & "'"
'    Set rs = oSaccoMaster.GetRecordset(sql)
'    If Not rs.EOF Then
'    If rs!Levels <> "Intake" Then
'    MsgBox "You are not allowed to Use complain Module ", vbInformation
'    Exit Sub
'    End If
'    End If
frmcomplaintdesk.Show vbModal
End Sub

Private Sub mnuContainerType_Click()
frmContainer.Show vbModal
End Sub

Private Sub mnuControlReport_Click()
reportname = "d_ControlReport.rpt"
 
 Show_Sales_Crystal_Report "", reportname, ""
End Sub

Private Sub mnucostcentre_Click()
frmCostcentre.Show vbModal
End Sub

Private Sub mnucreateinvoice_Click()
  frminvoice.Show vbModal
End Sub

Private Sub mnucreaterequisition_Click()
frmcreaterequisition.Show vbModal
End Sub
Private Sub mnucreditdebitmemos_Click()
'frmvendorscreditdebitmemos.Show vbModal
frmreceiveinvoice.Show vbModal
End Sub

Private Sub mnuCurrentstock_Click()
reportname = "stockbalances as at.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub mnudailys_Click()
Dim ans As String
ans = MsgBox("Do you Want a Report as per price??", vbYesNo)
If ans = vbYes Then
reportname = "SALES PER DAY.rpt"
Else
reportname = "dailysales.rpt"
End If
Show_Sales_Crystal_Report "", reportname, ""
End Sub

Private Sub mnudailySummary_Click()

Dim ans As String
ans = MsgBox("Do you Want List of Suppliers from All Branches?", vbYesNo)
If ans = vbYes Then
 reportname = "d_Dailysummary2.rpt"

 Else
 reportname = "d_Dailysummary.rpt"
'  ReportTitle = "TO :" & UCase(Location) & " ; " & vbNewLine & " Please pay the following transporters the amount indicated: (Our Ref is code)"
'    '{d_Payroll.NPay} > 0 and {d_Payroll.Bank} <> '' and month({d_Payroll.EndofPeriod})= month(30/09/2010)  AND year({d_Payroll.EndofPeriod}) = Year(30/09/2010)
' STRFORMULA = "{d_TransportersPayRoll.NetPay} > 0 and {d_TransportersPayRoll.BankName} = '" & cboBank & "' and month({d_TransportersPayRoll.EndPeriod})=" & month(dtpEndPeriod) & " AND year({d_TransportersPayRoll.EndPeriod}) =" & year(dtpEndPeriod)
' Show_Sales_Crystal_Report STRFORMULA, reportname, ReportTitle
 End If
 Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub mnudebtormilk_Click()
'frmMilkControl.Show vbModal
End Sub

Private Sub mnuDebtorsDetails_Click()
frmDebtorsDetails.Show vbModal
End Sub

Private Sub mnuDebtorsReport_Click()
reportname = "d_DebtorsInvoice.rpt"
 
 Show_Sales_Crystal_Report "", reportname, ""
End Sub

Private Sub mnudairyi_Click()
On Error GoTo ErrorHandler
Dim rst, rstg, rsa As Recordset
Dim txtdateenterered As Date
txtdateenterered = Date
Startdate = DateSerial(Year(txtdateenterered), month(txtdateenterered), 1)
Enddate = DateSerial(Year(txtdateenterered), month(txtdateenterered) + 1, 1 - 1)
sql = ""
sql = "set dateformat dmy delete from d_incomestate where Date >= '" & Startdate & "' And Date<='" & Enddate & "'"
cn.Execute sql
     
     sql = ""
     sql = "set dateformat dmy Select distinct(TransDate) from   d_Milkintake WHERE TransDate >='" & Startdate & "' And TransDate<='" & Enddate & "' order by TransDate asc"
     Set rstg = cn.Execute(sql)
     While Not rstg.EOF
      sql = ""
      sql = "set dateformat dmy Select isnull(sum(PAmount),0) from   d_Milkintake WHERE TransDate ='" & rstg.Fields(0) & "'"
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

Private Sub mnudairyprocessi_Click()
    sql = ""
    sql = "SELECT UserLoginIDs, UserGroup, SUPERUSER From UserAccounts where UserLoginIDs='" & User & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
 If Not rs.EOF Then
  If rs!SuperUser <> "1" Then
    MsgBox "You are not allowed to Process", vbInformation
    Exit Sub
  End If
 End If
'frmdebtor1.Show vbModal
End Sub

Private Sub mnudebsta_Click()
    reportname = "d_DebtorsInvoice.rpt"
    Show_Sales_Crystal_Report "", reportname, ""
End Sub

Private Sub mnudebtorslist_Click()
    reportname = "debtorslist.rpt"
    Show_Sales_Crystal_Report "", reportname, ""
End Sub

Private Sub mnudebtorssale_Click()
'check the user
    sql = "SELECT     UserLoginIDs, UserGroup, SUPERUSER From UserAccounts where UserLoginIDs='" & User & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
    If rs!SuperUser <> "1" Then
    MsgBox "You are not allowed to Reverse Sales", vbInformation
    Exit Sub
    End If
    End If
frmdebtor1.Show vbModal
End Sub

Private Sub mnuDeductionassignment_Click()
frmFarmerDeductionAssign.Show
End Sub

Private Sub mnuDeductions_Click()
frmDCodes.Show vbModal
End Sub

Private Sub mnuDeductionSettings_Click()
frmPresETS.Show vbModal
End Sub

Private Sub mnudepreciation_Click()
frmassetstransactions.Show vbModal
End Sub

Private Sub mnudispatch_Click()
   reportname = "d_controlreport2.rpt"
   Show_Sales_Crystal_Report "", reportname, ""
End Sub

Private Sub mnudispatch1_Click()
reportname = "Dispatch.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub mnudispatchstock_Click()
reportname = "Transfer agrovet sales1.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub mnuDistricts_Click()
frmDistricts.Show vbModal
End Sub

Private Sub mnudistrict_Click()
frmDistricts.Show vbModal
End Sub

Private Sub mnuDrawnstock_Click()
'check the user
    sql = "SELECT     UserLoginIDs, UserGroup, SUPERUSER From UserAccounts where UserLoginIDs='" & User & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
    If rs!SuperUser <> "1" And rs!SuperUser <> "2" Then
    MsgBox "You are not allowed to Draw stock", vbInformation
    Exit Sub
    End If
    End If
frmdrawnstock.Show vbModal
End Sub

Private Sub mnudispatchbmik_Click()
frmbranchdisp.Show vbModal
End Sub

Private Sub mnudispatchre_Click()
   reportname = "Transfer agrovet sales1.rpt"
' End If
  Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub mnueasysacco_Click()
'Shell "D:\HRSource\HR1.EXE", vbMinimizedFocus
'Dim h
'Shell "E:\SOURCE CODES\EASYSACCO 7.14Test.exe", vbMaximizedFocus
'frmPettyCash.Show vbModal
End Sub

Private Sub mnuEnquiry_Click()
frmEnquery.Show vbModal
End Sub

Private Sub mnuexit_Click()
End
End Sub

Private Sub mnufarmeradjustment_Click()
frmfarmeradjustment.Show vbModal
End Sub

Private Sub mnufarmersadjustments_Click()
frmfarmeradjustment.Show vbModal
End Sub

Private Sub mnuFarmersPayment_Click()
frmPayment.Show vbModal
End Sub

Private Sub mnuFsaaccountinquiry_Click()
'frminquiry.Show vbModal
Shell "\\Main-server\shared\OLE E-FOSA.exe", vbMaximizedFocus
End Sub

Private Sub mnufundsources_Click()
frmBankSetup.Show vbModal
End Sub

Private Sub mnuGeneralshares_Click()
'reportname = "general shares.rpt"
reportname = "Totalshares.rpt"
 
 Show_Sales_Crystal_Report "", reportname, ""
End Sub

Private Sub mnugemijare_Click()
reportname = "MonthlyDebtors Report1.rpt"
Show_Sales_Crystal_Report "", reportname, ""
End Sub

Private Sub mnuglinquiry_Click()
GlinqueryTransaction.Show vbModal
End Sub

Private Sub mnuglposting_Click()
frmglpostings.Show vbModal
End Sub

Private Sub mnuGLSetup_Click()
frmglsetup.Show vbModal
End Sub

Private Sub mnuGLTransactions_Click()
frmreversalofcashbookentries.Show vbModal
End Sub

Private Sub mnuImportExport_Click()
On Error GoTo SysError
Shell App.path & "/DataImport.exe"
'Shell "shutdown.exe -s -f -t 10"
Exit Sub
SysError:
MsgBox "Cannot find export feature", vbCritical
End Sub

Private Sub mnuindividualShares_Click()
frmLedgerFees.Show vbModal
End Sub

Private Sub mnuindividualre_Click()
reportname = "Planindividual.rpt"
Show_Sales_Crystal_Report "", reportname, ""
End Sub

Private Sub mnuIntakeAudit_Click()
 reportname = "d_MilkAfter4.rpt"
  Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub mnuinterstoretransfer_Click()
frminterstoretransfer.Show vbModal
End Sub

Private Sub mnuinvoicepayment_Click()
frminvoicepayment.Show vbModal
End Sub

Private Sub mnuinvoicestools_Click()
frmreceiptadjustment.Show vbModal
End Sub

Private Sub mnuiventory_Click()
End Sub

Private Sub mnuissueinventory_Click()
frmissueinventory.Show vbModal
End Sub

Private Sub mnuItemsSold_Click()
reportname = "d_sold.rpt"
  Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub mnujournals_Click()
frmJournals.Show vbModal
End Sub

Private Sub mnujournaltypes_Click()
frmJournalTypes.Show vbModal
End Sub

Private Sub mnukiarie_Click()
    reportname = "Kiarie reports.rpt"
    Show_Sales_Crystal_Report "", reportname, ""
End Sub

Private Sub mnuliabilities_Click()
frmpendingliabilities.Show vbModal
End Sub

Private Sub mnuLoanSettings_Click()
frmLoanSet.Show vbModal, Me
End Sub

Private Sub mnuLocations_Click()
frmLocation.Show vbModal
End Sub

Private Sub mnulocation_Click()
frmLocation.Show vbModal
End Sub

Private Sub mnuLogOff_Click()
Me.Hide
frmlogin.Show vbModal
End Sub

Private Sub mnulogoutt_Click()
frmlogout.Show vbModal
End Sub

Private Sub mnuMainAccounts_Click()
frmMainAccounts.Show vbModal
End Sub

Private Sub mnumanagementreports_Click()
frmAccounts.Show vbModal
'frmAccounts1.Show vbModal
End Sub

Private Sub mnuMaximumShares_Click()
frmMaxShares.Show vbModal
End Sub

Private Sub mnumidmonthprocessing_Click()
frmmidmonthpaymentprocessing.Show vbModal
End Sub

Private Sub mnumilkcollection_Click()
frmMilkCollection.Show vbModeless
End Sub

Private Sub mnuMilkControl_Click()
'Dim ans As String
'ans = MsgBox("Please for Debtors Reply with Yes and No to receive from Branches?", vbYesNo)
'If ans = vbYes Then
'frmMilkControl.Show vbModal
'Else
'frmbranchmilk.Show vbModal
'frmMilkControl.Show vbModal
'End If
End Sub

Private Sub mnumilkdipach_Click()
frmmilkdispatch.Show vbModal
End Sub

Private Sub mnuMilkIntake_Click()

MainForm.Caption = "EasyMa " & "--(Milk Intake Form)"
frmMilkCollection.Show vbModeless
End Sub

Private Sub mnuMilkIntakeSummary_Click()
'ReportTitle = "ALA"
  reportname = "d_DailyIntake.rpt"
 Show_Sales_Crystal_Report "", reportname, ""
End Sub

Private Sub mnumilkquality_Click()
Frmquality.Show vbModal
End Sub

Private Sub mnumilksales1_Click(Index As Integer)
frmmilksales.Show vbModal
End Sub

Private Sub mnuMilkTest_Click()
    frmMilkTests.Show vbModal
End Sub

Private Sub mnumilktestsetup_Click()
Frmqualitysetup.Show vbModal
End Sub

Private Sub mnuMilkTestsParam_Click()
    frmMilkTestSettings.Show vbModal
End Sub

Private Sub mnunewcashaccount_Click()
    frmnewcashaccount.Show vbModal
End Sub

Private Sub mnunewtender_Click()
    frmnewtender.Show vbModal
End Sub

Private Sub mnumpes_Click()
 reportname = "Mpesa agrovet sales.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub mnuMonthlysalpre_Click()
Dim ans As String
ans = MsgBox("Do you Want Summary Report as per Outlet?", vbYesNo)
 If ans = vbYes Then
     reportname = "COMBINESALES1.rpt"
 Else
   reportname = "COMBINESALES.rpt"
 End If
' Show_Sales_Crystal_Report STRFORMULA, reportname, ""
    Show_Sales_Crystal_Report "", reportname, ""
End Sub

Private Sub mnumpesar_Click()
    reportname = "Mpesa agrovet sales.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub mnumsalesre_Click()
reportname = "Plan sales Report1.rpt"
Show_Sales_Crystal_Report "", reportname, ""
End Sub

Private Sub mnunominal_Click()
frmNominals.Show vbModal
End Sub

Private Sub mnunonmembersuppliers_Click()
    frmcustomers.Show vbModal
End Sub

Private Sub mnuothers_Click()
reportname = "others.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub mnupackagegroup_Click()
'//TCHP_MemberList
reportname = "TCHP_MemberListin.rpt"
 
 Show_Sales_Crystal_Report STRFORMULA, reportname, ""

End Sub

Private Sub mnuoutletdis_Click()
reportname = "OutletVehicledispatch Report.rpt"
Show_Sales_Crystal_Report "", reportname, ""
End Sub

Private Sub mnuoutletm_Click()
reportname = "Outletdispatch Report.rpt"
Show_Sales_Crystal_Report "", reportname, ""
End Sub

Private Sub mnuoutsale_Click()
reportname = "Outlet Report.rpt"
Show_Sales_Crystal_Report "", reportname, ""
End Sub

Private Sub mnuParameters_Click()
frmSysParam.Show vbModal
End Sub

Private Sub mnupartlycashcheckoff_Click()
    reportname = "CheckOff_PartlyPayment.rpt"
    Show_Sales_Crystal_Report "", reportname, ""
End Sub

Private Sub mnupaybill_Click()
'reportname = "paybillpayement.rpt"
reportname = "B2CPAYMENTREPORT.rpt"
Show_Sales_Crystal_Report STRFORMULA, reportname, ""
'Print_reports reportname, "", CompanyName
End Sub

Private Sub mnupayment_Click()
frmpaymentrequisition.Show vbModal
End Sub

Private Sub mnupayroll_Click()
Dim h
Shell "C:\Maziwa\payroll.exe", vbMaximizedFocus

End Sub


Private Sub mnupaysum_Click()
reportname = "olepay.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub mnuperiods_Click()
frmPeriods.Show vbModal
End Sub

Private Sub mnupettycash_Click()
frmPettyCash.Show vbModal
End Sub

Private Sub mnuplansare_Click()
reportname = "Plan sales Report.rpt"
Show_Sales_Crystal_Report "", reportname, ""
End Sub

Private Sub mnuposting_Click()
   'BackColor = #&HC0FFC0##
    sql = "SELECT     UserLoginIDs, UserGroup, SUPERUSER From UserAccounts where UserLoginIDs='" & User & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
    If rs!SuperUser <> "1" Then
    MsgBox "You are not allowed to Post ", vbInformation
    Exit Sub
    End If
    End If

frmImportPhotos.Show vbModal
'frmImportOtherDeduction.Show vbModal
End Sub

Private Sub mnuPostings_Click()
frmNonmemberTransactionposting.Show vbModal
End Sub

Private Sub mnupostinvetory_Click()
frmpostinventoryissue.Show vbModal
End Sub

Private Sub mnupostmidmonthpayment_Click()
frmpostmidmonthpayment.Show vbModal
End Sub

Private Sub mnupostmilkpayments_Click()
frmpostmemberspayments.Show vbModal
End Sub



Private Sub mnupptran_Click()
reportname = "tran_PartlyPayment.rpt"
    Show_Sales_Crystal_Report "", reportname, ""
End Sub

Private Sub mnuprintmidmonthschedule_Click()
frmprintmidmonthschedule.Show vbModal
End Sub

Private Sub mnuprocessmilkpayment_Click()
frmmilkprocesspayment.Show vbModal
End Sub

Private Sub mnuProcessPayroll_Click()
frmProcess.Show vbModal
End Sub

Private Sub mnuQbmps_Click()
frmqbmps.Show vbModal
End Sub

Private Sub MNUQBMPSIMPORT_Click()
FRMQBMPSIMPORTS.Show vbModal
End Sub

Private Sub mnupurchaseRep_Click()
reportname = "MILK PURCHASE REPORT.rpt"
Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub mnuraisepurchaseorder_Click()
'frmpurchaseorders.Show vbModal
End Sub

Private Sub mnurecievegoods_Click()
frmitemreciepts.Show vbModal
End Sub

Private Sub mnuRecievePayment_Click()
frmcustomerpayment.Show vbModal
End Sub

Private Sub mnurecievesupplierinvoice_Click()
frmvendersinvoices.Show vbModal
End Sub

Private Sub mnurecieveutiltybills_Click()
frmrecievevendorbill.Show vbModal
End Sub

Private Sub mnuReconcilition_Click()
frmreconcilition.Show vbModal
End Sub

Private Sub mnurecivstock_Click()
    reportname = "stoock Report.rpt"
    Show_Sales_Crystal_Report "", reportname, ""
End Sub

Private Sub mnuregistertransporter_Click()
frmTransportersDetails.Show vbModal
End Sub

Private Sub mnuRegistration_Click()
'    Set rs = oSaccoMaster.GetRecordset(sql)
'    If Not rs.EOF Then
'    If rs!enable = 1 Then
'    MsgBox "You are not allowed to Register suppliers", vbInformation
'    Exit Sub
'    End If
'    End If
sql = "SELECT     UserLoginIDs,levels, UserGroup, SUPERUSER From UserAccounts where UserLoginIDs='" & User & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
If rs!SuperUser <> "1" Then
MsgBox "You are not allowed to Register suppliers", vbInformation
Exit Sub
End If
End If

frmSupplies.Show vbModal
End Sub

Private Sub mnurejectedmilk_Click()
'frmRejectedmilk.Show vbModal
End Sub

Private Sub mnurepackaging_Click()
frmproductrepackaging.Show vbModal
End Sub

Private Sub mnureportpath_Click()
frmReportPath.Show vbModal
End Sub

Private Sub mnurollbackmidmonthpayment_Click()
frmrollbackmidmonthpayment.Show vbModal
End Sub

Private Sub mnuSales_Click()

'MainForm.Caption = "EasyMa " & "--(Milk Intake Form)"

frmreceipts.Show vbModal
End Sub

Private Sub mnuSalesAnalysis_Click()
    reportname = "Sales analysis.rpt"
    Show_Sales_Crystal_Report "", reportname, ""
End Sub

Private Sub mnusalesreturn_Click()
'check the user
'    sql = "SELECT     UserLoginIDs, UserGroup, SUPERUSER From UserAccounts where UserLoginIDs='" & User & "'"
'    Set rs = oSaccoMaster.GetRecordset(sql)
'    If Not rs.EOF Then
'    If rs!SuperUser <> "1" Then
'    MsgBox "You are not allowed to Reverse Sales", vbInformation
'    Exit Sub
'    End If
'    End If

frmsalesreturn.Show vbModal
End Sub

Private Sub mnuSendSMS_Click()
frmSendSMS.Show vbModal
End Sub

Private Sub mnuShares_Click()
frmshares.Show vbModal, Me
End Sub

Private Sub mnusharestransactions_Click()
frmsharestransactions.Show vbModal, Me
'frmsharestransactions
End Sub

Private Sub mnusalesanare_Click()
    reportname = "Sales analysis.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub mnusalesR_Click()
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

Private Sub mnuSMS_Click()
frmSendSMS.Show vbModal
End Sub

Private Sub mnuspecialpurchacepayment_Click()
frmspecialpurchasepayment.Show vbModal, Me
End Sub

Private Sub mnuSpecificDed_Click()
reportname = "d_SpecificDeduct.rpt"
 
 Show_Sales_Crystal_Report "", reportname, ""
End Sub

Private Sub mnustaffcash_Click()
reportname = "agrovet staffc.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, ""

End Sub

Private Sub mnuSPrice_Click()
frmsmssettings.Show vbModal, Me
End Sub

Private Sub mnustaffre_Click()
    reportname = "Agrovet staffs.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub mnuStaffregistration_Click()
frmstaffregistration.Show vbModal, Me
End Sub

Private Sub mnuStaffsagrovet_Click()
reportname = "Agrovet staffs.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub mnustandingorders_Click()
'frmstandingorders.Show vbModal, Me
'frmPresETS.Show vbModal, Me
frmSTOS.Show vbModal, Me
End Sub

Private Sub mnustock_Click()
 'check the user
    sql = "SELECT     UserLoginIDs, UserGroup, SUPERUSER From UserAccounts where UserLoginIDs='" & User & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
    If rs!SuperUser <> "1" Then
    MsgBox "You are not allowed to Receive stock", vbInformation
    Exit Sub
    End If
    End If

'****************'
frmproduct1s.Show vbModal, Me
End Sub

Private Sub mnustockana_Click()
frmstockbalance.Show vbModal
End Sub

Private Sub mnustockba_Click()
    reportname = "d_StockBal.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub mnustockbal_Click()
Dim ans As String
frmstockbalance.Show vbModal
'ans = MsgBox("Do you Want List of all Branches", vbYesNo)
'If ans = vbYes Then
' reportname = "d_StockBal1.rpt"
' Else
' reportname = "d_StockBal.rpt"
' End If
'  Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub mnuStockBalance_Click()
reportname = "d_StockBal.rpt"
 
 Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub mnustockrunningbalance_Click()

 reportname = "runningbal.rpt"
 Show_Sales_Crystal_Report "", reportname, ""
 
End Sub

Private Sub mnustocktaking_Click()
frmstocktaking.Show vbModal, Me
End Sub

Private Sub mnuSubscribe_Click()
'frmSMSSubscribe.Show vbModal
End Sub

Private Sub mnusupap_Click()
Frmapproval.Show vbModal
End Sub

Private Sub mnusuppliersdeductions_Click()
 
 reportname = "d_suppliersdeductions.rpt"
 
 Show_Sales_Crystal_Report "", reportname, ""

  
End Sub

Private Sub mnuSuppliersReg_Click()

 
 reportname = "suppliersregister.rpt"

 Show_Sales_Crystal_Report "", reportname, ""

 

End Sub

Private Sub mnuSuppliersStatement_Click()
frmSupplierStmt.Show vbModal
End Sub

Private Sub mnusuppliespayroll_Click()
'//d_payroll\
'//call the companyname
Dim ans As String
ans = MsgBox("Do you Want supplier With Routes", vbYesNo)
If ans = vbYes Then
 reportname = "d_payroll.rpt"
 Else
 reportname = "d_payroll.rpt"
 End If
 
 Show_Sales_Crystal_Report STRFORMULA, reportname, ""
  
End Sub

Private Sub mnusychronizestock_Click()
frmsychronize.Show vbModal
End Sub

Private Sub mnutender_Click()
frmtenders.Show vbModal
End Sub


Private Sub mnusyppliersss_Click()
reportname = "stocksup.rpt"
    Show_Sales_Crystal_Report "", reportname, ""
End Sub

Private Sub mnutchpdailyanalysis_Click()
reportname = "Trackingmoney.rpt"
 STRFORMULA = ""
 Show_Sales_Crystal_Report STRFORMULA, reportname, ""

End Sub

Private Sub mnutchpmemberlist_Click()
'//TCHP_MemberList
reportname = "TCHP_MemberList.rpt"
 
 Show_Sales_Crystal_Report STRFORMULA, reportname, ""

End Sub

Private Sub mnutchpmemreport_Click()
'//tchp_itmr
'd_transdeduction_report
''generate that report from here
'Dim sno As String
'Dim ds As Date
'Dim dt As Date
'Dim r As New ADODB.Recordset
'Dim s As New ADODB.Recordset
'Dim durations As Integer
'sql = ""
'sql = "SELECT   distinct  sno   FROM         tchp_durations  order by sno "
'Set rs = oSaccoMaster.GetRecordset(sql)
'While Not rs.EOF
'sno = rs.Fields(0)
''SELECT     sno, dthcps, dthcpd, durations   FROM         tchp_durations1
'
'
'sql = ""
'sql = "select dthcps,dthcpd from tchp_durations where sno='" & sno & "' order by id"
'Set Rst = oSaccoMaster.GetRecordset(sql)
'While Not Rst.EOF
'
'ds = IIf(IsNull(Rst.Fields(0)), "01/01/1900", Rst.Fields(0))
'dt = IIf(IsNull(Rst.Fields(1)), "01/01/1900", Rst.Fields(1))
'If ds <> "" And dt <> "" Then
'If dt = "01/01/1900" Then GoTo h
'durations = DateDiff("d", ds, dt)
'sql = "INSERT INTO tchp_durations1"
'sql = sql & "                   (sno, dthcps, dthcpd, durations)"
'sql = sql & " VALUES     (" & sno & ",'" & ds & "','" & dt & "'," & durations & ")"
'oSaccoMaster.ExecuteThis (sql)
'
'ds = ""
'dt = ""
'h:
'End If
'
'Rst.MoveNext
'Wend
'
'
'
'
'
'rs.MoveNext
'Wend
'

reportname = "tchp_itmr.rpt"
 
 Show_Sales_Crystal_Report STRFORMULA, reportname, ""

End Sub

Private Sub mnutchprates_Click()
frmtchprates.Show vbModal, Me

End Sub

Private Sub mnutchptracker_Click()
frmtchpinquiry.Show vbModal, Me
End Sub

Private Sub mnuTPayroll_Click()
 reportname = "d_TransportersPayRoll.rpt"
 
 Show_Sales_Crystal_Report STRFORMULA, reportname, ""
  
'd_TransportersPayRoll.rpt
End Sub

Private Sub mnuTradersPayroll_Click()
reportname = "d_PayrollTraders.rpt"
 
 Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub mnuTraders_Click()
frmTraders.Show vbModal
End Sub

Private Sub mnuTraderap_Click()
'frmTraders.Show vbModal
End Sub

Private Sub mnutransactionslistings_Click()
frmtransactionlisting.Show vbModal
End Sub

Private Sub mnutransdeducreport_Click()
'd_transdeduction_report
reportname = "d_transdeduction_report.rpt"
 
 Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub mnuTransDetailed_Click()
'd_TransDetailed
reportname = "d_TransDetailed.rpt"
 
 Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub mnutransferofkilos_Click()
frmtransfer.Show vbModal, Me
End Sub

Private Sub mnuTransportassignment_Click()
frmTransAssign.Show vbModal
End Sub

Private Sub mnutransportdeductionsassignment_Click()
frmtransportdeductions.Show
End Sub

Private Sub mnuTransporterEnquiry_Click()
frmTransEnquery.Show vbModal
End Sub

Private Sub mnutransporterperiodicreport_Click()
'milkdeliverypertransporter
reportname = "milkdeliverypertransporter.rpt"
  Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub mnuTransportersDailyintake_Click()
'd_TransDetailed
reportname = "Transportersdailyintake.rpt"
 
 Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub mnuTransportersStatement_Click()
frmTransportStmts.Show vbModal
End Sub

Private Sub mnuTransportMode_Click()
frmTransport.Show vbModal
End Sub

Private Sub mnutrendanalysis_Click()
 reportname = "trendanalyis.rpt"
  Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub mnutypes_Click()
frmtypes.Show vbModal

End Sub

Private Sub mnuupdatefrombranch_Click()
frmimports.Show vbModal
End Sub

Private Sub mnuupdateprice_Click()
frmupdatesellingprice.Show vbModal
End Sub

Private Sub mnuUserGroups_Click()
'check the user
sql = "SELECT     UserLoginIDs,levels, UserGroup, SUPERUSER From UserAccounts where UserLoginIDs='" & User & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
If rs!Levels <> "Manager" Then
MsgBox "Access Denied", vbInformation
Exit Sub
End If
End If
frmusergroup.Show vbModal
End Sub

Private Sub mnuusermenus_Click()
sql = "SELECT     UserLoginIDs,levels, UserGroup, SUPERUSER From UserAccounts where UserLoginIDs='" & User & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
If rs!Levels <> "Manager" Then
MsgBox "Access Denied", vbInformation
Exit Sub
End If
End If
frmmenuregister.Show vbModal
End Sub

Private Sub mnuuserprevilleges_Click()
sql = "SELECT     UserLoginIDs,levels, UserGroup, SUPERUSER From UserAccounts where UserLoginIDs='" & User & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
If rs!usergroup <> "Manager" Then
MsgBox "Access Denied", vbInformation
Exit Sub
End If
End If
frmusermenuregistration.Show vbModal
End Sub

Private Sub mnuUsers_Click()
sql = "SELECT     UserLoginIDs,levels, UserGroup, SUPERUSER From UserAccounts where UserLoginIDs='" & User & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
If rs!SuperUser <> "1" Then
MsgBox "Access Denied", vbInformation
Exit Sub
End If
End If
frmsystemuser.Show vbModal
End Sub

Private Sub mnuUsersSummary_Click()
frmDUserSummary.Show vbModal
End Sub
Private Sub mnuvatremittance_Click()
frmVATremittance.Show vbModal
End Sub
Private Sub mnuVendors_Click()
frmSupplier.Show vbModal
End Sub

Private Sub mnuvehicleexp_Click()
    reportname = "incomevsevehicle.rpt"
    Show_Sales_Crystal_Report "", reportname, ""
End Sub

Private Sub mnuvehiclere_Click()
reportname = "monthlyvehc.rpt"
Show_Sales_Crystal_Report "", reportname, ""
End Sub

Private Sub mnuviewfarmerpayment_Click()
frmunpostedtransaction.Show vbModal
End Sub
Private Sub mnuviewgoodsrecived_Click()
frmrecieptslistings.Show vbModal
End Sub
Private Sub mnuviewincomestatement_Click()
'frmviewincomestatements.Show vbModal
End Sub

Private Sub mnuviewinterstorefransfer_Click()
frmviewinterstoretransfer.Show vbModal
End Sub

Private Sub mnuviewmilkcollection_Click()
frmviewmilkcollectiondetails.Show vbModal
End Sub

Private Sub mnuviewmilksales_Click()
'frmviewmilksales.Show vbModal
End Sub

Private Sub mnuviewrejectedmilk_Click()
'frmviewRejectedmilk.Show vbModal
End Sub

Private Sub mnuviewsupplierinvoice_Click()
frmviewvendorinvoices.Show vbModal
End Sub

Private Sub mnuvoidcheque_Click()
frmchequevoid.Show vbModal
End Sub


Private Sub mnuwritetodisk_Click()
frmUpdate.Show vbModal, Me
End Sub

Public Sub tchptrackers()
Dim sno As String
Dim premium As Double
Dim tmdate As Date
Dim DTPEndDate As Date
Dim balance As Double
Dim balaa As Double
DTPEndDate = Format(Get_Server_Date, "dd/mm/yyyy")
'//before all this let us clear this table called
sql = "truncate table         tchp_trxsreport"
oSaccoMaster.ExecuteThis (sql)

sql = ""
sql = "select sno,mpremium,Tmdate,statusr,balance from tchp_members where tchpactive=1 order by sno "
Set rs = oSaccoMaster.GetRecordset(sql)
While Not rs.EOF
'tchp_trxsreport
sno = rs.Fields(0)
balance = rs.Fields(4)
Dim nb As New ADODB.Recordset
sql = "select top 1 Balance  from tchp_trxs where sno='" & sno & "' order by Id desc"
Set nb = oSaccoMaster.GetRecordset(sql)
If Not nb.EOF Then
balaa = IIf(IsNull(nb.Fields(0)), 0, nb.Fields(0))
Else
balaa = 0
End If
'Debug.Print sno
premium = rs.Fields(1)
tmdate = rs.Fields(2)
Dim a As Double, b As Double, C As Double, status As String
Dim Debits As Double, creditsd As Double, creditsc As Double
status = rs.Fields(3)
Dim mtn As New ADODB.Recordset, mtn1 As New ADODB.Recordset
'//get the total debits and credits for the month in the tchp_trx table
'Debug.Print sno
sql = ""
sql = "SELECT     SUM(Debits) AS debits, SUM(CreditsD) AS creiditsd, SUM(CreditsC) AS creditc   FROM         tchp_trxs   WHERE     sno='" & sno & "' AND month(transdate)=" & month(DTPEndDate) & " AND YEAR(transdate)=" & Year(DTPEndDate) & ""
Set mtn1 = oSaccoMaster.GetRecordset(sql)
If Not mtn1.EOF Then
Debits = IIf(IsNull(mtn1.Fields(0)), 0, mtn1.Fields(0))
creditsd = IIf(IsNull(mtn1.Fields(1)), 0, mtn1.Fields(1))
creditsc = IIf(IsNull(mtn1.Fields(2)), 0, mtn1.Fields(2))
Else
Debits = 0
creditsd = 0
creditsc = 0
End If
sql = "SELECT     sno, mmonth, yyear, status FROM         tchp_status WHERE  sno='" & sno & "' and MMONTH=" & month(DTPEndDate) & " AND YYEAR=" & Year(DTPEndDate) & ""
Set mtn = oSaccoMaster.GetRecordset(sql)
If mtn.EOF Then
sql = "set dateformat dmy insert into tchp_status(sno, mmonth, yyear, status,balance,debits,creditsd,creditsc,premium) values ('" & sno & "', " & month(DTPEndDate) & ", " & Year(DTPEndDate) & ", '" & status & "'," & balaa & "," & Debits & "," & creditsd & "," & creditsc & "," & premium & ")"
oSaccoMaster.ExecuteThis (sql)
Else

sql = ""
sql = "set dateformat dmy update tchp_status set  status='" & status & "',balance=" & balaa & ",debits=" & Debits & ",creditsd=" & creditsd & ",creditsc=" & creditsc & ", premium=" & premium & " where mmonth= " & month(DTPEndDate) & " and yyear= " & Year(DTPEndDate) & " and sno='" & sno & "'"
oSaccoMaster.ExecuteThis (sql)
End If
'Debug.Print status
'tchp_trxs
Set rst = oSaccoMaster.GetRecordset("SELECT     SUM(Debits) AS a, SUM(CreditsD) AS b, SUM(CreditsC) AS c  FROM         tchp_trxs  WHERE     (sno = '" & sno & "')  GROUP BY sno")
If Not rst.EOF Then
            a = rst.Fields(0)
            b = rst.Fields(1)
            C = rst.Fields(2)
            
            '//get balance here
            
                    sql = "SELECT     balance   FROM         tchp_trxs  WHERE     sno ='" & sno & "'  ORDER BY transdate DESC, id DESC "
                    Dim rr As New ADODB.Recordset
                    Set rr = oSaccoMaster.GetRecordset(sql)
                    If Not rr.EOF Then
                    balance = rr.Fields(0)
                    Else
                    balance = 0
                    End If
                If status = "Pending" Then
                  status = "Pending"
                  GoTo kapjoel
                  End If
                 ' Debug.Print status
                  'Debug.Print sno
                  If balance = 0 And status = "New" Then
                  status = "Valid"
                  GoTo kapjoel
                  End If

                  If premium = balance Then
                  status = "Suspend"
                  GoTo kapjoel
                  End If
                  
                  If balance <= 0 Then
                  status = "Valid"
                  GoTo kapjoel
                  End If
                  If balance > premium Then
                  status = "Terminate"
                  GoTo kapjoel
                  End If
                  '//check the dates to determine if it is new or pending
                  
                  '//get the phone no and content to be send to each one
kapjoel:
                  Dim MsgContent As String, Phone As String, rt As New ADODB.Recordset
            sql = ""
            sql = "select phoneno from d_suppliers where sno='" & sno & "'"
            Set rt = oSaccoMaster.GetRecordset(sql)
            If Not rt.EOF Then
            Phone = rt.Fields(0)
            Else
            Phone = ""
            End If
            If Phone <> "" Then
                    If status = "Terminated" Then
                    
                         MsgContent = "Supplier No. " & sno & ", You have an outstanding TCHP balance of " & Format(balance, "###,###.00") & " and will be terminated from the scheme. You can rejoin the scheme in 6 months time. We will not deduct any more money from your Milk account"

                        
                    End If
                    
                    If status = "Suspend" Then
                    
                    MsgContent = "Supplier No. " & sno & ", You have an outstanding TCHP balance of " & Format(balance, "###,###.00") & " and will be Suspended from cover next month. Please pay two premiums next month to regain cover"

                    End If
                    If status = "status" Then
                    MsgContent = ""
                    End If
                    Else
                    Phone = ""
                    MsgContent = ""
            End If
                  'status here

            sql = ""
            sql = "INSERT INTO tchp_trxsreport"
            sql = sql & "                   (sno, Debits, CreditsD, CreditsC, Balance, status,premium,phone,content,msgtype)"
            sql = sql & "  VALUES     ('" & sno & "'," & a & "," & b & "," & C & "," & balance & ",'" & status & "'," & premium & ",'" & Phone & "','" & MsgContent & "','Outbox')"
            oSaccoMaster.ExecuteThis (sql)
            sql = ""
            sql = "UPDATE    tchp_members  SET      statusr='" & status & "'         where sno='" & sno & "'"
            oSaccoMaster.ExecuteThis (sql)
            
            '//insert into audit reports
            Dim rm As New ADODB.Recordset
            Set rm = oSaccoMaster.GetRecordset("select sno from tchp_audit where sno='" & sno & "'")
            If rm.EOF Then
             sql = ""
            sql = "INSERT INTO tchp_audit"
            sql = sql & "                   (sno, Debits, CreditsD, CreditsC, Balance, status,premium,jan2012,Febstatus)"
            sql = sql & "  VALUES     ('" & sno & "'," & a & "," & b & "," & C & "," & balance & ",'" & status & "'," & premium & "," & balance & ",'" & status & "')"
            oSaccoMaster.ExecuteThis (sql)
            Else
            '//do all the update here.
            'SELECT     jan2012, Febstatus, Feb2012, Marstatus, Mar2012, Aprilstatus, Apr2012, Maystatus, May2012
            'From tchp_audit
            If month(DTPEndDate) = 2 And Year(DTPEndDate) = 2012 Then
                sql = ""
                sql = "update tchp_audit set Feb2012=" & balance & ",Marstatus='" & status & "' where sno='" & sno & "' "
                oSaccoMaster.ExecuteThis (sql)
            End If
            
            If month(DTPEndDate) = 3 And Year(DTPEndDate) = 2012 Then
                sql = ""
                sql = "update tchp_audit set Mar2012=" & balance & ",Aprilstatus='" & status & "' where sno='" & sno & "' "
                oSaccoMaster.ExecuteThis (sql)
            End If
            
            If month(DTPEndDate) = 4 And Year(DTPEndDate) = 2012 Then
                sql = ""
                sql = "update tchp_audit set Apr2012=" & balance & ",Maystatus='" & status & "' where sno='" & sno & "' "
                oSaccoMaster.ExecuteThis (sql)
            End If
            
            If month(DTPEndDate) = 5 And Year(DTPEndDate) = 2012 Then
                sql = ""
                sql = "update tchp_audit set May2012=" & balance & ",junstatus='" & status & "' where sno='" & sno & "' "
                oSaccoMaster.ExecuteThis (sql)
            End If
            
            If month(DTPEndDate) = 6 And Year(DTPEndDate) = 2012 Then
                sql = ""
                sql = "update tchp_audit set june2012=" & balance & ",julstatus='" & status & "' where sno='" & sno & "' "
                oSaccoMaster.ExecuteThis (sql)
            End If
            '////*****************************other data
            If month(DTPEndDate) = 7 And Year(DTPEndDate) = 2012 Then
                sql = ""
                sql = "update tchp_audit set jul2012=" & balance & ",augstatus='" & status & "' where sno='" & sno & "' "
                oSaccoMaster.ExecuteThis (sql)
            End If
            If month(DTPEndDate) = 8 And Year(DTPEndDate) = 2012 Then
                sql = ""
                sql = "update tchp_audit set aug2012=" & balance & ",sepstatus='" & status & "' where sno='" & sno & "' "
                oSaccoMaster.ExecuteThis (sql)
            End If
            If month(DTPEndDate) = 9 And Year(DTPEndDate) = 2012 Then
                sql = ""
                sql = "update tchp_audit set sep2012=" & balance & ",octstatus='" & status & "' where sno='" & sno & "' "
                oSaccoMaster.ExecuteThis (sql)
            End If
            If month(DTPEndDate) = 10 And Year(DTPEndDate) = 2012 Then
                sql = ""
                sql = "update tchp_audit set oct2012=" & balance & ",novstatus='" & status & "' where sno='" & sno & "' "
                oSaccoMaster.ExecuteThis (sql)
            End If
            If month(DTPEndDate) = 11 And Year(DTPEndDate) = 2012 Then
                sql = ""
                sql = "update tchp_audit set nov2012=" & balance & ",decstatus='" & status & "' where sno='" & sno & "' "
                oSaccoMaster.ExecuteThis (sql)
            End If
            If month(DTPEndDate) = 12 And Year(DTPEndDate) = 2012 Then
                sql = ""
                sql = "update tchp_audit set dec2012=" & balance & ",janstatus1='" & status & "' where sno='" & sno & "' "
                oSaccoMaster.ExecuteThis (sql)
            End If
           
           '2013
            If month(DTPEndDate) = 2 And Year(DTPEndDate) = 2013 Then
                sql = ""
                sql = "update tchp_audit set Feb2013=" & balance & ",Marstatus1='" & status & "' where sno='" & sno & "' "
                oSaccoMaster.ExecuteThis (sql)
            End If
            
            If month(DTPEndDate) = 3 And Year(DTPEndDate) = 2013 Then
                sql = ""
                sql = "update tchp_audit set Mar2013=" & balance & ",Aprilstatus1='" & status & "' where sno='" & sno & "' "
                oSaccoMaster.ExecuteThis (sql)
            End If
            
            If month(DTPEndDate) = 4 And Year(DTPEndDate) = 2013 Then
                sql = ""
                sql = "update tchp_audit set Apr2013=" & balance & ",Maystatus1='" & status & "' where sno='" & sno & "' "
                oSaccoMaster.ExecuteThis (sql)
            End If
            
            If month(DTPEndDate) = 5 And Year(DTPEndDate) = 2013 Then
                sql = ""
                sql = "update tchp_audit set May2013=" & balance & ",junstatus1='" & status & "' where sno='" & sno & "' "
                oSaccoMaster.ExecuteThis (sql)
            End If
            
            If month(DTPEndDate) = 6 And Year(DTPEndDate) = 2013 Then
                sql = ""
                sql = "update tchp_audit set june2013=" & balance & ",julstatus1='" & status & "' where sno='" & sno & "' "
                oSaccoMaster.ExecuteThis (sql)
            End If
            '////*****************************other data
            If month(DTPEndDate) = 7 And Year(DTPEndDate) = 2012 Then
                sql = ""
                sql = "update tchp_audit set jul2013=" & balance & ",augstatus1='" & status & "' where sno='" & sno & "' "
                oSaccoMaster.ExecuteThis (sql)
            End If
            If month(DTPEndDate) = 8 And Year(DTPEndDate) = 2013 Then
                sql = ""
                sql = "update tchp_audit set aug2013=" & balance & ",sepstatus1='" & status & "' where sno='" & sno & "' "
                oSaccoMaster.ExecuteThis (sql)
            End If
            If month(DTPEndDate) = 9 And Year(DTPEndDate) = 2013 Then
                sql = ""
                sql = "update tchp_audit set sep2013=" & balance & ",octstatus1='" & status & "' where sno='" & sno & "' "
                oSaccoMaster.ExecuteThis (sql)
            End If
            If month(DTPEndDate) = 10 And Year(DTPEndDate) = 2013 Then
                sql = ""
                sql = "update tchp_audit set oct2013=" & balance & ",novstatus1='" & status & "' where sno='" & sno & "' "
                oSaccoMaster.ExecuteThis (sql)
            End If
            If month(DTPEndDate) = 11 And Year(DTPEndDate) = 2013 Then
                sql = ""
                sql = "update tchp_audit set nov2013=" & balance & ",decstatus1='" & status & "' where sno='" & sno & "' "
                oSaccoMaster.ExecuteThis (sql)
            End If
            If month(DTPEndDate) = 12 And Year(DTPEndDate) = 2013 Then
                sql = ""
                sql = "update tchp_audit set dec2013=" & balance & ",janstatus1='" & status & "' where sno='" & sno & "' "
                'oSaccoMaster.ExecuteThis (sql)
            End If
           
            End If
            '//got and update all the months
            
    Else
    a = 0
    b = 0
    C = 0
    balance = 0
                sql = ""
            sql = "INSERT INTO tchp_trxsreport"
            sql = sql & "                   (sno, Debits, CreditsD, CreditsC, Balance, status,premium,phone,content,msgtype)"
            sql = sql & "  VALUES     ('" & sno & "'," & a & "," & b & "," & C & "," & balance & ",'" & status & "'," & premium & ",'" & Phone & "','" & MsgContent & "','Outbox')"
            oSaccoMaster.ExecuteThis (sql)
            sql = ""
            sql = "UPDATE    tchp_members  SET      statusr='" & status & "'         where sno='" & sno & "'"
            oSaccoMaster.ExecuteThis (sql)
            
            
            
            '//put also in the audit report
            
            Set rm = oSaccoMaster.GetRecordset("select sno from tchp_audit where sno='" & sno & "'")
            If rm.EOF Then
             sql = ""
            sql = "INSERT INTO tchp_audit"
            sql = sql & "                   (sno, Debits, CreditsD, CreditsC, Balance, status,premium,jan2012,Febstatus)"
            sql = sql & "  VALUES     ('" & sno & "'," & a & "," & b & "," & C & "," & balance & ",'" & status & "'," & premium & "," & balance & ",'" & status & "')"
            oSaccoMaster.ExecuteThis (sql)
            Else
            '//do all the update here.
            'SELECT     jan2012, Febstatus, Feb2012, Marstatus, Mar2012, Aprilstatus, Apr2012, Maystatus, May2012
            'From tchp_audit
            If month(DTPEndDate) = 2 And Year(DTPEndDate) = 2012 Then
                sql = ""
                sql = "update tchp_audit set Feb2012=" & balance & ",Marstatus='" & status & "' where sno='" & sno & "' "
                oSaccoMaster.ExecuteThis (sql)
            End If
            
            If month(DTPEndDate) = 3 And Year(DTPEndDate) = 2012 Then
                sql = ""
                sql = "update tchp_audit set Mar2012=" & balance & ",Aprilstatus='" & status & "' where sno='" & sno & "' "
                oSaccoMaster.ExecuteThis (sql)
            End If
            
            If month(DTPEndDate) = 4 And Year(DTPEndDate) = 2012 Then
                sql = ""
                sql = "update tchp_audit set Apr2012=" & balance & ",Maystatus='" & status & "' where sno='" & sno & "' "
                oSaccoMaster.ExecuteThis (sql)
            End If
            
            If month(DTPEndDate) = 5 And Year(DTPEndDate) = 2012 Then
                sql = ""
                sql = "update tchp_audit set May2012=" & balance & ",junstatus='" & status & "' where sno='" & sno & "' "
                oSaccoMaster.ExecuteThis (sql)
            End If
            
            If month(DTPEndDate) = 6 And Year(DTPEndDate) = 2012 Then
                sql = ""
                sql = "update tchp_audit set june2012=" & balance & ",julstatus='" & status & "' where sno='" & sno & "' "
                oSaccoMaster.ExecuteThis (sql)
            End If
            
            End If
            
            
End If
sql = ""
Phone = ""
rs.MoveNext
Wend
End Sub

Private Sub mnuwritetodisc_Click()
frmUpdate.Show vbModal
End Sub

Private Sub mnuyyyyyyyyyy_Click()
reportname = "Milkintakeyearly.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

'Private Sub Timer1_Timer()
'''Label1
'Dim sno As String, prem As Double, NetP As Double, statusr As String
'Dim txtTCHPBalances As Double, balance As Double, rr As New ADODB.Recordset
'Dim RTS As New ADODB.Recordset, ans1 As String
'Dim myclass As cdbase
'Set myclass = New cdbase
'Set cn = CreateObject("adodb.connection")
'Provider = myclass.OpenCon
'cn.Open Provider, "atm"
'Set rs = CreateObject("adodb.recordset")
'sql = "select server from d_company"
'rs.Open sql, cn
'If Not rs.EOF Then
'
'If Not IsNull(rs!server) Then sserver = rs!server
'End If
'Dim rrt As New ADODB.Recordset
'    On Error Resume Next
'    lbl.Move lbl.Left - 100
'    Label1.Move Label1.Left - 100
'    If lbl.Left <= -7000 Then
'        lbl.Visible = True
'        Label1.Visible = True
'
'        If sserver = 2 Then
'        GoTo grockings
'        End If
'
'        '////cash sms
'        If Day(tdate) = 99 Then '15 of the month
'        '//get the tchp memebers
'        Dim mm1, mm2, mm3 As Date
'
'        mm1 = Mid(tdate, 4, 2) 'month
'        mm2 = Mid(tdate, 7, 4) 'year
'        'construct date
'        mm3 = csmsdate & "/" & mm1 & "/" & mm2
'        Dim hcp As Long, aarno As String
'        sql = ""
'            sql = "set dateformat dmy delete  from         Messages1 where transdate<>'" & mm3 & "'"
'            oSaccoMaster.ExecuteThis (sql)
'
'        sql = "SELECT     sno,mpremium,statusr,tchpactive,aarno  FROM         tchp_members  where tchpactive=1 and smssend=0 ORDER BY sno"
'            Set rs = oSaccoMaster.GetRecordset(sql)
'            While Not rs.EOF
'                sno = rs.Fields(0)
'                prem = rs.Fields(1)
'                statusr = Trim(rs.Fields(2))
'                hcp = rs.Fields(3)
'                aarno = rs.Fields(4)
'                'Debug.Print sno
'                If statusr = "Terminate" Then GoTo Noterm
'
'                '//get the balance then you will add up
'
'                sql = "SELECT     balance   FROM         tchp_trxs  WHERE     sno ='" & sno & "'  ORDER BY transdate DESC, id DESC "
'
'                Set rr = oSaccoMaster.GetRecordset(sql)
'                If Not rr.EOF Then
'                txtTCHPBalances = rr.Fields(0)
'                Else
'                txtTCHPBalances = 0
'                End If
'
'                Dim Phone As String, content As String
'                '//get the phone no from the supplier
'                sql = ""
'                Dim rsms As New ADODB.Recordset
'                Set rsms = oSaccoMaster.GetRecordset("select phoneno from d_suppliers where sno='" & sno & "'")
'                If Not rsms.EOF Then
'                Phone = rsms.Fields(0)
'                'content = "Please remember to pay  for TCHP Premium before 25th: Your Current Balance is: " & txtTCHPBalances & ""
'                Else
'                Phone = ""
'                End If
'
'                If txtTCHPBalances > 0 And hcp = 1 Then
'                If txtTCHPBalances = prem Then
'                '//message 1
'                   content = "Supplier No " & sno & ". You have an outstanding TCHP balance of " & txtTCHPBalances & ", Please Pay this with Milk or Cash Before 25th to avoid being Suspended"
'                End If
'                If txtTCHPBalances > prem Then
'                '//message 2
'                content = "Supplier No " & sno & ". You have an outstanding TCHP balance of " & txtTCHPBalances & ", Please Pay this with Milk or Cash Before 25th to avoid being Terminated"
'                End If
'                End If
'                '//insert into the sms tables
'        If txtTCHPBalances > 0 Then
'        If Phone <> "" Then
'        If Len(Phone) >= 10 Then
'
'        '//do populate to the trx warning table
'        Dim os As New ADODB.Recordset
'        sql = "set dateformat dmy select * from Messages1 where transdate='" & mm3 & "' and sno='" & sno & "'"
'        Set os = oSaccoMaster.GetRecordset(sql)
'        If os.EOF Then
'        strSQL = "INSERT INTO Messages1(Telephone,Content,ProcessTime, MsgType,Source,sno,status,aarno,balance,premium,transdate)"
'        strSQL = strSQL & "Values ('" & Phone & "','" & content & "',getDate(),'Outbox','" & User & "','" & sno & "','" & statusr & "','" & aarno & "'," & txtTCHPBalances & "," & prem & ",'" & mm3 & "')"
'        oSaccoMaster.ExecuteThis (strSQL)
'        sql = ""
'        sql = "update tchp_members set smssend=1 where sno='" & sno & "'"
'        oSaccoMaster.ExecuteThis (sql)
'        End If
'        End If
'        End If
'        End If
'Noterm:
'            rs.MoveNext
'            Wend
'
'        End If
'        '/////end the cash sms
'        If Day(tdate) = 99 Then ''//deal with debits
'        sql = "SELECT     sno,mpremium  FROM         tchp_members  where tchpactive=1  ORDER BY sno"
'            Set rs = oSaccoMaster.GetRecordset(sql)
'            While Not rs.EOF
'                sno = rs.Fields(0)
'                prem = rs.Fields(1)
'
'                '//get the balance then you will add up
'
'                sql = "SELECT     balance   FROM         tchp_trxs  WHERE     sno ='" & sno & "'  ORDER BY transdate DESC, id DESC "
'                Dim sta As String, mts As New ADODB.Recordset
'                Set rr = oSaccoMaster.GetRecordset(sql)
'                If Not rr.EOF Then
'                txtTCHPBalances = rr.Fields(0)
'                Else
'                txtTCHPBalances = 0
'                End If
'                balance = prem + txtTCHPBalances
'                sql = "SELECT     *  FROM         tchp_trxs  WHERE     (sno = '" & sno & "') AND (MONTH(transdate) = " & month(tdate) & ") AND (YEAR(transdate) = " & Year(tdate) & ") AND (description = 'Debit')"
'
'                Set rrt = oSaccoMaster.GetRecordset(sql)
'                If rrt.EOF Then
'                sql = ""
'                sql = "set dateformat dmy INSERT INTO tchp_trxs"
'                sql = sql & "     (sno,transdate, description, Debits, CreditsD, CreditsC, Balance, auditid)"
'                sql = sql & " VALUES     ('" & sno & "','" & tdate & "','Debit'," & prem & ",0,0," & balance & ",'System')"
'                oSaccoMaster.ExecuteThis (sql)
'                sql = ""
'                sql = "select statusr from tchp_members where sno='" & sno & "'"
'                Set mts = oSaccoMaster.GetRecordset(sql)
'                If Not mts.EOF Then
'                sta = mts.Fields(0)
'                If sta = "Pending" Then
'                sql = ""
'                sql = "update tchp_members set statusr='New' where sno='" & sno & "'"
'                oSaccoMaster.ExecuteThis (sql)
'                End If
'                End If
'                End If
'                '//update the status if it was pending to new
'
'                sql = ""
'                sql = "update tchp_members set smssend=0 where sno='" & sno & "'"
'                oSaccoMaster.ExecuteThis (sql)
'                rs.MoveNext
'            Wend
'        End If
'
'        '//end if the debits
'        lbl.Visible = False
'        Label1.Visible = False
'        lbl.Move lbl.Left - 100
'        Label1.Move Label1 - 100
'        If lbl.Left <= -13500 Then
'            lbl.Left = 12000
'            If Day(tdate) = 99 Then  '// sms purely and the audit report
'           '' tchptrackers
'           If Day(tdate) > 28 And Day(tdate) <= 31 Then
'GoTo herll
'End If
'If Day(tdate) >= 1 And Day(tdate) <= 2 Then
'GoTo herll
'End If
'            sql = "SELECT     sno,mpremium,statusr  FROM         tchp_members  where tchpactive=1  ORDER BY sno"
'            Set rs = oSaccoMaster.GetRecordset(sql)
'            While Not rs.EOF
'            sno = rs.Fields(0)
'            prem = rs.Fields(1)
'
'            '//get milk balance first
'                Startdate = DateSerial(Year(tdate), month(tdate), 1)
'                Enddate = DateSerial(Year(tdate), month(tdate) + 1, 1 - 1)
'
'                Set Rst = oSaccoMaster.GetRecordset("d_sp_SupNet '" & sno & "','" & Startdate & "','" & Enddate & "', 0")
'
'                If Not IsNull(Rst.Fields(1)) Then
'                NetP = Rst.Fields(1)
'                Else
'                NetP = "0.00"
'                End If
'                Set Rst = oSaccoMaster.GetRecordset("d_sp_SupNet '" & sno & "','" & Startdate & "','" & Enddate & "', 1")
'                If Not IsNull(Rst.Fields(0)) Then
'                NetP = NetP - Rst.Fields(0)
'                Else
'                NetP = NetP - 0
'                End If
'            '//get the balance of the premium
'
'
'                sql = "SELECT     balance   FROM         tchp_trxs  WHERE     sno ='" & sno & "'  ORDER BY transdate DESC, id DESC "
'
'                Set rr = oSaccoMaster.GetRecordset(sql)
'                If Not rr.EOF Then
'                txtTCHPBalances = rr.Fields(0)
'                Else
'                txtTCHPBalances = 0
'                End If
'                '//two months
'                   If txtTCHPBalances > 0 Then '//if the amount is able to cover the all amount
'                    If NetP >= txtTCHPBalances Then
'                    prem = txtTCHPBalances
'                    End If
'                    End If
'                    If txtTCHPBalances > 0 Then
'                    If NetP >= prem Then
'                    prem = prem
'                    End If
'                    End If
'
'
'
'            '//check if it has been deduction before
'            'Debug.Print sno
'            If NetP >= prem Then
'            If txtTCHPBalances > 0 Then
'               balance = txtTCHPBalances - CDbl(prem)
'
''//put the deduction
'                Set cn = New ADODB.Connection
'                sql = "d_sp_SupplierDeduct '" & sno & "','" & tdate & "','TCHP'," & prem & ",'" & Startdate & "','" & Enddate & "'," & Year(tdate) & ",'" & User & "','TCHP','HQ'"
'                oSaccoMaster.ExecuteThis (sql)
'                '//put into the premium table
'                Dim ru As New ADODB.Recordset
'                'check if it has already been deducted on the same day
'                sql = ""
'                sql = "SET              dateformat dmy   SELECT     *    FROM         tchp_trxs   WHERE     description = 'Deduction(Auto)' AND sno = '" & sno & "' AND transdate = '" & tdate & "'"
'                Set ru = oSaccoMaster.GetRecordset(sql)
'                If ru.EOF Then
'                sql = ""
'                sql = "set dateformat dmy INSERT INTO tchp_trxs"
'                sql = sql & "     (sno,transdate, description, Debits, CreditsD, CreditsC, Balance, auditid)"
'                sql = sql & " VALUES     ('" & sno & "','" & tdate & "','Deduction(Auto)',0," & prem & ",0," & balance & ",'System')"
'                oSaccoMaster.ExecuteThis (sql)
'                If balance <= 0 Then
'                sql = ""
'                sql = "update tchp_members set statusr='Valid' where sno='" & sno & "'"
'                'oSaccoMaster.ExecuteThis (sql)
'                End If
'                End If
'                Else
'
'                '///put here those stuff
'
'            sql = ""
'                sql = "select statusr from tchp_members where sno='" & sno & "'"
'                Set mts = oSaccoMaster.GetRecordset(sql)
'                If Not mts.EOF Then
'                sta = mts.Fields(0)
'                If sta = "New" And Day(tdate) > 28 Then
'                sql = ""
'                sql = "update tchp_members set statusr='Valid' where sno='" & sno & "'"
'                oSaccoMaster.ExecuteThis (sql)
'                End If
'                End If
'                End If
'                End If
'            '//update the live report
'            '**********************************************start herer
'           ' truncate the report table
'           'Debug.Print sno
'Dim a As Double, b As Double, C As Double, status As String, premium As Double
'Dim fr As New ADODB.Recordset
'sql = "delete from          tchp_trxsreport where sno='" & sno & "'"
'oSaccoMaster.ExecuteThis (sql)
'
'sql = "select statusr,mpremium from tchp_members where sno='" & sno & "'"
'    Set mts = oSaccoMaster.GetRecordset(sql)
'    If Not mts.EOF Then
'    status = mts.Fields(0)
'    premium = mts.Fields(1)
'    End If
'
'Set fr = oSaccoMaster.GetRecordset("SELECT     SUM(Debits) AS a, SUM(CreditsD) AS b, SUM(CreditsC) AS c  FROM         tchp_trxs  WHERE     (sno = '" & sno & "')  GROUP BY sno")
'If Not fr.EOF Then
'a = fr.Fields(0)
'b = fr.Fields(1)
'C = fr.Fields(2)
'Else
'a = 0
'b = 0
'C = 0
'End If
'
''get the balance
'
'sql = "SELECT     balance   FROM         tchp_trxs  WHERE     sno ='" & sno & "'  ORDER BY transdate DESC, id DESC "
'
'Set rr = oSaccoMaster.GetRecordset(sql)
'If Not rr.EOF Then
'balance = rr.Fields(0)
'
'Else
'balance = 0
'
'End If
'
''insert into the table for the report
'
'sql = ""
'sql = "INSERT INTO tchp_trxsreport"
'sql = sql & "                   (sno, Debits, CreditsD, CreditsC, Balance, status,premium,phone,content,msgtype)"
'sql = sql & "  VALUES     ('" & sno & "'," & a & "," & b & "," & C & "," & balance & ",'" & status & "'," & premium & ",'" & Phone & "','" & MsgContent & "','Outbox')"
'oSaccoMaster.ExecuteThis (sql)
'sql = ""
'sql = "UPDATE    tchp_members  SET      statusr='" & status & "'         where sno='" & sno & "'"
'oSaccoMaster.ExecuteThis (sql)
'
''// UPDATE THE STATUS TABLE
'Dim mtn As New ADODB.Recordset
'sql = "SELECT     sno, mmonth, yyear, status FROM         tchp_status WHERE  sno='" & sno & "' and MMONTH=" & month(tdate) & " AND YYEAR=" & Year(tdate) & ""
'Set mtn = oSaccoMaster.GetRecordset(sql)
'If mtn.EOF Then
'sql = "set dateformat dmy insert into tchp_status(sno, mmonth, yyear, status,balance) values ('" & sno & "', " & month(tdate) & ", " & Year(tdate) & ", '" & status & "'," & balance & ")"
'oSaccoMaster.ExecuteThis (sql)
'Else
'sql = ""
'sql = "set dateformat dmy update tchp_status set  status='" & status & "',balance=" & balance & " where mmonth= " & month(tdate) & " and yyear= " & Year(tdate) & " and sno='" & sno & "'"
'oSaccoMaster.ExecuteThis (sql)
'End If
'
'           ' ***********************************************end here
'            '//get the dedcution for d_supplier_deduc
'            NetP = 0
'            prem = 0
'            txtTCHPBalances = 0
'            balance = 0
'            premium = 0
'            status = ""
'            a = 0
'            b = 0
'            C = 0
'            rs.MoveNext
'            Wend
'            End If
'herll:
'
'
'            Label1.Left = 12000
'            lbl.Visible = False
'grockings:
'            If Day(tdate) = deddate Then
'                If sserver = 2 And hhhh = 0 Then
'                tchptrackers
'                hhhh = hhhh + 1
'                ElseIf sserver = 1 Then
'                tchptrackers
'                End If
'
'
'            End If
'horai:
'            Label1.Visible = False
'            lbl.Visible = True
'            Label1.Visible = True
'
'            '//check if the milk is sufficeint then you can deduct immediately
'
'
'
'            lbl.Left = 12000
'            Label1.Left = 12000
'        End If
'    End If
'
'End Sub



Private Sub Picture2_Click()
 
End Sub

