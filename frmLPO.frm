VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmLPO 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Local Purchase Order"
   ClientHeight    =   7785
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   8220
   LinkTopic       =   "Form1"
   ScaleHeight     =   7785
   ScaleWidth      =   8220
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   7320
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   7320
      Width           =   975
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6735
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   11880
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   12648447
      TabCaption(0)   =   "Purchase Order Info"
      TabPicture(0)   =   "frmLPO.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label12"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label9"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label8"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label7"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label6"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label5"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label4"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label3"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label13"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label14"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label15"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "DTPlpodate"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "DTPduedate"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtRemarks"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtLPOSerial"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cbovendors"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtPoNo"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtIName"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtRefNo"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtQnty"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtPPrice"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).ControlCount=   21
      TabCaption(1)   =   "Purchase Order Items"
      TabPicture(1)   =   "frmLPO.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label11"
      Tab(1).Control(1)=   "Label10"
      Tab(1).Control(2)=   "Lvwitems"
      Tab(1).Control(3)=   "lvwselecteditems"
      Tab(1).Control(4)=   "cmdremove"
      Tab(1).Control(5)=   "cmdadd"
      Tab(1).ControlCount=   6
      Begin VB.TextBox txtPPrice 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4320
         TabIndex        =   30
         Top             =   4080
         Width           =   1815
      End
      Begin VB.TextBox txtQnty 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4320
         TabIndex        =   28
         Top             =   3720
         Width           =   1815
      End
      Begin VB.TextBox txtRefNo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4320
         TabIndex        =   27
         Top             =   3360
         Width           =   1815
      End
      Begin VB.TextBox txtIName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4320
         TabIndex        =   25
         Top             =   840
         Width           =   3735
      End
      Begin VB.TextBox txtPoNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   375
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   420
         Width           =   2295
      End
      Begin VB.ComboBox cbovendors 
         Height          =   315
         Left            =   4320
         TabIndex        =   12
         Top             =   1380
         Width           =   2295
      End
      Begin VB.TextBox txtLPOSerial 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4320
         TabIndex        =   9
         Top             =   2940
         Width           =   1815
      End
      Begin VB.TextBox txtRemarks 
         Height          =   2055
         Left            =   360
         TabIndex        =   8
         Top             =   4500
         Width           =   7575
      End
      Begin VB.CommandButton cmdadd 
         Caption         =   "Add"
         Height          =   375
         Left            =   -69960
         TabIndex        =   5
         Top             =   2700
         Width           =   975
      End
      Begin VB.CommandButton cmdremove 
         Caption         =   "Remove"
         Height          =   375
         Left            =   -68760
         TabIndex        =   4
         Top             =   2700
         Width           =   975
      End
      Begin MSComctlLib.ListView lvwselecteditems 
         Height          =   1755
         Left            =   -74760
         TabIndex        =   6
         Top             =   3240
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   3096
         View            =   3
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ReqID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Item Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Quantity"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "POCOST"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView Lvwitems 
         Height          =   1815
         Left            =   -74760
         TabIndex        =   7
         Top             =   780
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   3201
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ReqID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Item Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Quantity"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Price"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComCtl2.DTPicker DTPduedate 
         Height          =   375
         Left            =   4320
         TabIndex        =   10
         Top             =   2340
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         Format          =   20185089
         CurrentDate     =   40110
      End
      Begin MSComCtl2.DTPicker DTPlpodate 
         Height          =   375
         Left            =   4320
         TabIndex        =   11
         Top             =   1860
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         Format          =   20185089
         CurrentDate     =   40110
      End
      Begin VB.Label Label15 
         Caption         =   "Price Per Item"
         Height          =   255
         Left            =   2640
         TabIndex        =   31
         Top             =   4080
         Width           =   1215
      End
      Begin VB.Label Label14 
         Caption         =   "Quantity"
         Height          =   255
         Left            =   2640
         TabIndex        =   29
         Top             =   3720
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "Ref. No."
         Height          =   255
         Left            =   2640
         TabIndex        =   26
         Top             =   3360
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Header"
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
         Left            =   240
         TabIndex        =   23
         Top             =   420
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "PO No"
         Height          =   375
         Left            =   2520
         TabIndex        =   22
         Top             =   540
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Vendor"
         Height          =   375
         Left            =   2520
         TabIndex        =   21
         Top             =   1500
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "LPO Date"
         Height          =   375
         Left            =   2520
         TabIndex        =   20
         Top             =   1980
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Due Date"
         Height          =   375
         Left            =   2520
         TabIndex        =   19
         Top             =   2460
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "LPO Serial No"
         Height          =   255
         Left            =   2520
         TabIndex        =   18
         Top             =   2940
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Remarks"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   17
         Top             =   4260
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Available Items"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74760
         TabIndex        =   16
         Top             =   420
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "Selected Items"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74760
         TabIndex        =   15
         Top             =   2940
         Width           =   1695
      End
      Begin VB.Label Label12 
         Caption         =   "Item Name :"
         Height          =   375
         Left            =   2520
         TabIndex        =   14
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Purchase Orders"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   9015
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Purchase Orders"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9015
   End
End
Attribute VB_Name = "frmLPO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmSave_Click()

If Trim(txtIName) = "" Then
    MsgBox "Please selected Item Name."
        txtIName.SetFocus
    Exit Sub
End If

If Trim(txtQnty) = "" Then
    MsgBox "Please enter the quantity."
        txtQnty.SetFocus
    Exit Sub
End If

If Trim(txtRefNo) = "" Then
    MsgBox "Please enter the reference number."
        txtRefNo.SetFocus
    Exit Sub
End If

If Trim(txtPPrice) = "" Then
    MsgBox "Please enter price per item."
        txtPPrice.SetFocus
    Exit Sub
End If

If Trim(cbovendors.Text) = "" Then
    MsgBox "Please select vendor."
        cbovendors.SetFocus
    Exit Sub
End If
'd_sp_Requisition @RNo char(35), @TransDate varchar(12), @CostCentre varchar(150), @ServiceReq bit, @IName varchar (150), @Make varchar(150), @Qnty float, @Description varchar(300), @AuditID varchar (50),@pricing money,@Date varchar(12)   AS
sql = ""
sql = sql & "d_sp_Requisition '" & txtRefNo & "','" & Date & "',' ',0,'" & txtIName & "',' '," & txtQnty & ",'Direct Order','" & User & "'," & txtPPrice & ",'" & DTPlpodate & "'"
oSaccoMaster.ExecuteThis (sql)


'd_sp_LPO @PNo bigint, @TransDate varchar(12),@DueDate varchar(12), @Serial varchar(100),@RefNo varchar(50),@user varchar(35) as
oSaccoMaster.ExecuteThis ("d_sp_LPO " & txtPoNo & ",'" & DTPlpodate & "','" & DTPduedate & "','" & txtLPOSerial & "','" & txtRefNo & "','" & User & "','" & txtRemarks & "','" & cbovendors & "'")
MsgBox "Records saved successfully!"
Form_Load
End Sub
Private Sub Command1_Click()
Unload Me
End Sub
Private Sub Form_Load()
txtLPOSerial = ""
txtIName = ""
cbovendors.Text = ""
txtRemarks = ""

DTPduedate = Format(Get_Server_Date + 7, "dd/mm/yyyy")
DTPlpodate = Format(Get_Server_Date, "dd/mm/yyyy")

'//LOAD THE VENDORS
sql = "SELECT  CompanyName  FROM         ag_Supplier1 order by companyname"
Set rs = oSaccoMaster.GetRecordset(sql)
                Do While Not rs.EOF
                cbovendors.AddItem rs.Fields(0)
                        rs.MoveNext
                    Loop

Set rs = oSaccoMaster.GetRecordset("d_sp_PoNo")
If Not rs.EOF Then
txtPoNo = CCur(rs.Fields(0)) + 1
Else
txtPoNo = "1"
End If
End Sub

