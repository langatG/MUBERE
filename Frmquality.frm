VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frmquality 
   Caption         =   "Quality form"
   ClientHeight    =   6630
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10080
   LinkTopic       =   "Form1"
   ScaleHeight     =   6630
   ScaleWidth      =   10080
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Opttransporter 
      Caption         =   "Transporter"
      Height          =   495
      Left            =   1920
      TabIndex        =   42
      Top             =   120
      Width           =   1215
   End
   Begin VB.OptionButton Optsupplier 
      Caption         =   "Supplier"
      Height          =   495
      Left            =   480
      TabIndex        =   41
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Close"
      Height          =   375
      Left            =   5640
      TabIndex        =   40
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Save"
      Height          =   375
      Left            =   4080
      TabIndex        =   39
      Top             =   5880
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   375
      Left            =   2640
      TabIndex        =   38
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Frame Frmtransporterdetails 
      Caption         =   "Transporter Details"
      Height          =   2655
      Left            =   360
      TabIndex        =   1
      Top             =   3120
      Width           =   8535
      Begin VB.TextBox txtamount 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1440
         TabIndex        =   37
         Top             =   1800
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DTPenddate 
         Height          =   375
         Left            =   3240
         TabIndex        =   35
         Top             =   1320
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   118030337
         CurrentDate     =   42715
      End
      Begin MSComCtl2.DTPicker DTPstartdate 
         Height          =   375
         Left            =   960
         TabIndex        =   33
         Top             =   1320
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   118030337
         CurrentDate     =   42715
      End
      Begin VB.TextBox txtquan 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5760
         TabIndex        =   31
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtrti 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3240
         TabIndex        =   29
         Top             =   840
         Width           =   1095
      End
      Begin VB.ComboBox Cboqlity 
         Height          =   315
         ItemData        =   "Frmquality.frx":0000
         Left            =   960
         List            =   "Frmquality.frx":000D
         TabIndex        =   27
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox Txtcno 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7200
         TabIndex        =   25
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox Txttname 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3240
         TabIndex        =   23
         Top             =   360
         Width           =   3255
      End
      Begin VB.TextBox Txttno 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   960
         TabIndex        =   21
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label17 
         Caption         =   "Actual Amount"
         Height          =   195
         Left            =   240
         TabIndex        =   36
         Top             =   1920
         Width           =   1035
      End
      Begin VB.Label Label16 
         Caption         =   "Enddate"
         Height          =   375
         Left            =   2520
         TabIndex        =   34
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label15 
         Caption         =   "Startdate"
         Height          =   375
         Left            =   240
         TabIndex        =   32
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label14 
         Caption         =   "Quantity(kgs)"
         Height          =   255
         Left            =   4560
         TabIndex        =   30
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label13 
         Caption         =   "Rate"
         Height          =   255
         Left            =   2640
         TabIndex        =   28
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "Quality"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "Canno"
         Height          =   255
         Left            =   6600
         TabIndex        =   24
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "T Name"
         Height          =   255
         Left            =   2400
         TabIndex        =   22
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "Tno"
         Height          =   255
         Left            =   360
         TabIndex        =   20
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Frmdetails 
      Caption         =   "Suppliers quality details"
      Height          =   2295
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   8535
      Begin VB.TextBox txtamt 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1320
         TabIndex        =   19
         Top             =   1680
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DTPedate 
         Height          =   375
         Left            =   3120
         TabIndex        =   17
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   118030337
         CurrentDate     =   42715
      End
      Begin MSComCtl2.DTPicker DTPsdate 
         Height          =   375
         Left            =   840
         TabIndex        =   15
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   118030337
         CurrentDate     =   42715
      End
      Begin VB.TextBox txtqty 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5640
         TabIndex        =   13
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtrate 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3120
         TabIndex        =   11
         Top             =   720
         Width           =   1095
      End
      Begin VB.ComboBox cboquality 
         Height          =   315
         ItemData        =   "Frmquality.frx":0023
         Left            =   840
         List            =   "Frmquality.frx":0030
         TabIndex        =   9
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtcanno 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7080
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtsname 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3120
         TabIndex        =   5
         Top             =   240
         Width           =   3255
      End
      Begin VB.TextBox txtsno 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   840
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Actual Amount"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   1800
         Width           =   1035
      End
      Begin VB.Label Label7 
         Caption         =   "Enddate"
         Height          =   375
         Left            =   2400
         TabIndex        =   16
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Startdate"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Quantity(kgs)"
         Height          =   255
         Left            =   4440
         TabIndex        =   12
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblrate 
         Caption         =   "Rate"
         Height          =   255
         Left            =   2520
         TabIndex        =   10
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Quality"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Canno"
         Height          =   255
         Left            =   6480
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "SName"
         Height          =   255
         Left            =   2280
         TabIndex        =   4
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Sno"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   375
      End
   End
End
Attribute VB_Name = "Frmquality"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cboqlity_Change()
Set rst = New ADODB.Recordset
sql = "select * from Qsetup where Quality='" & Cboqlity & "'"
Set rst = oSaccoMaster.GetRecordset(sql)
If Not rst.EOF Then
txtrti = Trim(rst!irate)
'txtcanno = rst!canno
End If
End Sub

Private Sub Cboqlity_Click()
Set rst = New ADODB.Recordset
sql = "select * from Qsetup where Quality='" & Cboqlity & "'"
Set rst = oSaccoMaster.GetRecordset(sql)
If Not rst.EOF Then
txtrti = Trim(rst!irate)
'txtcanno = rst!canno
End If
End Sub

Private Sub cboquality_Change()
Set rst = New ADODB.Recordset
sql = "select * from Qsetup where Quality='" & Cboquality & "'"
Set rst = oSaccoMaster.GetRecordset(sql)
If Not rst.EOF Then
txtrate = Trim(rst!irate)
'txtcanno = rst!canno
End If
End Sub

Private Sub cboquality_Click()
Set rst = New ADODB.Recordset
sql = "select * from Qsetup where Quality='" & Cboquality & "'"
Set rst = oSaccoMaster.GetRecordset(sql)
If Not rst.EOF Then
txtrate = Trim(rst!irate)
'txtcanno = rst!canno
End If
End Sub

Private Sub Command1_Click()
txtsname = ""
txtsno = ""
txtcanno = ""
txtamt = ""
txtqty = ""
Txttno = ""
Txttname = ""
Txtcno = ""
txtamount = ""
txtquan = ""
End Sub

Private Sub Command2_Click()
If Optsupplier.value = True Then
If txtsno = "" Then
MsgBox "Sno is empty please enter supplier number ", vbInformation
txtsno.SetFocus
Exit Sub
End If
If txtcanno = "" Then
MsgBox " please enter can number ", vbInformation
txtcanno.SetFocus
Exit Sub
End If
If Cboquality = "" Then
MsgBox " please select the quality ", vbInformation

Exit Sub
End If
If txtqty = "" Then
MsgBox " please enter the quantity ", vbInformation
txtqty.SetFocus
Exit Sub
End If
sql = ""
sql = "Insert Into d_Quality(sno,name, canno, rate, Quantity, Quality, Startdate, Enddate, amount,  auditid)" _
 & " Values ('" & txtsno & "','" & txtsname & "', '" & txtcanno & "', '" & txtrate & "', '" & txtqty & "', '" & Cboquality & "', '" & DTPsdate & "', '" & DTPedate & "', " & txtamt & ", '" & User & "')"
 oSaccoMaster.ExecuteThis (sql)
 
MsgBox "Record save succesfully", vbInformation
txtsname = ""
txtsno = ""
txtcanno = ""
txtamt = ""
txtqty = ""
End If

'Else
If Opttransporter.value = True Then
If Txttno = "" Then
MsgBox "tno is empty please enter transporter number ", vbInformation
Txttno.SetFocus
Exit Sub
End If
If Txtcno = "" Then
MsgBox " please enter can number ", vbInformation
Txtcno.SetFocus
Exit Sub
End If
If Cboqlity = "" Then
MsgBox " please select the quality ", vbInformation

Exit Sub
End If
If txtquan = "" Then
MsgBox " please enter the quantity ", vbInformation
txtqty.SetFocus
Exit Sub
End If
sql = ""
sql = "Insert Into d_Quality(sno,name, canno, rate, Quantity, Quality, Startdate, Enddate, amount,  auditid)" _
 & " Values ('" & Txttno & "','" & Txttname & "', '" & Txtcno & "', '" & txtrti & "', '" & txtquan & "', '" & Cboqlity & "', '" & DTPstartdate & "', '" & DTPenddate & "', " & txtamount & ", '" & User & "')"
 oSaccoMaster.ExecuteThis (sql)
 

MsgBox "Record save succesfully", vbInformation
Txttno = ""
Txttname = ""
Txtcno = ""
txtamount = ""
txtquan = ""
End If
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub optSupplier_Click()
If Optsupplier.value = True Then
Frmtransporterdetails.Visible = False
Frmdetails.Visible = True
End If
End Sub

Private Sub optTransporter_Click()

If Opttransporter.value = True Then
Frmdetails.Visible = False
Frmtransporterdetails.Visible = True
End If
End Sub

Private Sub txtqty_Change()
'txtqty = ""
End Sub

Private Sub txtqty_Click()
If txtqty <> "" Then
txtamt = CDbl(txtrate) * CDbl(txtqty)
End If
End Sub

Private Sub txtquan_Change()
If txtquan <> "" Then
txtamount = CDbl(txtrti) * CDbl(txtquan)
End If
End Sub

Private Sub txtquan_Click()
If txtquan <> "" Then
txtamount = CDbl(txtrti) * CDbl(txtquan)
End If
End Sub

Private Sub txtSNo_Change()
Set rst = New ADODB.Recordset
sql = "select * from d_suppliers where sno='" & txtsno & "'"
Set rst = oSaccoMaster.GetRecordset(sql)
If Not rst.EOF Then
txtsname = rst!names
txtcanno = IIf(IsNull(rst!canno), "", (rst!canno))
End If
End Sub

Private Sub Txttno_Change()
Set rst = New ADODB.Recordset
sql = "select * from d_Transporters where transcode='" & Txttno & "'"
Set rst = oSaccoMaster.GetRecordset(sql)
If Not rst.EOF Then
Txttname = rst!TransName
Txtcno = IIf(IsNull(rst!canno), "", rst!canno)
End If
End Sub

Private Sub Txttno_Click()
Set rst = New ADODB.Recordset
sql = "select * from d_Transporters where transcode='" & Txttno & "'"
Set rst = oSaccoMaster.GetRecordset(sql)
If Not rst.EOF Then
Txttname = rst!TransName
Txtcno = IIf(IsNull(rst!canno), "", rst!canno)
End If
End Sub
