VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmmilkinvoice 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      Height          =   255
      Left            =   960
      TabIndex        =   4
      Top             =   2520
      Width           =   1575
   End
   Begin VB.ComboBox cbodcode 
      Height          =   315
      Left            =   720
      TabIndex        =   3
      Text            =   "D002"
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton cmdprintinvoice 
      Caption         =   "Print Invoice"
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   1680
      Width           =   2415
   End
   Begin MSComCtl2.DTPicker dtpenddate 
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   95289345
      CurrentDate     =   41771
   End
   Begin MSComCtl2.DTPicker dtpstartdate 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   95289345
      CurrentDate     =   41771
   End
End
Attribute VB_Name = "frmmilkinvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdprintinvoice_Click()
If cbodcode.Text = "" Then
  MsgBox "Select the Debtors Code first", vbInformation, Me.Caption
  cbodcode.SetFocus
 Exit Sub
End If
Set rs = oSaccoMaster.GetRecordset("Truncate table PrintMilkInvoice")

sql = "Set dateformat dmy  Select Dcode,isnull(SUM(DispQnty),0),isnull(SUM(DispQnty * price),0) from d_MilkControl where DispDate>='" & dtpstartdate & "' and DispDate<='" & dtpenddate & "' and dcode='" & cbodcode.Text & "' group by Dcode"
     Set rs2 = oSaccoMaster.GetRecordset(sql)
     If Not rs2.EOF Then
     sql = "set dateformat dmy  Insert into PrintMilkinvoice (DCode,DispQnty,Total,StartDate,EndDate) Values"
     sql = sql & "('" & cbodcode & "'," & rs2.Fields(1) & "," & rs2.Fields(2) & ",'" & dtpstartdate & "','" & dtpenddate & "')"
     Set rs = oSaccoMaster.GetRecordset(sql)
     End If

    reportname = "printmilkinvoice.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, title
     
     
End Sub

Private Sub Form_Load()
transdate = Format(Get_Server_Date, "DD/MM/YYYY")
dtpstartdate = DateSerial(year(transdate), month(transdate), 1)
dtpenddate = DateSerial(year(transdate), month(transdate) + 1, -1)

cbodcode.Clear
sql = "select dcode from d_Debtors"
Set rs = oSaccoMaster.GetRecordset(sql)
 With rs
  While Not .EOF
   cbodcode.AddItem (rs.Fields(0))
   .MoveNext
   Wend
 End With
End Sub
