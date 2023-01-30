VERSION 5.00
Begin VB.Form frmmilkdidprev 
   Caption         =   "MILKCONTROL REVERSAL"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cndrem 
      Caption         =   "Remove"
      Height          =   375
      Left            =   600
      TabIndex        =   7
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox txtdesc 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      TabIndex        =   5
      Top             =   720
      Width           =   2415
   End
   Begin VB.TextBox txtquantity 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox txtdcode 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lbldescrip 
      Caption         =   "Description"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   855
   End
   Begin VB.Label lblquantity 
      Caption         =   "Quantity"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label lbldcode 
      Caption         =   "Dcode"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmmilkdidprev"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdClose_Click()
Unload Me
End Sub
Private Sub Cndrem_Click()

sql = ""
sql = "set dateformat dmy delete from d_MilkControl  where dcode ='" & txtdcode & "'and DispQnty ='" & txtquantity & "' and DispDate ='" & frmMilkControl.DTPDispatchDate & "' "
'cn.Execute sql
oSaccoMaster.ExecuteThis (sql)
'sql = ""
'sql = "set dateformat dmy delete from d_DetailDispatch  where DCode ='" & txtdcode & "'and dispatch ='" & txtquantity & "'and Transdate ='" & frmMilkControl.DTPDispatchDate & "' "
''cn.Execute sql
'oSaccoMaster.ExecuteThis (sql)
sql = ""
sql = "set dateformat dmy UPDATE    d_dispatch  SET   dipping =dipping + '" & txtquantity & "',dispatch=dispatch-'" & txtquantity & "'  WHERE     (Transdate = '" & frmMilkControl.DTPDispatchDate & "')"
oSaccoMaster.ExecuteThis (sql)

'mysql = "set dateformat dmy Insert into Receiptno(Receiptno,Auditdate,auditid)values('" & txtRefNo & "','" & Format(Get_Server_Date, "dd/MM/yyyy") & "','" & User & "')"
'mysql = "set dateformat dmy delete from Receiptno where Receiptno,Auditdate,auditid)values('" & txtRefNo & "','" & Format(Get_Server_Date, "dd/MM/yyyy") & "','" & User & "')"
'oSaccoMaster.ExecuteThis (mysql)

sql = ""
sql = "set dateformat dmy select  DCode, DName,price from d_Debtors where (dcode ='" & frmmilkdidprev.txtdcode & "')"
Set rs = oSaccoMaster.GetRecordset(sql)
Price = rs!Price

sql = ""
sql = "set dateformat dmy delete gltransactions WHERE DocumentNo='Milk Sales'  and (Amount='" & (CCur(Price) * CCur(txtquantity)) & "') and (TransDate = '" & frmMilkControl.DTPDispatchDate & "')"
oSaccoMaster.ExecuteThis (sql)

MsgBox "Item successfully removed", vbInformation, "Stocks"
Unload Me
End Sub

