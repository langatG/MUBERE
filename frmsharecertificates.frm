VERSION 5.00
Begin VB.Form frmsharecertificates 
   Caption         =   "SHARE CERTIFICATES"
   ClientHeight    =   4260
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7575
   Icon            =   "frmsharecertificates.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   7575
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdprint 
      Caption         =   "Print"
      Height          =   375
      Left            =   5520
      TabIndex        =   12
      Top             =   3720
      Width           =   1335
   End
   Begin VB.ComboBox CBOYEAR 
      Height          =   315
      ItemData        =   "frmsharecertificates.frx":014A
      Left            =   2640
      List            =   "frmsharecertificates.frx":0166
      TabIndex        =   10
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox txtpremium 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2640
      TabIndex        =   8
      Top             =   1800
      Width           =   3135
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmsharecertificates.frx":019A
      Left            =   2640
      List            =   "frmsharecertificates.frx":01A4
      TabIndex        =   7
      Top             =   2400
      Width           =   2655
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      Height          =   435
      Left            =   3720
      TabIndex        =   5
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdgeneratecerts 
      Caption         =   "Generate Certs"
      Height          =   495
      Left            =   1920
      TabIndex        =   4
      Top             =   3720
      Width           =   1335
   End
   Begin VB.TextBox txtpervalue 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   1200
      Width           =   3135
   End
   Begin VB.TextBox txtdivided 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   480
      Width           =   3135
   End
   Begin VB.Label Label5 
      Caption         =   "Year"
      Height          =   255
      Left            =   600
      TabIndex        =   11
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Premium:"
      Height          =   255
      Left            =   480
      TabIndex        =   9
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Type"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Per Value:"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Divided Into:"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   2055
   End
End
Attribute VB_Name = "frmsharecertificates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdgeneratecerts_Click()
Dim pervalue As Currency
Dim Shares As Double
Dim sno As String
Dim bal As Double
Dim initial As Double
sql = ""
sql = "truncate table sharescerts"
oSaccoMaster.ExecuteThis (sql)
'SELECT     sno, shares, pervalue, premium, div, yyear   FROM         sharescerts
sql = ""
sql = "SELECT     sno,bal,AMNT,premium,code     FROM         d_Shares where sno='" & frmsharestransactions.txtto & "' "
Set rs = oSaccoMaster.GetRecordset(sql)
While Not rs.EOF
sno = IIf(IsNull(rs.Fields(0)), rs.Fields(4), rs.Fields(0))
'Shares=
'If sno = "2921" Then
'MsgBox ""
'End If
bal = rs.Fields(1)
'initial = rs.Fields(2) / CDbl(txtpervalue)
Shares = (rs.Fields(2)) / CDbl(rs.Fields(3))
sql = "select * from sharescerts where sno='" & sno & "'"
Set Rst = oSaccoMaster.GetRecordset(sql)
If Rst.EOF Then
sql = " set dateformat dmy INSERT INTO sharescerts"
sql = sql & "               (sno, shares, pervalue, premium, div, yyear)"
sql = sql & "VALUES     ('" & sno & "'," & (Shares) & "," & txtpervalue & "," & rs.Fields(2) & ",0," & CBOYEAR & ")"
oSaccoMaster.ExecuteThis (sql)
Else
sql = "update sharescerts set shares=shares+" & Shares & " where sno='" & sno & "'"
oSaccoMaster.ExecuteThis (sql)
End If
rs.MoveNext
Wend

MsgBox "Certs successfully generated", vbInformation

End Sub

Private Sub cmdprint_Click()
'sharescert.rpt

 reportname = "sharescert.rpt"
 Show_Sales_Crystal_Report "", reportname, ""
 
End Sub
