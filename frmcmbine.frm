VERSION 5.00
Begin VB.Form frmcmbine 
   BackColor       =   &H8000000B&
   Caption         =   "PRINT ASSETS & LIABILITIES"
   ClientHeight    =   1875
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4305
   LinkTopic       =   "Form1"
   ScaleHeight     =   1875
   ScaleWidth      =   4305
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "CLOSE"
      Height          =   435
      Left            =   3240
      TabIndex        =   2
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Print Asset Balance Sheet"
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   480
      Width           =   2295
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Print Liabilities Balance Sheet"
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   1320
      Width           =   2295
   End
End
Attribute VB_Name = "frmcmbine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command3_Click()
    '//kimberbalancesheet
    reportname = "BalanceSheeetA.rpt"
    STRFORMULA = ""
    Show_Sales_Crystal_Report STRFORMULA, reportname, CompanyName

'IF YOU WANT TO DO A CSV FILE
'Command1_Click

End Sub

Private Sub Command5_Click()
    '//kimberbalancesheet
    reportname = "BalanceSheeetL.rpt"
    STRFORMULA = ""
    Show_Sales_Crystal_Report STRFORMULA, reportname, CompanyName


'IF YOU WANT TO DO A CSV FILE
'Command1_Click

End Sub
