Attribute VB_Name = "Decla"
Option Explicit
Public strFileName  As String
Public rs1 As Recordset
Public rs2 As Recordset
Public rs3 As Recordset
Public rs4 As Recordset
Public bcode  As String
Public rs5 As Recordset
Public res As String
Public Mach As String
Public transno As String
'Public rsNumbers As String
Public I As Integer
Public SuspenseAcc As String
Public mMemberNo As String
Public action As String
Public power As Long
Public cname As String
'Public STRFORMULA As String
Public reportname As String
Public Rst1 As New ADODB.Recordset
Public Title As String
Public SelectedDsn As String
Public ErrorMessage As String
Public rsNumbers() As String
Public rscompany As Recordset
Public rsPayslip As Recordset
'Public a As CRAXDRT.Application
'Public r As CRAXDRT.Report
Public clsClass As Object
Public CompanyName As String, CompanyPhone As String, CompanyTown As String, CompanyTagLine As String
Public STRFORMULA As String
Public li As ListItem
Dim mvarConnection As String
Public oldconnection As String
Public co As String
Public NewRecord As Boolean
Public MyBookMark
Public SearchValue As String
Public SearchValue1 As String
Public cnn
Public continue As Boolean
Public Editing As Boolean
Public oSaccoMaster As New CSaccoData 'reference to the class
Public Provider As String
Public sql As String
Public mysql As String
'Public a As CRAXDRT.Application
'Public r As CRAXDRT.Report
Public DSource As String
Public sel As String
Public RCanc As Boolean
Public FMonth As Long
Public TMonth As Long
Public ByPass As Boolean
Public MDBase As String
Public Direct As Boolean
Public Const Cfmt = "###,###,###,###,##0.00;(0.00)"
Dim path As String
Dim myclass As Object
'Dim reportname As String
'Dim CSecurity As CConnect
Public DelimiterConstant As Integer
Public rs As Recordset
Public rsEmployee As Recordset
Public rst As Recordset
