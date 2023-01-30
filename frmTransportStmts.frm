VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTransportStmts 
   Caption         =   "Transporters Statement"
   ClientHeight    =   3585
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   8640
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox ports 
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "frmTransportStmts.frx":0000
      Left            =   3960
      List            =   "frmTransportStmts.frx":0010
      TabIndex        =   13
      Text            =   "\\127.0.0.1\E-PoS 80mm Thermal Printer"
      Top             =   480
      Width           =   4575
   End
   Begin VB.CheckBox chprint 
      Caption         =   "Use LPT1 Printer"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6120
      TabIndex        =   12
      Top             =   120
      Width           =   2295
   End
   Begin VB.TextBox txtTCode 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1920
      TabIndex        =   6
      Top             =   480
      Width           =   735
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Transporter's Statements"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   1815
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   5775
      Begin VB.OptionButton optAdvanceSlip 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "Print Advance Slip"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   480
         TabIndex        =   11
         Top             =   1440
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.OptionButton OptNormalStmt 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "Normal Statement (Use POS Printer)"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   480
         TabIndex        =   5
         Top             =   360
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.OptionButton OptDetailedStmt 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "Detailed Statement (Use POS Printer)"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   480
         TabIndex        =   4
         Top             =   720
         Value           =   -1  'True
         Width           =   4095
      End
      Begin VB.OptionButton OptNormalA4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "Normal Statement (Use Normal Printer (A4))"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   480
         TabIndex        =   3
         Top             =   1080
         Visible         =   0   'False
         Width           =   4095
      End
   End
   Begin VB.CommandButton cmdPrint 
      Appearance      =   0  'Flat
      Caption         =   "Print"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1920
      TabIndex        =   1
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton cmClose 
      Appearance      =   0  'Flat
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2640
      TabIndex        =   0
      Top             =   3240
      Width           =   735
   End
   Begin MSComCtl2.DTPicker DTPStmts 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   7
      Top             =   840
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
      _Version        =   393216
      Format          =   119865345
      CurrentDate     =   40109
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      Caption         =   "Transport Code :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Width           =   1560
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      Caption         =   "End of Period :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Width           =   1365
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      Caption         =   "Enter Transporter code and select end of period to print statement"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   5970
   End
End
Attribute VB_Name = "frmTransportStmts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Enddate As Date
Dim Startdate As Date

Private Sub chprint_Click()
ports.Clear
ports = ""
'//If the drivers are installed it won't matter whether the Port is indicated
' or not it will just work.

If chprint.value = vbChecked Then
ports.AddItem "LPT1"
ports = "LPT1"
ports.AddItem "LPT2"
ports.AddItem "LPT3"
ports.AddItem "LPT4"
ports.AddItem "LPT5"
Else
'Share the printer first the use of 127.0.0.1 which is
'standard IP address for a loopback network connection
'instead of getting the computer name or IP Address
'
Dim prnPrinter As Printer
Dim pr As String
ports.Clear

For Each prnPrinter In Printers
   If InStr(prnPrinter.DeviceName, "\\") Then
    ports.AddItem prnPrinter.DeviceName
    If InStr(prnPrinter.DeviceName, "G") Then
    ports.Text = prnPrinter.DeviceName
    End If
    Else
    ports.AddItem "\\127.0.0.1\" & prnPrinter.DeviceName
    If InStr(prnPrinter.DeviceName, "G") Then
    ports.Text = "\\127.0.0.1\" & prnPrinter.DeviceName
    End If
    End If
   
   
Next
End If
'This code will work only if there is a connection e.g LAN or modem.
'It is not a must that it is an internet connection because
'computer's network interface card has to be functional
End Sub

Private Sub cmClose_Click()
Unload Me
End Sub
Private Sub updatetransporter()
    '''UPDATE FOR THOSE HAVING TWO PRICES ON THE SAME MONTHE
    Startdate = DateSerial(Year(DTPStmts), month(DTPStmts), 1)
    Enddate = DateSerial(Year(DTPStmts), month(DTPStmts) + 1, 1 - 1) 'txtTCode
        sql = "set dateformat dmy SELECT Sno  FROM d_Transport  WHERE Trans_Code='" & txtTCode & "' AND (Active = 0) AND DateInactivate BETWEEN '" & Format(Startdate, "dd/mm/yyyy") & "' AND '" & Format(Enddate, "dd/mm/yyyy") & "' order by sno asc"
        Set rschecking = oSaccoMaster.GetRecordset(sql)
        While Not rschecking.EOF
        DoEvents
        Dim kgamount, qnt, qntamount, qntkgs As Double

        kgamount = 0
        qnt = 0
        qntamount = 0
        qntkgs = 0
            sql = "set dateformat dmy select Sno,Rate,startdate from d_Transport  WHERE Sno='" & rschecking.Fields(0) & "' and Trans_Code='" & txtTCode & "' AND (Active = 1) AND startdate BETWEEN '" & Format(Startdate, "dd/mm/yyyy") & "' AND '" & Format(Enddate, "dd/mm/yyyy") & "'"
            Set rsconfirm = oSaccoMaster.GetRecordset(sql)
            If Not rsconfirm.EOF Then
              'sql = "SET DATEFORMAT DMY SELECT isnull(SUM(QSupplied),0) AS Kgs from d_Milkintake WHERE Sno='" & rsconfirm!sno & "' and TransDate BETWEEN '" & Format(rsconfirm!Startdate, "dd/mm/yyyy") & "' AND '" & Format(Enddate, "dd/mm/yyyy") & "'"
              Set rsp = oSaccoMaster.GetRecordset("d_sp_UpdateTkgs '" & rsconfirm!sno & "','" & Format(rsconfirm!Startdate, "dd/mm/yyyy") & "', '" & Format(Enddate, "dd/mm/yyyy") & "'")
              If rsp!kgs > 0 Then
                 kgamount = rsp!kgs * rsconfirm!rate
                sql = "SET DATEFORMAT DMY SELECT QNTY, Amount from d_TransDetailed where Sno='" & rsconfirm!sno & "' and EndPeriod BETWEEN '" & Format(Startdate, "dd/mm/yyyy") & "' AND '" & Format(Enddate, "dd/mm/yyyy") & "'"
                Set rsz = oSaccoMaster.GetRecordset(sql)
                If Not rsz.EOF Then
                   qnt = rsz!qnty + rsp!kgs
                   qntamount = rsz!amount + kgamount
                   qntkgs = 0
                  oSaccoMaster.ExecuteThis ("d_sp_UpdateT '" & qnt & "','" & qntamount & "','" & rsconfirm!sno & "','" & Format(Startdate, "dd/mm/yyyy") & "','" & Format(Enddate, "dd/mm/yyyy") & "'")
                End If
              End If
            End If
        rschecking.MoveNext
        Wend
    '''END
End Sub

Private Sub cmdprint_Click()
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
    Dim CummulKgs As Double
     
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
       
If txtTCode = "" Then
    MsgBox "Please enter Transporter code.", vbCritical
        txtTCode.SetFocus
    Exit Sub
End If

Startdate = DateSerial(Year(DTPStmts), month(DTPStmts), 1)
Enddate = DateSerial(Year(DTPStmts), month(DTPStmts) + 1, 1 - 1)

If optAdvanceSlip.value = True Then
'--Net amount as at date
    Set rst1 = New ADODB.Recordset
        sql = "d_sp_PrintTransStmt '" & txtTCode & "','" & DTPStmts & "'"
    Set rst1 = oSaccoMaster.GetRecordset(sql)
    
    If rst1.EOF Then
    MsgBox "There is no current record."
        txtTCode.SetFocus
        Exit Sub
    End If
    
        ttt = "\\127.0.0.1\E-PoS 80mm Thermal Printer" 'LPT1,LPT2....
        Set fso = CreateObject("Scripting.FileSystemObject")
        On Error GoTo err
        
        'Set chkPrinter = fso.GetFile(ttt)
       
        
    Set txtFile = fso.CreateTextFile(ttt, True)
    txtFile.WriteLine escAlignCenter
    txtFile.WriteLine "Advance Slip"
    txtFile.WriteLine "" & cname & ""
    txtFile.WriteLine "........................................"
    txtFile.WriteLine escAlignLeft
    txtFile.WriteLine "Trans Code : " & txtTCode
    txtFile.WriteLine "Names : " & rst1.Fields("TransName")
    txtFile.WriteLine "Issue Items/Services worth not more than"
    Dim rsg As New ADODB.Recordset
    Set rsg = oSaccoMaster.GetRecordset("SET  dateformat dmy  SELECT     SUM(Amount) + SUM(Subsidy) AS Gross   FROM  d_TransDetailed  WHERE     (Trans_Code = '" & txtTCode & "') AND (EndPeriod = '" & DTPStmts & "')")
    Dim Gro As Currency
    If Not IsNull(rsg.Fields(0)) Then Gro = CCur(rsg.Fields(0))
    Set rsg = oSaccoMaster.GetRecordset("SET dateformat dmy SELECT SUM(amount) FROM d_Transport_Deduc   WHERE     tdate_deduc BETWEEN '" & Startdate & "' AND '" & Enddate & "' AND transcode = '" & txtTCode & "'")
    Dim Ded As Currency
    If Not IsNull(rsg.Fields(0)) Then Ded = CCur(rsg.Fields(0))
    
    txtFile.WriteLine "Kshs. : " & Format((Gro - Ded), "#,##0.00") & ""
    txtFile.WriteLine "Sign"
    txtFile.WriteLine "___________________________"
    txtFile.WriteLine UCase(username)
    txtFile.Write "Date " & Format(Get_Server_Date, "dd/mm/yyyy")
    txtFile.WriteLine ", Time : " & Format(Get_Server_Date, "HH:MM:SS")
    txtFile.WriteLine "........................................"
    txtFile.WriteLine escFeedAndCut
    
    
    
End If

If OptNormalStmt.value = True Then
    Set rst = New ADODB.Recordset
        sql = "d_sp_PrintTransStmt '" & txtTCode & "','" & Enddate & "'"
    Set rst = oSaccoMaster.GetRecordset(sql)
If rst.EOF Then
    MsgBox "There is no record in the payroll for Transporter with code " & txtTCode, vbInformation
        txtTCode.SetFocus
    Exit Sub
End If



        ttt = "\\127.0.0.1\E-PoS 80mm Thermal Printer" 'LPT1,LPT2....
        Set fso = CreateObject("Scripting.FileSystemObject")
        On Error GoTo err
        
'        Set chkPrinter = fso.GetFile(ttt)
       
        
        Set txtFile = fso.CreateTextFile(ttt, True)
    txtFile.WriteLine escAlignCenter
    txtFile.WriteLine "" & cname & ""
    txtFile.WriteLine "" & paddress & ""
    txtFile.WriteLine "" & town & ""
    txtFile.WriteLine "TRANSPORTER STATEMENT FOR " & UCase(Format(DTPStmts, "MMMM/YYYY"))
    txtFile.WriteLine escAlignLeft
    txtFile.WriteLine "........................................"
    txtFile.WriteLine "Trans Code :" & txtTCode
    txtFile.WriteLine "Tans Name :" & rst!TransName
    txtFile.WriteLine "........................................"
    txtFile.WriteLine "Total Kgs Delivered :" & Format(rst!QntySup, "#,##0.00" & " Kgs")
    txtFile.WriteLine "Gross Amount               Kshs: " & Format(rst!Amnt, "#,##0.00") & ""
    txtFile.WriteLine "Subsidy                    Kshs: " & Format(rst!subsidy, "#,##0.00") & ""
    txtFile.WriteLine "Net Gross                  Kshs: " & Format(rst!GrossPay, "#,##0.00") & ""
   
    txtFile.Write escBoldOn
    txtFile.WriteLine "DEDUCTIONS"
    txtFile.WriteLine "-------------"
    txtFile.Write escBoldOff
    
    txtFile.WriteLine "Agrovet          Kshs: " & Format(rst!agrovet, "#,##0.00") & ""
    txtFile.WriteLine "TM Shares        Kshs: " & Format(rst!TMShares, "#,##0.00") & ""
    txtFile.WriteLine "H Shares         Kshs: " & Format(rst!HShares, "#,##0.00") & ""
    txtFile.WriteLine "Advances         Kshs: " & Format(rst!Advance, "#,##0.00") & ""
    txtFile.WriteLine "FSA              Kshs: " & Format(rst!FSA, "#,##0.00") & ""
    txtFile.WriteLine "AI               Kshs: " & Format(rst!AI, "#,##0.00") & ""
    txtFile.WriteLine "Others           Kshs: " & Format(rst!Others, "#,##0.00") & ""
    txtFile.WriteLine "Total Deductions Kshs: " & Format(rst!Totaldeductions, "#,##0.00") & ""
    txtFile.WriteLine "........................................"
    txtFile.WriteLine "NET PAY                    Kshs: " & Format((rst!NetPay), "#,##0.00") & ""
    txtFile.WriteLine "........................................"
    txtFile.WriteLine "BANK DETAILS"
    txtFile.WriteLine "-------------"
    txtFile.WriteLine "Bank Name :" & rst!BankName & ""
    txtFile.WriteLine "Bank Branch :" & rst!Branch
    txtFile.WriteLine "Account Number :" & rst!ACCNO

    txtFile.WriteLine "---------------------------------------"
    txtFile.WriteLine "        Date :" & Format(Get_Server_Date, "dd/mm/yyyy HH:MM:SS AM/PM")
    txtFile.WriteLine "         " & motto & ""
    txtFile.WriteLine "---------------------------------------"
    txtFile.WriteLine escFeedAndCut
    txtFile.Close
'End If
Exit Sub
err: MsgBox err.description & " or There is no printer connected."
Exit Sub
End If

'--Print detailed statement
If OptDetailedStmt.value = True Then

    
  ttt = "\\127.0.0.1\E-PoS 80mm Thermal Printer" 'LPT1,LPT2....
        Set fso = CreateObject("Scripting.FileSystemObject")
        On Error GoTo err
        
 '       Set chkPrinter = fso.GetFile(ttt)
       
        
        Set txtFile = fso.CreateTextFile(ttt, True)
    txtFile.WriteLine escAlignCenter
    txtFile.WriteLine "" & cname & ""
    txtFile.WriteLine "" & paddress & ""
    txtFile.WriteLine "" & town & ""
    txtFile.WriteLine "TRANSPORT DETAILED STATEMENT FOR " & UCase(Format(DTPStmts, "MMMM/YYYY"))
    txtFile.WriteLine escAlignLeft
    txtFile.WriteLine "........................................"
    txtFile.WriteLine "Trans Code :" & UCase(txtTCode)
    '---d_sp_TransDetStmt @Code varchar(35), @EPeriod varchar(12)
    Set rs = New ADODB.Recordset
    'sql = "d_sp_TransDetStmt '" & txtTCode & "','" & Enddate & "'"
    'sql = "  set dateformat dmy  SELECT  dbo.d_TransDetailed.Trans_Code, dbo.d_Transporters.TransName, dbo.d_TransDetailed.SNo, dbo.d_TransDetailed.QNTY, dbo.d_TransDetailed.Amount , dbo.d_TransDetailed.subsidy From dbo.d_TransDetailed, dbo.d_Transporters WHERE     dbo.d_TransDetailed.Trans_Code = dbo.d_Transporters.TransCode AND   (dbo.d_TransDetailed.Trans_Code ='" & txtTCode & "' ) AND (dbo.d_TransDetailed.EndPeriod ='" & Enddate & "')"
'    sql = "  set dateformat dmy  SELECT  dbo.d_TransDetailed.Trans_Code, dbo.d_Transporters.TransName,dbo.d_Transport.sno,dbo.d_Transport.Active, dbo.d_TransDetailed.SNo, dbo.d_TransDetailed.QNTY, dbo.d_TransDetailed.Amount , dbo.d_TransDetailed.subsidy From dbo.d_TransDetailed, dbo.d_Transporters,d_Transport WHERE     dbo.d_TransDetailed.Trans_Code = dbo.d_Transporters.TransCode and dbo.d_Transporters.TransCode=dbo.d_Transport.Trans_Code AND   (dbo.d_TransDetailed.Trans_Code ='" & txtTCode & "' ) AND dbo.d_Transport.Active=0 and(dbo.d_TransDetailed.EndPeriod ='" & Enddate & "')"
'    Set rs = oSaccoMaster.GetRecordset(sql)
     oSaccoMaster.ExecuteThis ("d_sp_TransUpdate '" & Startdate & "','" & Enddate & "','" & User & "','" & txtTCode & "'")
     updatetransporter
     
    sql = "  set dateformat dmy  SELECT  dbo.d_TransDetailed.Trans_Code, dbo.d_Transporters.TransName, dbo.d_TransDetailed.SNo, dbo.d_TransDetailed.QNTY, dbo.d_TransDetailed.Amount , dbo.d_TransDetailed.subsidy From dbo.d_TransDetailed, dbo.d_Transporters WHERE     dbo.d_TransDetailed.Trans_Code = dbo.d_Transporters.TransCode AND   (dbo.d_TransDetailed.Trans_Code ='" & txtTCode & "' ) AND (dbo.d_TransDetailed.EndPeriod ='" & Enddate & "')"
    Set rs = oSaccoMaster.GetRecordset(sql)
    txtFile.WriteLine "Name :" & rs.Fields(1)
    txtFile.WriteLine "........................................"
    txtFile.WriteLine "SNo " & vbTab & "SUBSIDY" & vbTab & "QNTY" & vbTab & "PAYABLE"
    txtFile.WriteLine "........................................"
    
    
         
        While Not rs.EOF
        'MsgBox rs!QSupplied
        txtFile.WriteLine rs.Fields(2) & " " & vbTab & " " & Format(rs.Fields(5), "#,##0.00") & " " & vbTab & " " & Format(rs.Fields(3), "#,##0.0#") & " " & vbTab & " " & Format(rs.Fields(4), "#,##0.00") & " "
        'txtfile.WriteLine rs!PPU
         rs.MoveNext
        
        Wend
    
    
     
        
    Set rst = New ADODB.Recordset
        'sql = "d_sp_PrintTransStmt '" & txtTCode & "','" & Enddate & "'"
        sql = "SET              dateformat dmy SELECT     dbo.d_TransportersPayRoll.Code, dbo.d_Transporters.TransName, dbo.d_TransportersPayRoll.QntySup, dbo.d_TransportersPayRoll.Amnt, dbo.d_TransportersPayRoll.Subsidy, dbo.d_TransportersPayRoll.GrossPay,  dbo.d_TransportersPayRoll.Agrovet, dbo.d_TransportersPayRoll.TMShares, dbo.d_TransportersPayRoll.FSA, dbo.d_TransportersPayRoll.HShares, dbo.d_TransportersPayRoll.Advance, dbo.d_TransportersPayRoll.AI,dbo.d_TransportersPayRoll.Others, dbo.d_TransportersPayRoll.Totaldeductions, dbo.d_TransportersPayRoll.NetPay,dbo.d_TransportersPayRoll.BankName , dbo.d_TransportersPayRoll.Branch, dbo.d_TransportersPayRoll.accno   FROM dbo.d_Transporters INNER JOIN dbo.d_TransportersPayRoll ON dbo.d_Transporters.TransCode = dbo.d_TransportersPayRoll.Code   WHERE     (dbo.d_TransportersPayRoll.Code = '" & txtTCode & "') AND (dbo.d_TransportersPayRoll.EndPeriod = '" & Enddate & "')"
    Set rst = oSaccoMaster.GetRecordset(sql)
   txtFile.WriteLine "........................................"
   txtFile.WriteLine "TOTAL Kgs : " & Format(rst!QntySup, "#,##0.00 ") & vbNewLine & "Amount " & vbTab & "" & vbTab & "Kshs " & Format(rst!Amnt, "#,##0.00") & " " & vbNewLine & "Subsidy " & vbTab & "Kshs " & Format(rst!subsidy, "#,##0.00") & vbNewLine & "Gross Amount " & vbTab & "Kshs " & Format((rst!subsidy + rst!Amnt), "#,##0.00")
   txtFile.WriteLine "........................................"
    txtFile.WriteLine escBoldOn
    txtFile.WriteLine "DEDUCTIONS"
    txtFile.WriteLine "........................................"
    txtFile.WriteLine escBoldOff
    Dim Gross As Currency
    Gross = rst!subsidy + rst!Amnt
    
  '--d_sp_PrintTransDeduc @Code varchar(35), @SDate varchar(12), @EDate varchar(12) AS
Set rst = New ADODB.Recordset
sql = "d_sp_PrintTransDeduc " & txtTCode & ",'" & Startdate & "','" & Enddate & "'"
Set rst = oSaccoMaster.GetRecordset(sql)
    

    txtFile.WriteLine "........................................"
    txtFile.WriteLine "DATE " & vbTab & "" & vbTab & "AMOUNT" & vbTab & "DESCRIPTION"
    txtFile.WriteLine "........................................"

    While Not rst.EOF
        'MsgBox rs!QSupplied
        txtFile.WriteLine rst.Fields(0) & " " & vbTab & " " & Format(rst.Fields(2), "#,##0.00" & vbTab & " " & rst.Fields(1) & " ")
        'txtfile.WriteLine rs!PPU
         rst.MoveNext
        
        Wend
    Set rst1 = New ADODB.Recordset
        sql = "SET dateformat dmy SELECT SUM(Amount) AS TOTAL From d_Transport_Deduc WHERE     (TransCode = '" & txtTCode & "') AND (TDate_Deduc BETWEEN '" & Startdate & "' AND '" & Enddate & "')"
    Set rst1 = oSaccoMaster.GetRecordset(sql)
    txtFile.WriteLine "........................................"
    If Not IsNull(rst1!total) Then
    txtFile.WriteLine "Total Deductions Kshs: " & Format(rst1!total, "#,##0.00") & ""
    Else
    txtFile.WriteLine "Total Deductions Kshs: 0.00"
    End If
    txtFile.WriteLine "........................................"
    If Not IsNull(rst1!total) Then
    txtFile.WriteLine "NET PAY                   Kshs :" & Format(Gross - rst1!total, "#,##0.00") & ""
    Else
    txtFile.WriteLine "NET PAY                   Kshs :" & Format(Gross, "#,##0.00") & ""
    End If
    txtFile.WriteLine "-----------------------------------------"
    txtFile.WriteLine "BANK DETAILS"
    txtFile.WriteLine "-------------"
'    txtfile.WriteLine "Bank Name :" & Rst1!Bank & ""
'    txtfile.WriteLine "Bank Branch :" & Rst1!BBranch
'    txtfile.WriteLine "Account Number :" & Rst1!accountnumber
    txtFile.WriteLine "---------------------------------------"
    txtFile.WriteLine "        Date :" & Format(Get_Server_Date, "dd/mm/yyyy HH:MM:SS AM/PM")
    txtFile.WriteLine "         " & motto & ""
    txtFile.WriteLine "---------------------------------------"
    txtFile.WriteLine escFeedAndCut
    txtFile.Close
    End If
    
    If OptNormalA4.value = True Then
    reportname = "d_StmtA4.rpt"
    STRFORMULA = "{d_Payroll.SNo}= " & txtTCode & " and {d_Payroll.MMonth }=" & month(DTPStmts) & " AND {d_Payroll.YYear} =" & Year(DTPStmts)
    Show_Sales_Crystal_Report STRFORMULA, reportname, ""

'    d_StmtA4
    End If
    txtTCode = ""
End Sub

Private Sub DTPStmts_Validate(Cancel As Boolean)
DTPStmts = DateSerial(Year(DTPStmts), month(DTPStmts) + 1, 1 - 1)
End Sub

Private Sub Form_Load()
DTPStmts = Format(Get_Server_Date, "dd/mm/yyyy")
DTPStmts = DateSerial(Year(DTPStmts), month(DTPStmts) + 1, 1 - 1)
End Sub
