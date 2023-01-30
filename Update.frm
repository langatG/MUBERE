VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmUpdate 
   BackColor       =   &H80000000&
   Caption         =   "Update "
   ClientHeight    =   3360
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6825
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Update.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   6825
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdADest 
      Caption         =   "..."
      Height          =   300
      Left            =   6360
      TabIndex        =   15
      Top             =   2175
      Width           =   345
   End
   Begin VB.TextBox txtADestHeader 
      Height          =   315
      Left            =   0
      TabIndex        =   14
      Top             =   2160
      Width           =   6330
   End
   Begin MSComCtl2.DTPicker DTPto 
      Height          =   315
      Left            =   2355
      TabIndex        =   13
      Top             =   840
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   " dd-MM-yyyy"
      Format          =   130416643
      CurrentDate     =   38251
   End
   Begin MSComCtl2.DTPicker DTPfrom 
      Height          =   315
      Left            =   2355
      TabIndex        =   12
      Top             =   375
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   " dd-MM-yyyy"
      Format          =   130416643
      CurrentDate     =   38251
   End
   Begin MSComDlg.CommonDialog DLG 
      Left            =   5685
      Top             =   1500
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdwrite 
      Caption         =   " Write To Disk"
      Height          =   360
      Left            =   1125
      TabIndex        =   11
      Top             =   2880
      Width           =   1440
   End
   Begin VB.CommandButton cmdcub 
      Appearance      =   0  'Flat
      Caption         =   "&Cub"
      Height          =   360
      Left            =   825
      TabIndex        =   9
      Top             =   4200
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.OptionButton Optdebit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Debit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   -600
      Width           =   1575
   End
   Begin VB.OptionButton Optcredits 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Credit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   600
      TabIndex        =   7
      Top             =   -480
      Width           =   1575
   End
   Begin VB.TextBox txtto 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   720
      TabIndex        =   4
      Top             =   -600
      Width           =   2415
   End
   Begin VB.TextBox txtfrom 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2160
      TabIndex        =   3
      Top             =   -1200
      Width           =   2415
   End
   Begin MSComctlLib.ProgressBar Pb 
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Visible         =   0   'False
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Enabled         =   0   'False
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdclose 
      Appearance      =   0  'Flat
      Caption         =   "&Close"
      Height          =   360
      Left            =   2775
      TabIndex        =   1
      Top             =   2880
      Width           =   1305
   End
   Begin VB.CommandButton cmdupdate 
      Appearance      =   0  'Flat
      Caption         =   "Update"
      Height          =   360
      Left            =   1890
      TabIndex        =   0
      Top             =   4200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1560
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Finish Date:"
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
      Left            =   1290
      TabIndex        =   6
      Top             =   885
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Start Date:"
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
      Left            =   1395
      TabIndex        =   5
      Top             =   420
      Width           =   840
   End
End
Attribute VB_Name = "frmUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim bb As String
Dim maxRec As Integer
Dim sql As String
Dim teller As String
Public br As String
Private Type acInfo
    ACCNO As String
    custBal As String
    AccName As String
    custno As String
    Amnt As Currency
    transdes As String
    transdate As Date
    chequeno As String
    Commission As Currency
    transtype As Variant
    vno As String
    audid As String
    audittime As Date
    id As String
End Type
Dim rs As Object
Dim cn As Object

Dim ctdata As acInfo
Dim myclass As New cdbase


Dim Provider As String

Private Sub cmdADest_Click()
'txtADestHeader
On Error GoTo SysError
    With DLG
        .ShowOpen
        If .FileName <> "" Then
            txtADestHeader = .FileName
            txtADestHeader = StrReverse(txtADestHeader)
            If InStr(1, txtADestHeader, "\", vbTextCompare) > 0 Then
                txtADestHeader = StrReverse(Right(txtADestHeader, Len(txtADestHeader) - InStr(1, txtADestHeader, "\", vbTextCompare)))
            End If
        End If
    End With
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub cmdclose_Click()
Unload Me
End Sub
'Private Sub getdata(data() As acInfo)
'Dim myindex As Integer
' Dim tempRs As Object
'    Dim myrec As Object, X As Integer
'    maxrec = 0
'
'    Set MyClass = New cdbase
'    Set cn = CreateObject("adodb.connection")
'
'    If Provider = "" Then Provider = MyClass.OpenCon
'
'    cn.Open = Provider
'    Set myrec = CreateObject("adodb.recordset")
'
'
'
'            sql = "SELECT DISTINCT customerno,accname,amount,accno,transdescription,"
'            sql = sql & " transdate,commission,chequeno,transtype from cashtransactions1 where accno >='" & txtfrom & "' and accno < = '" & txtto & "'"
'            myrec.Open sql, cn
'
'
'            With myrec
'                While Not .EOF
'                    ctdata().accNo = !accNo & ""
'                    ctdata().accName = !accName & ""
'                    ctdata().custNo = !customerno & ""
'                    ctdata().amnt = !amount
'                    ctdata().transdes = !transdescription
'                    ctdata().transdate = !transdate
'                    ctdata().commission = !commission
'                    ctdata().Chequeno = !Chequeno
'                    ctdata().transtype = !transtype
'                    ctdata().custbal = getCurBal(ctdata().accNo)
'
'                    ReDim Preserve ctdata(myindex + 1)
'                    myindex = myindex + 1
'                    .MoveNext
'                Wend
'            End With
'End Sub

Private Sub cmdcub_Click()
Dim myrec As Object
        Dim myclass As cdbase
        Set myclass = New cdbase
        Set myrec = CreateObject("adodb.recordset")
        Set cn = CreateObject("adodb.connection")
        Provider = myclass.OpenCon
        cn.Open Provider, "atm", "atm"
        Dim TELL As String
        Dim avail As Currency
        Dim cust As Variant
        Dim cub As Object
        Dim com As Currency
        Dim amount As Currency
        Dim trn As String
        Dim v As String
        Dim tdate As Date
        Dim chequeno As String
        Dim ttype As String
        Dim number As Integer
        Dim xxxx
        TELL = GetSetting("FOSAdll", "Teller", "Name")
    sql = "select distinct accNO from customerbalance  where accno >='" & txtFrom & "' and  accno <='" & txtTo & "' order by accno asc"
    myrec.Open sql, cn
     Pb.Min = 0

    While Not myrec.EOF
    
                  Pb.max = number + 1
                  Pb.Visible = True
    
                          Set rs = CreateObject("adodb.Recordset")
                            sql = "select top 1 *  from customerbalance where accno ='" & myrec!ACCNO & "' order by  CUSTOMERBALANCEID desc"
                            
                            rs.Open sql, cn
                          If rs.EOF Then
                          MsgBox "No record selected for updating", vbExclamation, "Customer Balance Update"
                          Exit Sub
                          Else
                            If Not rs.EOF Then rs.MoveFirst
                                If Not IsNull(rs!availablebalance) Then avail = rs!availablebalance Else avail = 0
                                If Not IsNull(rs!customerbalanceid) Then cust = rs!customerbalanceid Else cust = 0
                                If Not IsNull(rs!Commission) Then com = rs!Commission Else com = 0
                                If Not IsNull(rs!amount) Then amount = rs!amount Else amount = 0
                                If Not IsNull(rs!transDescription) Then trn = rs!transDescription Else trn = "other"
                                If Not IsNull(rs!vno) Then v = rs!vno Else v = "oth"
                                If Not IsNull(rs!transdate) Then tdate = rs!transdate Else tdate = "& now &"
                                If Not IsNull(rs!chequeno) Then chequeno = rs!chequeno Else chequeno = "na"
                                If Not IsNull(rs!transtype) Then ttype = rs!transtype
                            
                            Set cub = CreateObject("adodb.recordset")
                            
                            sql = ""
                            sql = "select * from cub WHERE ACCNO='" & myrec!ACCNO & "'"
                            cub.Open sql, cn
                             If Not cub.EOF Then
                              
                              sql = "update cub set amount=" & amount & ",transdescription='" & trn & "',availablebalance=" & avail & ",commission=" & com & ",transdate='" & tdate & "',vno='" & v & "',chequeno='" & chequeno & "',period='" & month(Date) & "',auditid='" & TELL & "',auditdate='" & Date & "',moduleid=2,active=1 where accno='" & myrec!ACCNO & "'"
                              myclass.save sql
                             Else
                                sql = " INSERT INTO CUB"
                                sql = sql & " (AccNo, AccountName, Name, Amount, Transdescription, AvailableBalance, Commission, Transdate, VNO, ChequeNo, Transtype, period, AUDITID,"
                                sql = sql & "  Auditdate, moduleid)"
                                sql = sql & "   VALUES ('" & myrec!ACCNO & "', ' Savings account', 'NAH', " & amount & ", '" & trn & "', " & avail & ", " & com & ", '" & tdate & "', '" & v & "', '" & chequeno & "', '" & ttype & "', '" & month(Date) & "', '" & TELL & "', '" & Now & "', '2')"
                                myclass.save sql
                            
                              End If
                End If
                number = number + 1
                Pb.value = number
               Label3.Caption = " Processing Record " & CStr(myrec!ACCNO)
               Label3.Refresh
              
        myrec.MoveNext
        
        
    Wend

End Sub

Private Sub cmdupdate_Click()
On Error Resume Next
teller = GetSetting("FOSAdll", "Teller", "Name")
Dim myIndex As Integer
 Dim temprs As Object
 Dim rsb As Object
    Dim myrec As Object, X As Integer
    maxRec = 0
    Dim cub As Object
   
Dim Net As Currency
Dim net1 As Currency
Dim number As Integer
Dim Amnt, ACCNO
Dim bal As Currency
Dim tot As Currency
Dim vno As String
Dim id As Long
Dim avail As Currency
Dim com As Currency
 Dim TABI As Object
Dim myclass As cdbase
Set myclass = New cdbase
Set myrec = CreateObject("adodb.recordset")
Set cn = CreateObject("adodb.connection")
Provider = myclass.OpenCon
cn.Open Provider, "atm", "atm"
            sql = "SELECT  cashid,customerno,accname,amount,accno,transdescription,"
            sql = sql & " transdate,commission,chequeno,transtype,vno,auditid,audittime from cashtransactions1 where accno >='" & txtFrom & "' and accno < = '" & txtTo & "' and transdescription <> ' ChDeposit' and transdescription <>'Cheque Deposit(uncleared)'order by accno asc"
            myrec.Open sql, cn

                    While Not myrec.EOF
                       number = number + 1
                       myrec.MoveNext
                    Wend
                    Pb.max = number + 1
                    Pb.Min = 0
                    Pb.Visible = True
                    myrec.Requery
                    number = 0
                
            With ctdata
            
                While Not myrec.EOF
                myrec.Refresh
                        If Not IsNull(myrec!ACCNO) Then .ACCNO = myrec!ACCNO
                        If Not IsNull(myrec!AccName) Then .AccName = myrec!AccName
                        If Not IsNull(myrec!custno) Then .custno = myrec!CustomerNo
                        If Not IsNull(myrec!amount) Then .Amnt = myrec!amount Else .Amnt = 0
                        If Not IsNull(myrec!transDescription) Then .transdes = myrec!transDescription Else .transdes = "Non"
                        If Not IsNull(myrec!transdate) Then .transdate = Format(myrec!transdate, "mm-dd-yyyy")
                        If Not IsNull(myrec!Commission) Then .Commission = myrec!Commission Else .Commission = 0
                        If Not IsNull(myrec!chequeno) Then .chequeno = myrec!chequeno Else .chequeno = "Non"
                        If Not IsNull(myrec!transtype) Then .transtype = myrec!transtype
                        If Not IsNull(myrec!vno) Then .vno = myrec!vno Else .vno = myrec!transDescription
                        If Not IsNull(myrec!auditid) Then .audid = myrec!auditid
                        If Not IsNull(myrec!audittime) Then .audittime = myrec!audittime
                        If Not IsNull(myrec!cashid) Then .id = myrec!cashid
                
                
         
                If .transdes = "TR" Then '//RETURNS
                
                        sql = ""
                        sql = "insert into [Teller Transactions](CubicleNumber,tellername,transactiondate,Deposits,accountnumber,auditid,audittime,transdescription,balance,name,accname,printed)"
                        sql = sql & "select'B12','" & .audid & "','" & .transdate & "'," & .Amnt & ",'" & .ACCNO & "','" & .audid & "','" & Now & "','Imprest ',0,'" & .audid & "','" & .audid & "',0"
                        myclass.save sql
                        
                       sql = "select top 1 * from treasuarytrans  order by treasuryid desc"
                        Set rsb = CreateObject("adodb.recordset")
                        rsb.Open sql, cn
                        
                        If rsb.EOF Then
                        tot = .Amnt
                        Else
                        tot = rsb!balance
                        End If
                        rsb.Close
   'update the bank transaction by subtracting that amount from the bank
   
   
                        sql = ""
                        sql = "INSERT INTO Treasuarytrans"
                        sql = sql & "( Rdate, Operator,Amount,balance,Description, Moduleid,mode,Auditdatetime, Auditid,posted,locked)"
                        sql = sql & "VALUES     ( '" & .transdate & "', '" & .audid & "'," & .Amnt & "," & tot & ",'Tellers Returns', '2','Cash','" & Now & "', '" & .audid & "',0,0)"
                        myclass.save sql
                
                   ElseIf .transdes = "TAR" Then '// TELLER AMOUNT ISSUED
                   
                   sql = ""
                        sql = "insert into [Teller Transactions](CubicleNumber,tellername,transactiondate,Deposits,accountnumber,auditid,audittime,transdescription,balance,name,accname,printed)"
                        sql = sql & "select'B12','" & .audid & "','" & .transdate & "'," & .Amnt & ",'" & .ACCNO & "','" & .audid & "','" & Now & "','Imprest ',0,'" & .audid & "','" & .audid & "',0"
                        myclass.save sql
                        
                        sql = "select top 1 * from treasuarytrans  order by treasuryid desc"
                        Set rsb = CreateObject("adodb.recordset")
                        rsb.Open sql, cn
                        
                        If rsb.EOF Then
                        tot = .Amnt
                        Else
                        tot = rsb!balance
                         End If
                        rsb.Close
                        'update the bank transaction by subtracting that amount from the bank
   
   
                        sql = ""
                        sql = "INSERT INTO Treasuarytrans"
                        sql = sql & "( Rdate, Operator,Amount,balance,Description, Moduleid,mode,Auditdatetime, Auditid,posted,locked)"
                        sql = sql & "VALUES     ( '" & Now & "', '" & .audid & "'," & .Amnt & "," & tot & ",'Tellers Returns', '2','Cash','" & Now & "', '" & .audid & "',0,0)"
                        myclass.save sql
                   
                   Else
                   
                            Set rs = CreateObject("adodb.Recordset")
                            sql = "select top 1 availablebalance,customerbalanceid,commission from customerbalance where accno ='" & myrec!ACCNO & "' order by  transdate desc"
                            
                            rs.Open sql, cn
                            If Not rs.EOF Then 'rs.movefirst
                            Dim cust
                            
                            avail = rs!availablebalance - .Commission
                            cust = rs!customerbalanceid
   
                            End If
                           
                         
                            sql = ""
                            

                            
                            

                        If .transtype = "DR " Then
                        Net = avail - myrec!amount
                         If rs.EOF Then
                            'sql = "update customerbalance set AvailableBalance= " & avail & " + " & myrec!amount & "  where accno='" & ctdata.accNo & "' and customerbalanceid='" & cust & "' "
                                
                                    sql = ""
                                    sql = "insert into customerbalance (customerno,accname,amount,availablebalance,accno,transdescription,transdate,commission,transType,posted,locked,status,vno,auditid,auditdate,moduleid,gene) "
                                    sql = sql & " values ('" & .custno & "','" & .AccName & "'," & .Amnt & "," & .Amnt & ",'" & .ACCNO & "','" & .transdes & "','" & .transdate & "'," & .Commission & ",'" & .transtype & "',0,0,0,'" & .vno & "','" & .audid & "','" & .audittime & "','2','" & br & "' )"
                                  cn.Execute sql
                                    
                                    
                                    
                                     Set cub = CreateObject("adodb.recordset")
                            
                                     sql = ""
                                     sql = "select * from cub WHERE ACCNO='" & .ACCNO & "'"
                                        cub.Open sql, cn
                                    If Not cub.EOF Then
                              
                                     sql = "update cub set amount=" & .Amnt & ",transdescription='" & .transdes & "',availablebalance=" & .Amnt & ",commission=" & .Commission & ",transdate='" & .transdate & "',vno='" & .vno & "',chequeno='" & .chequeno & "',period='" & month(Date) & "',auditid='" & .audid & "',auditdate='" & .audittime & "',moduleid=2,active=1 where accno='" & .ACCNO & "'"
                                     cn.Execute sql
                                     Else
                                     sql = " INSERT INTO CUB"
                                     sql = sql & " (AccNo, AccountName, Name, Amount, Transdescription, AvailableBalance, Commission, Transdate, VNO, ChequeNo, Transtype, period, AUDITID,"
                                     sql = sql & "  Auditdate, moduleid)"
                                     sql = sql & "   VALUES ('" & .ACCNO & "', '" & .AccName & "', '" & .AccName & "', " & .Amnt & ", '" & .transdes & "', " & .Amnt & ", " & .Commission & ", '" & .transdate & "', '" & .vno & "', '" & .chequeno & "', '" & .transtype & "', '" & month(Date) & "', '" & .audid & "', '" & .audittime & "', '2')"
                                     cn.Execute sql
                                     
                                     
                                      
                                    sql = ""
                                    sql = "delete from cashtransactions1 where cashid='" & .id & "'"
                                    cn.Execute sql
                                    End If
                            Else
                               sql = ""
                                    sql = "insert into customerbalance (customerno,accname,amount,availablebalance,accno,transdescription,transdate,commission,transType,posted,locked,status,vno,auditid,auditdate,moduleid,gene) "
                                    sql = sql & " values ('" & .custno & "','" & .AccName & "'," & .Amnt & "," & Net & ",'" & .ACCNO & "','" & .transdes & "','" & .transdate & "'," & .Commission & ",'" & .transtype & "',0,0,0,'" & .vno & "','" & .audid & "','" & .audittime & "','2','" & br & "' )"
                                   cn.Execute sql
                                 
                                    Set cub = CreateObject("adodb.recordset")
                            
                                    sql = ""
                                    sql = "select * from cub WHERE ACCNO='" & .ACCNO & "'"
                                     cub.Open sql, cn
                                    If Not cub.EOF Then
                              
                                    sql = "update cub set amount=" & .Amnt & ",transdescription='" & .transdes & "',availablebalance=" & Net & ",commission=" & .Commission & ",transdate='" & .transdate & "',vno='" & .vno & "',chequeno='" & .chequeno & "',period='" & month(Date) & "',auditid='" & .audid & "',auditdate='" & .audittime & "',moduleid=2,active=1 where accno='" & .ACCNO & "'"
                                     cn.Execute sql
                                     sql = ""
                                        sql = "delete from cashtransactions1 where cashid='" & .id & "'"
                                        cn.Execute sql
                                     Else
                                     sql = " INSERT INTO CUB"
                                     sql = sql & " (AccNo, AccountName, Name, Amount, Transdescription, AvailableBalance, Commission, Transdate, VNO, ChequeNo, Transtype, period, AUDITID,"
                                     sql = sql & "  Auditdate, moduleid)"
                                     sql = sql & "   VALUES ('" & .ACCNO & "', '" & .AccName & "', 'NAH', " & .Amnt & ", '" & .transdes & "', " & Net & ", " & .Commission & ", '" & .transdate & "', '" & .vno & "', '" & .chequeno & "', '" & .transtype & "', '" & month(Date) & "', '" & .audid & "', '" & .audittime & "', '2')"
                                     cn.Execute sql
                                     sql = ""
                                     sql = "delete from cashtransactions1 where cashid='" & .id & "'"
                                     cn.Execute sql
                                     
'                                     sql = ""
'                           sql = "delete from cashtransactions1 where accno='" & .accno & "'"
'                           MyClass.Delete sql
                                    End If
                           End If
                           
                        End If
                        If .transtype = "CR " Then
                         Net = avail + myrec!amount
                          If rs.EOF Then
                            'sql = "update customerbalance set AvailableBalance= " & avail & " + " & myrec!amount & "  where accno='" & ctdata.accNo & "' and customerbalanceid='" & cust & "' "
                                
                                    sql = ""
                                    sql = "insert into customerbalance (customerno,accname,amount,availablebalance,accno,transdescription,transdate,commission,transType,posted,locked,status,vno,auditid,auditdate,moduleid,gene) "
                                    sql = sql & " values ('" & .custno & "','" & .AccName & "'," & .Amnt & "," & .Amnt & ",'" & .ACCNO & "','" & .transdes & "','" & .transdate & "'," & .Commission & ",'" & .transtype & "',0,0,0,'" & .vno & "','" & .audid & "','" & .audittime & "','2','" & br & "' )"
                                    cn.Execute sql
                                    
                                    Set cub = CreateObject("adodb.recordset")
                            
                                    sql = ""
                                    sql = "select * from cub WHERE ACCNO='" & .ACCNO & "'"
                                    cub.Open sql, cn
                                    If Not cub.EOF Then
                              
                                    sql = "update cub set amount=" & .Amnt & ",transdescription='" & .transdes & "',availablebalance=" & Net & ",commission=" & .Commission & ",transdate='" & .transdate & "',vno='" & .vno & "',chequeno='" & .chequeno & "',period='" & month(Date) & "',auditid='" & .audid & "',auditdate='" & .audittime & "',moduleid=2,active=1 where accno='" & .ACCNO & "'"
                                      cn.Execute sql
                                     Else
                                     sql = " INSERT INTO CUB"
                                     sql = sql & " (AccNo, AccountName, Name, Amount, Transdescription, AvailableBalance, Commission, Transdate, VNO, ChequeNo, Transtype, period, AUDITID,"
                                     sql = sql & "  Auditdate, moduleid)"
                                     sql = sql & "   VALUES ('" & .ACCNO & "', '" & .AccName & "', 'NAH', " & .Amnt & ", '" & .transdes & "', " & Net & ", " & .Commission & ", '" & .transdate & "', '" & .vno & "', '" & .chequeno & "', '" & .transtype & "', '" & month(Date) & "', '" & .audid & "', '" & .audittime & "', '2')"
                                    cn.Execute sql
                                     
                                    sql = ""
                                    sql = "delete from cashtransactions1 where cashid='" & .id & "'"
                                    cn.Execute sql
                                    End If
                                    
                            Else
                               sql = ""
                                    sql = "insert into customerbalance (customerno,accname,amount,availablebalance,accno,transdescription,transdate,commission,transType,posted,locked,status,vno,auditid,auditdate,moduleid,gene) "
                                    sql = sql & " values ('" & .custno & "','" & .AccName & "'," & .Amnt & "," & Net & ",'" & .ACCNO & "','" & .transdes & "','" & .transdate & "'," & .Commission & ",'" & .transtype & "',0,0,0,'" & .vno & "','" & .audid & "','" & .audittime & "','2','" & br & "' )"
                                    cn.Execute sql
                                    
                                    
                                    
                                    
                                    Set cub = CreateObject("adodb.recordset")
                            
                                    sql = ""
                                    sql = "select * from cub WHERE ACCNO='" & .ACCNO & "'"
                                    cub.Open sql, cn
                                    If Not cub.EOF Then
                              
                                    sql = "update cub set amount=" & .Amnt & ",transdescription='" & .transdes & "',availablebalance=" & Net & ",commission=" & .Commission & ",transdate='" & .transdate & "',vno='" & .vno & "',chequeno='" & .chequeno & "',period='" & month(Date) & "',auditid='" & .audid & "',auditdate='" & .audittime & "',moduleid=2,active=1 where accno='" & .ACCNO & "'"
                                      cn.Execute sql
                                     
                                     sql = ""
                                        sql = "delete from cashtransactions1 where cashid='" & .id & "'"
                                        cn.Execute sql
                                     Else
                                     sql = " INSERT INTO CUB"
                                     sql = sql & " (AccNo, AccountName, Name, Amount, Transdescription, AvailableBalance, Commission, Transdate, VNO, ChequeNo, Transtype, period, AUDITID,"
                                     sql = sql & "  Auditdate, moduleid)"
                                     sql = sql & "   VALUES ('" & .ACCNO & "', '" & .AccName & "', 'NAH', " & .Amnt & ", '" & .transdes & "', " & Net & ", " & .Commission & ", '" & .transdate & "', '" & .vno & "', '" & .chequeno & "', '" & .transtype & "', '" & month(Date) & "', '" & .audid & "', '" & .audittime & "', '2')"
                                      cn.Execute sql
                                        
                                        sql = ""
                                        sql = "delete from cashtransactions1 where cashid='" & .id & "'"
                                        cn.Execute sql
                                    End If
                           End If
'                            sql = ""
'                            sql = "delete from cashtransactions1 where accno='" & .accno & "'"
'                            cn.Execute sql
                          
                        End If
                        
                        
                        'Else
                        
                        
                    ' Else
                     
                     '//TELLERS TRANSACTIONS TAR
                     
                     
                     
                     End If
                           
                    number = number + 1
                    Pb.value = number
                    myIndex = myIndex + 1
                    myrec.MoveNext
                     Label3.Caption = " Processing Record " & CStr(.ACCNO)
                     Label3.Refresh
                     rs.Refresh
                     myrec.Refresh

                Wend
                
            End With
        
        
      

        'If rs.EOF Then Exit Sub
        
        'get the number of records
       
       MsgBox "Updating Transactions " & number & "  Complete", vbInformation, "Transactions update"
       
       
       
        
        Pb.Visible = False
End Sub

Private Sub mnucub_Click()

End Sub

Private Sub cmdwrite_Click()
    On Error GoTo ErrorHandler
    Dim I As Long, r As Double
    Dim total
    total = 0
    Dim ttt
    Dim number As String, rsCustBal As New Recordset, Has_Trans As Boolean
    number = 1
    Dim fso As New FileSystemObject, txtFile As TextStream
    Dim myclass As cdbase
    Set myclass = New cdbase
    Set rs = CreateObject("adodb.recordset")
    Set cn = CreateObject("adodb.connection")
    Provider = myclass.OpenCon
   cn.Open Provider, "atm", "atm"
    '// place the exe
    DLG.Filter = "Text Files|*.txt"
   ' DLG.ShowSave
    If DLG.FileName <> "" Then
       ' ttt = DLG.FileName
    Else
        ttt = ""
        MsgBox "Please select the file to write to.", vbInformation, Me.Caption
        Exit Sub
    End If
'    If ttt = "" Then
'        MsgBox "File should not be blank", vbCritical, "Data transfer"
'        Exit Sub
'    End If
    Dim as5 As String, ADESCHEADER As String
      
    
    'ttt = "D:\PROJECTS\FOSA\DAILY" & Date & ""
    
    total = 0
    sql = ""
    Dim custno As String, rsNewAccounts As New Recordset
    Dim amount  As Currency, ACCNO As String, TransDesc As String, transdate As Date
    Dim chequeno As String, Period As String, transtype As String, vno As String
    Dim audit As String, accd As String, valdate As String, cash As Boolean
    Dim idno As String, name As String, AccName As String, empName As String, Address As String, _
    Telephone As String, NomID1 As String, NomiID2 As String, NomiID3 As String, _
    NomiName1 As String, NomiName2 As String, NomiName3 As String, Sig1ID As String, _
    Sig2ID As String, Sig3ID As String, Sig4ID As String, SigName1 As String, SigName2 _
    As String, SigName3 As String, SigName4 As String
   Dim DestFsoHeader As New FileSystemObject, DestFileHeader As TextStream
    Dim sno As String, qs As Currency, ppu As Currency, PAmount As Currency, transtime As String, auditid As String, auditdatetime As String, paid As Integer, lr As Integer, remark As String, br As String
    sql = ""
    sql = "set dateformat dmy SELECT * FROM d_milkintake WHERE"
    sql = sql & "(TransDate >= '" & DTPfrom & "') AND (TransDate <= '" & DTPto & "') and br='" & Trim(bb) & "' "
    rs.Open sql, cn
      as5 = Replace(Date & Time, "/", "")
    as5 = Trim(Replace(as5, ":", ""))
    ADESCHEADER = "d_milkintake" & as5 & ".txt"
   Set fso = CreateObject("Scripting.FileSystemObject")
    'Set txtfile = fso.CreateTextFile(ADESCHEADER, True)
   Set DestFileHeader = DestFsoHeader.CreateTextFile(txtADestHeader & "\" & ADESCHEADER)
    While Not rs.EOF
        If Not IsNull(rs!sno) Then sno = rs!sno
        If Not IsNull(rs!transdate) Then transdate = rs!transdate
        If Not IsNull(rs!QSupplied) Then qs = rs!QSupplied
        If Not IsNull(rs!ppu) Then ppu = rs!ppu
        If Not IsNull(rs!PAmount) Then PAmount = rs!PAmount
        If Not IsNull(rs!transtime) Then transtime = rs!transtime
        If Not IsNull(rs!auditid) Then auditid = rs!auditid
        If Not IsNull(rs!auditdatetime) Then auditdatetime = rs!auditdatetime
        If Not IsNull(rs!paid) Then paid = rs!paid
        If Not IsNull(rs!lr) Then lr = rs!lr
        If Not IsNull(rs!remark) Then remark = rs!remark
        If Not IsNull(rs!br) Then br = rs!br
       
        
        DestFileHeader.WriteLine (" " & sno & "," & transdate & " ," & qs & " ," & ppu & " ," & _
        PAmount & " ," & transtime & " ," & auditid & " ," & auditdatetime & " ," & paid & " ," & _
        lr & " ," & remark & "," & br & "")
        I = 0
        number = number + 1
        'Pb.value = number
        rs.MoveNext
    Wend
    
    '//suppliers table*************************START SUPPLIERS
    Dim sno1 As String, Regdate As Date, NAMES As String
    Dim bcode As String, BBranch As String, Type1 As String, Village As String, Location As String
    Dim Division As String, District As String, Trader As String, Active As String, Branch As String
    Dim PhoneNo As String, town As String, Email As String, TransCode As String
    Dim scode As String, loan As Integer, Compare As String, Isfrate As String, frate As String, rate As String, hast As Integer
    sql = ""
            sql = "SET              DATEFORMAT DMY"
            sql = sql & " SELECT     SNo, Regdate, IdNo, [Names], AccNo, Bcode, BBranch, Type, Village, Location, Division, District, Trader, active, Branch, PhoneNo,"
            sql = sql & " Address , town, Email, TransCode, AuditID, auditdatetime, scode, loan, Compare, isfrate, frate, rate, hast, br"
            sql = sql & "  From d_Suppliers"
            sql = sql & "  WHERE     REGDATE >= '" & DTPfrom & "' AND REGDATE <= '" & DTPto & "'  and br='" & bb & "' "
    Set rs = oSaccoMaster.GetRecordset(sql)
      as5 = Replace(Date & Time, "/", "")
    as5 = Trim(Replace(as5, ":", ""))
    ADESCHEADER = "d_Suppliers" & as5 & ".txt"
   Set fso = CreateObject("Scripting.FileSystemObject")
    'Set txtfile = fso.CreateTextFile(ADESCHEADER, True)
   Set DestFileHeader = DestFsoHeader.CreateTextFile(txtADestHeader & "\" & ADESCHEADER)
    While Not rs.EOF
        If Not IsNull(rs!sno) Then sno1 = rs!sno
        If Not IsNull(rs!Regdate) Then Regdate = rs!Regdate
        If Not IsNull(rs!idno) Then idno = rs!idno
        If Not IsNull(rs!NAMES) Then NAMES = rs!NAMES
        If Not IsNull(rs!ACCNO) Then ACCNO = rs!ACCNO
        If Not IsNull(rs!bcode) Then bcode = rs!bcode
        If Not IsNull(rs!BBranch) Then BBranch = rs!BBranch
        If Not IsNull(rs!Type) Then Type1 = rs!Type
        If Not IsNull(rs!Village) Then Village = rs!Village
        If Not IsNull(rs!Location) Then Location = rs!Location
        If Not IsNull(rs!Division) Then Division = rs!Division
        If Not IsNull(rs!District) Then District = rs!District
        If Not IsNull(rs!Trader) Then Trader = rs!Trader
        If Not IsNull(rs!Active) Then Active = rs!Active
        If Not IsNull(rs!Branch) Then Branch = rs!Branch
        If Not IsNull(rs!PhoneNo) Then Phone = rs!PhoneNo
        If Not IsNull(rs!Address) Then Address = rs!Address
        If Not IsNull(rs!town) Then town = rs!town
        If Not IsNull(rs!Email) Then Email = rs!Email
        If Not IsNull(rs!TransCode) Then TransCode = rs!TransCode
        If Not IsNull(rs!auditid) Then auditid = rs!auditid
        If Not IsNull(rs!auditdatetime) Then auditdatetime = rs!auditdatetime
        If Not IsNull(rs!scode) Then scode = rs!scode
        'loan, Compare, isfrate, frate, rate, hast, br
        If Not IsNull(rs!loan) Then loan = rs!loan
        If Not IsNull(rs!Compare) Then Compare = rs!Compare
        If Not IsNull(rs!Isfrate) Then Isfrate = rs!Isfrate
        If Not IsNull(rs!frate) Then frate = rs!frate
        If Not IsNull(rs!rate) Then rate = rs!rate
        If Not IsNull(rs!hast) Then hast = rs!hast
        If Not IsNull(rs!br) Then br = rs!br
       
        
        DestFileHeader.WriteLine (" " & sno1 & "," & Regdate & " ," & idno & " ," & NAMES & " ," & _
        ACCNO & " ," & bcode & " ," & BBranch & " ," & Type1 & " ," & Village & " ," & _
        Location & " ," & Division & "," & District & "," & Trader & " ," & Active & " ," & Branch & " ," & Phone & " ," & _
        Address & "," & town & " ," & Email & " ," & TransCode & " ," & auditid & " ," & auditdatetime & " ," & scode & " ," & loan & " ," & Compare & "," _
         & Isfrate & "," & frate & " ," & rate & " ," & hast & " ," & br & "")
        I = 0
        number = number + 1
        'Pb.value = number
        rs.MoveNext
    Wend

    
    '*******************************************END SUPPLIERS
    Dim sdate As Date, edate As Date, yYear As Long, certno As String, subsidy As Double

    '************************start supplier deduction *******************************
        sql = ""
    sql = "set dateformat dmy SELECT * FROM d_supplier_deduc WHERE"
    sql = sql & "(Date_Deduc >= '" & DTPfrom & "') AND (Date_Deduc <= '" & DTPto & "')  and br='" & bb & "' "
    Set rs = oSaccoMaster.GetRecordset(sql)
      as5 = Replace(Date & Time, "/", "")
    as5 = Trim(Replace(as5, ":", ""))
    ADESCHEADER = "d_supplier_deduc" & as5 & ".txt"
   Set fso = CreateObject("Scripting.FileSystemObject")
    'Set txtfile = fso.CreateTextFile(ADESCHEADER, True)
   Set DestFileHeader = DestFsoHeader.CreateTextFile(txtADestHeader & "\" & ADESCHEADER)
    While Not rs.EOF
        If Not IsNull(rs!sno) Then sno = rs!sno
        If Not IsNull(rs!Date_Deduc) Then transdate = rs!Date_Deduc
        If Not IsNull(rs!description) Then TransDesc = rs!description
        If Not IsNull(rs!amount) Then amount = rs!amount
        If Not IsNull(rs!Period) Then Period = rs!Period
        If Not IsNull(rs!Startdate) Then sdate = rs!Startdate
        If Not IsNull(rs!Enddate) Then edate = rs!Enddate
        If Not IsNull(rs!auditid) Then auditid = rs!auditid
        If Not IsNull(rs!auditdatetime) Then auditdatetime = rs!auditdatetime
        If Not IsNull(rs!yYear) Then yYear = rs!yYear
        If Not IsNull(rs!Remarks) Then remark = rs!Remarks
        If Not IsNull(rs!br) Then br = rs!br
       
        
        DestFileHeader.WriteLine (" " & sno & "," & transdate & " ," & TransDesc & " ," & amount & " ," & _
        Period & " ," & sdate & " ," & edate & " ," & auditid & " ," & auditdatetime & " ," & _
        yYear & " ," & remark & "," & br & "")
        I = 0
        number = number + 1
        'Pb.value = number
        rs.MoveNext
    Wend
    '************************end supplier deductions*********************************
     '************************start transporter deduction *******************************
        sql = ""
    sql = "set dateformat dmy SELECT * FROM d_Transport_Deduc WHERE"
    sql = sql & "(TDate_Deduc >= '" & DTPfrom & "') AND (TDate_Deduc <= '" & DTPto & "')  and br='" & bb & "' "
    Set rs = oSaccoMaster.GetRecordset(sql)
      as5 = Replace(Date & Time, "/", "")
    as5 = Trim(Replace(as5, ":", ""))
    ADESCHEADER = "d_Transport_Deduc" & as5 & ".txt"
   Set fso = CreateObject("Scripting.FileSystemObject")
    'Set txtfile = fso.CreateTextFile(ADESCHEADER, True)
   Set DestFileHeader = DestFsoHeader.CreateTextFile(txtADestHeader & "\" & ADESCHEADER)
    While Not rs.EOF
        If Not IsNull(rs!TransCode) Then sno = rs!TransCode
        If Not IsNull(rs!tdate_deduc) Then transdate = rs!tdate_deduc
        If Not IsNull(rs!description) Then TransDesc = rs!description
        If Not IsNull(rs!amount) Then amount = rs!amount
        If Not IsNull(rs!Period) Then Period = rs!Period
        If Not IsNull(rs!Startdate) Then sdate = rs!Startdate
        If Not IsNull(rs!Enddate) Then edate = rs!Enddate
        If Not IsNull(rs!auditid) Then auditid = rs!auditid
        If Not IsNull(rs!auditdatetime) Then auditdatetime = rs!auditdatetime
        If Not IsNull(rs!yYear) Then yYear = rs!yYear
        If Not IsNull(rs!rate) Then r = rs!rate
        If Not IsNull(rs!br) Then br = rs!br
       
        
        DestFileHeader.WriteLine (" " & sno & "," & transdate & " ," & TransDesc & " ," & amount & " ," & _
        Period & " ," & sdate & " ," & edate & " ," & auditid & " ," & auditdatetime & " ," & _
        yYear & " ," & r & "," & br & "")
        I = 0
        number = number + 1
        'Pb.value = number
        rs.MoveNext
    Wend

    
    
    '**********************end of transport deductions***************************
       sno1 = ""
       NAMES = ""
       certno = ""
       Location = ""
       'Regdate = ""
       Email = ""
       Phone = ""
       town = ""
       Address = ""
       subsidy = 0
       ACCNO = ""
       bcode = ""
       BBranch = ""
       Active = ""
       Branch = ""
       auditid = ""
       auditdatetime = ""
       Isfrate = ""
       r = 0
       br = ""

    '//d_Transporters table*************************START TRANSPORTERES
    'Dim sno1 As String, Regdate As Date, Names As String
    'Dim bcode As String, BBranch As String, Type1 As String, Village As String, Location As String
    'Dim Division As String, District As String, Trader As String, Active As String, Branch As String
   ' Dim PhoneNo As String, town As String, Email As String, TransCode As String, scode As String, loan As Integer, Compare As String, isfrate As String, frate As String, rate As String, hast As Integer
    sql = ""
        sql = " SET              dateformat dmy"
        sql = sql & "        SELECT     TransCode, TransName, CertNo, Locations, TregDate, email, Phoneno, Town, Address, Subsidy, Accno, Bcode, BBranch, Active, TBranch,"
        sql = sql & "                           auditid , auditdatetime, isfrate, rate, BR"
        sql = sql & "   From d_Transporters"
        sql = sql & "  WHERE     tregdate >= '" & DTPfrom & "' AND tregdate <= '" & DTPto & "'  and br='" & bb & "' "
    Set rs = oSaccoMaster.GetRecordset(sql)
      as5 = Replace(Date & Time, "/", "")
    as5 = Trim(Replace(as5, ":", ""))
    ADESCHEADER = "d_Transporters" & as5 & ".txt"
   Set fso = CreateObject("Scripting.FileSystemObject")
    'Set txtfile = fso.CreateTextFile(ADESCHEADER, True)
   Set DestFileHeader = DestFsoHeader.CreateTextFile(txtADestHeader & "\" & ADESCHEADER)
   'sno1,names,certno,location,regdate,email,phone,town,address,subsidy,accno,bcode,bbranch,active,branch,auditid,auditdatetime,isfrate,r,br
   
    While Not rs.EOF
        If Not IsNull(rs!TransCode) Then sno1 = rs!TransCode
        If Not IsNull(rs!TransName) Then NAMES = rs!TransName
        If Not IsNull(rs!certno) Then certno = rs!certno
        If Not IsNull(rs!Locations) Then Location = rs!Locations
        If Not IsNull(rs!tregdate) Then Regdate = rs!tregdate
        If Not IsNull(rs!Email) Then Email = rs!Email
        If Not IsNull(rs!PhoneNo) Then Phone = rs!PhoneNo
        If Not IsNull(rs!town) Then town = rs!town
        If Not IsNull(rs!Address) Then Address = rs!Address
        If Not IsNull(rs!subsidy) Then subsidy = rs!subsidy
        If Not IsNull(rs!ACCNO) Then ACCNO = rs!ACCNO
        If Not IsNull(rs!bcode) Then bcode = rs!bcode
        If Not IsNull(rs!BBranch) Then BBranch = rs!BBranch
        If Not IsNull(rs!Active) Then Active = rs!Active
        If Not IsNull(rs!tBranch) Then Branch = rs!tBranch
        If Not IsNull(rs!auditid) Then auditid = rs!auditid
        If Not IsNull(rs!auditdatetime) Then auditdatetime = rs!auditdatetime
        If Not IsNull(rs!Isfrate) Then Isfrate = rs!Isfrate
        If Not IsNull(rs!rate) Then r = rs!rate
        If Not IsNull(rs!br) Then br = rs!br
       
        
        DestFileHeader.WriteLine (" " & sno1 & "," & NAMES & " ," & certno & " ," & Location & " ," & _
        Regdate & " ," & Email & " ," & Phone & " ," & town & " ," & Address & " ," & _
        subsidy & " ," & ACCNO & "," & bcode & "," & BBranch & " ," & Active & " ," & Branch & "," & auditid & " ," & auditdatetime & " ," _
         & Isfrate & "," & r & " ," & br & "")
        I = 0
        number = number + 1
        'Pb.value = number
        rs.MoveNext
    Wend

    
    '*******************************************END TRANSPORT DETAILS
    
    '//////////////////////////////////////////////////////////new definitions
    
    
    '//d_Transporters table*************************START FARMER TRANSPORT ASSIGNMENT
    Dim trans_code As String, Startdate As Date
    Dim dateinactivate As Date
  
    
sql = "SET              dateformat dmy"
sql = sql & "         SELECT     trans_code, sno, rate, startdate, active, dateinactivate, auditid, auditdatetime, isfrate, br"
sql = sql & "         From d_Transport"
sql = sql & "      WHERE     startdate >= '" & DTPfrom & "' AND startdate <= '" & DTPto & "'  and br='" & bb & "'"
sql = sql & "      ORDER BY ID DESC"
    
    Set rs = oSaccoMaster.GetRecordset(sql)
      as5 = Replace(Date & Time, "/", "")
    as5 = Trim(Replace(as5, ":", ""))
    ADESCHEADER = "d_Transport" & as5 & ".txt"
   Set fso = CreateObject("Scripting.FileSystemObject")
    'Set txtfile = fso.CreateTextFile(ADESCHEADER, True)
   Set DestFileHeader = DestFsoHeader.CreateTextFile(txtADestHeader & "\" & ADESCHEADER)
   'sno1,names,certno,location,regdate,email,phone,town,address,subsidy,accno,bcode,bbranch,active,branch,auditid,auditdatetime,isfrate,r,br
   
    While Not rs.EOF
        If Not IsNull(rs!trans_code) Then trans_code = rs!trans_code
        If Not IsNull(rs!sno) Then sno = rs!sno
        If Not IsNull(rs!rate) Then rate = rs!rate
        If Not IsNull(rs!Startdate) Then Startdate = rs!Startdate
        If Not IsNull(rs!Active) Then Active = rs!Active
        If Not IsNull(rs!dateinactivate) Then dateinactivate = rs!dateinactivate
        If Not IsNull(rs!auditid) Then auditid = rs!auditid
        If Not IsNull(rs!auditdatetime) Then auditdatetime = rs!auditdatetime
        If Not IsNull(rs!Isfrate) Then Isfrate = rs!Isfrate
        If Not IsNull(rs!br) Then br = rs!br
       
        
        DestFileHeader.WriteLine (" " & trans_code & "," & sno & " ," & rate & " ," & Startdate & " ," & _
        Active & " ," & dateinactivate & " ," & auditid & " ," & auditdatetime & " ," & Isfrate & " ," & br & "")
        I = 0
        number = number + 1
        'Pb.value = number
        rs.MoveNext
    Wend

    
    '*******************************************END FARMER TRANSPORT ASSIGNMENTS................
    
    MsgBox "Records successfullly copied"
    Pb.Visible = False
    Exit Sub
ErrorHandler:
    MsgBox err.description
End Sub

Private Sub Generate_Back_Office_Transactions(Startdate As Date, FinishDate As Date)
    On Error GoTo SysError
    Dim FileName As String, rsLoanRep As New Recordset, Loanno As String, memberno As String, _
    datereceived As Date, paymentno As Long, amount As Double, principal As Double, interest As _
    Double, IntrCharged As Double, IntrOwed As Double, loanbalance As Double, ReceiptNo As _
    String, Remarks As String, auditid As String, transby As String, MyFso As New FileSystemObject, _
    BackOffFile As TextStream
    With DLG
        .Filter = "Text Files|*.txt|CSV Files|*.csv"
        .DialogTitle = "Select Back Office File"
        .ShowSave
        If .FileName <> "" Then
            FileName = .FileName
        End If
    End With
    If FileName <> "" Then
        Set BackOffFile = MyFso.CreateTextFile(FileName, False)
        'Set rsLoanRep = Get_Bosa_Records("Set DateFormat DMY Exec " _
        & "Get_Cash_Loan_Transactions '" & StartDate & "','" & FinishDate _
        & "'", ErrorMessage)
        With rsLoanRep
            If .State = adStateOpen Then
                While Not .EOF
                    DoEvents
                    Loanno = IIf(IsNull(!Loanno), "", !Loanno)
                    memberno = IIf(IsNull(!memberno), "", !memberno)
                    datereceived = IIf(IsNull(!datereceived), Date, !datereceived)
                    paymentno = IIf(IsNull(!paymentno), 1, !paymentno)
                    amount = IIf(IsNull(!amount), 0, !amount)
                    principal = IIf(IsNull(!principal), 0, !principal)
                    interest = IIf(IsNull(!interest), 0, !interest)
                    IntrCharged = IIf(IsNull(!IntrCharged), 0, !IntrCharged)
                    IntrOwed = IIf(IsNull(!IntrOwed), 0, !IntrOwed)
                    loanbalance = IIf(IsNull(!loanbalance), 0, !loanbalance)
                    ReceiptNo = IIf(IsNull(!ReceiptNo), "", !ReceiptNo)
                    Remarks = IIf(IsNull(!Remarks), "", !Remarks)
                    auditid = IIf(IsNull(!auditid), "", !auditid)
                    transby = IIf(IsNull(!transby), "", !transby)
                    BackOffFile.WriteLine Loanno & "," & memberno & "," & datereceived _
                    & "," & paymentno & "," & amount & "," & principal & "," & interest _
                    & "," & IntrCharged & "," & IntrOwed & "," & loanbalance & "," & _
                    ReceiptNo & "," & Remarks & "," & auditid & "," & transby
                    .MoveNext
                Wend
            Else
                If ErrorMessage <> "" Then
                    MsgBox ErrorMessage, vbInformation, Me.Caption
                    ErrorMessage = ""
                End If
            End If
        End With
    End If
    MsgBox "Completed successfully", vbInformation
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub Form_Load()
DTPfrom = Format(Get_Server_Date, "DD/MM/yy")
DTPto = Format(Get_Server_Date, "dd/MM/yy")
 sql = "SELECT     TOP 1 BR  FROM         d_company"
    Set rs = oSaccoMaster.GetRecordset(sql)
   If Not rs.EOF Then
        bb = IIf(IsNull(rs.Fields(0)), "A", rs.Fields(0))
   End If
End Sub
