VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPosttoGls 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Batch Posting to General Ledgers"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9150
   Icon            =   "frmPosttoGls.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   9150
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      Height          =   615
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   7695
      Begin MSComCtl2.DTPicker dtpYear 
         Height          =   375
         Left            =   4680
         TabIndex        =   7
         Top             =   120
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "yyyy"
         Format          =   122355715
         CurrentDate     =   39121
      End
      Begin MSComCtl2.DTPicker dtpMonth 
         Height          =   375
         Left            =   1680
         TabIndex        =   4
         Top             =   120
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "MM"
         Format          =   122355715
         CurrentDate     =   39121
      End
      Begin VB.Label Label2 
         Caption         =   "Year"
         Height          =   375
         Left            =   3600
         TabIndex        =   6
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Month"
         Height          =   255
         Left            =   720
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   6000
      Width           =   7695
      Begin VB.CommandButton CmdPost 
         Caption         =   "Post"
         Height          =   375
         Left            =   3720
         TabIndex        =   19
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   375
         Left            =   1560
         TabIndex        =   18
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Ledgers Control"
      Height          =   2775
      Left            =   120
      TabIndex        =   1
      Top             =   3240
      Width           =   9015
      Begin VB.PictureBox Picture10 
         Height          =   255
         Left            =   5880
         Picture         =   "frmPosttoGls.frx":27A2
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   32
         Top             =   600
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox Picture9 
         Height          =   255
         Left            =   5880
         Picture         =   "frmPosttoGls.frx":2A64
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   31
         Top             =   960
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox Picture8 
         Height          =   255
         Left            =   5880
         Picture         =   "frmPosttoGls.frx":2D26
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   30
         Top             =   1680
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox Picture7 
         Height          =   255
         Left            =   5880
         Picture         =   "frmPosttoGls.frx":2FE8
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   29
         Top             =   1320
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox Picture6 
         Height          =   255
         Left            =   5880
         Picture         =   "frmPosttoGls.frx":32AA
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   28
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox Picture4 
         Height          =   255
         Left            =   1440
         Picture         =   "frmPosttoGls.frx":356C
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   21
         Top             =   1320
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox Picture3 
         Height          =   255
         Left            =   1440
         Picture         =   "frmPosttoGls.frx":382E
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   20
         Top             =   1680
         Width           =   255
      End
      Begin VB.PictureBox Picture2 
         Height          =   255
         Left            =   1440
         Picture         =   "frmPosttoGls.frx":3AF0
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   11
         Top             =   960
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Height          =   255
         Left            =   1440
         Picture         =   "frmPosttoGls.frx":3DB2
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   10
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox Picture5 
         Height          =   255
         Left            =   1440
         Picture         =   "frmPosttoGls.frx":4074
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   9
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblmemberid 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6240
         TabIndex        =   40
         Top             =   600
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Label Label15 
         Caption         =   "Memberid Control"
         Height          =   255
         Left            =   4560
         TabIndex        =   39
         Top             =   600
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label14 
         Caption         =   "Passbook Acc"
         Height          =   255
         Left            =   4560
         TabIndex        =   38
         Top             =   960
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label lblPassbook 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6240
         TabIndex        =   37
         Top             =   960
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Label Label12 
         Caption         =   "Other Charges Control"
         Height          =   255
         Left            =   4560
         TabIndex        =   36
         Top             =   1680
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label lblOthercharges 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6240
         TabIndex        =   35
         Top             =   1680
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Label Label10 
         Caption         =   "Christmas Acc"
         Height          =   255
         Left            =   4560
         TabIndex        =   34
         Top             =   1320
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label lblchristmas 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6240
         TabIndex        =   33
         Top             =   1320
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Label lblEntranceFee 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1800
         TabIndex        =   27
         Top             =   1320
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Label Label7 
         Caption         =   "Entrance Acc"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1320
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label lblLoanForm 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1800
         TabIndex        =   25
         Top             =   1680
         Width           =   2655
      End
      Begin VB.Label Label5 
         Caption         =   "Loan form Control"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Calender Acc"
         Height          =   255
         Left            =   4560
         TabIndex        =   23
         Top             =   240
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label lblCalender 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6240
         TabIndex        =   22
         Top             =   240
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Label lblSharescash 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1800
         TabIndex        =   17
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label Label18 
         Caption         =   "Shares Control Acc"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label19 
         Caption         =   "Interest Control"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label lblintcontrol 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1800
         TabIndex        =   14
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label Label20 
         Caption         =   "Loan Control Acc"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lblloancontrolacc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1800
         TabIndex        =   12
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   8895
      Begin MSComctlLib.ListView lsvMdeduct 
         Height          =   2415
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   4260
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
End
Attribute VB_Name = "frmPosttoGls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdPost_Click()
''// this is the hardest thing
''//select sum all
Dim principal As Currency
Dim interest As Currency
Dim Shares As Currency
Dim entrancefee As Currency
Dim loanforma As Currency
Dim memberid As Currency
Dim passbook As Currency
Dim Othercharges As Currency
Dim christmas As Currency
Dim NA As String
''//
Dim mysql As String
Dim RsMdeduct As ADODB.Recordset
Dim amt As Currency
''//
Dim mMonth As Integer
Dim yYear  As Integer
Dim txtreceiptno As String

mMonth = Format(dtpMonth.value, "MM")
yYear = Format(dtpYear.value, "yyyy")

mysql = "SELECT     SUM(Principal) AS principal, SUM(Interest) AS Interest, SUM(SHARES) AS shares, SUM(entrancefee) AS entrancefeem, SUM(loanform) AS loanform" _
                      & ",SUM(calender) AS calender, SUM(memberid) AS memberid, SUM(passbook) AS passbook, SUM(othercharges) AS othercharges, SUM(Christmas)" _
                      & " AS christmas From MDEDUCT WHERE     (Mmonth = '" & mMonth & "') and yyear ='" & yYear & "'"

Set RsMdeduct = oSaccoMaster.GetRecordset(mysql)

If Not RsMdeduct.EOF Then
    If Not IsNull(RsMdeduct!principal) Then
        principal = RsMdeduct!principal
    Else
        principal = 0
    End If
    If Not IsNull(RsMdeduct!interest) Then
        interest = RsMdeduct!interest
    Else
        interest = 0
    End If
    If Not IsNull(RsMdeduct!Shares) Then
        Shares = RsMdeduct!Shares
    Else
        Shares = 0
    End If
    If Not IsNull(RsMdeduct!entrancefeem) Then
        entrancefee = RsMdeduct!entrancefeem
    Else
        entrancefee = 0
    End If
    If Not IsNull(RsMdeduct!loanform) Then
        loanform = RsMdeduct!loanform
    Else
        loanform = 0
    End If
    If Not IsNull(RsMdeduct!calender) Then
        calender = RsMdeduct!calender
    Else
        calender = 0
    End If
    If Not IsNull(RsMdeduct!memberid) Then
        memberid = RsMdeduct!memberid
    Else
        memberid = 0
    End If
    If Not IsNull(RsMdeduct!passbook) Then
        passbook = RsMdeduct!passbook
    Else
        passbook = 0
    End If
    If Not IsNull(RsMdeduct!Othercharges) Then
        Othercharges = RsMdeduct!Othercharges
    Else
        Othercharges = 0
    End If
    If Not IsNull(RsMdeduct!christmas) Then
        christmas = RsMdeduct!christmas
    Else
        christmas = 0
    End If
    
    ''// save it to customer balance
    
    If lblloancontrolacc <> "" Then '  LOAN ACCOUNT FIRST
        sql = ""
        
        If txtreceiptno = "" Then txtreceiptno = "L. Repayment " & mMonth
    
         sql = ""
         NA = lblloancontrolacc
         getde NA
        
        sql = "set dateformat dmy insert into customerbalance (customerno,accname,amount,availablebalance,accno,transdescription,transdate,commission,period,transType,posted,locked,status,vno,auditid,auditdate,moduleid,accd) "
        sql = sql & " values ('" & glmemno & "','" & glnamE1 & "'," & principal & "," & bookba + principal & ",'" & glaccno & "','" & txtreceiptno & "','" & Format(Get_Server_Date, "dd/mm/yyyy") & "',0,'" & month(Get_Server_Date) & "','CR',0,0,0,'" & txtreceiptno & "','" & User & "','" & Get_Server_Date & "','3','" & glaccno & "' )"
        
        
        oSaccoMaster.ExecuteThis sql
        
        sql = ""
        sql = "set dateformat dmy update cub set amount=" & principal & ",Active=1,transdescription='" & txtreceiptno & "',availablebalance=" & bookba + txtPrincipal & ",transdate='" & Format(Get_Server_Date, "dd/mm/yyyy") & "',vno='" & txtreceiptno & "',period='" & month(Get_Server_Date) & "',auditid='" & User & "',auditdate='" & Now & "',moduleid=2 where accno='" & glaccno & "'"
        oSaccoMaster.ExecuteThis sql
    End If
    
    If lblintcontrol <> "" Then '  LOAN ACCOUNT FIRST
        sql = ""
        
        
        txtreceiptno = "Interest. Repayment " & mMonth
    
         sql = ""
         NA = lblintcontrol
         
         getde NA
        
        sql = "set dateformat dmy insert into customerbalance (customerno,accname,amount,availablebalance,accno,transdescription,transdate,commission,period,transType,posted,locked,status,vno,auditid,auditdate,moduleid,accd) "
        sql = sql & " values ('" & glmemno & "','" & glnamE1 & "'," & interest & "," & bookba + interest & ",'" & glaccno & "','" & txtreceiptno & "','" & Format(Get_Server_Date, "dd/mm/yyyy") & "',0,'" & month(Get_Server_Date) & "','CR',0,0,0,'" & txtreceiptno & "','" & User & "','" & Get_Server_Date & "','3','" & glaccno & "' )"
        oSaccoMaster.ExecuteThis sql
        
        sql = ""
        sql = "set dateformat dmy update cub set amount=" & interest & ",Active=1,transdescription='" & txtreceiptno & "',availablebalance=" & bookba + txtPrincipal & ",transdate='" & Format(Get_Server_Date, "dd/mm/yyyy") & "',vno='" & txtreceiptno & "',period='" & month(Get_Server_Date) & "',auditid='" & User & "',auditdate='" & Now & "',moduleid=2 where accno='" & glaccno & "'"
        oSaccoMaster.ExecuteThis sql
    End If
    
    If lblSharescash <> "" Then '  LOAN ACCOUNT FIRST
        sql = ""
        
        
        txtreceiptno = "shares payment " & mMonth
    
         sql = ""
         NA = lblSharescash
         
         getde NA
        
        sql = "set dateformat dmy insert into customerbalance (customerno,accname,amount,availablebalance,accno,transdescription,transdate,commission,period,transType,posted,locked,status,vno,auditid,auditdate,moduleid,accd) "
        sql = sql & " values ('" & glmemno & "','" & glnamE1 & "'," & Shares & "," & bookba + Shares & ",'" & glaccno & "','" & txtreceiptno & "','" & Format(Get_Server_Date, "dd/mm/yyyy") & "',0,'" & month(Get_Server_Date) & "','CR',0,0,0,'" & txtreceiptno & "','" & User & "','" & Get_Server_Date & "','3','" & glaccno & "' )"
        oSaccoMaster.ExecuteThis sql
        
        sql = ""
        sql = "set dateformat dmy update cub set amount=" & Shares & ",Active=1,transdescription='" & txtreceiptno & "',availablebalance=" & bookba + txtPrincipal & ",transdate='" & Format(Get_Server_Date, "dd/mm/yyyy") & "',vno='" & txtreceiptno & "',period='" & month(Get_Server_Date) & "',auditid='" & User & "',auditdate='" & Now & "',moduleid=2 where accno='" & glaccno & "'"
        oSaccoMaster.ExecuteThis sql
    End If
    If lblEntranceFee <> "" Then '  LOAN ACCOUNT FIRST
        sql = ""
        
        
        txtreceiptno = "Entrance fee for  " & mMonth
    
         sql = ""
         NA = lblEntranceFee
         
         getde NA
        
        sql = "set dateformat dmy insert into customerbalance (customerno,accname,amount,availablebalance,accno,transdescription,transdate,commission,period,transType,posted,locked,status,vno,auditid,auditdate,moduleid,accd) "
        sql = sql & " values ('" & glmemno & "','" & glnamE1 & "'," & entrancefee & "," & bookba + entrancefee & ",'" & glaccno & "','" & txtreceiptno & "','" & Format(Get_Server_Date, "dd/mm/yyyy") & "',0,'" & month(Get_Server_Date) & "','CR',0,0,0,'" & txtreceiptno & "','" & User & "','" & Get_Server_Date & "','3','" & glaccno & "' )"
        oSaccoMaster.ExecuteThis sql
        
        sql = ""
        sql = "set dateformat dmy update cub set amount=" & entrancefee & ",Active=1,transdescription='" & txtreceiptno & "',availablebalance=" & bookba + txtPrincipal & ",transdate='" & Format(Get_Server_Date, "dd/mm/yyyy") & "',vno='" & txtreceiptno & "',period='" & month(Get_Server_Date) & "',auditid='" & User & "',auditdate='" & Now & "',moduleid=2 where accno='" & glaccno & "'"
        oSaccoMaster.ExecuteThis sql
    End If
    If lblLoanForm <> "" Then '  LOAN ACCOUNT FIRST
        sql = ""
        
        
        txtreceiptno = "loan form for  " & mMonth
    
         sql = ""
         NA = lblLoanForm
         
         getde NA
        
        sql = "set dateformat dmy insert into customerbalance (customerno,accname,amount,availablebalance,accno,transdescription,transdate,commission,period,transType,posted,locked,status,vno,auditid,auditdate,moduleid,accd) "
        sql = sql & " values ('" & glmemno & "','" & glnamE1 & "'," & loanforma & "," & bookba + loanforma & ",'" & glaccno & "','" & txtreceiptno & "','" & Format(Get_Server_Date, "dd/mm/yyyy") & "',0,'" & month(Get_Server_Date) & "','CR',0,0,0,'" & txtreceiptno & "','" & User & "','" & Get_Server_Date & "','3','" & glaccno & "' )"
        oSaccoMaster.ExecuteThis sql
        
        sql = ""
        sql = "set dateformat dmy update cub set amount=" & loanforma & ",Active=1,transdescription='" & txtreceiptno & "',availablebalance=" & bookba + txtPrincipal & ",transdate='" & Format(Get_Server_Date, "dd/mm/yyyy") & "',vno='" & txtreceiptno & "',period='" & month(Get_Server_Date) & "',auditid='" & User & "',auditdate='" & Now & "',moduleid=2 where accno='" & glaccno & "'"
        oSaccoMaster.ExecuteThis sql
    End If
    If lblCalender <> "" Then '  LOAN ACCOUNT FIRST
        sql = ""
        
        
        txtreceiptno = "Calender fee for  " & mMonth
    
         sql = ""
         NA = lblCalender
         
         getde NA
        
        sql = "set dateformat dmy insert into customerbalance (customerno,accname,amount,availablebalance,accno,transdescription,transdate,commission,period,transType,posted,locked,status,vno,auditid,auditdate,moduleid,accd) "
        sql = sql & " values ('" & glmemno & "','" & glnamE1 & "'," & Calendar & "," & bookba + Calendar & ",'" & glaccno & "','" & txtreceiptno & "','" & Format(Get_Server_Date, "dd/mm/yyyy") & "',0,'" & month(Get_Server_Date) & "','CR',0,0,0,'" & txtreceiptno & "','" & User & "','" & Get_Server_Date & "','3','" & glaccno & "' )"
        oSaccoMaster.ExecuteThis sql
        
        sql = ""
        sql = "set dateformat dmy update cub set amount=" & Calendar & ",Active=1,transdescription='" & txtreceiptno & "',availablebalance=" & bookba + txtPrincipal & ",transdate='" & Format(Get_Server_Date, "dd/mm/yyyy") & "',vno='" & txtreceiptno & "',period='" & month(Get_Server_Date) & "',auditid='" & User & "',auditdate='" & Now & "',moduleid=2 where accno='" & glaccno & "'"
        oSaccoMaster.ExecuteThis sql
    End If
    
    If lblmemberid <> "" Then '  LOAN ACCOUNT FIRST
        sql = ""
        
        
        txtreceiptno = "memberid fee for  " & mMonth & " " & yYear
    
         sql = ""
         NA = lblmemberid
         
         getde NA
        
        sql = "set dateformat dmy insert into customerbalance (customerno,accname,amount,availablebalance,accno,transdescription,transdate,commission,period,transType,posted,locked,status,vno,auditid,auditdate,moduleid,accd) "
        sql = sql & " values ('" & glmemno & "','" & glnamE1 & "'," & memberid & "," & bookba + txtPrincipal & ",'" & glaccno & "','" & txtreceiptno & "','" & Format(Get_Server_Date, "dd/mm/yyyy") & "',0,'" & month(Get_Server_Date) & "','CR',0,0,0,'" & txtreceiptno & "','" & User & "','" & Get_Server_Date & "','3','" & glaccno & "' )"
        oSaccoMaster.ExecuteThis sql
        
        sql = ""
        sql = "set dateformat dmy update cub set amount=" & memberid & ",Active=1,transdescription='" & txtreceiptno & "',availablebalance=" & bookba + txtPrincipal & ",transdate='" & Format(Get_Server_Date, "dd/mm/yyyy") & "',vno='" & txtreceiptno & "',period='" & month(Get_Server_Date) & "',auditid='" & User & "',auditdate='" & Now & "',moduleid=2 where accno='" & glaccno & "'"
        oSaccoMaster.ExecuteThis sql
    End If
    If lblPassbook <> "" Then '  LOAN ACCOUNT FIRST
        sql = ""
        
        
        txtreceiptno = "Passbook fee for  " & mMonth & " " & yYear
    
         sql = ""
         NA = lblPassbook
         
         getde NA
        
        sql = "set dateformat dmy insert into customerbalance (customerno,accname,amount,availablebalance,accno,transdescription,transdate,commission,period,transType,posted,locked,status,vno,auditid,auditdate,moduleid,accd) "
        sql = sql & " values ('" & glmemno & "','" & glnamE1 & "'," & passbook & "," & bookba + passbook & ",'" & glaccno & "','" & txtreceiptno & "','" & Format(Get_Server_Date, "dd/mm/yyyy") & "',0,'" & month(Get_Server_Date) & "','CR',0,0,0,'" & txtreceiptno & "','" & User & "','" & Get_Server_Date & "','3','" & glaccno & "' )"
        oSaccoMaster.ExecuteThis sql
        
        sql = ""
        sql = "set dateformat dmy update cub set amount=" & passbook & ",Active=1,transdescription='" & txtreceiptno & "',availablebalance=" & bookba + txtPrincipal & ",transdate='" & Format(Get_Server_Date, "dd/mm/yyyy") & "',vno='" & txtreceiptno & "',period='" & month(Get_Server_Date) & "',auditid='" & User & "',auditdate='" & Now & "',moduleid=2 where accno='" & glaccno & "'"
        oSaccoMaster.ExecuteThis sql
    End If
    If lblchristmas <> "" Then '  LOAN ACCOUNT FIRST
        sql = ""
        
        
        txtreceiptno = "Christmas fee for  " & mMonth & " " & yYear
    
         sql = ""
         NA = lblchristmas
         
         getde NA
        
        sql = "set dateformat dmy insert into customerbalance (customerno,accname,amount,availablebalance,accno,transdescription,transdate,commission,period,transType,posted,locked,status,vno,auditid,auditdate,moduleid,accd) "
        sql = sql & " values ('" & glmemno & "','" & glnamE1 & "'," & christmas & "," & bookba + christmas & ",'" & glaccno & "','" & txtreceiptno & "','" & Format(Get_Server_Date, "dd/mm/yyyy") & "',0,'" & month(Get_Server_Date) & "','CR',0,0,0,'" & txtreceiptno & "','" & User & "','" & Get_Server_Date & "','3','" & glaccno & "' )"
        oSaccoMaster.ExecuteThis sql
        
        sql = ""
        sql = "set dateformat dmy update cub set amount=" & christmas & ",Active=1,transdescription='" & txtreceiptno & "',availablebalance=" & bookba + txtPrincipal & ",transdate='" & Format(Get_Server_Date, "dd/mm/yyyy") & "',vno='" & txtreceiptno & "',period='" & month(Get_Server_Date) & "',auditid='" & User & "',auditdate='" & Now & "',moduleid=2 where accno='" & glaccno & "'"
        oSaccoMaster.ExecuteThis sql
    End If
    If lblOthercharges <> "" Then '  LOAN ACCOUNT FIRST
        sql = ""
        
        
        txtreceiptno = "Other charges for  " & mMonth & " " & yYear
    
         sql = ""
         NA = lblOthercharges
         
         getde NA
        
        sql = "set dateformat dmy insert into customerbalance (customerno,accname,amount,availablebalance,accno,transdescription,transdate,commission,period,transType,posted,locked,status,vno,auditid,auditdate,moduleid,accd) "
        sql = sql & " values ('" & glmemno & "','" & glnamE1 & "'," & Othercharges & "," & bookba + Othercharges & ",'" & glaccno & "','" & txtreceiptno & "','" & Format(Get_Server_Date, "dd/mm/yyyy") & "',0,'" & month(Get_Server_Date) & "','CR',0,0,0,'" & txtreceiptno & "','" & User & "','" & Get_Server_Date & "','3','" & glaccno & "' )"
        oSaccoMaster.ExecuteThis sql
        
        sql = ""
        sql = "set dateformat dmy update cub set amount=" & Othercharges & ",Active=1,transdescription='" & txtreceiptno & "',availablebalance=" & bookba + txtPrincipal & ",transdate='" & Format(Get_Server_Date, "dd/mm/yyyy") & "',vno='" & txtreceiptno & "',period='" & month(Get_Server_Date) & "',auditid='" & User & "',auditdate='" & Now & "',moduleid=2 where accno='" & glaccno & "'"
        oSaccoMaster.ExecuteThis sql
    End If
    
    sql = "Update mededuct set Postedtoledgers =1 where mmonth ='" & mMonth & "' and yyear ='" & yYear & "'"
    
    oSaccoMaster.ExecuteThis sql
    
End If


End Sub

Private Sub dtpMonth_Change()
    Dim li As ListItem
    Dim thisyear As Integer
    Dim thismonth As Integer
    Dim RsMdeduct As New ADODB.Recordset
    Dim cn As New ADODB.Connection
    Dim sql As String
    On Error GoTo ErrorHandler
    
    thisyear = Format(dtpYear.value, "yyyy")
    thismonth = Format(dtpMonth.value, "MM")
    
    sql = "select * from Mdeduct where mmonth =" & thismonth & " and yyear ='" & thisyear & "' and Postedtoledgers =0"
    
    
    Set RsMdeduct = oSaccoMaster.GetRecordset(sql)
    
    If Not RsMdeduct.EOF Then
    
    lsvMdeduct.ListItems.Clear
    
        Do While Not RsMdeduct.EOF
        
            MousePointer = vbHourglass
            
            ''///
            ''//get rid of null values
        
        If IsNull(RsMdeduct!memberno) Then RsMdeduct!memberno = 0
        If IsNull(RsMdeduct!principal) Then RsMdeduct!principal = 0
        If IsNull(RsMdeduct!interest) Then RsMdeduct!interest = 0
        If IsNull(RsMdeduct!Shares) Then RsMdeduct!Shares = 0
        If IsNull(RsMdeduct!entrancefee) Then RsMdeduct!entrancefee = 0
        If IsNull(RsMdeduct!loanform) Then RsMdeduct!loanform = 0
        If IsNull(RsMdeduct!calender) Then RsMdeduct!calender = 0
        If IsNull(RsMdeduct!passbook) Then RsMdeduct!passbook = 0
        If IsNull(RsMdeduct!Othercharges) Then RsMdeduct!Othercharges = 0
        If IsNull(RsMdeduct!christmas) Then RsMdeduct!christmas = 0
        
        
            
            
          Set li = lsvMdeduct.ListItems.Add(, , RsMdeduct!memberno & "")
              li.ListSubItems.Add , , RsMdeduct!principal
              li.ListSubItems.Add , , RsMdeduct!interest
              li.ListSubItems.Add , , RsMdeduct!Shares
              li.ListSubItems.Add , , RsMdeduct!entrancefee
              li.ListSubItems.Add , , RsMdeduct!loanform
              li.ListSubItems.Add , , RsMdeduct!calender
              li.ListSubItems.Add , , RsMdeduct!passbook
              li.ListSubItems.Add , , RsMdeduct!Othercharges
              li.ListSubItems.Add , , RsMdeduct!christmas
              
              RsMdeduct.MoveNext
        Loop
    Else
        lsvMdeduct.ListItems.Clear
    End If
    MousePointer = vbDefault
    Exit Sub
ErrorHandler:
    
    
End Sub

Private Sub Form_Load()

    With lsvMdeduct
        .ColumnHeaders.Add , , "memberno"
        .ColumnHeaders.Add , , "Principal"
        .ColumnHeaders.Add , , "interest"
        .ColumnHeaders.Add , , "shares"
        .ColumnHeaders.Add , , "entrancefee"
        .ColumnHeaders.Add , , "loanform"
        .ColumnHeaders.Add , , "calender"
        .ColumnHeaders.Add , , "passbook"
        .ColumnHeaders.Add , , "othercharges"
        .ColumnHeaders.Add , , "Christmas"
        .View = lvwReport
        .GridLines = True
        
    End With
    
End Sub


Private Sub Picture1_Click()
Dim Z
Dim rs As Recordset
frmsearchrecords.Show vbModal
 Z = strName
    If Z <> "" Then
     lblintcontrol = Z
        
        End If
        
        Set cn = CreateObject("adodb.connection")
    'If accdr = "" Then Exit Sub
   cn.Open Provider, "atm", "atm"
    sql = ""
   sql = "select * from glsetup where glaccname='" & Z & "'"
   Set rs = New ADODB.Recordset
   rs.Open sql, cn, adOpenKeyset, adLockOptimistic
   If Not rs.EOF Then
        If Not IsNull(rs.Fields("glaccname")) Then lblintcontrol = rs.Fields("glaccname")
        If Not IsNull((rs.Fields("ACCNO"))) Then ACCNO2 = rs.Fields("ACCNO")
  End If

End Sub

Private Sub Picture10_Click()
Dim Z
Dim rs As Recordset
frmsearchrecords.Show vbModal
 Z = strName
    If Z <> "" Then
     lblmemberid = Z
        
        End If
        
        Set cn = CreateObject("adodb.connection")
    'If accdr = "" Then Exit Sub
   cn.Open Provider, "atm", "atm"
    sql = ""
   sql = "select * from cub where accno='" & Z & "'"
   Set rs = New ADODB.Recordset
   rs.Open sql, cn, adOpenKeyset, adLockOptimistic
   If Not rs.EOF Then
        If Not IsNull(rs.Fields("glaccname")) Then lblloancontrolacc = rs.Fields("glaccname")
        If Not IsNull((rs.Fields("ACCNO"))) Then ACCNO1 = rs.Fields("ACCNO")
        End If

End Sub

Private Sub Picture2_Click()
'// lblSharescash
Dim Z
Dim rs As Recordset
frmsearchrecords.Show vbModal
 Z = strName
    If Z <> "" Then
     lblSharescash = Z
        
        End If
        
        Set cn = CreateObject("adodb.connection")
    'If accdr = "" Then Exit Sub
    cn.Open Provider, "atm", "atm"
    sql = ""
   sql = "select * from glsetup where glaccname='" & Z & "'"
   Set rs = New ADODB.Recordset
   rs.Open sql, cn, adOpenKeyset, adLockOptimistic
   If Not rs.EOF Then
        If Not IsNull(rs.Fields("glaccname")) Then lblSharescash = rs.Fields("glaccname")
        If Not IsNull((rs.Fields("ACCNO"))) Then ACCNO1 = rs.Fields("ACCNO")
        End If

End Sub

Private Sub Picture3_Click()
Dim Z
Dim rs As Recordset
frmsearchrecords.Show vbModal
 Z = strName
    If Z <> "" Then
     lblLoanForm = Z
        
        End If
        
        Set cn = CreateObject("adodb.connection")
    'If accdr = "" Then Exit Sub
   cn.Open Provider, "atm", "atm"
    sql = ""
   sql = "select * from cub where accno='" & Z & "'"
   Set rs = New ADODB.Recordset
   rs.Open sql, cn, adOpenKeyset, adLockOptimistic
   If Not rs.EOF Then
        If Not IsNull(rs.Fields("name")) Then lblLoanForm = rs.Fields("name")
        If Not IsNull((rs.Fields("ACCNO"))) Then ACCNO1 = rs.Fields("ACCNO")
        End If

End Sub

Private Sub Picture4_Click()
Dim Z
Dim rs As Recordset
frmsearchrecords.Show vbModal
 Z = strName
    If Z <> "" Then
     lblEntranceFee = Z
        
        End If
        
        Set cn = CreateObject("adodb.connection")
    'If accdr = "" Then Exit Sub
    cn.Open Provider, "atm", "atm"
    sql = ""
   sql = "select * from cub where accno='" & Z & "'"
   Set rs = New ADODB.Recordset
   rs.Open sql, cn, adOpenKeyset, adLockOptimistic
   If Not rs.EOF Then
        If Not IsNull(rs.Fields("name")) Then lblEntranceFee = rs.Fields("name")
        If Not IsNull((rs.Fields("ACCNO"))) Then ACCNO1 = rs.Fields("ACCNO")
        End If

End Sub

Private Sub Picture5_Click()
Dim Z
Dim rs As Recordset
frmsearchrecords.Show vbModal
 Z = strName
    If Z <> "" Then
     lblloancontrolacc = Z
        
        End If
        
        Set cn = CreateObject("adodb.connection")
    'If accdr = "" Then Exit Sub
   cn.Open Provider, "atm", "atm"
    sql = ""
   sql = "select * from glsetup where glaccname='" & Z & "'"
   Set rs = New ADODB.Recordset
   rs.Open sql, cn, adOpenKeyset, adLockOptimistic
   If Not rs.EOF Then
        If Not IsNull(rs.Fields("glaccname")) Then lblloancontrolacc = rs.Fields("glaccname")
        If Not IsNull((rs.Fields("ACCNO"))) Then ACCNO1 = rs.Fields("ACCNO")
        End If
End Sub

Private Sub Picture6_Click()
Dim Z
Dim rs As Recordset
frmsearchrecords.Show vbModal
 Z = strName
    If Z <> "" Then
     lblCalender = Z
        
        End If
        
        Set cn = CreateObject("adodb.connection")
    'If accdr = "" Then Exit Sub
  cn.Open Provider, "atm", "atm"
    sql = ""
   sql = "select * from cub where accno='" & Z & "'"
   Set rs = New ADODB.Recordset
   rs.Open sql, cn, adOpenKeyset, adLockOptimistic
   If Not rs.EOF Then
        If Not IsNull(rs.Fields("name")) Then lblCalender = rs.Fields("name")
        If Not IsNull((rs.Fields("ACCNO"))) Then ACCNO1 = rs.Fields("ACCNO")
        End If

End Sub

Private Sub Picture7_Click()
Dim Z
Dim rs As Recordset
frmsearchrecords.Show vbModal
 Z = strName
    If Z <> "" Then
     lblchristmas = Z
        
        End If
        
        Set cn = CreateObject("adodb.connection")
    'If accdr = "" Then Exit Sub
  cn.Open Provider, "atm", "atm"
    sql = ""
   sql = "select * from glsetup where glaccname='" & Z & "'"
   Set rs = New ADODB.Recordset
   rs.Open sql, cn, adOpenKeyset, adLockOptimistic
   If Not rs.EOF Then
        If Not IsNull(rs.Fields("glaccname")) Then lblchristmas = rs.Fields("glaccname")
        If Not IsNull((rs.Fields("ACCNO"))) Then ACCNO1 = rs.Fields("ACCNO")
        End If

End Sub

Private Sub Picture8_Click()
Dim Z
Dim rs As Recordset
frmsearchrecords.Show vbModal
 Z = strName
    If Z <> "" Then
     lblOthercharges = Z
        
        End If
        
        Set cn = CreateObject("adodb.connection")
    'If accdr = "" Then Exit Sub
  cn.Open Provider, "atm", "atm"
    sql = ""
   sql = "select * from glsetup where glaccname='" & Z & "'"
   Set rs = New ADODB.Recordset
   rs.Open sql, cn, adOpenKeyset, adLockOptimistic
   If Not rs.EOF Then
        If Not IsNull(rs.Fields("glaccname")) Then lblOthercharges = rs.Fields("glaccname")
        If Not IsNull((rs.Fields("ACCNO"))) Then ACCNO1 = rs.Fields("ACCNO")
        End If

End Sub

Private Sub Picture9_Click()
Dim Z
Dim rs As Recordset
frmsearchrecords.Show vbModal
 Z = strName
    If Z <> "" Then
     lblPassbook = Z
        
        End If
        
        Set cn = CreateObject("adodb.connection")
    'If accdr = "" Then Exit Sub
  cn.Open Provider, "atm", "atm"
    sql = ""
   sql = "select * from glsetup where glaccname='" & Z & "'"
   Set rs = New ADODB.Recordset
   rs.Open sql, cn, adOpenKeyset, adLockOptimistic
   If Not rs.EOF Then
        If Not IsNull(rs.Fields("glaccname")) Then lblPassbook = rs.Fields("glaccname")
        If Not IsNull((rs.Fields("ACCNO"))) Then ACCNO1 = rs.Fields("ACCNO")
        End If

End Sub
Private Sub getde(NA As String)
Dim myclass As Object
 Set myclass = New cdbase
    Set cn = CreateObject("adodb.connection")
    Provider = myclass.OpenCon
   cn.Open Provider, "atm", "atm"
sql = "select * from cuB where name='" & NA & "'"
Set rs = New ADODB.Recordset

rs.Open sql, cn
If Not rs.EOF Then
If Not IsNull(rs.Fields("accno")) Then glaccno = rs.Fields("accno")
If Not IsNull(rs.Fields("accountname")) Then glnamE1 = rs.Fields("name")
If Not IsNull(rs.Fields("idno")) Then glidno = rs.Fields("idno")
If Not IsNull(rs.Fields("memberno")) Then glmemno = rs.Fields("memberno")
If Not IsNull(rs.Fields("payno")) Then glpayno = rs.Fields("payno")
End If
'glPremium = Scheme_GL_Field(AccountCode, "glPremium")
'bookba = cub_balance(glaccno)
End Sub
