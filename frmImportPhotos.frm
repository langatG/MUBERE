VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmImportPhotos 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Posting Importation"
   ClientHeight    =   7020
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9435
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7020
   ScaleWidth      =   9435
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Close"
      Height          =   375
      Left            =   7920
      TabIndex        =   17
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox txtFile 
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   1200
      TabIndex        =   12
      Top             =   720
      Width           =   4095
   End
   Begin VB.CommandButton cmdSearchFile 
      Caption         =   "Search"
      Height          =   375
      Left            =   5640
      TabIndex        =   11
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "Import"
      Height          =   375
      Left            =   5640
      TabIndex        =   10
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   "Select All"
      Height          =   375
      Left            =   7800
      TabIndex        =   9
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdPost 
      Caption         =   "Post"
      Height          =   375
      Left            =   7800
      TabIndex        =   8
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "update member names"
      Height          =   615
      Left            =   11520
      TabIndex        =   5
      Top             =   5400
      Width           =   3375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Update CustomerBalance"
      Height          =   255
      Left            =   11760
      TabIndex        =   4
      Top             =   2160
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Update jan loan"
      Height          =   495
      Left            =   7800
      TabIndex        =   3
      Top             =   3960
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdProcess 
      Caption         =   "Process"
      Height          =   375
      Left            =   7920
      TabIndex        =   2
      Top             =   5520
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdLoadInterest 
      Caption         =   "Load Interest"
      Height          =   375
      Left            =   7800
      TabIndex        =   1
      Top             =   5040
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker DTPTransdate 
      Height          =   375
      Left            =   5640
      TabIndex        =   0
      Top             =   0
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      Format          =   31653889
      CurrentDate     =   43885
   End
   Begin MSComctlLib.ListView lvwMemberDetails 
      Height          =   4815
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   8493
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "SNO"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Phoneno"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Net"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Year"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Month"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ProgressBar importationProgress 
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   1560
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7560
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "NOTE: ENSURE YOUR DATA DONT HAVE THIS '"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   360
      TabIndex        =   19
      Top             =   1200
      Width           =   4695
   End
   Begin VB.Label Label4 
      Caption         =   "Excel to be in this format: Staff/Mno,Name,Phone,Net,Year,Month"
      Height          =   375
      Left            =   720
      TabIndex        =   18
      Top             =   240
      Width           =   4095
   End
   Begin VB.Label lbSourceFile 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000004&
      Caption         =   "Source File:"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label1 
      Height          =   375
      Left            =   7800
      TabIndex        =   15
      Top             =   1200
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   495
      Left            =   4080
      TabIndex        =   14
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000004&
      Caption         =   "Date:"
      Height          =   375
      Left            =   4800
      TabIndex        =   13
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmImportPhotos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdImport_Click()
On Error GoTo ErrorHandler
Dim Graders As String

If Trim(txtFile) = "" Then
    MsgBox "Please Select The File To Import", vbInformation, Me.Caption
Exit Sub
End If
 With lvwMemberDetails
    
       .ListItems.Clear
    
       .ColumnHeaders.Clear

  End With
                     lvwMemberDetails.ColumnHeaders.Add , , "SNO"
                     lvwMemberDetails.ColumnHeaders.Add , , "Name"
                     lvwMemberDetails.ColumnHeaders.Add , , "Phoneno", 2000
                     lvwMemberDetails.ColumnHeaders.Add , , "Net", 2000
                     lvwMemberDetails.ColumnHeaders.Add , , "Year", 3000
                     lvwMemberDetails.ColumnHeaders.Add , , "Month.", 2000
'                     lvwMemberDetails.ColumnHeaders.Add , , "Status", 2000
'lvwMemberDetails.ListItems.Clear
Dim totals As Double

Dim tam As Double
    Dim PrFSO As New FileSystemObject, sData As String, sno As String, companycode As String, StaffNo As String, _
     name As String, PrFile As TextStream, idno As Long, DEPT As Long, mtype As String, Period As Date, _
     myrec As Long, MyPos As Long, mycase6 As String, tripno As Integer, MisFso As New FileSystemObject, FndFso As New FileSystemObject, MisFile As TextStream, _
    FndFile As TextStream, ApplicDate As Date, OtherNames As String, Year As Integer, month As Integer, Photo As String, sex As String, auditid As String, NetPay As Double
    
    Set PrFile = PrFSO.OpenTextFile(txtFile, ForReading, False)
    
    Do Until PrFile.AtEndOfStream
      Dim Add As String
      Add = ","
        myrec = myrec + 1
        sData = PrFile.ReadLine
        sData = sData + Add
'    Set PrFile = PrFSO.OpenTextFile(txtFile, ForReading, False)
'
'    Do Until PrFile.AtEndOfStream
'      Dim Add As String
'      Add = ","
'        myrec = myrec + 1
'        sData = PrFile.ReadLine
'         sData = sData + Add
         
    Loop
    importationProgress.max = myrec
    Set PrFile = PrFSO.OpenTextFile(txtFile, ForReading, False)
 Do Until PrFile.AtEndOfStream
        sData = PrFile.ReadLine
         sData = sData + Add
        MyPos = MyPos + 1
        importationProgress.value = MyPos
        DoEvents
        Do Until InStr(1, sData, Chr(44), vbBinaryCompare) < 0
            MyField = MyField + 1
            If InStr(1, sData, Chr(44), vbBinaryCompare) = 0 Then
                MyField = 0

                        Set li = lvwMemberDetails.ListItems.Add(, , sno)
                        
                         'li.SubItems(2) = transdate
                         li.SubItems(1) = name
                         li.SubItems(2) = PhoneNo
                         li.SubItems(3) = Net
                         li.SubItems(4) = Year
                         li.SubItems(5) = month
'                         li.SubItems(6) = status
'                         li.SubItems(7) = ApplicDate
'                         li.SubItems(8) = AuditID
'                       li.SubItems(9) = DEPT
'                         li.SubItems(10) = Mheader

               
                
                Exit Do
            Else
                Select Case MyField
                   Case 1
                      sno = Left(sData, InStr(1, sData, Chr(44), vbBinaryCompare) - 1)
                   Case 2 'ApplicDate
                       name = Left(sData, InStr(1, sData, Chr(44), vbBinaryCompare) - 1)
                   Case 3 'CompanyCode
                      PhoneNo = Left(sData, InStr(1, sData, Chr(44), vbBinaryCompare) - 1)
                    Case 4 'StaffNo
                       Net = Left(sData, InStr(1, sData, Chr(44), vbBinaryCompare) - 1)

                   Case 5 'IDNo
                       Year = Left(sData, InStr(1, sData, Chr(44), vbBinaryCompare) - 1)
                   Case 6 'Surname
                       month = Left(sData, InStr(1, sData, Chr(44), vbBinaryCompare) - 1)
                   Case 7 'OtherNames
'                        status = Left(sData, InStr(1, sData, Chr(44), vbBinaryCompare) - 1)
'                   Case 8 'Sex
'                        Sex = Left(sData, InStr(1, sData, Chr(44), vbBinaryCompare) - 1)
'                   Case 8 'ApplicDate
'                       ApplicDate = Left(sData, InStr(1, sData, Chr(44), vbBinaryCompare) - 1)
'                   Case 9 'AuditID
'                       auditid = Left(sData, InStr(1, sData, Chr(44), vbBinaryCompare) - 1)
'                   Case 10 'DEPT
'                        DEPT = Left(sData, InStr(1, sData, Chr(44), vbBinaryCompare) - 1)
'                   Case 11 'MTYPE
'                       MType = Left(sData, InStr(1, sData, Chr(44), vbBinaryCompare) - 1)
'''                   Case 12 'gender
'''                        gender = Left(sData, InStr(1, sData, Chr(44), vbBinaryCompare) - 1)
'''                   Case 13 'next of kin
''                        nextOfKin = Left(sData, InStr(1, sData, Chr(44), vbBinaryCompare) - 1)
                End Select
                sData = Right(sData, Len(sData) - InStr(1, sData, Chr(44), vbBinaryCompare))
            End If
        Loop
    Loop
    MsgBox "Records Imported Successfully", vbInformation, Me.Caption
    
    
    Exit Sub
ErrorHandler:
MsgBox err.description
End Sub

 

Private Sub cmdPost_Click()
  On Error GoTo ErrorHandler
Dim totals As Double
 Dim name As String
 Dim month As Integer
 Dim Year As Integer
totals = 0
   'mMonth = month(DTPTransdate)
   'yYear = year(DTPTransdate)

 'oSaccoMaster.Execute "truncate  table Payout"
 'importationProgress.Min = 0
 'importationProgress.Max = 1000


'importationProgress.value
  If lvwMemberDetails.ListItems.Count < 1 Then
     MsgBox "  No Records to be posted", vbInformation
        Exit Sub
    End If
'      If lvwMemberDetails.ListItems.Item(I).Checked = False Then
'        MsgBox "  No Records selected", vbInformation
'        End If
        
       MyPos = importationProgress.value
       
  For I = lvwMemberDetails.ListItems.Count To 1 Step -1
        If lvwMemberDetails.ListItems.Item(I).Checked = True Then
          MyPos = MyPos - 1
           importationProgress.value = MyPos
            Set li = lvwMemberDetails.ListItems(I)
               'transdate = li
                sno = lvwMemberDetails.ListItems(I)
                name = lvwMemberDetails.ListItems(I).SubItems(1)
                PhoneNo = lvwMemberDetails.ListItems(I).SubItems(2)
                Net = lvwMemberDetails.ListItems(I).SubItems(3)
                Year = lvwMemberDetails.ListItems(I).SubItems(4)
                month = lvwMemberDetails.ListItems(I).SubItems(5)
               ' chequeno = lvwMemberDetails.ListItems(I).SubItems(6)

        sql = ""
        sql = "set dateformat dmy  INSERT INTO   d_PayrollCopy"
        sql = sql & " (SNo, Names, PhoneNo, NetPay, Yyear, Mmonth,Status1,User1)"
        sql = sql & "VALUES ('" & lvwMemberDetails.ListItems(I) & "','" & lvwMemberDetails.ListItems(I).SubItems(1) & "','" & lvwMemberDetails.ListItems(I).SubItems(2) & "','" & lvwMemberDetails.ListItems(I).SubItems(3) & "','" & lvwMemberDetails.ListItems(I).SubItems(4) & "','" & lvwMemberDetails.ListItems(I).SubItems(5) & "','1','" & User & "')"
        ','" & lvwMemberDetails.ListItems(I).SubItems(4) & "','" & lvwMemberDetails.ListItems(I).SubItems(5) & "','" & lvwMemberDetails.ListItems(I).SubItems(6) & "','" & lvwMemberDetails.ListItems(I).SubItems(7) & "')"
        ''" & Amount & "' ,'" & DrAccNo & "','" & crAccNo & "','" & DocumentNo & "','" & crAccNo & "'"
'        sql = sql & "'" & txtMonth.Text & "', " & txtyear.Text & "," & txtPosted.Text & ",'" & Now & "')"
        
        oSaccoMaster.Execute (sql)
            
           'sql = "  update MEMBERS set Photo='" & Photo & "'    where MemberNo = '" & sno & "'"
            ' If Not oSaccoMaster.Execute(sql) Then
             '       GoTo errorhandler
               ' End If
                         
                         lvwMemberDetails.ListItems.Remove (lvwMemberDetails.ListItems(I).Index)
    
          Else
              MsgBox "  No Records selected", vbInformation
               Exit Sub
        End If
'importationProgress.value -1
     'importationProgress.value = MyPos
  Next I
        
        'oSaccoMaster.Execute " UPDATE   MEMBERS  SET   BranchCode =1"
         MsgBox "Records Imported", vbInformation
         cmdSelectAll.Caption = "Select All"
    Exit Sub
ErrorHandler:
MsgBox err.description
End Sub

Private Sub cmdSearchFile_Click()
With CommonDialog1
        .Filter = "Comma Seperated Values|*.csv"
        .ShowOpen
        If .FileName <> "" Then
            txtFile = .FileName
        Else
            MsgBox "No file selected", vbInformation, Me.Caption
            txtFile = ""
            Exit Sub
        End If
        .FileName = ""
    End With
End Sub

Private Sub cmdSelectAll_Click()
If lvwMemberDetails.ListItems.Count = 0 Then
  MsgBox "Please there is no item to select/Deselect", vbExclamation
Else
  If cmdSelectAll.Caption = "Select All" Then
       cmdSelectAll.Caption = "Deselect All"
       For I = 1 To lvwMemberDetails.ListItems.Count
             If lvwMemberDetails.ListItems.Item(I).Checked = False Then
                 lvwMemberDetails.ListItems.Item(I).Checked = True
             Else
                 lvwMemberDetails.ListItems.Item(I).Checked = False
             End If
       Next I
       
  ElseIf cmdSelectAll.Caption = "Deselect All" Then
        cmdSelectAll.Caption = "Select All"
        For I = 1 To lvwMemberDetails.ListItems.Count
             If lvwMemberDetails.ListItems.Item(I).Checked = False Then
                 lvwMemberDetails.ListItems.Item(I).Checked = True
             Else
                 lvwMemberDetails.ListItems.Item(I).Checked = False
             End If
       Next I
        
  End If
End If
End Sub
'get grader number
Private Function getGraderNo(subroute1 As String) As String
   Set rst1 = oSaccoMaster.GetRecordset("select LCode from d_Location where LName = '" & subroute1 & "' ")
    If Not rst1.EOF Then
        Set rst = oSaccoMaster.GetRecordset("select graderNo from d_GraderRoutesAllocation where locationId = '" & rst1(0) & "'")
          If Not rst.EOF Then
              getGraderNo = rst(0)
              Exit Function
          Else
              MsgBox "please Allocate grader to route : '" & subroute & "'", vbExclamation
          End If
    Else
            MsgBox "Please the Route is Not available", vbExclamation
    End If

End Function

'get grader number
Private Function getGraderNumber(subroute As String) As Integer
   Set rst1 = oSaccoMaster.GetRecordset("select  graderNo from d_Graders where Route = '" & subroute & "'")
     If Not rst1.EOF Then
         getGraderNumber = rst1(0)
     Else
           MsgBox "Please SubRoute Has not been allocated to SubRoute:>> '" & subroute & "'", vbExclamation
     End If
End Function


Private Sub Command1_Click()
  Dim rstww As New ADODB.Recordset
    Set rstww = oSaccoMaster.GetRecordset(" select m.mno , m.surNaeme,  m.First +' '+m.second as names  from  mweamem m ")
       If Not rstww.EOF Then
          While Not rstww.EOF
            oSaccoMaster.Execute " update members set Surname = '" & rstww(1) & "',OtherNames = '" & rstww(2) & "'  where MemberNo= '" & rstww(0) & "' "
              'Label1.Caption = rstww(0)
            rstww.MoveNext
             'Label1.Caption = rstww(0)
           Wend
       End If
      
     MsgBox " done"

End Sub

Private Sub Command2_Click()
   Dim rscub As New ADODB.Recordset
      'Set rscub = oSaccoMaster.GetRecordSet("SELECT      AccNo, Balance FROM         CUB WHERE     (AccountCode = 07)")
     Set rscub = oSaccoMaster.GetRecordset("SELECT     AccNo, Amount FROM         CUSTOMERBALANCE  WHERE     (AccNo LIKE '107%')")

       If Not rscub.EOF Then
          While Not rscub.EOF
          oSaccoMaster.Execute ("Update  CUB  set Balance='" & rscub(1) & "' where AccNo='" & rscub(0) & "' and AccountCode = 07")
         rscub.MoveNext
         Wend
       End If
   MsgBox "This is Good Place!"
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Command5_Click()

sql = ""
sql = "set dateformat dmy (SELECT     TOP (2) TransDate, Amount, DrAccNo, CrAccNo, DocumentNo, TransDescript, ChequeNo From GLTRANSACTIONS Where (auditid = 'mary'))"
cn.Execute sql
'Export

End Sub

Private Sub Form_Load()
  DTPTransdate = Format(Get_Server_Date, "dd/mm/yyyy")
End Sub

Private Sub mnuPayoutReport_Click()
  STRFORMULA = ""
        reportname = "Payout List.rpt"
        Show_Sales_Crystal_Report STRFORMULA, reportname, CompanyName
End Sub


