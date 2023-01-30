VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FRMQBMPSIMPORTS 
   Caption         =   "QBMPSIMPORTS"
   ClientHeight    =   4680
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   ScaleHeight     =   4680
   ScaleWidth      =   7335
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSearchFile 
      Caption         =   "Search"
      Height          =   375
      Left            =   5280
      TabIndex        =   6
      Top             =   720
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DTPTransdate 
      Height          =   375
      Left            =   4320
      TabIndex        =   5
      Top             =   120
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      Format          =   124583937
      CurrentDate     =   42748
   End
   Begin VB.TextBox txtFile 
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   1080
      TabIndex        =   1
      Top             =   720
      Width           =   4095
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "Import"
      Height          =   375
      Left            =   2880
      TabIndex        =   0
      Top             =   2280
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar importationProgress 
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5880
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lbSourceFile 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000004&
      Caption         =   "Source File:"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000004&
      Caption         =   "Date:"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "FRMQBMPSIMPORTS"
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

'If ChkRefno.value = vbUnchecked And ChkAccno.value = vbUnchecked Then
'    MsgBox "Please checked one of the option", vbInformation, Me.Caption
'Exit Sub
'End If

'lvwMemberDetails.ListItems.Clear
sql = "delete from qbmps"
If Not oSaccoMaster.Execute(sql) Then
    GoTo ErrorHandler
End If
Dim totals As Double

Dim tam As Double
    Dim PrFSO As New FileSystemObject, sData As String, dateenntered As Date, canno As String, Tpc As Double, Tss As Double, ads As Double, anr As Double, Remarks As String, score As Double, Company As String, companycode As String, memAccno As String, StaffNo As String, _
     NAMES As String, PrFile As TextStream, idno As Long, DEPT As Long, mtype As String, Period As Date, _
     myrec As Long, MyPos As Long, mycase6 As String, tripno As Integer, MisFso As New FileSystemObject, FndFso As New FileSystemObject, MisFile As TextStream, _
    FndFile As TextStream, ApplicDate As Date, OtherNames As String, Photo As String, sex As String, auditid As String, NetPay As Double
    
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
    importationProgress.Max = myrec
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
'                        Set li = lvwMemberDetails.ListItems.Add(, , memAccno)
'                            If ChkRefno.value = vbChecked Then
'                            sno = getrefno(memAccno)
'                            Else
'                           sno = Getmembrno(memAccno)
'                            End If
'                         li.SubItems(1) = sno
                        ' li.SubItems(2) = names
'                       li.SubItems(2) = NetPay
                          'li.SubItems(4) = Netpay
'                         li.SubItems(5) = OtherNames
'                         li.SubItems(6) = Sex
'                         li.SubItems(7) = ApplicDate
'                         li.SubItems(8) = AuditID
'                       li.SubItems(9) = DEPT
'                         li.SubItems(10) = Mheader

               
                
                Exit Do
            Else
                Select Case MyField
                   Case 1 'MemberNo
                        dateenntered = Left(sData, InStr(1, sData, Chr(44), vbBinaryCompare) - 1)
                   Case 2 'CompanyCode
                      canno = Left(sData, InStr(1, sData, Chr(44), vbBinaryCompare) - 1)
                    Case 3 'StaffNo
                       Tpc = Left(sData, InStr(1, sData, Chr(44), vbBinaryCompare) - 1)

                   Case 4 'IDNo
                       Tss = Left(sData, InStr(1, sData, Chr(44), vbBinaryCompare) - 1)
                   Case 5 'Surname
                       ads = Left(sData, InStr(1, sData, Chr(44), vbBinaryCompare) - 1)
                   Case 6 'OtherNames
                        anr = Left(sData, InStr(1, sData, Chr(44), vbBinaryCompare) - 1)
                   
                   Case 7 'ApplicDate
                       Remarks = Left(sData, InStr(1, sData, Chr(44), vbBinaryCompare) - 1)
                   Case 8 'AuditID
                       score = Left(sData, InStr(1, sData, Chr(44), vbBinaryCompare) - 1)
                   Case 9 'DEPT
                        Company = Left(sData, InStr(1, sData, Chr(44), vbBinaryCompare) - 1)
'                   Case 11 'MTYPE
'                       MType = Left(sData, InStr(1, sData, Chr(44), vbBinaryCompare) - 1)
'''                   Case 12 'gender
'''                        gender = Left(sData, InStr(1, sData, Chr(44), vbBinaryCompare) - 1)
'''                   Case 13 'next of kin
''                        nextOfKin = Left(sData, InStr(1, sData, Chr(44), vbBinaryCompare) - 1)
                sql = "set dateformat dmy Insert Into QBMPS(Date, canno, Tpc, TSC, ALC, anr, remarks, Pscore, cname) VALUES     ('" & dateenntered & "','" & canno & "', '" & Tpc & "', '" & Tss & "', '" & ads & "', '" & anr & "', '" & Remarks & "', '" & score & "', '" & Company & "')"
                If Not oSaccoMaster.Execute(sql) Then
                    GoTo ErrorHandler
                End If
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

