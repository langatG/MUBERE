VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim cnnPayroll As Connection
Dim rsGroups As Recordset
Dim rsSystUsers As Recordset
Dim rsCUser As Recordset
Dim Form As Object
Dim i As Object
Public EmpModify As Boolean
Public BenModify As Boolean
Public DedModify As Boolean
Public OtherModify As Boolean
Public SetupModify As Boolean
Public CurrentUser As String
Dim Ex As String
Dim Pwd As Variant
Dim Temp As String, PwdChr As Long
Dim EncryptKey As Long


Private Sub Class_Initialize()
On Error GoTo errorhandler
Set cnnPayroll = New Connection
Set rsGroups = New Recordset
Set rsSystUsers = New Recordset
Set rsCUser = New Recordset

cnnPayroll.Open modCommon.pConnection
rsGroups.Open "Select * from Groups order by GNo", cnnPayroll, adOpenKeyset, adLockOptimistic
rsSystUsers.Open "Select * from Security order by UID", cnnPayroll, adOpenKeyset, adLockOptimistic
rsCUser.Open "Select * from CUser", cnnPayroll, adOpenKeyset, adLockOptimistic

Call GlobSecurity
Call GetUser
Exit Sub
errorhandler:
MsgBox Err.Description
End Sub



Public Sub GlobSecurity()


With rsCUser
    If .RecordCount > 0 Then
        .MoveFirst
        If Not !UName = "" Then
            With rsSystUsers
                If .RecordCount > 0 Then
                    .MoveFirst
                    .Find "UID like '*" & rsCUser!UName & "*'", , adSearchForward, adBookmarkFirst
                    If Not .EOF Then
                        
                        Ex = rsSystUsers!Excc
                        
                        Pwd = Ex
                        EncryptKey = Int(Sqr(Len(Pwd) * 81)) + 23
                        
                        For PwdChr = 1 To Len(Pwd)
                            Temp = Temp + Chr(Asc(Mid(Pwd, PwdChr, 1)) Xor EncryptKey)
                        Next PwdChr
                        
                        EncryptPass = Temp

                        Ex = EncryptPass
                             
                        If !Exc = True And Ex = 1 Then

                            
                            SetupModify = True
                            OtherModify = True
                            DedModify = True
                            BenModify = True
                            EmpModify = True
                           
                            
                        Else
                            If Not !GNo = "" Then
                                With rsGroups
                                    If .RecordCount > 0 Then
                                        .MoveFirst
                                        .Find "GNo like '*" & rsSystUsers!GNo & "*'", , adSearchForward, adBookmarkFirst
                                        If Not .EOF Then
    '
    
                                            If !Setup = "Modify" Then
                                                SetupModify = True
                                            ElseIf !Setup = "View" Then
                                                SetupModify = False
                                            Else
                                               
                                            End If
                                            
                                            If !Other = "Modify" Then
                                                OtherModify = True
                                            ElseIf !Other = "View" Then
                                                OtherModify = False
                                            Else

                                            End If
                                            
                                            If !Ded = "Modify" Then
                                                DedModify = True
                                            ElseIf !Ded = "View" Then
                                                DedModify = False
                                            Else

                                            End If
                                            
                                            If !Ben = "Modify" Then
                                                BenModify = True
                                            ElseIf !Ben = "View" Then
                                                BenModify = False
                                            Else

                                            End If
                                            
                                            If !Emp = "Modify" Then
                                                EmpModify = True
                                            ElseIf !Emp = "View" Then
                                                EmpModify = False
    
                                            Else

                                            End If
    
    
                                        End If
                                    End If
                                End With
    
    
                            End If
                    
                        End If
                        
                      

                    End If
                End If
            End With
        End If
    End If
End With
End Sub

Public Sub GetUser()
With rsCUser
    If .RecordCount > 0 Then
        .MoveFirst
        If Not !UName = "" Then
            CurrentUser = !UName & ""
        End If
    End If
End With
        
        
End Sub

Private Sub Class_Terminate()
    'cnnPayroll.Close
End Sub


Function EEncryptPassword()
Dim Pwd As Variant
Dim Temp As String, PwdChr As Long
Dim EncryptKey As Long
Pwd = Ex
EncryptKey = Int(Sqr(Len(Pwd) * 81)) + 23

For PwdChr = 1 To Len(Pwd)
    Temp = Temp + Chr(Asc(Mid(Pwd, PwdChr, 1)) Xor EncryptKey)
Next PwdChr

EncryptPass = Temp

End Function
