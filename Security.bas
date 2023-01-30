Attribute VB_Name = "Security"
'Dim cnnPayroll As Connection
'Dim rsGroups As Recordset
'Dim rsSystUsers As Recordset
'Dim rsCUser As Recordset
'Dim Form As String
'
'
'Private Sub Class_Initialize()
'Set cnnPayroll = New Connection
'Set rsGroups = New Recordset
'Set rsSystUsers = New Recordset
'Set rsCUser = New Recordset
'
'cnnPayroll.Open MDBase
'rsGroups.Open "Select * from Groups order by GNo", cnnPayroll, adOpenKeyset, adLockOptimistic
'rsSystUsers.Open "Select * from Security order by UID", cnnPayroll, adOpenKeyset, adLockOptimistic
'rsCUser.Open "Select * from CUser", cnnPayroll, adOpenKeyset, adLockOptimistic
'
'End Sub
'
'Public Sub Disbcmd()
'Dim i As Object
'    For Each i In Me
'        If TypeOf i Is CommandButton Then
'            i.Enabled = False
'        End If
'    Next i
'End Sub
'
'
'Public Sub GlobalSecurity()
'
'
'With rsCUser
'    If .RecordCount > 0 Then
'        .MoveFirst
'        If Not !UName = "" Then
'            With rsSystUsers
'                If .RecordCount > 0 Then
'                    .MoveFirst
'                    .Find "UID like '*" & rsCUser!UName & "*'", , adSearchForward, adBookmarkFirst
'                    If Not .EOF Then
'                        If !GNo = "" Then
'                            With rsGroups
'                                If .RecordCount > 0 Then
'                                    .MoveFirst
'                                    .Find "GNo like '*" & rsSystUsers!GNo & "*'", , adSearchForward, adBookmarkFirst
'                                    If Not .EOF Then
'                                        If !Emp = "Modify" Then
'
'                                        ElseIf !Emp = "View" Then
'                                            Form = "frmEmp"
'                                            Call Disbcmd
'                                        Else
'                                            frmMain.CoolBar1.Bands(1).Visible = False
'                                        End If
''
'                                    End If
'                                End If
'                            End With
'
'
'                        End If
'
'                    End If
'                End If
'            End With
'        End If
'    End If
'End With
'
'End Sub
