VERSION 5.00
Begin VB.Form frmSQLSRVBackup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Backup SQL Server Database"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10005
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSQLSRVBackup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   10005
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraMain 
      Caption         =   "Backup SQL Server Database"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   4575
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   9615
      Begin VB.CommandButton cmdBackup 
         Caption         =   "&Backup"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   5520
         TabIndex        =   9
         Top             =   3750
         Width           =   1335
      End
      Begin VB.Frame fraPathDetails 
         Caption         =   "Path && File Information"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   1400
         Left            =   4320
         TabIndex        =   21
         Top             =   1320
         Width           =   5055
         Begin VB.TextBox txtFileName 
            Height          =   315
            Left            =   1200
            MaxLength       =   30
            TabIndex        =   8
            Top             =   870
            Width           =   3135
         End
         Begin VB.CommandButton cmdGetPath 
            Caption         =   "..."
            Height          =   375
            Left            =   4395
            TabIndex        =   7
            Top             =   360
            Width           =   375
         End
         Begin VB.TextBox txtPath 
            Height          =   315
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   23
            Top             =   390
            Width           =   3135
         End
         Begin VB.Label lblFileName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "File Name:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   24
            Top             =   930
            Width           =   750
         End
         Begin VB.Label lblSelectPath 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Path:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   22
            Top             =   450
            Width           =   390
         End
      End
      Begin VB.Frame fraDatabase 
         Caption         =   "Database Information"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   855
         Left            =   4320
         TabIndex        =   19
         Top             =   360
         Width           =   5055
         Begin VB.ComboBox cboDBName 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   320
            Width           =   3615
         End
         Begin VB.Label lblDatabase 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Database:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   20
            Top             =   375
            Width           =   750
         End
      End
      Begin VB.Frame fraConnection 
         Caption         =   "Connection"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   855
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   3855
         Begin VB.TextBox txtServer 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1320
            TabIndex        =   0
            Top             =   300
            Width           =   2295
         End
         Begin VB.Label lblServer 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Server:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   18
            Top             =   360
            Width           =   540
         End
      End
      Begin VB.Frame fraAction 
         Caption         =   "Action"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   855
         Left            =   240
         TabIndex        =   14
         Top             =   3480
         Width           =   3855
         Begin VB.CommandButton cmdClose 
            Caption         =   "C&lose"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2760
            TabIndex        =   16
            Top             =   320
            Width           =   975
         End
         Begin VB.CommandButton cmdDisconnect 
            Caption         =   "&Disconnect"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1440
            TabIndex        =   15
            Top             =   320
            Width           =   1215
         End
         Begin VB.CommandButton cmdConnect 
            Caption         =   "&Connect"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   5
            Top             =   320
            Width           =   1215
         End
      End
      Begin VB.Frame fraAuthentication 
         Caption         =   "Authentication"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   2055
         Left            =   240
         TabIndex        =   2
         Top             =   1320
         Width           =   3855
         Begin VB.OptionButton optSQLAuth 
            Caption         =   "Use S&QL Server authentication"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   120
            TabIndex        =   11
            Top             =   720
            Width           =   3135
         End
         Begin VB.OptionButton optWinNTAuth 
            Caption         =   "Use Windows &NT authentication"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   120
            TabIndex        =   10
            Top             =   360
            Width           =   3015
         End
         Begin VB.TextBox txtPassword 
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   1320
            PasswordChar    =   "*"
            TabIndex        =   4
            Top             =   1500
            Width           =   2292
         End
         Begin VB.TextBox txtLoginName 
            Height          =   315
            Left            =   1320
            TabIndex        =   3
            Top             =   1050
            Width           =   2292
         End
         Begin VB.Label lblPassword 
            AutoSize        =   -1  'True
            Caption         =   "Password:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   1560
            Width           =   750
         End
         Begin VB.Label lblLoginName 
            AutoSize        =   -1  'True
            Caption         =   "Login name:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   1110
            Width           =   870
         End
      End
      Begin VB.Label lblAuthInfo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "xx"
         ForeColor       =   &H00C0FFFF&
         Height          =   195
         Left            =   4830
         TabIndex        =   25
         Top             =   3600
         Width           =   4035
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "frmSQLSRVBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim blnConnect As Boolean
Dim AdoCn As New ADODB.Connection
Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

Const BIF_RETURNONLYFSDIRS = &H1      'Only file system directories
Const MAX_PATH = 260

Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
'   Variables for Get short pathname.
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal lBuffer As Long) As Long

'   Purpose:    Displays the Error Messages.
Private Sub ErrorMessage()
    Screen.MousePointer = vbDefault
    MsgBox "Error No:   " & err.number & vbCrLf & _
            "Message:  " & err.description, vbCritical
End Sub

'   The following function is used to restrict the FileName contains Invalid Characters.
Private Function FileNameValidation(ByVal KeyAscii As Integer) As Long
    Dim strInValid As String
    strInValid = "\/:*?<>|."""
    If InStr(strInValid, Chr(KeyAscii)) > 0 Then Exit Function
    FileNameValidation = KeyAscii
End Function

Private Sub FillDBList()
    On Error GoTo ErrorHandler
    Dim AdoRs As New ADODB.Recordset
    Dim strSQL As String
    Dim lngR As Long
    Set AdoRs = Nothing
    strSQL = vbNullString
    With cboDBName
        .Clear
        .AddItem vbNullString
    End With
    '   Display one User Created Database(s).
    strSQL = "Select DBID, NAME From MASTER..SYSDATABASES "
'    strSQL = strSQL & "Where Name Not In ('master', 'model', 'msdb', 'tempdb')"
    With AdoRs
        .CursorLocation = adUseClient
        .ActiveConnection = AdoCn
        .Open strSQL, , adOpenForwardOnly, adLockReadOnly
        .Sort = "Name Asc"
        For lngR = 1 To .RecordCount
            cboDBName.AddItem Trim(!name & "")
            .MoveNext
        Next
        Set .ActiveConnection = Nothing
    End With
    Set AdoRs = Nothing
    strSQL = vbNullString
    Exit Sub
ErrorHandler:
    Call ErrorMessage
End Sub

Private Function GetShortPath(strFileName As String) As String
    Dim lngRes As Long, strPath As String
    'Create a buffer
    strPath = String$(165, 0)
    'retrieve the short pathname
    lngRes = GetShortPathName(strFileName, strPath, 164)
    'remove all unnecessary chr$(0)'s
    GetShortPath = Left$(strPath, lngRes)
End Function

Private Sub SQLAuthOptionsOn()
    lblLoginName.Enabled = True
    lblPassword.Enabled = True
    txtLoginName.Enabled = True
    txtpassword.Enabled = True
    txtpassword.Visible = True
    txtLoginName.Visible = True
    lblLoginName.Visible = True
    lblPassword.Visible = True
    txtLoginName = "sa"
End Sub

Private Sub WinNTAuthOptionsOn()
    lblLoginName.Enabled = False
    lblPassword.Enabled = False
    lblLoginName.Visible = False
    lblPassword.Visible = False
    With txtLoginName
        .Enabled = False
        .Text = vbNullString
        .Visible = False
    End With
    With txtpassword
        .Enabled = False
        .Text = vbNullString
        .Visible = False
    End With
End Sub

Private Sub cmdBackup_Click()
    On Error GoTo ErrorHandler
    Dim AdoCmd As New ADODB.Command
    Dim strDPath As String      ' Destination path.
    Dim strExecute As String    ' For execution of Backup Database Command.
    '   Validation for Database.
    With cboDBName
        If Len(Trim(.Text)) = 0 Then
            MsgBox "Select Database From List.", vbInformation
            If .Enabled Then .SetFocus
            Exit Sub
        End If
    End With
    '   Validation for Path.
    With txtPath
        If Len(Trim(.Text)) = 0 Then
            MsgBox "Select Path From Browser.", vbInformation
            If cmdGetPath.Enabled Then cmdGetPath.SetFocus
            Exit Sub
        End If
    End With
    '   Validation for File Name.
    With txtFileName
        If Len(Trim(.Text)) = 0 Then
            MsgBox "Enter File Name.", vbInformation
            If .Enabled Then .SetFocus
            Exit Sub
        End If
    End With
    '   Collect the required information.
    If Right(Trim(txtPath.Text), 1) <> "\" Then
        strDPath = GetShortPath(Trim(txtPath.Text)) & "\" & Trim(txtFileName.Text)
    Else
        strDPath = GetShortPath(Trim(txtPath.Text) & Trim(txtFileName.Text))
    End If
    DoEvents
    If Right(Trim(txtPath.Text), 1) <> "\" Then
        strDPath = Trim(txtPath.Text) & "\" & Trim(txtFileName.Text)
    Else
        strDPath = Trim(txtPath.Text) & Trim(txtFileName.Text)
    End If
'    MsgBox strDPath
    ' Delete the datafile to allow the application to create a brand new file.
    ' This will prevent attaching the new backup data to the old data if there
    ' is any.
    If Len(Dir(Trim(strDPath))) > 0 Then
        Kill strDPath
    End If
    '   Starts Backup process.
    Screen.MousePointer = vbHourglass
    strExecute = "BACKUP DATABASE [" & Trim(cboDBName.Text) & "] "
    strExecute = strExecute & "TO DISK = '" & Trim(strDPath) & "' "
    strExecute = strExecute & "WITH INIT" ', STATS "
    With AdoCmd
        .ActiveConnection = AdoCn
        .CommandType = adCmdText
        .CommandTimeout = 0
        .CommandText = "use master"
        .Execute
        .CommandText = strExecute
        .Execute
        Set .ActiveConnection = Nothing
    End With
    Set AdoCmd = Nothing
    Screen.MousePointer = vbDefault
    MsgBox "Backup Completed Successfully.", vbInformation
    Exit Sub
ErrorHandler:
    Set AdoCmd = Nothing
    Call ErrorMessage
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdConnect_Click()
    On Error GoTo ErrorHandler
    Dim strServerName As String, strLoginName As String, strPassword As String
    Dim strADOCnn As String
    Set AdoCn = Nothing
    '   Validation for Server Name.
    With txtServer
        If Len(Trim(.Text)) = 0 Then
            MsgBox "Enter Server Name.", vbInformation
            If .Enabled Then .SetFocus
            Exit Sub
        End If
    End With
    '   Validation for Login Name only when SQL Authentaction.
    If optSQLAuth.value = True Then
        With txtLoginName
            If Len(Trim(.Text)) = 0 Then
                MsgBox "Enter Login User Name.", vbInformation
                If .Enabled Then .SetFocus
                Exit Sub
            End If
        End With
    End If
    '   Collect the connection information into local variables.
    strServerName = Trim(txtServer.Text)
    strLoginName = Trim(txtLoginName.Text)
    strPassword = Trim(txtpassword.Text)
    strADOCnn = "Provider=SQLOLEDB.1;"
    '   For Windows NT Authentication.
    If optWinNTAuth.value = True Then
        strADOCnn = strADOCnn & "Integrated Security=SSPI;"
        strADOCnn = strADOCnn & "Persist Security Info=False;"
        strADOCnn = strADOCnn & "Trusted_Connection=yes;"
    End If
    '   For SQL Server Authentication.
    If optSQLAuth.value = True Then
        strADOCnn = strADOCnn & "Persist Security Info=False;"
        strADOCnn = strADOCnn & "User ID=" & strLoginName & ";Password=" & strPassword & ";"
    End If
    strADOCnn = strADOCnn & "Data Source=" & strServerName
    Screen.MousePointer = vbHourglass
    '   Connection establishment to Server.
    With AdoCn
        If .State = adStateOpen Then .Close
        .ConnectionString = strADOCnn
'        .ConnectionTimeout = 30
        .CursorLocation = adUseClient
        .Open
    End With
    blnConnect = True
    '   Now filling the Databases.
    Call FillDBList
    Screen.MousePointer = vbDefault
    MsgBox "Connection made to Server successfully.", vbInformation
    fraConnection.Enabled = False
    fraAuthentication.Enabled = False
    cmdConnect.Enabled = False
    cmdDisconnect.Enabled = True
    fraDatabase.Enabled = True
    fraPathDetails.Enabled = True
    cmdBackup.Enabled = True
    cboDBName.SetFocus
    Exit Sub
ErrorHandler:
    Set AdoCn = Nothing
    Screen.MousePointer = vbDefault
    If err.number = "-2147467259" Then      '   ADO Error No.
        MsgBox "Specified SQL Server does not exist or access denied...!", vbExclamation
        If txtServer.Enabled Then txtServer.SetFocus
        Exit Sub
    End If
    If err.number = "-2147217843" Then      '   ADO Error No.
        MsgBox "Login failed for User < " & Trim(txtLoginName.Text) & " > ", vbExclamation
        If txtLoginName.Enabled Then txtLoginName.SetFocus
        Exit Sub
    End If
    Call ErrorMessage
End Sub

Private Sub cmdDisconnect_Click()
    On Error GoTo ErrorHandler
    
    If MsgBox("Do you want Disconnect from Server ?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    Set AdoCn = Nothing
    
    fraConnection.Enabled = True
    fraAuthentication.Enabled = True
    fraDatabase.Enabled = False
    fraPathDetails.Enabled = False
'    txtServer.Text = vbNullString
    cboDBName.ListIndex = -1
    txtPath.Text = vbNullString
    txtFileName.Text = vbNullString
    cmdBackup.Enabled = False
    cmdDisconnect.Enabled = False
    cmdConnect.Enabled = True
    txtServer.SetFocus
    
    Exit Sub
ErrorHandler:
    Call ErrorMessage
End Sub

Private Sub cmdGetPath_Click()
    On Error GoTo ErrorHandler
    
    Dim iNull As Integer, lpIDList As Long, lResult As Long
    Dim sPath As String, udtBI As BrowseInfo
    
    With udtBI
        'Set the owner window
        .hWndOwner = Me.hwnd
        'lstrcat appends the two strings and returns the memory address
        .lpszTitle = lstrcat("Select Destination Folder...", "")
        'Return only if the user selected a directory
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With
    
    'Show the 'Browse for folder' dialog
    lpIDList = SHBrowseForFolder(udtBI)
    
    If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        'Get the path from the IDList
        SHGetPathFromIDList lpIDList, sPath
        'free the block of memory
        CoTaskMemFree lpIDList
        iNull = InStr(sPath, vbNullChar)
        If iNull Then
            sPath = Left$(sPath, iNull - 1)
        End If
    End If
    
    '   If user selects cancel then simple exit.
    If Len(Trim(sPath)) = 0 Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    txtPath.Text = sPath
    txtFileName.SetFocus
    
    Exit Sub
ErrorHandler:
    Call ErrorMessage
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandler
    
    Set AdoCn = Nothing
    
    txtServer.Text = vbNullString
    txtServer.Text = "(local)"
    blnConnect = False
    
    optSQLAuth.value = True
    Call SQLAuthOptionsOn
    cmdDisconnect.Enabled = False
    fraDatabase.Enabled = False
    fraPathDetails.Enabled = False
    cmdBackup.Enabled = False
    
    lblAuthInfo.Caption = "Designed && Developed By " & vbCrLf & _
                        "Amtech Technologies Ltd" & vbCrLf & "E-mail: info@amtechafrica.com"
    
    Exit Sub
ErrorHandler:
    Call ErrorMessage
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set AdoCn = Nothing
End Sub

Private Sub optSQLAuth_Click()
    Call SQLAuthOptionsOn
End Sub

Private Sub optWinNTAuth_Click()
    Call WinNTAuthOptionsOn
End Sub

Private Sub txtFileName_GotFocus()
    With txtFileName
        If Len(Trim(.Text)) > 0 Then
            .SelStart = 0
            .SelLength = Len(Trim(.Text))
        End If
    End With
End Sub

Private Sub txtFileName_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    KeyAscii = FileNameValidation(KeyAscii)
End Sub

Private Sub txtLoginName_GotFocus()
    With txtLoginName
        If Len(Trim(.Text)) > 0 Then
            .SelStart = 0
            .SelLength = Len(Trim(.Text))
        End If
    End With
End Sub

Private Sub txtPassword_GotFocus()
    With txtpassword
        If Len(Trim(.Text)) > 0 Then
            .SelStart = 0
            .SelLength = Len(Trim(.Text))
        End If
    End With
End Sub

Private Sub txtServer_GotFocus()
    With txtServer
        If Len(Trim(.Text)) > 0 Then
            .SelStart = 0
            .SelLength = Len(Trim(.Text))
        End If
    End With
End Sub

Private Sub txtServer_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
