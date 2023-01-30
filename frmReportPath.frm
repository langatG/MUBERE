VERSION 5.00
Begin VB.Form frmReportPath 
   Caption         =   "Change Report Path"
   ClientHeight    =   945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5700
   Icon            =   "frmReportPath.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   945
   ScaleWidth      =   5700
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPath 
      Height          =   315
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   270
      Width           =   3135
   End
   Begin VB.CommandButton cmdGetPath 
      Caption         =   "Change"
      Height          =   375
      Left            =   4635
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Current Path"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   300
      Width           =   975
   End
End
Attribute VB_Name = "frmReportPath"
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
    MsgBox "Error No:   " & Err.number & vbCrLf & _
            "Message:  " & Err.Description, vbCritical
End Sub

'   The following function is used to restrict the FileName contains Invalid Characters.
Private Function FileNameValidation(ByVal KeyAscii As Integer) As Long
    Dim strInValid As String
    strInValid = "\/:*?<>|."""
    If InStr(strInValid, Chr(KeyAscii)) > 0 Then Exit Function
    FileNameValidation = KeyAscii
End Function

Private Sub cmdGetPath_Click()
    On Error GoTo ErrorHandler
    Dim iNull As Integer, lpIDList As Long, lResult As Long
    Dim sPath As String, udtBI As BrowseInfo
    
    If cmdGetPath.Caption = "Change" Then
        cmdGetPath.Caption = "Save"
        txtPath.Locked = False
    Else
        oSaccoMaster.ExecuteThis ("if exists (select * from reportpath) update reportpath set reportpath='" & txtPath & "\" & "' insert into reportpath (reportpath) values('" & txtPath & "\" & "')")
        cmdGetPath.Caption = "Change"
        MsgBox "Record Saved successfully"
        txtPath.Locked = True
        Exit Sub
    End If
    
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
    'txtFileName.SetFocus
    
    Exit Sub
ErrorHandler:
    Call ErrorMessage

End Sub

Private Sub Form_Load()
    On Error GoTo SysError:
        Set rst = oSaccoMaster.GetRecordSet("Select reportpath from reportPath")
        If Not rst.EOF Then
            txtPath = rst(0)
        Else
            txtPath = "DEFAULT"
        End If
    Exit Sub
SysError:
    
End Sub
