VERSION 5.00
Begin VB.Form frmODBCLogon1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ODBC Logon"
   ClientHeight    =   1545
   ClientLeft      =   2850
   ClientTop       =   1755
   ClientWidth     =   3960
   ControlBox      =   0   'False
   Icon            =   "frmODBCLogon1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   3960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   450
      Left            =   2475
      TabIndex        =   3
      Top             =   975
      Width           =   1440
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   450
      Left            =   930
      TabIndex        =   2
      Top             =   975
      Width           =   1440
   End
   Begin VB.Frame fraStep3 
      Caption         =   "Connection Values"
      Height          =   750
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3795
      Begin VB.ComboBox cboDSNList 
         Height          =   315
         ItemData        =   "frmODBCLogon1.frx":030A
         Left            =   1110
         List            =   "frmODBCLogon1.frx":0311
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2625
      End
      Begin VB.Label lblStep3 
         AutoSize        =   -1  'True
         Caption         =   "&DataBase"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   0
         Top             =   270
         Width           =   705
      End
   End
End
Attribute VB_Name = "frmODBCLogon1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SQLDataSources Lib "ODBC32.DLL" (ByVal henv&, ByVal fDirection%, ByVal szDSN$, ByVal cbDSNMax%, pcbDSN%, ByVal szDescription$, ByVal cbDescriptionMax%, pcbDescription%) As Integer
Private Declare Function SQLAllocEnv% Lib "ODBC32.DLL" (env&)
Const SQL_SUCCESS As Long = 0
Const SQL_FETCH_NEXT As Long = 1


Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    SelectedDsn = cboDSNList.Text
    MainForm.Show vbModal
    Unload Me
End Sub

Private Sub Form_Load()
    GetDSNsAndDrivers
End Sub

Private Sub cboDSNList_Click()
    On Error Resume Next
    
End Sub

Sub GetDSNsAndDrivers()
    Dim i As Integer
    Dim sDSNItem As String * 1024
    Dim sDRVItem As String * 1024
    Dim sDSN As String
    Dim sDRV As String
    Dim iDSNLen As Integer
    Dim iDRVLen As Integer
    Dim lHenv As Long         'handle to the environment

    On Error Resume Next
    cboDSNList.AddItem "(None)"

    'get the DSNs
    If SQLAllocEnv(lHenv) <> -1 Then
        Do Until i <> SQL_SUCCESS
            sDSNItem = Space$(1024)
            sDRVItem = Space$(1024)
            i = SQLDataSources(lHenv, SQL_FETCH_NEXT, sDSNItem, 1024, iDSNLen, sDRVItem, 1024, iDRVLen)
            sDSN = Left$(sDSNItem, iDSNLen)
            sDRV = Left$(sDRVItem, iDRVLen)
                
            If sDSN <> Space(iDSNLen) Then
                cboDSNList.AddItem sDSN
                'cboDrivers.AddItem sDRV
            End If
        Loop
    End If
    'remove the dupes
    If cboDSNList.ListCount > 0 Then
'        With cboDrivers
'            If .ListCount > 1 Then
'                i = 0
'                While i < .ListCount
'                    If .List(i) = .List(i + 1) Then
'                        .RemoveItem (i)
'                    Else
'                        i = i + 1
'                    End If
'                Wend
'            End If
'        End With
    End If
    cboDSNList.ListIndex = 0
End Sub
