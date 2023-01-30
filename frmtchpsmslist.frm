VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmtchpsmslist 
   Caption         =   "TCHP SMS LIST"
   ClientHeight    =   9195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9195
   ScaleWidth      =   14940
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Select All"
      Height          =   495
      Left            =   13680
      TabIndex        =   3
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      Height          =   375
      Left            =   13920
      TabIndex        =   2
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton cmdsendsms 
      Caption         =   "Send SMS"
      Height          =   375
      Left            =   13920
      TabIndex        =   1
      Top             =   1920
      Width           =   975
   End
   Begin MSComctlLib.ListView LVTCHPSMS 
      Height          =   9015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13335
      _ExtentX        =   23521
      _ExtentY        =   15901
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Sno"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Status"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Phone"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Content"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "MsgType"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "AAR TK#"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmtchpsmslist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim I As Integer
Private Sub Check1_Click()

With LVTCHPSMS
    If Check1.value = vbChecked Then
        If LVTCHPSMS.ListItems.Count >= 1 Then
        For I = 1 To .ListItems.Count
        LVTCHPSMS.ListItems(I).Checked = True
        Next I
        End If
    Else
    If LVTCHPSMS.ListItems.Count >= 1 Then
        For I = 1 To .ListItems.Count
        LVTCHPSMS.ListItems(I).Checked = False
        Next I
        End If
    End If
End With
End Sub

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdsendsms_Click()
Dim id As Long
Dim sno As String, Phone As String, content As String, mtype As String, status As String, nam As String
On Error GoTo errorhandler
 For I = 1 To LVTCHPSMS.ListItems.Count
        If LVTCHPSMS.ListItems.Item(I).Checked = True Then
        Set li = LVTCHPSMS.ListItems(I)
        
        id = li
        sno = LVTCHPSMS.ListItems(I).SubItems(1)
        status = LVTCHPSMS.ListItems(I).SubItems(3)
        nam = Replace(LVTCHPSMS.ListItems(I).SubItems(2), ",", "")
        Phone = LVTCHPSMS.ListItems(I).SubItems(4)
        content = LVTCHPSMS.ListItems(I).SubItems(5)
        mtype = LVTCHPSMS.ListItems(I).SubItems(6)
        
        strSQL = "INSERT INTO Messages(Telephone,Content,ProcessTime, MsgType,Source,names,sno)"
        strSQL = strSQL & "Values ('" & Phone & "','" & content & "',getDate(),'Outbox','" & user & "','" & nam & "','" & sno & "')"
        oSaccoMaster.ExecuteThis (strSQL)
        
        End If
    Next I
    MsgBox "Items successfully qued"
                        Exit Sub
errorhandler:
                        MsgBox err.description
End Sub

Private Sub Form_Load()
Dim sno As String
Dim NAMES As String
sql = "SELECT     sno, status, phone, Content, msgtype,ID,aarno  FROM         tchp_trxsreport  WHERE     (status <> 'Valid') AND (LEN(phone) > 9)"
Set rs = oSaccoMaster.GetRecordset(sql)
With LVTCHPSMS
        
        
 
    
        While Not rs.EOF
        sno = rs.Fields(0)
         Set li = .ListItems.Add(, , Trim(rs.Fields("id")))
           If Not IsNull(rs.Fields("SNo")) Then
        
           li.ListSubItems.Add , , Trim(rs.Fields("sno"))
           'get name here
           sql = ""
           sql = "select names from d_suppliers where sno='" & sno & "'"
           Set Rst = oSaccoMaster.GetRecordset(sql)
           If Not Rst.EOF Then
           li.ListSubItems.Add , , Trim(Rst.Fields(0))
           End If
            End If
            If Not IsNull(rs.Fields("status")) Then
            li.ListSubItems.Add , , Trim(rs.Fields("status"))
            End If
            If Not IsNull(rs.Fields("phone")) Then
            li.ListSubItems.Add , , Trim(rs.Fields("phone"))
           
            End If
            If Not IsNull(rs.Fields("Content")) Then
            li.ListSubItems.Add , , Trim(rs.Fields("Content"))
           
            End If
            If Not IsNull(rs.Fields("msgtype")) Then
             li.ListSubItems.Add , , Trim(rs.Fields("msgtype"))
          
            End If
             If Not IsNull(rs.Fields("aarno")) Then
             li.ListSubItems.Add , , Trim(rs.Fields("aarno"))
          
            End If

            
                    rs.MoveNext
        
        Wend
        
    End With

End Sub
