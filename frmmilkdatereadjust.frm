VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmmilkdatereadjust 
   Caption         =   "Change Date milk Receive"
   ClientHeight    =   7050
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   ScaleHeight     =   7050
   ScaleWidth      =   6720
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      Height          =   495
      Left            =   2280
      TabIndex        =   7
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdupdate 
      Caption         =   "Update"
      Height          =   495
      Left            =   480
      TabIndex        =   6
      Top             =   2040
      Width           =   1335
   End
   Begin VB.ComboBox cbolocation 
      Height          =   315
      ItemData        =   "frmmilkdatereadjust.frx":0000
      Left            =   1560
      List            =   "frmmilkdatereadjust.frx":0002
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   840
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      MouseIcon       =   "frmmilkdatereadjust.frx":0004
      CalendarBackColor=   8454016
      Format          =   130416641
      CurrentDate     =   40095
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   1440
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      MouseIcon       =   "frmmilkdatereadjust.frx":08DE
      CalendarBackColor=   8454016
      Format          =   130416641
      CurrentDate     =   40095
   End
   Begin MSComctlLib.ListView lvWBranch2 
      Height          =   4215
      Left            =   120
      TabIndex        =   8
      Top             =   2640
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   7435
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label3 
      Caption         =   "Date To :"
      BeginProperty Font 
         Name            =   "News701 BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Date From:"
      BeginProperty Font 
         Name            =   "News701 BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Branch:"
      BeginProperty Font 
         Name            =   "News701 BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmmilkdatereadjust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdupdate_Click()
On Error GoTo ErrorHandler
 If cbolocation = "" Then
  MsgBox "Please Select The Branch."
 Exit Sub
 End If
    Set rst = New ADODB.Recordset
    sql = "set dateformat dmy select * from  d_Milkintake where LOCATION='" & cbolocation & "' and TransDate='" & DTPicker1 & "'"
    Set rst = oSaccoMaster.GetRecordset(sql)
 If Not rst.EOF Then
     Set cn = New ADODB.Connection
     sql = "set dateformat dmy Update d_Milkintake set TransDate='" & DTPicker2 & "' where LOCATION='" & cbolocation & "' and TransDate='" & DTPicker1 & "'"
     oSaccoMaster.ExecuteThis (sql)
     
     sql = ""
     sql = "insert into d_milkintakechange(Branch, [Date From], [Date To], userid) values('" & cbolocation & "','" & DTPicker1 & "','" & DTPicker2 & "','" & User & "')"
     oSaccoMaster.ExecuteThis (sql)
 Else
  MsgBox "No milk intake for this Branch " & cbolocation & " on '" & DTPicker1 & "'"
  Exit Sub
 End If
 MsgBox "Record updated succesfully."
 loadpro
Exit Sub
ErrorHandler:
MsgBox err.description
End Sub
Public Sub loadpro()
    
    With lvWBranch2
    
        .ListItems.Clear
        
        .ColumnHeaders.Clear
    
    End With

    Set rs2 = CreateObject("adodb.recordset")
    
    sql = "set dateformat dmy Select Branch, [Date From], [Date To], userid from d_milkintakechange where [Date From]='" & DTPicker1 & "'"
    
    Set rs2 = CreateObject("adodb.recordset")
    
    Set clsClass = New cdbase
    
    Provider = clsClass.OpenCon
    
    Set cn = CreateObject("adodb.connection")
    
   cn.Open Provider, "atm", "atm"
    
    rs2.Open sql, cn
    
    With lvWBranch2
        
        .ColumnHeaders.Add , , "Branch"
        .ColumnHeaders.Add , , "Date From"
        .ColumnHeaders.Add , , "Date To"
        .ColumnHeaders.Add , , "User"
        '.ColumnHeaders.Add , , "R.Price"
        '.ColumnHeaders.Add , , "Branch"
        While Not rs2.EOF
        Set li = .ListItems.Add(, , Trim(rs2.Fields("Branch")))
            li.ListSubItems.Add , , Trim(rs2.Fields("Date From"))
            li.ListSubItems.Add , , Trim(rs2.Fields("Date To"))
            li.ListSubItems.Add , , Trim(rs2.Fields("userid"))
            'li.ListSubItems.Add , , Trim(rs2.Fields("Rprice"))
            'li.ListSubItems.Add , , Trim(rs2.Fields("Branch"))
            rs2.MoveNext
        Wend
        
    End With
    
    rs2.Close
    
    Set rs2 = Nothing
    
lvWBranch2.View = lvwReport

End Sub

Private Sub DTPicker1_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
loadpro
End Sub

Private Sub Form_Load()
DTPicker1 = Format(Get_Server_Date, "dd/mm/yyyy")
DTPicker2 = Format(Get_Server_Date, "dd/mm/yyyy")
'DTPMilkDate.MaxDate = Format(Get_Server_Date, "dd/mm/yyyy")
DTPcomplaintperiod = DTPicker1

    Set rst = New Recordset
    Dim cn As Connection
    Set cn = New ADODB.Connection
    Provider = "MAZIWA"
    'Provider = cnn
    cn.Open Provider, "atm", "atm"
    Set rst = New Recordset
    
    sql = "Select Bname from   d_Branch order by Bname asc"
    rst.Open sql, cn, adOpenKeyset, adLockOptimistic
    While Not rst.EOF
    cbolocation.AddItem rst.Fields(0)
    rst.MoveNext
    Wend
End Sub

