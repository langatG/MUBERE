VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLsimportpayments1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public rs As ADODB.Recordset

Public Function ImportFile(ByVal frm As Form)
    On Error GoTo h
    Dim fld As ADODB.Field
    Dim strRow As String
    Dim strField As String
    Dim intPos As Integer
    Dim space
    Dim DelimiterConstant   As String
    DelimiterConstant = Chr$(9)
    DelimiterConstant = ","
    Set rs = New ADODB.Recordset
    With rs
        .Fields.Append "txtsno1", adChar, 50, adFldUpdatable
        .Fields.Append "txtregdate", adDate, 10, adFldUpdatable
        .Fields.Append "txtidno", adChar, 50, adFldUpdatable
        .Fields.Append "txtnames", adVarChar, 100, adFldUpdatable
        .Fields.Append "txtaccno", adVarChar, 20, adFldUpdatable
        .Fields.Append "txtbcode", adChar, 50, adFldUpdatable
        .Fields.Append "txtbbranch", adChar, 50, adFldUpdatable
        .Fields.Append "txttype", adChar, 30, adFldUpdatable
        .Fields.Append "txtvillage", adVarChar, 50, adFldUpdatable
        .Fields.Append "txtlocations", adVarChar, 50, adFldUpdatable
        .Fields.Append "txtdivision", adVarChar, 50, adFldUpdatable
        .Fields.Append "txtdistrict", adChar, 50, adFldUpdatable
        .Fields.Append "txttrader", adBoolean, 1, adFldUpdatable
        .Fields.Append "txtactive", adBoolean, 1, adFldUpdatable
        .Fields.Append "txtbranch", adChar, 20, adFldUpdatable
        .Fields.Append "txtphone", adVarChar, 15, adFldUpdatable
        .Fields.Append "txtaddress", adVarChar, 20, adFldUpdatable
        .Fields.Append "txttown", adChar, 50, adFldUpdatable
        .Fields.Append "txtemail", adChar, 50, adFldUpdatable
        .Fields.Append "txttranscode", adChar, 30, adFldUpdatable
        .Fields.Append "txtauditid1", adVarChar, 50, adFldUpdatable
        .Fields.Append "txtauditdatetime1", adVarChar, 50, adFldUpdatable
        .Fields.Append "txtscode", adVarChar, 50, adFldUpdatable
        .Fields.Append "txtloan", adInteger, 4, adFldUpdatable
        .Fields.Append "txtcompare", adVarChar, 10, adFldUpdatable
        .Fields.Append "txtisfrare", adVarChar, 10, adFldUpdatable
        .Fields.Append "txtfrate", adVarChar, 10, adFldUpdatable
        .Fields.Append "txtrate", adVarChar, 10, adFldUpdatable
        .Fields.Append "txthast", adInteger, 4, adFldUpdatable
        .Fields.Append "txtbr", adChar, 50, adFldUpdatable
        
        
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open
    End With
    strFileName = frmimports.txtImportedFile
    Open strFileName For Input As #1
    Do Until EOF(1)
        Line Input #1, strRow
        If Left(strRow, 5) = "SIGNA" Then
            Exit Do
        End If
        With rs
            .AddNew
            For Each fld In .Fields
                If InStr(1, strRow, DelimiterConstant, vbTextCompare) <> 0 Then
                    'Move position to tab delimiter.
                    intPos = InStr(1, strRow, DelimiterConstant, vbTextCompare)
                    'Assign field text to strField variable.
                    strField = Left(strRow, intPos - 1)
                Else
                    'If a tab delimiter isn't found, field text is the
                    'last field in the row.
                    strField = strRow
                End If            ' Strip off quotation marks.
                If Left(strField, 1) = Chr(0) Then
                    strField = Left(strField, Len(strField) - 1)
                    strField = Right(strField, Len(strField) - 1)
                End If
                fld.value = Trim(strField)
                'Strip off field value text from text row.
                strRow = Right(strRow, Len(strRow) - intPos)
                intPos = 0
            Next
            If strRow <> "" Then .Update
            .MoveFirst
        End With
    Loop
closeFile:
    Close
    Exit Function
h:
    MsgBox err.description
End Function
Private Sub Class_GetDataMember(DataMember As String, Data As Object)
Set Data = rs
End Sub
Private Sub Class_Initialize()
    ImportFile frmimports
End Sub

