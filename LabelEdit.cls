VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LabelEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'===============================================================
'ListView LabelEdit
'© 2004 by Michiel Meulendijk
'
'This code enables label editing for SubItems in ListViews.
'By default, when using the LabelEdit property, only the first
'ListItem of a ListView can be edited. With this code all
'ListSubItems can be edited as well.
'
'This code is contained within a class, so multiple instances
'can run at the same time (e.g. more ListViews on one form can
'all support label editing).
'
'This file is provided "as is" with no expressed or implied
'warranty. The author accepts no liability for any damage caused
'to your system because of using this code.
'===============================================================
Private Declare Function GetScrollInfo Lib "user32.dll" ( _
                ByVal hwnd As Long, ByVal n As Long, _
                lpScrollInfo As SCROLLINFO) As Long
Private Declare Function SetWindowLong& Lib "user32" _
                Alias "SetWindowLongA" (ByVal hwnd As Long, _
                ByVal nIndex As Long, ByVal dwNewLong As Long)

Private Type SCROLLINFO
    cbSize As Long
    fMask As Long
    nMin As Long
    nMax As Long
    nPage As Long
    nPos As Long
    nTrackPos As Long
End Type
Private Const GWL_WNDPROC = (-4)
Private Const WM_VSCROLL = &H115
Private Const WM_HSCROLL = &H114

Private Const SB_HORZ = 0
Private Const SB_VERT = 1
Private Const SIF_POS = &H4
Dim WithEvents txtEdit As TextBox
Attribute txtEdit.VB_VarHelpID = -1
Dim WithEvents ltvListView As ListView
Attribute ltvListView.VB_VarHelpID = -1
Dim objItem As Object
Dim WndProcOld As Long
Dim colWnd As Collection
Dim colClass As Collection

Public Sub Init(ByRef ctlForm As Form, ByRef ctlListView As ListView)
'Initiates object. Adds textbox control.
On Error Resume Next
Set ltvListView = ctlListView
Set txtEdit = ctlForm.Controls.Add("VB.TextBox", "txtLabelEdit_" & ctlListView.name)
Set txtEdit.Container = ctlListView.Container
Set txtEdit.Font = ctlListView.Font
txtEdit.Appearance = 0
txtEdit.ForeColor = vbHighlight
'Subclass listview
'SubClassWnd ltvListView.hwnd, Me
End Sub

Private Function GetHorizontalScroll() As Long
'Returns the position of the horizontal scroll bar
Dim scrInfo As SCROLLINFO
scrInfo.cbSize = LenB(scrInfo)
scrInfo.fMask = SIF_POS
GetScrollInfo ltvListView.hwnd, SB_HORZ, scrInfo
GetHorizontalScroll = scrInfo.nPos
End Function

Private Sub EditText(ByVal X As Integer, ByVal Y As Integer)
'Handles label editing
    On Error GoTo EndSub
    Dim I As Integer, objCol As ColumnHeader, lngScroll As Long
    lngScroll = GetHorizontalScroll * Screen.TwipsPerPixelX
    X = X + lngScroll
    For I = 1 To ltvListView.ColumnHeaders.Count
        If X < ltvListView.ColumnHeaders.Item(1).Width Or ltvListView.ColumnHeaders.Count = 1 Then
            Set objCol = ltvListView.ColumnHeaders.Item(1)
            Set objItem = ltvListView.SelectedItem
            Exit For
        ElseIf X < ltvListView.ColumnHeaders.Item(I).Left Then
            Set objCol = ltvListView.ColumnHeaders.Item(I - 1)
            Set objItem = ltvListView.SelectedItem.ListSubItems.Item(I - 2)
            Exit For
        ElseIf I = ltvListView.ColumnHeaders.Count Then
            Set objCol = ltvListView.ColumnHeaders(I)
            Set objItem = ltvListView.SelectedItem.ListSubItems.Item(I - 1)
            Exit For
        End If
    Next I
    txtEdit.BorderStyle = 0
    txtEdit.Left = ltvListView.Left + objCol.Left - lngScroll + 40
    txtEdit.Top = ltvListView.Top + ltvListView.SelectedItem.Top + 60
    txtEdit.Width = objCol.Width
    txtEdit.Height = ltvListView.SelectedItem.Height
    txtEdit.BorderStyle = 0
    txtEdit.Text = objItem.Text
    txtEdit.SelStart = 0
    txtEdit.SelLength = Len(txtEdit)
    txtEdit.Visible = True
    txtEdit.SetFocus
EndSub:
End Sub

Public Sub SetText()
On Error Resume Next
objItem.Text = txtEdit.Text
txtEdit.Visible = False
End Sub

Private Sub Class_Terminate()
On Error Resume Next
UnSubClassWnd ltvListView.hwnd
Set txtEdit = Nothing
Set ltvListView = Nothing
End Sub

Private Sub ltvListView_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    EditText X, Y
End Sub

Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SetText
End Sub

Private Sub txtEdit_LostFocus()
SetText
End Sub
Public Sub UnSubClassWnd(hwnd As Long)
SetWindowLong hwnd, GWL_WNDPROC, WndProcOld&
WndProcOld& = 0
End Sub
Public Sub InitSubClass()
Set colClass = New Collection
End Sub



