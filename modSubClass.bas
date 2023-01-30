Attribute VB_Name = "modSubClass"
Option Explicit

'===============================================================
'ListView LabelEdit
'© 2004 by Michiel Meulendijk
'
'This module handles the subclassing and belongs to the
'LabelEdit class module.
'===============================================================

Private Declare Function GetWindowLong Lib "user32" _
                Alias "GetWindowLongA" (ByVal hwnd As Long, _
                ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong& Lib "user32" _
                Alias "SetWindowLongA" (ByVal hwnd As Long, _
                ByVal nIndex As Long, ByVal dwNewLong As Long)
Private Declare Function CallWindowProc Lib "user32" Alias _
                "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
                ByVal hwnd As Long, ByVal msg As Long, _
                ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const GWL_WNDPROC = (-4)
Private Const WM_VSCROLL = &H115
Private Const WM_HSCROLL = &H114

Dim WndProcOld As Long
Dim colWnd As Collection
Public colClass As Collection

'SubClass Code
'Public Function WindProc(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'If wMsg = WM_VSCROLL Or wMsg = WM_HSCROLL Then colClass.Item("H" & hwnd).SetText
'WindProc = CallWindowProc(WndProcOld&, hwnd&, wMsg&, wParam&, lParam&)
'End Function

Public Sub InitSubClass()
Set colClass = New Collection
End Sub

Public Sub CloseSubClass()
Set colClass = Nothing
End Sub

'Public Sub SubClassWnd(hwnd As Long, Class As Object)
'colClass.Add Class, "H" & hwnd
'WndProcOld& = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WindProc)
'End Sub

Public Sub UnSubClassWnd(hwnd As Long)
SetWindowLong hwnd, GWL_WNDPROC, WndProcOld&
WndProcOld& = 0
End Sub



