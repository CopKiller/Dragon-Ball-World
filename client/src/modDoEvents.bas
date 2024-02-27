Attribute VB_Name = "modDoEvents"
Option Explicit

Private Type Msg
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    Time As Long
    pt As POINTAPI
End Type

Private M As Msg

Private Const WM_SYSCOMMAND As Long = &H112
Private Const WM_CLOSE As Long = &H10
Private Const WM_DESTROY As Long = &H2
Private Const PM_NOREMOVE As Long = &H0

Private Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As Msg, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long

Public Sub GoPeekMessage()
    If PeekMessage(M, 0, 0, 0, PM_NOREMOVE) Then DoEvents
End Sub
