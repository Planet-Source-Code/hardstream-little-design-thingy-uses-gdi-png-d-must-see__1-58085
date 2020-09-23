Attribute VB_Name = "MOT"
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOP = -2

Sub SetOnTop(hwnd, OnTop As Boolean)
If OnTop = True Then
    SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
        ElseIf OnTop = False Then
    SetWindowPos hwnd, HWND_NOTOP, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End If
End Sub
