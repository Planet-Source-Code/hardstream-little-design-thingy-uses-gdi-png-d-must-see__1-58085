Attribute VB_Name = "modMove"
Option Explicit

Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wparam As Integer, ByVal iparam As Long) As Long

Public Sub FormMove(Frm As Form)
ReleaseCapture
Call SendMessage(Frm.hWnd, &HA1, 2, 0&)
End Sub
