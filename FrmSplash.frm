VERSION 5.00
Begin VB.Form FrmSplash 
   BorderStyle     =   0  'None
   ClientHeight    =   5400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2520
      Top             =   2520
   End
End
Attribute VB_Name = "FrmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Call the GDI class
Private PNG As LayeredWindow

'Close the form
Private Sub Form_DblClick()
End
End Sub

Private Sub Form_Load()
'Set the GDI class
Set PNG = New LayeredWindow

'Make the form transparent
PNG.MakeTrans App.Path & "\Globe.png", Me
End Sub

'Move the form
'TAKES A LOT OF CPU USAGE
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then FormMove Me
End Sub

'Unload the PNG form
'IF U DON'T USE THIS, VB WILL CRASH
'WORKS BETTER IF COMPILED
Private Sub Form_Unload(Cancel As Integer)
PNG.UnloadPNGForm
End Sub

'Make the form topmost, even on top of the Start menu
'Refreshed ever msec
Private Sub Timer1_Timer()
SetOnTop Me.hwnd, True
End Sub
