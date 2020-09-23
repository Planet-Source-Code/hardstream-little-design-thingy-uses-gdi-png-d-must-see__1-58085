VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LayeredWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const ULW_OPAQUE = &H4
Private Const ULW_COLORKEY = &H1
Private Const ULW_ALPHA = &H2
Private Const BI_RGB As Long = 0&
Private Const DIB_RGB_COLORS As Long = 0
Private Const AC_SRC_ALPHA As Long = &H1
Private Const AC_SRC_OVER = &H0
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_STYLE As Long = -16
Private Const GWL_EXSTYLE As Long = -20
Private Const HWND_TOPMOST As Long = -1
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOSIZE As Long = &H1

Private Type BLENDFUNCTION
    BlendOp As Byte
    BlendFlags As Byte
    SourceConstantAlpha As Byte
    AlphaFormat As Byte
End Type

Private Type Size
    cx As Long
    cy As Long
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type

Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function AlphaBlend Lib "Msimg32.dll" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal lnYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hdcSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal bf As Long) As Boolean
Private Declare Function UpdateLayeredWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, ByVal crKey As Long, ByRef pblend As BLENDFUNCTION, ByVal dwFlags As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32.dll" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByRef lplpVoid As Any, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function GetDIBits Lib "gdi32.dll" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetDIBits Lib "gdi32.dll" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long

Dim mDC As Long
Dim MainBitmap As Long
Dim blendFunc32bpp As BLENDFUNCTION
Dim Token As Long
Dim OldBitmap As Long

Sub UnloadPNGForm()
Call GdiplusShutdown(Token)

SelectObject mDC, OldBitmap
DeleteObject MainBitmap
DeleteObject OldBitmap
DeleteDC mDC
End Sub

Private Sub Class_Initialize()
Dim GpInput As GdiplusStartupInput

GpInput.GdiplusVersion = 1
If GdiplusStartup(Token, GpInput) <> 0 Then
    MsgBox "Error loading GDI+!", vbCritical, "GDI Plus Error"
End If
End Sub

Public Function MakeTrans(pngPath As String, Frm As Object) As Boolean
Dim tempBI As BITMAPINFO
Dim tempBlend As BLENDFUNCTION
Dim lngHeight As Long, lngWidth As Long
Dim curWinLong As Long
Dim img As Long
Dim graphics As Long
Dim winSize As Size
Dim srcPoint As POINTAPI
   
With tempBI.bmiHeader
    .biSize = Len(tempBI.bmiHeader)
    .biBitCount = 32
    .biHeight = Frm.ScaleHeight
    .biWidth = Frm.ScaleWidth
    .biPlanes = 1
    .biSizeImage = .biWidth * .biHeight * (.biBitCount / 8)
End With

mDC = CreateCompatibleDC(Frm.hdc)
MainBitmap = CreateDIBSection(mDC, tempBI, DIB_RGB_COLORS, ByVal 0, 0, 0)
OldBitmap = SelectObject(mDC, MainBitmap)
    
Call GdipCreateFromHDC(mDC, graphics)
Call GdipLoadImageFromFile(StrConv(pngPath, vbUnicode), img)
Call GdipGetImageHeight(img, lngHeight)
Call GdipGetImageWidth(img, lngWidth)
Call GdipDrawImageRect(graphics, img, 0, 0, lngWidth, lngHeight)

curWinLong = GetWindowLong(Frm.hwnd, GWL_EXSTYLE)
SetWindowLong Frm.hwnd, GWL_EXSTYLE, curWinLong Or WS_EX_LAYERED

SetWindowPos Frm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE

srcPoint.x = 0
srcPoint.y = 0
winSize.cx = Frm.ScaleWidth
winSize.cy = Frm.ScaleHeight
    
With blendFunc32bpp
    .AlphaFormat = AC_SRC_ALPHA
    .BlendFlags = 0
    .BlendOp = AC_SRC_OVER
    .SourceConstantAlpha = 255
End With
    
Call GdipDisposeImage(img)
Call GdipDeleteGraphics(graphics)
Call UpdateLayeredWindow(Frm.hwnd, Frm.hdc, ByVal 0&, winSize, mDC, srcPoint, 0, blendFunc32bpp, ULW_ALPHA)
End Function
