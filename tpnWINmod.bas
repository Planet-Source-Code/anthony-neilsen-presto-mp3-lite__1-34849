Attribute VB_Name = "tpnWINmod"
Option Explicit

Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Declare Function ReleaseCapture Lib "user32" () As Long


Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Global Const HWND_NOTOPMOST = -2
Global Const SWP_NOACTIVATE = &H10
Global Const SWP_SHOWWINDOW = &H40
Global Const HWND_BOTTOM = 1
Global Const HWND_TOP = 0
Global Const HWND_TOPMOST = -1
Global Const SWP_NOOWNERZORDER = &H200
Global Const SWP_NOSIZE = &H1
Global Const SWP_NOZORDER = &H4
Global Const SWP_NOMOVE = &H2


Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type


Public Function GetBitmapRegion(cPicture As StdPicture, cTransparent As Long)
Dim hRgn As Long, tRgn As Long
Dim x As Integer, y As Integer, X0 As Integer
Dim hDC As Long, BM As BITMAP
hDC = CreateCompatibleDC(0)
If hDC Then
   SelectObject hDC, cPicture
   GetObject cPicture, Len(BM), BM
   hRgn = CreateRectRgn(0, 0, BM.bmWidth, BM.bmHeight)
        
   For y = 0 To BM.bmHeight
       For x = 0 To BM.bmWidth
           While x <= BM.bmWidth And GetPixel(hDC, x, y) <> cTransparent
               x = x + 1
           Wend
           X0 = x
           While x <= BM.bmWidth And GetPixel(hDC, x, y) = cTransparent
               x = x + 1
           Wend
           If X0 < x Then
               tRgn = CreateRectRgn(X0, y, x, y + 1)
               CombineRgn hRgn, hRgn, tRgn, 4
               DeleteObject tRgn
           End If
       Next x
   Next y
   GetBitmapRegion = hRgn
   DeleteObject SelectObject(hDC, cPicture)
End If
DeleteDC hDC
End Function


