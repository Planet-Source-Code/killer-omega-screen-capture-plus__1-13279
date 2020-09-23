Attribute VB_Name = "Module1"
Public Declare Function BmpToJpeg Lib "JPeg32.dll" (ByVal BmpFilename As String, ByVal JpegFilename As String, ByVal Quality As Integer) As Integer
Public Declare Function BringWindowToTop Lib "user32.dll" (ByVal hwnd As Long) As Long
Declare Function BitBlt Lib "gdi32.dll" (ByVal hDCDest As Long, ByVal nXDest As Long, ByVal nYDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hdcSrc As Long, ByVal nXSrc As Long, ByVal nYSrc As Long, ByVal dwRop As Long) As Long
Const BLACKNESS = &H42
Const DSTINVERT = &H550009
Const MERGECOPY = &HC000CA
Const MERGEPAINT = &HBB0226
Const NOTSRCCOPY = &H330008
Const NOTSRCERASE = &H1100A6
Const PATCOPY = &HF00021
Const PATINVERT = &H5A0049
Const SRCCOPY = &HCC0020

Public Declare Function GetDesktopWindow Lib "user32.dll" () As Long
Public Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Public Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal vKey As Long) As Integer

Const VK_SCROLL = &H91
Public HOTKEY As String
Const PAUSE_TIME = 750
Public HWNDOFSAVE As Variant

Public Function CaptureScreen(PicDest As Object)

Dim DeskhWnd As Long
Dim DeskhDC As Long

    DeskhWnd = GetDesktopWindow
    DeskhDC = GetDC(DeskhWnd)
    
    Call BitBlt(PicDest.hdc, 0, 0, Screen.Width, Screen.Height, _
                DeskhDC, 0&, 0&, SRCCOPY)
    
    Call ReleaseDC(DeskhWnd, DeskhDC)
End Function
