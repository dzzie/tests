Attribute VB_Name = "mGDI"
Option Explicit

' ===== Screen capture / GDI =====
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, _
    ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, _
    ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Const SM_CXSCREEN = 0
Public Const SM_CYSCREEN = 1
Public Const SRCCOPY = &HCC0020

' Pens / brushes
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal lpPoint As Long) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function Polygon Lib "gdi32" (ByVal hDC As Long, lpPoint As Any, ByVal nCount As Long) As Long
Public Declare Function PatBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Public Const PATCOPY = &HF00021
Public Const BLACKNESS = &H42

' Regions (the key new piece)
Public Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As Any, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hrgnDest As Long, ByVal hrgnSrc1 As Long, ByVal hrgnSrc2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function SelectClipRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long) As Long
Public Declare Function FillRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Public Const ALTERNATE = 1
Public Const RGN_AND = 1
Public Const RGN_OR = 2
Public Const RGN_DIFF = 4

Public Type POINTAPI
    x As Long
    y As Long
End Type

Public Const PS_SOLID = 0

' ===== Topmost window =====
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, _
    ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_SHOWWINDOW = &H40

' ===== Sound =====
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" _
    (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_FILENAME = &H20000

' ===== GDI+ flat API (for PNG with alpha) =====
Public Declare Function GdiplusStartup Lib "gdiplus" (token As Long, inputbuf As Any, Optional ByVal outputbuf As Long = 0) As Long
Public Declare Function GdiplusShutdown Lib "gdiplus" (ByVal token As Long) As Long
Public Declare Function GdipCreateBitmapFromFile Lib "gdiplus" (ByVal filename As Long, bitmap As Long) As Long
Public Declare Function GdipDisposeImage Lib "gdiplus" (ByVal image As Long) As Long
Public Declare Function GdipGetImageWidth Lib "gdiplus" (ByVal image As Long, width As Long) As Long
Public Declare Function GdipGetImageHeight Lib "gdiplus" (ByVal image As Long, height As Long) As Long
Public Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hDC As Long, graphics As Long) As Long
Public Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal graphics As Long) As Long
Public Declare Function GdipDrawImageRectI Lib "gdiplus" (ByVal graphics As Long, ByVal image As Long, _
    ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long) As Long
Public Declare Function GdipSetInterpolationMode Lib "gdiplus" (ByVal graphics As Long, ByVal interpolationMode As Long) As Long

Public Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Public g_gdipToken As Long

Sub Main()
    
    On Error Resume Next
    Dim i As Long
    Const max = 9
    
    i = CLng(Command)
    'MsgBox i & " " & Command
    
    If i < 1 Or i > max Then i = 5
    
    Select Case i
        Case 1: Form1.Show
        Case 2: Form2.Show
        Case 3: Form3.Show
        Case 4: Form4.Show
        Case 5: Form5.Show
        Case 6: Form6.Show
        Case 7: Form7.Show
        Case 8: Form8.Show
        Case 9: Form9.Show
    End Select
    
    
End Sub



Public Sub GdipInit()
    Dim si As GdiplusStartupInput
    si.GdiplusVersion = 1
    GdiplusStartup g_gdipToken, si
End Sub

Public Sub GdipShutdown()
    If g_gdipToken <> 0 Then GdiplusShutdown g_gdipToken
    g_gdipToken = 0
End Sub

Public Function LoadPNG(ByVal path As String) As Long
    Dim h As Long
    If GdipCreateBitmapFromFile(StrPtr(path), h) = 0 Then
        LoadPNG = h
    Else
        LoadPNG = 0
    End If
End Function

Public Sub DrawPNG(ByVal hDC As Long, ByVal img As Long, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long)
    If img = 0 Then Exit Sub
    Dim g As Long
    If GdipCreateFromHDC(hDC, g) = 0 Then
        GdipSetInterpolationMode g, 7
        GdipDrawImageRectI g, img, x, y, w, h
        GdipDeleteGraphics g
    End If
End Sub

Public Function FileThere(ByVal path As String) As Boolean
    On Error Resume Next
    FileThere = (Len(Dir$(path)) > 0)
End Function
