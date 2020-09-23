Attribute VB_Name = "mod_Star"
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal DWROP As Long) As Long

Public Const WHITENESS = &HFF0062
Public Const BLACKNESS = &H42
Public Const SRCAND = &H8800C6
Public Const SRCCOPY = &HCC0020
Public Const SRCINVERT = &H660046
Public Const SRCERASE = &H440328
Public Const SRCPAINT = &HEE0086

Public Type Point
    X As Integer
    Y As Integer
End Type

Public Type Star
    X As Integer
    Y As Integer
    vx As Integer
    vy As Integer
End Type

Public bHole As Star
Public hbH As Boolean

Public sArr(1 To 100) As Star
Public sNum As Integer
