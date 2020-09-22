Attribute VB_Name = "SpriteBas"
Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

Declare Function BitBlt Lib "gdi32" ( _
        ByVal hDestDC As Long, _
        ByVal X As Long, _
        ByVal Y As Long, _
        ByVal nWidth As Long, _
        ByVal nHeight As Long, _
        ByVal hSrcDC As Long, _
        ByVal xSrc As Long, _
        ByVal ySrc As Long, _
        ByVal dwRop As Long) As Long

Public Const SRCCOPY = &HCC0020
Public Const SRCAND = &H8800C6
Public Const SRCINVERT = &H660046
Public Const SRCPAINT = &HEE0086
Public Const SRCERASE = &H4400328
Public Const WHITENESS = &HFF0062
Public Const BLACKNESS = &H42

Function Mask(PicSrc As PictureBox, picDEST As PictureBox, bColor As OLE_COLOR)
    Dim looper As Long
    Dim looper2 As Long
    Dim bColor2 As OLE_COLOR
    picDEST.Cls
    For looper = 0 To PicSrc.ScaleHeight
    picDEST.Refresh
        For looper2 = 0 To PicSrc.ScaleWidth
            If PicSrc.Point(looper2, looper) = bColor Then
                bColor2 = RGB(255, 255, 255)
            Else
                bColor2 = RGB(0, 0, 0)
            End If
            SetPixel picDEST.hdc, looper2, looper, bColor2
        Next looper2
    Next looper
    picDEST.Refresh
End Function
Function Sprite(PicSrc As PictureBox, picDEST As PictureBox, bColor As OLE_COLOR)
    Dim looper As Long
    Dim looper2 As Long
    Dim bColor2 As OLE_COLOR
    picDEST.Cls
    For looper = 0 To PicSrc.ScaleHeight
    picDEST.Refresh
        For looper2 = 0 To PicSrc.ScaleWidth
            If PicSrc.Point(looper2, looper) = bColor Then
                bColor2 = RGB(0, 0, 0)
            Else
                bColor2 = GetPixel(PicSrc.hdc, looper2, looper)
            End If
            SetPixel picDEST.hdc, looper2, looper, bColor2
        Next looper2
    Next looper
    picDEST.Refresh
End Function
