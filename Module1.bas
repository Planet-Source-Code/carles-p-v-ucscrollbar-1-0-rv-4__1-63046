Attribute VB_Name = "Module1"
Option Explicit

'-- Some helper functions for Custom-Draw style...

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type RECT
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
End Type

Private Const DFC_SCROLL    As Long = 3
Private Const DFCS_MONO     As Long = &H8000
Private Const DFCS_FLAT     As Long = &H4000
Private Const PS_SOLID      As Long = 0
Private Const COLOR_BTNTEXT As Long = 18

Private Declare Function DrawFrameControl Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Integer) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function InvertRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long
Private Declare Function OleTranslateColor Lib "olepro32" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, ColorRef As Long) As Long

'//

Public Enum ArrowFlagCts
    [afDFCS_SCROLLUP] = &H0
    [afDFCS_SCROLLDOWN] = &H1
    [afDFCS_SCROLLLEFT] = &H2
    [afDFCS_SCROLLRIGHT] = &H3
    [afDFCS_PUSHED] = &H200
End Enum

Public Sub DrawEdgeEx(ByVal hDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal clr As Long, Optional ByVal bPressed As Boolean)
    
  Dim clrHighLight As Long
  Dim clrShadow    As Long
  
    clrHighLight = ShiftColor(clr, &H40)
    clrShadow = ShiftColor(clr, -&H40)
    
    If (bPressed) Then
        Call pvDrawLine(hDC, x1, y2 - 2, x1, y1, clrShadow)
        Call pvDrawLine(hDC, x1, y1, x2, y1, clrShadow)
        Call pvDrawLine(hDC, x1, y2 - 1, x2 - 1, y2 - 1, clrHighLight)
        Call pvDrawLine(hDC, x2 - 1, y2 - 1, x2 - 1, y1, clrHighLight)
    Else
        Call pvDrawLine(hDC, x1, y2 - 2, x1, y1, clrHighLight)
        Call pvDrawLine(hDC, x1, y1, x2, y1, clrHighLight)
        Call pvDrawLine(hDC, x1, y2 - 1, x2 - 1, y2 - 1, clrShadow)
        Call pvDrawLine(hDC, x2 - 1, y2 - 1, x2 - 1, y1, clrShadow)
    End If
End Sub

Public Sub DrawRectEx(ByVal hDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal clr As Long)
    
    Call pvDrawLine(hDC, x1, y2 - 2, x1, y1, clr)
    Call pvDrawLine(hDC, x1, y1, x2, y1, clr)
    Call pvDrawLine(hDC, x1, y2 - 1, x2 - 1, y2 - 1, clr)
    Call pvDrawLine(hDC, x2 - 1, y2 - 1, x2 - 1, y1, clr)
End Sub

Public Sub FillRectEx(ByVal hDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal clr As Long)
    
  Dim uRct   As RECT
  Dim hBrush As Long
    
    Call SetRect(uRct, x1, y1, x2, y2)
    hBrush = CreateSolidBrush(clr)
    Call FillRect(hDC, uRct, hBrush)
    Call DeleteObject(hBrush)
End Sub

Public Sub DrawArrowEx(ByVal hDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal lArrowFlags As ArrowFlagCts, ByVal clrBk As Long, ByVal clrFr As Long)

  Dim uRctMem    As RECT
  
  Dim hDCMem1    As Long
  Dim hDCMem2    As Long
  Dim hBmp1      As Long
  Dim hBmp2      As Long
  Dim hBmpOld1   As Long
  Dim hBmpOld2   As Long
  
  Dim clrBkOld   As Long
  Dim clrTextOld As Long
        
    '-- Monochrome bitmap to convert the arrow to black/white mask
    hDCMem1 = CreateCompatibleDC(hDC)
    hBmp1 = CreateBitmap(x2 - x1, y2 - y1, 1, 1, ByVal 0)
    hBmpOld1 = SelectObject(hDCMem1, hBmp1)
    
    '-- Normal bitmap to draw the arrow into
    hDCMem2 = CreateCompatibleDC(hDC)
    hBmp2 = CreateCompatibleBitmap(hDC, x2 - x1, y2 - y1)
    hBmpOld2 = SelectObject(hDCMem2, hBmp2)
    
    '-- Draw frame normaly
    Call SetRect(uRctMem, 0, 0, x2 - x1, y2 - y1)
    Call InvertRect(hDCMem1, uRctMem)
    Call DrawFrameControl(hDCMem2, uRctMem, DFC_SCROLL, DFCS_FLAT Or DFCS_MONO Or lArrowFlags)
    
    '-- Magic!
    Call SetBkColor(hDCMem2, GetSysColor(COLOR_BTNTEXT))
    Call BitBlt(hDCMem1, 0, 0, x2 - x1, y2 - y1, hDCMem2, 0, 0, vbSrcCopy)
    clrBkOld = SetBkColor(hDC, clrFr)
    clrTextOld = SetTextColor(hDC, clrBk)
    Call BitBlt(hDC, x1, y1, x2 - x1, y2 - y1, hDCMem1, 0, 0, vbSrcCopy)
        
    '-- Clean up
    Call SetBkColor(hDC, clrBkOld)
    Call SetTextColor(hDC, clrTextOld)
    Call DeleteObject(SelectObject(hDCMem1, hBmpOld1))
    Call DeleteObject(SelectObject(hDCMem2, hBmpOld2))
    Call DeleteDC(hDCMem1)
    Call DeleteDC(hDCMem2)
End Sub

Public Function ShiftColor(ByVal clr As Long, ByVal d As Long) As Long

  Dim R As Long, B As Long, G As Long

    R = (clr And &HFF) + d
    G = ((clr \ &H100) Mod &H100) + d
    B = ((clr \ &H10000) Mod &H100) + d
    
    If (d > 0) Then
        If (R > &HFF) Then R = &HFF
        If (G > &HFF) Then G = &HFF
        If (B > &HFF) Then B = &HFF
    ElseIf (d < 0) Then
        If (R < 0) Then R = 0
        If (G < 0) Then G = 0
        If (B < 0) Then B = 0
    End If
    ShiftColor = R + &H100& * G + &H10000 * B
End Function

Public Function TranslateColor(ByVal clr As OLE_COLOR) As Long
    
    Call OleTranslateColor(clr, 0, TranslateColor)
End Function

Private Sub pvDrawLine(ByVal hDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal clr As Long)
        
  Dim uPt     As POINTAPI
  Dim hPen    As Long
  Dim hOldPen As Long
        
    hPen = CreatePen(PS_SOLID, 1, clr)
    hOldPen = SelectObject(hDC, hPen)
    Call MoveToEx(hDC, x1, y1, uPt)
    Call LineTo(hDC, x2, y2)
    Call DeleteObject(SelectObject(hDC, hOldPen))
End Sub
