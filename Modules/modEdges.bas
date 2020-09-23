Attribute VB_Name = "modEdges"
' This module is used for drawing edges.

Option Explicit

Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hDc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hDc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDc As Long, ByVal x As Long, ByVal y As Long) As Long

Const PS_SOLID = 0

Private Type POINTAPI
   x As Long
   y As Long
End Type

Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Sub FlatEdge(ByVal hDc As Long, Box As RECT, lColour As Long)
   Dim lPen As Long
   lPen = CreatePen(PS_SOLID, 1, TranslateColour(lColour))
   SelectObject hDc, lPen
   Rectangle hDc, Box.Left, Box.Top, Box.Right, Box.Bottom
   DeleteObject lPen
End Sub

' EtchedEdge()
' -----------------------------
' I WAS going to use the DrawFrame API, except it only uses system
' colours. But there's a really easy way to obtain the same effect.
Sub EtchedEdge(ByVal hDc As Long, Box As RECT, hdColour As Long, _
   sdColour As Long)
   
   Dim lPen As Long
   
   lPen = CreatePen(PS_SOLID, 1, TranslateColour(hdColour))
   SelectObject hDc, lPen
   Rectangle hDc, Box.Left + 1, Box.Top + 1, Box.Right, Box.Bottom
   DeleteObject lPen
   
   lPen = CreatePen(PS_SOLID, 1, TranslateColour(sdColour))
   SelectObject hDc, lPen
   Rectangle hDc, Box.Left, Box.Top, Box.Right - 1, Box.Bottom - 1
   DeleteObject lPen
   
   SetPixel hDc, Box.Right - 1, Box.Top, TranslateColour(hdColour)
   SetPixel hDc, Box.Left, Box.Bottom - 1, TranslateColour(hdColour)
End Sub

' ThinEdge()
' -----------------------------
' This time there's no choice - we have to use lines.
Sub ThinEdge(ByVal hDc As Long, Box As RECT, lightColour As Long, _
   darkColour As Long)
   
   Dim lPen As Long
   Dim oldPoint As POINTAPI
   
   lPen = CreatePen(PS_SOLID, 1, TranslateColour(lightColour))
   DeleteObject SelectObject(hDc, lPen)
   MoveToEx hDc, Box.Right, Box.Top, oldPoint
   LineTo hDc, Box.Left, Box.Top
   LineTo hDc, Box.Left, Box.Bottom - 1
   DeleteObject lPen
   
   lPen = CreatePen(PS_SOLID, 1, TranslateColour(darkColour))
   DeleteObject SelectObject(hDc, lPen)
   LineTo hDc, Box.Right - 1, Box.Bottom - 1
   LineTo hDc, Box.Right - 1, Box.Top
   DeleteObject lPen
End Sub

Sub ThickEdge(ByVal hDc As Long, Box As RECT, hlightColour As Long, _
   lightColour As Long, sdarkColour As Long, darkColour As Long)
   
   Dim lPen As Long
   Dim oldPoint As POINTAPI
   
   lPen = CreatePen(PS_SOLID, 1, TranslateColour(hlightColour))
   SelectObject hDc, lPen
   MoveToEx hDc, Box.Right, Box.Top, oldPoint
   LineTo hDc, Box.Left, Box.Top
   LineTo hDc, Box.Left, Box.Bottom - 1
   DeleteObject lPen
   
   lPen = CreatePen(PS_SOLID, 1, TranslateColour(darkColour))
   SelectObject hDc, lPen
   LineTo hDc, Box.Right - 1, Box.Bottom - 1
   LineTo hDc, Box.Right - 1, Box.Top
   DeleteObject lPen
   
   lPen = CreatePen(PS_SOLID, 1, TranslateColour(lightColour))
   SelectObject hDc, lPen
   MoveToEx hDc, Box.Right - 2, Box.Top + 1, oldPoint
   LineTo hDc, Box.Left + 1, Box.Top + 1
   LineTo hDc, Box.Left + 1, Box.Bottom - 2
   DeleteObject lPen
   
   lPen = CreatePen(PS_SOLID, 1, TranslateColour(sdarkColour))
   SelectObject hDc, lPen
   LineTo hDc, Box.Right - 2, Box.Bottom - 2
   LineTo hDc, Box.Right - 2, Box.Top + 1
   DeleteObject lPen
End Sub

