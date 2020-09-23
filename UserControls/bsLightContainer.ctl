VERSION 5.00
Begin VB.UserControl bsLightContainer 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
End
Attribute VB_Name = "bsLightContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function Rectangle Lib "gdi32" (ByVal hDc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateHatchBrush Lib "gdi32" (ByVal nIndex As Long, ByVal crColor As Long) As Long

Const HS_DIAGCROSS = 2
Const RUNTIME = True

Public Event Resize()

Private Sub UserControl_Paint()
   Highlight
End Sub

Private Sub UserControl_Resize()
   RaiseEvent Resize
    Highlight
End Sub

Private Sub Highlight()
   Cls
   If Ambient.UserMode <> RUNTIME Then
      hPen = CreatePen(0, 1, vbBlack)
      hBrush = CreateHatchBrush(HS_DIAGCROSS, vbBlack)
      DeleteObject SelectObject(UserControl.hDc, hPen)
      DeleteObject SelectObject(UserControl.hDc, hBrush)
      Rectangle UserControl.hDc, 0, 0, ScaleX(ScaleWidth, ScaleMode, vbPixels), _
         ScaleY(ScaleHeight, ScaleMode, vbPixels)
      DeleteObject lPen
      DeleteObject hBrush
    End If
End Sub

Private Sub UserControl_Show()
    Highlight
End Sub

Public Sub Refresh()
   Highlight
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleMode
Public Property Get ScaleMode() As ScaleModeConstants
Attribute ScaleMode.VB_Description = "Returns/sets a value indicating measurement units for object coordinates when using graphics methods or positioning controls."
   ScaleMode = UserControl.ScaleMode
End Property

Public Property Let ScaleMode(ByVal New_ScaleMode As ScaleModeConstants)
   UserControl.ScaleMode() = New_ScaleMode
   PropertyChanged "ScaleMode"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleWidth
Public Property Get ScaleWidth() As Single
Attribute ScaleWidth.VB_Description = "Returns/sets the number of units for the horizontal measurement of an object's interior."
   ScaleWidth = UserControl.ScaleWidth
End Property

Public Property Let ScaleWidth(ByVal New_ScaleWidth As Single)
   UserControl.ScaleWidth() = New_ScaleWidth
   PropertyChanged "ScaleWidth"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleHeight
Public Property Get ScaleHeight() As Single
Attribute ScaleHeight.VB_Description = "Returns/sets the number of units for the vertical measurement of an object's interior."
   ScaleHeight = UserControl.ScaleHeight
End Property

Public Property Let ScaleHeight(ByVal New_ScaleHeight As Single)
   UserControl.ScaleHeight() = New_ScaleHeight
   PropertyChanged "ScaleHeight"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

   UserControl.ScaleMode = PropBag.ReadProperty("ScaleMode", 3)
   UserControl.ScaleWidth = PropBag.ReadProperty("ScaleWidth", 320)
   UserControl.ScaleHeight = PropBag.ReadProperty("ScaleHeight", 240)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

   Call PropBag.WriteProperty("ScaleMode", UserControl.ScaleMode, 3)
   Call PropBag.WriteProperty("ScaleWidth", UserControl.ScaleWidth, 320)
   Call PropBag.WriteProperty("ScaleHeight", UserControl.ScaleHeight, 240)
End Sub

