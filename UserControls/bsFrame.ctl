VERSION 5.00
Begin VB.UserControl bsFrame 
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ClipControls    =   0   'False
   ControlContainer=   -1  'True
   MaskColor       =   &H00808000&
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "bsFrame.ctx":0000
End
Attribute VB_Name = "bsFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : BadSoft bsFrame control
' DateTime  : 14/02/2002
' Author    : Drew (aka The Bad One)
'             ©2002-2003 BadSoft Entertainment, all rights reserved.
'---------------------------------------------------------------------------------------

'Default Property Values:
Const m_def_MousePointer = 0
Const m_def_CaptionBGColour = vbWhite
Const m_def_CaptionBG = 0
Const m_def_CaptionPosition = 0
Const m_def_FrameStyle = 0
Const m_def_BorderStyle = 6
Const m_def_GradientColour1 = vbWhite
Const m_def_HighlightColour = vb3DHighlight
Const m_def_HighlightDKColour = vb3DLight
Const m_def_ShadowColour = vb3DShadow
Const m_def_ShadowDKColour = vb3DDKShadow
Const m_def_FlatBorderColour = vbInactiveBorder
Const m_def_GradientColour2 = vbButtonFace
Const m_def_CaptionColour = vbButtonText
Const m_def_BackColour = vbButtonFace

'Property Variables:
Dim m_MousePointer As MousePointerConstants
Dim m_MouseIcon As Picture
Dim m_CaptionBGColour As OLE_COLOR
Dim m_CaptionBG As bsfCaptionBG
Dim m_GradientColour1 As OLE_COLOR
Dim m_GradientColour2 As OLE_COLOR
Dim m_CaptionPosition As bsfCaptionPosition
Dim m_FlatBorderColour As OLE_COLOR
Dim m_FrameStyle As bsfFrameType
Dim m_CaptionColour As OLE_COLOR
Dim m_HighlightColour As OLE_COLOR
Dim m_HighlightDKColour As OLE_COLOR
Dim m_ShadowColour As OLE_COLOR
Dim m_ShadowDKColour As OLE_COLOR
Dim m_BackColour As OLE_COLOR
Dim m_Caption As String
Dim m_BorderStyle As bsfBorderStyle


' ENUMERATIONS
' ----------------------------
Public Enum bsfBorderStyle
   bsfNone
   bsfFlat
   bsfRaisedThin
   bsfRaised3D
   bsfSunkenThin
   bsfSunken3D
   bsfEtched
   bsfBump
End Enum

Public Enum bsfFrameType
   bsfStandardFrame
   bsfPlainFrame
   bsfHeaderFrame
   bsfHeaderOnly
End Enum

Public Enum bsfCaptionPosition
   bsfOnTop
   bsfOnBottom
   bsfOnLeft
   bsfOnRight
End Enum

Public Enum bsfCaptionBG
   bsfDefault
   bsfColour
   bsfGradientH
   bsfGradientV
End Enum

'Public Enum bsfCaptionAlign
'   bsfAlignLeft
'   bsfAlignCentre
'   bsfAlignRight
'End Enum

Private Type RGB
   Red   As Integer
   Green As Integer
   Blue  As Integer
End Type

Private Declare Function GradientFillRect Lib "msimg32" Alias "GradientFill" (ByVal hDc As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_RECT, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
Private Type TRIVERTEX
    X As Long
    Y As Long
    Red As Integer 'Ushort value
    Green As Integer 'Ushort value
    Blue As Integer 'ushort value
    Alpha As Integer 'ushort
End Type

Private Type GRADIENT_RECT
    UpperLeft As Long  'In reality this is a UNSIGNED Long
    LowerRight As Long 'In reality this is a UNSIGNED Long
End Type

Const GRADIENT_FILL_RECT_H As Long = &H0 'In this mode, two endpoints describe a rectangle. The rectangle is
'defined to have a constant color (specified by the TRIVERTEX structure) for the left and right edges. GDI interpolates
'the color from the top to bottom edge and fills the interior.
Const GRADIENT_FILL_RECT_V  As Long = &H1 'In this mode, two endpoints describe a rectangle. The rectangle
' is defined to have a constant color (specified by the TRIVERTEX structure) for the top and bottom edges. GDI interpolates
' the color from the top to bottom edge and fills the interior.
'Event Declarations:
Event Click()
Event DblClick()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Attribute MouseDown.VB_UserMemId = -605
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Attribute MouseMove.VB_UserMemId = -606
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Attribute MouseUp.VB_UserMemId = -607

Const DT_CALCRECT = &H400
Const DT_SINGLELINE = &H20

Private Type LOGFONT
   lfHeight As Long
   lfWidth As Long
   lfEscapement As Long
   lfOrientation As Long
   lfWeight As Long
   lfItalic As Byte
   lfUnderline As Byte
   lfStrikeOut As Byte
   lfCharSet As Byte
   lfOutPrecision As Byte
   lfClipPrecision As Byte
   lfQuality As Byte
   lfPitchAndFamily As Byte
   lfFacename As String * 33
End Type

Private Declare Function FillRect Lib "user32" (ByVal hDc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDc As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDc As Long, ByVal lpString As String, ByVal nCount As Long, lpRect As RECT, ByVal uFormat As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long

Private Sub DrawStandardFrame()

   Dim rectCaption As RECT
   Dim rectBlank As RECT
   Dim textHeight As Integer
   Dim strCaption As String
   
   strCaption = m_Caption
   Call DrawText(UserControl.hDc, strCaption, _
      Len(strCaption), rectCaption, DT_CALCRECT)
   textHeight = rectCaption.Bottom

   DoStandardEdges textHeight
            
   'rectangle behind text
   With rectCaption
      Select Case m_CaptionPosition
         Case bsfOnTop
            SetRect rectBlank, 6, 0, .Right + 10, .Bottom
         Case bsfOnBottom
            SetRect rectBlank, 6, ScaleHeight - textHeight, _
               .Right + 10, ScaleHeight
         Case bsfOnRight
            SetRect rectBlank, ScaleWidth - .Bottom, 6, _
               ScaleWidth, .Right + 10
         Case bsfOnLeft
            SetRect rectBlank, 0, ScaleHeight - .Right - 10, _
               textHeight, ScaleHeight - 6
      End Select
   End With
   lBrush = CreateSolidBrush(TranslateColour(BackColour))
   FillRect UserControl.hDc, rectBlank, lBrush
   DeleteObject lBrush
   
   'Draw text
   Select Case m_CaptionPosition
      Case bsfOnTop
         SetRect rectCaption, 8, rectCaption.Top, _
            rectCaption.Right + 8, rectCaption.Bottom
      Case bsfOnBottom
         SetRect rectCaption, 8, ScaleHeight - textHeight, _
            rectCaption.Right + 8, ScaleHeight
   End Select
   
   If UserControl.Enabled = False And _
      UserControl.Ambient.UserMode = True Then
   'We don't need a special API function, here's how to draw
   'disabled text.
      With rectCaption
         Select Case m_CaptionPosition
            Case bsfOnTop, bsfOnBottom
               SetRect rectCaption, .Left + 1, .Top + 1, _
                  .Right + 1, .Bottom + 1
               SetTextColor UserControl.hDc, _
                  TranslateColour(m_HighlightColour)
               Call DrawText(UserControl.hDc, strCaption, _
                  Len(strCaption), rectCaption, 0)
                  
               SetRect rectCaption, .Left - 1, .Top - 1, _
                  .Right - 1, .Bottom - 1
               SetTextColor UserControl.hDc, _
                  TranslateColour(m_ShadowColour)
               Call DrawText(UserControl.hDc, strCaption, _
                  Len(strCaption), rectCaption, 0)
            Case bsfOnLeft
               SetTextColor UserControl.hDc, _
                  TranslateColour(m_HighlightColour)
               DrawRotatedText 90, strCaption, 1, ScaleHeight - 5
               SetTextColor UserControl.hDc, _
                  TranslateColour(m_ShadowColour)
               DrawRotatedText 90, strCaption, 0, ScaleHeight - 6
            Case bsfOnRight
               SetTextColor UserControl.hDc, _
                  TranslateColour(m_HighlightColour)
               DrawRotatedText 270, strCaption, ScaleWidth + 1, 7
               SetTextColor UserControl.hDc, _
                  TranslateColour(m_ShadowColour)
               DrawRotatedText 270, strCaption, ScaleWidth, 6
         End Select
      End With
   Else
   'Draw text as normal.
      SetTextColor UserControl.hDc, _
         TranslateColour(m_CaptionColour)
      Select Case m_CaptionPosition
         Case bsfOnTop, bsfOnBottom
            Call DrawText(UserControl.hDc, strCaption, _
               Len(strCaption), rectCaption, DT_SINGLELINE)
         Case bsfOnLeft
            DrawRotatedText 90, strCaption, 0, ScaleHeight - 8
         Case bsfOnRight
            DrawRotatedText 270, strCaption, ScaleWidth, 6
      End Select
   End If

End Sub

Private Function GetRGBColours(lColour As Long) As RGB

   Dim HexColour As String
   lColour = TranslateColour(lColour)
   HexColour = String(6 - Len(Hex$(lColour)), "0") & Hex$(lColour)
   GetRGBColours.Red = "&H" & Mid$(HexColour, 5, 2) & "00"
   GetRGBColours.Green = "&H" & Mid$(HexColour, 3, 2) & "00"
   GetRGBColours.Blue = "&H" & Mid$(HexColour, 1, 2) & "00"

End Function
' BackColour()
' --------------------------
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get BackColour() As OLE_COLOR
Attribute BackColour.VB_Description = "The colour of the bsFrame's background."
Attribute BackColour.VB_ProcData.VB_Invoke_Property = ";Colour"
   BackColour = m_BackColour
End Property

Public Property Let BackColour(ByVal New_BackColour As OLE_COLOR)
   m_BackColour = New_BackColour
   PropertyChanged "BackColour"
   DrawFrame
End Property

' Caption()
' --------------------------
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get Caption() As String
Attribute Caption.VB_Description = "The text at the top of the bsFrame."
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Text"
   Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
   m_Caption = New_Caption
   PropertyChanged "Caption"
   DrawFrame
End Property

' BorderStyle()
' --------------------------
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=1,0,0,0
Public Property Get BorderStyle() As bsfBorderStyle
Attribute BorderStyle.VB_Description = "The style in which the bsFrame's edges are drawn."
Attribute BorderStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
   BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As bsfBorderStyle)
   m_BorderStyle = New_BorderStyle
   PropertyChanged "BorderStyle"
   DrawFrame
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
   m_Caption = UserControl.Extender.Name
   m_BorderStyle = m_def_BorderStyle
   m_FrameStyle = m_def_FrameStyle
   m_FlatBorderColour = m_def_FlatBorderColour
   m_HighlightColour = m_def_HighlightColour
   m_HighlightDKColour = TranslateColour(m_def_HighlightDKColour)
   m_ShadowColour = m_def_ShadowColour
   m_ShadowDKColour = m_def_ShadowDKColour
   m_CaptionColour = m_def_CaptionColour
   m_BackColour = m_def_BackColour
   Set UserControl.Font = Ambient.Font
   m_CaptionBG = m_def_CaptionBG
   m_GradientColour1 = m_def_GradientColour1
   m_GradientColour2 = m_def_GradientColour2
   m_CaptionPosition = m_def_CaptionPosition
   m_CaptionBGColour = m_def_CaptionBGColour
   m_MousePointer = m_def_MousePointer
   Set m_MouseIcon = LoadPicture("")
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   m_BackColour = PropBag.ReadProperty("BackColour", m_def_BackColour)
   m_Caption = PropBag.ReadProperty("Caption", UserControl.Extender.Name)
   m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
   m_HighlightColour = PropBag.ReadProperty("HighlightColour", m_def_HighlightColour)
   m_HighlightDKColour = PropBag.ReadProperty("HighlightDKColour", m_def_HighlightDKColour)
   m_ShadowColour = PropBag.ReadProperty("ShadowColour", m_def_ShadowColour)
   m_ShadowDKColour = PropBag.ReadProperty("ShadowDKColour", m_def_ShadowDKColour)
   m_CaptionColour = PropBag.ReadProperty("CaptionColour", m_def_CaptionColour)
   Set UserControl.Font = PropBag.ReadProperty("Fount", Ambient.Font)
   m_FrameStyle = PropBag.ReadProperty("FrameStyle", m_def_FrameStyle)
   m_FlatBorderColour = PropBag.ReadProperty("FlatBorderColour", m_def_FlatBorderColour)
   UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
   m_CaptionBG = PropBag.ReadProperty("CaptionBG", m_def_CaptionBG)
   m_GradientColour1 = PropBag.ReadProperty("GradientColour1", m_def_GradientColour1)
   m_GradientColour2 = PropBag.ReadProperty("GradientColour2", m_def_GradientColour2)
   m_CaptionPosition = PropBag.ReadProperty("CaptionPosition", m_def_CaptionPosition)
   m_CaptionBGColour = PropBag.ReadProperty("CaptionBGColour", m_def_CaptionBGColour)
   m_MousePointer = PropBag.ReadProperty("MousePointer", m_def_MousePointer)
   Set m_MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
End Sub

Private Sub UserControl_Resize()
   DrawFrame
End Sub

Private Sub UserControl_Show()
   'DrawFrame
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   Call PropBag.WriteProperty("BackColour", m_BackColour, m_def_BackColour)
   Call PropBag.WriteProperty("Caption", m_Caption, UserControl.Extender.Name)
   Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
   Call PropBag.WriteProperty("FlatBorderColour", m_FlatBorderColour, m_def_FlatBorderColour)
   Call PropBag.WriteProperty("HighlightColour", m_HighlightColour, m_def_HighlightColour)
   Call PropBag.WriteProperty("HighlightDKColour", m_HighlightDKColour, m_def_HighlightDKColour)
   Call PropBag.WriteProperty("ShadowColour", m_ShadowColour, m_def_ShadowColour)
   Call PropBag.WriteProperty("ShadowDKColour", m_ShadowDKColour, m_def_ShadowDKColour)
   Call PropBag.WriteProperty("CaptionColour", m_CaptionColour, m_def_CaptionColour)
   Call PropBag.WriteProperty("Fount", UserControl.Font, Ambient.Font)
   Call PropBag.WriteProperty("FrameStyle", m_FrameStyle, m_def_FrameStyle)
   Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
   Call PropBag.WriteProperty("CaptionBG", m_CaptionBG, m_def_CaptionBG)
   Call PropBag.WriteProperty("GradientColour1", m_GradientColour1, m_def_GradientColour1)
   Call PropBag.WriteProperty("GradientColour2", m_GradientColour2, m_def_GradientColour2)
   Call PropBag.WriteProperty("CaptionPosition", m_CaptionPosition, m_def_CaptionPosition)
   Call PropBag.WriteProperty("CaptionBGColour", m_CaptionBGColour, m_def_CaptionBGColour)
   Call PropBag.WriteProperty("MousePointer", m_MousePointer, m_def_MousePointer)
   Call PropBag.WriteProperty("MouseIcon", m_MouseIcon, Nothing)
End Sub

Private Sub DoPlainEdges()

   Dim rectArea As RECT
   
   rectArea.Right = UserControl.ScaleWidth
   rectArea.Bottom = UserControl.ScaleHeight
   Select Case m_BorderStyle
      Case bsfFlat
         FlatEdge UserControl.hDc, rectArea, m_FlatBorderColour
         
      Case bsfRaisedThin
         ThinEdge UserControl.hDc, rectArea, m_HighlightColour, _
            m_ShadowColour

      Case bsfSunkenThin
         ThinEdge UserControl.hDc, rectArea, m_ShadowColour, _
            m_HighlightColour
         
      Case bsfRaised3D
         ThickEdge UserControl.hDc, rectArea, m_HighlightColour, _
            m_HighlightDKColour, m_ShadowColour, m_ShadowDKColour
      
      Case bsfSunken3D
         ThickEdge UserControl.hDc, rectArea, m_ShadowDKColour, _
            m_ShadowColour, m_HighlightDKColour, m_HighlightColour
         
      Case bsfEtched
         EtchedEdge UserControl.hDc, rectArea, m_HighlightColour, _
            m_ShadowColour

      Case bsfBump
         EtchedEdge UserControl.hDc, rectArea, m_ShadowColour, _
            m_HighlightColour
   End Select
End Sub

Private Sub DoHeaderEdges(ByVal textHeight As Long)

   Dim rect1 As RECT
   Dim rect2 As RECT
   
   With rect1
      Select Case m_CaptionPosition
         Case bsfOnTop
            .Right = UserControl.ScaleWidth
            .Bottom = textHeight + 4
         Case bsfOnBottom
            .Top = UserControl.ScaleHeight - (textHeight + 4)
            .Right = UserControl.ScaleWidth
            .Bottom = UserControl.ScaleHeight
         Case bsfOnLeft
            .Right = textHeight + 4
            .Bottom = ScaleHeight
         Case bsfOnRight
            .Left = ScaleWidth - textHeight - 4
            .Right = ScaleWidth
            .Bottom = ScaleHeight
      End Select
   End With
   
   With rect2
      Select Case m_CaptionPosition
         Case bsfOnTop
            .Top = textHeight + IIf(m_BorderStyle = bsfFlat, 3, 4)
            .Right = UserControl.ScaleWidth
            .Bottom = ScaleHeight
         Case bsfOnBottom
            .Bottom = UserControl.ScaleHeight - (textHeight + _
               IIf(m_BorderStyle = bsfFlat, 3, 4))
            .Right = UserControl.ScaleWidth
         Case bsfOnLeft
            .Left = textHeight + IIf(m_BorderStyle = bsfFlat, 3, 4)
            .Right = ScaleWidth
            .Bottom = ScaleHeight
         Case bsfOnRight
            .Right = ScaleWidth - textHeight - _
               IIf(m_BorderStyle = bsfFlat, 3, 4)
            .Bottom = ScaleHeight
      End Select
   End With
   
   ' Header bit
   Select Case BorderStyle
      Case bsfFlat
         'header
         FlatEdge UserControl.hDc, rect1, m_FlatBorderColour
      Case bsfRaisedThin
         ThinEdge UserControl.hDc, rect1, m_HighlightColour, _
            m_ShadowColour
      Case bsfSunkenThin
         ThinEdge UserControl.hDc, rect1, m_ShadowColour, _
            m_HighlightColour
      Case bsfRaised3D
         ThickEdge UserControl.hDc, rect1, m_HighlightColour, _
            m_HighlightDKColour, m_ShadowColour, m_ShadowDKColour
      Case bsfSunken3D
         ThickEdge UserControl.hDc, rect1, m_ShadowDKColour, _
            m_ShadowColour, m_HighlightDKColour, m_HighlightColour
      Case bsfEtched
         EtchedEdge UserControl.hDc, rect1, m_HighlightColour, _
            m_ShadowColour
      Case bsfBump
         EtchedEdge UserControl.hDc, rect1, m_ShadowColour, _
            m_HighlightColour
   End Select
   
   If m_FrameStyle = bsfHeaderOnly Then
      Exit Sub
   End If
   
   ' Body bit
   Select Case BorderStyle
      Case bsfFlat
         FlatEdge UserControl.hDc, rect2, m_FlatBorderColour
      Case bsfRaisedThin
         ThinEdge UserControl.hDc, rect2, m_HighlightColour, _
            m_ShadowColour
      Case bsfSunkenThin
         ThinEdge UserControl.hDc, rect2, m_ShadowColour, _
            m_HighlightColour
      Case bsfRaised3D
         ThickEdge UserControl.hDc, rect2, m_HighlightColour, _
            m_HighlightDKColour, m_ShadowColour, m_ShadowDKColour
      Case bsfSunken3D
         ThickEdge UserControl.hDc, rect2, m_ShadowDKColour, _
            m_ShadowColour, m_HighlightDKColour, m_HighlightColour
      Case bsfEtched
         EtchedEdge UserControl.hDc, rect2, m_HighlightColour, _
            m_ShadowColour
      Case bsfBump
         EtchedEdge UserControl.hDc, rect2, m_ShadowColour, _
            m_HighlightColour
   End Select
      
End Sub

' DoStandardEdges()
' --------------------------
Private Sub DoStandardEdges(ByVal textHeight As Long)

   Dim rectFrame As RECT
   
   SetRect rectFrame, 0, 0, ScaleWidth, ScaleHeight
   
   With rectFrame
      Select Case m_CaptionPosition
         Case bsfOnTop
            .Top = .Top + textHeight / 2
         Case bsfOnBottom
            .Bottom = .Bottom - textHeight / 2
         Case bsfOnLeft
            .Left = .Left + textHeight / 2
         Case bsfOnRight
            .Right = .Right - textHeight / 2
      End Select
   End With
   
   Select Case BorderStyle
      Case bsfFlat
         FlatEdge UserControl.hDc, rectFrame, m_FlatBorderColour
         
      Case bsfRaisedThin
         ThinEdge UserControl.hDc, rectFrame, m_HighlightColour, _
            m_ShadowColour

      Case bsfSunkenThin
         ThinEdge UserControl.hDc, rectFrame, m_ShadowColour, _
            m_HighlightColour
         
      Case bsfRaised3D
         ThickEdge UserControl.hDc, rectFrame, m_HighlightColour, _
            m_HighlightDKColour, m_ShadowColour, m_ShadowDKColour
      
      Case bsfSunken3D
         ThickEdge UserControl.hDc, rectFrame, m_ShadowDKColour, _
            m_ShadowColour, m_HighlightDKColour, m_HighlightColour
         
      Case bsfEtched
         EtchedEdge UserControl.hDc, rectFrame, m_HighlightColour, _
            m_ShadowColour

      Case bsfBump
         EtchedEdge UserControl.hDc, rectFrame, m_ShadowColour, _
            m_HighlightColour
   End Select
End Sub

' DrawFrame()
' --------------------------
Private Sub DrawFrame()

   Dim rectCaption As RECT
   Dim strCaption As String
   Dim rgb1 As RGB, rgb2 As RGB
   Dim corner(1) As TRIVERTEX
   Dim gRect As GRADIENT_RECT
   Dim hBrush As Long
   Dim cbgRect As RECT
   Dim textHeight As Integer
   
'Clear everything
'--------------------------------
   UserControl.BackColor = TranslateColour(m_BackColour)
   Cls
   
' Drawing a standard frame is WAY different from drawing the other
' ones, so I've given it its own subroutine.
   If m_FrameStyle = bsfStandardFrame Then
      DrawStandardFrame
      Exit Sub
   End If
   
'Before we draw anything we need to calculate the size space the
'text will take up. This is so that the side that holds the text
'can be drawn properly.
'We use a separate string for checking the size of the caption
'because if it's too long we can truncate it with an ellipsis
'(...). Calling DrawText with the DT_CALCRECT flag doesn't draw
'any text.
   
   strCaption = m_Caption
   Call DrawText(UserControl.hDc, strCaption, Len(strCaption), rectCaption, DT_CALCRECT)
   textHeight = rectCaption.Bottom
   
'Caption background
'--------------------------------
   Select Case m_CaptionPosition
      Case bsfOnTop
         SetRect cbgRect, 0, 0, ScaleWidth, textHeight + 4
      Case bsfOnBottom
         SetRect cbgRect, 0, ScaleHeight - (textHeight + 4), ScaleWidth, ScaleHeight
      Case bsfOnLeft
         SetRect cbgRect, 0, 0, textHeight + 4, ScaleHeight
      Case bsfOnRight
         SetRect cbgRect, ScaleWidth - textHeight - 4, 0, ScaleWidth, ScaleHeight
   End Select
   
   Select Case m_CaptionBG
      Case bsfColour
         hBrush = _
            CreateSolidBrush(TranslateColour(m_CaptionBGColour))
         FillRect UserControl.hDc, cbgRect, hBrush
         DeleteObject hBrush
         
      Case bsfGradientH, bsfGradientV
         rgb1 = GetRGBColours(m_GradientColour1)
         rgb2 = GetRGBColours(m_GradientColour2)
         With corner(0)
            .Red = rgb1.Red
            .Green = rgb1.Green
            .Blue = rgb1.Blue
            .X = cbgRect.Left
            .Y = cbgRect.Top
         End With
         With corner(1)
            .Red = rgb2.Red
            .Green = rgb2.Green
            .Blue = rgb2.Blue
            .X = cbgRect.Right
            .Y = cbgRect.Bottom
         End With
         gRect.UpperLeft = 0
         gRect.LowerRight = 1
         GradientFillRect UserControl.hDc, corner(0), 2, gRect, 1, _
            IIf(m_CaptionBG = bsfGradientH, _
            GRADIENT_FILL_RECT_H, GRADIENT_FILL_RECT_V)
   End Select
   
   Select Case m_FrameStyle
      Case bsfPlainFrame
         DoPlainEdges
         rectCaption.Left = 4
         rectCaption.Top = 2
         
      Case Else
         DoHeaderEdges textHeight
         rectCaption.Left = 4
         rectCaption.Top = 2
   End Select
         
'Draw text
'--------------------------------
   rectCaption.Right = ScaleWidth - 2
   rectCaption.Top = rectCaption.Top + cbgRect.Top
   rectCaption.Bottom = rectCaption.Bottom + cbgRect.Top + 2
   
   If UserControl.Enabled = False And _
      UserControl.Ambient.UserMode = True Then
   'We don't need a special API function, here's how to draw
   'disabled text.
      With rectCaption
         Select Case m_CaptionPosition
            Case bsfOnTop, bsfOnBottom
               SetRect rectCaption, .Left + 1, .Top + 1, _
                  .Right + 1, .Bottom + 1
               SetTextColor UserControl.hDc, _
                  TranslateColour(m_HighlightColour)
               Call DrawText(UserControl.hDc, strCaption, _
                  Len(strCaption), rectCaption, 0)
                  
               SetRect rectCaption, .Left - 1, .Top - 1, _
                  .Right - 1, .Bottom - 1
               SetTextColor UserControl.hDc, _
                  TranslateColour(m_ShadowColour)
               Call DrawText(UserControl.hDc, strCaption, _
                  Len(strCaption), rectCaption, 0)
            Case bsfOnLeft
               SetTextColor UserControl.hDc, _
                  TranslateColour(m_HighlightColour)
               DrawRotatedText 90, strCaption, 3, ScaleHeight - 5
               SetTextColor UserControl.hDc, _
                  TranslateColour(m_ShadowColour)
               DrawRotatedText 90, strCaption, 2, ScaleHeight - 6
            Case bsfOnRight
               SetTextColor UserControl.hDc, _
                  TranslateColour(m_HighlightColour)
               DrawRotatedText 270, strCaption, ScaleWidth - 1, 7
               SetTextColor UserControl.hDc, _
                  TranslateColour(m_ShadowColour)
               DrawRotatedText 270, strCaption, ScaleWidth - 2, 6
         End Select
      End With
   Else
   'Draw text as normal.
      SetTextColor UserControl.hDc, _
         TranslateColour(m_CaptionColour)
      Select Case m_CaptionPosition
         Case bsfOnTop, bsfOnBottom
            Call DrawText(UserControl.hDc, strCaption, _
               Len(strCaption), rectCaption, DT_SINGLELINE)
         Case bsfOnLeft
            DrawRotatedText 90, strCaption, 2, ScaleHeight - 6
         Case bsfOnRight
            DrawRotatedText 270, strCaption, ScaleWidth - 2, 6
      End Select
   End If
End Sub

Private Sub DrawRotatedText(Angle As Integer, strText As String, X As Integer, Y As Integer)
  On Error GoTo GetOut
  
  Dim F As LOGFONT, hPrevFont As Long, hFont As Long
  Dim FontName As String

  F.lfEscapement = 10 * Angle
  FontName = UserControl.FontName + Chr$(0)
  F.lfFacename = FontName
  F.lfHeight = (UserControl.FontSize * -20) / Screen.TwipsPerPixelY
  F.lfWeight = IIf(UserControl.FontBold, 700, 0)
  F.lfItalic = UserControl.FontItalic
  F.lfStrikeOut = UserControl.FontStrikethru
  F.lfUnderline = UserControl.FontUnderline
  hFont = CreateFontIndirect(F)
  hPrevFont = SelectObject(UserControl.hDc, hFont)
  CurrentX = X
  CurrentY = Y
  Print strText

  hFont = SelectObject(UserControl.hDc, hPrevFont)
  DeleteObject hFont
  
  Exit Sub
GetOut:
  Exit Sub

End Sub
' FlatBorderColour()
' --------------------------
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get FlatBorderColour() As OLE_COLOR
Attribute FlatBorderColour.VB_Description = "The colour of the edges when BorderStyle is set to bsfFlat."
Attribute FlatBorderColour.VB_ProcData.VB_Invoke_Property = ";Colour"
   FlatBorderColour = m_FlatBorderColour
End Property

Public Property Let FlatBorderColour(ByVal New_FlatBorderColour As OLE_COLOR)
   m_FlatBorderColour = New_FlatBorderColour
   PropertyChanged "FlatBorderColour"
   DrawFrame
End Property

' HighlightColour()
' --------------------------
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get HighlightColour() As OLE_COLOR
Attribute HighlightColour.VB_Description = "The colour of the lightest border colour."
Attribute HighlightColour.VB_ProcData.VB_Invoke_Property = ";Colour"
   HighlightColour = m_HighlightColour
End Property

Public Property Let HighlightColour(ByVal New_HighlightColour As OLE_COLOR)
   m_HighlightColour = New_HighlightColour
   PropertyChanged "HighlightColour"
   DrawFrame
End Property

' HighlightDKColour()
' --------------------------
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get HighlightDKColour() As OLE_COLOR
Attribute HighlightDKColour.VB_Description = "The colour of the second lightest border colour."
Attribute HighlightDKColour.VB_ProcData.VB_Invoke_Property = ";Colour"
   HighlightDKColour = m_HighlightDKColour
End Property

Public Property Let HighlightDKColour(ByVal New_HighlightDKColour As OLE_COLOR)
   m_HighlightDKColour = New_HighlightDKColour
   PropertyChanged "HighlightDKColour"
   DrawFrame
End Property

' ShadowColour()
' --------------------------
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get ShadowColour() As OLE_COLOR
Attribute ShadowColour.VB_Description = "The second darkest border colour."
Attribute ShadowColour.VB_ProcData.VB_Invoke_Property = ";Colour"
   ShadowColour = m_ShadowColour
End Property

Public Property Let ShadowColour(ByVal New_ShadowColour As OLE_COLOR)
   m_ShadowColour = New_ShadowColour
   PropertyChanged "ShadowColour"
   DrawFrame
End Property

' ShadowDKColour()
' --------------------------
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get ShadowDKColour() As OLE_COLOR
Attribute ShadowDKColour.VB_Description = "The darkest border colour."
Attribute ShadowDKColour.VB_ProcData.VB_Invoke_Property = ";Colour"
   ShadowDKColour = m_ShadowDKColour
End Property

Public Property Let ShadowDKColour(ByVal New_ShadowDKColour As OLE_COLOR)
   m_ShadowDKColour = New_ShadowDKColour
   PropertyChanged "ShadowDKColour"
   DrawFrame
End Property

' CaptionColour()
' --------------------------
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,vbButtonText
Public Property Get CaptionColour() As OLE_COLOR
Attribute CaptionColour.VB_Description = "The colour of the bsFrame's Caption text."
Attribute CaptionColour.VB_ProcData.VB_Invoke_Property = ";Colour"
   CaptionColour = m_CaptionColour
End Property

Public Property Let CaptionColour(ByVal New_CaptionColour As OLE_COLOR)
   m_CaptionColour = New_CaptionColour
   PropertyChanged "CaptionColour"
   DrawFrame
End Property

' Fount()
' --------------------------
' Before you say anything, Fount is the English word for Font.
' Hence Font is American. Because I'm British, I use British
' words. Unlike many of you I haven't sold myself out.

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Fount() As Font
Attribute Fount.VB_Description = "The font used for the Caption. (Fount is the English word for font.)"
Attribute Fount.VB_ProcData.VB_Invoke_Property = ";Fount"
   Set Fount = UserControl.Font
End Property

Public Property Set Fount(ByVal New_Fount As Font)
   Set UserControl.Font = New_Fount
   PropertyChanged "Fount"
   DrawFrame
End Property

' FrameStyle()
' --------------------------
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=21,0,0,0
Public Property Get FrameStyle() As bsfFrameType
Attribute FrameStyle.VB_Description = "The design of the bsFrame."
Attribute FrameStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
   FrameStyle = m_FrameStyle
End Property

Public Property Let FrameStyle(ByVal New_FrameStyle As bsfFrameType)
   m_FrameStyle = New_FrameStyle
   PropertyChanged "FrameStyle"
   DrawFrame
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Whether or not the control can respond to the user's actions."
Attribute Enabled.VB_UserMemId = -514
   Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
   UserControl.Enabled() = New_Enabled
   PropertyChanged "Enabled"
   For I = 0 To ContainedControls.Count - 1
      UserControl.ContainedControls(I).Enabled = New_Enabled
   Next
End Property

Public Property Get CaptionBG() As bsfCaptionBG
Attribute CaptionBG.VB_Description = "The style of the frame caption's background."
Attribute CaptionBG.VB_ProcData.VB_Invoke_Property = ";Appearance"
   CaptionBG = m_CaptionBG
End Property

Public Property Let CaptionBG(ByVal New_CaptionBG As bsfCaptionBG)
   m_CaptionBG = New_CaptionBG
   PropertyChanged "CaptionBG"
   DrawFrame
End Property

Public Property Get GradientColour1() As OLE_COLOR
Attribute GradientColour1.VB_Description = "The first of the gradient colours."
Attribute GradientColour1.VB_ProcData.VB_Invoke_Property = ";Colour"
   GradientColour1 = m_GradientColour1
End Property

Public Property Let GradientColour1(ByVal New_GradientColour1 As OLE_COLOR)
   m_GradientColour1 = New_GradientColour1
   PropertyChanged "GradientColour1"
   DrawFrame
End Property

Public Property Get GradientColour2() As OLE_COLOR
Attribute GradientColour2.VB_Description = "The second of the gradient colours."
Attribute GradientColour2.VB_ProcData.VB_Invoke_Property = ";Colour"
   GradientColour2 = m_GradientColour2
   DrawFrame
End Property

Public Property Let GradientColour2(ByVal New_GradientColour2 As OLE_COLOR)
   m_GradientColour2 = New_GradientColour2
   PropertyChanged "GradientColour2"
   DrawFrame
End Property

Public Property Get CaptionPosition() As bsfCaptionPosition
Attribute CaptionPosition.VB_Description = "The position of the caption."
Attribute CaptionPosition.VB_ProcData.VB_Invoke_Property = ";Appearance"
   CaptionPosition = m_CaptionPosition
End Property

Public Property Let CaptionPosition(ByVal New_CaptionPosition As bsfCaptionPosition)
   m_CaptionPosition = New_CaptionPosition
   PropertyChanged "CaptionPosition"
   DrawFrame
End Property

Public Property Get CaptionBGColour() As OLE_COLOR
Attribute CaptionBGColour.VB_Description = "The background colour of the caption section."
Attribute CaptionBGColour.VB_ProcData.VB_Invoke_Property = ";Colour"
   CaptionBGColour = m_CaptionBGColour
End Property

Public Property Let CaptionBGColour(ByVal New_CaptionBGColour As OLE_COLOR)
   m_CaptionBGColour = New_CaptionBGColour
   PropertyChanged "CaptionBGColour"
   DrawFrame
End Property
'
'Public Property Get CaptionAlign() As bsfCaptionAlign
'   CaptionAlign = m_CaptionAlign
'End Property
'
'Public Property Let CaptionAlign(ByVal New_CaptionAlign As bsfCaptionAlign)
'   m_CaptionAlign = New_CaptionAlign
'   PropertyChanged "CaptionAlign"
'End Property
'
Private Sub UserControl_Click()
   RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
   RaiseEvent DblClick
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
Attribute MousePointer.VB_ProcData.VB_Invoke_Property = ";Misc"
   MousePointer = m_MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
   m_MousePointer = New_MousePointer
   PropertyChanged "MousePointer"
End Property

Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
Attribute MouseIcon.VB_ProcData.VB_Invoke_Property = ";Misc"
   Set MouseIcon = m_MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
   Set m_MouseIcon = New_MouseIcon
   PropertyChanged "MouseIcon"
End Property

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

