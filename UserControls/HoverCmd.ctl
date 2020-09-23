VERSION 5.00
Begin VB.UserControl HoverCommand 
   ClientHeight    =   330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1335
   ScaleHeight     =   330
   ScaleWidth      =   1335
End
Attribute VB_Name = "HoverCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ************************************************************************
'  The EliteVB Highlight button control
' ************************************************************************
'
'      Written By: The Hand aka G.D. Sever
'      Written On: 09/13/01
'      Updated On: 01/15/02
'
'     Description: a button control that changes states when the
'                  cursor is hovering over the control. More than one
'                  style can be set at a time.
'
'  Terms of use:
'    Ahhh.... the sticky part. Here's the deal: Use this code in your projects.
'    For the love of god, use it and make your apps prettier. We need to start
'    showing people that VB has every bit as much potential as the C++ apps out
'    there.
'
'    Change the code if you want to! Don't like the way I do something?
'    Want to have different highlight colors available?  GO NUTS! Want to make
'    the image placeable left/right/top/bottom/center? ITS ALL YOU! Can't live
'    without a built-in animation loop with an array of StdPictures? That's
'    totally up to you.
'
'    HOWEVER: If you happen to post this code or a portion of it somewhere, give
'    me credit for the parts I am responsible for. Saying that I was
'    "an inspiration" for the code when 70% of it was cut & paste from here is NOT
'    adequate to me. Put a 1 or two line comment at the beginning of the subs and
'    functions you use and name me as the author! And let me tell you something...
'    Doing a global "Replace all" on a couple of variable names and function names
'    does not suddenly make the code something you wrote.
'
'    I really do not think I'm asking too much - it all boils down to one simple
'    principle: Give credit where its due. I do wherever necessary and possible.
'
'    That being said, API declarations are almost 100% from the ALLAPI.NET guide
'    which is a fantastic resource. Go out and download it immediately from
'    www.allapi.net
'
' ********************************************************************************
'      Visit http://www.EliteVB.com for more high-powered solutions!!
' ********************************************************************************

Option Explicit

' Declare styles available to the control. By using 2^x powers, we can
'  easily create combinations of styles and detect them in the code.
Public Enum hcStyle
    hcsNone = 0
    hcsPopupBorder = 1
    hcsThinBorder = 2
    hcsThickBorder = 4
    hcsTextColor = 8
    hcsNoBorder = 16
    hcsNoFocusRect = 32
    hcsHighlightBG = 64
End Enum

' Horizontal alignment constants
Public Enum hcHorizAlign
    hcaLeft = 0
    hcaCenter = 1
    hcaRight = 2
End Enum

' A rectangle type used by FillRect and DrawEdge.
Private Type RECT
   left As Long
   top As Long
   right As Long
   bottom As Long
End Type

' Type used by GetTextMetrics to determine various font attributes.
' Really only used so we know what values to pass to the CreateFont API.
Private Type TEXTMETRIC
    tmHeight As Long
    tmAscent As Long
    tmDescent As Long
    tmInternalLeading As Long
    tmExternalLeading As Long
    tmAveCharWidth As Long
    tmMaxCharWidth As Long
    tmWeight As Long
    tmOverhang As Long
    tmDigitizedAspectX As Long
    tmDigitizedAspectY As Long
    tmFirstChar As Byte
    tmLastChar As Byte
    tmDefaultChar As Byte
    tmBreakChar As Byte
    tmItalic As Byte
    tmUnderlined As Byte
    tmStruckOut As Byte
    tmPitchAndFamily As Byte
    tmCharSet As Byte
End Type

' Various Constants used by the DrawEdge function.
Private Const BF_LEFT = &H1
Private Const BF_BOTTOM = &H8
Private Const BF_RIGHT = &H4
Private Const BF_TOP = &H2
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Private Const BDR_INNER = &HC
Private Const BDR_OUTER = &H3
Private Const BDR_RAISED = &H5
Private Const BDR_RAISEDINNER = &H4
Private Const BDR_RAISEDOUTER = &H1
Private Const BDR_SUNKEN = &HA
Private Const BDR_SUNKENINNER = &H8
Private Const BDR_SUNKENOUTER = &H2
Private Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
Private Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Private Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Private Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
' Used to create the back buffer and memory DCs required when drawing
'  DC = Device context, which is basically like an easel which holds the canvas on which to draw.
'  Bmp = Bitmap, which is the paper or canvas on which we draw
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
' Used to manipulate our graphics resources and clean up afterwards
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
' Used to draw bitmaps on the back buffer (memory DC)
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
' Used to draw icons on the back buffer (memory DC)
Private Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal hIcon As Long) As Long
' Used to size the font and print to our back buffer (memory DC)
Private Type POINTAPI
    x As Long
    Y As Long
End Type
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As POINTAPI) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hdc As Long, lpMetrics As TEXTMETRIC) As Long
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal I As Long, ByVal u As Long, ByVal S As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long
' Used to create the background for the control
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
' These are used to create the mask for our picture one pixel at a time in the
'  doPicture method. While not the fastest method, what do you expect for a
'  $1.50 control.
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal crColor As Long) As Long
' These are used in the MouseMove event to detect when the cursor is inside &
'  outside the control.
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
' *************************
'   For disabled items:
' *************************
Private Const DST_PREFIXTEXT = &H2
Private Const DST_ICON = &H3
Private Const DST_BITMAP = &H4
Private Const DSS_NORMAL = &H0
Private Const DSS_DISABLED = &H20
Private Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hdc As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal flags As Long) As Long
Private Declare Function DrawStateText Lib "user32" Alias "DrawStateA" (ByVal hdc As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lString As String, ByVal wParam As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal flags As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBrushOrgEx Lib "gdi32" (ByVal hdc As Long, ByVal nXOrg As Long, ByVal nYOrg As Long, lppt As POINTAPI) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Sub OleTranslateColor Lib "oleaut32.dll" (ByVal clr As Long, ByVal hpal As Long, ByRef lpcolorref As Long)

' ******************************************************************
'  Local variables for setting the UserControl's properties.
' ******************************************************************
Private mCap As String                  ' Caption string
Private mPic As StdPicture              ' Regular picture image
Private mHovPic As StdPicture           ' Highlight picture image
Private mBGPic As StdPicture            ' Background image
Private mStyle As Long                  ' Style of the button. Can be a combination
                                        '  of various constants.
Private mState As Integer               ' State of the button.
                                        '  0 = no focus, cursor outside
                                        '  1 = mouse hovering over button
                                        '  2 = button depressed
Private mGrabBG As Boolean              ' Draw the background by grabbing its container's hDC
Private mHAlign As hcHorizAlign         ' Horizontal alignment (left/center/right)
Private mImageOffset As Integer         ' Offset of the picture from the caption
Private mBorderOffset As Integer        ' Offset of pic & caption from the border
Private mEnabled As Boolean             ' Whether or not control is disabled
Private mGotFocus As Boolean            ' Whether the usercontrol has focus or not
Private mForeColor As OLE_COLOR         ' Foreground (text) color
Private mBackColor As OLE_COLOR         ' Background (button) color
Private mHighlightColor As OLE_COLOR    ' Highlight color

' ******************************************************************
'  UserControl's generated events
' ******************************************************************
Public Event Click()
Public Event MouseLeave()
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

' ******************************************************************
'  Routines for setting the UserControl's properties.
' ******************************************************************
Public Property Let Tag(aVal As Variant)
    UserControl.Extender.Tag = aVal
End Property
Public Property Get Tag() As Variant
    Tag = UserControl.Extender.Tag
End Property
Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property
Public Property Let GrabBG(aBool As Boolean)
    If mGrabBG <> aBool Then
        mGrabBG = aBool
        PropertyChanged "GrabBG"
        UserControl_Paint
    End If
End Property

Public Property Get GrabBG() As Boolean
    GrabBG = mGrabBG
End Property
Public Property Let HAlign(anAlign As hcHorizAlign)
    If mHAlign <> anAlign Then
        mHAlign = anAlign
        PropertyChanged "HAlign"
        UserControl_Paint
    End If
End Property
Public Property Get HAlign() As hcHorizAlign
    HAlign = mHAlign
End Property
Public Property Let ImageOffset(anOffset As Integer)
    If mImageOffset <> anOffset Then
        mImageOffset = anOffset
        PropertyChanged "ImageOffset"
        UserControl_Paint
    End If
End Property
Public Property Get ImageOffset() As Integer
    ImageOffset = mImageOffset
End Property
Public Property Let BorderOffset(anOffset As Integer)
    If mBorderOffset <> anOffset Then
        mBorderOffset = anOffset
        PropertyChanged "BorderOffset"
        UserControl_Paint
    End If
End Property
Public Property Get BorderOffset() As Integer
    BorderOffset = mBorderOffset
End Property
Public Property Set Font(aFont As StdFont)
    Set UserControl.Font = aFont
    PropertyChanged "Font"
    UserControl_Paint
End Property
Public Property Get Font() As StdFont
    Set Font = UserControl.Font
End Property
Public Property Let Style(aStyle As Long)
    If mStyle <> aStyle Then
        mStyle = aStyle
        PropertyChanged "Style"
        UserControl_Paint
    End If
End Property
Public Property Get Style() As Long
    Style = mStyle
End Property
Public Property Set Picture(aStdPic As StdPicture)
    Set mPic = aStdPic
    PropertyChanged "Picture"
    UserControl_Paint
End Property
Public Property Get Picture() As StdPicture
    Set Picture = mPic
End Property
Public Property Set HovPicture(aStdPic As StdPicture)
    Set mHovPic = aStdPic
    PropertyChanged "HovPicture"
    UserControl_Paint
End Property
Public Property Get HovPicture() As StdPicture
    Set HovPicture = mHovPic
End Property
Public Property Set BGPicture(aStdPic As StdPicture)
    Set mBGPic = aStdPic
    PropertyChanged "BGPicture"
    UserControl_Paint
End Property
Public Property Get BGPicture() As StdPicture
    Set BGPicture = mBGPic
End Property
Public Property Let Caption(aCap As String)
    If aCap <> mCap Then
        mCap = aCap
        PropertyChanged "Caption"
        DoEvents
        UserControl_Paint
    End If
End Property
Public Property Get Caption() As String
    Caption = mCap
End Property
Public Property Let Enabled(aBool As Boolean)
    If mEnabled <> aBool Then
        mEnabled = aBool
        PropertyChanged "Enabled"
        DoEvents
        UserControl_Paint
    End If
    UserControl.Enabled = aBool
End Property
Public Property Get Enabled() As Boolean
    Enabled = mEnabled
End Property
Public Property Let ForeColor(aColor As OLE_COLOR)
    If mForeColor <> aColor Then
        mForeColor = aColor
        PropertyChanged "ForeColor"
        UserControl_Paint
    End If
End Property
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = mForeColor
End Property
Public Property Let BackColor(aColor As OLE_COLOR)
    If mBackColor <> aColor Then
        mBackColor = aColor
        PropertyChanged "BackColor"
        UserControl_Paint
    End If
End Property
Public Property Get BackColor() As OLE_COLOR
    BackColor = mBackColor
End Property
Public Property Let HighlightColor(aColor As OLE_COLOR)
    If mHighlightColor <> aColor Then
        mHighlightColor = aColor
        PropertyChanged "HighlightColor"
        UserControl_Paint
    End If
End Property
Public Property Get HighlightColor() As OLE_COLOR
    HighlightColor = mHighlightColor
End Property
' ******************************************************************
'  Routines for to save & retrieve the UserControl's properties,
'   giving these values some persistance.
' ******************************************************************
Public Sub Refresh()
    UserControl_Paint
End Sub

Private Sub UserControl_InitProperties()
    mCap = UserControl.Name
    mStyle = hcsThinBorder
    Set mHovPic = Nothing
    Set mPic = Nothing
    Set mBGPic = Nothing
    mImageOffset = 10
    mBorderOffset = 10
    mHAlign = hcaLeft
    mGrabBG = False
    mEnabled = True
    mForeColor = CLng("&H80000007")
    mBackColor = CLng("&H8000000F")
    mHighlightColor = CLng("&H8000000D")
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mCap = PropBag.ReadProperty("Caption", UserControl.Name)
    mStyle = PropBag.ReadProperty("Style", hcsThinBorder)
    Set mHovPic = PropBag.ReadProperty("HovPicture", Nothing)
    Set mPic = PropBag.ReadProperty("Picture", Nothing)
    Set mBGPic = PropBag.ReadProperty("BGPicture", Nothing)
    Set UserControl.Font = PropBag.ReadProperty("Font", UserControl.Font)
    mImageOffset = PropBag.ReadProperty("ImageOffset", 10)
    mBorderOffset = PropBag.ReadProperty("BorderOffset", 10)
    mHAlign = PropBag.ReadProperty("HAlign", hcaLeft)
    mGrabBG = PropBag.ReadProperty("GrabBG", False)
    mEnabled = PropBag.ReadProperty("Enabled", True)
    mForeColor = PropBag.ReadProperty("ForeColor", CLng("&H80000007"))
    mBackColor = PropBag.ReadProperty("BackColor", CLng("&H8000000F"))
    mHighlightColor = PropBag.ReadProperty("HighlightColor", CLng("&H8000000D"))
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Caption", mCap, UserControl.Name
    PropBag.WriteProperty "Style", mStyle, hcsThinBorder
    PropBag.WriteProperty "Picture", mPic, Nothing
    PropBag.WriteProperty "HovPicture", mHovPic, Nothing
    PropBag.WriteProperty "BGPicture", mBGPic, Nothing
    PropBag.WriteProperty "Font", UserControl.Font
    PropBag.WriteProperty "ImageOffset", mImageOffset, 10
    PropBag.WriteProperty "BorderOffset", mBorderOffset, 10
    PropBag.WriteProperty "HAlign", mHAlign, hcaLeft
    PropBag.WriteProperty "GrabBG", mGrabBG, False
    PropBag.WriteProperty "Enabled", mEnabled, True
    PropBag.WriteProperty "ForeColor", mForeColor, CLng("&H80000007")
    PropBag.WriteProperty "BackColor", mBackColor, CLng("&H8000000F")
    PropBag.WriteProperty "HighlightColor", mHighlightColor, CLng("&H8000000D")
End Sub
' ******************************************************************
'  Routines to handle the UserControl's events and paint the proper
'   picture.
' ******************************************************************
Private Sub UserControl_Click()
    mState = 0
    UserControl_Paint
    If mEnabled Then RaiseEvent Click
End Sub
Private Sub UserControl_GotFocus()
    If Not mEnabled Then Exit Sub
    mGotFocus = True
    UserControl_Paint
End Sub
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not mEnabled Then Exit Sub
    If (KeyCode = 32 Or KeyCode = 13) And mState <> 2 Then
        mState = 2
        UserControl_Paint
    End If
End Sub
Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    If Not mEnabled Then Exit Sub
    If (KeyCode = 32 Or KeyCode = 13) And mState = 2 Then
        mState = 0
        UserControl_Paint
        RaiseEvent Click
    End If
End Sub
Private Sub UserControl_LostFocus()
    If Not mEnabled Then Exit Sub
    mGotFocus = False
    UserControl_Paint
End Sub
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    ' Button is down. Draw the control in a "depressed" state.
    If Not mEnabled Then Exit Sub
    If Button = 1 Then
        mState = 2
        UserControl_Paint
    End If
    RaiseEvent MouseDown(Button, Shift, x, Y)
End Sub
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Not mEnabled Then Exit Sub
    ' If our mouse is moving inside of the control's borders...
    If Ambient.UserMode And mEnabled Then
        If x > 0 And x < UserControl.Width And _
           Y > 0 And Y < UserControl.Height Then
            ' Make all messages get sent to the UserControl for a while
            SetCapture UserControl.hwnd
            If mState <> 1 And Button <> 1 Then
                ' Repaint because we were in a different state
                mState = 1
                UserControl_Paint
            End If
            RaiseEvent MouseMove(Button, Shift, x, Y)
        Else
            ' Cursor went outside of the control. Release messages to be sent
            '  to wherever. Repaint the control with a "Lost focus" state
            ReleaseCapture
            mState = 0
            UserControl_Paint
            RaiseEvent MouseLeave
        End If
    End If
End Sub
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    ' Button is up... we assume over the button b/c of the SetCapture.
    '  So set our state to 1 and redraw...
    If Not mEnabled Then Exit Sub
    mState = 1
    UserControl_Paint
    RaiseEvent MouseUp(Button, Shift, x, Y)

End Sub

Private Sub UserControl_Paint()
    Dim aRect As RECT       ' a rectangle structure. Duh.
    Dim lBrush As Long      ' a brush used to paint the background
    Dim memDC As Long       ' Our memory device context, which we use as a
                            '  back buffer for flicker-free drawing.
    Dim memBmp As Long      ' The bitmap contained in our back buffer.
    Dim aPt As POINTAPI
    Dim lOffSet As Long
    
    ' Calculate the size of the control in pixels.
    aRect.right = UserControl.ScaleX(UserControl.Width, UserControl.ScaleMode, vbPixels)
    aRect.bottom = UserControl.ScaleY(UserControl.Height, UserControl.ScaleMode, vbPixels)
      
    ' Create a back buffer for the user control. This allows us to paint all
    '  of the control in memory first, and then copy it to the control
    '  once we've finished.
    memDC = CreateCompatibleDC(UserControl.hdc)
    memBmp = CreateCompatibleBitmap(UserControl.hdc, aRect.right, aRect.bottom)
    DeleteObject SelectObject(memDC, memBmp)
    ' If the background property is set, use the bitmap. Otherwise, just use
    '  the system's color value for ButtonFace... (&HF or SysColor 15)
    If mState = 2 Then lOffSet = IIf((mStyle And hcsThinBorder) = hcsThinBorder, 1, 2)
    If mState = 1 And (mStyle And hcsHighlightBG) = hcsHighlightBG Then
        lBrush = CreateSolidBrush(TranslateColor(mHighlightColor))
        ' Paint the background onto our control's back buffer.
        FillRect memDC, aRect, lBrush
        ' Clean up our graphics brush resource
        DeleteObject lBrush
    Else
        If mGrabBG And mBGPic Is Nothing Then
            BitBlt memDC, lOffSet, lOffSet, _
                   aRect.right, aRect.bottom, UserControl.Extender.Parent.hdc, _
                   UserControl.ScaleX(UserControl.Extender.left, UserControl.ScaleMode, vbPixels), _
                   UserControl.ScaleX(UserControl.Extender.top, UserControl.ScaleMode, vbPixels), _
                   vbSrcCopy
        Else
            If mBGPic Is Nothing Then
                lBrush = CreateSolidBrush(TranslateColor(mBackColor))
            Else
                lBrush = CreatePatternBrush(mBGPic.Handle)
                If mState = 2 Then SetBrushOrgEx memDC, lOffSet, lOffSet, aPt
                DeleteObject SelectObject(memDC, lBrush)
            End If
            ' Paint the background onto our control's back buffer.
            FillRect memDC, aRect, lBrush
            ' Clean up our graphics brush resource
            DeleteObject lBrush
        End If
    End If
    ' Paint the picture, the text, and the border.
    doPicture memDC
    doText memDC
    doBorder memDC, aRect

    If mGotFocus Then doFocusRect memDC, aRect
    
    ' copy the contents of our back buffer(memory device context) onto the control.
    BitBlt UserControl.hdc, 0, 0, aRect.right, aRect.bottom, memDC, 0, 0, vbSrcCopy
    ' clean up our memory resources
    DeleteDC memDC
    DeleteObject memBmp

End Sub
' ******************************************************************
'  Routines for painting the UserControl's various states.
' ******************************************************************
Private Sub doText(aDC As Long)

    Dim hFont           As Long         'Handle to the font object
    Dim hOrigFont       As Long         'Original font of the DC
    Dim x               As Long         'X coordinate for the text
    Dim Y               As Long         'Y coordinate for the text
    Dim tm              As TEXTMETRIC   'Type used to get the Height & Weight of the font
    Dim aPt             As POINTAPI     'Used to get text width
    Dim isPic           As Boolean      'Whether there is a picture or not
    Dim aPic            As StdPicture   'picture being used
    
    ' Calculate the X and Y Locations for our caption.
    ' Offset the X value by mImageOffset pixels from the picture's width
    Y = UserControl.ScaleY((UserControl.Height - UserControl.TextHeight(mCap)) / 2, UserControl.ScaleMode, vbPixels)
    
    ' Select the correct picture for dimensions
    If mState = 1 And Not (mHovPic Is Nothing) Then
        Set aPic = mHovPic
    Else
        If Not mPic Is Nothing Then
            Set aPic = mPic
        End If
    End If
    
    ' Determine the left location of the Caption text
    If mHAlign = hcaLeft Then
        If aPic Is Nothing Then
            x = mBorderOffset
        Else
            x = mBorderOffset + mImageOffset + UserControl.ScaleX(aPic.Width, vbHimetric, vbPixels)
        End If
    ElseIf mHAlign = hcaRight Then
        GetTextExtentPoint32 hdc, mCap, Len(mCap), aPt
        x = UserControl.ScaleX(UserControl.Width, UserControl.ScaleMode, vbPixels) - aPt.x - mBorderOffset
    ElseIf HAlign = hcaCenter Then
        GetTextExtentPoint32 hdc, mCap, Len(mCap), aPt
        If aPic Is Nothing Then
            x = ((UserControl.ScaleX(UserControl.Width, UserControl.ScaleMode, vbPixels) - (aPt.x)) / 2) + x
        Else
            x = ((UserControl.ScaleX(UserControl.Width, UserControl.ScaleMode, vbPixels) - (aPt.x - IIf(aPic Is Nothing, 0, mImageOffset + UserControl.ScaleX(aPic.Width, vbHimetric, vbPixels)))) / 2) + x
        End If
    End If
    
    ' If the button is in a depressed state, offset the caption by the
    '  number of pixels in the border (2 for thick, 1 for thin)
    If mState = 2 Then
        x = x + IIf((mStyle And hcsThinBorder) = hcsThinBorder, 1, 2)
        Y = Y + IIf((mStyle And hcsThinBorder) = hcsThinBorder, 1, 2)
    End If
    
    ' Get some of the font's values, namely height and weight
    GetTextMetrics UserControl.hdc, tm
    With UserControl.Font
        ' Create a font to use in the memory DC
        hFont = CreateFont(tm.tmHeight, _
                            0, 0, 0, _
                            tm.tmWeight, .Italic, _
                            .Underline, .Strikethrough, 0, 0, 16, _
                            0, 0, .Name)
    End With
    
    ' Put our font into the memory DC
    hOrigFont = SelectObject(aDC, hFont)
    ' Figure out which color to use when drawing the font
    If ((mStyle And hcsTextColor) = hcsTextColor) And mState = 1 Then
        SetTextColor aDC, TranslateColor(mHighlightColor)
    Else
        SetTextColor aDC, TranslateColor(mForeColor)
    End If
    
    ' Make our text print out transparent (only letters, no background)
    SetBkMode aDC, 1
    ' Print our caption on the memory DC at the correct location X,Y
    DrawStateText aDC, 0, 0, mCap, Len(mCap), x, Y, 0, 0, DST_PREFIXTEXT Or IIf(mEnabled, DSS_NORMAL, DSS_DISABLED)
    ' Replace the original font and delete our new one
    SelectObject aDC, hOrigFont
    DeleteObject hFont
End Sub
Private Sub doPicture(aDC As Long)
    Dim x               As Long         ' X location of the picture
    Dim Y               As Long         ' Y location of the picture
    Dim picDC           As Long         ' Device context to hold the pic
    Dim picOldBmp       As Long         ' original bmp in the picDC
    Dim wid             As Long         ' Width of the bitmap
    Dim hgt             As Long         ' Height of the bitmap
    Dim aPic            As StdPicture   ' The bitmap picture
    
    Dim transColor      As Long         ' Transparent color, from (0,0) of the pic
    Dim maskDC          As Long         ' DC of the mask image
    Dim maskBmp         As Long         ' Bitmap of the mask image
    Dim I               As Long         ' used for creating the mask image
    Dim j               As Long         ' same as above
    Dim tempDC          As Long         ' DC for "cleaned" pic image (no transColor)
    Dim tempBmp         As Long         ' Bitmap in the tempDC
    Dim aPt             As POINTAPI     ' Point for caption width
    Dim origBGCol       As Long
    
    ' Leave if we don't have a picture
    If mHovPic Is Nothing And mPic Is Nothing Then Exit Sub
    
    ' Select which picture to use. If we have a hover picture and
    '  we're in a hover state, then use that. Otherwise, use the
    '  standard picture
    If (mState = 1) And Not (mHovPic Is Nothing) Then
        Set aPic = mHovPic
    Else
        If mPic Is Nothing Then Exit Sub
        Set aPic = mPic
    End If
    
    ' Calculate the dimensions of the picture, as well as the the
    '  X and Y destination locations
    wid = UserControl.ScaleX(aPic.Width, vbHimetric, vbPixels)
    hgt = UserControl.ScaleY(aPic.Height, vbHimetric, vbPixels)
    Y = (UserControl.ScaleY(UserControl.Height, UserControl.ScaleMode, vbPixels) - hgt) / 2
    
    If mHAlign = hcaLeft Then
        x = mBorderOffset
    ElseIf mHAlign = hcaRight Then
        GetTextExtentPoint32 hdc, mCap, Len(mCap), aPt
        x = UserControl.ScaleX(UserControl.Width, UserControl.ScaleMode, vbPixels) - aPt.x - mBorderOffset - mImageOffset - wid
    ElseIf HAlign = hcaCenter Then
        GetTextExtentPoint32 hdc, mCap, Len(mCap), aPt
        x = (UserControl.ScaleX(UserControl.Width, UserControl.ScaleMode, vbPixels) - (wid + mImageOffset + aPt.x)) / 2
    End If
    
    ' Offset the picture's destination X and Y by the border's width if
    '  it is in a depressed state
    If mState = 2 Then
        x = x + IIf((mStyle And hcsThinBorder) = hcsThinBorder, 1, 2)
        Y = Y + IIf((mStyle And hcsThinBorder) = hcsThinBorder, 1, 2)
    End If
    
    If aPic.Type = vbPicTypeIcon Then
        If mEnabled Then
            DrawIcon aDC, x, Y, aPic.Handle
        Else
            DrawState aDC, 0, 0, aPic.Handle, 0, x, Y, 0, 0, DST_ICON Or DSS_DISABLED
        End If
        GoTo doPicture_exitSub
    End If
    
    ' Create a device context to hold and manipulate our picture
    picDC = CreateCompatibleDC(aDC)
    ' Pull the picture into our device context.
    picOldBmp = SelectObject(picDC, aPic.Handle)
    
    ' Create some other graphics resources:
    '  1) a DC and B&W Bitmap to use as our "Mask" image
    '  2) a DC and bitmap to store a "cleaned" version of the picture w/o
    '       the transparent color.
    maskDC = CreateCompatibleDC(aDC)
    maskBmp = CreateBitmap(wid, hgt, 1, 1, ByVal 0&)
    tempDC = CreateCompatibleDC(aDC)
    tempBmp = CreateCompatibleBitmap(aDC, wid, hgt)
    DeleteObject SelectObject(maskDC, maskBmp)
    DeleteObject SelectObject(tempDC, tempBmp)
    
    ' Get the mask color, or "transparent" color from pixel 0,0
    transColor = GetPixel(picDC, 0, 0)
    ' Generate the pixel mask. We're going to do it one pixel at a time
    '  rather than BitBlting. Its slower, but what do you expect for a
    '  $1.50 control?
    For j = 0 To hgt
        For I = 0 To wid
            SetPixel maskDC, I, j, IIf(GetPixel(picDC, I, j) = transColor, vbWhite, vbBlack)
        Next I
    Next j
    
    ' Create the image of our picture WITHOUT the transparent color
    BitBlt tempDC, 0, 0, wid, hgt, maskDC, 0, 0, vbSrcCopy
    BitBlt tempDC, 0, 0, wid, hgt, picDC, 0, 0, vbSrcPaint
    
        
    If mEnabled Then
        ' Punch the hole in our control's DC where we want to put the picture
        origBGCol = SetBkColor(aDC, vbWhite)
        SetTextColor aDC, vbBlack
        BitBlt aDC, x, Y, wid, hgt, maskDC, 0, 0, vbMergePaint
        ' Put our picture into that area
        BitBlt aDC, x, Y, wid, hgt, tempDC, 0, 0, vbSrcAnd
        ' Release the picture back to its normal state by selecting the
        '  picDC's original bitmap back
        SetBkColor aDC, origBGCol
        ' Clean up all of our graphics resources... MUY IMPORTANTE!!!
        SelectObject picDC, picOldBmp
    Else
        BitBlt picDC, 0, 0, wid, hgt, tempDC, 0, 0, vbSrcCopy
        SelectObject picDC, picOldBmp
        DeleteDC picDC
        DeleteObject picOldBmp
        origBGCol = SetBkColor(aDC, transColor)
        DrawState aDC, 0, 0, tempBmp, 0, x, Y, 0, 0, DST_BITMAP Or DSS_DISABLED
        SetBkColor aDC, origBGCol
    End If
    
    DeleteDC maskDC
    DeleteObject maskBmp
    DeleteDC tempDC
    DeleteObject tempBmp
    DeleteDC picDC
    DeleteObject picOldBmp

doPicture_exitSub:

End Sub
Private Sub doBorder(aDC As Long, ByRef aRect As RECT)
    If (mStyle And hcsNoBorder) = hcsNoBorder Then Exit Sub
    If ((mStyle And hcsPopupBorder) = hcsPopupBorder) And mState = 0 Then
        ' If our style has "popupBorder" and the cursor isn't on the control,
        '  then don't draw anything
    ElseIf (mState = 0 And (mStyle And hcsPopupBorder) = 0) Or _
           mState = 1 Then
        ' If we don't have the "PopupBorder" style and the cursor isn't on the
        '  control, OR if its in a hover state, draw a raised border of the
        '  appropriate thickness
        If (mStyle And hcsThinBorder) = hcsThinBorder Then
            DrawEdge aDC, aRect, BDR_RAISEDINNER, BF_RECT
        Else
            DrawEdge aDC, aRect, EDGE_RAISED, BF_RECT
        End If
    ElseIf mState = 2 Then
        ' If the button is in a depressed state, draw a sunken border of the
        '  appropriate thickness.
        If (mStyle And hcsThinBorder) = hcsThinBorder Then
            DrawEdge aDC, aRect, BDR_SUNKENINNER, BF_RECT
        Else
            DrawEdge aDC, aRect, EDGE_SUNKEN, BF_RECT
        End If
    End If
End Sub
Private Sub doFocusRect(aDC As Long, ByRef aRect As RECT)
    Dim aRect2 As RECT
    
    If ((mStyle And hcsPopupBorder) = hcsPopupBorder) And mState = 0 Then
        ' If our style has "popupBorder" and the cursor isn't on the control,
        '  then don't draw anything
        If (mStyle And hcsThinBorder) = hcsThinBorder Then
            DrawEdge aDC, aRect, BDR_RAISEDINNER, BF_RECT
        Else
            DrawEdge aDC, aRect, EDGE_RAISED, BF_RECT
        End If
    End If

    If mGotFocus And ((mStyle And hcsNoFocusRect) <> hcsNoFocusRect Or (mHovPic Is Nothing)) Then
        If (mStyle And hcsNoFocusRect) <> hcsNoFocusRect Then
            aRect2.left = aRect.left + 3
            aRect2.top = aRect.top + 3
            aRect2.bottom = aRect.bottom - 3
            aRect2.right = aRect.right - 3
        Else
            aRect2.left = aRect.left
            aRect2.right = aRect.right
            aRect2.top = aRect.top
            aRect2.bottom = aRect.bottom
        End If
        SetTextColor aDC, vbBlack
        DrawFocusRect aDC, aRect2
    
    End If

End Sub

Private Function TranslateColor(aColor As OLE_COLOR) As Long
    Dim newcolor As Long
    OleTranslateColor aColor, UserControl.Palette, newcolor
    TranslateColor = newcolor
End Function
