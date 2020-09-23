Attribute VB_Name = "Module1"
Option Explicit
    
' demo project showing how to create splitter bars in VB
' by Bryan Stafford of New Vision Software® - newvision@imt.net
' this demo is released into the public domain "as is" without
' warranty or guaranty of any kind.  In other words, use at
' your own risk.

'sets the width of the splitter bars
Const SPLT_WDTH As Integer = 4

'sets the horizontal & vertical
'offsets of the controls
Const CTRL_OFFSET As Integer = 5

' flag to indicate that a splitter recieved a mousedown
Dim fInitiateDrag As Boolean

' Point struct for ClientToScreen
Private Type POINTAPI
   X As Long
   Y As Long
End Type

' structure to pass to ClipCursor
Private Type RECT
   left As Long
   top As Long
   right As Long
   bottom As Long
End Type

' RECT structs to hold the area to contain the cursor
Dim CurVertRect As RECT
Dim CurHorzRect As RECT

Private Declare Sub ClipCursor Lib "user32" (lpRect As Any)
Private Declare Function LockWindowUpdate& Lib "user32" (ByVal hwndLock&)
Private Declare Function ClientToScreen& Lib "user32" (ByVal hwnd&, lpPoint As POINTAPI)
Private Declare Function SendMessage& Lib "user32" Alias "SendMessageA" (ByVal hwnd&, ByVal wMsg&, ByVal wParam&, lParam As Any)
Private Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hwnd&, ByVal nIndex&)
Private Declare Function SetWindowLong& Lib "user32" Alias "SetWindowLongA" (ByVal hwnd&, ByVal nIndex&, ByVal dwNewLong&)

Private Sub Form_Resize()

    ' minimized do nothing
    If WindowState = vbMinimized Then Exit Sub
    
    ' maximized, lock update to avoid nasty window flashing
    If WindowState = vbMaximized Then Call LockWindowUpdate(hwnd)
    
    Screen.MousePointer = vbHourglass
    
    ' handle minimum height. if you were to remove the
    ' controlbox you would need to handle minimum width also
    If Height < 1500 Then Height = 2000


    Dim FrameWidth As Integer
    ' the width of the window frame
    FrameWidth = ((Width \ Screen.TwipsPerPixelX) - ScaleWidth) \ 2

    ' handle a form resize that hides the vertical splitter
    If ((ScaleWidth - CTRL_OFFSET) - (Splitter(False).left + Splitter(False).Width)) < 12 Then
      Splitter(False).left = ScaleWidth - ((CTRL_OFFSET * 4) + (FrameWidth * 2))
    End If
    
    ' handle a form resize that hides the horizontal splitter
    If ((ScaleHeight - CTRL_OFFSET) - (Splitter(1).top + Splitter(1).Height)) < 12 Then
      Splitter(1).top = ScaleHeight - ((textHeight("A") + (FrameWidth * 2)) + (CTRL_OFFSET * 4))
    End If


   ' set vars for figuring the height of the controls
    Dim X2 As Integer
    Dim height1 As Integer
    Dim width1 As Integer
    Dim width2 As Integer
    
    height1 = ScaleHeight - (CTRL_OFFSET * 2)
    width1 = ListLeft.Width
    X2 = CTRL_OFFSET + width1 + SPLT_WDTH - 1
    width2 = ScaleWidth - X2 - CTRL_OFFSET
    
    ' move the left list
    ListLeft.Move CTRL_OFFSET - 1, CTRL_OFFSET, Splitter(False).left - Splitter(False).Width, height1

    ' handle the form height since IntegralHeight for the list is set to true.
    If WindowState = vbNormal Then Height = (((ListLeft.top + ListLeft.Height) + CTRL_OFFSET) _
             + (((Height \ Screen.TwipsPerPixelY) - ScaleHeight))) * Screen.TwipsPerPixelY
   
   
    ' resize the verticle splitter
    Splitter(False).Height = ListLeft.Height - 4
    
    ' move the right top textbox
    TextRight(False).Move (Splitter(False).left + Splitter(False).Width), CTRL_OFFSET, _
                   ((ScaleWidth - (Splitter(False).left + Splitter(False).Width)) _
                   - CTRL_OFFSET) + 2, Splitter(1).top - CTRL_OFFSET
                   
    ' check to see if this textbox needs scrollbar modifications
    CheckScrollBar TextRight(False)


    ' resize the horizontal splitter
    Splitter(1).Move (Splitter(False).left + Splitter(False).Width) + 2, _
                                        Splitter(1).top, TextRight(False).Width - 4


    ' move the right bottom textbox
    TextRight(1).Move TextRight(False).left, Splitter(1).top + Splitter(1).Height, _
                   TextRight(False).Width, (ListLeft.top + ListLeft.Height) - _
                   (Splitter(1).top + Splitter(1).Height)
                   
    ' check to see if this textbox needs scrollbar modifications
    CheckScrollBar TextRight(1)

                   
    ' sets the distance from the edge of the wall that the splitter may be moved.
    ' necessary because each of the controls have a minimum width and height.
    Const MIN_VERT_BUFFER As Integer = 20
    Const MIN_HORZ_BUFFER As Integer = 13
    ' the width of the cursor
    Const CURSOR_DEDUCT As Integer = 10
               
    ' set the ClipCursor vertical Rect
    With CurVertRect
      .left = GetScreenPoint(ListLeft.left + MIN_VERT_BUFFER, ListLeft.top, True)
      .top = GetScreenPoint(ListLeft.left, ListLeft.top + CURSOR_DEDUCT, False)
      .right = GetScreenPoint((TextRight(False).left + TextRight(False).Width) _
                                        - MIN_VERT_BUFFER, TextRight(False).top, True)
      .bottom = GetScreenPoint(ListLeft.left, (ListLeft.top + ListLeft.Height) _
                                        - CURSOR_DEDUCT, False)
    End With
    
    ' set the ClipCursor horizontal Rect
    With CurHorzRect
      .left = GetScreenPoint((Splitter(False).left + Splitter(False).Width) _
                                        + CURSOR_DEDUCT, ListLeft.top, True)
      .top = CurVertRect.top + MIN_HORZ_BUFFER
      .right = (CurVertRect.right + MIN_VERT_BUFFER) - CURSOR_DEDUCT
      .bottom = CurVertRect.bottom - MIN_HORZ_BUFFER
    End With
    
    Screen.MousePointer = vbDefault
    
    ' if it's locked unlock the window
    If WindowState = vbMaximized Then Call LockWindowUpdate(0&)

End Sub

Sub Splitter_ActivateDrag(state As Boolean)
   fInitiateDrag = state
End Sub

Sub Splitter_Move(X As Single, Y As Single)

  ' if the flag isn't set then the left button wasn't
  ' pressed while the mouse was over one of the splitters
  If fInitiateDrag <> True Then Exit Sub

  ' if the left button is down then we want to move the splitter
  If Button = 1 Then ' if the Tag is false then we need to set
    If Splitter(index).Tag = False Then ' the color and clip the cursor.
    
      Splitter(index).BackColor = &H808080 '<- set the "dragging" color here
      
      Select Case index
        Case False       ' figure out which Rect to use
          ClipCursor CurVertRect
          
        Case 1
          ClipCursor CurHorzRect
          
      End Select
      
      Splitter(index).Tag = True
    End If
    Select Case index
      Case False         ' move the appropriate splitter
        Splitter(index).left = (Splitter(index).left + X) - (SPLT_WDTH \ 3)
        
        ' For an interesting effect you can uncomment the next line.  You will
        ' also need to add code to change the color of both splitters when the
        ' vertical splitter is moved if you wish to implement this effect.
        'Splitter(index + 1).Move Splitter(index).left + Splitter(index).Width, _
                 Splitter(index + 1).top, ((ScaleWidth - (Splitter(index).left _
                 + Splitter(index).Width)) - CTRL_OFFSET) + 2
      Case 1
        Splitter(index).top = (Splitter(index).top + Y) - (SPLT_WDTH \ 3)
        
    End Select
  End If
  
End Sub

Private Sub Splitter_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  ' if the left button is the one being released we need to reset
  ' the color, Tag, flag, cancel ClipCursor and call form_resize
  If Button = 1 Then           ' to move the list and text boxes
    Splitter(index).Tag = False
    fInitiateDrag = False
    ClipCursor ByVal 0&
    Splitter(index).BackColor = &H8000000F  '<- set to original color
    Form_Resize
  End If
  
End Sub

Function GetScreenPoint(X As Long, Y As Long, bReturn As Boolean)
  ' this function calls ClientToScreen to convert the passed client point to
  ' a screen point and returns the x or y point depending on the value of bReturn
  
  Dim pt As POINTAPI
  
  ' plug the point into the point struct
  pt.X = X
  pt.Y = Y
  
  ' call for the conversion
  Call ClientToScreen(hwnd, pt)
  
  ' return the desired value
  If bReturn Then
    GetScreenPoint = pt.X
  Else
    GetScreenPoint = pt.Y
  End If
  
End Function
