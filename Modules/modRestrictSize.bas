Attribute VB_Name = "modRestrictSize"
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2003 VBnet, Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public DefWindowProc As Long
Public minX As Long
Public minY As Long
Public maxX As Long
Public maxY As Long

Public Const GWL_WNDPROC As Long = (-4)
Public Const WM_GETMINMAXINFO As Long = &H24

Public Type POINTAPI
    x As Long
    y As Long
End Type

Type MINMAXINFO
    ptReserved As POINTAPI
    ptMaxSize As POINTAPI
    ptMaxPosition As POINTAPI
    ptMinTrackSize As POINTAPI
    ptMaxTrackSize As POINTAPI
End Type

Public Declare Function SetWindowLong Lib "user32" _
    Alias "SetWindowLongA" _
   (ByVal hWnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long
   
Public Declare Function CallWindowProc Lib "user32" _
    Alias "CallWindowProcA" _
   (ByVal lpPrevWndFunc As Long, _
    ByVal hWnd As Long, _
    ByVal uMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long

Public Declare Sub CopyMemory Lib "kernel32" _
   Alias "RtlMoveMemory" _
   (hpvDest As Any, _
    hpvSource As Any, _
    ByVal cbCopy As Long)


Public Sub rsSubClass(hWnd As Long)

  'assign our own window message
  'procedure (WindowProc)
   On Error Resume Next
   DefWindowProc = SetWindowLong(hWnd, _
                                 GWL_WNDPROC, _
                                 AddressOf WindowProc)

End Sub


Public Sub rsUnSubClass(hWnd As Long)

  'restore the default message handling
  'before exiting
   If DefWindowProc Then
      SetWindowLong hWnd, GWL_WNDPROC, DefWindowProc
      DefWindowProc = 0
   End If

End Sub


Public Function WindowProc(ByVal hWnd As Long, _
                           ByVal uMsg As Long, _
                           ByVal wParam As Long, _
                           ByVal lParam As Long) As Long

  'window message procedure

   On Error Resume Next
  
   Select Case hWnd
   
     'If the handle returned is to our form,
     'perform form-specific message handling
     'to deal with the notifications. If it
     'is a general system message, pass it
     'on to the default window procedure.
      Case frmMain.hWnd
         
         On Error Resume Next
          
        'form-specific handler
         Select Case uMsg
            
            Case WM_GETMINMAXINFO
                 
                  Dim MMI As MINMAXINFO
                  
                  CopyMemory MMI, ByVal lParam, LenB(MMI)
      
                  'set the MINMAXINFO data to the
                  'minimum and maximum values set
                  'by the option choice
                   With MMI
                      .ptMinTrackSize.x = minX
                      .ptMinTrackSize.y = minY
                      .ptMaxTrackSize.x = maxX
                      .ptMaxTrackSize.y = maxY
                  End With
      
                  CopyMemory ByVal lParam, MMI, LenB(MMI)
                 
                 'the MSDN tells us that if we process
                 'the message, to return 0
                  WindowProc = 0
                    
              Case Else
              
                  'this takes care of all the other messages
                  'coming to the form and not specifically
                  'handled above.
                   WindowProc = CallWindowProc(DefWindowProc, _
                                               hWnd, _
                                               uMsg, _
                                               wParam, _
                                               lParam)
                  
          End Select
   
   End Select
   
End Function

