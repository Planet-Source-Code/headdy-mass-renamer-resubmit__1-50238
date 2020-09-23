Attribute VB_Name = "modMain"
Option Explicit

Public Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean

Public Const ICC_USEREX_CLASSES = &H200

Public Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type

Sub Main()

   Dim iccex As tagInitCommonControlsEx

   ' Initialise XP styles for Windows XP...
   On Error Resume Next
   With iccex
       .lngSize = LenB(iccex)
       .lngICC = ICC_USEREX_CLASSES
   End With
   InitCommonControlsEx iccex
   On Error GoTo 0
   
   
   ' Show the splash screen...
   Load frmSplash
   frmSplash.Show vbModal
   
   ' ...until the main form has been loaded.
   Load frmMain

   ' Then show the main form.
   Unload frmSplash
   frmMain.Show
End Sub
