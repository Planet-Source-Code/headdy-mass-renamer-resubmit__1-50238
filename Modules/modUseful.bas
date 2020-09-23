Attribute VB_Name = "modUseful"
'---------------------------------------------------------------------------------------
' Module    : modUseful (Module)
' DateTime  : 18/11/2003
' Author    : Drew (aka The Bad One)
' Purpose   : A small repository of useful code.
'---------------------------------------------------------------------------------------

Option Explicit

Private Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, col As Long) As Long
Private Const CLR_INVALID = &HFFFF
Const Pi = 3.1415927
Public Const Author = "Drew (aka The Bad One)"

'---------------------------------------------------------------------------------------
' Procedure : Version
' DateTime  : 12/10/2003
' Author    : Drew (aka The Bad One)
' Purpose   : Returns a preformatted version string.
'---------------------------------------------------------------------------------------
'
Function Version$()
   Version = App.Major & "." & App.Minor & " build " & App.Revision
End Function

'---------------------------------------------------------------------------------------
' Procedure : TranslateColour
' DateTime  : 12/10/2003
' Author    : Drew (aka The Bad One)
' Purpose   : Used to convert Automation colours to a Windows (long) colour.
'---------------------------------------------------------------------------------------
'
Function TranslateColour(ByVal oClr As OLE_COLOR, Optional hPal As Long = 0) As Long
   If TranslateColor(oClr, hPal, TranslateColour) Then
       TranslateColour = CLR_INVALID
   End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : FileNameFromPath
' DateTime  : 12/10/2003 12:37
' Author    : Drew (aka The Bad One)
' Purpose   : Returns the filename from a full path, with extension.
'---------------------------------------------------------------------------------------
'
Function FileNameFromPath(ByVal Path As String) As String
   FileNameFromPath = Right(Path, Len(Path) - InStrRev(Path, "\"))
End Function

'---------------------------------------------------------------------------------------
' Procedure : FileBodyFromPath
' DateTime  : 12/10/2003 12:37
' Author    : Drew (aka The Bad One)
' Purpose   : Returns the file body of a filename from a full path.
'---------------------------------------------------------------------------------------
'
Function FileBodyFromPath(ByVal Path As String) As String
   Dim A As Integer
   If InStrRev(Path, "\") > 0 Then
      A = InStrRev(Path, "\")
      Path = Right(Path, Len(Path) - A)
   End If
   
   If InStrRev(Path, ".") > 0 Then
      A = InStrRev(Path, ".")
      FileBodyFromPath = VBA.Left$(Path, A - 1)
   Else
      FileBodyFromPath = Path
   End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : TrimFromEnd
' DateTime  : 12/10/2003
' Author    : Drew (aka The Bad One)
' Purpose   : Trims a specified charactere from the end of a string.
'---------------------------------------------------------------------------------------
'
Function TrimFromEnd(what As String, char As String) As String
   While Right(what, 1) = char
      what = VBA.Left$(what, Len(what) - 1)
   Wend
   TrimFromEnd = what
End Function

'---------------------------------------------------------------------------------------
' Procedure : RPad
' DateTime  : 12/10/2003
' Author    : Drew (aka The Bad One)
' Purpose   : Equivalent to RSet command - this function returns a string padded from
'             the left with the specified character.
'---------------------------------------------------------------------------------------
'
Function RPad$(Expression As String, Number As Long, Optional Character As String = " ")
   Dim L As Long
   L = Len(Expression)
   If L > Number Then
      ' String is too long, truncate it!
      RPad = Right(Expression, Number)
   Else
      RPad = String$(Number - L, Character) & Expression
   End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : modUseful.DegreesToRadians
' DateTime  : 08/11/2003
' Author    : Drew (aka The Bad One)
' Purpose   : Converts a value in degrees to radians, as used by Visual Basic.
' Assuming  : nothing
'---------------------------------------------------------------------------------------
'
Function DegreesToRadians(ByVal sngAngle As Single) As Single
   DegreesToRadians = sngAngle * (Pi / 180)
End Function

'---------------------------------------------------------------------------------------
' Procedure : modExif.IsJPEGFile
' DateTime  : 15/11/2003
' Author    : ?
' Purpose   : Looks for certain bytes in a file to determine if a file is a JPEG file.
' Assuming  : nothing
'---------------------------------------------------------------------------------------
'
Function IsJPEGFile(strFileName As String) As Boolean
   Dim First2Bytes   As String
   Dim Last2Bytes    As String
   Dim F As Long
   
   On Error GoTo IsJPEGFile_Error

   F = FreeFile
   Open strFileName For Binary Access Read As #F
   
   First2Bytes = Input(2, #F)
   Seek #F, FileLen(strFileName) - 1 ' Jump to near end
   Last2Bytes = Input(2, #F)
   IsJPEGFile = Not (Byte2Hex(First2Bytes) <> "FFD8" Or Byte2Hex(Last2Bytes) <> "FFD9")
   Close #F

   On Error GoTo 0
   Exit Function

IsJPEGFile_Error:

End Function

'---------------------------------------------------------------------------------------
' Procedure : modUseful.Byte2Hex
' DateTime  : 18/11/2003
' Author    : ?
' Purpose   : Converts byte values to hex.
' Assuming  : nothing
'---------------------------------------------------------------------------------------
'
Function Byte2Hex(InBytes As String) As String
   Dim i As Long
   Dim tmp As String
   
   For i = 1 To Len(InBytes)
      tmp = Hex(Asc(Mid$(InBytes, i, 1)))
      Byte2Hex = Byte2Hex & String$(2 - Len(tmp), "0") & tmp
   Next i
End Function
