Attribute VB_Name = "modMassRename"
Option Explicit


Global MR_SerialCounter As Long


' NewName()
' ------------------------------------
' This replaces the fields in strFormat with the revelant info,
' and returns only the FILENAME of the new image.

Function NewName(strOriginal As String, strFormat As Variant) As String
   Dim I As Integer
   For I = 0 To 5
      NewName = NewName + strFormat(I)
   Next
   
   NewName = Replace(NewName, "{file name}", LCase(FileBodyFromPath(strOriginal)))
   NewName = Replace(NewName, "{FILE NAME}", UCase(FileBodyFromPath(strOriginal)))
   NewName = Replace(NewName, "{File Name}", SentenceCase(FileBodyFromPath(strOriginal)))
   NewName = Replace(NewName, "{Date (dd)}", Format(Day(Now), "00"))
   NewName = Replace(NewName, "{Date (mm)}", Format(Month(Now), "00"))
   NewName = Replace(NewName, "{Date (yy)}", Format(Year(Now) Mod 100, "00"))
   NewName = Replace(NewName, "{Date (yyyy)}", Format(Year(Now), "0000"))
   NewName = Replace(NewName, "{Date (mmyy)}", Format(Month(Now), "00") & Format(Year(Now) Mod 100, "00"))
   NewName = Replace(NewName, "{Date (mmdd)}", Format(Month(Now), "00") & Format(Day(Now), "00"))
   NewName = Replace(NewName, "{Date (ddmm)}", Format(Day(Now), "00") & Format(Month(Now), "00"))
   NewName = Replace(NewName, "{Date (mmddyyyy)}", Format(Month(Now), "00") & Format(Day(Now), "00") & Format(Year(Now), "0000"))
   NewName = Replace(NewName, "{Date (mmddyy)}", Format(Month(Now), "00") & Format(Day(Now), "00") & Format(Year(Now) Mod 100, "00"))
   NewName = Replace(NewName, "{Date (ddmmyyyy)}", Format(Day(Now), "00") & Format(Month(Now), "00") & Format(Year(Now), "0000"))
   NewName = Replace(NewName, "{Date (ddmmyy)}", Format(Day(Now), "00") & Format(Month(Now), "00") & Format(Year(Now) Mod 100, "00"))
   NewName = Replace(NewName, "{1 digit serial number}", MR_SerialCounter Mod 10)
   NewName = Replace(NewName, "{2 digit serial number}", Format(MR_SerialCounter Mod 100, "00"))
   NewName = Replace(NewName, "{3 digit serial number}", Format(MR_SerialCounter Mod 1000, "000"))
   NewName = Replace(NewName, "{4 digit serial number}", Format(MR_SerialCounter Mod 10000, "0000"))
   NewName = Replace(NewName, "{Serial letter (a, b, c...)}", LCase(Chr$(vbKeyA + (MR_SerialCounter Mod 26))))
   NewName = Replace(NewName, "{Serial letter (A, B, C...)}", UCase(Chr$(vbKeyA + (MR_SerialCounter Mod 26))))

   ' extension - if at the end of the format string we replace it with a preceding dot.
   I = InStrRev(LCase(NewName), "{extension}", -1, vbTextCompare)
   If I > 0 And FileExtFromPath(strOriginal) <> Empty Then
      If Right(NewName, Len("{extension}")) = "{extension}" Then
         NewName = Left(NewName, I - 1) + "." & LCase(FileExtFromPath(strOriginal))
      End If
      If Right(NewName, Len("{extension}")) = "{EXTENSION}" Then
         NewName = Left(NewName, I - 1) + "." & UCase(FileExtFromPath(strOriginal))
      End If
   End If

   NewName = Replace(NewName, "{extension}", LCase(FileExtFromPath(strOriginal)))
   NewName = Replace(NewName, "{EXTENSION}", UCase(FileExtFromPath(strOriginal)))
   
   MR_SerialCounter = MR_SerialCounter + 1
End Function


Function FileExtFromPath(ByVal Path As String) As String
   Dim A As Integer
   If InStrRev(Path, ".") > 0 Then
      A = InStrRev(Path, ".")
      FileExtFromPath = VBA.Right$(Path, Len(Path) - A)
   Else
      FileExtFromPath = Empty
   End If
   FileExtFromPath = LCase$(FileExtFromPath)
End Function

' SentenceCase()
' -----------------------
' Converts strings into sentence case (each word beginning with
' a captial letter, and the other letters lower case).

Function SentenceCase(strWhat As String) As String
   Dim I As Integer
   If strWhat = Empty Then
      Exit Function
   End If
   Mid(strWhat, 1, 1) = UCase(Mid(strWhat, 1, 1))
   For I = 2 To Len(strWhat)
      If I = 1 Or Mid(strWhat, I - 1, 1) = " " Then
         Mid(strWhat, I, 1) = UCase(Mid(strWhat, I, 1))
      Else
         Mid(strWhat, I, 1) = LCase(Mid(strWhat, I, 1))
      End If
   Next
   SentenceCase = strWhat
End Function
