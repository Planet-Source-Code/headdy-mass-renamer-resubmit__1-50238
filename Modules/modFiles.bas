Attribute VB_Name = "modFiles"
'---------------------------------------------------------------------------------------
' Module    : bsMassRenamer.modFiles (Module)
' DateTime  : 18/11/2003
' Author    : Drew (aka The Bad One)
' Purpose   : One of the backbones of Mass Renamer - this module handles many
'             of the file operations, including listing the files in a
'             ListView control.
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
' Updates
'---------------------------------------------------------------------------------------
'
'---------------------------------------------------------------------------------------

Option Explicit

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Const SHGFI_TYPENAME = &H400
Const MAX_PATH = 260
Const INVALID_HANDLE_VALUE = -1
Const FILE_ATTRIBUTE_DIRECTORY = &H10

' These are only some of the constants for retrieving special folders.
Private Const CSIDL_BITBUCKET = &HA
Private Const CSIDL_CONTROLS = &H3
Private Const CSIDL_DESKTOP = &H0
Private Const CSIDL_DESKTOPDIRECTORY = &H10
Private Const CSIDL_DRIVES = &H11
Private Const CSIDL_FAVORITES = &H6
Private Const CSIDL_FONTS = &H14
Private Const CSIDL_NETWORK = &H12
Private Const CSIDL_PERSONAL = &H5
Private Const CSIDL_PRINTERS = &H4
Private Const CSIDL_PROGRAMS = &H2
Private Const CSIDL_RECENT = &H8
Private Const CSIDL_SENDTO = &H9
Private Const CSIDL_STARTMENU = &HB
Private Const CSIDL_STARTUP = &H7
Private Const CSIDL_SYSTEM = &H25
Private Const CSIDL_TEMPLATES = &H15

Enum SpecialFolderConstants
   sfDesktopRoot = CSIDL_DESKTOP
   sfProgramFiles = CSIDL_PROGRAMS
   sfControlPanel = CSIDL_CONTROLS
   sfPrinters = CSIDL_PRINTERS
   sfMyDocuments = CSIDL_PERSONAL
   sfFavourites = CSIDL_FAVORITES
   sfStatupFolder = CSIDL_STARTUP
   sfRecentFiles = CSIDL_RECENT
   sfSendTo = CSIDL_SENDTO
   sfRecycleBin = CSIDL_BITBUCKET
   sfStartMenu = CSIDL_STARTMENU
   sfDesktopFolder = CSIDL_DESKTOPDIRECTORY
   sfMyComputer = CSIDL_DRIVES
   sfNetwork = CSIDL_NETWORK
   sfFonts = CSIDL_FONTS
   sfTemplates = CSIDL_TEMPLATES
   sfSystem = CSIDL_SYSTEM
End Enum

Private Type SHITEMID
    cb As Long
    abID As Byte
End Type

Private Type ITEMIDLIST
   mkid As SHITEMID
End Type


Private Type FILETIME
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
   dwFileAttributes As Long
   ftCreationTime As FILETIME
   ftLastAccessTime As FILETIME
   ftLastWriteTime As FILETIME
   nFileSizeHigh As Long
   nFileSizeLow As Long
   dwReserved0 As Long
   dwReserved1 As Long
   cFileName As String * MAX_PATH
   cAlternate As String * 14
End Type

Private Type SHFILEINFO
   hIcon As Long                       '  out: icon
   iIcon As Long                       '  out: icon index
   dwAttributes As Long                '  out: SFGAO_ flags
   szDisplayName As String * MAX_PATH  '  out: display name (or path)
   szTypeName As String * 80           '  out: type name
End Type

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

'---------------------------------------------------------------------------------------
' Procedure : modFiles.StripNulls
' DateTime  : 18/11/2003
' Author    : Drew (aka The Bad One)
' Purpose   : Removes null characters from a string.
' Assuming  : nothing
'---------------------------------------------------------------------------------------
'
Function StripNulls(OriginalStr As String) As String
    If (InStr(OriginalStr, Chr(0)) > 0) Then
        OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
    End If
    StripNulls = OriginalStr
End Function


'---------------------------------------------------------------------------------------
' Procedure : modFiles.ListFiles
' DateTime  : 17/11/2003
' Author    : Drew (aka The Bad One) (with help from the KPD Team)
' Purpose   : Lists the files within the specified directory, by placing the
'             details within a ListView control.
' Assuming  : Target ListView control already has the columns added.
'---------------------------------------------------------------------------------------
'
Sub ListFiles(strDirectory As String, lsvOutput As ListView)

   Dim lsiNew As ListItem
   
   lsvOutput.ListItems.Clear

   ' Correct the directory string (you don't know how many problems this
   ' causes)
   If Right(strDirectory, 1) <> "\" Then
      strDirectory = strDirectory & "\"
   End If

   Dim FileName As String ' Walking filename variable...
   Dim strFileName As String ' SubDirectory Name
   Dim dirNames() As String ' Buffer for directory name entries
   Dim nDir As Integer ' Number of directories in this path
   Dim i As Integer ' For-loop counter...
   Dim hSearch As Long ' Search Handle
   Dim WFD As WIN32_FIND_DATA
   Dim Cont As Boolean
   Dim exifHeader As ExifReader
   Dim FileCount As Integer
    
   ' Search for subdirectories, ignoring "." (current directory) and ".." (parent
   ' directory).
   nDir = 0
   Cont = True
   ReDim dirNames(nDir)
   
   ' Do a file count
   hSearch = FindFirstFile(strDirectory & "*", WFD)
   If hSearch <> INVALID_HANDLE_VALUE Then
      While Cont = True
         If (strFileName <> ".") And (strFileName <> "..") Then
            FileCount = FileCount + 1
         End If
         Cont = FindNextFile(hSearch, WFD)
      Wend
   End If
   Cont = FindClose(hSearch)
   
   If FileCount = 0 Then
      Exit Sub
   Else
      With frmMain.prgListFiles
         .Min = 0
         .Max = FileCount
         .Value = 0
         .Visible = True
      End With
   End If
   
   ' now list the files.
   Cont = True
   hSearch = FindFirstFile(strDirectory & "*", WFD)
   LockWindowUpdate frmMain.lsvFiles.hWnd
   
   On Error Resume Next
   
   If hSearch <> INVALID_HANDLE_VALUE Then
      Do While Cont
         strFileName = StripNulls(WFD.cFileName)
         If (strFileName <> ".") And (strFileName <> "..") Then
         
            Set lsiNew = lsvOutput.ListItems.Add(Text:=strFileName)
            
            If Not (GetFileAttributes(strFileName) And FILE_ATTRIBUTE_DIRECTORY) Then
               ' File extension - folders don't have them
               lsiNew.SubItems(1) = LCase(FileExtFromPath(strFileName))
            End If
            
            ' Type name of file
            lsiNew.SubItems(2) = GetFileTypeName(strDirectory & strFileName)
            
            ' Modified date
            lsiNew.SubItems(3) = GetLocalFileTime(WFD.ftLastWriteTime)
            
            ' Creation date
            lsiNew.SubItems(4) = GetLocalFileTime(WFD.ftCreationTime)
            
            ' Image taken on date, for JPEG images
            If IsJPEGFile(strDirectory & strFileName) Then
               Set exifHeader = New ExifReader
               exifHeader.Load strDirectory & strFileName
               lsiNew.SubItems(5) = exifHeader.Tag(&H9003&) ' slight problem, so used raw value
               Set exifHeader = Nothing
            End If
            
            frmMain.prgListFiles.Value = frmMain.prgListFiles.Value + 1
         End If
         Cont = FindNextFile(hSearch, WFD)
      Loop
      Cont = FindClose(hSearch)
   End If
   
   On Error GoTo 0
   
   frmMain.prgListFiles.Visible = False
   LockWindowUpdate False
End Sub


'---------------------------------------------------------------------------------------
' Procedure : modFiles.GetFileTypeName
' DateTime  : 18/11/2003
' Author    : Drew (aka The Bad One)
' Purpose   : Returns the type name of the specified file (eg. "Word Document").
' Assuming  : A full path is needed, or this function won't work.
'---------------------------------------------------------------------------------------
'
Function GetFileTypeName(strFullName As String) As String
   Dim FI As SHFILEINFO
   SHGetFileInfo strFullName, 0, FI, Len(FI), SHGFI_TYPENAME
   GetFileTypeName = StripNulls(FI.szTypeName)
End Function

'---------------------------------------------------------------------------------------
' Procedure : modFiles.GetLocalFileTime
' DateTime  : 18/11/2003
' Author    : Drew (aka The Bad One)
' Purpose   : Converts a FILETIME structure into a readable date and time string.
' Assuming  : nothing
'---------------------------------------------------------------------------------------
'
Function GetLocalFileTime(ftFileTime As FILETIME)
   Dim ftConvert As FILETIME, stSystemTime As SYSTEMTIME
   
   FileTimeToLocalFileTime ftFileTime, ftConvert
   FileTimeToSystemTime ftConvert, stSystemTime
   
   GetLocalFileTime = Format$(stSystemTime.wDay, "00") + "/" + _
      Format$(stSystemTime.wMonth, "00") + "/" + Format$(stSystemTime.wYear, "00") + _
      " " + Format$(stSystemTime.wHour, "00") + ":" + Format$(stSystemTime.wMinute, "00")
End Function

'---------------------------------------------------------------------------------------
' Procedure : modFiles.GetSpecialFolderPath
' DateTime  : 18/11/2003
' Author    : Drew (aka The Bad One) w/KPD Team's help
' Purpose   : Returns the directory of a special folder.
' Assuming  : nothing
'---------------------------------------------------------------------------------------
'
Function GetSpecialFolderPath(sfFolder As SpecialFolderConstants)
   Dim IDL As ITEMIDLIST
   Dim result As Long
   Dim strPath As String
   
   result = SHGetSpecialFolderLocation(100, sfFolder, IDL)
   If Not result Then
      strPath = Space$(512)
      result = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal strPath)
      GetSpecialFolderPath = StripNulls(strPath)
   End If
End Function
