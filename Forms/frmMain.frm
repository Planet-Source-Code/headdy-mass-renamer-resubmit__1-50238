VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{19B7F2A2-1610-11D3-BF30-1AF820524153}#1.2#0"; "ccrpftv6.ocx"
Begin VB.Form frmMain 
   Caption         =   "Mass Renamer"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9255
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   9255
   StartUpPosition =   2  'CenterScreen
   Begin CCRPFolderTV6.FolderTreeview FolderView 
      Height          =   3420
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   6033
      Indent          =   20
      VirtualFolders  =   0   'False
   End
   Begin bsMassRenamer.bsLightContainer lctFiles 
      Height          =   2775
      Left            =   2880
      TabIndex        =   1
      Top             =   120
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   4895
      ScaleMode       =   0
      ScaleWidth      =   2775
      ScaleHeight     =   2775
      Begin MSComctlLib.Toolbar tbrSelect 
         Height          =   330
         Left            =   0
         TabIndex        =   3
         Top             =   1680
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "All"
               Object.ToolTipText     =   "Select All"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Invert"
               Object.ToolTipText     =   "Invert selection"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "None"
               Object.ToolTipText     =   "Select None"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lsvFiles 
         Height          =   1695
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   2990
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "filename"
            Text            =   "Filename"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Key             =   "extension"
            Text            =   "ext."
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Key             =   "filetype"
            Text            =   "Type"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Key             =   "datemodify"
            Text            =   "Modified date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Key             =   "datecreate"
            Text            =   "Creation date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Key             =   "dateimagetaken"
            Text            =   "Image taken on"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.PictureBox picSplitter 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      HasDC           =   0   'False
      Height          =   3135
      Left            =   2880
      MousePointer    =   9  'Size W E
      ScaleHeight     =   3135
      ScaleWidth      =   135
      TabIndex        =   4
      Top             =   120
      Width           =   135
   End
   Begin MSComctlLib.ImageList imlToolbar 
      Left            =   6240
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1CFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1E54
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1FAE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin bsMassRenamer.bsLightContainer lctOptions 
      Height          =   1935
      Left            =   0
      TabIndex        =   5
      Top             =   3720
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   3413
      ScaleMode       =   0
      ScaleWidth      =   311.808
      ScaleHeight     =   294.857
      Begin VB.CommandButton cmdAbout 
         Caption         =   "About..."
         Height          =   375
         Left            =   7080
         TabIndex        =   30
         Top             =   1440
         Width           =   2055
      End
      Begin MSComctlLib.ProgressBar prgListFiles 
         Height          =   255
         Left            =   7080
         TabIndex        =   7
         Top             =   360
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.CheckBox chkTestMode 
         Caption         =   "Test Mode"
         Height          =   315
         Left            =   7080
         TabIndex        =   28
         Top             =   600
         Width           =   2055
      End
      Begin VB.CommandButton cmdWork 
         Caption         =   "Work!"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7080
         TabIndex        =   29
         Top             =   960
         Width           =   2055
      End
      Begin bsMassRenamer.bsLightContainer lctRenameOptions 
         Height          =   975
         Index           =   1
         Left            =   240
         TabIndex        =   22
         Top             =   480
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   1720
         ScaleMode       =   0
         ScaleWidth      =   113.882
         ScaleHeight     =   0.002
         Begin VB.CheckBox chkRCExtension 
            Caption         =   "Include file extension"
            Height          =   195
            Left            =   0
            TabIndex        =   27
            Top             =   600
            Width           =   4815
         End
         Begin VB.TextBox txtRCReplaceTo 
            Height          =   285
            Left            =   2640
            TabIndex        =   26
            Text            =   " "
            Top             =   240
            Width           =   2295
         End
         Begin VB.TextBox txtRCReplaceFrom 
            Height          =   285
            Left            =   0
            TabIndex        =   25
            Text            =   "_"
            Top             =   240
            Width           =   2295
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "with character(s)"
            Height          =   195
            Left            =   2640
            TabIndex        =   24
            Top             =   0
            Width           =   1230
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "replace character(s)"
            Height          =   195
            Left            =   0
            TabIndex        =   23
            Top             =   0
            Width           =   1455
         End
      End
      Begin bsMassRenamer.bsLightContainer lctRenameOptions 
         Height          =   1380
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   2434
         ScaleMode       =   0
         ScaleWidth      =   441
         ScaleHeight     =   92
         Begin VB.TextBox txtSerialStart 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2057
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   2400
            TabIndex        =   21
            Text            =   "0"
            Top             =   1080
            Width           =   615
         End
         Begin VB.ComboBox cboRenameFormat 
            Height          =   315
            Index           =   0
            Left            =   0
            TabIndex        =   9
            Tag             =   "Rename"
            Text            =   "{file name}"
            Top             =   0
            Width           =   3015
         End
         Begin VB.ComboBox cboRenameFormat 
            Height          =   315
            Index           =   1
            Left            =   3360
            TabIndex        =   11
            Tag             =   "Rename"
            Text            =   "{extension}"
            Top             =   0
            Width           =   3015
         End
         Begin VB.ComboBox cboRenameFormat 
            Height          =   315
            Index           =   5
            Left            =   3360
            TabIndex        =   19
            Tag             =   "Rename"
            Top             =   720
            Width           =   3015
         End
         Begin VB.ComboBox cboRenameFormat 
            Height          =   315
            Index           =   4
            Left            =   0
            TabIndex        =   17
            Tag             =   "Rename"
            Top             =   720
            Width           =   3015
         End
         Begin VB.ComboBox cboRenameFormat 
            Height          =   315
            Index           =   3
            Left            =   3360
            TabIndex        =   15
            Tag             =   "Rename"
            Top             =   360
            Width           =   3015
         End
         Begin VB.ComboBox cboRenameFormat 
            Height          =   315
            Index           =   2
            Left            =   0
            TabIndex        =   13
            Tag             =   "Rename"
            Top             =   360
            Width           =   3015
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Serial number starts at"
            Height          =   195
            Left            =   0
            TabIndex        =   20
            Top             =   1125
            Width           =   1635
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "+"
            Height          =   195
            Index           =   3
            Left            =   6480
            TabIndex        =   16
            Top             =   420
            Width           =   120
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "+"
            Height          =   195
            Index           =   0
            Left            =   3120
            TabIndex        =   10
            Top             =   60
            Width           =   120
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "+"
            Height          =   195
            Index           =   4
            Left            =   3120
            TabIndex        =   18
            Top             =   780
            Width           =   120
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "+"
            Height          =   195
            Index           =   2
            Left            =   3120
            TabIndex        =   14
            Top             =   420
            Width           =   120
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "+"
            Height          =   195
            Index           =   1
            Left            =   6480
            TabIndex        =   12
            Top             =   60
            Width           =   120
         End
      End
      Begin MSComctlLib.TabStrip tbsRenameOptions 
         Height          =   1935
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   3413
         TabWidthStyle   =   2
         MultiRow        =   -1  'True
         TabFixedWidth   =   3528
         HotTracking     =   -1  'True
         TabStyle        =   1
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   2
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Rename format"
               Key             =   "Format"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Replace characters"
               Key             =   "Chars"
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : bsMassRenamer.frmMain (Form)
' DateTime  : 18/11/2003
' Author    : Drew (aka The Bad One)
' Purpose   : The main form for Mass Renamer.
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
' Module    : frmMain (Form)
' DateTime  : 16/11/2003
' Author    : Drew (aka The Bad One)
'---------------------------------------------------------------------------------------

Option Explicit

Private Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
Const UNIT_TWIPS = 120

Private WithEvents mSplit As cSplitter
Attribute mSplit.VB_VarHelpID = -1

Private Sub cmdAbout_Click()
   frmAbout.Show vbModal, Me
End Sub

'---------------------------------------------------------------------------------------
' Procedure : frmMain.cmdWork_Click
' DateTime  : 18/11/2003
' Author    : Drew (aka The Bad One)
' Purpose   : Starts the renaming of selected files.
' Assuming  : nothing
'---------------------------------------------------------------------------------------
'
Private Sub cmdWork_Click()
   Dim i As Integer, c As Integer
   
   For i = 1 To lsvFiles.ListItems.Count
      If lsvFiles.ListItems(1).Checked Then
         c = c + 1
      End If
   Next
   
   If c = 0 Then
      MsgBox "No files have been selected!" & vbCrLf & _
         "You must select at least one file to rename."
      Exit Sub
   End If
   
   Select Case tbsRenameOptions.SelectedItem.Key
      Case "Format"
         MassRenameFormat
      Case "Chars"
         MassRenameChars
   End Select
   
   If chkTestMode.Value Then
      frmTestMode.ResizeColumns
   Else
      If FolderView.AreSameFolders(FolderView.SelectedFolder, FolderView.RootFolder) Then
         ListFiles GetSpecialFolderPath(sfDesktopFolder), lsvFiles
      Else
         ListFiles FolderView.SelectedFolder, lsvFiles
      End If
   End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : frmMain.MassRenameFormat
' DateTime  : 18/11/2003
' Author    : Drew (aka The Bad One)
' Purpose   : Renames all the selected files, based on Mass Renamer fields. The results
'             are then used to rename the existing files, or displayed as items.
' Assuming  : modMassRenamer is present
'---------------------------------------------------------------------------------------
'
Private Sub MassRenameFormat()
   Dim i As Long
   Dim lsiTestItem As ListItem
   Dim strFileName As String, strPath As String
   
   If FolderView.AreSameFolders(FolderView.SelectedFolder, FolderView.RootFolder) Then
      strPath = GetSpecialFolderPath(sfDesktopFolder)
   Else
      strPath = FolderView.SelectedFolder
   End If
   
   If Right(strPath, 1) <> "\" Then
      strPath = strPath & "\"
   End If
   
   If chkTestMode.Value = 1 Then
      frmTestMode.Show
      frmTestMode.ListView1.ListItems.Clear
   End If
   
   MR_SerialCounter = Val(txtSerialStart.Text)
   
   With lsvFiles.ListItems
      For i = 1 To .Count
         If .Item(i).Checked = True Then
            strFileName = lsvFiles.ListItems(i).Text
            If chkTestMode.Value = 1 Then
               ' Just list the file
               Set lsiTestItem = frmTestMode.ListView1.ListItems.Add(Text:=strFileName)
               lsiTestItem.SubItems(1) = NewName(strFileName, cboRenameFormat)
            Else
               ' Really do it! CAUTION
               MoveFile strPath & strFileName, strPath & NewName(strFileName, cboRenameFormat)
            End If
         End If
      Next
   End With
End Sub

'---------------------------------------------------------------------------------------
' Procedure : frmMain.FolderView_FolderClick
' DateTime  : 18/11/2003
' Author    : Drew (aka The Bad One)
' Purpose   : When an item is clicked on with this control, the files and folders
'             for the path are displayed.
' Assuming  : nothing
'---------------------------------------------------------------------------------------
'
Private Sub FolderView_FolderClick(Folder As CCRPFolderTV6.Folder, Location As CCRPFolderTV6.ftvHitTestConstants)
   ' We need to watch out for the desktop folder, because by default it returns the
   ' root directory (which is nothing).
   If FolderView.AreSameFolders(FolderView.SelectedFolder, FolderView.RootFolder) Then
      ListFiles GetSpecialFolderPath(sfDesktopFolder), lsvFiles
   Else
      ListFiles Folder.FullPath, lsvFiles
   End If
End Sub

Private Sub Form_Load()
   Dim i As Integer
   
   ' Initialise form
   ' -----------------------------------------------------------------------
   lctFiles.ScaleMode = vbTwips
   lctOptions.ScaleMode = vbTwips
   
   ' Initialise FolderView
   ' -----------------------------------------------------------------------
   
   tbsRenameOptions_Click
    
   
   ' Form resizing restrictions
   ' -----------------------------------------------------------------------
   minX = (lctOptions.Width + UNIT_TWIPS * 4) \ Screen.TwipsPerPixelX
   minY = Me.Height \ Screen.TwipsPerPixelY
   maxX = Screen.Width \ Screen.TwipsPerPixelX
   maxY = Screen.Height \ Screen.TwipsPerPixelY
   Me.Width = minX * Screen.TwipsPerPixelX
      
   ' Initialise splitter control
   ' -----------------------------------------------------------------------
   Set mSplit = New cSplitter
   With mSplit
      .Initialise picSplitter, Me
      .BorderSize = picSplitter.Left
      .Orientation = cSPLTOrientationVertical
   End With
   picSplitter.BackColor = vbButtonFace
   picSplitter.Width = UNIT_TWIPS
   
   ' Initialise toolbar
   ' -----------------------------------------------------------------------
   With tbrSelect
      .ImageList = imlToolbar
      For i = 1 To 3
         .Buttons(i).Image = i
      Next
   End With
   
   ' Initialise combo boxes
   ' -----------------------------------------------------------------------
   FillCombos
   
   
   mSplit_SplitComplete
   Call rsSubClass(Me.hWnd)
End Sub

'---------------------------------------------------------------------------------------
' Procedure : frmMain.FillCombos
' DateTime  : 18/11/2003
' Author    : Drew (aka The Bad One)
' Purpose   : Fills the relevant combo boxes with Mass Renamer fields.
' Assuming  : nothing
'---------------------------------------------------------------------------------------
'
Private Sub FillCombos()
   Dim A As Integer
    
   For A = 0 To 5
      With cboRenameFormat(A)
         .AddItem ""
         .AddItem "{file name}"
         .AddItem "{File Name}"
         .AddItem "{FILE NAME}"
         .AddItem "{extension}"
         .AddItem "{EXTENSION}"
         .AddItem "{1 digit serial number}"
         .AddItem "{2 digit serial number}"
         .AddItem "{3 digit serial number}"
         .AddItem "{4 digit serial number}"
         .AddItem "{Serial letter (a, b, c...)}"
         .AddItem "{Serial letter (A, B, C...)}"
         .AddItem "{Date (ddmmyy)}"
         .AddItem "{Date (ddmmyyyy)}"
         .AddItem "{Date (mmddyy)}"
         .AddItem "{Date (mmddyyyy)}"
         .AddItem "{Date (ddmm)}"
         .AddItem "{Date (mmdd)}"
         .AddItem "{Date (mmyy)}"
         .AddItem "{Date (dd)}"
         .AddItem "{Date (mm)}"
         .AddItem "{Date (yy)}"
         .AddItem "{Date (yyyy)}"
      End With
   Next
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   mSplit.MouseMove x
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   mSplit.MouseUp x
   FolderView.Width = x - FolderView.Left
End Sub

Private Sub Form_Resize()
   ' The splitter control will size itself relative to the form automatically.
   If Me.WindowState <> vbMinimized Then
      FolderView.Height = ScaleHeight - lctOptions.Height - UNIT_TWIPS * 3
      picSplitter.Height = FolderView.Height
      lctFiles.Move picSplitter.Left + picSplitter.Width, UNIT_TWIPS, _
         ScaleWidth - picSplitter.Left - UNIT_TWIPS, picSplitter.Height
           
      lctOptions.Move UNIT_TWIPS, _
         Me.ScaleHeight - lctOptions.Height - UNIT_TWIPS, _
         Me.ScaleWidth
   End If
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call rsUnSubClass(Me.hWnd)                              ' VERY important
   Set mSplit = Nothing
   Unload frmTestMode
   Unload Me
End Sub

Private Sub lsvFiles_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   With lsvFiles
      ' Toggle the clicked column's sort order only if the active colum is clicked
      ' (iow, don't reverse the sort order when different columns are clicked).
      If (.SortKey = ColumnHeader.Index - 1) Then
         ColumnHeader.Tag = Not Val(ColumnHeader.Tag)
      End If
      
      ' Set sort order to that of the respective SortOrderConstants value
      .SortOrder = Abs(Val(ColumnHeader.Tag))
      
      ' Get the zero-based index of the clicked column (ColumnHeader.Index is one-based).
      .SortKey = ColumnHeader.Index - 1
   End With
   
End Sub

Private Sub mSplit_SplitComplete()
   FolderView.Width = picSplitter.Left - FolderView.Left
   lctFiles.Move picSplitter.Left + picSplitter.Width, UNIT_TWIPS, _
      ScaleWidth - picSplitter.Left - UNIT_TWIPS
End Sub


Private Sub lctFiles_Resize()
   lsvFiles.Move 0, 0, lctFiles.ScaleWidth, lctFiles.ScaleHeight - tbrSelect.Height
   tbrSelect.Move 0, lctFiles.ScaleHeight - tbrSelect.Height
End Sub

Private Sub picSplitter_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   mSplit.MouseDown x ' Use Y if the splitter is horizontal
End Sub

Private Sub tbrSelect_ButtonClick(ByVal Button As MSComctlLib.Button)
   Dim i As Integer
   
   For i = 1 To lsvFiles.ListItems.Count
      Select Case Button.Key
         Case "All"
            lsvFiles.ListItems(i).Checked = True
         Case "None"
            lsvFiles.ListItems(i).Checked = False
         Case "Invert"
            lsvFiles.ListItems(i).Checked = Not lsvFiles.ListItems(i).Checked
      End Select
   Next
End Sub

Private Sub tbsRenameOptions_Click()
   Select Case tbsRenameOptions.SelectedItem.Key
      Case "Format"
         lctRenameOptions(0).Visible = True
         lctRenameOptions(1).Visible = False
      Case "Chars"
         lctRenameOptions(0).Visible = False
         lctRenameOptions(1).Visible = True
   End Select
End Sub

'---------------------------------------------------------------------------------------
' Procedure : frmMain.MassRenameChars
' DateTime  : 06/11/2003
' Author    : Drew (aka The Bad One)
' Purpose   : Replaces characters in all the selected filenames in lsvFiles, and either
'             displays them as results or renames the files.
' Assuming  : nothing
'---------------------------------------------------------------------------------------
'
Private Sub MassRenameChars()
   Dim strPath As String
   Dim i As Integer
   Dim lsiTestItem As ListItem
   
   If FolderView.AreSameFolders(FolderView.SelectedFolder, FolderView.RootFolder) Then
      strPath = GetSpecialFolderPath(sfDesktopFolder)
   Else
      strPath = FolderView.SelectedFolder
   End If
   
   If Right(strPath, 1) <> "\" Then
      strPath = strPath + "\"
   End If

   If chkTestMode.Value = 1 Then
      frmTestMode.Show
      frmTestMode.ListView1.ListItems.Clear
   End If

   With lsvFiles.ListItems
      For i = 1 To .Count
         If .Item(i).Checked Then
            If chkTestMode.Value = 1 Then
               ' Just a test
               Set lsiTestItem = frmTestMode.ListView1.ListItems.Add(Text:=.Item(i).Text)
               lsiTestItem.SubItems(1) = ModifiedName_Char(.Item(i).Text)
            Else
               ' Do it for real! CAUTION
               MoveFile strPath & .Item(i).Text, strPath & ModifiedName_Char(.Item(i).Text)
            End If
         End If
      Next
   End With
   
End Sub


'---------------------------------------------------------------------------------------
' Procedure : frmMain.ModifiedName_Char
' DateTime  : 18/11/2003
' Author    : Drew (aka The Bad One)
' Purpose   : Returns the name of the file once the characters have been replaced.
' Assuming  : nothing
'---------------------------------------------------------------------------------------
'
Function ModifiedName_Char(strFileName As String) As String
   Dim strFileNameModify As String
   Dim intExtStart As Integer
   
   If chkRCExtension.Value = 1 Then
      ' Replace characters all over, including file extension
      strFileNameModify = Replace(strFileName, txtRCReplaceFrom.Text, _
         txtRCReplaceTo.Text)
   Else
      ' Don't replace characters in file extension
      intExtStart = InStrRev(strFileName, ".")
      If intExtStart Then
         strFileNameModify = Replace(FileBodyFromPath(strFileName), _
            txtRCReplaceFrom.Text, txtRCReplaceTo.Text) & "." & FileExtFromPath(strFileName)
      Else
         strFileNameModify = Replace(strFileName, txtRCReplaceFrom.Text, _
            txtRCReplaceTo.Text)
      End If
   End If
   
   ModifiedName_Char = strFileNameModify
End Function
