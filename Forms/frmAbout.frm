VERSION 5.00
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BackColor       =   &H009D774C&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5415
   ClientLeft      =   2340
   ClientTop       =   1650
   ClientWidth     =   7575
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3737.53
   ScaleMode       =   0  'User
   ScaleWidth      =   7113.316
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin bsMassRenamer.bsFrame fraCopyright 
      Height          =   975
      Left            =   120
      Top             =   4080
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   1720
      BackColour      =   10319692
      Caption         =   "bsFrame3"
      BorderStyle     =   2
      HighlightDKColour=   13160660
      CaptionColour   =   -2147483634
      BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FrameStyle      =   3
      Begin VB.CommandButton cmdOK 
         Cancel          =   -1  'True
         Caption         =   "&Okay"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6000
         TabIndex        =   4
         Top             =   240
         Width           =   1260
      End
      Begin VB.Label lblDisclaimer 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmAbout.frx":0000
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0E0FF&
         Height          =   705
         Left            =   120
         TabIndex        =   5
         Top             =   248
         Width           =   5775
      End
   End
   Begin VB.PictureBox picBG 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   6840
      Picture         =   "frmAbout.frx":012D
      ScaleHeight     =   660
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   3360
      Visible         =   0   'False
      Width           =   540
   End
   Begin bsMassRenamer.bsFrame bsFrame2 
      Height          =   975
      Left            =   120
      Top             =   2760
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   1720
      BackColour      =   10319692
      Caption         =   "Special Thanks"
      BorderStyle     =   2
      HighlightDKColour=   13160660
      BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FrameStyle      =   3
      Begin bsMassRenamer.RichLabel RichLabel1 
         Height          =   615
         Left            =   120
         TabIndex        =   3
         Top             =   248
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   1085
         Caption         =   "<C1>KPD Team<BR>Common Controls Replacement Project<BR>Author of ExifReader class"
         BorderStyle     =   0
         Colour1         =   16777215
      End
   End
   Begin bsMassRenamer.bsFrame bsFrame1 
      Height          =   615
      Left            =   120
      Top             =   1800
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   1085
      BackColour      =   10319692
      Caption         =   "created by"
      BorderStyle     =   2
      HighlightDKColour=   13160660
      BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FrameStyle      =   3
      Begin bsMassRenamer.RichLabel RichLabel2 
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   248
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   450
         Caption         =   "<C1><B>Drew</B> aka The Bad One    <B>drew</B>badsoft.co.uk    <B>www.</B>badsoft.co.uk"
         BorderStyle     =   0
         Colour1         =   16777215
      End
   End
   Begin VB.Image Image1 
      Height          =   1275
      Left            =   120
      Picture         =   "frmAbout.frx":0AEF
      Top             =   120
      Width           =   7350
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   3600
      TabIndex        =   0
      Top             =   1440
      Width           =   3765
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Sub cmdOK_Click()
   Unload Me
End Sub

Private Sub Form_Load()

   Dim X As Long, Y As Long
   ' Tile the background image
   For Y = 0 To ScaleHeight + picBG.ScaleHeight * 3 Step picBG.ScaleHeight
      For X = 0 To ScaleWidth + picBG.ScaleWidth Step picBG.ScaleWidth
         BitBlt Me.hDc, X / Screen.TwipsPerPixelX, Y / Screen.TwipsPerPixelY, _
            picBG.ScaleWidth, picBG.ScaleHeight, picBG.hDc, 0, 0, vbSrcCopy
      Next
   Next
   Refresh

   ' Fix information
   lblVersion.Caption = "Version " & App.Major & "." & App.Minor & App.Revision
   fraCopyright.Caption = App.LegalCopyright
End Sub

