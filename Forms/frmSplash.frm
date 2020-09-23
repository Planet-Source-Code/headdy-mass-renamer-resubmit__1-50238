VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2220
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8250
   LinkTopic       =   "Form1"
   Picture         =   "frmSplash.frx":0000
   ScaleHeight     =   2220
   ScaleWidth      =   8250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrDisplay 
      Interval        =   1800
      Left            =   3720
      Top             =   480
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "version number"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   6255
      TabIndex        =   0
      Top             =   840
      Width           =   1080
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : bsMassRenamer.frmSplash (Form)
' DateTime  : 18/11/2003
' Author    : Drew (aka The Bad One)
' Purpose   : Splash screen.
'---------------------------------------------------------------------------------------

Option Explicit

Private Sub Form_Load()
   lblVersion.Caption = App.Major & "." & App.Minor
End Sub

Private Sub Form_Resize()
   Me.Width = ScaleX(Me.Picture.Width, vbHimetric, vbTwips)
   Me.Height = ScaleY(Me.Picture.Height, vbHimetric, vbTwips)
End Sub

Private Sub tmrDisplay_Timer()
   Unload Me
End Sub
