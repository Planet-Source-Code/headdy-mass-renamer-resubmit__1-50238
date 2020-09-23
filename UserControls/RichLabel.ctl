VERSION 5.00
Begin VB.UserControl RichLabel 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   HasDC           =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "RichLabel.ctx":0000
End
Attribute VB_Name = "RichLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum BorderStyleValues
    None = 0
    [Fixed Single] = 1
End Enum

Private ms_Caption As String
Private mb_WordWrap As Boolean
Private mb_AutoSize As Boolean

Private ml_Colour1 As Long
Private ml_Colour2 As Long
Private ml_Colour3 As Long



Public Event Click()

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_UserMemId = -501
       BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal lNewValue As OLE_COLOR)
       UserControl.BackColor = lNewValue
       DrawMe
End Property

Public Property Get AutoSize() As Boolean
       AutoSize = mb_AutoSize
End Property

Public Property Let AutoSize(ByVal bNewValue As Boolean)
       mb_AutoSize = bNewValue
       
       If bNewValue = True Then
            WordWrap = False
       End If
       DrawMe
        
End Property

Public Property Get Font() As IFontDisp
Attribute Font.VB_UserMemId = -512
       
       Set Font = UserControl.Font
       Font.Name = UserControl.Font.Name
       Font.Size = UserControl.Font.Size
       
End Property
Public Property Set Font(ByVal NewValue As IFontDisp)
       
       UserControl.Font.Name = NewValue.Name
       UserControl.Font.Size = NewValue.Size
       DrawMe
       
End Property

Public Property Get Colour3() As OLE_COLOR
       Colour3 = ml_Colour3
End Property

Public Property Let Colour3(ByVal lNewValue As OLE_COLOR)
       ml_Colour3 = lNewValue
       DrawMe
End Property

Public Property Get Colour2() As OLE_COLOR
       Colour2 = ml_Colour2
End Property

Public Property Let Colour2(ByVal lNewValue As OLE_COLOR)
       ml_Colour2 = lNewValue
       DrawMe
End Property

Public Property Get Colour1() As OLE_COLOR
       Colour1 = ml_Colour1
End Property

Public Property Let Colour1(ByVal lNewValue As OLE_COLOR)
       ml_Colour1 = lNewValue
       DrawMe
End Property


Public Property Get BorderStyle() As BorderStyleValues
Attribute BorderStyle.VB_UserMemId = -504
    
    BorderStyle = UserControl.BorderStyle
    
End Property

Public Property Let BorderStyle(ByVal iNewValue As BorderStyleValues)
    
    UserControl.BorderStyle = iNewValue
    DrawMe
End Property

Public Property Get WordWrap() As Boolean
       WordWrap = mb_WordWrap
End Property

Public Property Let WordWrap(ByVal bNewValue As Boolean)
       mb_WordWrap = bNewValue
       
       If bNewValue = True Then AutoSize = False
       DrawMe
       
End Property

Public Property Get Caption() As String
Attribute Caption.VB_UserMemId = -518
       
       Caption = ms_Caption
       
End Property

Public Property Let Caption(ByVal sNewValue As String)
       
       ms_Caption = sNewValue
       PropertyChanged "Caption"
       DrawMe

End Property
Private Sub DrawMe()
Dim s As String, i As Integer, c As String * 1

s = Caption
s = Replace(s, "<B>", Chr$(1), 1, -1, vbTextCompare)    'BOLD
s = Replace(s, "</B>", Chr$(2), 1, -1, vbTextCompare)
s = Replace(s, "<I>", Chr$(3), 1, -1, vbTextCompare)    'ITALIC
s = Replace(s, "</I>", Chr$(4), 1, -1, vbTextCompare)
s = Replace(s, "<U>", Chr$(5), 1, -1, vbTextCompare)    'UNDERLINE
s = Replace(s, "</U>", Chr$(6), 1, -1, vbTextCompare)

'COLORS: C1,C2 and C3
s = Replace(s, "<C1>", Chr$(15), 1, -1, vbTextCompare)
s = Replace(s, "<C2>", Chr$(16), 1, -1, vbTextCompare)
s = Replace(s, "<C3>", Chr$(17), 1, -1, vbTextCompare)

s = Replace(s, "</C>", Chr$(8), 1, -1, vbTextCompare)
s = Replace(s, "<lt>", "<", 1, -1, vbTextCompare)
s = Replace(s, "<gt>", ">", 1, -1, vbTextCompare)
s = Replace(s, "<BR>", vbCr, 1, -1, vbTextCompare)

'Font.Name = "verdana"
'Font.Size = 8

Cls

For i = 1 To Len(s)
    c = Mid(s, i, 1)
    
    If c = Chr$(1) Then
        UserControl.Font.Bold = True
    ElseIf c = Chr$(2) Then
        UserControl.Font.Bold = False
    
    ElseIf c = Chr$(3) Then
        UserControl.Font.Italic = True
    ElseIf c = Chr$(4) Then
        UserControl.Font.Italic = False
    
    ElseIf c = Chr$(5) Then
        UserControl.Font.Underline = True
    ElseIf c = Chr$(6) Then
        UserControl.Font.Underline = False
    
    ElseIf c = Chr$(15) Then
        UserControl.ForeColor = Colour1
    ElseIf c = Chr$(16) Then
        UserControl.ForeColor = Colour2
    ElseIf c = Chr$(17) Then
        UserControl.ForeColor = Colour3
    ElseIf c = Chr$(8) Then
        UserControl.ForeColor = vbBlack
    
    
    Else
        Print c;
        If WordWrap Then
            If (Width - CurrentX) < TextWidth("W") Then Print
        End If
        If AutoSize Then
            If (Width - CurrentX) < TextWidth("W") Then Width = Width + TextWidth("W")
        End If

    End If
Next


End Sub

Private Sub UserControl_Click()

RaiseEvent Click

End Sub

Private Sub UserControl_Initialize()
    
    Colour1 = vbRed
    Colour2 = vbGreen
    Colour3 = vbBlue

End Sub

Private Sub UserControl_InitProperties()

Caption = "<B>" & Extender.Name & "</B>"

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

Caption = PropBag.ReadProperty("Caption", Extender.Name)

BackColor = PropBag.ReadProperty("BackColor", Ambient.BackColor)
BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
Font.Name = PropBag.ReadProperty("Font.Name", Ambient.Font.Name)
Font.Size = PropBag.ReadProperty("Font.Size", Ambient.Font.Size)

Colour1 = PropBag.ReadProperty("Colour1", vbRed)
Colour2 = PropBag.ReadProperty("Colour2", vbGreen)
Colour3 = PropBag.ReadProperty("Colour3", vbBlue)

DrawMe

End Sub


Private Sub UserControl_Resize()

DrawMe

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

PropBag.WriteProperty "Caption", Caption, Extender.Name

PropBag.WriteProperty "BackColor", BackColor, Ambient.BackColor
PropBag.WriteProperty "BorderStyle", BorderStyle, 1

PropBag.WriteProperty "Colour1", Colour1, vbRed
PropBag.WriteProperty "Colour2", Colour2, vbGreen
PropBag.WriteProperty "Colour3", Colour3, vbBlue

PropBag.WriteProperty "Font.Name", Font.Name, Ambient.Font.Name
PropBag.WriteProperty "Font.size", Font.Size, Ambient.Font.Size

End Sub

