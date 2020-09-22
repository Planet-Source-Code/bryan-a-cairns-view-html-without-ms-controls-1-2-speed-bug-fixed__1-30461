Attribute VB_Name = "mod_ViewStyle"
'Default Page Variables
Global lBGCOLOR As Long
Global lTextColor As Long
Global lLinkColor As Long
Global lActiveColor As Long
Global lVisitedColor As Long
Global sBGImage As String
Global sHTML As String
Global lFontSize As Long
Global sFontFace As String
Global sPageTitle As String
Global sDir As String
Global lAlign As Long

'Variables for Rendering the HTML
Global TCurrentX As Single
Global TCurrentY As Single
Global TOldX As Single
Global TOldY As Single

'Link Type Delcaration
Type HTMLLink
sLink As String 'where it links to
rBounds As RECT 'the x1,y1,x2,y2 area of the screen the link is on (in twips)
End Type
Global AllLinks() As HTMLLink
Public Sub SetDefaultColors(Optional bReset As Boolean)
lBGCOLOR = &HFFFFFF
lTextColor = &H0


lFontSize = 9
sFontFace = "Arial"
If bReset = False Then
sBGImage = ""
sPageTitle = ""
lLinkColor = &HFF0000
lActiveColor = &HFF8080
lVisitedColor = &H800000
End If
End Sub

Public Sub SetBold(Obj As PictureBox, bBold As Boolean)
Obj.FontBold = bBold
End Sub

Public Sub SetItalic(Obj As PictureBox, bItalic As Boolean)
Obj.FontItalic = bItalic
End Sub

Public Sub SetUnderLine(Obj As PictureBox, bUnderLine As Boolean)
Obj.FontUnderline = bUnderLine
End Sub

Public Sub SetFontColor(Obj As PictureBox, lColor As Long)
Obj.ForeColor = lColor
End Sub

Public Sub SetFontSize(Obj As PictureBox, lSize As Long)
Obj.FontSize = lSize
End Sub

Public Sub SetFontFace(Obj As PictureBox, sFont As String)
Obj.Font.Name = sFont
End Sub
