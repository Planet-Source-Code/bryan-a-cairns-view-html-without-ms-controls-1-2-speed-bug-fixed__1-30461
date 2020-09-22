Attribute VB_Name = "mod_HTMLRender"
Dim isInLink As Boolean
Dim OldColor As Long
Dim bUnderLinned As Boolean
Dim LLastLink As Long
Public Sub DisplayWebPage(Obj As PictureBox)
'This is where we parse the HTML and render the web page
Dim Ipos As Single
Dim Epos As Single
Dim sTAG As String
Dim sTXT As String
Dim rc As RECT
Dim lWidth As Long
Dim lHeight As Long

'Clear our links array
ReDim AllLinks(0)

sHTML = Replace(sHTML, vbCrLf, "")


Form1.Label1.Move ViewMargin, ViewMargin

TOldY = ViewMargin
TOldX = ViewMargin

rc.Top = TOldY
rc.Left = TOldX
rc.Bottom = (Obj.Height / Screen.TwipsPerPixelY) - rc.Top
rc.Right = (Obj.Width / Screen.TwipsPerPixelX) - rc.Left

Obj.Visible = False
Obj.Cls
Epos = 1
If sBGImage <> "" Then
TileBitmap Obj, Form1.PicImage
TileBitmap Form1.PicBacking, Form1.PicImage
End If

Ipos = InStr(1, sHTML, "<BODY")
If Ipos = 0 Then Exit Sub
Epos = InStr(Ipos, sHTML, ">")
If Epos = 0 Then Epos = 1
Do
DoEvents
'Get the start and end pos of the tag
Ipos = InStr(Epos, sHTML, "<", vbBinaryCompare)
If Ipos = 0 Then Exit Do

'Get the plain text before the tag

sTXT = Mid(sHTML, Epos + 1, (Ipos - Epos))
sTXT = Replace(sTXT, "<", "")


'Get the Ending Pos of the tag
Epos = InStr(Ipos, sHTML, ">", vbBinaryCompare)
If Epos = 0 Then Exit Do

sTAG = Mid(sHTML, Ipos, (Epos - Ipos) + 1)
'AddText Obj, sTag
AddText Obj, sTXT
'format the view to match the html
DoHTMLFormatting Obj, sTAG
'print the text
'If Trim(sTXT) <> "" Then

'End If
Loop
Obj.Visible = True


End Sub

Private Sub DoHTMLFormatting(Obj As PictureBox, sTAG As String)
Dim sTXT As String
Dim Ipos As Long
sTXT = LCase(sTAG)

Select Case sTXT
Case Is = "<br>", "<p>", "</p>" ' add new line
    AddNewLine Obj
Case Is = "<b>" ' start bold face
    SetBold Obj, True
Case Is = "</b>" ' stop bold face
    SetBold Obj, False
Case Is = "<i>" 'start italics
    SetItalic Obj, True
Case Is = "</i>" 'stop italics
    SetItalic Obj, False
Case Is = "<u>" 'start uderline
    SetUnderLine Obj, True
Case Is = "</u>" 'stop underline
    SetUnderLine Obj, False
Case Is = "</a>" 'end link
    EndLink Obj
Case Is = "</font>" 'font reset
    SetDefaultColors True
    SetBold Obj, False
    SetItalic Obj, False
    SetUnderLine Obj, False
    SetFontColor Obj, lTextColor
    SetFontSize Obj, lFontSize
    SetFontFace Obj, sFontFace
Case Else
    'here we process all other types of complex tags like links, fonts, and pictures
    If InStr(1, sTXT, "<font") <> 0 Then
    'it is a font command
    ProcessFont Obj, sTXT
    End If
    If InStr(1, sTXT, "<a ") <> 0 Then
    'it is a link command
    ProcessLink Obj, sTXT
    End If
    If InStr(1, sTXT, "<img ") <> 0 Then
    'it is a image command
    
    End If
End Select

End Sub

Sub EndLink(Obj As PictureBox)
'End the Hyperlink
    isInLink = False
    SetUnderLine Obj, bUnderLinned
    SetFontColor Obj, OldColor
End Sub

Sub ProcessLink(Obj As PictureBox, sTXT As String)
Dim ILink As Long
Dim sLink As String
Dim rc As RECT
'Set the link flag
isInLink = True

'Underline the link and change the color to the link color
OldColor = Obj.ForeColor
bUnderLinned = Obj.Font.Bold
SetUnderLine Obj, True
SetFontColor Obj, lLinkColor
    
'Add the new link
ILink = UBound(AllLinks) + 1
ReDim Preserve AllLinks(ILink)

'Get Link location
sTAG = GetHTMLEle(sTXT, "href=" & Chr(34), Chr(34))
sTAG = Replace(sTAG, Chr(34) & ">", "")
If sTAG <> "" Then
sLink = sTAG
End If

LLastLink = ILink
With AllLinks(ILink)
.sLink = sLink
.rBounds.Left = TOldX * Screen.TwipsPerPixelX
.rBounds.Top = TOldY * Screen.TwipsPerPixelY
End With

End Sub
Private Sub ProcessFont(Obj As PictureBox, sTXT As String)
Dim sTAG As String

'Get the Font Color
sTAG = GetHTMLEle(sTXT, "color=" & Chr(34), Chr(34))
sTAG = Replace(sTAG, "#", "")
If sTAG <> "" Then
Obj.ForeColor = MakeHexRGB(sTAG)
End If

'Get the Font Face
sTAG = GetHTMLEle(sTXT, "face=" & Chr(34), Chr(34))
If sTAG <> "" Then
Obj.Font.Name = sTAG
End If

'Get the Font Size
sTAG = GetHTMLEle(sTXT, "size=" & Chr(34), Chr(34))
sTAG = StripNonNumeric(sTAG)
If sTAG <> "" Then
Obj.Font.Size = GetHTMLFontSize(CLng(sTAG))
End If
End Sub

Private Function GetHTMLFontSize(lNUm As Long) As Long
Select Case lNUm
Case Is = 1
    GetHTMLFontSize = 8
Case Is = 2
    GetHTMLFontSize = 10
Case Is = 3
    GetHTMLFontSize = 12
Case Is = 4
    GetHTMLFontSize = 14
Case Is = 5
    GetHTMLFontSize = 18
Case Is = 6
    GetHTMLFontSize = 24
Case Is = 7
    GetHTMLFontSize = 36
Case Else
GetHTMLFontSize = lFontSize
End Select
End Function

Private Function StripNonNumeric(sTXT As String) As Long
Dim sFinal As String
Dim I As Long
Dim S As String
For I = 1 To Len(sTXT)
S = Mid(sTXT, I, 1)
If IsNumeric(S) = True Then
sFinal = sFinal & S
End If
Next I
If Len(sFinal) = 0 Then sFinal = "0"
StripNonNumeric = CLng(sFinal)
End Function

Private Sub AddText(Obj As PictureBox, sTXT As String)
'Adds plain text to the HTML View
Dim rc As RECT
Dim lWidth As Long
Dim lHeight As Long

AlignFonts Obj, Form1.Label1
Form1.Label1.Caption = sTXT
lWidth = (Form1.Label1.Width / Screen.TwipsPerPixelX) + TOldX 'ViewMargin
lHeight = TOldY + (Form1.Label1.Height / Screen.TwipsPerPixelY) + ViewMargin
rc.Top = TOldY
rc.Left = TOldX
rc.Right = lWidth
rc.Bottom = lHeight
If (TOldX * Screen.TwipsPerPixelX) + Form1.Label1.Width >= Obj.Width Then
Obj.Width = (TOldX * Screen.TwipsPerPixelX) + Form1.Label1.Width + ((ViewMargin * Screen.TwipsPerPixelX) * 2)
End If
TCurrentY = DrawText(Obj.hdc, sTXT, -1, rc, DT_LEFT)
'TOldY = TOldY + LineBreakHeight
TOldX = rc.Right
'Obj.Line (TOldX * Screen.TwipsPerPixelX, TOldY * Screen.TwipsPerPixelY)-(lWidth * Screen.TwipsPerPixelX, TOldY * Screen.TwipsPerPixelY), vbRed
Obj.Refresh
If isInLink = True Then
With AllLinks(LLastLink)
.rBounds.Right = .rBounds.Right + Form1.Label1.Width
.rBounds.Bottom = Form1.Label1.Height
End With
End If
End Sub

Private Sub AlignFonts(sOBJ As PictureBox, sLBL As Label)
'Makes the label the same font as  the picturebox

With sLBL
.AutoSize = True
.Font.Bold = sOBJ.Font.Bold
.Font.Charset = sOBJ.Font.Charset
.Font.Italic = sOBJ.Font.Italic
.Font.Name = sOBJ.Font.Name
.Font.Size = sOBJ.Font.Size
.Font.Strikethrough = sOBJ.Font.Strikethrough
.Font.Underline = sOBJ.Font.Underline
.Font.Weight = sOBJ.Font.Weight
End With
End Sub



Private Sub AddNewLine(Obj As PictureBox)
'Adds a vbCrlf / <BR> to the HTML View
TOldY = TOldY + LineBreakHeight * 2
TOldX = ViewMargin
If TOldY * Screen.TwipsPerPixelY >= Obj.Height Then
Obj.Height = (TOldY * Screen.TwipsPerPixelY)
End If
'obj.Height = obj.Height + LineBreakHeight + ViewMargin
End Sub

Public Sub RenderHTML(Obj As PictureBox)
'Windows 95/98/Me: len(sHTML) = This number may not exceed 8192.
Dim rc As RECT
Dim result As Long


rc.Top = ViewMargin / Screen.TwipsPerPixelY
rc.Left = ViewMargin / Screen.TwipsPerPixelX
rc.Bottom = (Obj.Height / Screen.TwipsPerPixelY) - rc.Top
rc.Right = (Obj.Width / Screen.TwipsPerPixelX) - rc.Left

'get the current height, if it is bigger then the canvas, then resize canvas
TCurrentY = CSng(DrawText(Obj.hdc, sHTML, -1, rc, DT_CALCRECT)) * Screen.TwipsPerPixelY
If TCurrentY >= Obj.Height Then AddNewLine Obj
'draw the text
TCurrentY = CSng(DrawText(Obj.hdc, sHTML, -1, rc, DT_CHARSTREAM)) * Screen.TwipsPerPixelY


End Sub

Public Sub ProcessBODY(Obj As PictureBox)
'Find the <BODY> tag and get all the colors and the background
Dim Ipos As Long
Dim Epos As Long
Dim sBODY As String
Dim sTAG As String
sHTML = Replace(sHTML, "<Body", "<BODY")
sHTML = Replace(sHTML, "<body", "<BODY")
Ipos = InStr(1, sHTML, "<BODY")
If Ipos = 0 Then Exit Sub
Epos = InStr(Ipos, sHTML, ">")
If Epos = 0 Then Exit Sub

sBODY = Mid(sHTML, Ipos, Epos - Ipos + 1)
sBODY = Replace(sBODY, Chr(34), "")
sBODY = LCase(sBODY)
sTAG = GetHTMLEle(sBODY, "background=", " ")

If CheckFile(sDir & sTAG) = True Then
sBGImage = sDir & sTAG
Form1.PicImage.Picture = LoadPicture(sDir & sTAG)
End If

sTAG = GetHTMLEle(sBODY, "bgcolor=", " ")
sTAG = Replace(sTAG, "#", "")
If sTAG <> "" Then
lBGCOLOR = MakeHexRGB(sTAG)
Obj.BackColor = lBGCOLOR
End If

sTAG = GetHTMLEle(sBODY, "text=", " ")
sTAG = Replace(sTAG, "#", "")
If sTAG <> "" Then
lTextColor = MakeHexRGB(sTAG)
Obj.ForeColor = lTextColor
End If

sTAG = GetHTMLEle(sBODY, "link=", " ")
sTAG = Replace(sTAG, "#", "")
If sTAG <> "" Then
lLinkColor = MakeHexRGB(sTAG)
End If

sTAG = GetHTMLEle(sBODY, "vlink=", " ")
sTAG = Replace(sTAG, "#", "")
If sTAG <> "" Then
lVisitedColor = MakeHexRGB(sTAG)
End If

sTAG = GetHTMLEle(sBODY, "alink=", " ")
sTAG = Replace(sTAG, "#", "")
If sTAG <> "" Then
lActiveColor = MakeHexRGB(sTAG)
End If
End Sub

