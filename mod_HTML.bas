Attribute VB_Name = "mod_HTML"
'Various Subs for Working With HTML TAGS

Public Function GetHTMLEle(Origin As String, Sep1 As String, Sep2 As String) As String
'Parses a Line of text
On Error GoTo EH
Dim Bpos As Long
Dim Epos As Long
Dim SPacePOS As Long

Bpos = InStr(1, Origin, Sep1, vbBinaryCompare)
If Bpos = 0 Then Exit Function
Epos = InStr(Bpos + Len(Sep1), Origin, Sep2, vbBinaryCompare)
SPacePOS = InStr(Bpos + Len(Sep1), Origin, " ", vbBinaryCompare)

If Epos = 0 Then
 If SPacePOS = 0 Then Exit Function
 Epos = SPacePOS
End If
If SPacePOS < Epos Then
Epos = SPacePOS
End If
Bpos = Bpos + Len(Sep1)
If Epos > 0 Then
GetHTMLEle = Mid(Origin, Bpos, Epos - Bpos)
Else
GetHTMLEle = Mid(Origin, Bpos, Len(Origin))
End If
Exit Function
EH:
GetHTMLEle = ""
Exit Function
End Function

Public Function StripHTML(sHTML As String) As String
    Dim sTemp As String, lSpot1 As Long, lSpot2 As Long, lSpot3 As Long
    sTemp$ = sHTML$


    Do
        lSpot1& = InStr(lSpot3& + 1, sTemp$, "<")
        lSpot2& = InStr(lSpot1& + 1, sTemp$, ">")
        
        If lSpot1& = lSpot3& Or lSpot1& < 1 Then Exit Do
        If lSpot2& < lSpot1& Then lSpot2& = lSpot1& + 1
        sTemp$ = Left$(sTemp$, lSpot1& - 1) + Right$(sTemp$, Len(sTemp$) - lSpot2&)
        lSpot3& = lSpot1& - 1
    Loop
    StripHTML$ = sTemp$
End Function


Public Function GETHex(stColor As Long) As String
  GETHex = "#" & Hex(stColor)
End Function

Public Function GETHexOLD(stColor As Long) As String
'This is Obsolete
On Error Resume Next
'stColor = m_CurHex
       '     'If r > 255 Then Exit Sub
       '     'If g > 255 Then Exit Sub
       '     'If b > 255 Then Exit Sub
       Dim r, b, g As Long
       
       Dim dts As Variant
       Dim q, w, e As Variant
       Dim qw, we, gq As Variant
       Dim lCol As Long
       lCol = stColor
       r = lCol Mod &H100
       lCol = lCol \ &H100
       g = lCol Mod &H100
       lCol = lCol \ &H100
       b = lCol Mod &H100
       
       '     'Get Red Hex
       q = Hex(r)

              If Len(q) < 2 Then
                     qw = q
                     q = "0" & qw
              End If

       '     'Get Blue Hex
       w = Hex(b)

              If Len(w) < 2 Then
                     we = w
                     w = "0" & we
              End If

       '     'Get Green Hex
       e = Hex(g)

              If Len(e) < 2 Then
                     gq = e
                     e = "0" & gq
              End If

       'GETRGB = "#" & q & e & w
       GETHexOLD = "#" & q & e & w   '"#" &
End Function
Public Function RgbToHsv(r, g, b, h, S, V) As Long
    'Convert RGB to HSV values
    Dim vRed, vGreen, vBlue
    Dim Mx, Mn, Va, Sa, rc, gc, bc

    vRed = r / 255
    vGreen = g / 255
    vBlue = b / 255

    Mx = vRed
    If vGreen > Mx Then Mx = vGreen
    If vBlue > Mx Then Mx = vBlue

    Mn = vRed
    If vGreen < Mn Then Mn = vGreen
    If vBlue < Mn Then Mn = vBlue

    Va = Mx
    If Mx Then
        Sa = (Mx - Mn) / Mx
    Else
        Sa = 0
    End If
    If Sa = 0 Then
        h = 0
    Else
        rc = (Mx - vRed) / (Mx - Mn)
        gc = (Mx - vGreen) / (Mx - Mn)
        bc = (Mx - vBlue) / (Mx - Mn)
        Select Case Mx
        Case vRed
            h = bc - gc
        Case vGreen
            h = 2 + rc - bc
        Case vBlue
            h = 4 + gc - rc
         End Select
        h = h * 60
        If h < 0 Then h = h + 360
    End If

    S = Sa * 100
    V = Va * 100
    RgbToHsv = r + g + b
End Function

Sub HsvToRgb(h, S, V, r, g, b)
    'Convert HSV to RGB values
    Dim Sa, Va, Hue, I, f, P, q, t

    Sa = S / 100
    Va = V / 100
    If S = 0 Then
        r = Va
        g = Va
        b = Va
    Else
        Hue = h / 60
        If Hue = 6 Then Hue = 0
        I = Int(Hue)
        f = Hue - I
        P = Va * (1 - Sa)
        q = Va * (1 - (Sa * f))
        t = Va * (1 - (Sa * (1 - f)))
        Select Case I
        Case 0
            r = Va
            g = t
            b = P
        Case 1
            r = q
            g = Va
            b = P
        Case 2
            r = P
            g = Va
            b = t
        Case 3
            r = P
            g = q
            b = Va
        Case 4
            r = t
            g = P
            b = Va
        Case 5
            r = Va
            g = P
            b = q
        End Select
    End If
    
    r = Int(255.9999 * r)
    g = Int(255.9999 * g)
    b = Int(255.9999 * b)
End Sub
Public Sub GETRGB(SColor As Variant)
Dim r, g, b As Integer
r = SColor Mod &H100
g = (SColor \ &H100) Mod &H100
b = (SColor \ &H10000) Mod &H100
frmColor.Text3.Text = r
frmColor.Text4.Text = g
frmColor.Text5.Text = b

Dim rr, gg, bb, maincolor As Long


'frmColor.BackColor = (r * 255) + (g * 255) + (b * 255)
End Sub



Sub FadeSide2Side(Form As Object, Color1 As Long, Color2 As Long)
Dim X!, x2!, Y%, I%, red1%, green1%, blue1%, red2%, green2%, blue2%, pat1!, pat2!, pat3!, c1!, c2!, c3!
' find the length of the form and cut it into 80 pieces
x2 = (Form.Width / 80) / 2
Y% = Form.ScaleHeight
' separating red, green, and blue in each of the two colors
red1% = Color1 And 255
green1% = Color1 \ 256 And 255
blue1% = Color1 \ 65536 And 255
red2% = Color2 And 255
green2% = Color2 \ 256 And 255
blue2% = Color2 \ 65536 And 255
' cut the difference between the two colors into 100 pieces
pat1 = (red2% - red1%) / 80
pat2 = (green2% - green1%) / 80
pat3 = (blue2% - blue1%) / 80
' set the c variables at the starting colors
c1 = red1%
c2 = green1%
c3 = blue1%
' draw 80 different lines on the form
For I% = 1 To 80
Form.Line (X, 0)-(X + x2, Y%), RGB(c1, c2, c3), BF
X = X + x2 ' draw the Next line one step up from the old step
c1 = c1 + pat1 ' make the c variable equal 2 it's Next step
c2 = c2 + pat2
c3 = c3 + pat3
 Next
Form.CurrentX = 0
Form.CurrentY = 0
'Form.Print "Click And Resize Me!" ' Note: remove this line when making your own projects
End Sub

Sub FadeSide2Side2(Form As Object, Color1 As Long, Color2 As Long)
Dim X!, x2!, Y%, I%, red1%, green1%, blue1%, red2%, green2%, blue2%, pat1!, pat2!, pat3!, c1!, c2!, c3!
' find the length of the form and cut it into 80 pieces
x2 = (Form.Width / 80) / 2
Y% = Form.ScaleHeight
' separating red, green, and blue in each of the two colors
red1% = Color1 And 255
green1% = Color1 \ 256 And 255
blue1% = Color1 \ 65536 And 255
red2% = Color2 And 255
green2% = Color2 \ 256 And 255
blue2% = Color2 \ 65536 And 255
' cut the difference between the two colors into 100 pieces
pat1 = (red2% - red1%) / 80
pat2 = (green2% - green1%) / 80
pat3 = (blue2% - blue1%) / 80
' set the c variables at the starting colors
c1 = red1%
c2 = green1%
c3 = blue1%
' draw 80 different lines on the form
X = Form.Width / 2 - 55
For I% = 1 To 80
'picture1.Line
'Form.Line (x, 0)-(x + x2, y%), RGB(c1, c2, c3), BF
Form.Line (X, 0)-(X + x2, Y%), RGB(c1, c2, c3), BF
X = X + x2 ' draw the Next line one step up from the old step
c1 = c1 + pat1 ' make the c variable equal 2 it's Next step
c2 = c2 + pat2
c3 = c3 + pat3
 Next
Form.CurrentX = 0
Form.CurrentY = 0

'Form.Print "Click And Resize Me!" ' Note: remove this line when making your own projects
End Sub

Public Function MakeRGBHex(stColor As Long) As String
On Error GoTo EH
Dim r, b, g As Long
Dim dts As Variant
Dim q, w, e As Variant
Dim qw, we, gq As Variant
Dim lCol As Long
lCol = stColor
r = lCol Mod &H100
lCol = lCol \ &H100
g = lCol Mod &H100
lCol = lCol \ &H100
b = lCol Mod &H100
q = Hex(r)
If Len(q) < 2 Then
qw = q
q = "0" & qw
End If
w = Hex(b)
If Len(w) < 2 Then
we = w
w = "0" & we
End If
e = Hex(g)
If Len(e) < 2 Then
gq = e
e = "0" & gq
End If
MakeRGBHex = "#" & q & e & w
Exit Function
EH:
MakeRGBHex = "#" & q & e & w
Exit Function
End Function

''''''''''''''''''''''''''''''''''''''''''''''
Public Function MakeHexRGB(sHex As String) As Long
On Error GoTo Errh
Dim Ipos As Integer
Dim tmpStr As String
Dim P1 As String
Dim P2 As String
Dim P3 As String
Dim pFin As String
Ipos = InStr(1, sHex, "#", vbBinaryCompare)
If Ipos = 0 Then
tmpStr = sHex
Else
If Ipos <> 0 Then
tmpStr = Mid(sHex, Ipos + 1, Len(sHex))
End If
End If
P1 = Mid(tmpStr, 1, 2)
P2 = Mid(tmpStr, 3, 2)
P3 = Mid(tmpStr, 5, 2)
pFin = P3 & P2 & P1
MakeHexRGB = CLng("&H" & pFin)
Errh:
Exit Function
End Function

