Attribute VB_Name = "Module1"

'Global FontFace As String
'Global FontSize As Integer
'Global FontColor As Long
'Global LinkColor As Long
'Global BGColor As Long



Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Public Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Public Declare Function CreateFontIndirectA Lib "gdi32" (lpLogFont As LOGFONT) As Long

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
' Logical Font
Public Const LF_FACESIZE = 32
Public Const LF_FULLFACESIZE = 64
Public Const FONT_SIZE = 12
Public Const NO_ERROR = 0
Public Const ANSI_CHARSET = 0
Public Const OUT_DEFAULT_PRECIS = 0
Public Const CLIP_DEFAULT_PRECIS = 0
Public Const DEFAULT_QUALITY = 0
Public Const DEFAULT_PITCH = 0
Public Const FF_DONTCARE = 0
Public Const LOGPIXELSX = 88        '  Logical pixels/inch in X
Public Const LOGPIXELSY = 90        '  Logical pixels/inch in Y
Public Const TRANSPARENT = 1

Public Type LOGFONT
   lfHeight As Long
   lfWidth As Long
   lfEscapement As Long
   lfOrientation As Long
   lfWeight As Long
   lfItalic As Byte
   lfUnderline As Byte
   lfStrikeOut As Byte
   lfCharSet As Byte
   lfOutPrecision As Byte
   lfClipPrecision As Byte
   lfQuality As Byte
   lfPitchAndFamily As Byte
   lfFaceName(LF_FACESIZE) As Byte
End Type




Public Const ViewMargin = 5 'pixels
Public Const LineBreakHeight = 10 'pixels

Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_TRANSPARENT = &H20&

Public Const DT_BOTTOM = &H8
Public Const DT_CALCRECT = &H400
Public Const DT_CENTER = &H1
Public Const DT_CHARSTREAM = 4          '  Character-stream, PLP
Public Const DT_DISPFILE = 6            '  Display-file
Public Const DT_EXPANDTABS = &H40
Public Const DT_EXTERNALLEADING = &H200
Public Const DT_INTERNAL = &H1000
Public Const DT_LEFT = &H0
Public Const DT_METAFILE = 5            '  Metafile, VDM
Public Const DT_NOCLIP = &H100
Public Const DT_NOPREFIX = &H800
Public Const DT_PLOTTER = 0             '  Vector plotter
Public Const DT_RASCAMERA = 3           '  Raster camera
Public Const DT_RASDISPLAY = 1          '  Raster display
Public Const DT_RASPRINTER = 2          '  Raster printer
Public Const DT_RIGHT = &H2
Public Const DT_SINGLELINE = &H20
Public Const DT_TABSTOP = &H80
Public Const DT_TOP = &H0
Public Const DT_VCENTER = &H4
Public Const DT_WORDBREAK = &H10



Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Const SRCCOPY = &HCC0020

Public Sub TileBitmap(Target As Object, Source As Object)

BackupInformation_ScaleMode = Target.ScaleMode
BackupInformation_ScaleMode2 = Source.ScaleMode
Source.ScaleMode = 3
Target.ScaleMode = 3
Target.Cls
Target.AutoRedraw = True


For yDraw = 0 To Target.Height Step Source.ScaleHeight

    For Xdraw = 0 To Target.ScaleWidth Step Source.ScaleWidth

        BitBlt Target.hdc, Xdraw, yDraw, Source.ScaleWidth, Source.ScaleHeight, Source.hdc, 0, 0, SRCCOPY



    Next Xdraw

Next yDraw


Target.ScaleMode = BackupInformation_ScaleMode
Source.ScaleMode = BackupInformation_ScaleMode2

End Sub

Public Function OpenTextFile(sFile As String) As String
'Reads an entire file into a string
On Error GoTo EH
Dim TMPTXT As String
Dim FinTxt As String
Dim iFile As Integer
iFile = FreeFile
Open sFile For Binary Access Read As #iFile
TMPTXT = Space$(LOF(iFile))
Get #iFile, , TMPTXT
Close #iFile
OpenTextFile = TMPTXT
Exit Function
EH:
OpenTextFile = ""
Exit Function
End Function

Public Function CheckFile(sFile As String) As Boolean
'Does a file exist TRUE / FALSE
On Error Resume Next
If sFile = "" Then
CheckFile = False
Exit Function
End If
Dim Iret
Iret = Dir(sFile)
If Iret > "" Then
CheckFile = True
Else
If Iret = "" Then
CheckFile = False
End If
End If

End Function
