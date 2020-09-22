VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Form1 
   Caption         =   "View HTML 1.2"
   ClientHeight    =   7710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8025
   LinkTopic       =   "Form1"
   ScaleHeight     =   7710
   ScaleWidth      =   8025
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PicImage 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   6000
      ScaleHeight     =   1665
      ScaleWidth      =   1425
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSComCtl2.FlatScrollBar HScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   4440
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   450
      _Version        =   393216
      MousePointer    =   1
      Appearance      =   2
      Arrows          =   65536
      LargeChange     =   500
      Min             =   1
      Orientation     =   8323073
      SmallChange     =   120
      Value           =   1
   End
   Begin VB.PictureBox PicBacking 
      AutoRedraw      =   -1  'True
      Height          =   4455
      Left            =   0
      ScaleHeight     =   4395
      ScaleWidth      =   5355
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      Begin VB.PictureBox picCanvas 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   735
         Left            =   0
         MouseIcon       =   "Form1.frx":0000
         ScaleHeight     =   735
         ScaleWidth      =   4575
         TabIndex        =   1
         Top             =   0
         Width           =   4575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000007&
         Height          =   195
         Left            =   0
         TabIndex        =   5
         Top             =   720
         Visible         =   0   'False
         Width           =   1125
      End
   End
   Begin MSComCtl2.FlatScrollBar VScroll1 
      Height          =   3375
      Left            =   5520
      TabIndex        =   3
      Top             =   0
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   5953
      _Version        =   393216
      MousePointer    =   1
      Appearance      =   2
      LargeChange     =   500
      Min             =   1
      Orientation     =   8323072
      SmallChange     =   120
      Value           =   1
   End
   Begin VB.Image imgCursor 
      Height          =   480
      Left            =   6120
      Picture         =   "Form1.frx":030A
      Top             =   2520
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OverLink As Boolean
Dim iOverLink As Long
Dim OldHeight As Single
Dim OldWidth As Single
Dim bLoaded As Boolean
Private Sub Form_Load()
bLoaded = False
SetDefaultColors
sHTML = OpenTextFile(App.Path & "\test.html")
sDir = App.Path & "\"
'process the <BODY> tag first
Me.Show
DoEvents
ProcessBODY picCanvas
'DisplayWebPage picCanvas

bLoaded = True
picCanvas_Resize
Form_Resize

End Sub

Private Sub Form_Resize()
On Error Resume Next
PicBacking.Move 0, 0, (Me.Width - 120) - VScroll1.Width, (Me.Height - 425) - HScroll1.Height
VScroll1.Move PicBacking.Width, 0, 255, PicBacking.Height
HScroll1.Move 0, PicBacking.Height, PicBacking.Width, 255
'resize the scroll bars
VScroll1.Max = picCanvas.Height
HScroll1.Max = picCanvas.Width
If picCanvas.Height <= PicBacking.Height Then VScroll1.Max = VScroll1.Min
If picCanvas.Width <= PicBacking.Width Then HScroll1.Max = HScroll1.Min
If picCanvas.Height > PicBacking.Height Then VScroll1.Max = PicBacking.Height - picCanvas.Height
If picCanvas.Width > PicBacking.Width Then HScroll1.Max = PicBacking.Width - picCanvas.Width


End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim Ret As Object
For Each Ret In Forms
Unload Ret
Next Ret
End
End Sub

Private Sub HScroll1_Change()
If HScroll1.Max = HScroll1.Min Then Exit Sub
If HScroll1.Value = HScroll1.Min Then picCanvas.Top = 0
picCanvas.Left = HScroll1.Value
End Sub

Private Sub PicBacking_Resize()
If PicBacking.Width > 4575 Then
picCanvas.Width = PicBacking.Width
End If
If PicBacking.Height > picCanvas.Height Then picCanvas.Height = PicBacking.Height
End Sub

Private Sub picCanvas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'check to see if we are in a link
On Error GoTo EH
Dim rc As RECT
Dim I As Long
For I = 0 To UBound(AllLinks)
rc = AllLinks(I).rBounds
OverLink = False
'Me.Caption = X & "-" & RC.Left & " " & Y & "-" & RC.Top & " " & AllLinks(I).sLink
If X >= rc.Left And X < rc.Left + rc.Right Then
    If Y >= rc.Top And Y < rc.Top + rc.Bottom Then
    'we are over a link
        OverLink = True
        iOverLink = I
    Exit For
    End If
End If
Next I

'show the mouseover cursor
If OverLink = True Then
picCanvas.MousePointer = 99
Else
iOverLink = 0
picCanvas.MousePointer = vbArrow
End If

Exit Sub
EH:
'do nothing
Exit Sub
End Sub

Private Sub picCanvas_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'see if we are clicking a link
Dim sFile As String
If OverLink = True Then
sFile = sDir & AllLinks(iOverLink).sLink

If CheckFile(sFile) = True Then
'local file exists, load it
    SetDefaultColors
    sHTML = OpenTextFile(sFile)
    'process the <BODY> tag first
    ProcessBODY picCanvas
    DisplayWebPage picCanvas
Else
    MsgBox "File not found:" & vbCrLf & sFile, vbInformation, "Error"
End If
End If
End Sub

Private Sub picCanvas_Resize()
If bLoaded = True Then
TileBitmap picCanvas, PicImage
DisplayWebPage picCanvas
End If
End Sub

Private Sub VScroll1_Change()
If VScroll1.Max = VScroll1.Min Then Exit Sub
If VScroll1.Value = VScroll1.Min Then picCanvas.Top = 0
picCanvas.Top = VScroll1.Value
End Sub



