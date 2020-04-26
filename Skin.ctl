VERSION 5.00
Begin VB.UserControl SKin 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.PictureBox bEnd 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      BorderStyle     =   0  '없음
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   1665
      Picture         =   "Skin.ctx":0000
      ScaleHeight     =   225
      ScaleWidth      =   255
      TabIndex        =   1
      Top             =   80
      Width           =   255
   End
   Begin VB.PictureBox bMinimize 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      BorderStyle     =   0  '없음
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   960
      Picture         =   "Skin.ctx":00AC
      ScaleHeight     =   225
      ScaleWidth      =   255
      TabIndex        =   0
      Top             =   80
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   225
      Left            =   1320
      Picture         =   "Skin.ctx":0147
      Top             =   80
      Width           =   255
   End
   Begin VB.Label lblDrag 
      BackStyle       =   0  '투명
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   30
      Width           =   1215
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  '투명
      Caption         =   "lblCaption"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   8.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1560
   End
   Begin VB.Image iEndC 
      Height          =   225
      Left            =   3120
      Picture         =   "Skin.ctx":01EA
      Top             =   1200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image iEndR 
      Height          =   225
      Left            =   3120
      Picture         =   "Skin.ctx":02EA
      Top             =   960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image iEndD 
      Height          =   225
      Left            =   3120
      Picture         =   "Skin.ctx":0469
      Top             =   720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image iMinimizeC 
      Height          =   225
      Left            =   2640
      Picture         =   "Skin.ctx":0515
      Top             =   1200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image iMinimizeR 
      Height          =   225
      Left            =   2640
      Picture         =   "Skin.ctx":0609
      Top             =   960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image iMinimizeD 
      Height          =   225
      Left            =   2640
      Picture         =   "Skin.ctx":06F3
      Top             =   720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image iDownRight 
      Height          =   210
      Left            =   1680
      Picture         =   "Skin.ctx":078E
      Top             =   1680
      Width           =   270
   End
   Begin VB.Image iDown 
      Height          =   210
      Left            =   240
      Picture         =   "Skin.ctx":07EA
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   1530
   End
   Begin VB.Image iDownLeft 
      Height          =   210
      Left            =   0
      Picture         =   "Skin.ctx":0858
      Top             =   1680
      Width           =   270
   End
   Begin VB.Image iRight 
      Height          =   1200
      Left            =   1680
      Picture         =   "Skin.ctx":08B4
      Stretch         =   -1  'True
      Top             =   480
      Width           =   270
   End
   Begin VB.Image iCenter 
      Height          =   1200
      Left            =   240
      Picture         =   "Skin.ctx":094C
      Stretch         =   -1  'True
      Top             =   480
      Width           =   1530
   End
   Begin VB.Image iLeft 
      Height          =   1200
      Left            =   0
      Picture         =   "Skin.ctx":09D6
      Stretch         =   -1  'True
      Top             =   480
      Width           =   270
   End
   Begin VB.Image iUpRight 
      Height          =   525
      Left            =   1680
      Picture         =   "Skin.ctx":0A6E
      Top             =   0
      Width           =   270
   End
   Begin VB.Image iUp 
      Height          =   525
      Left            =   240
      Picture         =   "Skin.ctx":0BA8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1530
   End
   Begin VB.Image iUpLeft 
      Height          =   525
      Left            =   0
      Picture         =   "Skin.ctx":0CE4
      Top             =   0
      Width           =   270
   End
End
Attribute VB_Name = "SKin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'//이미지 롤오버
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetCapture Lib "user32" () As Long

 '//폼드래그
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const HTCAPTION = 2
Private Const WM_NCLBUTTONDOWN = &HA1

Private CName As String


Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Function RoundFRM(frm As Form)
Dim Result As Long
Result = CreateRoundRectRgn(0, 0, (frm.Width + 10) / Screen.TwipsPerPixelX, (frm.Height + 100) / Screen.TwipsPerPixelY, 5, 5)
SetWindowRgn frm.hwnd, Result, True
End Function



Private Sub lblDrag_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'//폼드래그
Dim lngReturnValue As Long

If Button = 1 Then
Call ReleaseCapture
lngReturnValue = SendMessage(Parent.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End If
End Sub


Private Sub bEnd_Click()
Unload Parent
End Sub

Private Sub bMinimize_Click()
Parent.WindowState = 1
End Sub

Private Sub bMinimize_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (X < 0) Or (Y < 0) Or (X > bMinimize.Width) Or (Y > bMinimize.Height) Then
           ReleaseCapture
                '마우스가 이미지에서 벗어남
           bMinimize.Picture = iMinimizeD.Picture
    ElseIf GetCapture() <> bMinimize.hwnd Then
           SetCapture bMinimize.hwnd
                '마우스가 이미지 위에있음
           bMinimize.Picture = iMinimizeR.Picture
    End If
End Sub


Private Sub bMinimize_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'클릭했음
bMinimize.Picture = iMinimizeC.Picture
End Sub


Private Sub bEND_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (X < 0) Or (Y < 0) Or (X > bEnd.Width) Or (Y > bEnd.Height) Then
           ReleaseCapture
                '마우스가 이미지에서 벗어남
           bEnd.Picture = iEndD.Picture
    ElseIf GetCapture() <> bEnd.hwnd Then
           SetCapture bEnd.hwnd
                '마우스가 이미지 위에있음
           bEnd.Picture = iEndR.Picture
    End If
End Sub


Private Sub bEND_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'클릭했음
bEnd.Picture = iEndC.Picture

End Sub

Private Sub UserControl_Resize()
On Error Resume Next

'//상단
iUp.Left = iUpLeft.Width
iUp.Width = Width - iUpLeft.Width - iUpRight.Width
iUpRight.Left = iUpLeft.Width + iUp.Width

'//중간
iLeft.Top = iUp.Height
iCenter.Top = iUp.Height
iRight.Top = iUp.Height

iCenter.Left = iUp.Left
iCenter.Width = iUp.Width
iRight.Left = iUpRight.Left

iCenter.Height = Height - iUp.Height - iDown.Height
iLeft.Height = Height - iUp.Height - iDown.Height
iRight.Height = Height - iUp.Height - iDown.Height



'//하단
iDownLeft.Top = iUp.Height + iCenter.Height
iDown.Top = iUp.Height + iCenter.Height
iDownRight.Top = iUp.Height + iCenter.Height

iDown.Left = iUp.Left
iDown.Width = iUp.Width
iDownRight.Left = iRight.Left

'//버튼
bMinimize.Left = Width - 900
Image1.Left = Width - 600
bEnd.Left = Width - 300

'//폼캡션
lblCaption.Width = Width

'//폼드래그
lblDrag.Width = Width

End Sub



Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
CName = PropBag.ReadProperty("Caption", UserControl.Name)

lblCaption = CName

End Sub

Private Sub UserControl_Show()
RoundFRM Parent
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Call PropBag.WriteProperty("Caption", CName, Empty)
End Sub

Public Property Get Caption() As String
Caption = CName
End Property

Public Property Let Caption(Str As String)
CName = Str
lblCaption = CName
End Property

