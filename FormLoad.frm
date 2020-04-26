VERSION 5.00
Begin VB.Form FormLoad 
   BorderStyle     =   0  '없음
   Caption         =   "로딩중"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10275
   Icon            =   "FormLoad.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   10275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   9720
      Top             =   120
   End
   Begin VB.PictureBox Picture1 
      Height          =   5895
      Left            =   -120
      Picture         =   "FormLoad.frx":1857E
      ScaleHeight     =   5835
      ScaleWidth      =   10395
      TabIndex        =   1
      Top             =   -120
      Width           =   10455
   End
   Begin VB.PictureBox P 
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   675
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin VB.Menu msys 
      Caption         =   "msys"
      Visible         =   0   'False
      Begin VB.Menu moption 
         Caption         =   "관리자 설정 (&Option)"
      End
      Begin VB.Menu line1 
         Caption         =   "-"
      End
      Begin VB.Menu mabout 
         Caption         =   "정보 (&About)"
      End
   End
End
Attribute VB_Name = "FormLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type NOTIFYICONDATA
  cbSize As Long
  hWnd As Long
  uID As Long
  uFlags As Long
  uCallbackMessage As Long
  hIcon As Long
  szTip As String * 128
  
  dwState As Long
  dwStateMask As Long
  szInfo As String * 256
  uTimeoutOrVersion As Long
  szInfoTitle As String * 64
  dwInfoFlags As Long
End Type
  

Private Const NIIF_WARNING = 2
Private Const NIIF_ERROR = 3
Private Const NIIF_INFO = 1

Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Dim SysTrayT As NOTIFYICONDATA
Private Sub Form_Load()
On Error Resume Next
If ReadINI("mj", "bj", Environ$("AppData") & "\KidsLock.ini") = 0 Then
MsgBox "KidsLock 를 설치해주셔서 감사합니다.", vbDefaultButton1, "Thankyou!"
MsgBox "초기 관리자 비밀번호는 kidslockadmin 입니다 설정에서 꼭 변경해주세요.", vbInformation, "알림!"
SaveSetting "KidLock", "admin", "password", "1fec1b134299a83474b480d3d60a9621"
Call WriteINI("mj", "bj", "1", Environ$("AppData") & "\KidsLock.ini")
Else
End If

Unload FormKidsLoad
Unload FormKidsTry
Me.Hide

With SysTrayT
        .cbSize = Len(SysTrayT)
        .hWnd = P.hWnd
        .uID = 1
        .uFlags = &H2 Or &H1 Or &H10 Or &H4
        .hIcon = Me.Icon
        .uCallbackMessage = &H200
        
        .szTip = "Kids Lock -부모님 모드" & Chr(0) ' 도움 팁
        .szInfoTitle = "Kids Lock" & Chr(0) ' 풍선 테이블 네임
        .szInfo = "트레이 모드로 동작합니다!" & Chr(0)   ' 풍선 메세지
        .uTimeoutOrVersion = 15000 '풍선을 몇초 동안 보여줌 (1000 = 1초)
    End With
    
        Shell_NotifyIcon &H0, SysTrayT ' Action to take: &H0 = ADD, &H1 = MODIFY, &H2 = DELETE

End Sub


Private Sub mabout_Click()
FormAbout.Show
End Sub

Private Sub moption_Click()
FormSetting.Show
End Sub

Private Sub P_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Static rec As Boolean, Msg As Long
    Msg = X / Screen.TwipsPerPixelX
    If rec = False Then
        rec = True
        Select Case Msg
            Case &H202: FormAbout.Show '왼쪽마우스 클릭하면 발생하는 이벤트
            Case &H205: PopupMenu msys ' 오른쪽 마우스 클릭하면 발생하는 이벤트
        End Select
        rec = False
    End If
End Sub


Private Sub Timer1_Timer()
FormLoad.Visible = False
End Sub
