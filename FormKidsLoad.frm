VERSION 5.00
Begin VB.Form FormKidsLoad 
   BackColor       =   &H80000010&
   BorderStyle     =   4  '고정 도구 창
   Caption         =   "자녀보호 모드"
   ClientHeight    =   2520
   ClientLeft      =   10500
   ClientTop       =   2955
   ClientWidth     =   4800
   Icon            =   "FormKidsLoad.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Visible         =   0   'False
   Begin KidsLock.isButton isButton1 
      Height          =   300
      Left            =   3240
      TabIndex        =   9
      Top             =   0
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   529
      Icon            =   "FormKidsLoad.frx":1857E
      Style           =   2
      Caption         =   "부모님 모드 전환"
      IconAlign       =   1
      iNonThemeStyle  =   0
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000010&
      Caption         =   "관리자 로그인"
      Height          =   615
      Left            =   240
      TabIndex        =   8
      Top             =   1680
      Visible         =   0   'False
      Width           =   4335
      Begin KidsLock.isButton isButton2 
         Height          =   300
         Left            =   3120
         TabIndex        =   11
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
         Icon            =   "FormKidsLoad.frx":1859A
         Style           =   0
         Caption         =   "확인"
         IconAlign       =   1
         iNonThemeStyle  =   0
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   0
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox Text3 
         Height          =   270
         IMEMode         =   3  '사용 못함
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   10
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   300
      Left            =   1920
      TabIndex        =   7
      Top             =   2760
      Value           =   1  '확인
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Timer Timer2 
      Interval        =   10
      Left            =   4080
      Top             =   720
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "0"
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "1"
      Top             =   2760
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   3600
      Top             =   720
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000010&
      Caption         =   "남은 PC 사용 가능시간"
      Height          =   1335
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   4335
      Begin VB.Label Label3 
         BackStyle       =   0  '투명
         Caption         =   "남은시간...."
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   855
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   3615
      End
   End
   Begin VB.PictureBox P 
      Height          =   255
      Left            =   8160
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   0
      Top             =   4560
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '투명
      Caption         =   "동작중이 아님..."
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000010&
      Caption         =   "자녀보호 모드가 "
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   2775
   End
   Begin VB.Menu msys 
      Caption         =   "msys"
      Visible         =   0   'False
      Begin VB.Menu mabout 
         Caption         =   "정보 (&About)"
      End
   End
End
Attribute VB_Name = "FormKidsLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type NOTIFYICONDATA
  cbSize As Long
  hwnd As Long
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
Dim md5Test As MD5 'md5 암호화 선언
Dim SysTrayT As NOTIFYICONDATA

Private Sub Form_Load()
On Error Resume Next
Set md5Test = New MD5 'md5 선언
Timer1.Enabled = True
Timer2.Enabled = True
Check1.Value = 1
Text2.Text = 0
Dim hour As Integer
Dim min As Integer
hour = ReadINI("setting", "hour", Environ$("AppData") & "\KidsLock.ini")
min = ReadINI("setting", "min", Environ$("AppData") & "\KidsLock.ini")

 Text1.Text = hour * 60 + min
        Timer1.Enabled = True
        Label2.Caption = "동작중..."
        Label3.Caption = "남은시간:" & Text1.Text - Text2.Text & "분"
        
With SysTrayT
        .cbSize = Len(SysTrayT)
        .hwnd = P.hwnd
        .uID = 1
        .uFlags = &H2 Or &H1 Or &H10 Or &H4
        .hIcon = Me.Icon
        .uCallbackMessage = &H200
        
        .szTip = "Kids Lock -자녀보호 모드" & Chr(0) ' 도움 팁
        .szInfoTitle = "Kids Lock" & Chr(0) ' 풍선 테이블 네임
        .szInfo = "자녀보호 모드가 실행되었습니다!" & vbCrLf & "설정된 시간 : " & hour & "시간" & min & "분" & Chr(0) ' 풍선 메세지
        .uTimeoutOrVersion = 15000 '풍선을 몇초 동안 보여줌 (1000 = 1초)
    End With
    
        Shell_NotifyIcon &H0, SysTrayT ' Action to take: &H0 = ADD, &H1 = MODIFY, &H2 = DELETE
        ProtectProcess
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Check1.Value = 1 Then
ProtectProcess
Cancel = 1
Else
Cancel = 0
 RestoreProcess '윈도우 중요 프로세스 해재 정상 종료 가능 하게 만듬
End If
End Sub

Private Sub isButton1_Click()
Frame2.Visible = True
End Sub

Private Sub isButton2_Click()
On Error Resume Next
Text3.Text = LCase(md5Test.DigestStrToHexStr(Text3.Text))
If GetSetting("KidLock", "admin", "password") = Text3.Text Then
MsgBox "관리자 인증 성공!", vbDefaultButton1, "성공!"
Timer1.Enabled = False
Check1.Value = 0
RestoreProcess '윈도우 중요 프로세스 해재 정상 종료 가능 하게 만듬
Unload Me
FormLoad.Show
Else
MsgBox "관리자 암호가 일치하지 않습니다!", vbCritical, "ERROR"
Label1.Visible = True
Text3.Text = vbNullString
Frame2.Visible = False
End If
End Sub

Private Sub mabout_Click()
FormAbout.Show
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

Text2.Text = Text2.Text + 1 '1분마다 더하기 1하기

Label3.Caption = "남은시간:" & Text1.Text - Text2.Text & "분"
End Sub

Private Sub Timer2_Timer()

On Error Resume Next

If Text2.Text = Text1.Text Then '만약 텍스트1과2가 같을경우 실행하고 타이머 끄기
Text2.Text = 999
Label2.Caption = "동작중이 아님..."
Label3.Caption = "현재 자녀보호 모드가 아닙니다."

Unload FormLoad
Unload FormMain
Unload Formps
Unload FormSetting
Unload FormChange
Unload FormAbout
Unload FormConfirm
Unload FormKidsTry

FormLock.Show

Timer1.Enabled = False
Check1.Value = 0
RestoreProcess '윈도우 중요 프로세스 해재 정상 종료 가능 하게 만듬
Unload Me

End If
ProcessKill GetProcess(ReadINI("kill", "default1", Environ$("AppData") & "\KidsLock.ini"))
ProcessKill GetProcess(ReadINI("kill", "default2", Environ$("AppData") & "\KidsLock.ini"))
ProcessKill GetProcess(ReadINI("kill", "kill1", Environ$("AppData") & "\KidsLock.ini"))
ProcessKill GetProcess(ReadINI("kill", "kill1", Environ$("AppData") & "\KidsLock.ini"))
ProcessKill GetProcess(ReadINI("kill", "kill2", Environ$("AppData") & "\KidsLock.ini"))
ProcessKill GetProcess(ReadINI("kill", "kill3", Environ$("AppData") & "\KidsLock.ini"))
ProcessKill GetProcess(ReadINI("kill", "kill4", Environ$("AppData") & "\KidsLock.ini"))
ProcessKill GetProcess(ReadINI("kill", "kill5", Environ$("AppData") & "\KidsLock.ini"))
End Sub

