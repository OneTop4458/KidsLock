VERSION 5.00
Begin VB.Form FormLock 
   BorderStyle     =   4  '고정 도구 창
   Caption         =   "사용가능 시간이 종료되었습니다...."
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10260
   Icon            =   "FormLock.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   10260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   5760
      Value           =   1  '확인
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   5775
      Left            =   -120
      Picture         =   "FormLock.frx":1857E
      ScaleHeight     =   5715
      ScaleWidth      =   10395
      TabIndex        =   0
      Top             =   -120
      Width           =   10455
      Begin VB.Timer Timer3 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   8880
         Top             =   3720
      End
      Begin VB.Timer Timer2 
         Interval        =   10000
         Left            =   8880
         Top             =   3240
      End
      Begin VB.Timer Timer1 
         Interval        =   10
         Left            =   8880
         Top             =   2760
      End
      Begin KidsLock.isButton isButton3 
         Height          =   300
         Left            =   3120
         TabIndex        =   4
         Top             =   5160
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         Icon            =   "FormLock.frx":221A7
         Style           =   5
         Caption         =   "컴퓨터 다시시작"
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
      Begin KidsLock.isButton isButton2 
         Height          =   300
         Left            =   1200
         TabIndex        =   3
         Top             =   5160
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         Icon            =   "FormLock.frx":221C3
         Style           =   5
         Caption         =   "컴퓨터 종료"
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
      Begin KidsLock.isButton isButton1 
         Height          =   855
         Left            =   7440
         TabIndex        =   2
         Top             =   4080
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1508
         Icon            =   "FormLock.frx":221DF
         Style           =   4
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
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         IMEMode         =   3  '사용 못함
         Left            =   2280
         PasswordChar    =   "*"
         TabIndex        =   1
         ToolTipText     =   "비밀번호를 입력하여 주세요."
         Top             =   4320
         Width           =   4815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "ERROR: 관리자 비밀번호가 일치하지 않습니다!!"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   2280
         TabIndex        =   5
         Top             =   4800
         Visible         =   0   'False
         Width           =   4935
      End
   End
End
Attribute VB_Name = "FormLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim md5Test As MD5 'md5 암호화 선언

Private Sub Form_Load()
Timer3.Enabled = False
On Error Resume Next
Set md5Test = New MD5 'md5 선언
Check1.Value = 1
ProtectProcess
AlwaysTop FormLock, True
End Sub

Private Sub isButton1_Click()
On Error Resume Next

Text1.Text = LCase(md5Test.DigestStrToHexStr(Text1.Text))

If GetSetting("KidLock", "admin", "password") = Text1.Text Then

Label1.Visible = True
Label1.Caption = "5초후 자동으로 부모모드로 전환됩니다."

Timer1.Enabled = False
AlwaysTop FormLock, False
CreateObject("WScript.Shell").Run "C:\Windows\explorer.exe" '죽였던 explorer 부활
Text1.Text = vbNullString '텍스트 박스 값 초기화

Timer3.Enabled = True

 RestoreProcess '윈도우 중요 프로세스 해재 정상 종료 가능 하게 만듬
 
 Check1.Value = 0
 
 
 
 Else

Label1.Visible = True
Label1.Caption = "ERROR: 관리자 비밀번호가 일치하지 않습니다!!"
Text1.Text = vbNullString

End If
End Sub

Private Sub isButton2_Click()
AlwaysTop FormLock, False
FormLock.Hide
Check1.Value = 0
Text1.Text = vbNullString '텍스트 박스 값 초기화

Timer1.Enabled = False '타이머 1 작동중지
RestoreProcess
    Shell "shutdown -s -t 1" '종료
    End
End Sub

Private Sub isButton3_Click()
AlwaysTop FormLock, False
FormLock.Hide
Check1.Value = 0
Text1.Text = vbNullString '텍스트 박스 값 초기화

Timer1.Enabled = False '타이머 1 작동중지
RestoreProcess
    Shell "shutdown -r -t 1" '다시시작
    End
End Sub


Private Sub Timer1_Timer()
On Error Resume Next
ProcessKill GetProcess("explorer.exe") '//explorer 차단
'최대 10 까지 ini 읽고 종료
ProcessKill GetProcess(ReadINI("kill", "default1", Environ$("AppData") & "\KidsLock.ini"))
ProcessKill GetProcess(ReadINI("kill", "default2", Environ$("AppData") & "\KidsLock.ini"))
ProcessKill GetProcess(ReadINI("kill", "kill1", Environ$("AppData") & "\KidsLock.ini"))
ProcessKill GetProcess(ReadINI("kill", "kill2", Environ$("AppData") & "\KidsLock.ini"))
ProcessKill GetProcess(ReadINI("kill", "kill3", Environ$("AppData") & "\KidsLock.ini"))
ProcessKill GetProcess(ReadINI("kill", "kill4", Environ$("AppData") & "\KidsLock.ini"))
ProcessKill GetProcess(ReadINI("kill", "kill5", Environ$("AppData") & "\KidsLock.ini"))
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
isButton1_Click
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Check1.Value = 1 Then
ProtectProcess
Cancel = 1
Else
Cancel = 0
 RestoreProcess '윈도우 중요 프로세스 해재 정상 종료 가능 하게 만듬
 AlwaysTop FormLock, False
End If
End Sub

Private Sub Timer2_Timer()
FormLock.Show
End Sub

Private Sub Timer3_Timer()
Label1.Visible = True
Label1.Caption = "5초후 자동으로 부모모드로 전환됩니다."
FormLoad.Show
Unload Me
End Sub
