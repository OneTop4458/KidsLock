VERSION 5.00
Begin VB.Form FormChange 
   BorderStyle     =   0  '없음
   Caption         =   "Form1"
   ClientHeight    =   1935
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   ScaleHeight     =   1935
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '사용 못함
      Left            =   3240
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   840
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '사용 못함
      Left            =   3240
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1320
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "변경"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5880
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin KidsLock.SKin SKin1 
      Height          =   1935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   3413
      Caption         =   "관리자 비밀번호 설정"
      Begin VB.Label Label4 
         BackStyle       =   0  '투명
         Caption         =   "※주의 : 대소문자 특부문자 구분"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label Label3 
         BackStyle       =   0  '투명
         Caption         =   "새 미빌번호 확인"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   15.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "새 관리자 비밀번호"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   15.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   2895
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '투명
         Caption         =   "보안등급:0"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   6240
         TabIndex        =   4
         Top             =   480
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FormChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Shadow As clsShadow
Dim md5Test As MD5 'md5 암호화 선언
Private Type KeyboardBytes
     kbByte(0 To 255) As Byte
End Type
Dim LCd(1 To 3) As Byte
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare Function GetKeyboardState Lib "user32" (kbArray As KeyboardBytes) As Long
Private Sub Command1_Click()
On Error Resume Next

If Text.Text = Text1.Text Then
Text.Text = LCase(md5Test.DigestStrToHexStr(Text.Text))
SaveSetting "KidLock", "admin", "password", Text.Text
MsgBox "관리자 비밀번호 변경완료!", vbDefaultButton1, "변경 완료"
FormSetting.Show
Unload Me
Else
MsgBox "새 관리자 비밀번호와 비밀번호 확인이 일치하지 않습니다!", vbCritical, "오류!"
Text.Text = vbNullString
Text1.Text = vbNullString
End If
End Sub

Private Sub Form_Load()
Set md5Test = New MD5 'md5 선언
Set Shadow = New clsShadow
Call Shadow.Shadow(Me)
Shadow.Color = vbBlack
Shadow.Depth = 5
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command1_Click
End If
End Sub
Private Sub Text_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command1_Click
End If
End Sub

Private Sub Text_KeyUp(KeyCode As Integer, Shift As Integer)
Call StrChk(Text)
End Sub


Function StrChk(Str As String) As Integer
    Dim i As Integer
    Dim Chk(1 To 5) As Boolean
    
    If Len(Str) > 7 Then Chk(1) = True
    
    For i = 1 To Len(Str)
        If Asc(Mid(Str, i, 1)) >= 97 And Asc(Mid(Str, i, 1)) <= 122 Then Chk(2) = True
        If Asc(Mid(Str, i, 1)) >= 65 And Asc(Mid(Str, i, 1)) <= 90 Then Chk(3) = True
        If Asc(Mid(Str, i, 1)) >= 48 And Asc(Mid(Str, i, 1)) <= 57 Then Chk(4) = True
        If Asc(Mid(Str, i, 1)) < 48 Or Asc(Mid(Str, i, 1)) > 57 And Asc(Mid(Str, i, 1)) < 65 Or Asc(Mid(Str, i, 1)) > 90 And Asc(Mid(Str, i, 1)) < 97 Or Asc(Mid(Str, i, 1)) > 122 Then Chk(5) = True
    Next
    
    For i = 1 To 5
        If Chk(i) = True Then StrChk = StrChk + 1
    Next
    Label2 = "보안등급:" & StrChk
End Function
