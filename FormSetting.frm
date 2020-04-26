VERSION 5.00
Begin VB.Form FormSetting 
   BorderStyle     =   0  '없음
   Caption         =   "Form1"
   ClientHeight    =   2250
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4860
   LinkTopic       =   "Form1"
   ScaleHeight     =   2250
   ScaleWidth      =   4860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin KidsLock.SKin SKin1 
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   4048
      Caption         =   "자녀보호 설정"
      Begin KidsLock.CandyButton CandyButton4 
         Height          =   495
         Left            =   3000
         TabIndex        =   4
         Top             =   960
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "자녀모드 전환"
         IconHighLiteColor=   0
         CaptionHighLiteColor=   0
         Style           =   2
         Checked         =   0   'False
         ColorButtonHover=   255
         ColorButtonUp   =   192
         ColorButtonDown =   8421631
         BorderBrightness=   0
         ColorBright     =   12632319
         DisplayHand     =   0   'False
         ColorScheme     =   0
      End
      Begin KidsLock.CandyButton CandyButton3 
         Height          =   255
         Left            =   3120
         TabIndex        =   3
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "donation"
         IconHighLiteColor=   0
         CaptionHighLiteColor=   0
         Style           =   2
         Checked         =   0   'False
         ColorButtonHover=   65280
         ColorButtonUp   =   49152
         ColorButtonDown =   8454016
         BorderBrightness=   0
         ColorBright     =   12648384
         DisplayHand     =   0   'False
         ColorScheme     =   0
      End
      Begin KidsLock.CandyButton CandyButton2 
         Height          =   495
         Left            =   360
         TabIndex        =   2
         Top             =   1320
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "프로그램 설정"
         IconHighLiteColor=   0
         CaptionHighLiteColor=   0
         Checked         =   0   'False
         ColorButtonHover=   16760976
         ColorButtonUp   =   15309136
         ColorButtonDown =   15309136
         BorderBrightness=   0
         ColorBright     =   16772528
         DisplayHand     =   0   'False
         ColorScheme     =   0
      End
      Begin KidsLock.CandyButton CandyButton1 
         Height          =   495
         Left            =   360
         TabIndex        =   1
         Top             =   600
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "관리자 비밀번호 설정"
         IconHighLiteColor=   0
         CaptionHighLiteColor=   0
         Checked         =   0   'False
         ColorButtonHover=   16760976
         ColorButtonUp   =   15309136
         ColorButtonDown =   15309136
         BorderBrightness=   0
         ColorBright     =   16772528
         DisplayHand     =   0   'False
         ColorScheme     =   0
      End
   End
End
Attribute VB_Name = "FormSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// API  상수선언 (경우에따라서 GetCommandLine , GetModuleFIleName 도 필요)
Dim Shadow As clsShadow
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWDEFAULT = 10
Private Sub CandyButton1_Click()
FormSetting.Hide
Call WriteINI("confirm", "show", "1", Environ$("AppData") & "\KidsLock.ini")
FormConfirm.Show
End Sub

Private Sub CandyButton2_Click()
FormSetting.Hide
Call WriteINI("confirm", "show", "2", Environ$("AppData") & "\KidsLock.ini")
FormConfirm.Show
End Sub

Private Sub CandyButton3_Click()
If MsgBox("제작자를 위해 후원해주시겠습니까?", vbInformation + vbYesNo, "후원하기") = vbYes Then
MsgBox "기업은행 048-072153-01-023 이병준"
MsgBox "커피 한잔 값 이라도 감사한 마음으로 받겠습니다! 감사합니다"
End If
End Sub

Private Sub CandyButton4_Click()

Unload Me
Unload FormLoad
Unload FormMain
Unload Formps
Unload FormSetting
Unload FormChange
Unload FormAbout
Unload FormConfirm

FormKidsTry.Show
End Sub
Private Sub Form_Load()
Set Shadow = New clsShadow
Call Shadow.Shadow(Me)
Shadow.Color = vbBlack
Shadow.Depth = 5
End Sub
