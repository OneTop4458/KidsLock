VERSION 5.00
Begin VB.Form FormAbout 
   BorderStyle     =   0  '없음
   Caption         =   "About"
   ClientHeight    =   5745
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10275
   LinkTopic       =   "Form3"
   ScaleHeight     =   5745
   ScaleWidth      =   10275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.PictureBox Picture1 
      Height          =   5895
      Left            =   -120
      Picture         =   "FormAbout.frx":0000
      ScaleHeight     =   5835
      ScaleWidth      =   10395
      TabIndex        =   0
      Top             =   -120
      Width           =   10455
      Begin VB.Timer Timer3 
         Interval        =   3000
         Left            =   480
         Top             =   4200
      End
      Begin VB.Timer Timer2 
         Left            =   480
         Top             =   3720
      End
      Begin VB.Timer Timer1 
         Left            =   480
         Top             =   3240
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '투명
         Caption         =   "프로그램 버전 : beta3895"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7080
         TabIndex        =   2
         Top             =   2760
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "special thanks to : 평화님 (폼스킨 제공)"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         TabIndex        =   1
         Top             =   5280
         Width           =   3855
      End
   End
End
Attribute VB_Name = "FormAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_ALPHA = &H2&

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Private Const fade As Integer = 8
Dim Alpha As Integer
Dim exitable As Boolean

Private Sub Form_Load()
    Call SetWindowLong(Me.hwnd, GWL_EXSTYLE, GetWindowLong(Me.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
    Alpha = 0
    Timer1.Enabled = True
    Timer1.Interval = 10
    Call Timer1_Timer
    
    Timer2.Enabled = False
    Timer2.Interval = 10
    exitable = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If exitable = False Then
        Timer2.Enabled = True
        Cancel = True
    End If
End Sub


Private Sub Timer1_Timer()
    Alpha = Alpha + fade
    If Alpha >= 255 Then
        Alpha = 255
        Timer1.Enabled = False
    End If
    Call SetLayeredWindowAttributes(Me.hwnd, 0, Alpha, LWA_ALPHA)
End Sub

Private Sub Timer2_Timer()
    Alpha = Alpha - fade
    If Alpha <= 0 Then
        Alpha = 0
        Timer2.Enabled = False
        exitable = True
    End If
    Call SetLayeredWindowAttributes(Me.hwnd, 0, Alpha, LWA_ALPHA)
    Unload Me
End Sub

Private Sub Timer3_Timer()
    If exitable = False Then
        Timer2.Enabled = True
        End If
End Sub

