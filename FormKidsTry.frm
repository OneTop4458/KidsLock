VERSION 5.00
Begin VB.Form FormKidsTry 
   BorderStyle     =   0  '없음
   Caption         =   "자녀보호모드 동작중..."
   ClientHeight    =   5745
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   10185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.PictureBox Picture1 
      Height          =   5895
      Left            =   -120
      Picture         =   "FormKidsTry.frx":0000
      ScaleHeight     =   5835
      ScaleWidth      =   10275
      TabIndex        =   0
      Top             =   -120
      Width           =   10335
      Begin VB.Timer Timer3 
         Interval        =   3000
         Left            =   240
         Top             =   3960
      End
      Begin VB.Timer Timer2 
         Left            =   240
         Top             =   3480
      End
      Begin VB.Timer Timer1 
         Left            =   240
         Top             =   3000
      End
   End
End
Attribute VB_Name = "FormKidsTry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_ALPHA = &H2&

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Private Const fade As Integer = 8
Dim Alpha As Integer
Dim exitable As Boolean

Private Sub Form_Load()
    Call SetWindowLong(Me.hWnd, GWL_EXSTYLE, GetWindowLong(Me.hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
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
    Call SetLayeredWindowAttributes(Me.hWnd, 0, Alpha, LWA_ALPHA)
End Sub

Private Sub Timer2_Timer()
    Alpha = Alpha - fade
    If Alpha <= 0 Then
        Alpha = 0
        Timer2.Enabled = False
        exitable = True
    End If
    Call SetLayeredWindowAttributes(Me.hWnd, 0, Alpha, LWA_ALPHA)
    Unload Me
End Sub

Private Sub Timer3_Timer()
    If exitable = False Then
        Timer2.Enabled = True
        Unload FormSetting
        FormKidsLoad.Show
        End If
End Sub

