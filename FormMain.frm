VERSION 5.00
Begin VB.Form FormMain 
   BorderStyle     =   0  '없음
   Caption         =   "Tray"
   ClientHeight    =   5805
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   10320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.PictureBox Picture1 
      Height          =   6015
      Left            =   -120
      Picture         =   "FormMain.frx":0000
      ScaleHeight     =   5955
      ScaleWidth      =   10515
      TabIndex        =   0
      Top             =   -120
      Width           =   10575
      Begin VB.Timer Timer4 
         Enabled         =   0   'False
         Interval        =   3000
         Left            =   9720
         Top             =   4680
      End
      Begin VB.Timer Timer3 
         Interval        =   3000
         Left            =   9720
         Top             =   4200
      End
      Begin VB.Timer Timer2 
         Left            =   9720
         Top             =   3720
      End
      Begin VB.Timer Timer1 
         Left            =   9720
         Top             =   3240
      End
   End
End
Attribute VB_Name = "FormMain"
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
On Error Resume Next
If ReadINI("mj", "lv", Environ$("AppData") & "\KidsLock.ini") = True Then
Timer3.Enabled = False
Timer4.Enabled = True
Else
Timer3.Enabled = True
Timer4.Enabled = False
End If

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
        FormLoad.Show
        Unload Me
        End If
End Sub

Private Sub Timer4_Timer()
    If exitable = False Then
        Timer2.Enabled = True
        FormKidsTry.Show
        Unload Me
        End If
End Sub
