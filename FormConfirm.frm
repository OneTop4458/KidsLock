VERSION 5.00
Begin VB.Form FormConfirm 
   BorderStyle     =   0  '����
   Caption         =   "Form1"
   ClientHeight    =   960
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   ScaleHeight     =   960
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '������ ���
   Begin KidsLock.SKin SKin1 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   1720
      Caption         =   "������ ��й�ȣ Ȯ��"
      Begin VB.TextBox Text1 
         Height          =   390
         IMEMode         =   3  '��� ����
         Left            =   240
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   480
         Width           =   4095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Ȯ��"
         Height          =   375
         Left            =   4560
         TabIndex        =   1
         Top             =   480
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FormConfirm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Shadow As clsShadow
Dim md5Test As MD5 'md5 ��ȣȭ ����

Private Sub Command1_Click()
On Error Resume Next
Text1.Text = LCase(md5Test.DigestStrToHexStr(Text1.Text))
If GetSetting("KidLock", "admin", "password") = Text1.Text Then
MsgBox "������ ���� ����!", vbDefaultButton1, "����!"

If ReadINI("confirm", "show", Environ$("AppData") & "\KidsLock.ini") = 1 Then
ShowF FormChange
Else
End If
If ReadINI("confirm", "show", Environ$("AppData") & "\KidsLock.ini") = 2 Then
ShowF FormKidsSetting
Else
End If

Else
MsgBox "���� ������ ��й�ȣ�� ��ġ���� �ʽ��ϴ�!", vbCritical, "����!"
Text1.Text = vbNullString
End If
End Sub

Private Sub Form_Load()
Set md5Test = New MD5 'md5 ����
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

