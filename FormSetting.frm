VERSION 5.00
Begin VB.Form FormSetting 
   BorderStyle     =   0  '����
   Caption         =   "Form1"
   ClientHeight    =   2250
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4860
   LinkTopic       =   "Form1"
   ScaleHeight     =   2250
   ScaleWidth      =   4860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '������ ���
   Begin KidsLock.SKin SKin1 
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   4048
      Caption         =   "�ڳຸȣ ����"
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
         Caption         =   "�ڳ��� ��ȯ"
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
         Caption         =   "���α׷� ����"
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
         Caption         =   "������ ��й�ȣ ����"
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
'// API  ������� (��쿡���� GetCommandLine , GetModuleFIleName �� �ʿ�)
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
If MsgBox("�����ڸ� ���� �Ŀ����ֽðڽ��ϱ�?", vbInformation + vbYesNo, "�Ŀ��ϱ�") = vbYes Then
MsgBox "������� 048-072153-01-023 �̺���"
MsgBox "Ŀ�� ���� �� �̶� ������ �������� �ްڽ��ϴ�! �����մϴ�"
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
