VERSION 5.00
Begin VB.Form FormKidsSetting 
   BackColor       =   &H80000010&
   BorderStyle     =   4  '���� ���� â
   Caption         =   "Kids Lock ����"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '������ ���
   Begin VB.Frame Frame1 
      BackColor       =   &H80000010&
      Caption         =   "���� ���۸��"
      Height          =   735
      Left            =   1920
      TabIndex        =   14
      Top             =   480
      Width           =   2535
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   2040
         Top             =   240
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '����
         Caption         =   "���� ���:"
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   2055
      End
   End
   Begin KidsLock.CandyButton CandyButton4 
      Height          =   615
      Left            =   2520
      TabIndex        =   13
      Top             =   3000
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "������ ���۽� �θ���"
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
   Begin KidsLock.CandyButton CandyButton3 
      Height          =   615
      Left            =   360
      TabIndex        =   12
      Top             =   3000
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "������ ���۽� �ڳ���"
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
   Begin KidsLock.CandyButton CandyButton2 
      Height          =   375
      Left            =   3840
      TabIndex        =   11
      Top             =   0
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "��� ���"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Checked         =   0   'False
      ColorButtonHover=   14704640
      ColorButtonUp   =   13668448
      ColorButtonDown =   11108432
      BorderBrightness=   0
      ColorBright     =   16775930
      DisplayHand     =   0   'False
      ColorScheme     =   2
   End
   Begin KidsLock.CandyButton CandyButton1 
      Height          =   615
      Left            =   240
      TabIndex        =   10
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "���� ����"
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
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   8
      Top             =   4680
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   6
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H80000010&
      Caption         =   "�������α׷� ���"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H80000010&
      Caption         =   "�⺻ ����"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Value           =   1  'Ȯ��
      Width           =   1335
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '����
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   9
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '����
      Caption         =   "�ð�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Top             =   4800
      Width           =   735
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '����
      Caption         =   "(�ڳ��� PC ��� ���ɽð��� �����մϴ�.)"
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   4320
      Width           =   3615
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '����
      Caption         =   "-�ڳ� PC ��� �ð� ���� - /�⺻ ���� : 99�ð� 99��"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   4080
      Width           =   4455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '����
      Caption         =   "Kid Lock �� ������ ���۽� �ڵ� ����/���� �մϴ�"
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   2280
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '����
      Caption         =   "�⺻������ üũ �Ǿ��ֽ��ϴ� �۾�������,cmd �� �����մϴ�."
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   1440
      Width           =   2655
   End
End
Attribute VB_Name = "FormKidsSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// API  ������� (��쿡���� GetCommandLine , GetModuleFIleName �� �ʿ�)
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWDEFAULT = 10

Private Sub CandyButton1_Click()
'���� ����
Call WriteINI("setting", "hour", Text1.Text, Environ$("AppData") & "\KidsLock.ini")
Call WriteINI("setting", "min", Text2.Text, Environ$("AppData") & "\KidsLock.ini")
Call WriteINI("setting", "default", Check1.Value, Environ$("AppData") & "\KidsLock.ini")
Call WriteINI("setting", "winstart", Check2.Value, Environ$("AppData") & "\KidsLock.ini")
MsgBox "���������� �����Ͽ����ϴ�", vbDefaultButton1, "����Ϸ�!"
Unload Me
FormSetting.Show
End Sub

Private Sub CandyButton2_Click()
MsgBox "�� ���� �ڳຸȣ ������ ����Ǵ� ���α׷��� �����Ҽ� �ִ� ����Դϴ�" & vbCrLf & "�ִ� 5���� ���α׷��� �����Ҽ��ֽ��ϴ�.", vbDefaultButton1, "�˸�"
Formps.Show
End Sub

Private Sub CandyButton3_Click()
Call WriteINI("mj", "lv", "True", Environ$("AppData") & "\KidsLock.ini")
End Sub

Private Sub CandyButton4_Click()
Call WriteINI("mj", "lv", "False", Environ$("AppData") & "\KidsLock.ini")
End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then
Call WriteINI("kill", "default1", "cmd.exe", Environ$("AppData") & "\KidsLock.ini")
Call WriteINI("kill", "default2", "Taskmgr.exe", Environ$("AppData") & "\KidsLock.ini")
Else
Call WriteINI("kill", "default1", "", Environ$("AppData") & "\KidsLock.ini")
Call WriteINI("kill", "default2", "", Environ$("AppData") & "\KidsLock.ini")
End If
End Sub

Private Sub Check2_Click()
'üũ 3 ����
Dim Path As String
Dim juso As String
juso = Environ$("ProgramFiles") & "\KidsLock\KidsLock.exe"
Path = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\"

If Check2.Value = 0 Then
On Error Resume Next '������ ���� ����
Set WshShell = CreateObject("WScript.Shell")
WshShell.RegDelete Path & "KidsLock"
End If

If Check2.Value = 1 Then
Set WshShell = CreateObject("WScript.Shell")
WshShell.RegWrite Path & "KidsLock", juso, "REG_SZ"

End If
End Sub

Private Sub Form_Load()
On Error Resume Next
'���� �ҷ�����
Text1.Text = ReadINI("setting", "hour", Environ$("AppData") & "\KidsLock.ini")
Text2.Text = ReadINI("setting", "min", Environ$("AppData") & "\KidsLock.ini")

If (ReadINI("setting", "default", Environ$("AppData") & "\KidsLock.ini")) = 1 Then
Check1.Value = 1
Else
Check1.Value = 0
End If

If (ReadINI("setting", "winstart", Environ$("AppData") & "\KidsLock.ini")) = 0 Then
Check2.Value = 0
Else
Check2.Value = 1
End If

End Sub



Private Sub Timer1_Timer()
On Error Resume Next
If ReadINI("mj", "lv", Environ$("AppData") & "\KidsLock.ini") = True Then '�ڳ����Ͻ�
Label2.Caption = "������: ���۽� �ڳ���� ����"
Else
Label2.Caption = "������: ���۽� �θ���� ����"
End If
End Sub
