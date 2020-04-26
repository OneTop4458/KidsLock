VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form Formps 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '단일 고정
   Caption         =   "Task Manager"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5655
   Icon            =   "Formps.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   5655
   StartUpPosition =   1  '소유자 가운데
   Begin VB.CommandButton Command2 
      Caption         =   "등록된 프로세스 일괄 제거"
      Height          =   300
      Left            =   2640
      TabIndex        =   7
      Top             =   6120
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      Caption         =   "등록된 프로세스 목록 (추가항목 있을시 새고고침 필요)"
      Height          =   1455
      Left            =   240
      TabIndex        =   5
      Top             =   4920
      Width           =   5055
      Begin VB.ListBox List1 
         Height          =   960
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   4575
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5040
      Top             =   4440
   End
   Begin VB.CommandButton Command1 
      Caption         =   "선택한 프로세스 자녀보호시 금지"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   3960
      Width           =   3375
   End
   Begin MSComctlLib.ListView lvProcess 
      Height          =   3975
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   7011
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "새로고침"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "0"
      Height          =   255
      Left            =   3480
      TabIndex        =   4
      Top             =   4440
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "현재 등록된 프로그램 수 (자동 동기화) :"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   4440
      Width           =   3375
   End
End
Attribute VB_Name = "Formps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PName As String, PID As Long
'// API  상수선언 (경우에따라서 GetCommandLine , GetModuleFIleName 도 필요)
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWDEFAULT = 10



Private Sub cmdRefresh_Click()
Dim Process
Dim lv As ListItem
On Error Resume Next
List1.Clear
List1.AddItem ReadINI("kill", "default1", Environ$("AppData") & "\KidsLock.ini")
List1.AddItem ReadINI("kill", "default2", Environ$("AppData") & "\KidsLock.ini")
List1.AddItem ReadINI("kill", "kill1", Environ$("AppData") & "\KidsLock.ini")
List1.AddItem ReadINI("kill", "kill2", Environ$("AppData") & "\KidsLock.ini")
List1.AddItem ReadINI("kill", "kill3", Environ$("AppData") & "\KidsLock.ini")
List1.AddItem ReadINI("kill", "kill4", Environ$("AppData") & "\KidsLock.ini")
List1.AddItem ReadINI("kill", "kill5", Environ$("AppData") & "\KidsLock.ini")


lvProcess.ListItems.Clear

For Each Process In GetObject("winmgmts:"). _
    ExecQuery("select * from Win32_Process")
    
    Set lv = lvProcess.ListItems.Add(, , Process.Name)
    lv.SubItems(1) = Process.ProcessID
   
Next



End Sub

Private Sub Command1_Click()
On Error Resume Next
Dim no As String
no = 0
Dim no2 As String
no2 = 0

no = ReadINI("killsetting", "no", Environ$("AppData") & "\KidsLock.ini") '처음 설치시 no 값은 1
If no = 1 Then 'no가 1일때는  등록 가능
Call WriteINI("kill", "kill" & no, PName, Environ$("AppData") & "\KidsLock.ini")
MsgBox "선택한 프로세스" & PName & "이 정상적으로 등록되었습니다", vbDefaultButton1, "등록완료"
Else
no2 = no + 1 'no2 에 현재 no 값에 1을 추가함
Call WriteINI("killsetting", "no", no2, Environ$("AppData") & "\KidsLock.ini") 'ini no 부분에 no2 값으로 변경
End If

If no = 2 Then 'no가 2일때는  등록 가능
Call WriteINI("kill", "kill" & no, PName, Environ$("AppData") & "\KidsLock.ini")
MsgBox "선택한 프로세스" & PName & "이 정상적으로 등록되었습니다", vbDefaultButton1, "등록완료"
Else
no2 = no + 1 'no2 에 현재 no 값에 1을 추가함
Call WriteINI("killsetting", "no", no2, Environ$("AppData") & "\KidsLock.ini") 'ini no 부분에 no2 값으로 변경
End If

If no = 3 Then 'no가 3일때는  등록 가능
Call WriteINI("kill", "kill" & no, PName, Environ$("AppData") & "\KidsLock.ini")
MsgBox "선택한 프로세스" & PName & "이 정상적으로 등록되었습니다", vbDefaultButton1, "등록완료"
Else
no2 = no + 1 'no2 에 현재 no 값에 1을 추가함
Call WriteINI("killsetting", "no", no2, Environ$("AppData") & "\KidsLock.ini") 'ini no 부분에 no2 값으로 변경
End If

If no = 4 Then 'no가 4일때는  등록 가능
Call WriteINI("kill", "kill" & no, PName, Environ$("AppData") & "\KidsLock.ini")
MsgBox "선택한 프로세스" & PName & "이 정상적으로 등록되었습니다", vbDefaultButton1, "등록완료"
Else
no2 = no + 1 'no2 에 현재 no 값에 1을 추가함
Call WriteINI("killsetting", "no", no2, Environ$("AppData") & "\KidsLock.ini") 'ini no 부분에 no2 값으로 변경
End If

If no = 5 Then 'no가 5일때는  등록 가능
Call WriteINI("kill", "kill" & no, PName, Environ$("AppData") & "\KidsLock.ini")
MsgBox "선택한 프로세스" & PName & "이 정상적으로 등록되었습니다", vbDefaultButton1, "등록완료"
Else
no2 = no + 1 'no2 에 현재 no 값에 1을 추가함
Call WriteINI("killsetting", "no", no2, Environ$("AppData") & "\KidsLock.ini") 'ini no 부분에 no2 값으로 변경
End If

If no2 >= 7 Then
MsgBox "등록 가능한 프로세스가 초과되었습니다! ", vbCritical, "등록실패"
Else
End If

End Sub

Private Sub Command2_Click()
Call WriteINI("kill", "kill1", "", Environ$("AppData") & "\KidsLock.ini")
Call WriteINI("kill", "kill2", "", Environ$("AppData") & "\KidsLock.ini")
Call WriteINI("kill", "kill3", "", Environ$("AppData") & "\KidsLock.ini")
Call WriteINI("kill", "kill4", "", Environ$("AppData") & "\KidsLock.ini")
Call WriteINI("kill", "kill5", "", Environ$("AppData") & "\KidsLock.ini")
Call WriteINI("killsetting", "no", "1", Environ$("AppData") & "\KidsLock.ini")
End Sub

Private Sub Form_Load()

With lvProcess.ColumnHeaders

    .Add , , "프로세스", 3900
    .Add , , "프로세스 ID", 1000
    
End With

Call cmdRefresh_Click

End Sub



Private Sub lvProcess_ItemClick(ByVal Item As MSComctlLib.ListItem)

PName = Item.Text
PID = Item.SubItems(1)


End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Dim ge As Integer
ge = ReadINI("killsetting", "no", Environ$("AppData") & "\KidsLock.ini")
Label2.Caption = ge - 1 & "개"

Dim ge2 As Integer
ge2 = ge

If ge2 >= 6 Then
Label2.Caption = "5개" & " " & "※등록 가능 횟수 초과"
Else
End If
End Sub


