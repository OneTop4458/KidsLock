Attribute VB_Name = "Top"
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, Y, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
Public Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Public Function AlwaysTop(TheForm As Form, Use As Boolean)
     If Use = True Then
          SetWindowPos TheForm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
     Else
          SetWindowPos TheForm.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
     End If
End Function
