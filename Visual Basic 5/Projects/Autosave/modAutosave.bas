Attribute VB_Name = "Various"
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function OpenIcon Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

' Constants for Screen Metrics
Private Const SM_CXFULLSCREEN = 16
Private Const SM_CYFULLSCREEN = 17

' Constants for Window Position
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const FLAGS = SWP_NOMOVE & SWP_NOSIZE
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Public Sub CenterForm(frm As Form)
    ' Centers a form on the screen taking the taskbar into consideration.
    ' Basically what this does is take the full screen size excluding the taskbar
    ' and centers your form accordingly.
    
    Dim Left As Long, Top As Long
    
    Left = (Screen.TwipsPerPixelX * (GetSystemMetrics(SM_CXFULLSCREEN) / 2)) - _
        (frm.Width / 2)
    Top = (Screen.TwipsPerPixelY * (GetSystemMetrics(SM_CYFULLSCREEN) / 2)) - _
        (frm.Height / 2)
    frm.Move Left, Top
End Sub

Public Sub AlwaysOnTop(f As Form, pos As Boolean)

    If pos = True Then
        SetWindowPos frmMain.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS
        ' SetWindowPos takes a hWnd of a window, and moves it to a specific
        ' screen location (specified by the 3rd through 6th parameters) and
        ' changes the Z-Order as requested.
    Else
        SetWindowPos frmMain.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS
    End If
    
End Sub
