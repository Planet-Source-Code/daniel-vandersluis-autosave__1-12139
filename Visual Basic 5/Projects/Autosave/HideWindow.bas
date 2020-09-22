Attribute VB_Name = "Hide"
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Public hwnd As Long
Public Function HideWnd(hwnd)
    ' Makes a window invisible and hides it from the taskbar and ALT+TAB list
    HideWnd = ShowWindow(hwnd, 0)

End Function
Public Function ShowWnd(hwnd)
    ' Returns window to taskbar, shows window
    ShowWnd = ShowWindow(hwnd, 1)
    
End Function
