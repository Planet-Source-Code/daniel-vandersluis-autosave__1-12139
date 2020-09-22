Attribute VB_Name = "Enumeration"
Option Explicit

Public Type ProcData
    AppHwnd As Long
    Title As String
    Placement As String
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public ProcInfo() As ProcData
Public NumProcs As Integer
Public NID As NOTIFYICONDATA

Private Type POINTAPI
    x As Long
    y As Long
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type WINDOWPLACEMENT
    length As Long
    FLAGS As Long
    showCmd As Long
    ptMinPosition As POINTAPI
    ptMaxPosition As POINTAPI
    rcNormalPosition As RECT
End Type
Private Declare Function GetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long


Public Function EnumProc(ByVal app_hwnd As Long, ByVal lParam As Long) As Boolean
Const SW_SHOWNORMAL = 1
Const SW_SHOWMINIMIZED = 2
Const SW_SHOWMAXIMIZED = 3

Dim wp As WINDOWPLACEMENT
Dim buf As String * 256
Dim length As Long

    NumProcs = NumProcs + 1
    ReDim Preserve ProcInfo(1 To NumProcs)
    With ProcInfo(NumProcs)
        .AppHwnd = app_hwnd

        ' Get the window's title.
        length = GetWindowText(app_hwnd, buf, Len(buf))
        If length > 20 Then length = 20
        .Title = Left$(buf, length) & " "

        ' See where the window is.
        GetWindowPlacement app_hwnd, wp
        Select Case wp.showCmd
            Case SW_SHOWNORMAL
                .Placement = "Normal"
            Case SW_SHOWMINIMIZED
                .Placement = "Minimized"
            Case SW_SHOWMAXIMIZED
                .Placement = "Maximized"
        End Select

        .Left = wp.rcNormalPosition.Left
        .Top = wp.rcNormalPosition.Top
        .Right = wp.rcNormalPosition.Right
        .Bottom = wp.rcNormalPosition.Bottom
    End With
    
    ' Continue searching.
    EnumProc = 1
End Function
