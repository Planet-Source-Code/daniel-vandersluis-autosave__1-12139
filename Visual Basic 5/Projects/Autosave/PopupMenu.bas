Attribute VB_Name = "IconMenu"
' Menu Functions
Declare Function TrackPopupMenuEx& Lib "user32" (ByVal hMenu As Long, ByVal un As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal hwnd As Long, lpTPMParams As TPMPARAMS)
Declare Function InsertMenu& Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String)
Declare Function GetMenu& Lib "user32" (ByVal hwnd As Long)
Declare Function CreateMenu& Lib "user32" ()
Declare Function AppendMenu& Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String)
Declare Function CreatePopupMenu& Lib "user32" ()
Declare Function GetCursorPos& Lib "user32" (lpPoint As POINTAPI)
Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function CheckMenuRadioItem Lib "user32" (ByVal hMenu As Long, ByVal un1 As Long, ByVal un2 As Long, ByVal un3 As Long, ByVal un4 As Long) As Long
Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Declare Function FindWindow Lib "user32" (ByVal lpClassName As Any, ByVal lpWindowName As Any)
Declare Function CheckMenuItem Lib "user32" (ByVal hMenu%, ByVal wIdCheckItem%, ByVal Check%)
Declare Function EnableMenuItem Lib "user32" (ByVal hMenu%, ByVal wIdEnableItem%, ByVal Enable%)
Declare Function GetSubMenu Lib "user32" (ByVal hMenu%, ByVal nPos%)
Declare Function GetMenuItemID& Lib "user32" (ByVal hMenu&, ByVal nPos&)
Declare Function ModifyMenu& Lib "user32" Alias "ModifyMenuA" (ByVal hMenu&, ByVal nPosition&, ByVal wFlags&, ByVal wIDNewItem&, ByVal lpString$)

' Constants for menu position
Public Const MF_BYPOSITION = &H400&
Public Const MF_POPUP = &H10&
Public Const MF_STRING = &H0&
Public Const MF_MENUBARBREAK As Long = &H20&

' Constants to check or uncheck a menu item
Public Const MF_CHECKED = &H8&
Public Const MF_UNCHECKED = &H0&

' Constants dealing with systray icons
Public Const NIM_ADD = &H0
Public Const NIM_DELETE = &H2
Public Const NIM_MODIFY = &H1
Public Const NIF_ICON = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_TIP = &H4

Public Const TPM_LEFTBUTTON = &H0
Public Const TPM_RIGHTBUTTON = &H2
Public Const TPM_LEFTALIGN = &H0
Public Const TPM_CENTERALIGN = &H4
Public Const TPM_RIGHTALIGN = &H8
Public Const TPM_BOTTOMALIGN = &H20&
Public Const TPM_HORIZONTAL = &H0&
Public Const TPM_NONOTIFY = &H80&
Public Const TPM_RETURNCMD = &H100&
Public Const TPM_TOPALIGN = &H0&
Public Const TPM_VCENTERALIGN = &H10&
Public Const TPM_VERTICAL = &H40&

' Windows Mouse Constants
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_RBUTTONUP = &H205

Type NOTIFYICONDATA
        cbSize As Long
        hwnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type

Type RECT
    Top As Long
    Left As Long
    Bottom As Long
    Right As Long
End Type

Type TPMPARAMS
    cbSize As Long
    rcExclude As RECT
End Type

Type POINTAPI
    x As Long
    y As Long
End Type
Public Sub SetTrayTip(tip As String)
' Set the systray icon tip.
    With NID
        .szTip = tip & vbNullChar
        .uFlags = .uFlags Or NIF_TIP
    End With
    Shell_NotifyIcon NIM_MODIFY, NID
End Sub
