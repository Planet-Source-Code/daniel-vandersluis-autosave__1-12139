VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AutoSave"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6570
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   6570
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List2 
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Timer tmrRefresh 
      Interval        =   10
      Left            =   0
      Top             =   1080
   End
   Begin VB.CommandButton cmdSet 
      Caption         =   "&Set Timer"
      Height          =   495
      Left            =   3840
      TabIndex        =   5
      Top             =   4800
      Width           =   1215
   End
   Begin VB.TextBox txtInterval 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Text            =   "0"
      Top             =   4890
      Width           =   1215
   End
   Begin VB.CheckBox chkCurrent 
      Caption         =   "Autosave in Current Window"
      Height          =   195
      Left            =   360
      TabIndex        =   7
      Top             =   5760
      Width           =   2415
   End
   Begin VB.CheckBox chkSaveAs 
      Caption         =   "Activate Save As Dialog Box"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   5400
      Value           =   1  'Checked
      Width           =   2415
   End
   Begin VB.Timer tmrLoop 
      Interval        =   10
      Left            =   0
      Top             =   600
   End
   Begin VB.Timer tmrWait 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   120
   End
   Begin VB.OptionButton optOff 
      Caption         =   "Off"
      Height          =   495
      Left            =   5520
      TabIndex        =   3
      Top             =   4680
      Value           =   -1  'True
      Width           =   615
   End
   Begin VB.OptionButton optOn 
      Caption         =   "On"
      Height          =   495
      Left            =   5520
      TabIndex        =   2
      Top             =   4200
      Width           =   615
   End
   Begin VB.ListBox List1 
      Height          =   3765
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   5775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Refresh List"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   4200
      Width           =   1815
   End
   Begin VB.Label lblMilliseconds 
      AutoSize        =   -1  'True
      Caption         =   "milliseconds"
      Height          =   195
      Left            =   2760
      TabIndex        =   9
      Top             =   4920
      Width           =   840
   End
   Begin VB.Label lblTimerInterval 
      AutoSize        =   -1  'True
      Caption         =   "Timer Interval:"
      Height          =   195
      Left            =   360
      TabIndex        =   8
      Top             =   4920
      Width           =   1005
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===========================================================================
'AUTOSAVE PROGRAM
'
'Written by:    Daniel Vandersluis
'               E-Mail: dvandersluis215@yahoo.com
'               ICQ:    57887382
'Version:       1.1
'Finished:      16-Oct-2000
'
'Description:   A lot of programs do not automatically save your work
'               and therefore, if something like a crash occurs, all of
'               your ammendments to your file since your last save are lost.
'               This program automatically saves your file for you, saving
'               potentially losing your work.
'
'Programming:   Uses many system API calls, but everything is documented, so
'               don't worry. I tried to comment as much as I could. I'd say that
'               this is pretty Advanced if you would like a difficulty rating.
'
'Distribution:  I don't mind if you use part or all of this code within your
'               own programs, however, I ask that you do not distribute it
'               your own program and give credit where credit is deserved.
'
'Feedback:      I'd appreciate feedback or and comments you might have. Either
'               log them through PSC, email them to me, or ICQ me and I'll respond
'               as soon as I can.
'
'Thanks to:     * Pause Break [mofd4u@yahoo.com] for supplying PSC with his
'                 Registry module
'               * Nick Smith aka ImN0thing for the icon in systray code
'               * Bryan Stafford of New Vision Software® - newvision@imt.net - for
'                 the menu columns code
'               * The folks over at Planet Source Code for their great
'                 website
'
'By the way:    Check out http://www.angelfire.com/on3/infiniti/index.html
'               This is my homepage and it contains all my programs in C/C++ and
'               VB, as well as Descent 3 Levels, and much more, all with open
'               source code!
'============================================================================

Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Any, ByVal lParam As Long) As Long

Public TimerSet As Boolean
Public application As String
Public listNow As Integer
Public listNum As Integer

Private Sub chkCurrent_Click()

    retval = UpdateKey(HKEY_LOCAL_MACHINE, "\Software\iNFiNiTi Studios\Autosave", "Current", chkCurrent.Value)
    'UpdateKey creates/updates a registry entry
    
End Sub
Private Sub chkSaveAs_Click()
    
    retval = UpdateKey(HKEY_LOCAL_MACHINE, "\Software\iNFiNiTi Studios\Autosave", "SaveAs", chkSaveAs.Value)
    'See chkCurrent_Click()
    
End Sub
Private Sub cmdSet_Click()
    
    If txtInterval.Text > 0 Then
        tmrLoop.Interval = txtInterval.Text
        MsgBox "Timer Interval set to " & tmrLoop.Interval, vbInformation + vbOKOnly, "AutoSave"
        TimerSet = True
        ' Update Registry Key
        retval = UpdateKey(HKEY_LOCAL_MACHINE, "\Software\iNFiNiTi Studios\Autosave", "TimerInterval", txtInterval)
    End If

End Sub
Private Sub Command1_Click()

    ' For comments on window enumeration, see Form_Activate()
    
    Erase ProcInfo
    NumProcs = 0
    
    List1.Clear
    EnumWindows AddressOf EnumProc, 0

    tmrWait.Enabled = True
    
End Sub
Private Sub Form_Activate()

    ' Enumerating current windows into a list box.
    
    ' Erase any old information.
    Erase ProcInfo
    NumProcs = 0
    
    List1.Clear
    EnumWindows AddressOf EnumProc, 0

    ' Wait for the enumeration to finish.
    tmrWait.Enabled = True

End Sub
Private Sub Form_Load()
    
    'Hide window from TaskList (CTRL+ALT+DEL List)
    OwnerhWnd = GetWindow(Me.hwnd, 4)   ' Specifies what part of the window to get
                                        ' Here OwnerhWnd is the part of the window
                                        ' for the TaskList
    retval = ShowWindow(OwnerhWnd, 0)   ' Hides the window from the TaskList
    
    'Get Stored Settings from Registry
    Dim tInterval As String
    txtInterval = GetKeyValue(HKEY_LOCAL_MACHINE, "Software\iNFiNiTi Studios\Autosave\", "TimerInterval")
    chkSaveAs.Value = GetKeyValue(HKEY_LOCAL_MACHINE, "Software\iNFiNiTi Studios\Autosave\", "SaveAs")
    chkCurrent.Value = GetKeyValue(HKEY_LOCAL_MACHINE, "Software\iNFiNiTi Studios\Autosave\", "Current")

    'Create Systray icon
    NID.hwnd = Me.hwnd
    NID.cbSize = Len(NID)
    NID.uID = vbNull
    NID.uFlags = NIF_TIP Or NIF_ICON Or NIF_MESSAGE
    NID.hIcon = Me.Icon
    NID.uCallbackMessage = WM_MOUSEMOVE
    SetTrayTip "Autosave"
    
    Shell_NotifyIcon NIM_ADD, NID

    TimerSet = False
    
    listNum = NumProcs
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Dim NID As NOTIFYICONDATA
    
    ' Deletes the systray icon
    NID.hwnd = Me.hwnd
    NID.cbSize = Len(NID)
    NID.uID = vbNull
    NID.uFlags = NIF_TIP Or NIF_ICON Or NIF_MESSAGE
    NID.hIcon = Me.Icon
    NID.uCallbackMessage = WM_MOUSEMOVE
    NID.szTip = "Right-Click to display Popupmenu"
    
    Shell_NotifyIcon NIM_DELETE, NID
    
End Sub
Private Sub Form_Resize()

'If the form is minimized, hide it from the taskbar
    If Me.WindowState = 1 Then
        HideWnd Me.hwnd
    Else ' If Restore is pressed
        ShowWnd Me.hwnd
    End If

End Sub
Private Sub tmrLoop_Timer()
    Dim lRetVal As Long
    Dim bRetVal As Boolean
    Dim myHWnd As Long
    
    If optOn.Value = True Then
        On Error Resume Next
        If chkCurrent.Value = True And chkSaveAs.Value = True Then
            SendKeys ("%FAE")   ' This sends alt-F, then A, then E to the current
                                ' window. I'm sending both A and E to the file
                                ' menu because both are sometimes used as the
                                ' SaveAs accelerator in the file menu.
                                
        ElseIf chkCurrent.Value = True And chkSaveAs.Value = False Then
            SendKeys ("%FS")    ' This time we just save the file, and don't
                                ' confirm file name.
                                
        ElseIf chkSaveAs.Value = True And chkCurrent.Value = False Then
            myHWnd = CLng(List2.List(List1.ListIndex))
            lRetVal = SetForegroundWindow(myHWnd)
            ' SetForegroundWindow takes the equivilant hWnd from an invisible list
            ' and activates that window. However, if the window is minimized, then
            ' it is only activated in the taskbar, and not maximized
            bRetVal = OpenIcon(myHWnd)
            ' Maximizes the window
            SendKeys ("%FAE")
            
        ElseIf chkSaveAs.Value = False And chkCurrent.Value = False Then
            myHWnd = CLng(List2.List(List1.ListIndex))
            lRetVal = SetForegroundWindow(myHWnd)
            bRetVal = OpenIcon(myHWnd)
            SendKeys ("%FS")
        End If
    End If

End Sub
Private Sub tmrRefresh_Timer()
    ' Checks if a list item is selected and if the timer is set
    If TimerSet Or txtInterval.Text > 0 And List1.SelCount <> 0 Then
        optOn.Enabled = True
    Else
        optOn.Enabled = False
    End If
End Sub
Private Sub tmrWait_Timer()
Dim i As Integer
Dim txt As String
Dim txt2 As Long

'Enumerates open (including hidden) windows into a list box
    tmrWait.Enabled = False
      For i = 1 To NumProcs
        With ProcInfo(i)
            'txt = Format$(.Title, "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
            txt = .Title    ' List1 will get the Window Title
            txt2 = .AppHwnd ' List2 will get the Window's hWnd
        End With
        'Dump System and Blank Procs
        If (Left(txt, 1) <> " ") And (Left(txt, 2) <> "MS") And _
            (Left(txt, 3) <> "WIN") And (Left(txt, 2) <> "DD") And _
            (Left(txt, 3) <> "OLE") And (Left(txt, 3) <> "Ole") Then
        ' if a title is not blank, and is not a system window (MS, OLE, etc.) it
        ' is removed from the listbox
            List1.AddItem txt
            List2.AddItem txt2
        Else
            NumProcs = NumProcs - 1 ' Otherwise, the number of Procs is decreased
                                    ' by one
        End If
    Next i
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim Msg As Long
Msg = x / Screen.TwipsPerPixelX
    Select Case Msg
        Case WM_LBUTTONDOWN
        ' if the left mouse button is pressed on the systray icon, the window
        ' is restored
        Me.WindowState = 0
        ShowWnd (Me.hwnd)
        SetForegroundWindow (Me.hwnd)
        
        Case WM_RBUTTONUP
        ' if the right mouse button is pressed on the systray icon, a menu pops up
        Dim pAPI As POINTAPI
        Dim PMParams As TPMPARAMS
        
        'activate the program, this is the code'
        'that gets rid of the bugs of popup'
        'menus hiding behind the taskbar, or'
        'displaying the same time other popup'
        'menus are up'
        'AppActivate Me.Caption
        
        'get point at which the user clicked'
        GetCursorPos pAPI
        
        'create a new popup menu, and add entries'
        'and assign ids to each'
        Dim tmpPop&
        tmpPop& = CreatePopupMenu
        
        Dim ListCounter As Integer
        
        ' Insert all window titles from the list box into the first part of the
        ' popup menu
        
        'This is my old code that didn't include a submenu:
            'For ListCounter = 0 To NumProcs - 1
            '    InsertMenu tmpPop%, 1 + ListCounter, MF_BYPOSITION, 200 + ListCounter, List1.List(ListCounter)
            'Next ListCounter
            
            'InsertMenu tmpPop%, NumProcs + 1, MF_SEPARATOR, 71, vbNullString
            'InsertMenu tmpPop%, NumProcs + 2, MF_BYPOSITION, 72, "Restore"
            'InsertMenu tmpPop%, NumProcs + 3, MF_BYPOSITION, 73, "Exit"
        
        Dim subMenu As Long
        
        subMenu = CreatePopupMenu
        
        Call InsertMenu(tmpPop&, 1&, MF_POPUP Or MF_STRING Or MF_BYPOSITION, subMenu&, "Procs")
        InsertMenu tmpPop&, 2&, MF_SEPARATOR, 71, vbNullString
        InsertMenu tmpPop&, 3&, MF_BYPOSITION, 72, "Restore"
        InsertMenu tmpPop&, 4&, MF_BYPOSITION, 73, "Exit"
             
        For ListCounter = 0 To NumProcs - 1
            Call InsertMenu(subMenu, ListCounter + 0&, MF_STRING Or MF_BYPOSITION, 200& + ListCounter, List1.List(ListCounter))
            
            ' after every 20 menu items, create a new column in the menu:
            If (ListCounter Mod 20 = 0) And ListCounter <> 0 Then
                Call ModifyMenu(subMenu, ListCounter, MF_BYPOSITION Or MF_MENUBARBREAK, ListCounter, List1.List(ListCounter))
            End If
            
        Next ListCounter
        
        'this is a standard size required for
        'the popup menu to be displayed
        PMParams.cbSize = 20
        
        'display the popup menu, note the
        'flag "TMP_RETURNCMD", it sets the value
        'of tmpReply% to the id of the menu item
        'that was clicked
        tmpReply% = TrackPopupMenuEx(tmpPop&, TPM_LEFTALIGN Or TPM_LEFTBUTTON Or TPM_RETURNCMD, pAPI.x, pAPI.y, Me.hwnd, PMParams)
        
        Select Case tmpReply%
            Case 200 To (200 + NumProcs - 1)    ' Select title in list box
                List1.Selected(tmpReply% - 200) = True
                
            Case 72 'Restore
                ShowWnd Me.hwnd
                                
            Case 73 'Exit
                End
        End Select
    End Select
End Sub
Private Sub Form_Terminate()
    
    Dim NID As NOTIFYICONDATA
    
    NID.hwnd = Me.hwnd
    NID.cbSize = Len(NID)
    NID.uID = vbNull
    NID.uFlags = NIF_TIP Or NIF_ICON Or NIF_MESSAGE
    NID.hIcon = Me.Icon
    NID.uCallbackMessage = WM_MOUSEMOVE
    NID.szTip = "Right-Click to display Popupmenu"
    
    Shell_NotifyIcon NIM_DELETE, NID
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    
    Dim NID As NOTIFYICONDATA
    
    NID.hwnd = Me.hwnd
    NID.cbSize = Len(NID)
    NID.uID = vbNull
    NID.uFlags = NIF_TIP Or NIF_ICON Or NIF_MESSAGE
    NID.hIcon = Me.Icon
    NID.uCallbackMessage = WM_MOUSEMOVE
    NID.szTip = "Right-Click to display Popupmenu"
    
    Shell_NotifyIcon NIM_DELETE, NID

End Sub
