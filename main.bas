Attribute VB_Name = "main"
Option Explicit
Public Type POINTAPI
        x As Long
        y As Long
End Type

Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Type WINDOWPLACEMENT
        Length As Long
        flags As Long
        showCmd As Long
        ptMinPosition As POINTAPI
        ptMaxPosition As POINTAPI
        rcNormalPosition As RECT
End Type


Public StartButtonhwnd, TrayWndHwnd As Long
Public StartButtonWidth, StartButtonHeight, TrayWndWidth, TrayWndHeight As Long
Public CurrentSet As Byte, Delta, Delta2 As Single
Public StartBtnCurPos As Integer

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, _
                     ByVal lpWindowName As String) As Long

Public Const SWP_SHOWWINDOW = &H40
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTTOP = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_NOZORDER = &H4

Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function GetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Public Declare Function RegisterServiceProcess Lib "kernel32.dll" (ByVal dwProcessId As Long, ByVal dwType As Long) As Long

Public Declare Function ShellAbout Lib "Shell32" Alias "ShellAboutA" (ByVal ParentHWND As Long, _
ByVal sTtitle As String, ByVal sOtherStuff As String, ByVal hIcon As Long) As Integer

Public Sub InvisibleForTaskMngr()
On Local Error GoTo extsub
RegisterServiceProcess 0, 1 '1 means to make invisible
extsub:
Exit Sub
End Sub
Public Sub GetTraySize()
Dim PointPos As POINTAPI
Dim WndPlacement As WINDOWPLACEMENT
WndPlacement.Length = Len(WndPlacement)
Call GetWindowPlacement(StartButtonhwnd, WndPlacement)
StartButtonHeight = WndPlacement.rcNormalPosition.Bottom - WndPlacement.rcNormalPosition.Top
StartButtonWidth = WndPlacement.rcNormalPosition.Right - WndPlacement.rcNormalPosition.Left

WndPlacement.Length = Len(WndPlacement)
Call GetWindowPlacement(TrayWndHwnd, WndPlacement)
TrayWndHeight = WndPlacement.rcNormalPosition.Bottom - WndPlacement.rcNormalPosition.Top
TrayWndWidth = WndPlacement.rcNormalPosition.Right - WndPlacement.rcNormalPosition.Left

ClientToScreen StartButtonhwnd, PointPos
StartBtnCurPos = PointPos.x - StartButtonWidth
End Sub
Public Sub ShowAbout()
Dim sOtherStuff As String

sOtherStuff = "Windows TrayWnd stuff#WinTrayWndStuff for"

Call ShellAbout(frmMain.hwnd, sOtherStuff & vbNullChar, _
"This program was made by [REA]CoolCold especially for Allhack", frmMain.Icon)
End Sub
Public Sub GetWindowHandles()
Dim sTEMp As String * 255 'buffer
Dim sTEMP1, sTEMP2 As String

sTEMP1 = "Shell_TrayWnd" + Chr$(0)  'this is class name of TaskBar Window
                                    '(in the bottom of desktop)
sTEMP2 = "Button" + Chr$(0) '       'this is classname of Start Button!

'class names of this windows were taken from BC++ 5.01 WinSight for Win32
'and WinApi help was provided by Borland Delphi 4.3 Win32 SDK help
'rather good help & windows and messages tracer :)) Nothing more :))
TrayWndHwnd = FindWindow("Shell_TrayWnd", vbNullString)
'getting handle of TaskBar
StartButtonhwnd = FindWindowEx(TrayWndHwnd, 0, sTEMP2, Chr$(0))
'getting handle of Start Button
GetTraySize 'getting size of these windows

End Sub
