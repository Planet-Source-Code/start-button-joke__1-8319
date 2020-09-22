VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Window's 'Start' Button Joke!"
   ClientHeight    =   1440
   ClientLeft      =   4476
   ClientTop       =   3540
   ClientWidth     =   3744
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1440
   ScaleWidth      =   3744
   Visible         =   0   'False
   Begin VB.Timer timerCheckCoor 
      Enabled         =   0   'False
      Index           =   3
      Interval        =   5
      Left            =   3492
      Top             =   144
   End
   Begin VB.Timer timerCheckCoor 
      Enabled         =   0   'False
      Index           =   2
      Interval        =   5
      Left            =   3204
      Top             =   540
   End
   Begin VB.CommandButton cmdSet2Def 
      Caption         =   "Set to &default position"
      Height          =   444
      Left            =   288
      TabIndex        =   2
      Top             =   936
      Width           =   1416
   End
   Begin VB.Timer timerCheckCoor 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   5
      Left            =   3204
      Top             =   180
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "&Stop"
      Height          =   516
      Left            =   1908
      TabIndex        =   1
      Top             =   252
      Width           =   1200
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "&Go!"
      Height          =   516
      Left            =   504
      TabIndex        =   0
      Top             =   252
      Width           =   1056
   End
   Begin VB.Label lblMailTo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "oroman@mail.ru"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   1980
      MouseIcon       =   "frmMain.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   3
      ToolTipText     =   "Write me a letter"
      Top             =   1044
      Width           =   1560
   End
   Begin VB.Menu mnuFileItem 
      Caption         =   "&File"
      Begin VB.Menu mnuExitItem 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuOptionsItem 
      Caption         =   "&Options"
      Begin VB.Menu mnuSet1Item 
         Caption         =   "Set &1"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSet2Item 
         Caption         =   "Set &2"
      End
      Begin VB.Menu mnuSet3Item 
         Caption         =   "Set &3"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAboutItem 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGo_Click()
timerCheckCoor(1).Enabled = False
timerCheckCoor(2).Enabled = False
timerCheckCoor(3).Enabled = False
If StartButtonhwnd <> 0 Then timerCheckCoor(CurrentSet).Enabled = True Else _
    MsgBox "Can't find start button!" & vbCrLf & "Make sure that Windows Explorer is runned!", vbInformation Or &H40000: GetWindowHandles
End Sub

Private Sub cmdSet2Def_Click()

If StartButtonhwnd <> 0 Then
    SetWindowPos StartButtonhwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOZORDER Or SWP_SHOWWINDOW Or SWP_NOSIZE
Else
    MsgBox "Can't find start button!" & vbCrLf & "Make sure that Windows Explorer is runned!", vbInformation Or &H40000: GetWindowHandles
End If
End Sub

Private Sub cmdStop_Click()
timerCheckCoor(1).Enabled = False
timerCheckCoor(2).Enabled = False
timerCheckCoor(3).Enabled = False
End Sub

Private Sub Form_Load()
Dim a As Long
Me.Caption = "tmp"
If InStr(UCase(Command$), "/SHOW") <> 0 Then
    a = FindWindow(vbNullString, "Window's 'Start' Button Joke!")
    If a <> 0 Then
        SetWindowPos a, HWND_NOTTOP, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_SHOWWINDOW
        End
    Else
        MsgBox "Can't find any of my copies runned!"
    End If
End If
Me.Caption = "Window's 'Start' Button Joke!"
CurrentSet = 1
If InStr(UCase(Command$), "/SET2") <> 0 Then CurrentSet = 2: mnuSet2Item_Click
If InStr(UCase(Command$), "/SET3") <> 0 Then CurrentSet = 3: mnuSet3Item_Click
If InStr(UCase(Command$), "/STEP") <> 0 Then Delta = Abs(Val(Mid$(Command$, InStr(UCase(Command$), "/STEP") + 5, 8)))
If Delta <= 0 Then Delta = 5
Delta2 = Delta
GetWindowHandles
InvisibleForTaskMngr 'making invisible for task manager
                     'doesn't work on WinNT :(
If StartButtonhwnd = 0 Then MsgBox "Can't find start button!" & vbCrLf & "Make sure that Windows Explorer is runned!", vbInformation Or &H40000 Else cmdGo_Click
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
lblMailTo.FontUnderline = False
End Sub

Private Sub lblMailTo_Click()
Shell "start.exe " & Chr$(34) & "mailto:oroman@mail.ru?Subject=StartButtonJoke" & Chr$(34)
End Sub

Private Sub lblMailTo_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Not (lblMailTo.FontUnderline) Then lblMailTo.FontUnderline = True
End Sub

Private Sub mnuAboutItem_Click()
ShowAbout
End Sub

Private Sub mnuExitItem_Click()
Unload Me
End Sub

Private Sub mnuSet1Item_Click()
mnuSet1Item.Checked = True
mnuSet2Item.Checked = False
mnuSet3Item.Checked = False
CurrentSet = 1
End Sub

Private Sub mnuSet2Item_Click()
mnuSet2Item.Checked = True
mnuSet1Item.Checked = False
mnuSet3Item.Checked = False
CurrentSet = 2

End Sub

Private Sub mnuSet3Item_Click()
mnuSet3Item.Checked = True
mnuSet1Item.Checked = False
mnuSet2Item.Checked = False
CurrentSet = 3
End Sub

Private Sub timerCheckCoor_Timer(Index As Integer)
Dim GetStartButtonPos As POINTAPI
Dim GetCurPos As POINTAPI
GetCursorPos GetCurPos

Select Case Index

Case 1
ClientToScreen StartButtonhwnd, GetStartButtonPos
'this converts internal coordinates of the window to the screen coordinates


If (GetStartButtonPos.x - GetCurPos.x > -(StartButtonWidth + 5) And GetStartButtonPos.x - GetCurPos.x < 5) And _
(GetStartButtonPos.y - GetCurPos.y > -(StartButtonHeight + 5) And GetStartButtonPos.y - GetCurPos.y < 5) Then
    SetWindowPos StartButtonhwnd, HWND_TOPMOST, Int(Rnd * (TrayWndWidth - StartButtonWidth)), 0, 0, 0, SWP_NOZORDER Or SWP_SHOWWINDOW Or SWP_NOSIZE
    'if we use swp_nosize flag,width and height values are ignored
    'I adding 5 to size of the windows in order to prevent clicking
    'on it's border
End If
Case 2
    ClientToScreen TrayWndHwnd, GetStartButtonPos

    If (GetStartButtonPos.x - GetCurPos.x > -(TrayWndWidth + 5) And GetStartButtonPos.x - GetCurPos.x < 4) And _
    (GetStartButtonPos.y - GetCurPos.y > -(TrayWndHeight + 5) And GetStartButtonPos.y - GetCurPos.y < 4) Then
        SetWindowPos TrayWndHwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOZORDER Or SWP_HIDEWINDOW Or SWP_NOSIZE Or SWP_NOMOVE
    Else
        If IsWindowVisible(TrayWndHwnd) = 0 Then
            SetWindowPos TrayWndHwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOZORDER Or SWP_SHOWWINDOW Or SWP_NOSIZE Or SWP_NOMOVE
        End If
    End If
Case 3
    If StartBtnCurPos > TrayWndWidth - StartButtonWidth Then Delta2 = -Delta
    If StartBtnCurPos <= 0 Then Delta2 = Delta
    StartBtnCurPos = StartBtnCurPos + Delta2
    SetWindowPos StartButtonhwnd, HWND_TOPMOST, StartBtnCurPos, 0, 0, 0, SWP_NOZORDER Or SWP_SHOWWINDOW Or SWP_NOSIZE
End Select
End Sub
