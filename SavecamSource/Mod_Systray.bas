Attribute VB_Name = "Mod_Systray"
'A module for some systray options...
Option Explicit

'The API stuff...
Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NDats) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RDats, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DrawAnimatedRects Lib "user32" (ByVal hwnd As Long, ByVal idAni As Long, lprcFrom As RDats, lprcTo As RDats) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'Some constants...
Private Const IDANI_OPEN = &H1
Private Const IDANI_CLOSE = &H2
Private Const IDANI_CAPTION = &H3
Private Const NIM_ADD = &H0
Private Const NIM_DELETE = &H2
Private Const NIM_MODIFY = &H1
Private Const NIF_ICON = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_TIP = &H4
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205

'Some types...

Private Type RDats
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type NDats
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type


'The icon self...
Private TIcon As NDats

'Let's add the icon to the systray...
Public Sub ST_ON(formName As Form, Tooltip As String, Icon As Long)
    
    TIcon.cbSize = Len(TIcon)
    TIcon.hwnd = formName.hwnd
    TIcon.szTip = Tooltip & vbNullChar
    TIcon.hIcon = Icon
    TIcon.uID = vbNull
    TIcon.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
    TIcon.uCallbackMessage = WM_MOUSEMOVE
    
    Shell_NotifyIcon NIM_ADD, TIcon
End Sub

'To remove the icon from the systray we use this...
Public Sub ST_OFF()
    Shell_NotifyIcon NIM_DELETE, TIcon
End Sub

'To change the systrayicon we use this...
Public Sub ST_Icon(Icon As Long)
    TIcon.hIcon = Icon
    Shell_NotifyIcon NIM_MODIFY, TIcon
End Sub

'To change the tooltip we use this
Public Sub ST_ToolTip(Tooltip As String)
    TIcon.szTip = Tooltip & vbNullChar
    Shell_NotifyIcon NIM_MODIFY, TIcon
End Sub

'Now let's make a little animation for the users...
Public Sub ST_ANI(formName As Form, Minimize As Boolean)

  Dim rSource As RDats, rDest As RDats, ScreenWidth As Long, ScreenHeight As Long
  Dim FormLeft As Long, FormTop As Long, FormWidth As Long, FormHeight As Long
    FormWidth = formName.Width / Screen.TwipsPerPixelX
    FormHeight = formName.Height / Screen.TwipsPerPixelY
    FormLeft = formName.Left / Screen.TwipsPerPixelX
    FormTop = formName.Top / Screen.TwipsPerPixelY
    ScreenWidth = Screen.Width / Screen.TwipsPerPixelX
    ScreenHeight = Screen.Height / Screen.TwipsPerPixelY
    SetRect rSource, FormLeft, FormTop, FormWidth, FormHeight
    SetRect rDest, ScreenWidth - 50, ScreenHeight, ScreenWidth - 50, ScreenHeight
    If Minimize Then
        DrawAnimatedRects formName.hwnd, IDANI_CLOSE Or IDANI_CAPTION, rSource, rDest
      Else
        DrawAnimatedRects formName.hwnd, IDANI_CLOSE Or IDANI_CAPTION, rDest, rSource
    End If

End Sub
