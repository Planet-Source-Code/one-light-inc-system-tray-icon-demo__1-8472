Attribute VB_Name = "SysTrayBas"
'Declarations required to display a system tray icon
Public nid As NOTIFYICONDATA
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public Const WM_CHAR = &H102
Public Const WM_SETTEXT = &HC
Public Const WM_USER = &H400
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_CLEAR = &H303
Public Const WM_DESTROY = &H2
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
    End Type
    'constants required by Shell_NotifyIcon API call:
    Public Const NIM_ADD = &H0
    Public Const NIM_MODIFY = &H1
    Public Const NIM_DELETE = &H2
    Public Const NIF_MESSAGE = &H1
    Public Const NIF_ICON = &H2
    Public Const NIF_TIP = &H4
    Public Const WM_MOUSEMOVE = &H200

Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Sub InitializeTrayIcon()

      With nid
        .cbSize = Len(nid)
        .hwnd = frmMain.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .szTip = "Right click to display menu, left click to display window" & vbNullChar
        .hIcon = frmMain.Icon
    End With
    
    Shell_NotifyIcon NIM_ADD, nid
End Sub
