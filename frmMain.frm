VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   " System Tray Demo"
   ClientHeight    =   2880
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Tray Icon Demo"
      Height          =   2175
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4455
      Begin VB.Label Label2 
         Caption         =   "by Ryan Lederman - ryanled@pacbell.net"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1800
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   $"frmMain.frx":058A
         Height          =   1455
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   4215
      End
   End
   Begin VB.CommandButton btnHide 
      Caption         =   "&Hide"
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   2400
      Width           =   975
   End
   Begin VB.Menu MNU 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu Item1 
         Caption         =   "Menu Item 1"
      End
      Begin VB.Menu Item2 
         Caption         =   "Menu Item 2"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu ExitMNU 
         Caption         =   "Exit app and hide icon"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnHide_Click()
Me.Hide
End Sub

Private Sub ExitMNU_Click()
'Exit menu item was clicked.
Shell_NotifyIcon NIM_DELETE, nid
End
End Sub

Private Sub Form_Load()
InitializeTrayIcon 'This sub creates the icon in the system tray
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Result As Long
    Dim msg As Long

    If Me.ScaleMode = vbPixels Then
        msg = X
    Else
        msg = X / Screen.TwipsPerPixelX
    End If
    
    Select Case msg
    
        Case 517 '517 display popup menu
        
            Me.PopupMenu MNU
        
        Case 514 '514 left mouse button
        
            Result = SetForegroundWindow(Me.hwnd)
            Me.Show
            
    End Select
End Sub

Private Sub Form_Terminate()
Shell_NotifyIcon NIM_DELETE, nid 'Remove icon from system tray
End Sub

Private Sub Form_Unload(Cancel As Integer)
Shell_NotifyIcon NIM_DELETE, nid 'Remove icon from system tray
End Sub

Private Sub Item1_Click()
MsgBox "Menu Item 1", 64, "System Tray Icon Demo"
End Sub

Private Sub Item2_Click()
MsgBox "Menu Item 2", 64, "System Tray Icon Demo"
End Sub
