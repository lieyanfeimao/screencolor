VERSION 5.00
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FF0000&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   510
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   510
   ScaleWidth      =   510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   240
      Top             =   120
   End
   Begin VB.Image Image1 
      Height          =   512
      Left            =   1
      Picture         =   "Form2.frx":048A
      Stretch         =   -1  'True
      Top             =   1
      Width           =   512
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Const SPI_GETWORKAREA = 48

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type


Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1
Private Const HWND_TOPMOST = -1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1

Public has_scw As Boolean
Dim has_showmenu As Boolean

Private Sub Form_Load()
If App.PrevInstance Then
MsgBox "程序已开启，不支持重复打开", vbYes, "玄翼猫提示"
End
End If
Dim rtn As Long
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, &HFF0000, 0, LWA_COLORKEY
Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)

Dim lRes As Long
Dim rectVal As RECT
Dim TaskbarHeight As Integer
lRes = SystemParametersInfo(SPI_GETWORKAREA, 0, rectVal, 0)
TaskbarHeight = Screen.Height - rectVal.Bottom * Screen.TwipsPerPixelY
Me.Move Screen.Width - Me.Width - 100, Screen.Height - Me.Height - TaskbarHeight - 500, Me.Width, Me.Height

has_showmenu = False
Image1.ToolTipText = "鼠标左键拖动图标，右键显示菜单"
has_scw = False
End Sub


Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then '判断鼠标的左键被按下
    If has_showmenu Then
        Form3.Hide
        has_showmenu = False
    Else
        Call ReleaseCapture
        Call SendMessage(hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
    End If
ElseIf Button = vbRightButton Then
    '显示菜单
    If Form2.Top - Form3.Height < 0 Then
        Form3.Top = Form2.Height
    Else
        Form3.Top = Form2.Top - Form3.Height
    End If
    
    If Form2.Left + Form2.Width - Form3.Width < 0 Then
        Form3.Left = Form2.Left
    Else
        Form3.Left = Form2.Left + Form2.Width - Form3.Width
    End If
    Form3.Show
    has_showmenu = True
End If
End Sub

Private Sub Timer1_Timer()
    If GetAsyncKeyState(vbKeyControl) Then
        If has_scw = False Then
            Form1.Show
            has_scw = True
        End If
    End If
End Sub
