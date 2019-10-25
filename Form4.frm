VERSION 5.00
Begin VB.Form Form4 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000012&
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   1950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6045
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   720
      Top             =   1440
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   1440
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "玄翼猫"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "按住Ctrl键移动鼠标进行屏幕取色，Ctrl+Alt+C键关闭取色面板"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   795
      Width           =   5775
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "确定"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "本界面将在3秒后关闭，请点击右下角的托盘图标对软件进行操作"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   5175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "玄翼猫屏幕取色器V1.2"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Ctrl+Alt+X键调出屏幕取色面板"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   4455
   End
   Begin VB.Menu menu_main 
      Caption         =   "菜单"
      Visible         =   0   'False
      Begin VB.Menu menu_showcolor 
         Caption         =   "显示取色面板"
      End
      Begin VB.Menu menu_help 
         Caption         =   "软件帮助"
      End
      Begin VB.Menu menu_about 
         Caption         =   "关于作者"
      End
      Begin VB.Menu menu_newversion 
         Caption         =   "获取最新版本"
      End
      Begin VB.Menu menu_exit 
         Caption         =   "退出软件"
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const MAX_TOOLTIP As Integer = 64
Private Const NIF_ICON = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_TIP = &H4
Private Const NIM_ADD = &H0
Private Const NIM_DELETE = &H2
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_RBUTTONDBLCLK = &H206

Private Const SW_RESTORE = 9
Private Const SW_HIDE = 0

Private nfIconData As NOTIFYICONDATA


Private Type NOTIFYICONDATA
  cbSize As Long
  hwnd As Long
  uID As Long
  uFlags As Long
  uCallbackMessage As Long
  hIcon As Long
  szTip As String * MAX_TOOLTIP
End Type

Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Private Const WS_SYSMENU = &H80000
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_CAPTION = &HC00000
Private Const GWL_STYLE = (-16)

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public has_scw As Boolean
Dim lens As Integer
Private Sub Form_Load()
If App.PrevInstance Then
MsgBox "程序已开启，不支持重复打开", vbYes, "玄翼猫提示"
End
End If
'Form4.BackColor = RGB(78, 200, 59)
has_scw = True
lens = 3
Dim l As Long
l = GetWindowLong(hwnd, GWL_STYLE)
l = l And Not WS_SYSMENU
l = l And Not WS_MAXIMIZEBOX
l = l And Not WS_MINIMIZEBOX
l = l And Not WS_CAPTION
Call SetWindowLong(hwnd, GWL_STYLE, l)
End Sub

Private Sub Label4_Click()
    hideview
End Sub

Private Sub menu_about_Click()
MsgBox "作者：玄翼猫" & vbCrLf & "QQ:842417019" & vbCrLf & "email:842417019@qq.com", vbYes, "玄翼猫提示"
End Sub

Private Sub menu_exit_Click()
Call Shell_NotifyIcon(NIM_DELETE, nfIconData)
End
End Sub

Private Sub menu_help_Click()
MsgBox "屏幕取色器使用帮助：" & vbCrLf & "按下Ctrl+Alt+X组合键调出屏幕取色面板，按住Shift键移动鼠标，取色面板将跟随鼠标进行屏幕取色，按下Ctrl+Alt+C组合键关闭屏幕取色窗口" & vbCrLf & "在取色面板显示时，按下Ctrl+Alt+S组合键，软件会自动复制16进制颜色值到剪切板，您也可以单击颜色值或坐标值复制它们" & vbCrLf & "软件功能简单，未做界面美化" & vbCrLf & "软件后续可能会进行更新，本次发布时间:2016/7/32" & vbCrLf & "2017/1/16更新:修复颜色值R、B的值颠倒问题" & vbCrLf & "2019/10/19:修复Win10取色错位问题(可能某些系统仍有问题)"
End Sub

Private Sub menu_newversion_Click()
Shell "cmd /c start http://xuanyimao.com", vbHide
End Sub

Private Sub menu_showcolor_Click()
showcolorview
End Sub

Private Sub Timer1_Timer()
    If lens = 0 Then
        hideview
    Else
        Label3.Caption = "本界面将在" & lens - 1 & "秒后关闭，请点击右下角的托盘图标对软件进行操作"
        lens = lens - 1
    End If
End Sub

Private Sub hideview()
    With nfIconData
      .hwnd = Me.hwnd
      .uID = Me.Icon
      .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
      .uCallbackMessage = WM_MOUSEMOVE
      .hIcon = Me.Icon.Handle
      '定义鼠标移动到托盘上时显示的Tip
      .szTip = "单击显示取色面板，右击显示软件菜单"
      .cbSize = Len(nfIconData)
    End With
    Call Shell_NotifyIcon(NIM_ADD, nfIconData)
    Form4.Hide
    Timer1.Enabled = False
    Timer2.Enabled = True
    has_scw = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lMsg As Single
    lMsg = X / Screen.TwipsPerPixelX
    Select Case lMsg
    Case WM_LBUTTONUP
        'ShowWindow Me.hWnd, SW_RESTORE
        showcolorview
    Case WM_RBUTTONUP
        PopupMenu menu_main
    End Select
End Sub

Private Sub Timer2_Timer()
    If has_scw = False Then
        If GetAsyncKeyState(vbKeyControl) And GetAsyncKeyState(vbKeyMenu) And GetAsyncKeyState(vbKeyX) Then
            showcolorview
        End If
    End If
End Sub
'显示取色面板
Private Sub showcolorview()
    Form1.Show
    has_scw = True
    Form1.showDrawView
End Sub


