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
   StartUpPosition =   2  '��Ļ����
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
      Caption         =   "����è"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "��סCtrl���ƶ���������Ļȡɫ��Ctrl+Alt+C���ر�ȡɫ���"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "ȷ��"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "�����潫��3���رգ��������½ǵ�����ͼ���������в���"
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
      Caption         =   "����è��Ļȡɫ��V1.2"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "Ctrl+Alt+X��������Ļȡɫ���"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "�˵�"
      Visible         =   0   'False
      Begin VB.Menu menu_showcolor 
         Caption         =   "��ʾȡɫ���"
      End
      Begin VB.Menu menu_help 
         Caption         =   "�������"
      End
      Begin VB.Menu menu_about 
         Caption         =   "��������"
      End
      Begin VB.Menu menu_newversion 
         Caption         =   "��ȡ���°汾"
      End
      Begin VB.Menu menu_exit 
         Caption         =   "�˳����"
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
MsgBox "�����ѿ�������֧���ظ���", vbYes, "����è��ʾ"
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
MsgBox "���ߣ�����è" & vbCrLf & "QQ:842417019" & vbCrLf & "email:842417019@qq.com", vbYes, "����è��ʾ"
End Sub

Private Sub menu_exit_Click()
Call Shell_NotifyIcon(NIM_DELETE, nfIconData)
End
End Sub

Private Sub menu_help_Click()
MsgBox "��Ļȡɫ��ʹ�ð�����" & vbCrLf & "����Ctrl+Alt+X��ϼ�������Ļȡɫ��壬��סShift���ƶ���꣬ȡɫ��彫������������Ļȡɫ������Ctrl+Alt+C��ϼ��ر���Ļȡɫ����" & vbCrLf & "��ȡɫ�����ʾʱ������Ctrl+Alt+S��ϼ���������Զ�����16������ɫֵ�����а壬��Ҳ���Ե�����ɫֵ������ֵ��������" & vbCrLf & "������ܼ򵥣�δ����������" & vbCrLf & "����������ܻ���и��£����η���ʱ��:2016/7/32" & vbCrLf & "2017/1/16����:�޸���ɫֵR��B��ֵ�ߵ�����" & vbCrLf & "2019/10/19:�޸�Win10ȡɫ��λ����(����ĳЩϵͳ��������)"
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
        Label3.Caption = "�����潫��" & lens - 1 & "���رգ��������½ǵ�����ͼ���������в���"
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
      '��������ƶ���������ʱ��ʾ��Tip
      .szTip = "������ʾȡɫ��壬�һ���ʾ����˵�"
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
'��ʾȡɫ���
Private Sub showcolorview()
    Form1.Show
    has_scw = True
    Form1.showDrawView
End Sub


