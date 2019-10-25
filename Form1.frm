VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   3105
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2115
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3105
   ScaleWidth      =   2115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   240
      Top             =   120
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   6855
      Left            =   4800
      ScaleHeight     =   6855
      ScaleWidth      =   9375
      TabIndex        =   0
      Top             =   120
      Width           =   9375
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1920
      Top             =   3600
   End
   Begin VB.Label ldesc 
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lrgb 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1080
      TabIndex        =   6
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label lcolor 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   2445
      Width           =   975
   End
   Begin VB.Label lpoint 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   2175
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "RGB值："
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   15
      TabIndex        =   3
      Top             =   2760
      Width           =   1080
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "16进制颜色："
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   15
      TabIndex        =   2
      Top             =   2445
      Width           =   1080
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "屏幕坐标："
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   20
      TabIndex        =   1
      Top             =   2175
      Width           =   1080
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function getpixel Lib "gdi32" Alias "GetPixel" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetProcessDpiAwareness Lib "SHCORE.DLL" (ByVal DPImodel As Long) As Long

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

Const HWND_TOPMOST = -1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1

Private Const LOGPIXELSX = 88
Private Const LOGPIXELSY = 90

Private Type POINTAPI
X As Long
Y As Long
End Type
Dim p As POINTAPI
Dim swidth As Integer
Dim pointstr As String
Dim colorstr As String
Dim rgbstr As String
Dim stime As Integer
Dim colarr(21, 21) As Long

'屏幕实际宽高
Dim screen_width As Long
Dim screen_height As Long
'屏幕缩放比例
Dim screen_scale As Integer


Private Sub Form_Load()
If App.PrevInstance Then
MsgBox "程序已开启，不支持重复打开", vbYes, "玄翼猫提示"
End
End If
screen_scale = 100
SetProcessDpiAwareness 2
'Form1.BackColor = RGB(78, 200, 59)
'Label1.Alignment = 2
'Label1.Caption = "颜色值"
swidth = 100

dpi_x = GetDeviceCaps(GetDC(0), LOGPIXELSX)
'MsgBox dpi_x
If dpi_x = 120 Then
    screen_scale = 125
ElseIf dpi_x = 144 Then
    screen_scale = 150
ElseIf dpi_x = 192 Then
    screen_scale = 200
End If

screen_width = Screen.Width * screen_scale
screen_height = Screen.Height * screen_scale
'Picture1.Width = Screen.Width
'Picture1.Height = Screen.Height

Picture1.Width = screen_width
Picture1.Height = screen_height
'窗口置顶
Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
stime = 1000
setshowlocation

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cx = Int(X / swidth)
    cy = Int(Y / swidth)
    resetTabels cx, cy
End Sub

Private Sub Label1_Click()
copycontent pointstr
End Sub

Private Sub Label2_Click()
copycontent colorstr
End Sub

Private Sub Label3_Click()
copycontent rgbstr
End Sub

Private Sub lcolor_Click()
copycontent colorstr
End Sub

Private Sub lpoint_Click()
copycontent pointstr
End Sub

Private Sub lrgb_Click()
copycontent rgbstr
End Sub

Private Sub Timer1_Timer()
'vbKeyControl   vbKeyMenu：alt键
If Form4.has_scw Then
    If GetAsyncKeyState(vbKeyControl) And GetAsyncKeyState(vbKeyMenu) And GetAsyncKeyState(vbKeyC) Then
        Form4.has_scw = False
        Me.Hide
    ElseIf GetAsyncKeyState(vbKeyControl) And GetAsyncKeyState(vbKeyMenu) And GetAsyncKeyState(vbKeyS) Then
        copycontent colorstr
    ElseIf GetAsyncKeyState(vbKeyShift) Then
        setshowlocation
        'BitBlt Picture1.hdc, 0, 0, Screen.Width, Screen.Height, GetDC(0), 0, 0, vbSrcCopy
        BitBlt Picture1.hdc, 0, 0, screen_width, screen_height, GetDC(0), 0, 0, vbSrcCopy
        
        Picture1.Refresh
        getPointColor
    End If
End If
'getPointColor
'GetCursorPos p
'hd = GetDC(0)
'r = getpixel(hd, p.x, p.y)

'Print r
'csr = CStr(Hex(r))
't = Len(csr)
'Select Case t
'Case 1
'sr = "00000" & csr
'Case 2
'sr = "0000" & csr
'Case 3
'sr = "000" & csr
'Case 4
'sr = "00" & csr
'Case 5
'sr = "0" & csr
'Case 6
'sr = csr
'End Select

'Text1.Text = " &&H" & sr
'Text1.Text = Text1.Text & vbCrLf & getpixel(hd, p.x - 3, p.y - 3)
End Sub
Private Sub getPointColor()
GetCursorPos p
'hd = GetDC(0)
hd = Picture1.hdc
X = p.X
Y = p.Y
pointstr = X & "," & Y

sx = X - 10
sy = Y - 10
ex = X + 10
ey = Y + 10
For i = sx To ex
    For j = sy To ey
        If i < 0 Or j < 0 Then
            drawView i - sx, j - sy, 0, False
        Else
            r = getpixel(hd, i, j)
            If r < 0 Then
                r = 0
            End If
            colorhex = rgbToColor(r)
            If X = i And Y = j Then
                'Text1 = "#" & colorstr
                colorstr = colorhex
                setColorRGB colorhex
            End If
            colarr(i - sx + 1, j - sy + 1) = r
             '   drawView i - sx, j - sy, r, True
            'Else
                drawView i - sx, j - sy, r, False
            'End If
        End If
    Next
Next

Me.FillStyle = 1 '空心
Me.Line ((X - sx) * swidth, (Y - sy) * swidth)-((X - sx + 1) * swidth, (Y - sy + 1) * swidth), RGB(255, 0, 0), B
showdesc
'drawView x - sx, y - sy, r, True
End Sub


'绘制界面
Private Sub drawView(X, Y, col, sel As Boolean)
    If col < 0 Then
        col = 0
    End If
    Me.FillColor = col
    Me.FillStyle = 0 '填充
    DrawWidth = 1
    Me.Line (X * swidth, Y * swidth)-((X + 1) * swidth, (Y + 1) * swidth), col, B
    
    'Me.FillColor = 0
    Me.FillStyle = 1 '空心
    If sel Then '选中的孩子用红色边框
        Me.Line (X * swidth, Y * swidth)-((X + 1) * swidth, (Y + 1) * swidth), RGB(237, 28, 36), B
    Else
        Me.Line (X * swidth, Y * swidth)-((X + 1) * swidth, (Y + 1) * swidth), 0, B
    End If
    
End Sub

Function ColorRGB(Color As String) As Long
Dim A(2) As Long
A(0) = CLng("&H" & Mid(Color, 1, 2))
A(1) = CLng("&H" & Mid(Color, 3, 2))
A(2) = CLng("&H" & Mid(Color, 5, 2))
'MsgBox A(0) & "  " & A(1) & "  " & A(2)
ColorRGB = RGB(A(0), A(1), A(2))
End Function
'将十进制转换成16进制
Function rgbToColor(num) As String
red = num Mod 256
green = (num \ 256) Mod 256
blue = num \ 256 \ 256

csr = CStr(Hex(red))
If Len(csr) = 1 Then
sr = "0" & csr
Else
sr = csr
End If

csr = CStr(Hex(green))
If Len(csr) = 1 Then
sr = sr & "0" & csr
Else
sr = sr & csr
End If

csr = CStr(Hex(blue))
If Len(csr) = 1 Then
sr = sr & "0" & csr
Else
sr = sr & csr
End If

rgbToColor = sr
End Function


Function setColorRGB(Color) As Long

Dim A(2) As Long
A(0) = CLng("&H" & Mid(Color, 1, 2))
A(1) = CLng("&H" & Mid(Color, 3, 2))
A(2) = CLng("&H" & Mid(Color, 5, 2))
rgbstr = A(0) & "," & A(1) & "," & A(2)
'Text2 = A(0) & "," & A(1) & "," & A(2)

End Function

Private Sub showdesc()
lpoint.Caption = pointstr
lcolor.Caption = colorstr
lrgb.Caption = rgbstr
End Sub
Private Sub copycontent(str)
Clipboard.Clear
Clipboard.SetText str
ldesc.Caption = "已将数值：" & vbCrLf & "  " & str & vbCrLf & "复制到剪切板"
ldesc.Visible = True
Timer2.Enabled = True
stime = 1000
End Sub

Private Sub Timer2_Timer()
    If stime <= 0 Then
        Timer2.Enabled = False
        ldesc.Visible = False
    Else
        stime = stime - 1000
    End If
End Sub
'设置窗体的显示位置
Private Sub setshowlocation()
    GetCursorPos p
    mx = p.X + 20
    my = p.Y + 20
    If mx * 15 + Me.Width > Screen.Width Then
        mx = (p.X - 20) * 15 - Me.Width
    Else
        mx = mx * 15
    End If
    If my * 15 + Me.Height > Screen.Height Then
        my = (p.Y - 20) * 15 - Me.Height
    Else
        my = my * 15
    End If
    Me.Move mx, my
End Sub
'重绘表格
Private Sub resetTabels(X, Y)
    For i = 1 To 21
        For j = 1 To 21
            drawView i - 1, j - 1, colarr(i, j), False
        Next
    Next
    Me.FillStyle = 1 '空心
    Me.Line (X * swidth, Y * swidth)-((X + 1) * swidth, (Y + 1) * swidth), RGB(255, 0, 0), B
    colorstr = rgbToColor(colarr(X + 1, Y + 1))
    setColorRGB colorstr
    pointstr = ""
    showdesc
End Sub

Public Sub showDrawView()
    setshowlocation
    BitBlt Picture1.hdc, 0, 0, screen_width, screen_height, GetDC(0), 0, 0, vbSrcCopy
    
    Picture1.Refresh
    getPointColor
End Sub
