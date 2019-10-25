VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H8000000E&
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   1125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1215
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1125
   ScaleWidth      =   1215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "退出软件"
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   750
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "关于作者"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   375
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "使用帮助"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()
MsgBox "屏幕取色器使用帮助：" & vbCrLf & "按下shift键进行屏幕取色，按下ctrl键关闭屏幕取色窗口" & vbCrLf & "由于本软件仅花了半天时间开发，未做界面美化，如有不便，敬请谅解" & vbCrLf & "软件后续可能会进行更新，本次发布时间:2016/7/31"
End Sub

Private Sub Label2_Click()
MsgBox "作者：玄翼猫" & vbCrLf & "QQ:842417019" & vbCrLf & "email:opq842417019@163.com", vbYes, "玄翼猫提示"
Shell "cmd /c start http://xuanyimao.com", vbHide
End Sub

Private Sub Label3_Click()
End
End Sub
