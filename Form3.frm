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
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "�˳����"
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
      Caption         =   "��������"
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
      Caption         =   "ʹ�ð���"
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
MsgBox "��Ļȡɫ��ʹ�ð�����" & vbCrLf & "����shift��������Ļȡɫ������ctrl���ر���Ļȡɫ����" & vbCrLf & "���ڱ���������˰���ʱ�俪����δ���������������в��㣬�����½�" & vbCrLf & "����������ܻ���и��£����η���ʱ��:2016/7/31"
End Sub

Private Sub Label2_Click()
MsgBox "���ߣ�����è" & vbCrLf & "QQ:842417019" & vbCrLf & "email:opq842417019@163.com", vbYes, "����è��ʾ"
Shell "cmd /c start http://xuanyimao.com", vbHide
End Sub

Private Sub Label3_Click()
End
End Sub
