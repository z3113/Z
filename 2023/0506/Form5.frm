VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "����1"
   ClientHeight    =   5550
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9135
   LinkTopic       =   "Form5"
   ScaleHeight     =   5550
   ScaleWidth      =   9135
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      Height          =   495
      Left            =   6480
      TabIndex        =   1
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "�������к�Ӽ�������5����������˫������     ������������ݣ�����ʾ���е��������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   5880
      TabIndex        =   0
      Top             =   720
      Width           =   2175
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim a(5) As Integer, i As Integer
Private Sub Command1_Click()
    Unload Form5
End Sub

Private Sub Form_Activate()
    For i = 1 To 5
        a(i) = Val(InputBox("�������" & i & "����", "������", 10))
    Next i
End Sub

Private Sub Form_DblClick()
    Dim max As Integer
    max = 0
    Print "��������������ǣ�"
    For i = 1 To 5
        If max <= a(i) Then max = a(i)
        Print a(i);
    Next i
    Print
    Print "���ֵ�ǣ�" & max
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub
