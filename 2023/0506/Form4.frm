VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "�̶���ʽ���"
   ClientHeight    =   4455
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8925
   LinkTopic       =   "Form4"
   ScaleHeight     =   4455
   ScaleWidth      =   8925
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command3 
      Caption         =   "����"
      Height          =   495
      Left            =   7560
      TabIndex        =   2
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "���"
      Height          =   495
      Left            =   7560
      TabIndex        =   1
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      Height          =   495
      Left            =   7560
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim a(20) As Integer, i As Integer
Private Sub Command1_Click()
    Randomize
    For i = 1 To 20
        a(i) = Int(Rnd * (900) + 100)
    Next i
    Print "����A����20��Ԫ��ֵ"
End Sub

Private Sub Command2_Click()
    Cls
    For i = 1 To 20
        Print "A�����" & Format(i, "00") & "��Ԫ�ص�ֵ�� " & a(i)
    Next i
End Sub

Private Sub Command3_Click()
    Unload Form4
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub
