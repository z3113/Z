VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "ͼ�δ�ӡ"
   ClientHeight    =   6135
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8535
   LinkTopic       =   "Form4"
   ScaleHeight     =   6135
   ScaleWidth      =   8535
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command3 
      Caption         =   "ͼ��3"
      Height          =   495
      Left            =   6960
      TabIndex        =   2
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ͼ��2"
      Height          =   495
      Left            =   6960
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ͼ��1"
      Height          =   495
      Left            =   6960
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Cls
    Dim i%, j%, a%
    a = Val(InputBox("�����벻С��2��������", "����", 5))
    If a >= 2 Then
        Print "1234567890123456789012345678901234567890"
        For i = 1 To a
            Print Tab(3 * (a - i) + 1);
            For j = 1 To i
                Print i;
            Next j
        Next i
    Else
        MsgBox "������������", vbOKCancel + 16, "��ܰ��ʾ"
    End If
End Sub

Private Sub Command2_Click()
    Cls
    Dim i%, j%, a%
    a = Val(InputBox("�����벻С��2��������", "����", 5))
    If a >= 2 Then
        Print "1234567890123456789012345678901234567890"
        For i = 1 To a
            Print Tab(a - i + 1);
            For j = 1 To 2 * i - 1
                If j Mod 2 = 0 Then Print "*"; Else Print "$";
            Next j
        Next i
    Else
        MsgBox "�������ִ���", vbOKCancel + 16, "��ܰ��ʾ"
    End If
End Sub

Private Sub Command3_Click()
    Cls
    Dim i%, j%, a%
    a = Val(InputBox("�����벻С��2��������", "����", 5))
    If a >= 2 Then
        Print "1234567890123456789012345678901234567890"
        For i = -a To a
            Print Tab(a - Abs(i) + 1);
            For j = -Abs(i) To Abs(i)
                Print "$";
            Next j
        Next i
    Else
        MsgBox "�������ִ���", vbOKCancel + 16, "��ܰ��ʾ"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub
