VERSION 5.00
Begin VB.Form Form4 
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9735
   LinkTopic       =   "Form4"
   ScaleHeight     =   4935
   ScaleWidth      =   9735
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   495
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Height          =   495
      Left            =   7920
      TabIndex        =   5
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Height          =   495
      Left            =   7920
      TabIndex        =   4
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Height          =   495
      Left            =   7920
      TabIndex        =   3
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Height          =   495
      Left            =   7920
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Height          =   495
      Left            =   5040
      TabIndex        =   1
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Height          =   495
      Left            =   5040
      TabIndex        =   0
      Top             =   2160
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
    Text1.Text = InputBox("�����ĵ�һ�����ݣ����������ݣ�", "����", 1) + InputBox("�����ĵڶ������ݣ����������ݣ�", "����", 2)
End Sub

Private Sub Command2_Click()
    Text2.Text = Val(InputBox("��ӵĵڶ������ݣ����������ݣ�", "���", 2)) + Val(InputBox("��ӵĵ�һ�����ݣ����������ݣ�", "���", 1))
End Sub

Private Sub Command3_Click()
    MsgBox "���ǵ�һ����Ϣ���ı���text1�е������ǣ�" & Text1.Text, 1 + 48, "�������"
End Sub

Private Sub Command4_Click()
    MsgBox "���ǵڶ�����Ϣ���ı���text2�е������ǣ�" & Chr(10) & Chr(13) & Text2.Text, 3 + 64, "�������"
End Sub

Private Sub Command5_Click()
    Cls
    Print "123456789012345678901234567890"
    Print "1212ab"
    Print "123456789012345678901234567890"
    Print "12 12 ab"
    Print "123456789012345678901234567890"
    Print Tab(3); "12ab"; 12; "ab"
    Print "123456789012345678901234567890"
    Print Tab(6); "12ab", "12 12"
    Print "123456789012345678901234567890"
    Print "12ab12ab", 12; "abbb"
    Print
    Print "123456789012345678901234567890"
    Print 12; 12
    Print
    MsgBox "��Ļ��ӡ����ʲô��������" & Chr(10) & Chr(13) & "���㰴�մ����ݱ�дPRINT���", 1 + 32, "��һ��"
End Sub

Private Sub Command6_Click()
    Cls
    Print "123456789012345678901234567890���ַ���������������ı��������"
    Print "("; Text1.Text; ")"; Text2.Text, Text1.Text; Text2.Text
    Print
    Print "123456789012345678901234567890����ֵ������������ı��������"
    Print Val(Text1.Text); Val(Text2.Text), Val(Text1.Text); Val(Text2.Text)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Form4
    Form1.Show
End Sub
