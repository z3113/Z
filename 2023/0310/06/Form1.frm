VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "forѭ����ϰ5"
   ClientHeight    =   5325
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7815
   BeginProperty Font 
      Name            =   "����"
      Size            =   10.5
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5325
   ScaleWidth      =   7815
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command7 
      Caption         =   "�˳�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   7
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   6
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   5
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   4
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   3
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   2
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����һ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   1
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   3000
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   4575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Cls
    Dim i%, a&
    Text1.Text = "1����1+3^2+3^3+3^4+3^5+����+3^10 ֮������Ϣ�������"
    a = 1
    For i = 2 To 10
        a = a + 3 ^ i
    Next i
    Print "1+3^2+3^3+3^4+3^5+����+3^10="; a
    MsgBox "1+3^2+3^3+3^4+3^5+����+3^10=" & a, vbOKCancel + 64, "��ͽ��"
End Sub

Private Sub Command2_Click()
    Cls
    Dim i%, a%, b#
    b = 1
    Text1.Text = "2��inputbox��������n������������е�ǰn��֮����"
    a = Val(InputBox("������n��ֵ", "����", 10))
    For i = 1 To a
        b = b * 2 ^ (i - 1) / (i + 1)
    Next i
    Print "n="; a; "��="; b
End Sub

Private Sub Command3_Click()
    Cls
    Dim i%, n%, a%, b%, c%, d&
    a = 1
    b = 1
    d = 2
    Text1.Text = "3����һ���У�a1,a2,a3��an,����a1=1,a2=1,ai��������ai=ai-1+ai-2���Ӽ����������������n���ֵ��ǰn��ĺ͡�"
    n = Val(InputBox("������n��ֵ", "����", 10))
    If n >= 3 Then
        Print a; b;
        For i = 3 To n
            c = a + b
            d = d + c
            Print c;
            If i Mod 3 = 0 Then Print
            a = b
            b = c
        Next i
    Else
        MsgBox "�����벻С��3��������", vbOKCancel + 48, "��ܰ��ʾ��"
    End If
    Print "��"; n; "���ֵΪ��"; c
    Print "ǰ"; n; "��ĺ�Ϊ��"; d
End Sub

Private Sub Command4_Click()
    Cls
    Dim i%, a%, b%, min%, max%, e&
    min = 100
    max = 0
    Text1.Text = "4�� �ж�ȫ��ͬѧ�ĳɼ��ȼ����༶��������������룩"
    a = Val(InputBox("������༶����", "��������", 10))
    For i = 1 To a
        b = InputBox("�������" & i & "��ͬѧ�ĳɼ�", "����ɼ�", 60)
        If b < 60 Then
            Print "��"; i; "��ͬѧ�ķ����ǣ�"; b
            Print "�ɼ��ȼ�Ϊ��������"
        ElseIf b < 80 Then
            Print "��"; i; "��ͬѧ�ķ����ǣ�"; b
            Print "�ɼ��ȼ�Ϊ������"
        ElseIf b < 90 Then
            Print "��"; i; "��ͬѧ�ķ����ǣ�"; b
            Print "�ɼ��ȼ�Ϊ������"
        Else
            Print "��"; i; "��ͬѧ�ķ����ǣ�"; b
            Print "�ɼ��ȼ�Ϊ������"
        End If
        e = e + b
        If b >= max Then max = b
        If b <= min Then min = b
    Next i
    Print "�༶��߷�Ϊ��"; max
    Print "�༶��ͷ�Ϊ��"; min
    Print "�༶ƽ����Ϊ��"; e
End Sub

Private Sub Command5_Click()
    Cls
    Dim i%, a%, b$, c$, d$
    d = ""
    Text1.Text = "5����������N���ַ�������ӡԭ�ַ���"
    b = InputBox("������һ���ַ���", "����", "abcdef")
    a = Len(b)
    For i = a To 1 Step -1
        c = Mid(b, i, 1)
        d = d & c
    Next i
    Print "ԭ�ַ���Ϊ��"; b
    Print "���ú���ַ���Ϊ��"; d
End Sub

Private Sub Command6_Click()
    Cls
    Dim i%, a!, b!
    a = 200
    b = -200
    Text1.Text = "5��һС��� 200 �׸߶��������䣬ÿ����غ󷴵�ԭ�߶ȵ�һ�룬Ȼ�������¡��������С���ʮ�����ʱ�������˶����׵�·�̣�"
    For i = 1 To 10
        b = b + 2 * a
        a = a / 2
    Next i
    Print "��10����ؾ�����"; b; "��"
End Sub

Private Sub Command7_Click()
    Unload Form1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If MsgBox("ȷ���˳���", vbOKCancel + 64, "�˳���ʾ") = vbCancel Then Cancel = True
End Sub
