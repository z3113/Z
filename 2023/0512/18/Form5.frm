VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "�ɼ�ͳ��"
   ClientHeight    =   4230
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8790
   LinkTopic       =   "Form5"
   ScaleHeight     =   4230
   ScaleWidth      =   8790
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command2 
      Caption         =   "ͳ��"
      Height          =   495
      Left            =   7440
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����ɼ�"
      Height          =   495
      Left            =   7440
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "5��������ťCommand1���û������������������10��ѧ���ĳɼ���������ťCommand2��ͳ������ѧ����ƽ���֣�����ʾ��ͷּ���Щ�˵�λ�á�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   3600
      Width           =   8535
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim a(10) As Integer, i As Integer

Private Sub Command1_Click()
    Cls
    Print "ʮ��ͬѧ�ĳɼ����£�"
    For i = 1 To 10
        a(i) = Val(InputBox("�������" & i & "��ͬѧ�ĳɼ�", "����ɼ�"))
        Print a(i);
    Next i
End Sub

Private Sub Command2_Click()
    Dim b%, min%
    min = a(1)
    For i = 1 To 10
        b = b + a(i)
        If min >= a(i) Then min = a(i)
    Next i
    Print "ƽ����Ϊ��"; Round(b / 10, 1); "��"
    Print "��ͷ��ǣ�"; min; "�֣����ǵ�";
    For i = 1 To 10
        If a(i) = min Then Print i;
    Next i
    Print "��ͬѧ"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub
