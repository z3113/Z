VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "���������ַ�������"
   ClientHeight    =   6645
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6525
   LinkTopic       =   "Form4"
   ScaleHeight     =   6645
   ScaleWidth      =   6525
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command4 
      Cancel          =   -1  'True
      Caption         =   "���أ�ESC��"
      Height          =   495
      Left            =   4680
      TabIndex        =   7
      Top             =   5640
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   360
      TabIndex        =   6
      Text            =   "˫����ȡ�����ַ�"
      Top             =   4560
      Width           =   5775
   End
   Begin VB.CommandButton Command3 
      Caption         =   "��ʾ����ַ����������ַ����е�λ��"
      Height          =   495
      Left            =   1680
      TabIndex        =   5
      Top             =   3480
      Width           =   4455
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Text            =   "ABC"
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "������ʼλ�ú͸�������ȡ��Ӧ�ִ�"
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   2400
      Width           =   5775
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   1680
      Width           =   5775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "������ʼλ�ú͸�����ѡ�ж�Ӧ�ִ�"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   5775
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Text            =   "1234567890ABCDEFGHIJK"
      Top             =   240
      Width           =   5775
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   3
      X1              =   120
      X2              =   6360
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   3
      X1              =   120
      X2              =   6360
      Y1              =   3120
      Y2              =   3120
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim a%, b%
    a = InputBox("������ѡȡ����ʼλ�ã�", "��������", 3)
    b = InputBox("������ѡȡ�ĸ�����", "��������", 5)
    Text1.SetFocus
    Text1.SelStart = a
    Text1.SelLength = b
End Sub

Private Sub Command2_Click()
    Dim a%, b%
    a = InputBox("�������ȡ����ʼλ�ã�", "��������", 3)
    b = InputBox("�������ȡ�ĸ�����", "��������", 5)
    Text2.Text = Mid(Text1.Text, a, b)
End Sub

Private Sub Command3_Click()
    MsgBox "�ַ�����" & Text3.Text & "���ַ���" & Text1.Text & "�ĵ�" & InStr(Text1.Text, Text3.Text) & "��λ��"
    
End Sub

Private Sub Command4_Click()
    Unload Form4
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub

Private Sub Text4_DblClick()
    Text4.Text = StrReverse(Text1.Text)
End Sub
