VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "����һ"
   ClientHeight    =   4635
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9405
   LinkTopic       =   "Form2"
   ScaleHeight     =   4635
   ScaleWidth      =   9405
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   6120
      TabIndex        =   12
      Text            =   "��������"
      Top             =   240
      Width           =   3015
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00C0C0FF&
      Caption         =   "���Ӵ�"
      Height          =   495
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00FFFF80&
      Caption         =   "�Ӵ�1"
      Height          =   495
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H0000FFFF&
      Caption         =   "�Ӵ�2"
      Height          =   495
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Caption         =   "ɾ����"
      Height          =   495
      Left            =   7920
      TabIndex        =   8
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "�»���"
      Height          =   495
      Left            =   6240
      TabIndex        =   7
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Cancel          =   -1  'True
      Caption         =   "���أ�ESC��"
      Height          =   495
      Left            =   4680
      TabIndex        =   6
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "����˶�"
      Height          =   495
      Left            =   4680
      TabIndex        =   5
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "�ı���"
      Height          =   495
      Left            =   4680
      TabIndex        =   4
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "�ı�����"
      Height          =   495
      Left            =   4680
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�����Ϣ"
      Height          =   495
      Left            =   4680
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "������Ϣ"
      Height          =   495
      Left            =   4680
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
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
      Height          =   3255
      Left            =   240
      Picture         =   "Form2.frx":0000
      ScaleHeight     =   3195
      ScaleWidth      =   4155
      TabIndex        =   0
      Top             =   240
      Width           =   4215
      Begin VB.Shape Shape1 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000000FF&
         Height          =   855
         Left            =   1320
         Shape           =   3  'Circle
         Top             =   1320
         Width           =   855
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a$, b$
Option Explicit

Private Sub Command1_Click()
    a = InputBox("�������ַ��Ϣ", "��Ϣ¼��", "�㽭����")
    b = InputBox("������������Ϣ", "��Ϣ¼��", "����")
End Sub

Private Sub Command10_Click()
    Text1.FontBold = True
    Command10.Visible = False
    Command11.Visible = True
End Sub

Private Sub Command11_Click()
    Text1.FontBold = False
    Command10.Visible = True
    Command11.Visible = False
End Sub

Private Sub Command2_Click()
    Picture1.Cls
    Picture1.Print "����ĵ�ַ�ǣ�" & a
    Picture1.Print "����������ǣ�" & b
End Sub

Private Sub Command3_Click()
    Picture1.Print Text1.Text
End Sub

Private Sub Command4_Click()
    Picture1.Picture = LoadPicture("")
End Sub

Private Sub Command5_Click()
    Randomize
    Shape1.Move Rnd * (Picture1.Width - Shape1.Width + 1), Rnd * (Picture1.Height - Shape1.Height + 1)
End Sub

Private Sub Command6_Click()
    Unload Form2
End Sub

Private Sub Command7_Click()
    Text1.FontUnderline = Not Text1.FontUnderline
End Sub

Private Sub Command8_Click()
    Text1.FontStrikethru = Not Text1.FontStrikethru
End Sub

Private Sub Command9_Click()
    Text1.FontBold = Not Text1.FontBold
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub
