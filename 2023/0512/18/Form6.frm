VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "����ͳ��"
   ClientHeight    =   6015
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6975
   LinkTopic       =   "Form6"
   ScaleHeight     =   6015
   ScaleWidth      =   6975
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command4 
      Caption         =   "����"
      Height          =   495
      Left            =   5400
      TabIndex        =   9
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "���"
      Height          =   495
      Left            =   5400
      TabIndex        =   8
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "���ĸ���"
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "����ͳ��"
      Height          =   495
      Left            =   5400
      TabIndex        =   6
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�����ַ���"
      Height          =   495
      Left            =   5400
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "�����������"
      Top             =   360
      Width           =   1215
   End
   Begin VB.PictureBox Picture2 
      Height          =   2055
      Left            =   120
      ScaleHeight     =   1995
      ScaleWidth      =   4875
      TabIndex        =   2
      Top             =   3840
      Width           =   4935
   End
   Begin VB.PictureBox Picture1 
      Height          =   3135
      Left            =   120
      ScaleHeight     =   3075
      ScaleWidth      =   4875
      TabIndex        =   0
      Top             =   360
      Width           =   4935
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "�ַ�����"
      Height          =   180
      Left            =   120
      TabIndex        =   3
      Top             =   3600
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "�ַ�����"
      Height          =   180
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   720
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim a(20) As String, i As Integer

Private Sub Command1_Click()
    Picture1.Cls
    For i = 1 To Val(Text1.Text)
        a(i) = InputBox("�������" & i & "���ַ���", "����", "aaaaa")
        Picture1.Print a(i)
    Next i
End Sub

Private Sub Command2_Click()
    Dim b$, c%, j%
    For i = 1 To Val(Text1.Text)
        b = ""
        For j = Len(a(i)) To 1 Step -1
            b = b & Mid(a(i), j, 1)
        Next j
        If b = a(i) Then
            c = c + 1
            Picture2.Print a(i)
        End If
    Next i
    Text2.Text = c
End Sub

Private Sub Command3_Click()
    Picture1.Cls
    Picture2.Cls
    Text1.Text = "�����������"
    Text2.Text = "���ĸ���"
End Sub

Private Sub Command4_Click()
    Unload Form6
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub

Private Sub Text1_Click()
    Dim b%
    b = Val(InputBox("������Ҫ������ַ�������", "�������", 10))
    If 1 <= b And b <= 20 Then
        Text1.Text = b
        Text1.SetFocus
    Else
        MsgBox "����������ݲ�����Ҫ��", vbOKCancel + 16, "������ʾ"
        Text1.Text = ""
        Text1.SetFocus
    End If
End Sub

