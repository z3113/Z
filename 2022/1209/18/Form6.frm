VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "����"
   ClientHeight    =   4875
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8655
   LinkTopic       =   "Form6"
   ScaleHeight     =   4875
   ScaleWidth      =   8655
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command3 
      Caption         =   "��С���������������"
      Height          =   495
      Left            =   480
      TabIndex        =   6
      Top             =   4200
      Width           =   7695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "����ƻ���۸񣡣���IF��"
      Height          =   495
      Left            =   2160
      TabIndex        =   5
      Top             =   2400
      Width           =   6135
   End
   Begin VB.TextBox Text1 
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
      Left            =   240
      TabIndex        =   4
      Text            =   "2"
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�����Ի�������ɼ������������ж����öԻ�����Ƿ�ϸ񣿣���IF��"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   8175
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   8520
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   8520
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label3 
      Caption         =   "���������������ֱ����x��y�����У��Ƚ����Ǵ�С��Ȼ�󽫴�������x�У�С������y�С�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   3480
      Width           =   8055
   End
   Begin VB.Label Label2 
      Caption         =   $"Form6.frx":0000
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1095
      Left            =   2040
      TabIndex        =   2
      Top             =   960
      Width           =   6255
   End
   Begin VB.Label Label1 
      Caption         =   "ƻ��������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   1335
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim a%
    a = Val(InputBox("������ɼ�(����)", "����ɼ�", 60))
    If a >= 60 Then
        MsgBox "��ϲ��ϸ��ˣ�"
    Else
        MsgBox "���ź�����û�кϸ�"
    End If
End Sub

Private Sub Command2_Click()
    Dim a!
    a = Val(Text1.Text)
    If a < 2 Then MsgBox "������" & a & "ǧ��ƻ��������" & a * 1.5 & "Ԫ" Else MsgBox "������" & a & "ǧ��ƻ��������" & a * 1.5 * 0.8 & "Ԫ"
End Sub

Private Sub Command3_Click()
    Dim x%, y%, z%
    x = InputBox("��Ϊx��ֵ", "��������", 100)
    y = InputBox("��Ϊy��ֵ", "��������", 200)
    If x < y Then
        z = x
        x = y
        y = z
    End If
    MsgBox "������Ϊ" & y & "��" & x
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form2.Show
End Sub
