VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00C0E0FF&
   Caption         =   "�ۺ�Ӧ��"
   ClientHeight    =   5325
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8775
   LinkTopic       =   "Form4"
   ScaleHeight     =   5325
   ScaleWidth      =   8775
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command2 
      Caption         =   "�����Сֵ��λ��"
      Height          =   495
      Left            =   5040
      TabIndex        =   5
      Top             =   4320
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�������"
      Height          =   495
      Left            =   1560
      TabIndex        =   4
      Top             =   4320
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   2040
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   2160
      Width           =   6015
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   2040
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   6015
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��Сֵ��λ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   480
      TabIndex        =   3
      Top             =   2760
      Width           =   1170
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   480
      TabIndex        =   1
      Top             =   960
      Width           =   585
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim min%, d$

Private Sub Command1_Click()
    Randomize
    Text1.Text = ""
    Text2.Text = ""
    Dim i%, a%, b%, c%
    a = Val(InputBox("������һ����Сֵ", "����", 10))
    b = Val(InputBox("������һ�����ֵ", "����", 20))
    min = b
    d = ""
    For i = 1 To 30
        c = Int(Rnd * (b - a + 1) + a)
        Text1.Text = Text1.Text & c & " "
        If i Mod 10 = 0 Then Text1.Text = Text1.Text & vbCrLf
        If (min = c) Then d = d & " " & i
        If min > c Then min = c: d = i
    Next i
End Sub

Private Sub Command2_Click()
    Text2.Text = "��Сֵ " & min & vbCrLf & "λ�� " & d
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub
