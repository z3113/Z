VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FFC0C0&
   Caption         =   "��ֵ����"
   ClientHeight    =   6750
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11190
   LinkTopic       =   "Form2"
   ScaleHeight     =   6750
   ScaleWidth      =   11190
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command4 
      Caption         =   "����"
      Height          =   495
      Left            =   8520
      TabIndex        =   7
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "���"
      Height          =   495
      Left            =   6120
      TabIndex        =   6
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "������"
      Height          =   495
      Left            =   3720
      TabIndex        =   5
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1320
      TabIndex        =   4
      Top             =   5760
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   360
      ScaleHeight     =   3435
      ScaleWidth      =   5115
      TabIndex        =   2
      Top             =   1560
      Width           =   5175
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      TabIndex        =   1
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6480
      TabIndex        =   19
      Top             =   4800
      Width           =   2295
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "���ĸ�����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   18
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6480
      TabIndex        =   17
      Top             =   3480
      Width           =   2295
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��a���ַ��������������ִ�Сд����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   6480
      TabIndex        =   16
      Top             =   3120
      Width           =   3720
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   8400
      TabIndex        =   15
      Top             =   2400
      Width           =   120
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����ַ������ȣ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   6480
      TabIndex        =   14
      Top             =   2400
      Width           =   1800
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   8400
      TabIndex        =   13
      Top             =   1680
      Width           =   120
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "����ַ�����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   12
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   8400
      TabIndex        =   11
      Top             =   960
      Width           =   120
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ַ������ȣ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   6480
      TabIndex        =   10
      Top             =   960
      Width           =   1800
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   8400
      TabIndex        =   9
      Top             =   240
      Width           =   120
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "��ַ�����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   8
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�����ַ�����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      TabIndex        =   3
      Top             =   1200
      Width           =   1800
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "�����ı������벻����10����������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim max%, min%, b%, d$, e$, f$, g%

Private Sub Command1_Click()
    Picture1.Cls
    Dim i%, a$, j%
    b = 0
    d = ""
    e = ""
    g = 0
    max = 0
    min = 9999
    For i = 1 To Val(Text1.Text)
        a = InputBox("�������" & i & "���ַ���", "����", "aaaaa")
        Picture1.Print a
        If Len(a) <= min Then min = Len(a): e = a
        If Len(a) >= max Then max = Len(a): d = a
        For j = 1 To Len(a)
            If Mid(a, j, 1) = Chr(97) Then
                b = b + 1
                Exit For
            End If
        Next j
        For j = Len(a) To 1 Step -1
            f = f & Mid(a, j, 1)
        Next j
        Picture1.Print f
        If f = a Then g = g + 1
        f = ""
    Next i
End Sub

Private Sub Command2_Click()
    Label4.Caption = d
    Label6.Caption = max
    Label8.Caption = e
    Label10.Caption = min
    Label12.Caption = b
    Label14.Caption = g
End Sub

Private Sub Command3_Click()
    Text1.Text = ""
    Picture1.Cls
    Label4.Caption = ""
    Label6.Caption = ""
    Label8.Caption = ""
    Label10.Caption = ""
    Label12.Caption = ""
    Label14.Caption = ""
End Sub

Private Sub Command4_Click()
    Unload Form2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub

Private Sub Text1_Click()
    Dim n%
    n = Val(InputBox("������һ��������10��������", "����", 15))
    If n > 10 Or n <= 0 Then
        MsgBox "����������ݲ�����Ҫ��", vbOKCancel + 16, "������ʾ"
        Command1.Enabled = False
    Else
        Text1.Text = n
        Command1.Enabled = True
    End If
End Sub
