VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FFC0C0&
   Caption         =   "��������"
   ClientHeight    =   6375
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8205
   LinkTopic       =   "Form2"
   ScaleHeight     =   6375
   ScaleWidth      =   8205
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command7 
      Caption         =   "���أ�&E)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   7
      Top             =   4920
      Width           =   1695
   End
   Begin VB.CommandButton Command6 
      Caption         =   "ɾ��ȫ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   6
      Top             =   4200
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "��ʾѡ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   5
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "��ʾ��2��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   4
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "��ʾ�ܸ���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   3
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "��ǰ��һ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   2
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�����һ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   1
      Top             =   600
      Width           =   1695
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      ItemData        =   "Form2.frx":0000
      Left            =   1800
      List            =   "Form2.frx":0002
      TabIndex        =   0
      Top             =   600
      Width           =   2295
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim a$
    a = InputBox("��������������")
    List1.AddItem a
End Sub

Private Sub Command2_Click()
    Dim a$
    a = InputBox("��������������")
    List1.AddItem a, 0
End Sub

Private Sub Command3_Click()
    Cls
    Print "����"; List1.ListCount; "��������"
End Sub

Private Sub Command4_Click()
    Cls
    Print "�ڶ�����������"; List1.List(1)
End Sub

Private Sub Command5_Click()
    Cls
    If List1.ListIndex = -1 Then
        Print "��û��ѡ����"
    Else
        Print List1.Text
    End If
End Sub

Private Sub Command6_Click()
    List1.Clear
End Sub

Private Sub Command7_Click()
    Unload Form2
End Sub

Private Sub Form_Load()
    Dim i%
    For i = 1 To 9
        List1.AddItem i
    Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub
