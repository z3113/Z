VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "����������ֳ���"
   ClientHeight    =   5055
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8055
   LinkTopic       =   "Form3"
   ScaleHeight     =   5055
   ScaleWidth      =   8055
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command2 
      Caption         =   "����"
      Height          =   495
      Left            =   5520
      TabIndex        =   9
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      Height          =   495
      Left            =   5520
      TabIndex        =   8
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   4200
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   3360
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   2520
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   720
      Width           =   6855
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "���÷֣�"
      Height          =   180
      Left            =   480
      TabIndex        =   6
      Top             =   4200
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "ȥ��һ����ͷ֣�"
      Height          =   180
      Left            =   480
      TabIndex        =   4
      Top             =   3360
      Width           =   1440
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "ȥ��һ����߷֣�"
      Height          =   180
      Left            =   480
      TabIndex        =   2
      Top             =   2520
      Width           =   1440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "����ί���"
      Height          =   180
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   900
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim i%, a#, b#, c#, d#, min#, max#
    min = 999
    max = 0
    Text1.Text = ""
    For i = 1 To 10
        a = Val(InputBox("�������" & a & "����ί�Ĵ��", "���ִ��", 0))
        Text1.Text = Text1.Text & " " & a
        If a >= max Then max = a
        If a <= min Then min = a
        b = b + a
    Next i
    Text2.Text = max
    Text3.Text = min
    Text4.Text = Round((b - max - min) / 8, 3)
End Sub

Private Sub Command2_Click()
    Unload Form3
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form2.Show
End Sub
