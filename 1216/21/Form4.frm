VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "�ɼ�����"
   ClientHeight    =   4170
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7890
   LinkTopic       =   "Form4"
   ScaleHeight     =   4170
   ScaleWidth      =   7890
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5040
      TabIndex        =   4
      Top             =   1560
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1440
      TabIndex        =   3
      Top             =   2880
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1440
      TabIndex        =   1
      Top             =   960
      Width           =   3135
   End
   Begin VB.Label Label2 
      Caption         =   "��Ӧ�ĵȼ��ǣ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   2160
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "��������ĳɼ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim a!
    a = Val(Text1.Text)
    If a < 60 Then
        Text2.Text = "������"
    ElseIf a < 75 Then
        Text2.Text = "����"
    ElseIf a < 90 Then
        Text2.Text = "����"
    Else
        Text2.Text = "����"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form2.Show
End Sub
