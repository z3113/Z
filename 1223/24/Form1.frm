VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFC0C0&
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   ScaleHeight     =   3255
   ScaleWidth      =   7215
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command5 
      Caption         =   "������"
      Height          =   495
      Left            =   5880
      TabIndex        =   5
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "������"
      Height          =   495
      Left            =   4440
      TabIndex        =   4
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "������"
      Height          =   495
      Left            =   3000
      TabIndex        =   3
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�����"
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����һ"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "��ѡ��ѡ��ť�����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   7215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Form1.Hide
    Form2.Show
End Sub

Private Sub Command2_Click()
    Form1.Hide
    Form3.Show
End Sub

Private Sub Command3_Click()
    Form1.Hide
    Form4.Show
End Sub

Private Sub Command4_Click()
    Form1.Hide
    Form5.Show
End Sub

Private Sub Command5_Click()
    Form1.Hide
    Form6.Show
End Sub
