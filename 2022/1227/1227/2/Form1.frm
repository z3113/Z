VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "��ĩ����"
   ClientHeight    =   4575
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7470
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   4575
   ScaleWidth      =   7470
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command4 
      Caption         =   "�˳�"
      Height          =   735
      Left            =   5640
      TabIndex        =   4
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "�ۺ�Ӧ��"
      Height          =   735
      Left            =   3840
      TabIndex        =   3
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "��Ӧ��"
      Height          =   735
      Left            =   2040
      TabIndex        =   2
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��������"
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "2015ѧ���һѧ��VB��ĩ����"
      BeginProperty Font 
         Name            =   "��Բ"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   840
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
    Unload Form1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If MsgBox("ȷ���رմ�����", vbOKCancel + 16, "�رմ���") = vbCancel Then Cancel = True
End Sub
