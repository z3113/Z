VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0FFFF&
   Caption         =   "for����ѭ����ϰ"
   ClientHeight    =   3510
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   ScaleHeight     =   3510
   ScaleWidth      =   6270
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command3 
      Caption         =   "����˳�"
      Height          =   495
      Left            =   1920
      TabIndex        =   3
      Top             =   2760
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "��ٷ�"
      Height          =   495
      Left            =   3600
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��ϰ����"
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "21��������һ��09��Ԫ��for����ѭ����ϰ"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   6255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Form1.Hide
    Form2.Show
End Sub

Private Sub Command2_Click()
    Form1.Hide
    Form4.Show
End Sub

Private Sub Command3_Click()
    Unload Form1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If MsgBox("ȷ���˳���", vbOKCancel + 32, "�˳���ʾ") = vbCancel Then Cancel = True
End Sub
