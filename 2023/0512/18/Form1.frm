VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0E0FF&
   Caption         =   "����2��Ӧ��"
   ClientHeight    =   4095
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7605
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   7605
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFC0FF&
      Caption         =   "�˳�&Q"
      Height          =   495
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFC0FF&
      Caption         =   "����ͳ��&E"
      Height          =   495
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFC0FF&
      Caption         =   "�ɼ�ͳ��&D"
      Height          =   495
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFC0FF&
      Caption         =   "������ֵ&C"
      Height          =   495
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0FF&
      Caption         =   "����ֵ&B"
      Height          =   495
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0FF&
      Caption         =   "����&A"
      Height          =   495
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1680
      Width           =   1335
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

Private Sub Command6_Click()
    Unload Form1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If MsgBox("�Ƿ��˳���", vbYesNo + 48, "�˳�") = vbNo Then Cancel = True
End Sub

