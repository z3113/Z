VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFC0C0&
   Caption         =   "��Ͽ���ϰ"
   ClientHeight    =   4815
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9510
   LinkTopic       =   "Form1"
   ScaleHeight     =   4815
   ScaleWidth      =   9510
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command5 
      Caption         =   "�˳�"
      Height          =   615
      Left            =   7800
      TabIndex        =   4
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "�ۺ�Ӧ��"
      Height          =   615
      Left            =   5880
      TabIndex        =   3
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "����Ӧ��"
      Height          =   615
      Left            =   3960
      TabIndex        =   2
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "��Ӧ��"
      Height          =   615
      Left            =   2040
      TabIndex        =   1
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��ʶ��ʽ"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   2640
      Width           =   1575
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
    Unload Form1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If MsgBox("�Ƿ��˳���", vbYesNo + 32, "�˳�") = vbNo Then Cancel = True
End Sub
    
