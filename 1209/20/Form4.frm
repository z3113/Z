VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "3������ز�����Ϸ"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form4"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      Height          =   495
      Left            =   3360
      TabIndex        =   6
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   1920
      TabIndex        =   5
      Top             =   2040
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1920
      TabIndex        =   4
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1920
      TabIndex        =   3
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "���ۣ�"
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "���أ�kg����"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "��ߣ�cm����"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim a!, b!
    a = Val(Text1.Text)
    b = Val(Text2.Text)
    If a < 0 Or b < 0 Then
        Text3.Text = "����������"
    Else
        If a - b > 110 Or a - b < 100 Then
            Text3.Text = "����Ҫע�Ᵽ������"
        Else
            Text3.Text = "����Ĳ���"
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub
