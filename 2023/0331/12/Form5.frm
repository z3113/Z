VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "��Ǯ��ټ�"
   ClientHeight    =   3570
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7605
   LinkTopic       =   "Form5"
   ScaleHeight     =   3570
   ScaleWidth      =   7605
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command1 
      Caption         =   "��"
      Height          =   495
      Left            =   3120
      TabIndex        =   1
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "1������ÿֻ5Ԫ��ĸ��ÿֻ3Ԫ��С����ֻһԪ����100Ԫ��100ֻ����ÿ�ּ�������һֻ�����Ը���ֻ��"
      Height          =   735
      Left            =   600
      TabIndex        =   0
      Top             =   1680
      Width           =   6375
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Cls
    Print "���������ַ�����"
    Print "����", "ĸ��", "С��"
    Dim i%, j%
    For i = 1 To 18
        For j = 1 To 32
            If i * 5 + j * 3 + (100 - i - j) / 3 = 100 Then Print i, j, (100 - i - j)
        Next j
    Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form4.Show
End Sub
