VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "��Ǯ"
   ClientHeight    =   6735
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6915
   LinkTopic       =   "Form6"
   ScaleHeight     =   6735
   ScaleWidth      =   6915
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "���������ӡ��"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   5040
      TabIndex        =   1
      Top             =   5160
      Width           =   1440
   End
   Begin VB.Label Label1 
      Caption         =   $"Form6.frx":0000
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   5040
      TabIndex        =   0
      Top             =   960
      Width           =   1455
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
    Cls
    Dim i%, j%, k%, l%, a%
    For i = 1 To 7
        For j = 1 To 17
            For k = 1 To 37
                l = 40 - i - j - k
                If l > 0 And i * 10 + j * 5 + k * 2 + l = 100 Then Print "10Ԫ" & i & "��", "5Ԫ" & j & "��", "2Ԫ" & k & "��", "1Ԫ" & l & "��": a = a + 1
            Next k
        Next j
    Next i
    Print "����" & a & "�ֻ���"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form4.Show
End Sub
