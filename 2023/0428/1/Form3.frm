VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "ѭ���ṹ"
   ClientHeight    =   5430
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6165
   LinkTopic       =   "Form3"
   ScaleHeight     =   5430
   ScaleWidth      =   6165
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   3000
      TabIndex        =   6
      Top             =   2880
      Width           =   3015
   End
   Begin VB.CommandButton Command6 
      Caption         =   "������"
      Height          =   495
      Left            =   4560
      TabIndex        =   5
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "������"
      Height          =   495
      Left            =   2400
      TabIndex        =   4
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "������"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "������"
      Height          =   495
      Left            =   4560
      TabIndex        =   2
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�ڶ���"
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��һ��"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "����������"
      Height          =   180
      Left            =   1920
      TabIndex        =   7
      Top             =   3120
      Width           =   900
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Cls
    Dim i%, a%, b!
    a = 1
    For i = 1 To 200
        b = b + a / i
        a = -a
    Next i
    Print "ǰ200��֮��Ϊ" & b
End Sub

Private Sub Command2_Click()
    Cls
    Dim i%, a$, b$
    a = InputBox("������һ���ַ���")
    For i = Len(a) To 1 Step -1
        b = b & Mid(a, i, 1)
    Next i
    If b = a Then
        Print a & "�ǻ���"
    Else
        Print a & "���ǻ���"
    End If
End Sub

Private Sub Command3_Click()
    Cls
    Dim i%, a$, b%, c%, d%, e%
    For i = 1 To Len(Text1.Text)
        a = Mid(Text1.Text, i, 1)
        If "A" <= a And a <= "Z" Then
            b = b + 1
        ElseIf "A" <= a And a <= "z" Then
            c = c + 1
        ElseIf "0" <= a And a <= "9" Then
            d = d + 1
        Else
            e = e + 1
        End If
    Next i
    Print "��дӢ�ĸ���" & b
    Print "СдӢ�ĸ���" & c
    Print "���ָ���" & d
    Print "�����ַ�����" & e
End Sub

Private Sub Command4_Click()
    Cls
    Dim i%, a$, b%
    For i = 1 To 10
        a = InputBox("�������" & i & "���ַ���")
        If Mid(a, 1, 1) = "D" Then b = b + 1
    Next i
    Print "����ĸD��ͷ�ĵ�����" & b & "��"
End Sub

Private Sub Command5_Click()
    Cls
    Dim i%
    For i = 100 To 999
        If (i \ 100) ^ 3 + (i \ 10 Mod 10) ^ 3 + (i Mod 10) ^ 3 = i Then
            Print i;
        End If
    Next i
End Sub

Private Sub Command6_Click()
    Dim i%, j%, k%, a%, b%
    For i = 1 To 9
        For j = 0 To 9
            For k = 0 To 9
                a = i & j & k
                If (i = 2 Or j = 2 Or k = 2) And a Mod 9 = 0 Then
                    Print a,
                    b = b + 1
                    If b Mod 7 = 0 Then Print
                End If
            Next k
        Next j
    Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub
