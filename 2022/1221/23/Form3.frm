VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "��ϰ�����ķ�֧����"
   ClientHeight    =   4410
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7140
   LinkTopic       =   "Form3"
   ScaleHeight     =   4410
   ScaleWidth      =   7140
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command4 
      Caption         =   "����������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   4
      Top             =   3480
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Select Case���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   3
      Top             =   3480
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ELSEIF���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   2
      Top             =   2640
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��ͨIF���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form3.frx":0000
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1575
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   3735
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim a!
    a = Val(InputBox("������һ���ɼ�", "����ɼ�����ȵ�", 95))
    If a >= 90 Then Print "����"
    If a >= 75 Then Print "����"
    If a >= 60 Then Print "�ϸ�" Else Print "������"
End Sub

Private Sub Command2_Click()
    Dim a!
    a = Val(InputBox("������һ���ɼ�", "����ɼ�����ȵ�", 68))
    If a >= 90 Then
        Print "����"
    ElseIf a >= 75 Then
        Print "����"
    ElseIf a >= 60 Then
        Print "�ϸ�"
    Else
        Print "������"
    End If
End Sub

Private Sub Command3_Click()
    Dim a!
    a = Val(InputBox("������һ���ɼ�", "����ɼ�����ȵ�", 50))
    Select Case a
        Case Is >= 90
            Print "����"
        Case Is >= 75
            Print "����"
        Case Is >= 60
            Print "����"
        Case Is < 60
            Print "������"
    End Select
End Sub

Private Sub Command4_Click()
    Unload Form3
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub
