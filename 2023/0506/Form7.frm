VERSION 5.00
Begin VB.Form Form7 
   Caption         =   "�������2"
   ClientHeight    =   4215
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6150
   BeginProperty Font 
      Name            =   "΢���ź�"
      Size            =   15
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form7"
   ScaleHeight     =   4215
   ScaleWidth      =   6150
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command2 
      Caption         =   "����2"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   1
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����1"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   0
      Top             =   1200
      Width           =   1095
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim arr(3) As Integer, i As Integer, j As Integer
    For i = 1 To 3
        For j = 1 To 3
            arr(j) = arr(i) + 1
        Next j
    Next i
    Print "��һ����"
    Print arr(3)
End Sub

Private Sub Command2_Click()
    Dim m(10) As Integer, k As Integer, x As Integer
    For k = 1 To 10
        m(k) = 12 - k
    Next k
    x = 6
    Print "�ڶ�����"
    Print m(2 + m(x))
End Sub

