VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "ͼ�δ�ӡ1"
   ClientHeight    =   4350
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8055
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
   LinkTopic       =   "Form1"
   ScaleHeight     =   4350
   ScaleWidth      =   8055
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command5 
      Caption         =   "��һ������"
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
      Left            =   5760
      TabIndex        =   4
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "��ӡ������2"
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
      Left            =   4320
      TabIndex        =   3
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "��ӡ������1"
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
      Left            =   2880
      TabIndex        =   2
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "��ӡ2"
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
      Left            =   1560
      TabIndex        =   1
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��ӡ1"
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
      Left            =   240
      TabIndex        =   0
      Top             =   3720
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Cls
    Dim i%, a%
    Print "123456789012345678901234567890123456789012345678901234567890"
    a = InputBox("����Ҫ��ӡ��*����", "��������", 5)
    For i = 1 To a
        Print "*";
    Next i
End Sub

Private Sub Command2_Click()
    Cls
    Dim i%, j%, a%, b%
    Print "123456789012345678901234567890123456789012345678901234567890"
    a = InputBox("����Ҫ��ӡ������", "��������", 6)
    b = InputBox("����Ҫ��ӡ�ĸ���", "��������", 5)
    For i = 1 To a
        For j = 1 To b
            Print "A";
        Next j
        Print
    Next i
End Sub

Private Sub Command3_Click()
    Cls
    Dim i%, j%
    Print "123456789012345678901234567890123456789012345678901234567890"
    For i = 1 To 5
        For j = 1 To i
            Print "*";
        Next j
        Print
    Next i
End Sub

Private Sub Command4_Click()
    Cls
    Dim i%, j%
    Print "123456789012345678901234567890123456789012345678901234567890"
    For i = 5 To 1 Step -1
        For j = i To 1 Step -1
            Print "*";
        Next j
        Print
    Next i
End Sub

Private Sub Command5_Click()
    Form2.Show
End Sub
