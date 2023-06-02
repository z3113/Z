VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00C0C0FF&
   Caption         =   "排序任务一"
   ClientHeight    =   5145
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   ScaleHeight     =   5145
   ScaleWidth      =   6870
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command4 
      Caption         =   "返回"
      Height          =   495
      Left            =   2760
      TabIndex        =   6
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "产生"
      Height          =   495
      Left            =   2760
      TabIndex        =   5
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   4560
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   600
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   600
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   600
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "升序（选择）"
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "降序（冒泡）"
      Height          =   495
      Left            =   2760
      TabIndex        =   0
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "单击产生按钮，随机产生15个【1,999】的正整数"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a(15) As Integer

Private Sub Command1_Click()
    Dim i%
    Randomize
    Text1.Text = ""
    For i = 1 To 15
        a(i) = Int(Rnd * 999 + 1)
        Text1.Text = Text1.Text & a(i) & vbCrLf
    Next i
End Sub

Private Sub Command2_Click()
    Dim i%, j%, b%
    Text2.Text = ""
    For i = 1 To 14
        For j = 1 To 15 - i
            If a(j) < a(j + 1) Then b = a(j): a(j) = a(j + 1): a(j + 1) = b
        Next j
    Next i
    For i = 1 To 15
        Text2.Text = Text2.Text & a(i) & vbCrLf
    Next i
End Sub

Private Sub Command3_Click()
    Dim i%, j%, b%, c%
    Text2.Text = ""
    For i = 1 To 14
        b = i
        For j = i + 1 To 15
            If a(j) < a(b) Then b = j
        Next j
        If b <> i Then c = a(i): a(i) = a(b): a(b) = c
    Next i
    For i = 1 To 15
        Text2.Text = Text2.Text & a(i) & vbCrLf
    Next i
End Sub

Private Sub Command4_Click()
    Unload Form2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub
