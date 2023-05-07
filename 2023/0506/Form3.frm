VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "平均值处理"
   ClientHeight    =   7695
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8175
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form3"
   ScaleHeight     =   7695
   ScaleWidth      =   8175
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command4 
      Caption         =   "返回"
      Height          =   495
      Left            =   5520
      TabIndex        =   5
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "超平均"
      Height          =   495
      Left            =   4200
      TabIndex        =   4
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "平均"
      Height          =   495
      Left            =   2880
      TabIndex        =   3
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "产生"
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   6600
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   480
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   3960
      Width           =   7215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   480
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   1440
      Width           =   7215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   0
      TabIndex        =   6
      Top             =   360
      Width           =   8175
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim a(50) As Integer, i As Integer, b As Single, c As Integer
Private Sub Command1_Click()
    Randomize
    Text1.Text = ""
    For i = 1 To 50
        a(i) = Int(Rnd * (90) + 10)
        Text1.Text = Text1.Text & a(i) & " "
        If i Mod 10 = 0 Then Text1.Text = Text1.Text & vbCrLf
    Next i
End Sub

Private Sub Command2_Click()
    b = 0
    For i = 1 To 50
        b = b + a(i)
    Next i
    b = Format(b / 50, "0.00")
    'b = Round(b / 50, 2)
    Label1.Caption = "平均值为 " & b
End Sub

Private Sub Command3_Click()
    Text2.Text = ""
    c = 0
    For i = 1 To 50
        If a(i) > b Then
            Text2.Text = Text2.Text & a(i) & " "
            c = c + 1
            If c Mod 10 = 0 Then Text2.Text = Text2.Text & vbCrLf
        End If
    Next i
End Sub

Private Sub Command4_Click()
    Unload Form3
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub

