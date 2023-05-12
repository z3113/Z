VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "素数最值"
   ClientHeight    =   7215
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9030
   LinkTopic       =   "Form4"
   ScaleHeight     =   7215
   ScaleWidth      =   9030
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   10
      Top             =   6360
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   8
      Top             =   6360
      Width           =   1215
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
      Height          =   2895
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   3120
      Width           =   6375
   End
   Begin VB.CommandButton Command5 
      Caption         =   "返回"
      Height          =   495
      Left            =   7440
      TabIndex        =   5
      Top             =   5400
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "清空"
      Height          =   495
      Left            =   7440
      TabIndex        =   4
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "最值"
      Height          =   495
      Left            =   7440
      TabIndex        =   3
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "素数"
      Height          =   495
      Left            =   7440
      TabIndex        =   2
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "产生"
      Height          =   495
      Left            =   7440
      TabIndex        =   1
      Top             =   1080
      Width           =   1335
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
      Height          =   2775
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   6375
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "最小值"
      Height          =   180
      Left            =   3840
      TabIndex        =   9
      Top             =   6480
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "最大值"
      Height          =   180
      Left            =   720
      TabIndex        =   7
      Top             =   6480
      Width           =   540
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim a(100) As Integer, i As Integer

Private Sub Command1_Click()
    Randomize
    Text1.Text = ""
    For i = 1 To 100
        a(i) = Int(Rnd * 900) + 1
        Text1.Text = Text1.Text & a(i) & " "
        If i Mod 10 = 0 Then Text1.Text = Text1.Text & vbCrLf
    Next i
End Sub

Private Sub Command2_Click()
    Dim j%, b%
    Text2.Text = ""
    For i = 1 To 100
        For j = 2 To a(i)
            If a(i) Mod j = 0 Then Exit For
        Next j
        If a(i) = j Then
            Text2.Text = Text2.Text & a(i) & " "
            b = b + 1
            If b Mod 5 = 0 Then Text2.Text = Text2.Text & vbCrLf
        End If
    Next i
End Sub

Private Sub Command3_Click()
    Dim max%, min%
    max = a(1)
    min = a(1)
    For i = 1 To 100
        If max <= a(i) Then max = a(i)
        If min >= a(i) Then min = a(i)
    Next i
    Text3.Text = max
    Text4.Text = min
End Sub

Private Sub Command4_Click()
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
End Sub

Private Sub Command5_Click()
    Unload Form4
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub
