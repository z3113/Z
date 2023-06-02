VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00C0C0FF&
   Caption         =   "排序任务三"
   ClientHeight    =   8070
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9615
   LinkTopic       =   "Form1"
   ScaleHeight     =   8070
   ScaleWidth      =   9615
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "清除"
      Height          =   495
      Left            =   7440
      TabIndex        =   11
      Top             =   240
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      Height          =   1455
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   6360
      Width           =   6735
   End
   Begin VB.CommandButton Command5 
      Caption         =   "低于平均降序（冒泡）"
      Enabled         =   0   'False
      Height          =   615
      Left            =   7440
      TabIndex        =   9
      Top             =   6840
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "升序排列(选择）"
      Enabled         =   0   'False
      Height          =   735
      Left            =   7440
      TabIndex        =   8
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "求平均及小于平均的数"
      Enabled         =   0   'False
      Height          =   735
      Left            =   7440
      TabIndex        =   7
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "产生"
      Height          =   615
      Left            =   7440
      TabIndex        =   6
      Top             =   1440
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   1455
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   2760
      Width           =   6735
   End
   Begin VB.TextBox Text3 
      Height          =   1455
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   4440
      Width           =   6735
   End
   Begin VB.TextBox Text1 
      Height          =   1095
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1200
      Width           =   6735
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      LargeChange     =   5
      Left            =   2880
      Max             =   50
      Min             =   1
      SmallChange     =   2
      TabIndex        =   0
      Top             =   120
      Value           =   1
      Width           =   3255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "产生[20,90]之间的整数："
      Height          =   375
      Left            =   480
      TabIndex        =   12
      Top             =   840
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   5
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "产生数的个数"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a(50) As Integer

Private Sub Command1_Click()
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    HScroll1.Value = 1
    Command3.Enabled = False
    Command4.Enabled = False
    Command5.Enabled = False
End Sub

Private Sub Command2_Click()
    Dim i%
    Text1.Text = ""
    For i = 1 To Val(Label2.Caption)
        a(i) = Int(Rnd * 71 + 20)
        Text1.Text = Text1.Text & " " & a(i)
        If i Mod 15 = 0 Then Text1.Text = Text1.Text & vbCrLf
    Next i
End Sub

Private Sub Command3_Click()
    Dim i%, j%, b%
    Text2.Text = ""
    For i = 1 To Val(Label2.Caption) - 1
        b = i
        For j = i + 1 To Val(Label2.Caption)
            If a(j) < a(b) Then b = j
        Next j
        If i <> b Then a(0) = a(i): a(i) = a(b): a(b) = a(0)
    Next i
    For i = 1 To Val(Label2.Caption)
        Text2.Text = Text2.Text & " " & a(i)
        If i Mod 15 = 0 Then Text2.Text = Text2.Text & vbCrLf
    Next i
End Sub

Private Sub Command4_Click()
    Dim i%, b%, c%
    For i = 1 To Val(Label2.Caption)
        b = b + a(i)
    Next i
    b = Round(b / Val(Label2.Caption), 2)
    Text3.Text = " 平均值为：" & b & vbCrLf
    For i = 1 To Val(Label2.Caption)
        Text3.Text = Text3.Text & " " & a(i)
        If i Mod 10 = 0 Then Text3.Text = Text3.Text & vbCrLf
    Next i
End Sub

Private Sub Command5_Click()
    Dim i%, j%, b%, c(51) As Integer, d%
    d = 1
    For i = 1 To Val(Label2.Caption)
        b = b + a(i)
    Next i
    b = Round(b / Val(Label2.Caption), 2)
    For i = 1 To Val(Label2.Caption)
        If a(i) < b Then c(d) = a(i): d = d + 1
    Next i
    For i = 1 To d - 1
        For j = 1 To d - i
            If c(j) < c(j + 1) Then c(0) = a(j): c(j) = c(j + 1): c(j + 1) = c(0)
        Next j
    Next i
    For i = 1 To d - 1
        Text4.Text = Text4.Text & " " & c(i)
        If i Mod 10 = 0 Then Text4.Text = Text4.Text & vbCrLf
    Next i
End Sub

Private Sub HScroll1_Change()
    Label2.Caption = HScroll1.Value
End Sub

Private Sub HScroll1_Scroll()
    Label2.Caption = HScroll1.Value
End Sub

Private Sub Text1_Change()
    Command3.Enabled = True
    Command4.Enabled = True
End Sub

Private Sub Text3_Change()
    Command5.Enabled = True
End Sub
