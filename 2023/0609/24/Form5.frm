VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00FFC0C0&
   Caption         =   "综合应用"
   ClientHeight    =   5280
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7620
   LinkTopic       =   "Form5"
   ScaleHeight     =   5280
   ScaleWidth      =   7620
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command6 
      Caption         =   "返回（&E）"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      TabIndex        =   7
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "选择排序"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      TabIndex        =   6
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "找奇数"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   5
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "找最大值"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      TabIndex        =   4
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "求平均值"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      TabIndex        =   3
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "产生数据"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   2
      Top             =   3240
      Width           =   1575
   End
   Begin VB.ComboBox Combo2 
      Height          =   2520
      Left            =   4320
      Style           =   1  'Simple Combo
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   240
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      Height          =   2520
      Left            =   1680
      Style           =   1  'Simple Combo
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Randomize
    Dim i%
    For i = 0 To 99
        Combo1.AddItem Int(Rnd * 1000 + 1)
    Next i
End Sub

Private Sub Command2_Click()
    Dim i%, ave#
    Cls
    For i = 0 To Combo1.ListCount - 1
        ave = ave + Combo1.List(i)
    Next i
    Print "平均值为：" & Round(ave / 100, 2)
End Sub

Private Sub Command3_Click()
    Dim i%, max%
    Cls
    max = Combo1.List(0)
    For i = 1 To Combo1.ListCount - 1
        If max < Combo1.List(i) Then max = Combo1.List(i)
    Next i
    Print "最大值为：" & max
End Sub

Private Sub Command4_Click()
    Dim i%
    Combo2.Clear
    For i = 0 To Combo1.ListCount - 1
        If Combo1.List(i) Mod 2 <> 0 Then Combo2.AddItem Combo1.List(i)
    Next i
End Sub

Private Sub Command5_Click()
    Dim i%, j%, a%, b%
    For i = 0 To Combo2.ListCount - 2
        a = i
        For j = i + 1 To Combo2.ListCount - 1
            If Val(Combo2.List(a)) > Val(Combo2.List(j)) Then a = j
        Next j
        If a <> i Then b = Combo2.List(i): Combo2.List(i) = Combo2.List(a): Combo2.List(a) = b
    Next i
End Sub

Private Sub Command6_Click()
    Unload Form5
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub
