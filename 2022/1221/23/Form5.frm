VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "练习四、十二生肖"
   ClientHeight    =   5205
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6180
   LinkTopic       =   "Form5"
   ScaleHeight     =   5205
   ScaleWidth      =   6180
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      BackColor       =   &H00000000&
      Caption         =   "返回主窗体"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   5
      Top             =   3840
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "您的生肖"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2520
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3600
      TabIndex        =   3
      Top             =   2520
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3600
      TabIndex        =   2
      Text            =   "1996"
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "鼠牛虎兔龙蛇马羊猴鸡狗猪"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   1200
      TabIndex        =   6
      Top             =   4680
      Width           =   3855
   End
   Begin VB.Label Label2 
      Caption         =   "您出生年份"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "十二生肖"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Select Case Val(Text1.Text) Mod 12
        Case 0
            Text2.Text = "猴"
        Case 1
            Text2.Text = "鸡"
        Case 2
            Text2.Text = "狗"
        Case 3
            Text2.Text = "猪"
        Case 4
            Text2.Text = "鼠"
        Case 5
            Text2.Text = "牛"
        Case 6
            Text2.Text = "虎"
        Case 7
            Text2.Text = "兔"
        Case 8
            Text2.Text = "龙"
        Case 9
            Text2.Text = "蛇"
        Case 10
            Text2.Text = "马"
        Case 11
            Text2.Text = "羊"
    End Select
End Sub

Private Sub Command2_Click()
    Unload Form5
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub
