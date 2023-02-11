VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "任务一：油价计算"
   ClientHeight    =   4020
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7245
   LinkTopic       =   "Form2"
   ScaleHeight     =   4020
   ScaleWidth      =   7245
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "返回"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      TabIndex        =   9
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "计算"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      TabIndex        =   8
      Top             =   3120
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1800
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      TabIndex        =   0
      Top             =   840
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "种类"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   2415
      Begin VB.OptionButton Option3 
         Caption         =   "100号汽油"
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
         Left            =   240
         TabIndex        =   4
         Top             =   1560
         Width           =   1815
      End
      Begin VB.OptionButton Option2 
         Caption         =   "95号汽油"
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
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Caption         =   "90号汽油"
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
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   0
      TabIndex        =   10
      Top             =   120
      Width           =   7215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "总价（元）："
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
      Left            =   3120
      TabIndex        =   6
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "数量（升）："
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
      Left            =   3120
      TabIndex        =   5
      Top             =   960
      Width           =   1815
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim a!
    a = Val(Text1.Text)
    If Option1.Value = True Then
        Text2.Text = a * 2.3
        Label3.Caption = "90号汽油：2.3元/升"
    ElseIf Option2.Value = True Then
        Text2.Text = a * 2.45
        Label3.Caption = "95号汽油：2.45元/升"
    ElseIf Option3.Value = True Then
        Text2.Text = a * 2.6
        Label3.Caption = "100号汽油：2.6元/升"
    End If
End Sub

Private Sub Command2_Click()
    Unload Form2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub

Private Sub Option1_Click()
    If Text1.Text <> "" And (Option1.Value = True Or Option2.Value = True Or Option3.Value = True) Then
        Command1.Enabled = True
    End If
End Sub

Private Sub Option2_Click()
    If Text1.Text <> "" And (Option1.Value = True Or Option2.Value = True Or Option3.Value = True) Then
        Command1.Enabled = True
    End If
End Sub

Private Sub Option3_Click()
    If Text1.Text <> "" And (Option1.Value = True Or Option2.Value = True Or Option3.Value = True) Then
        Command1.Enabled = True
    End If
End Sub

Private Sub Text1_Change()
    If Text1.Text <> "" And (Option1.Value = True Or Option2.Value = True Or Option3.Value = True) Then
        Command1.Enabled = True
    End If
End Sub
