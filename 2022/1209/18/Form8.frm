VERSION 5.00
Begin VB.Form Form8 
   Caption         =   "循环滚动"
   ClientHeight    =   4500
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7350
   LinkTopic       =   "Form8"
   ScaleHeight     =   4500
   ScaleWidth      =   7350
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   5400
      Top             =   960
   End
   Begin VB.CommandButton Command3 
      Caption         =   "BACK"
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
      Left            =   5160
      TabIndex        =   4
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "STOP"
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
      Left            =   2520
      TabIndex        =   3
      Top             =   3240
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "START"
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
      Left            =   360
      TabIndex        =   2
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "（从左至右循环滚动）"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   600
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "（从上至下循环滚动）"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   1920
      Width           =   3375
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Timer1.Enabled = True
End Sub

Private Sub Command2_Click()
    Timer1.Enabled = False
End Sub

Private Sub Command3_Click()
    Unload Form8
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form2.Show
End Sub

Private Sub Timer1_Timer()
    Label1.Top = Label1.Top - 100
    Label2.Left = Label2.Left + 100
    If Label1.Top <= 0 Then Label1.Top = Form8.ScaleHeight + Label1.Height
    If Label2.Left >= Form8.ScaleWidth Then Label2.Left = -Label2.Width
End Sub
