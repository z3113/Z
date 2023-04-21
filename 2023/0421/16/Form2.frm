VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "基本操作"
   ClientHeight    =   4470
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9150
   LinkTopic       =   "Form2"
   ScaleHeight     =   4470
   ScaleWidth      =   9150
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "返回"
      Height          =   375
      Left            =   7800
      TabIndex        =   17
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "重写"
      Height          =   375
      Left            =   6840
      TabIndex        =   16
      Top             =   3120
      Width           =   735
   End
   Begin VB.Frame Frame3 
      Caption         =   "颜色"
      Height          =   2295
      Left            =   6960
      TabIndex        =   10
      Top             =   480
      Width           =   1455
      Begin VB.Label Label6 
         BackColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label5 
         BackColor       =   &H0000FFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label4 
         BackColor       =   &H000000C0&
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FF00FF&
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0FF&
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "字型"
      Height          =   1575
      Left            =   5040
      TabIndex        =   6
      Top             =   2160
      Width           =   1455
      Begin VB.CheckBox Check3 
         Caption         =   "下划线"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   1080
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         Caption         =   "倾斜"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "加粗"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "字体"
      Height          =   1815
      Left            =   5040
      TabIndex        =   2
      Top             =   240
      Width           =   1455
      Begin VB.OptionButton Option3 
         Caption         =   "(&Y)幼圆"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1320
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         Caption         =   "(&K)楷体"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "(&H)黑体"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "Form2.frx":0000
      Top             =   600
      Width           =   3855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "请您留下宝贵意见："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2565
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
    Text1.FontBold = Not Text1.FontBold
End Sub

Private Sub Check2_Click()
    Text1.FontItalic = Not Text1.FontItalic
End Sub

Private Sub Check3_Click()
    Text1.FontUnderline = Not Text1.FontUnderline
End Sub

Private Sub Command1_Click()
    Text1.Text = ""
End Sub

Private Sub Command2_Click()
    Unload Form2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub

Private Sub Label2_Click()
    Text1.ForeColor = Label2.BackColor
End Sub

Private Sub Label3_Click()
    Text1.ForeColor = Label3.BackColor
End Sub

Private Sub Label4_Click()
    Text1.ForeColor = Label4.BackColor
End Sub

Private Sub Label5_Click()
    Text1.ForeColor = Label5.BackColor
End Sub

Private Sub Label6_Click()
    Text1.ForeColor = Label6.BackColor
End Sub

Private Sub Option1_Click()
    Text1.FontName = "黑体"
End Sub

Private Sub Option2_Click()
    Text1.FontName = "楷体"
End Sub

Private Sub Option3_Click()
    Text1.FontName = "幼圆"
End Sub
