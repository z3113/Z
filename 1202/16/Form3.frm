VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "任务二：去两端空格"
   ClientHeight    =   5085
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7590
   LinkTopic       =   "Form3"
   ScaleHeight     =   5085
   ScaleWidth      =   7590
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command4 
      Cancel          =   -1  'True
      Caption         =   "返回（ESC）"
      Height          =   495
      Left            =   5040
      TabIndex        =   0
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "去掉空格"
      Height          =   2775
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   6255
      Begin VB.CommandButton Command3 
         Caption         =   "去掉两边空格"
         Height          =   495
         Left            =   3960
         TabIndex        =   6
         Top             =   2040
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "去掉右边空格"
         Height          =   495
         Left            =   2040
         TabIndex        =   5
         Top             =   2040
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "去掉左边空格"
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   0
         TabIndex        =   3
         Top             =   1200
         Width           =   5775
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   0
         TabIndex        =   2
         Text            =   "   123456  7890    "
         Top             =   360
         Width           =   5775
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Text2.Text = LTrim(Text1.Text)
    MsgBox "原字符串是：" & Text1.Text & "！" & vbCrLf & "新字符串是：" & Text2.Text & "！"
End Sub

Private Sub Command2_Click()
    Text2.Text = RTrim(Text1.Text)
    MsgBox "原字符串是：" & Text1.Text & "！" & vbCrLf & "新字符串是：" & Text2.Text & "！"
End Sub

Private Sub Command3_Click()
    Text2.Text = Trim(Text1.Text)
    MsgBox "原字符串是：" & Text1.Text & "！" & vbCrLf & "新字符串是：" & Text2.Text & "！"
End Sub

Private Sub Command4_Click()
    Unload Form3
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub
