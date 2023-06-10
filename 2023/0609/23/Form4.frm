VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00FFC0C0&
   Caption         =   "列表项移动"
   ClientHeight    =   6180
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8310
   LinkTopic       =   "Form4"
   ScaleHeight     =   6180
   ScaleWidth      =   8310
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command5 
      Caption         =   "返回（&E)"
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
      Left            =   3360
      TabIndex        =   6
      Top             =   4800
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "<<"
      Height          =   495
      Left            =   3600
      TabIndex        =   5
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   ">>"
      Height          =   495
      Left            =   3600
      TabIndex        =   4
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<"
      Height          =   495
      Left            =   3600
      TabIndex        =   3
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">"
      Height          =   495
      Left            =   3600
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3765
      ItemData        =   "Form4.frx":0000
      Left            =   5160
      List            =   "Form4.frx":0002
      TabIndex        =   1
      Top             =   600
      Width           =   1935
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3765
      ItemData        =   "Form4.frx":0004
      Left            =   1200
      List            =   "Form4.frx":0006
      TabIndex        =   0
      Top             =   600
      Width           =   1935
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If List1.ListIndex = -1 Then
        MsgBox "请在列表框1选择需要移动的项", vbOKCancel + 48, "警告！！！"
    Else
        List2.AddItem List1.Text
        List1.RemoveItem List1.ListIndex
    End If
End Sub

Private Sub Command2_Click()
    If List2.ListIndex = -1 Then
        MsgBox "请在列表框2选择需要移动的项", vbOKCancel + 48, "警告！！！"
    Else
        List1.AddItem List2.Text
        List2.RemoveItem List2.ListIndex
    End If
End Sub

Private Sub Command3_Click()
    Dim i%
    For i = 0 To List1.ListCount - 1
        List2.AddItem List1.List(i)
    Next i
    List1.Clear
End Sub

Private Sub Command4_Click()
    Dim i%
    For i = 0 To List2.ListCount - 1
        List1.AddItem List2.List(i)
    Next i
    List2.Clear
End Sub

Private Sub Command5_Click()
    Unload Form4
End Sub

Private Sub Form_Load()
    Dim i%
    For i = 1 To 10
        List1.AddItem i
        List2.AddItem Chr(64 + i)
    Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub
