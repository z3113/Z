VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "计算器"
   ClientHeight    =   3735
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4095
   LinkTopic       =   "Form6"
   ScaleHeight     =   3735
   ScaleWidth      =   4095
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command4 
      Caption         =   "="
      Height          =   495
      Left            =   2160
      TabIndex        =   16
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "."
      Height          =   495
      Left            =   1200
      TabIndex        =   15
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "/"
      Height          =   495
      Index           =   3
      Left            =   3120
      TabIndex        =   14
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "*"
      Height          =   495
      Index           =   2
      Left            =   3120
      TabIndex        =   13
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "-"
      Height          =   495
      Index           =   1
      Left            =   3120
      TabIndex        =   12
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "+"
      Height          =   495
      Index           =   0
      Left            =   3120
      TabIndex        =   11
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "9"
      Height          =   495
      Index           =   9
      Left            =   2160
      TabIndex        =   10
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "8"
      Height          =   495
      Index           =   8
      Left            =   1200
      TabIndex        =   9
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "7"
      Height          =   495
      Index           =   7
      Left            =   240
      TabIndex        =   8
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "6"
      Height          =   495
      Index           =   6
      Left            =   2160
      TabIndex        =   7
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "5"
      Height          =   495
      Index           =   5
      Left            =   1200
      TabIndex        =   6
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "4"
      Height          =   495
      Index           =   4
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "3"
      Height          =   495
      Index           =   3
      Left            =   2160
      TabIndex        =   4
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "2"
      Height          =   495
      Index           =   2
      Left            =   1200
      TabIndex        =   3
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   495
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   2280
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "0"
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   3000
      Width           =   735
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a#, b#, c#, d%

Private Sub Command1_Click(Index As Integer)
    Text1.Text = Text1.Text & Index
End Sub

Private Sub Command2_Click(Index As Integer)
    Select Case Index
        Case 0
            a = 0
            b = Val(Text1.Text)
            d = 0
        Case 1
            a = 1
            b = Val(Text1.Text)
            d = 0
        Case 2
            a = 2
            b = Val(Text1.Text)
            d = 0
        Case 3
            a = 3
            b = Val(Text1.Text)
            d = 0
    End Select
    Text1.Text = ""
End Sub

Private Sub Command3_Click()
    If d = 0 Then
        Text1.Text = Text1.Text & "."
    End If
    d = d + 1
End Sub

Private Sub Command4_Click()
    c = Val(Text1.Text)
    Select Case a
        Case 0
            Text1.Text = b + c
        Case 1
            Text1.Text = b - c
        Case 2
            Text1.Text = b * c
        Case 3
            If c <> 0 Then Text1.Text = b / c Else MsgBox "除数为0，无法计算！", vbOKOnly + 16, "错误"
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub
