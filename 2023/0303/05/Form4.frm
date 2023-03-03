VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "任务三、查找数"
   ClientHeight    =   4470
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7455
   LinkTopic       =   "Form4"
   ScaleHeight     =   4470
   ScaleWidth      =   7455
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "返回"
      Height          =   495
      Left            =   3960
      TabIndex        =   6
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "查找"
      Height          =   495
      Left            =   2280
      TabIndex        =   5
      Top             =   3600
      Width           =   1215
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   375
      Left            =   3720
      Max             =   999
      Min             =   300
      TabIndex        =   3
      Top             =   1680
      Value           =   300
      Width           =   3255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      Left            =   3720
      Max             =   999
      Min             =   300
      TabIndex        =   1
      Top             =   720
      Value           =   300
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   2775
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   3375
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   5400
      TabIndex        =   8
      Top             =   2160
      Width           =   120
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   5400
      TabIndex        =   7
      Top             =   1200
      Width           =   120
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "最大数："
      Height          =   180
      Left            =   4080
      TabIndex        =   4
      Top             =   2160
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "最小数："
      Height          =   180
      Left            =   4080
      TabIndex        =   2
      Top             =   1200
      Width           =   720
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim i%, a%, b%, c%
    a = Val(HScroll1.Value)
    b = Val(HScroll2.Value)
    If a > b Then
        Text1.Text = "范围不合理"
    Else
        For i = a To b
            If i \ 100 + i Mod 10 = i \ 10 Mod 10 Then
                Text1.Text = Text1.Text & " " & i
                c = c + 1
                If c Mod 5 = 0 Then Text1.Text = Text1.Text & vbCrLf
            End If
        Next i
    End If
End Sub

Private Sub Command2_Click()
    Unload Form4
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form2.Show
End Sub

Private Sub HScroll1_Change()
    Label3.Caption = HScroll1.Value
End Sub

Private Sub HScroll1_Scroll()
    Label3.Caption = HScroll1.Value
End Sub

Private Sub HScroll2_Change()
    Label4.Caption = HScroll2.Value
End Sub

Private Sub HScroll2_Scroll()
    Label4.Caption = HScroll2.Value
End Sub
