VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "图形打印"
   ClientHeight    =   5565
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8910
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   14.25
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form5"
   ScaleHeight     =   5565
   ScaleWidth      =   8910
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Caption         =   "沙漏"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7080
      TabIndex        =   2
      Top             =   3840
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "菱形"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7080
      TabIndex        =   1
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "三角形"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7080
      TabIndex        =   0
      Top             =   720
      Width           =   1455
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Cls
    Print "12345678901234567890123456789012345678901234567890123456789012345678901234567890"
    Dim i%, j%, a%
    a = Val(InputBox("请输入行数", "图形打印", 10))
    For i = 1 To a
        Print Tab(a + 1 - i);
        For j = 1 To 2 * i - 1
            Print Chr(64 + i);
        Next j
    Next i
End Sub

Private Sub Command2_Click()
    Cls
    Print "12345678901234567890123456789012345678901234567890123456789012345678901234567890"
    Dim i%, j%, a%
    a = Val(InputBox("请输入行数(奇数)", "图形打印", 11))
    If a Mod 2 = 0 Then
        MsgBox "请输入奇数行"
    Else
        a = (a - 1) / 2
        For i = -a To a
            Print Tab(Abs(i) + 1);
            For j = 1 To 11 - Abs(2 * i)
                If j Mod 2 = 0 Then Print "*"; Else Print "$";
            Next j
        Next i
    End If
End Sub

Private Sub Command3_Click()
    Cls
    Print "12345678901234567890123456789012345678901234567890123456789012345678901234567890"
        Dim i%, j%, a%
    a = Val(InputBox("请输入行数(奇数)", "图形打印", 11))
    a = (a - 1) / 2
    For i = -a To a
        Print Tab(a - Abs(i) + 1);
        For j = 1 To Abs(2 * i) + 1
            Print Chr(65 + Abs(i));
        Next j
    Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub
