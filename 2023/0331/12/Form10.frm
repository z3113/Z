VERSION 5.00
Begin VB.Form Form10 
   Caption         =   "走台阶"
   ClientHeight    =   4830
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6270
   LinkTopic       =   "Form10"
   ScaleHeight     =   4830
   ScaleWidth      =   6270
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "答案"
      Height          =   615
      Left            =   3840
      TabIndex        =   1
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "11、20个台阶，可以一步一阶，也可一步两阶。问走完20个台阶共有几种走法？（先后次序不同当相同走法处理）"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3240
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Cls
    Dim i%, j%, a%
    For i = 0 To 20
        For j = 0 To 10
            If i + 2 * j = 20 Then
                a = a + 1
            End If
        Next j
    Next i
    Print "共有走法： " & a & " 种"
End Sub
