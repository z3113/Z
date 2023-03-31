VERSION 5.00
Begin VB.Form Form7 
   Caption         =   "找特殊数"
   ClientHeight    =   11085
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6945
   LinkTopic       =   "Form7"
   ScaleHeight     =   11085
   ScaleWidth      =   6945
   StartUpPosition =   3  '窗口缺省
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "单击窗体打印答案"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   4800
      TabIndex        =   1
      Top             =   1800
      Width           =   1440
   End
   Begin VB.Label Label1 
      Caption         =   "找出所有能被6整除且至少有一位是7的四位自然数，按紧凑格式输出，每行7个。"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4320
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
    Cls
    Dim i%, a%
    For i = 1000 To 9999
        If i Mod 6 = 0 And (i Mod 10 = 7 Or i \ 10 Mod 10 = 7 Or i \ 100 Mod 10 = 7 Or i \ 1000 = 7) Then
            Print i;
            a = a + 1
            If a Mod 7 = 0 Then Print
        End If
    Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form4.Show
End Sub
