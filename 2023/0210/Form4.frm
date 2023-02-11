VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "文本框与素数"
   ClientHeight    =   3870
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7395
   LinkTopic       =   "Form4"
   ScaleHeight     =   3870
   ScaleWidth      =   7395
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "功能"
      Height          =   2535
      Left            =   5280
      TabIndex        =   2
      Top             =   480
      Width           =   1335
      Begin VB.CommandButton Command2 
         Caption         =   "返回"
         Height          =   855
         Left            =   240
         TabIndex        =   4
         Top             =   1440
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "产生"
         Height          =   855
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   960
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "10-200范围的素数:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim a%, b$, c%, i%
    For i = 10 To 200
        For a = 2 To i
            If i Mod a = 0 Then Exit For
        Next a
        If i = a Then
            c = c + 1
            b = b & " " & i
            If c Mod 4 = 0 Then b = b & vbCrLf
        End If
    Next i
    Text1.Text = b
End Sub

Private Sub Command2_Click()
    Unload Form4
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub
