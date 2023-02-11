VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "数据处理"
   ClientHeight    =   4695
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7125
   LinkTopic       =   "Form4"
   ScaleHeight     =   4695
   ScaleWidth      =   7125
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   4080
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   3480
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "返回"
      Height          =   615
      Left            =   5160
      TabIndex        =   4
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "最值"
      Height          =   615
      Left            =   5160
      TabIndex        =   3
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "平均值"
      Height          =   615
      Left            =   5160
      TabIndex        =   2
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "产生"
      Height          =   615
      Left            =   5160
      TabIndex        =   1
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   2415
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   240
      Width           =   4455
   End
   Begin VB.Label Label3 
      Caption         =   "最小值"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "最大值"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "平均值"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   3000
      Width           =   615
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim i As Long
    Dim sj(10 * 10) As Long
    Dim x
    Dim max
    Dim min
    Dim n

Private Sub Command1_Click()
    Randomize
    x = ""
    max = 0
    min = 100
    n = 0
    For i = 1 To 100
        sj(i) = Int(Rnd * 90 + 10)
        x = x & sj(i) & "  "
        n = n + sj(i)
        If max < sj(i) Then max = sj(i)
        If min > sj(i) Then min = sj(i)
        If i Mod 10 = 0 Then x = x & vbCrLf
    Next
    Text1.Text = x
End Sub

Private Sub Command2_Click()
    Text2.Text = n / 100
End Sub

Private Sub Command3_Click()
    Text3.Text = max
    Text4.Text = min
End Sub

Private Sub Command4_Click()
    Unload Form4
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub
