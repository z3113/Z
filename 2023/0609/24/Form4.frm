VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00FFC0C0&
   Caption         =   "基本应用"
   ClientHeight    =   6660
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7605
   LinkTopic       =   "Form4"
   ScaleHeight     =   6660
   ScaleWidth      =   7605
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command4 
      Caption         =   "返回"
      Height          =   540
      Left            =   6000
      TabIndex        =   9
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "排序（冒泡）"
      Height          =   540
      Left            =   4080
      TabIndex        =   8
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "找素数"
      Height          =   540
      Left            =   2160
      TabIndex        =   7
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "产生"
      Height          =   540
      Left            =   240
      TabIndex        =   6
      Top             =   5280
      Width           =   1335
   End
   Begin VB.ListBox List2 
      Height          =   3480
      Left            =   5400
      TabIndex        =   5
      Top             =   840
      Width           =   1815
   End
   Begin VB.ListBox List1 
      Height          =   3480
      Left            =   2880
      TabIndex        =   4
      Top             =   840
      Width           =   1815
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "Form4.frx":0000
      Left            =   360
      List            =   "Form4.frx":0002
      TabIndex        =   1
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "从小到大排序"
      Height          =   180
      Left            =   5400
      TabIndex        =   3
      Top             =   480
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "素数："
      Height          =   180
      Left            =   2880
      TabIndex        =   2
      Top             =   480
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "产生100个1-999随机数"
      Height          =   180
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   1800
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Randomize
    Command1.Enabled = False
    Combo1.Clear
    Dim i%, max%, a%
    max = 0
    For i = 0 To 99
        Combo1.AddItem Int(Rnd * 999 + 1)
        If Val(Combo1.List(max)) < Val(Combo1.List(i)) Then max = i
    Next i
    Combo1.ListIndex = max
End Sub

Private Sub Command2_Click()
    Dim i%, j%, a%
    List1.Clear
    For i = 0 To Combo1.ListCount - 1
        a = Val(Combo1.List(i))
        For j = 2 To a
            If a Mod j = 0 Then Exit For
        Next j
        If a = j Then List1.AddItem a
    Next i
End Sub

Private Sub Command3_Click()
    Dim i%, j%, a%, b%, c%
    List2.Clear
    For i = 0 To List1.ListCount - 1
        List2.AddItem List1.List(i)
    Next i
    For i = 1 To List2.ListCount - 1
        For j = 0 To List2.ListCount - 1 - i
            'a = List2.List(j)
            'b = List2.List(j + 1)
            'If a > b Then c = a: a = b: b = c
            'List2.List(j) = a
            'List2.List(j + 1) = b
            If Val(List2.List(j)) > Val(List2.List(j + 1)) Then c = List2.List(j): List2.List(j) = List2.List(j + 1): List2.List(j + 1) = c
        Next j
    Next i
End Sub

Private Sub Command4_Click()
    Unload Form4
End Sub

