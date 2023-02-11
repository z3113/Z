VERSION 5.00
Begin VB.Form Form2 
   ClientHeight    =   4575
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8655
   LinkTopic       =   "Form2"
   ScaleHeight     =   4575
   ScaleWidth      =   8655
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command4 
      Height          =   375
      Left            =   5640
      TabIndex        =   3
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   3840
      Width           =   1335
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Print "123456789012345678901234567890"
    Print "he", "is", "a", "good", "student"
    Print "1", "2", "3", "4", "5"
    Print 1, 2, 3, 4, 5
End Sub

Private Sub Command2_Click()
    Print "123456789012345678901234567890"
    Print "he"; "is"; "a"; "good"; "student"
    Print "1"; "2"; "3"; "4"; "5"
    Print 1; 2; 3; 4; 5
End Sub

Private Sub Command3_Click()
    Print "123456789012345678901234567890123456789012345678901234567890"
    Print Tab(5); "He", "is", "good!"
    Print Tab; "He"; Tab(12), "is", "good!"
    Print Space(5); "He"; Space(1); "is"; Space(1); "good!"
    Print Space(5); 1; Space(1); 2; Space(1); 3
End Sub

Private Sub Command4_Click()
    Cls
    Print "123456789012345678901234567890123456789012345678901234567890"
    Print Tab(20); "*"
    Print Tab(19); "***"
    Print Tab(18); "*****"
    Print Tab(17); "*******"
    Print Tab(16); "*********"
    Print Tab(15); "***********"
    Print Tab(19); "**"
    Print Tab(19); "**"
    Print Tab(19); "**"
    Print Tab(19); "**"
    Print Tab(19); "**"
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Form2
    Form1.Show
End Sub
