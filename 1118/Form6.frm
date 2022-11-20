VERSION 5.00
Begin VB.Form Form6 
   ClientHeight    =   3720
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5070
   LinkTopic       =   "Form6"
   ScaleHeight     =   3720
   ScaleWidth      =   5070
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command3 
      Height          =   495
      Left            =   3480
      TabIndex        =   2
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Height          =   495
      Left            =   3480
      TabIndex        =   1
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Height          =   495
      Left            =   3480
      TabIndex        =   0
      Top             =   840
      Width           =   1335
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Cls
    Print Tab(20); "*"
    Print Tab(20); "**"
    Print Tab(20); "***"
    Print Tab(20); "****"
    Print Tab(20); "*****"
End Sub

Private Sub Command2_Click()
    Cls
    Print Tab(20); "*"
    Print Tab(19); "**"
    Print Tab(18); "***"
    Print Tab(17); "****"
    Print Tab(16); "*****"
End Sub

Private Sub Command3_Click()
    Cls
    Print Tab(20); "********"
    Print Tab(19); "********"
    Print Tab(18); "********"
    Print Tab(17); "********"
    Print Tab(16); "********"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Form6
    Form1.Show
End Sub
