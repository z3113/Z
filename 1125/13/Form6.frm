VERSION 5.00
Begin VB.Form Form6 
   ClientHeight    =   4410
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6135
   LinkTopic       =   "Form6"
   ScaleHeight     =   4410
   ScaleWidth      =   6135
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   120
   End
   Begin VB.CommandButton Command2 
      Height          =   495
      Left            =   3360
      TabIndex        =   1
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      Height          =   1095
      Left            =   1080
      Shape           =   3  'Circle
      Top             =   1320
      Width           =   1095
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a%
Option Explicit

Private Sub Command1_Click()
    Timer1.Enabled = True
End Sub

Private Sub Command2_Click()
    Timer1.Enabled = False
End Sub

Private Sub Form_Load()
    a = 200
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Form6
    Form1.Show
End Sub

Private Sub Timer1_Timer()
    Shape1.Left = Shape1.Left + a
    If Shape1.Left <= 0 Or Shape1.Left >= Form6.Width - Shape1.Width Then
        a = -a
    End If
End Sub
