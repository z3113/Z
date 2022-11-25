VERSION 5.00
Begin VB.Form Form2 
   ClientHeight    =   3450
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   ScaleHeight     =   3450
   ScaleWidth      =   4560
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      TabIndex        =   2
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   2280
      Width           =   1215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a%
Option Explicit

Private Sub Command1_Click()
    a = InputBox("", "", 0)
End Sub

Private Sub Command2_Click()
    Text1.Text = a
    MsgBox "" & Text1.Text & Chr(10) & Chr(13) & "", 0 + 64, ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Form2
    Form1.Show
End Sub
