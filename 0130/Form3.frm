VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Timer And Scroll"
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7800
   LinkTopic       =   "Form3"
   ScaleHeight     =   4455
   ScaleWidth      =   7800
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   35
      Left            =   6120
      Top             =   1200
   End
   Begin VB.CommandButton Command3 
      Caption         =   "unload"
      Height          =   495
      Left            =   5280
      TabIndex        =   3
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "stop"
      Height          =   495
      Left            =   2880
      TabIndex        =   2
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "move"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   3480
      Width           =   1455
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      LargeChange     =   3
      Left            =   1200
      Max             =   155
      Min             =   35
      SmallChange     =   3
      TabIndex        =   0
      Top             =   2280
      Value           =   35
      Width           =   4935
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   6120
      Top             =   360
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   1320
      TabIndex        =   4
      Top             =   720
      Width           =   3255
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Timer2.Enabled = True
End Sub

Private Sub Command2_Click()
    Timer2.Enabled = False
End Sub

Private Sub Command3_Click()
    Unload Form3
End Sub

Private Sub Form_Activate()
    Label1.Caption = Hour(Now) & ":" & Minute(Now) & ":" & Second(Now)
End Sub

Private Sub HScroll1_Change()
    Timer2.Interval = HScroll1.Value
End Sub

Private Sub Timer1_Timer()
    Label1.Caption = Hour(Now) & ":" & Minute(Now) & ":" & Second(Now)
End Sub

Private Sub Timer2_Timer()
    Label1.Left = Label1.Left + 100
    If Label1.Left >= Form3.Width Then Label1.Left = -Label1.Width
End Sub
