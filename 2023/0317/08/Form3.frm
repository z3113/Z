VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "�������"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7815
   LinkTopic       =   "Form3"
   ScaleHeight     =   5415
   ScaleWidth      =   7815
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command2 
      Caption         =   "����������"
      Height          =   615
      Left            =   5040
      TabIndex        =   5
      Top             =   4560
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "������������ж�����"
      Height          =   615
      Left            =   1200
      TabIndex        =   4
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Caption         =   "����"
      Height          =   4215
      Left            =   4080
      TabIndex        =   1
      Top             =   120
      Width           =   3495
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3855
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "40����λ�������"
      Height          =   4215
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3855
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   3255
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Randomize
    Label1.Caption = ""
    Label2.Caption = ""
    Dim i%, a%, c%
    For i = 1 To 40
        a = Rnd * 1001 + 2000
        Label1.Caption = Label1.Caption & Str(a)
        If i Mod 4 = 0 Then Label1.Caption = Label1.Caption & vbCrLf
        If (a Mod 4 = 0 And a Mod 100 <> 0) Or (a Mod 400 = 0) Then
            Label2.Caption = Label2.Caption & Str(a)
            c = c + 1
            If c Mod 4 = 0 Then Label2.Caption = Label2.Caption & vbCrLf
        End If
    Next i
    Label2.Caption = Label2.Caption & vbCrLf & "�����������" & c & "��"
End Sub

Private Sub Command2_Click()
    Unload Form3
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub
