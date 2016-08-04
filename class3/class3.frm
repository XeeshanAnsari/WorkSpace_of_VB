VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "addition"
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   2400
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   600
      TabIndex        =   1
      Top             =   840
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF00&
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   1680
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim a As Integer
Dim b As Integer
Dim c As Integer
a = Text1.Text
b = Text2.Text
c = a + b

Label1.Caption = c

End Sub
