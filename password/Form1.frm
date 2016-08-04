VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4965
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   ScaleHeight     =   4965
   ScaleWidth      =   7590
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "confirm"
      Height          =   1215
      Left            =   1440
      TabIndex        =   2
      Top             =   1920
      Width           =   4215
   End
   Begin VB.TextBox txt1 
      Height          =   495
      Left            =   3000
      TabIndex        =   0
      Top             =   600
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "password enter"
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   600
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If txt1.Text = "xeeshan" Then
MsgBox "sucsessful!!!!", vbinformition, ""

Else
MsgBox "passwaord is not found", vbinformition, ""


End If




End Sub

Private Sub txt1_Change()

End Sub
