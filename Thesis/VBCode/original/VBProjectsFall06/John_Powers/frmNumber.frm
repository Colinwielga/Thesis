VERSION 5.00
Begin VB.Form frmNumber 
   Caption         =   "Numbers"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFinal 
      Caption         =   "GO!"
      Height          =   255
      Left            =   3840
      TabIndex        =   2
      Top             =   2640
      Width           =   735
   End
   Begin VB.PictureBox picFinal 
      Height          =   1095
      Left            =   600
      ScaleHeight     =   1035
      ScaleWidth      =   3555
      TabIndex        =   1
      Top             =   360
      Width           =   3615
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Compute!"
      Height          =   975
      Left            =   1080
      TabIndex        =   0
      Top             =   1680
      Width           =   2655
   End
End
Attribute VB_Name = "frmNumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdFinal_Click()
    frmNumber.Hide
    frmFinal.Show
End Sub

Private Sub cmdStart_Click()
    Dim a As Integer
    a = InputBox("Enter a number", "Number")
    If a > 100 Then
            MsgBox "Try a smaller number", , "Error !"
        ElseIf a > 50 Then
            a = a / 2
            picFinal.Print a
        ElseIf a > 0 Then
            a = a * 2
            picFinal.Print a
        Else
            MsgBox "Try a larger number", , "Error!"
    End If
End Sub
