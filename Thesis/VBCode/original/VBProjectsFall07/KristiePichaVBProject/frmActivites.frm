VERSION 5.00
Begin VB.Form frmActivities 
   BackColor       =   &H0000FFFF&
   Caption         =   "Form1"
   ClientHeight    =   4545
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   ScaleHeight     =   4545
   ScaleWidth      =   6135
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Back to Menu"
      Height          =   615
      Left            =   3000
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.PictureBox picResults 
      Height          =   3495
      Left            =   720
      ScaleHeight     =   3435
      ScaleWidth      =   4035
      TabIndex        =   1
      Top             =   840
      Width           =   4095
   End
   Begin VB.CommandButton cmdDisplay 
      Caption         =   "Display"
      Height          =   615
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmActivities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Activities(1 To 100) As String
Dim CTR As Integer


Private Sub cmdDisplay_Click()
CTR = 0
Open App.Path & "\Activities.txt" For Input As #1
Do Until EOF(1)
    CTR = CTR + 1
    Input #1, Activities(CTR)
    picResults.Print Activities(CTR)
Loop
Close #1

End Sub

Private Sub cmdMenu_Click()
frmActivities.Hide
frmMenu.Show
End Sub


