VERSION 5.00
Begin VB.Form frmGame3 
   BackColor       =   &H00000000&
   Caption         =   "Search and Matching Game"
   ClientHeight    =   8550
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18825
   LinkTopic       =   "Form1"
   Picture         =   "frmGame3.frx":0000
   ScaleHeight     =   8550
   ScaleWidth      =   18825
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   735
      Left            =   15000
      TabIndex        =   5
      Top             =   7320
      Width           =   3255
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back To Index"
      Height          =   735
      Left            =   16680
      TabIndex        =   4
      Top             =   6360
      Width           =   1575
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "Check Answer"
      Height          =   735
      Left            =   15000
      TabIndex        =   3
      Top             =   6360
      Width           =   1575
   End
   Begin VB.PictureBox picResults 
      Height          =   4095
      Left            =   15000
      ScaleHeight     =   4035
      ScaleWidth      =   3195
      TabIndex        =   2
      Top             =   2160
      Width           =   3255
   End
   Begin VB.TextBox txtObjects 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   15000
      TabIndex        =   1
      Top             =   1200
      Width           =   3255
   End
   Begin VB.Label lblWhatDoYouSee 
      BackColor       =   &H80000012&
      Caption         =   "What do you see? (Use singular)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   855
      Left            =   15120
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
End
Attribute VB_Name = "frmGame3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBack_Click()

frmGameScene.Show
frmGame3.Hide

End Sub

Private Sub cmdCheck_Click()

Dim pos As Integer, ctr As Integer
Dim Objects(1 To 100) As String
Dim found As Boolean
Dim inputName As String

Open App.Path & "\objects.txt" For Input As #1
Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, Objects(ctr)
Loop
Close #1
    
inputName = txtObjects.Text
    
picResults.Cls
found = False

Do While found = False And pos < ctr
    pos = pos + 1
    If LCase(Objects(pos)) = LCase(inputName) Then
        found = True
        picResults.Print Objects(pos) & " founded!"
    End If
Loop

If found = False Then
    picResults.Print inputName & " was not found!"
End If



End Sub

Private Sub cmdQuit_Click()

End

End Sub
