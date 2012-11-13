VERSION 5.00
Begin VB.Form ScoreCard 
   BackColor       =   &H000000FF&
   Caption         =   "Form1"
   ClientHeight    =   4200
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   ScaleHeight     =   4200
   ScaleWidth      =   6615
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PicScore 
      BackColor       =   &H00400040&
      FillColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0FF&
      Height          =   1215
      Left            =   720
      ScaleHeight     =   1155
      ScaleWidth      =   4875
      TabIndex        =   1
      Top             =   1560
      Width           =   4935
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Click to View Final Score"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   3135
   End
End
Attribute VB_Name = "ScoreCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
PicScore.Print "Your Final Score is " & Score
HighScore = InputBox("Input your name", "input")
FrmDog_Pound_Main.Show
ScoreCard.Hide
End Sub
