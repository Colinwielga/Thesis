VERSION 5.00
Begin VB.Form frmWinner 
   Caption         =   "Winner!"
   ClientHeight    =   7560
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3990
   LinkTopic       =   "Form1"
   Picture         =   "frmWinner.frx":0000
   ScaleHeight     =   7560
   ScaleWidth      =   3990
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEnd 
      BackColor       =   &H000080FF&
      Caption         =   "End"
      Height          =   855
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5160
      Width           =   975
   End
   Begin VB.PictureBox picValue 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   0
      ScaleHeight     =   1515
      ScaleWidth      =   3915
      TabIndex        =   0
      Top             =   6000
      Width           =   3975
   End
   Begin VB.OLE OLE1 
      BackColor       =   &H00FF0000&
      Class           =   "Package"
      Height          =   615
      Left            =   0
      OleObjectBlob   =   "frmWinner.frx":C684
      SourceDoc       =   "M:\CS130\Project 1-Deal or No Deal\Program Sounds\dond-deal.mp3"
      TabIndex        =   3
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label lblCongrats 
      BackColor       =   &H0000FF00&
      Caption         =   "Congratulations!!!!!!!"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "frmWinner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Displays the contestants winnings, including their case value if applicable and
'also ends the program

'The End command button ends the program
Private Sub cmdEnd_Click()

'End program
End

End Sub
