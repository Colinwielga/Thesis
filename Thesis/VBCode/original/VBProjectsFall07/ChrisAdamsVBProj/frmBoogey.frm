VERSION 5.00
Begin VB.Form frmBoogey 
   Caption         =   "Uh Oh"
   ClientHeight    =   8145
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8160
   LinkTopic       =   "Form1"
   Picture         =   "frmBoogey.frx":0000
   ScaleHeight     =   8145
   ScaleWidth      =   8160
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdWrong 
      Caption         =   "Game Over"
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   7080
      Width           =   1695
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Enabled         =   0   'False
      Height          =   615
      Left            =   6120
      TabIndex        =   0
      Top             =   7080
      Width           =   1695
   End
   Begin VB.Label lblCLick 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Double Click the Wild Icon above to recieve a special message from one of your Minnesota Wild Teammates"
      Height          =   495
      Left            =   1200
      TabIndex        =   3
      Top             =   4200
      Width           =   6015
   End
   Begin VB.OLE oleBoogey 
      BackColor       =   &H00008000&
      Class           =   "Package"
      DisplayType     =   1  'Icon
      Height          =   1215
      Left            =   3360
      OleObjectBlob   =   "frmBoogey.frx":A833
      SourceDoc       =   "M:\CS130\ChrisAdamsVBProj\video\BoogeymanCondensed.wmv"
      TabIndex        =   2
      Top             =   1920
      Width           =   1455
   End
End
Attribute VB_Name = "frmBoogey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Quest for The Cup~Minnesota Wild Trivia Game

'Author: Chris Adams

'Date: November 2007

'This form is shown if the user gets a question wrong to inform them that the game is over.

Private Sub cmdQuit_Click()

'Show form Sources
frmBoogey.Hide
frmSources.Show

End Sub

Private Sub cmdWrong_Click()

cmdQuit.Enabled = True

End Sub
