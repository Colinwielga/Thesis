VERSION 5.00
Begin VB.Form frmPos 
   Caption         =   "Select a Position"
   ClientHeight    =   5730
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   Picture         =   "frmPos.frx":0000
   ScaleHeight     =   5730
   ScaleWidth      =   5520
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   615
      Left            =   4080
      TabIndex        =   4
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton cmdGoalie 
      Caption         =   "Goalie (Easy)"
      Height          =   615
      Left            =   2040
      TabIndex        =   3
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdDefense 
      Caption         =   "Defense (Medium)"
      Height          =   615
      Left            =   2040
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdForward 
      Caption         =   "Forward (Hard) "
      Height          =   615
      Left            =   2040
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lblSelectPos 
      Alignment       =   2  'Center
      Caption         =   "Please select your position"
      Height          =   255
      Left            =   1680
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "frmPos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Quest for The Cup~Minnesota Wild Trivia Game

'Author: Chris Adams

'Date: November 2007

'This form is where the user selects their position which corresponds to the level of difficulty of the game.

Private Sub cmdForward_Click()
    
    Pos = "Forward"     'Sets the user's position to Forward
    Game = 1            'This sets the difficulty of the game at Hard
                        'and makes the user eligble for the Hall of Fame
    frmNumber.Show
    frmPos.Hide

End Sub

Private Sub cmdDefense_Click()
    
    Pos = "Defense"     'Sets the user's position to Defense
    Game = 2            'This sets the difficulty of the game at Medium
    frmNumber.Show
    frmPos.Hide
    
End Sub
Private Sub cmdGoalie_Click()
    
    Pos = "Goalie"      'Sets the user's position to Goalie
    Game = 3            'This sets the difficulty of the game at Easy
    frmNumber.Show
    frmPos.Hide
    
    End Sub

Private Sub cmdQuit_Click()
    
    'Show form Sources
    frmPos.Hide
    frmSources.Show

End Sub
