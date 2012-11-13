VERSION 5.00
Begin VB.Form frmhowtoplayfootball 
   BackColor       =   &H00400000&
   Caption         =   "How to Play Football"
   ClientHeight    =   7155
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9945
   LinkTopic       =   "Form1"
   ScaleHeight     =   7155
   ScaleWidth      =   9945
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdplays 
      Caption         =   "How to Pick Plays"
      Height          =   1335
      Left            =   120
      TabIndex        =   12
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton cmdnext6 
      BackColor       =   &H0000FFFF&
      Caption         =   "next"
      Enabled         =   0   'False
      Height          =   615
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5760
      Width           =   975
   End
   Begin VB.CommandButton cmdnext5 
      BackColor       =   &H0000FFFF&
      Caption         =   "next"
      Enabled         =   0   'False
      Height          =   615
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5040
      Width           =   975
   End
   Begin VB.CommandButton cmdnext4 
      BackColor       =   &H00C00000&
      Caption         =   "next"
      Enabled         =   0   'False
      Height          =   615
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5760
      Width           =   855
   End
   Begin VB.CommandButton cmdnext3 
      BackColor       =   &H000000FF&
      Caption         =   "Next"
      Enabled         =   0   'False
      Height          =   615
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5760
      Width           =   855
   End
   Begin VB.CommandButton cmdnext2 
      BackColor       =   &H00C00000&
      Caption         =   "next"
      Enabled         =   0   'False
      Height          =   615
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton cmdnext1 
      BackColor       =   &H000000FF&
      Caption         =   "Next"
      Enabled         =   0   'False
      Height          =   615
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5040
      Width           =   855
   End
   Begin VB.PictureBox picresults2 
      Height          =   3375
      Left            =   120
      ScaleHeight     =   3315
      ScaleWidth      =   3795
      TabIndex        =   5
      Top             =   120
      Width           =   3855
   End
   Begin VB.PictureBox picresults 
      Height          =   6855
      Left            =   4080
      ScaleHeight     =   6795
      ScaleWidth      =   5595
      TabIndex        =   4
      Top             =   120
      Width           =   5655
   End
   Begin VB.CommandButton cmdback 
      BackColor       =   &H000000FF&
      Caption         =   "Go Back"
      Height          =   495
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6480
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FFFF&
      Caption         =   "Special Teams"
      Height          =   1335
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C00000&
      Caption         =   "Defense"
      Height          =   1335
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton cmdoffense 
      BackColor       =   &H000000FF&
      Caption         =   "Offense"
      Height          =   1335
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3600
      Width           =   855
   End
End
Attribute VB_Name = "frmhowtoplayfootball"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: Super Tecmo Database
'Form name: frmhowtoplayfootball
'Author: Nate Johnson & Kevin Klein
'Date Written: October 11th, 2006
'Objective of project: This project will allow its users to learn more about the game of football
'and will also allow them the oppurtunity to learn how to play the game of football with the Nintendo
'video game, Tecmo Super Bowl.
'Objective of form: this form gives the user an oppurtunity to learn how to play the game
'of football. Nearly every aspect of the game is covered, from offense to defense, and
'even special teams play. If the user has never recieved any education about the game,
'this is a great place to learn.


Private Sub CmdBack_Click()
frmhowtoplayfootball.Hide
frmtutorial.Show
End Sub

Private Sub cmdnext1_Click()
picresults.Cls 'clears the picbox of any data'
picresults2.Picture = LoadPicture(App.Path & "\pics\more pics\run.bmp") 'displays new picture
picresults.Print "Running the ball relies on the blocking of the offensive" 'displays the written data
picresults.Print "linemen as they try to make holes in the defensive line."
picresults.Print "A running back will then try to run the ball through these"
picresults.Print "holes and up the field. Running is usually called when the"
picresults.Print "offense has small amounts of the field to cover or when"
picresults.Print "there is still lots of time on remaining in the game."
cmdnext3.Enabled = True
cmdnext1.Enabled = False

End Sub

Private Sub cmdnext2_Click()
picresults.Cls 'clears the picbox of any data'
picresults2.Picture = LoadPicture(App.Path & "\pics\more pics\sack.bmp") 'displays new picture
picresults.Print "The defense can try to get past the line of scrimmage quickly" 'displays the written data'
picresults.Print "and tackle the quarterback before he throws the ball, this is"
picresults.Print "called a sack."
cmdnext4.Enabled = True
cmdnext2.Enabled = False

End Sub

Private Sub cmdnext3_Click()
picresults.Cls 'clears the picbox of any data'
picresults2.Picture = LoadPicture(App.Path & "\pics\more pics\offense.bmp") 'displays new picture
picresults.Print "Passing the ball relies on the accuracy of a quarterback's throw" 'displays the written data'
picresults.Print "and the the reciever's hands. Passing usually is called when"
picresults.Print "offenses have to move the ball in large increments. Or when"
picresults.Print "little time is left in the game to play and the offense must"
picresults.Print "score quickly"
cmdnext3.Enabled = False

End Sub

Private Sub cmdnext4_Click()
picresults.Cls 'clears the picbox of any data'
picresults2.Picture = LoadPicture(App.Path & "\pics\more pics\catch2.bmp") 'displays new picture
picresults.Print "When the opposing team's quarterback throws the ball, a defender" 'displays the written data'
picresults.Print "may try to catch the ball. This is called and interception and it"
picresults.Print "results in the possession change."
cmdnext2.Enabled = False
cmdnext4.Enabled = False
End Sub

Private Sub cmdnext5_Click()
picresults.Cls 'clears the picbox of any data'
cmdnext6.Enabled = True
cmdnext5.Enabled = False
picresults2.Picture = LoadPicture(App.Path & "\pics\more pics\punt.bmp") 'displays new picture
picresults.Print "When a player does not reach a first down and decides to punt," 'displays the written data'
picresults.Print "a drop menu will appear in the play screen, and the player can "
picresults.Print "either choose to punt, kick a field goal, or go for it, by "
picresults.Print "pressing the A button to choose.  Punting the ball works the same"
picresults.Print "way with the power bar and then player pressing A to stop the bar"
picresults.Print "and kick the ball.  The only difference is the player has to hike"
picresults.Print "the ball by pressing the A button"



End Sub

Private Sub cmdnext6_Click()
picresults.Cls 'clears the picbox of any data'
cmdnext6.Enabled = False
picresults2.Picture = LoadPicture(App.Path & "\pics\more pics\fieldgoal2.bmp") 'displays new picture
picresults.Print "When the player chooses to kick a field goal there" 'displays the written data'
picresults.Print "is no power bar that appears.  Instead there is an arrow "
picresults.Print "that moves up and down indicated which way the kicker will"
picresults.Print "kick the ball.  The key to making the field goal is to stop "
picresults.Print "the arrow right in the middle.  The player also needs to hike"
picresults.Print "the ball and stop the arrow by pressing the A button."


End Sub

Private Sub cmdoffense_Click()
picresults2.Cls 'clears the picbox of any data'
picresults.Cls 'clears the picbox of any data'
picresults.Print "When a team is on offense, its goal is to advance the ball up the" 'displays written information
picresults.Print "field and into the other team's endzone. When this happens the"
picresults.Print "offense scores a touchdown."
picresults.Print

picresults2.Picture = LoadPicture(App.Path & "\pics\image5.bmp") 'loads the picture file into the picbox'
cmdnext1.Enabled = True


End Sub

Private Sub cmdplays_Click()
picresults.Cls 'clears the picbox of any data'
picresults2.Picture = LoadPicture(App.Path & "\pics\plays.gif")
picresults.Print "How to Pick Plays"
picresults.Print "When playing Tecmo Super Bowl, a player needs to pick plays on " 'displays the written data'
picresults.Print "offense and defense.   When a player gets to pick, a play a pop up "
picresults.Print "screen will appear and there will be 8 different plays to choose from."
picresults.Print "This occurs on both offense and defense.  The player can choose a "
picresults.Print "passing play or a running play on offense.  To choose one of the "
picresults.Print "plays a passing or running the player looks at a controller under "
picresults.Print "each play, which tells them which way to hold the arrow key down, and "
picresults.Print "which button to push.  For example if a controller had the left part "
picresults.Print "of the directional pad highlighted, and then the top button "
picresults.Print "highlighted, the player would hold down the left arrow and hit the "
picresults.Print "top button.  The player does the same thing when picking any sort "
picresults.Print "of play."
End Sub

Private Sub Command1_Click()
picresults2.Cls 'clears the picbox of any data'
picresults.Cls 'clears the picbox of any data'
picresults2.Picture = LoadPicture(App.Path & "\pics\image3.bmp")
picresults.Print "The side without possession of the ball, the defense, must" 'displays the written data'
picresults.Print "try to stop the offense's attempts to move the ball"
cmdnext2.Enabled = True

End Sub

Private Sub Command2_Click()
picresults.Cls
picresults2.Cls 'clears the picbox of any data'
picresults.Print "Special teams are incredibly important and can decide" 'displays the written data'
picresults.Print "the game.At the beginning of the game, after a player"
picresults.Print "scores, and at the start of the third quarter, a player"
picresults.Print "may need to kickoff.  In order to kickoff the game will "
picresults.Print "have a power bar that increases and then resets continuously.  "
picresults.Print "To kick the ball all the player needs to do is hit the A button."
picresults.Print "Depending on how far the ball goes, depends on how high the power "
picresults.Print "bar increased before resetting.  The player needs to be careful "
picresults.Print "when kicking off and make sure not to hit the A button too late; "
picresults.Print "otherwise it will be a very short kickoff."
cmdnext5.Enabled = True
picresults2.Picture = LoadPicture(App.Path & "\pics\image2.bmp")
End Sub
