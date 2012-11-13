VERSION 5.00
Begin VB.Form frmTitle 
   BackColor       =   &H00004080&
   Caption         =   "Yon Dungeonier"
   ClientHeight    =   3975
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   ScaleHeight     =   3975
   ScaleWidth      =   7275
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   735
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000000FF&
      Caption         =   "Quit"
      Height          =   735
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton cmdTutorial 
      BackColor       =   &H0000FFFF&
      Caption         =   "Learn How To Play"
      Enabled         =   0   'False
      Height          =   735
      Left            =   2640
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3000
      Width           =   1935
   End
   Begin VB.CommandButton cmdBegin 
      BackColor       =   &H0000C000&
      Caption         =   "Begin the Adventure"
      Enabled         =   0   'False
      Height          =   735
      Left            =   480
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3000
      Width           =   1575
   End
   Begin VB.PictureBox picTitle 
      Height          =   2655
      Left            =   960
      ScaleHeight     =   2595
      ScaleWidth      =   5235
      TabIndex        =   0
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "frmTitle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This form acts as the Title Screen.  From here, the user can quit the
'program, go through the tutorial to learn how to play, or begin the main
'aspect of the game.  To do the tutorial or main game, the user must first
'hit the load button.

Private Sub cmdBegin_Click()
'this loads the first room of the main dungeon.  It does this by hiding
'the title and showing the first form of the main section.
frmTitle.Hide
frmDungeon1.Show
End Sub

Private Sub cmdLoad_Click()
'the load button reads all files into their respective arrays, for use in the
'tutorial or main game. It also sets the starting value for some flag variables
'in case the user decides to move back to rooms they've previously been to without
'the initial puzzles reseting.
A = 0
Open App.Path & "\TutorialActions.txt" For Input As #1
Do Until EOF(1)
    A = A + 1
    Input #1, TutorialActions(A)
Loop
Close #1

Open App.Path & "\DungeonActions1.txt" For Input As #2
C = 0
Do Until EOF(2)
    C = C + 1
    Input #2, DungeonActions1(C)
Loop
Close #2

Open App.Path & "\DungeonActions2.txt" For Input As #3
E = 0
Do Until EOF(3)
    E = E + 1
    Input #3, DungeonActions2(E)
Loop
Close #3

Open App.Path & "\DungeonActions3.txt" For Input As #4
F = 0
Do Until EOF(4)
    F = F + 1
    Input #4, DungeonActions3(F)
Loop
Close

Open App.Path & "\DungeonActions4.txt" For Input As #5
g = 0
Do Until EOF(5)
    g = g + 1
    Input #5, DungeonActions4(g)
Loop
Close #5

Open App.Path & "\DungeonActions5.txt" For Input As #6
h = 0
Do Until EOF(6)
    h = h + 1
    Input #6, DungeonActions5(h)
Loop
Close #6
'this allows the user to start playing, now that the arrays have been properly set.
cmdLoad.Enabled = False
cmdBegin.Enabled = True
cmdTutorial.Enabled = True
puzzle1 = False
riddle = False
biglock = True
'this gives instructions to the user based on what they may want to do
picTitle.Print "If you have never played before, you should try the tutorial."
picTitle.Print "If you now what you're doing, go ahead and get started."
picTitle.Print "If you don't want to play, hit Quit."
End Sub

Private Sub cmdQuit_Click()
'quits the program
End
End Sub

Private Sub cmdTutorial_Click()
'loads the tutorial and disables it; the tutorial may only be done one time once the program is opened.
frmTitle.Hide
frmTutorial1.Show
cmdTutorial.Enabled = False
End Sub

