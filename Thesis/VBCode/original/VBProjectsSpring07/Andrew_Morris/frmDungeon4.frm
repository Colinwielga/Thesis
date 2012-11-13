VERSION 5.00
Begin VB.Form frmDungeon4 
   BackColor       =   &H00004000&
   Caption         =   "The Puzzler"
   ClientHeight    =   3990
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   ScaleHeight     =   3990
   ScaleWidth      =   7275
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDungeon4 
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   3000
      Width           =   2895
   End
   Begin VB.CommandButton cmdSubmit4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Submit"
      Height          =   375
      Left            =   1080
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton cmdLook4 
      BackColor       =   &H00C0C000&
      Caption         =   "Look Around"
      Height          =   375
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton cmdQuit4 
      BackColor       =   &H000000FF&
      Caption         =   "Quit Program"
      Height          =   375
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3480
      Width           =   2175
   End
   Begin VB.PictureBox picDungeon4 
      Height          =   2655
      Left            =   1080
      ScaleHeight     =   2595
      ScaleWidth      =   5235
      TabIndex        =   0
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "frmDungeon4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim solve As Boolean
Dim answer1 As Integer
Dim answer2 As String
Dim guess As Integer
'this room's puzzle is a riddle that the user must solve to get a key to open the
'locked door in the central chamber.  The user gets 3 guesses, if they are all wrong
'the player charcter is killed, and the game is over.

Private Sub cmdLook4_Click()
'this provides a description of the room based on what related conditions are met.
picDungeon4.Cls
If solve = False Then
    picDungeon4.Print "The room is like another world entirely. It is full"
    picDungeon4.Print "of all manner of wondrous things.  In the center is"
    picDungeon4.Print "an old man.  Maybe you should TALK to him."
End If
If solve = True Then
    picDungeon4.Print "You glance back at the old man, the Puzzler, and can't"
    picDungeon4.Print "help but wonder if there is more to him than there seems."
End If
picDungeon4.Print "There is a door leading back to the central room to the WEST."
End Sub

Private Sub cmdQuit4_Click()
'ends the program
End
End Sub

Private Sub cmdSubmit4_Click()
'this checks the user's submission to see if it is acceptable for this room and
'provides a description of the result depending on the current conditions.
picDungeon4.Cls
submit = txtDungeon4.Text
If LCase(submit) = DungeonActions4(1) And solve = False Then
    MsgBox "That old guy gives you the heeby-jeebies.  You head back to the center room.", , "Weird"
    frmDungeon4.Hide
    frmDungeon2.Show
End If
If LCase(submit) = DungeonActions4(1) And solve = True Then
    MsgBox "You head back to the center room, having successfully solved the riddle.", , "You're the Smart One"
    frmDungeon4.Hide
    frmDungeon2.Show
End If
If LCase(submit) = DungeonActions4(2) And solve = False Then
    answer1 = InputBox("The old man says 'I have a riddle for ye; are ye prepared to answer it?' 1=yes, 2=no", "Riddle Time")
    If answer1 = 1 Then
        Do Until solve = True Or guess = 3
        guess = guess + 1
            answer2 = InputBox("With every one of these you take you leave a few behind. What are they?", "Riddle Time")
            If LCase(answer2) = DungeonActions4(3) And solve = False Then
                solve = True
                riddle = True
                bigkey = True
                MsgBox "Correct! For your reward you may take this KEY!", , "Way to Go!"
            Else
                MsgBox "Incorrect, but you get a total of 3 guesses", , "Wrong Guess"
            End If
        Loop
    Else
        MsgBox "If you change your mind, I'll be here.", , "Creepy Old Guy"
    End If
    If solve = False And answer1 = 1 Then
        MsgBox "'You are a fool, and the universe shall suffer your existence no longer!'", , "Uh Oh"
        MsgBox "The old man fries you with some sort of lightning spell. You dead.", , "Ffffzzzzz....."
        MsgBox "Better luck next time", , "Game Over"
        End
    End If
End If
If LCase(submit) = DungeonActions4(2) And solve = True Then
    picDungeon4.Print "The old man says 'I having nothing more for you. Now go.'"
End If
End Sub

Private Sub Form_Load()
'sets the conditions and variables for this room upon first entry.
guess = 0
If riddle = False Then
    solve = False
End If
End Sub
