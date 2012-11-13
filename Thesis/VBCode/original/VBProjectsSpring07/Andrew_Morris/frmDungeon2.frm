VERSION 5.00
Begin VB.Form frmDungeon2 
   BackColor       =   &H00000040&
   Caption         =   "The Tower"
   ClientHeight    =   3990
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   ScaleHeight     =   3990
   ScaleWidth      =   7275
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picDungeon2 
      Height          =   2655
      Left            =   1080
      ScaleHeight     =   2595
      ScaleWidth      =   5235
      TabIndex        =   4
      Top             =   120
      Width           =   5295
   End
   Begin VB.CommandButton cmdQuit2 
      BackColor       =   &H000000FF&
      Caption         =   "Quit Program"
      Height          =   375
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3480
      Width           =   2175
   End
   Begin VB.CommandButton cmdLook2 
      BackColor       =   &H00C0C000&
      Caption         =   "Look Around"
      Height          =   375
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton cmdSubmit2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Submit"
      Height          =   375
      Left            =   1080
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3480
      Width           =   1335
   End
   Begin VB.TextBox txtDungeon2 
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   3000
      Width           =   2895
   End
End
Attribute VB_Name = "frmDungeon2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'this room is the central chamber.  There are no puzzles to solve here, but
'completing other tasks in other rooms will eventually allow the player to
'exit the dungeon and complete the game.

Private Sub cmdLook2_Click()
'this provides a description of the room based on what condition have or have not been
'fulfilled.  these conditions are met in other rooms.
picDungeon2.Cls
If puzzle1 = False And biglock = True Then
    picDungeon2.Print "This room is very elaborately decorated.  There are doors"
    picDungeon2.Print "to the EAST, WEST, and back SOUTH. There is another door"
    picDungeon2.Print "to the NORTH, but it is beyond a wide gap you couldn't possibly JUMP."
    picDungeon2.Print "It's locked too.  You'd best do some exploring. A large symbol of an eye"
    picDungeon2.Print "adorns the center of the floor.  Creepy."
End If
If puzzle1 = True And biglock = True Then
    picDungeon2.Print "The way NORTH now lies open, but the lock remains, preventing your progress."
    picDungeon2.Print "Other doors are to the EAST, WEST, and SOUTH."
    picDungeon2.Print "You could almost swear that eye is watching you..."
End If
If puzzle1 = True And biglock = False Then
    picDungeon2.Print "No more obstacles remain in your path NORTH, what are you waiting for?"
    picDungeon2.Print "Other doors are to the EAST, WEST, and SOUTH."
End If
End Sub

Private Sub cmdQuit2_Click()
'ends the program
End
End Sub

Private Sub cmdSubmit2_Click()
'this checks to see if the users input matches the action array for this room.
'the actions will do different things depending on what condition have been fulfilled.
submit = txtDungeon2.Text
picDungeon2.Cls
If LCase(submit) = DungeonActions2(1) Then
    MsgBox "You head WEST, hoping to make progress on your escape.", , "Door #1"
    frmDungeon2.Hide
    frmDungeon3.Show
End If
If LCase(submit) = DungeonActions2(2) And puzzle1 = True And biglock = False Then
    MsgBox "You open the door, and go through to the other side...", , "Door #3"
    MsgBox "You pass through the door and find yourself outside!  You've done it! You've escaped!", , "Booyakasha!"
    frmDungeon2.Hide
    frmEndScreen.Show
End If
If LCase(submit) = DungeonActions2(2) And puzzle1 = False Then
    picDungeon2.Print "You do your best Stretch Armstrong impression, but to no avail."
    picDungeon2.Print "Better try some other way."
End If
If LCase(submit) = DungeonActions2(2) And puzzle1 = True Then
    picDungeon2.Print "You can reach the NORTH door now, but can't go through it;"
    picDungeon2.Print "its still locked. You'll need a key."
End If
If LCase(submit) = DungeonActions2(4) Then
    MsgBox "You head EAST, hoping to make progress on your escape.", , "Door #2"
    frmDungeon2.Hide
    frmDungeon4.Show
End If
If LCase(submit) = DungeonActions2(3) Then
    MsgBox "You head back to the cage. There was something cozy about it you pine for, appatently.", , "Backpedaling"
    frmDungeon2.Hide
    frmDungeon1.Show
End If
If LCase(submit) = DungeonActions2(5) And puzzle1 = False Then
    MsgBox "You imagine you're in The Matrix and leap with all your might.  Too bad the Matrix was just a movie. You dead.", , "Good One"
    MsgBox "Better luck next time.", , "Game Over"
    End
End If
If LCase(submit) = DungeonActions2(5) And puzzle1 = True Then
    picDungeon2.Print "You leap mightily, but it doesn't really matter since"
    picDungeon2.Print "the chasm is gone anyway."
End If
If LCase(submit) = DungeonActions2(6) And puzzle1 = True And biglock = True Then
    biglock = False
    picDungeon2.Print "You unlock the NORTH door.  All paths now lie open."
End If
End Sub

