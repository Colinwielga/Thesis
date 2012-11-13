VERSION 5.00
Begin VB.Form frmTutorial1 
   BackColor       =   &H00FF8080&
   Caption         =   "A Simple Room"
   ClientHeight    =   3990
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   ScaleHeight     =   3990
   ScaleWidth      =   7275
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000000FF&
      Caption         =   "Quit Program"
      Height          =   375
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3480
      Width           =   2175
   End
   Begin VB.CommandButton cmdEndTutorial 
      BackColor       =   &H0000FFFF&
      Caption         =   "I'm done here, take me back"
      Height          =   375
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3000
      Width           =   2175
   End
   Begin VB.CommandButton cmdLook 
      BackColor       =   &H00C0C000&
      Caption         =   "Look Around"
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton cmdSubmit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Submit"
      Height          =   375
      Left            =   960
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3480
      Width           =   1335
   End
   Begin VB.TextBox txtTutorial1 
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   3000
      Width           =   2895
   End
   Begin VB.PictureBox picTutorial1 
      Height          =   2655
      Left            =   960
      ScaleHeight     =   2595
      ScaleWidth      =   5235
      TabIndex        =   0
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "frmTutorial1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Notes As Boolean
Dim key As Boolean
Dim lock1 As Boolean
Dim B As Integer
Dim flag1 As Boolean
Dim test As Integer
'this tutorial is designed to teach new users how to play the game.
'it introduces all basic actions and their respective input methods.

Private Sub cmdEndTutorial_Click()
'if the user decides they no longer want to do the tutorial, they can
'hit this button to return to the title screen.
frmTutorial1.Hide
frmTitle.Show
End Sub

Private Sub cmdLook_Click()
'this button provides a description of the room the player is currently
'in.  The description provided depends upon what conditions are met.
picTutorial1.Cls

If Notes = False And key = False Then
    picTutorial1.Print "You find yourself in a simple room. There are some NOTES"
    picTutorial1.Print "and a KEY on the table."
    MsgBox "You should now type GET NOTES in the text box. You can check the room any time by hitting 'Look Around'."
End If

If Notes = True And key = False Then
    picTutorial1.Print "A KEY lies on the table."
End If

If Notes = False And key = True Then
    picTutorial1.Print "Some NOTES lie on the table."
End If

If lock1 = True Then
    picTutorial1.Print "There is a locked door to the NORTH."
End If

If lock1 = False Then
    picTutorial1.Print "There is a door to the NORTH. It is now unlocked."
End If

End Sub

Private Sub cmdQuit_Click()
'quits the program
End
End Sub

Private Sub cmdSubmit_Click()
'this is the user's method of performing actions.  If the user's input
'matches the action stored in the action array for this room, then something
'in the room will change and the user will make progress.
submitTutorial = txtTutorial1.Text
picTutorial1.Cls
B = 0
Do Until B = A
    B = B + 1
    If LCase(submitTutorial) = TutorialActions(1) And Notes = False Then
        Notes = True
        picTutorial1.Print "You pick up the NOTES off the table."
        MsgBox "Good job. Now check the notes by typing USE NOTES", , "Notes"
    End If
    If LCase(submitTutorial) = TutorialActions(2) And Notes = True Then
        picTutorial1.Print "The Notes say: Nice work! Now repeat the process to get the KEY."
        picTutorial1.Print "You put the notes back down."
        Notes = False
    End If
    If LCase(submitTutorial) = TutorialActions(3) And key = False Then
        key = True
        picTutorial1.Print "You pick up the key.  Try using it on the door by typing USE KEY."
    End If
    If LCase(submitTutorial) = TutorialActions(4) And key = True And flag1 = False Then
        lock1 = False
        picTutorial1.Print "You use the key on the door. It is now unlocked."
        picTutorial1.Print "To go through the door type 'NORTH'."
        picTutorial1.Print "You can only use keys on a certain lock."
        flag1 = True
    End If
    If LCase(submitTutorial) = TutorialActions(5) And lock1 = False Then
    test = InputBox("Before you can finish, you must answer a question on the spot.  What is 1+1?", "Input Test")
        If test = 2 Then
            MsgBox "Good job, you've finished the tutorial.  You should now be able to complete the main game!", , "Finish!"
            lock1 = True
            frmTutorial1.Hide
            frmTitle.Show
        Else
            MsgBox "No, that's not right.  But I'll let you try again.", , "Lucky Break"
        End If
    End If
Loop

End Sub

Private Sub Form_Load()
'this message starts the walkthrough for the tutorial by telling the user
'to hit the "look around" button.
MsgBox "This will teach you how to play. Start by hitting 'Look Around'.", , "Tutorial"
lock1 = True
flag1 = False
End Sub
