VERSION 5.00
Begin VB.Form frmCasino 
   BackColor       =   &H00000000&
   Caption         =   "Casino"
   ClientHeight    =   7395
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8475
   LinkTopic       =   "Form1"
   ScaleHeight     =   7395
   ScaleWidth      =   8475
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      TabIndex        =   10
      Top             =   6720
      Width           =   1695
   End
   Begin VB.CommandButton cmdInstructions 
      Caption         =   "Instructions"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      TabIndex        =   9
      Top             =   2880
      Width           =   1935
   End
   Begin VB.PictureBox picScore 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1200
      ScaleHeight     =   495
      ScaleWidth      =   1935
      TabIndex        =   8
      Top             =   2880
      Width           =   1935
   End
   Begin VB.PictureBox pic8 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      Height          =   1215
      Left            =   6360
      ScaleHeight     =   1215
      ScaleWidth      =   1695
      TabIndex        =   7
      Top             =   5280
      Width           =   1695
   End
   Begin VB.PictureBox pic4 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   6360
      ScaleHeight     =   1215
      ScaleWidth      =   1695
      TabIndex        =   6
      Top             =   3840
      Width           =   1695
   End
   Begin VB.PictureBox pic7 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   4440
      ScaleHeight     =   1215
      ScaleWidth      =   1695
      TabIndex        =   5
      Top             =   5280
      Width           =   1695
   End
   Begin VB.PictureBox pic3 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      Height          =   1215
      Left            =   4440
      ScaleHeight     =   1215
      ScaleWidth      =   1695
      TabIndex        =   4
      Top             =   3840
      Width           =   1695
   End
   Begin VB.PictureBox pic6 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   2520
      ScaleHeight     =   1215
      ScaleWidth      =   1695
      TabIndex        =   3
      Top             =   5280
      Width           =   1695
   End
   Begin VB.PictureBox pic2 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   2520
      ScaleHeight     =   1215
      ScaleWidth      =   1695
      TabIndex        =   2
      Top             =   3840
      Width           =   1695
   End
   Begin VB.PictureBox pic5 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   600
      ScaleHeight     =   1215
      ScaleWidth      =   1695
      TabIndex        =   1
      Top             =   5280
      Width           =   1695
   End
   Begin VB.PictureBox pic1 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   600
      ScaleHeight     =   1215
      ScaleWidth      =   1695
      TabIndex        =   0
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Image imgCasino 
      Height          =   2280
      Left            =   1200
      Picture         =   "frmCasino.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   6435
   End
End
Attribute VB_Name = "frmCasino"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: RPGCraze
'Form name: frmCasino
'Author: Justin Roth
'Date Written: Sunday, November 4th, 2007
'Objective of form: This form brings the user to a color guessing game.
        'This is where the user gets their intellegince points and cash.

Option Explicit

Private Sub cmdInstructions_Click()
    MsgBox "If you click on one of the boxes you will be asked to guess the color on the other side of the box (guess either yellow, blue, red, or green).  For every color you guess correctly, you will get one point. One point is equal to $100.00 and 1 intelligence point.", , "Instructions" 'Gives instructions to the user on how to play.
End Sub

Private Sub cmdBack_Click()
    frmCasino.Hide  'Goes back to the Map form.
End Sub

Private Sub Form_Load()
    Cash = Cash + 10    'Adds $10.00 to the user's cash funds.
    MsgBox "You found $10.00!", , "$10.00!!"    'Notifies the user that they found $10.00.
End Sub

Private Sub pic1_Click()
    Guess = InputBox("What color do you think is on the other side?")   'Asks the user to guess which color is on the other side and defines what "Guess" is.
    If Guess = "yellow" Then    'If the user guesses the right color (yellow) then they get a point.
        Score = Score + 1   'Increases the user's score by one for guessing correctly.
        Cash = Cash + 50    'Adds $50.00 to the user's cash funds for guessing correctly.
        MsgBox "You're right! You gained 1 intelligence point!", , "+1 intelligence!"   'Notifies the user that they were right and that they gained an intelligence point.
        pic1.BackColor = &HFFFF&    'Switches the boxes background color to yellow after a correct guess.
        picScore.Cls    'Clears the score for the next count.
        picScore.Print "Score:"; Score  'Prints the user's current score.
        pic1.Enabled = False    'Disables the button after the user guesses correctly.
    Else
        MsgBox "Sorry, it was yellow!", , "Sorry!"  'Notifies the user that their guess was wrong and tells them what the correct one was.
        pic1.Visible = False    'Makes the button dissapear after an incorrect guess.
    End If
End Sub

Private Sub pic2_Click()
    Guess = InputBox("What color do you think is on the other side?")   'Asks the user to guess which color is on the other side and defines what "Guess" is.
    If Guess = "blue" Then  'If the user guesses the right color (blue) then they get a point.
        Score = Score + 1   'Increases the user's score by one for guessing correctly.
        Cash = Cash + 50    'Adds $50.00 to the user's cash funds for guessing correctly.
        MsgBox "You're right! You gained 1 intelligence point!", , "+1 intelligence!"   'Notifies the user that they were right and that they gained an intelligence point.
        pic2.BackColor = &HFF0000   'Switches the boxes background color to blue after a correct guess.
        picScore.Cls    'Clears the score for the next count.
        picScore.Print "Score:"; Score  'Prints the user's current score.
        pic2.Enabled = False    'Disables the button after the user guesses correctly.
    Else
        MsgBox "Sorry, it was blue!", , "Sorry!"    'Notifies the user that their guess was wrong and tells them what the correct one was.
        pic2.Visible = False    'Makes the button dissapear after an incorrect guess.
    End If
End Sub

Private Sub pic3_Click()
    Guess = InputBox("What color do you think is on the other side?")   'Asks the user to guess which color is on the other side and defines what "Guess" is.
    If Guess = "red" Then   'If the user guesses the right color (red) then they get a point.
        Score = Score + 1   'Increases the user's score by one for guessing correctly.
        Cash = Cash + 50    'Adds $50.00 to the user's cash funds for guessing correctly.
        MsgBox "You're right! You gained 1 intelligence point!", , "+1 intelligence!"   'Notifies the user that they were right and that they gained an intelligence point.
        pic3.BackColor = &HFF&  'Switches the boxes background color to red after a correct guess.
        picScore.Cls    'Clears the score for the next count.
        picScore.Print "Score:"; Score  'Prints the user's current score.
        pic3.Enabled = False    'Disables the button after the user guesses correctly.
    Else
        MsgBox "Sorry, it was red!", , "Sorry!" 'Notifies the user that their guess was wrong and tells them what the correct one was.
        pic3.Visible = False    'Makes the button dissapear after an incorrect guess.
    End If
End Sub

Private Sub pic4_Click()
    Guess = InputBox("What color do you think is on the other side?")   'Asks the user to guess which color is on the other side and defines what "Guess" is.
        If Guess = "yellow" Then    'If the user guesses the right color (yellow) then they get a point.
            Score = Score + 1   'Increases the user's score by one for guessing correctly.
            Cash = Cash + 50    'Adds $50.00 to the user's cash funds for guessing correctly.
            MsgBox "You're right! You gained 1 intelligence point!", , "+1 intelligence!"   'Notifies the user that they were right and that they gained an intelligence point.
            pic4.BackColor = &HFFFF&    'Switches the boxes background color to yellow after a correct guess.
            picScore.Cls    'Clears the score for the next count.
            picScore.Print "Score:"; Score  'Prints the user's current score.
            pic4.Enabled = False    'Disables the button after the user guesses correctly.
        Else
            MsgBox "Sorry, it was yellow!", , "Sorry!"  'Notifies the user that their guess was wrong and tells them what the correct one was.
            pic4.Visible = False    'Makes the button dissapear after an incorrect guess.
        End If
End Sub

Private Sub pic5_Click()
    Guess = InputBox("What color do you think is on the other side?")   'Asks the user to guess which color is on the other side and defines what "Guess" is.
    If Guess = "blue" Then  'If the user guesses the right color (blue) then they get a point.
        Score = Score + 1   'Increases the user's score by one for guessing correctly.
        Cash = Cash + 50    'Adds $50.00 to the user's cash funds for guessing correctly.
        MsgBox "You're right! You gained 1 intelligence point!", , "+1 intelligence!"   'Notifies the user that they were right and that they gained an intelligence point.
        pic5.BackColor = &HFF0000   'Switches the boxes background color to blue after a correct guess.
        picScore.Cls    'Clears the score for the next count.
        picScore.Print "Score:"; Score  'Prints the user's current score.
        pic5.Enabled = False    'Disables the button after the user guesses correctly.
    Else
        MsgBox "Sorry, it was blue!", , "Sorry!"    'Notifies the user that their guess was wrong and tells them what the correct one was.
        pic5.Visible = False    'Makes the button dissapear after an incorrect guess.
        End If
End Sub

Private Sub pic6_Click()
    Guess = InputBox("What color do you think is on the other side?")   'Asks the user to guess which color is on the other side and defines what "Guess" is.
    If Guess = "green" Then 'If the user guesses the right color (green) then they get a point.
        Score = Score + 1   'Increases the user's score by one for guessing correctly.
        Cash = Cash + 50    'Adds $50.00 to the user's cash funds for guessing correctly.
        MsgBox "You're right! You gained 1 intelligence point!", , "+1 intelligence!"   'Notifies the user that they were right and that they gained an intelligence point.
        pic6.BackColor = &HC000&    'Switches the boxes background color to green after a correct guess.
        picScore.Cls    'Clears the score for the next count.
        picScore.Print "Score:"; Score  'Prints the user's current score.
        pic6.Enabled = False    'Disables the button after the user guesses correctly.
    Else
        MsgBox "Sorry, it was green!", , "Sorry!"   'Notifies the user that their guess was wrong and tells them what the correct one was.
        pic6.Visible = False    'Makes the button dissapear after an incorrect guess.
    End If
End Sub

Private Sub pic7_Click()
    Guess = InputBox("What color do you think is on the other side?")   'Asks the user to guess which color is on the other side and defines what "Guess" is.
    If Guess = "red" Then   'If the user guesses the right color (red) then they get a point.
        Score = Score + 1   'Increases the user's score by one for guessing correctly.
        Cash = Cash + 50    'Adds $50.00 to the user's cash funds for guessing correctly.
        MsgBox "You're right! You gained 1 intelligence point!", , "+1 intelligence!"   'Notifies the user that they were right and that they gained an intelligence point.
        pic7.BackColor = &HFF&  'Switches the boxes background color to red after a correct guess.
        picScore.Cls    'Clears the score for the next count.
        picScore.Print "Score:"; Score  'Prints the user's current score.
        pic7.Enabled = False    'Disables the button after the user guesses correctly.
    Else
        MsgBox "Sorry, it was red!", , "Sorry!" 'Notifies the user that their guess was wrong and tells them what the correct one was.
        pic7.Visible = False    'Makes the button dissapear after an incorrect guess.
    End If
End Sub

Private Sub pic8_Click()
    Guess = InputBox("What color do you think is on the other side?")   'Asks the user to guess which color is on the other side and defines what "Guess" is.
    If Guess = "green" Then 'If the user guesses the right color (green) then they get a point.
        Score = Score + 1   'Increases the user's score by one.
        Cash = Cash + 50    'Adds $50.00 to the user's cash funds for guessing correctly.
        MsgBox "You're right! You gained 1 intelligence point!", , "+1 intelligence!"   'Notifies the user that they were right and that they gained an intelligence point.
        pic8.BackColor = &HC000&    'Switches the boxes background color to green after a correct guess.
        picScore.Cls    'Clears the score for the next count.
        picScore.Print "Score:"; Score  'Prints the user's current score.
        pic8.Enabled = False    'Disables the button after the user guesses correctly.
    Else
        MsgBox "Sorry, it was green!", , "Sorry!"   'Notifies the user that their guess was wrong and tells them what the correct one was.
        pic8.Visible = False    'Makes the button dissapear after an incorrect guess.
    End If
End Sub
