VERSION 5.00
Begin VB.Form frmMatchPictures 
   BackColor       =   &H00000080&
   Caption         =   "Match Pictures"
   ClientHeight    =   5220
   ClientLeft      =   4635
   ClientTop       =   3330
   ClientWidth     =   7170
   LinkTopic       =   "Form1"
   ScaleHeight     =   5220
   ScaleWidth      =   7170
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   4920
      ScaleHeight     =   1635
      ScaleWidth      =   1875
      TabIndex        =   8
      Top             =   2160
      Width           =   1935
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Batang"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1440
      Width           =   1935
   End
   Begin VB.CommandButton cmdVictoria 
      Enabled         =   0   'False
      Height          =   1095
      Left            =   240
      Picture         =   "frmMatchPictures.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdJames 
      Enabled         =   0   'False
      Height          =   1095
      Left            =   1800
      Picture         =   "frmMatchPictures.frx":0CF2
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdJacob 
      Enabled         =   0   'False
      Height          =   1095
      Left            =   3360
      Picture         =   "frmMatchPictures.frx":1ACD
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdAlice 
      Enabled         =   0   'False
      Height          =   1095
      Left            =   1800
      Picture         =   "frmMatchPictures.frx":27D7
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdBella 
      Enabled         =   0   'False
      Height          =   1095
      Left            =   240
      Picture         =   "frmMatchPictures.frx":7445
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdEdward 
      Enabled         =   0   'False
      Height          =   1095
      Left            =   3360
      Picture         =   "frmMatchPictures.frx":7C9E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdReturnMain 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Click to return to the main menu"
      BeginProperty Font 
         Name            =   "Batang"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Label lblInstructions 
      BackColor       =   &H00000080&
      Caption         =   $"frmMatchPictures.frx":8932
      BeginProperty Font 
         Name            =   "Batang"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   960
      TabIndex        =   10
      Top             =   840
      Width           =   5415
   End
   Begin VB.Label lblMatch 
      BackColor       =   &H00000080&
      Caption         =   "Match the Picture with the Correct Character Name"
      BeginProperty Font 
         Name            =   "Batang"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   1200
      TabIndex        =   9
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "frmMatchPictures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project Name: Twilight
'Form Name: frmMatchPictures
'Author: Mollie Land
'Date Written: 3/22/2009
'Objective: This is a program for the user to match a character's name with the associated picture using buttons
'Initially the user must click a start button to create a list of names to choose from when answering which character the picture is of
'Then the user is free to click a character's button where an input box will pop up asking which character it is
'There is also a points option in this game to tell the user how many points they have received

'Dim Global Variables
Dim Answer As String

Private Sub cmdAlice_Click()
'Get the answer from the user of which character this is
Answer = InputBox("Which Twilight character is this?", "Match Game Name")

'did they get the right answer?
'by using If/Then/Else the user can see if they got the correct answer and if not what the actual answer is
'the user's points are also shown with this statement
If Answer = "Alice" Then
    Points = Points + 1
    MsgBox "You are correct! You have " & Points & " points!", , "Correct Answer"
Else
    MsgBox "That is the incorrect answer. This is Alice.", , "Incorrect Answer"
End If

'disable this button since the user has already answered the question
cmdAlice.Enabled = False
End Sub

Private Sub cmdBella_Click()
'Get the answer from the user of which character this is
Answer = InputBox("Which Twilight character is this?", "Match Game Name")

'did they get the right answer?
'by using If/Then/Else the user can see if they got the correct answer and if not what the actual answer is
'the user's points are also shown with this statement
If Answer = "Bella" Then
    Points = Points + 1
    MsgBox "You are correct! You have " & Points & " points!", , "Correct Answer"
Else
    MsgBox "That is the incorrect answer. This is Bella.", , "Incorrect Answer"
End If

'disable this button since the user has already answered the question
cmdBella.Enabled = False

End Sub

Private Sub cmdEdward_Click()
'Get the answer from the user of which character this is
Answer = InputBox("Which Twilight character is this?", "Match Game Name")

'did they get the right answer?
'by using If/Then/Else the user can see if they got the correct answer and if not what the actual answer is
'the user's points are also shown with this statement
If Answer = "Edward" Then
    Points = Points + 1
    MsgBox "You are correct! You have " & Points & " points!", , "Correct Answer"
Else
    MsgBox "That is the incorrect answer. This is Edward.", , "Incorrect Answer"
End If

'disable this button since the user has already answered the question
cmdEdward.Enabled = False


End Sub

Private Sub cmdJacob_Click()
'Get the answer from the user of which character this is
Answer = InputBox("Which Twilight character is this?", "Match Game Name")

'did they get the right answer?
'by using If/Then/Else the user can see if they got the correct answer and if not what the actual answer is
'the user's points are also shown with this statement
If Answer = "Jacob" Then
    Points = Points + 1
    MsgBox "You are correct! You have " & Points & " points!", , "Correct Answer"
Else
    MsgBox "That is the incorrect answer. This is Jacob.", , "Incorrect Answer"
End If

'disable this button since the user has alreay answered the question
cmdJacob.Enabled = False

End Sub

Private Sub cmdJames_Click()
'Get the answer from the user of which character this is
Answer = InputBox("Which Twilight character is this?", "Match Game Name")

'did they get the right answer?
'by using If/Then/Else the user can see if they got the correct answer and if not what the actual answer is
'the user's points are also shown with this statement
If Answer = "James" Then
    Points = Points + 1
    MsgBox "You are correct! You have " & Points & " points!", , "Correct Answer"
Else
    MsgBox "That is the incorrect answer. This is James.", , "Incorrect Answer"
End If

'disable this button since the user has already answered the question
cmdJames.Enabled = False

End Sub

Private Sub cmdReturnMain_Click()
    'clear the picture box
    picResults.Cls
    
    'enable the start button
    cmdStart.Enabled = True
    
    'return to main menu, hiding the match characters form
    frmStart.Show
    frmMatchPictures.Hide
    
End Sub


Private Sub cmdStart_Click()
    'Dim variables subject to this button
    Dim MatchGameNames(1 To 10) As String
    
    'Enable the charcater picture buttons so once the file is read the user can begin the matching game
    cmdEdward.Enabled = True
    cmdBella.Enabled = True
    cmdAlice.Enabled = True
    cmdJacob.Enabled = True
    cmdJames.Enabled = True
    cmdVictoria.Enabled = True
    
    'initialize Points and CTR, resetting the numbers
    Points = 0
    CTR = 0
    
    'open the file with the names of the characters
    Open App.Path & "/MatchGameNames.txt" For Input As #1
    
    'read the file and close it once it has been read
    Do While Not EOF(1)
        CTR = CTR + 1
        Input #1, MatchGameNames(CTR)
    Loop
    Close (1)
    
    'print the character names in a picture box for future reference
    For J = 1 To CTR
        picResults.Print MatchGameNames(J)
    Next J
    
    'disable this button since the user has read the array and no longer needs to read it again
    cmdStart.Enabled = False
    
    
End Sub

Private Sub cmdVictoria_Click()
'Get the answer from the user of which character this is
Answer = InputBox("Which Twilight character is this?", "Match Game Name")

'did they get the right answer?
'by using If/Then/Else the user can see if they got the correct answer and if not what the actual answer is
'the user's points are also shown with this statement
If Answer = "Victoria" Then
    Points = Points + 1
    MsgBox "You are correct! You have " & Points & " points!", , "Correct Answer"
Else
    MsgBox "That is the incorrect answer. This is Victoria.", , "Incorrect Answer"
End If

'disable this button since the user has already answered this question
cmdVictoria.Enabled = False
End Sub
