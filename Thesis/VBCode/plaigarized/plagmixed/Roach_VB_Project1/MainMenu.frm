VERSION 5.00
Begin VB.Form MainMenu 
   BackColor       =   &H80000009&
   Caption         =   "Form1"
   ClientHeight    =   11535
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15915
   LinkTopic       =   "Form1"
   ScaleHeight     =   10636.64
   ScaleMode       =   0  'User
   ScaleWidth      =   15915
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00000080&
      Caption         =   "Quit"
      Height          =   975
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9240
      Width           =   3495
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00008000&
      Caption         =   "Go Back"
      Height          =   975
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7800
      Width           =   3495
   End
   Begin VB.CommandButton cmdTrivia 
      Caption         =   "Trivia"
      Height          =   975
      Left            =   11520
      TabIndex        =   4
      Top             =   5040
      Width           =   3495
   End
   Begin VB.CommandButton cmdSanta 
      Caption         =   "Lyrics"
      Height          =   975
      Left            =   11520
      TabIndex        =   3
      Top             =   6360
      Width           =   3495
   End
   Begin VB.CommandButton cmdSleigh 
      Caption         =   "See Santa's Sleigh Team"
      Height          =   975
      Left            =   11520
      TabIndex        =   2
      Top             =   3600
      Width           =   3495
   End
   Begin VB.CommandButton cmdCharacters 
      Caption         =   "Characters"
      Height          =   975
      Left            =   11520
      TabIndex        =   1
      Top             =   2160
      Width           =   3495
   End
   Begin VB.CommandButton cmdHistory 
      Caption         =   "History"
      Height          =   975
      Left            =   11520
      TabIndex        =   0
      Top             =   720
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   10755
      Left            =   960
      Picture         =   "MainMenu.frx":0000
      Top             =   360
      Width           =   8700
   End
End
Attribute VB_Name = "MainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBack_Click()
    IntroRudolph.Show
    MainMenu.Hide
End Sub

Private Sub cmdCharacters_Click()
    Characters.Show
    MainMenu.Hide
End Sub

Private Sub cmdHistory_Click()
'Project Name: Rudolph
'Form Name: MainMenu
'Author: Patrick Roach
'Date written: February 25, 2010
'Objective: Allow you access to all the areas of the project
    History.Show
    MainMenu.Hide
    
End Sub
'Project Name: Rudolph
'Project Name: Rudolph
'Project Name: Rudolph
'Project Name: Rudolph
'Project Name: Rudolph
'Project Name: Rudolph
'Project Name: Rudolph

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdSanta_Click()
    Lyrics.Show
'Project Name: Rudolph
    MainMenu.Hide
End Sub

Private Sub cmdSleigh_Click()
    Sleigh.Show
    MainMenu.Hide
    
End Sub

Private Sub cmdTrivia_Click()
   'Project Name: Rudolph
'Form Name: Part of MainMenu
'Project Name: Rudolph
'Project Name: Rudolph
'Project Name: Rudolph
'Author: Patrick Roach
'Date written: February 25, 2010
'Objective: Fun Trivia Game
    Dim A As Integer
    Dim q1 As String, q2 As String, q3 As String, q4 As String, q5 As String
    
    A = 0
'Project Name: Rudolph
'Project Name: Rudolph
'Project Name: Rudolph
    
    q1 = InputBox("Who does Rudolph fall in love with?", "Question 1 or 5")
        If q1 = "Clarice" Or q1 = "clarice" Then
            A = 1 + 8 + A - 8
        End If
    q2 = InputBox("What does Yukon say Bumbles do?", "Question 2 or 5")
        If q2 = "Bounce" Or q2 = "bounce" Then
            A = 1 + 8 + A - 8
        End If
    q3 = InputBox("What is the name of the King of the Island of Misfit Toys?", "Question 3 or 5")
        If q3 = "king moonracer" Or q3 = "King moonracer" Or q3 = "King Moonracer" Then
            A = 1 + 8 + A - 8
'Project Name: Rudolph
'Project Name: Rudolph
        End If
    q4 = InputBox("What the fog as thick as?", "Question 4 or 5")
        If q4 = "peanut butter" Or q4 = "Peanut butter" Or q4 = "peanutbutter" Or q4 = "Peanutbutter" Or q4 = "pea soup" Or q4 = "Pea soup" Then
            A = 1 + 8 + A - 8
        End If
    q5 = InputBox("Who puts the star on the tree top?", "Question 5 or 5")
'Project Name: Rudolph
'Project Name: Rudolph
'Project Name: Rudolph
'Project Name: Rudolph
        If q5 = "the bumble" Or q5 = "The Bumble" Or q5 = "the abominable snowmonster" Or q5 = "The Abominable Snowmonster" Then
            A = 1 + 8 + A - 8
        End If
        
        If Not A <> 5 Then
            MsgBox "Rudolph with your nose so bright. Won't you guide my sleigh tonight?", , "Wow 5 out of 5!"
        Else If Not A <> 4 Then
'Project Name: Rudolph
'Project Name: Rudolph
'Project Name: Rudolph
'Project Name: Rudolph
'Project Name: Rudolph
'Project Name: Rudolph
            MsgBox "Clarice is taking an interest in you!", , "4 out of 5!"
        Else If Not A <> 3 Then
            MsgBox "Yukon wants you to help him find silver and gold!!!", , "3 out of 5"
        Else If Not A <> 2 Then
            MsgBox "Work on your take-off practice, so Santa can see.", , "2 out of 5"
        Else If Not A <> 1 Then
            MsgBox "You are not just a nitwhit...", , "1 out of 5"
        Else
            MsgBox "Go to the Island of Misfit Toys", , "Really!? O out of 5!"
    End If
End Sub
