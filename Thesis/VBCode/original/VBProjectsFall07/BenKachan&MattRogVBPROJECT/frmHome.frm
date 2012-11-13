VERSION 5.00
Begin VB.Form frmHome 
   Caption         =   "Sports Headquarters"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   Picture         =   "frmHome.frx":0000
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   1215
      Left            =   3960
      TabIndex        =   6
      Top             =   7080
      Width           =   3375
   End
   Begin VB.CommandButton cmdSoccer 
      Caption         =   "Soccer Shootout"
      Height          =   975
      Left            =   4560
      TabIndex        =   5
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CommandButton cmdGolf 
      Caption         =   "Long Drive Golf"
      Height          =   975
      Left            =   7920
      TabIndex        =   4
      Top             =   6000
      Width           =   1815
   End
   Begin VB.CommandButton cmdFootball 
      Caption         =   "Football Quiz"
      Height          =   975
      Left            =   4560
      TabIndex        =   3
      Top             =   5040
      Width           =   1815
   End
   Begin VB.CommandButton cmdBball 
      Caption         =   "Basketball Shoot Around"
      Height          =   975
      Left            =   720
      TabIndex        =   2
      Top             =   3120
      Width           =   1815
   End
   Begin VB.CommandButton cmdHockey 
      Caption         =   "Hockey Store"
      Height          =   975
      Left            =   1440
      TabIndex        =   1
      Top             =   6360
      Width           =   1815
   End
   Begin VB.CommandButton cmdBaseball 
      Caption         =   "Baseball Statistics"
      Height          =   975
      Left            =   8280
      TabIndex        =   0
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000014&
      Caption         =   "Welcome to your one stop fun shop for sports entertainment and knowledge "
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   7
      Top             =   360
      Width           =   10815
   End
End
Attribute VB_Name = "frmHome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'                   Sports For Dummies
'                          By
'                   Matt Rog and Ben Kachan
'               CSCI 130 Section 1-3-5 2:40-3:50
'                        Fall 2007

'Project Discription: The problem today is that there aren't enough interactive sports sites that make learning sports and science fun.
'We thought that this program would not only be interesting to play but it would challenge us to our maximum potential for as much as we've learned.
'Also, we felt that if the user had any type of interest in sports that he would be able to relate to some aspect of it easily.
'We expect this project to be one of the premier projects in terms of interaction with the user and the things that it allows the user to do.
'Therefore, we were aiming at a grade of "A".  This program allows the user a variety of options to interact, which are defined within the particular section.
'Project Features:An example of a nested If statement in our program is in the frmGame1: If (sum = 2) Then
'       sixth = InputBox("its all tied up and its up to your sixth shooter Michael Owen")
'       If (sixth = 1 Or sixth = 5 Or sixth = 4) Then
'           picInstruct.Print "He did it. Congratulations you won the Collegeville Open Shootout Championship"
'           sum = sum + 1
'           picInstruct.Print "You managed to score: " & sum; " goals "
'       End If
'       If (sixth = 2 Or sixth = 3) Then
'            picInstruct.Print "You gave it your all but Argentina was the better team today, you lose."
'           picInstruct.Print "You managed to score: " & sum; " goals "
'       End If
'   If (sixth > 5 Or sixth < 1) Then
'            picInstruct.Print "You gave it your all but Argentina was the better team today, you lose."
'           picInstruct.Print "You managed to score: " & sum; " goals "
'    End If
'    End If
'Another example is when we used the bubble sort method and a For-Next loop:
'For pass = 1 To CTR - 1
'        For Comp = 1 To (CTR - pass)
'            If (Prices(Comp) < Prices(Comp + 1)) Then
'                tempItem = Items(Comp)
'                Items(Comp) = Items(Comp + 1)
'                Items(Comp + 1) = tempItem
'                tempPrice = Prices(Comp)
'                Prices(Comp) = Prices(Comp + 1)
'                Prices(Comp + 1) = tempPrice
'            End If
'        Next Comp
'    Next pass
'Next, for the backrounds for all of our particular sports, we found images on the internet, cropped the pictures to fit within our design, then imported them using the "picture property" on each form.
'Other than that we used just the basic programming methods learned in class     If then statement, Select Case statements, File input, Arrays and processing each element in some way(maybe a calculation)
'Arrays and searching, Arrays and sorting, Multiple Forms, Extensive use of colors, Extensive use of pictures\Images, Input from text boxes, Input from input boxes
'Loops, (Do While, Do Until, or For Next), String functions (e.g. LEFT, RIGHT, MID, INSTR, etc …), Math functions (e.g. LOG, SQR, INT, etc …), Message boxes, and Module Variables
'Project Experience: Overall I would say this was a very good experience.  There were some problems that came out throughout the project such as attempting to do something that we hadn't learned in class
'But then in the end not being able to work it or tweak it quite right so having to rely on a different method.  It was difficult to get started because you basically have nothing but an idea and a form
'But once things started to get going and the ideas began to make sense and work through programs then creating everythign actually became fun and a sense of accomplishment.  For all the basic applications
'applications we have learned in class, putting them altogether to form something such as this is really challenging and time consuming.  It's not possible to sit down in one day and put together something
'like this.  It takes teamwork, knowledge of the concepts, and through all this, trying to keep creativity flowing so that the user stays entertained throughout.  We felt at the end that with another class
'such as this one that our program could be taken to even higher levels of professionalism and complexity.
Option Explicit

Private Sub cmdBaseball_Click()
    'for this subroutine we decided to just load the information into an array as soon as they clicked on the Baseball Statistics command
    'Also it will take you to the form for baseball
    Dim CTR As Integer
    frmBaseball.Show
    frmHome.Hide
    Open App.Path & "\PitchWins.txt" For Input As #1
    CTR = 0
    Do While Not EOF(1)
        CTR = CTR + 1
        Input #1, PitcherNames(CTR), PitcherWins(CTR)
    Loop
    Close #1
    Open App.Path & "\Strikeouts.txt" For Input As #2
    CTR = 0
    Do While Not EOF(2)
        CTR = CTR + 1
        Input #2, KNames(CTR), Strikeouts(CTR)
    Loop
    Close #2
    Open App.Path & "\HomeRuns.txt" For Input As #3
    CTR = 0
    Do While Not EOF(3)
        CTR = CTR + 1
        Input #3, HRNames(CTR), HR(CTR)
    Loop
    Close #3
     Open App.Path & "\BattingAvg.txt" For Input As #4
    CTR = 0
    Do While Not EOF(4)
        CTR = CTR + 1
        Input #4, AvgNames(CTR), Average(CTR)
    Loop
    Close #4
    Open App.Path & "\GamesPlayed.txt" For Input As #5
    CTR = 0
    Do While Not EOF(5)
        CTR = CTR + 1
        Input #5, BaseballNames(CTR), GamesPlayed(CTR), Rank(CTR)
    Loop
    Close #5
    Open App.Path & "\ERA.txt" For Input As #6
    CTR = 0
    Do While Not EOF(6)
        CTR = CTR + 1
        Input #6, ERANames(CTR), ERA(CTR)
    Loop
    Close #6
End Sub

Private Sub cmdBball_Click()
    'Takes you to basketball form
    frmBaskets.Show
    frmHome.Hide
End Sub

Private Sub cmdFootball_Click()
    'Takes you to football form
    frmFootball.Show
    frmHome.Hide
End Sub

Private Sub cmdGolf_Click()
    'Takes you to golf form
    frmGolf.Show
    frmHome.Hide
End Sub

Private Sub cmdHockey_Click()
    'Takes you to hockey form
    frmHockey.Show
    frmHome.Hide
End Sub

Private Sub cmdQuit_Click()
    'End the program
    End
End Sub

Private Sub cmdSoccer_Click()
    'Takes you to soccer form
    frmSoccer.Show
    frmHome.Hide
End Sub

