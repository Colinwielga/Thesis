VERSION 5.00
Begin VB.Form frmSongs 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Song Lyrics"
   ClientHeight    =   8505
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10440
   LinkTopic       =   "Form1"
   ScaleHeight     =   8505
   ScaleWidth      =   10440
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start Page"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   18
      Top             =   7440
      Width           =   1335
   End
   Begin VB.CommandButton cmdScorcho 
      Caption         =   """El Scorcho"" Lyrics"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7200
      TabIndex        =   15
      Top             =   5160
      Width           =   3135
   End
   Begin VB.CommandButton cmdKeep 
      Caption         =   """Keep Fishin'"" Lyrics"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7200
      TabIndex        =   14
      Top             =   4440
      Width           =   3135
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test Yourself"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7320
      TabIndex        =   13
      Top             =   6240
      Width           =   2535
   End
   Begin VB.CommandButton cmdPicture 
      Caption         =   "Picture Page"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8760
      TabIndex        =   12
      Top             =   7440
      Width           =   1335
   End
   Begin VB.CommandButton cmdInfo 
      Caption         =   "Info Page"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7200
      TabIndex        =   11
      Top             =   7440
      Width           =   1335
   End
   Begin VB.CommandButton cmdTour 
      Caption         =   "Discography Page"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1800
      TabIndex        =   10
      Top             =   7440
      Width           =   1335
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Current Song"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   600
      TabIndex        =   9
      Top             =   6240
      Width           =   2415
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Exit This Rad Program"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4440
      TabIndex        =   8
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CommandButton cmdBeverly 
      Caption         =   """Beverly Hills"" Lyrics"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   5160
      Width           =   3015
   End
   Begin VB.CommandButton cmdIsland 
      Caption         =   """Island In the Sun"" Lyrics"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   4440
      Width           =   3015
   End
   Begin VB.CommandButton cmdBuddy 
      Caption         =   """Buddy Holly"" Lyrics"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   3720
      Width           =   3015
   End
   Begin VB.CommandButton cmdPork 
      Caption         =   """Pork and Beans"" Lyrics"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7200
      TabIndex        =   4
      Top             =   3720
      Width           =   3135
   End
   Begin VB.PictureBox picResults 
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Left            =   3240
      ScaleHeight     =   6795
      ScaleWidth      =   3795
      TabIndex        =   2
      Top             =   480
      Width           =   3855
   End
   Begin VB.PictureBox Picture2 
      Height          =   6015
      Left            =   -240
      Picture         =   "frmSongs.frx":0000
      ScaleHeight     =   5955
      ScaleWidth      =   6435
      TabIndex        =   1
      Top             =   0
      Width           =   6495
      Begin VB.Label lblEasier 
         Caption         =   "Easier Songs:"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   17
         Top             =   3240
         Width           =   1455
      End
      Begin VB.Label lblSongs 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "SONG LYRICS:"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3960
         TabIndex        =   3
         Top             =   0
         Width           =   2775
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   6015
      Left            =   3120
      Picture         =   "frmSongs.frx":985B
      ScaleHeight     =   5955
      ScaleWidth      =   7515
      TabIndex        =   0
      Top             =   0
      Width           =   7575
      Begin VB.Label lblHarder 
         Caption         =   "Harder songs:"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4920
         TabIndex        =   16
         Top             =   3240
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmSongs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Weezer
'Form Name: frmSongs.frm
'Author: Emily Balamut
'Date Written: 10/30/08
'Objective: This form allows the user to study the lyrics of 6 songs and then test
'themselves on if they remember which titles go with which lyrics.
Option Explicit

Private Sub cmdBeverly_Click()
    picResults.Cls
    picResults.Print "Where I come from isn't all that great"
    picResults.Print "My automobile is a piece of crap"
    picResults.Print "My fashion sense is a little whack"
    picResults.Print "And my friends are just as screwy as me"
    picResults.Print "I didn't go to boarding schools"
    picResults.Print "Preppie girls never looked at me"
    picResults.Print "Why should they?"
    picResults.Print "I ain't nobody"
    picResults.Print "Got nothing in my pocket"
    picResults.Print "Beverly Hills"
    picResults.Print "That's where I want to be"
    picResults.Print "Livin' in Beverly Hills'"
    picResults.Print "Beverly Hills"
    picResults.Print "Rolling like a celebrity"
    picResults.Print "Livin' in Beverly Hills"

End Sub

Private Sub cmdBuddy_Click()
    picResults.Cls
    picResults.Print "What's with these homies dissin' my girl?"
    picResults.Print "Why do they gotta front?"
    picResults.Print "What did we ever do to these guys"
    picResults.Print "That made them so violent?"
    picResults.Print "Woo-hoo, but you know I'm yours."
    picResults.Print "Woo-hoo, and I know you're mine."
    picResults.Print "Woo-hoo, and that's for all of time."
    picResults.Print "Woo-ee-oo, I look just like Buddy Holly."
    picResults.Print "Oh-Oh, and you're Mary Tyler Moore."
    picResults.Print "I don't care what they say about us anyway."
    picResults.Print "I don't care 'bout that."
End Sub

Private Sub cmdClear_Click()
    picResults.Cls
End Sub

Private Sub cmdGarage_Click()

End Sub

Private Sub cmdInfo_Click()
    frmSongs.Hide
    frmInfo.Show
End Sub

Private Sub cmdIsland_Click()
    picResults.Cls
    picResults.Print "When you’re on a holiday"
    picResults.Print "You can’t find the words to say"
    picResults.Print "All the things that come to you"
    picResults.Print "And I wanna feel it too"
    picResults.Print "On an island in the sun"
    picResults.Print "We’ll be playing and having fun"
    picResults.Print "And it makes me feel so fine"
    picResults.Print "I can’t control my brain"
End Sub

Private Sub cmdKeep_Click()
    picResults.Cls
    picResults.Print "You'll never be"
    picResults.Print "A better kind"
    picResults.Print "If you don't leave"
    picResults.Print "The world behind"
    picResults.Print "Waste my days"
    picResults.Print "Drown aways"
    picResults.Print "It's just the thought of you"
    picResults.Print "In love with someone else"
    picResults.Print "It breaks my heart to see you"
    picResults.Print "Hangin' from your shelf"
    picResults.Print "You'll never do"
    picResults.Print "The things you want"
    picResults.Print "If you don't move"
    picResults.Print "And get a job"
    picResults.Print "Waste my days"
    picResults.Print "Drown aways"
End Sub

Private Sub cmdPicture_Click()
    frmSongs.Hide
    frmPictures.Show
End Sub

Private Sub cmdPork_Click()
    picResults.Cls
    picResults.Print "They say"
    picResults.Print "I need some Rogaine"
    picResults.Print "To put in my hair"
    picResults.Print "Work it out at the gym"
    picResults.Print "To fit my underwear"
    picResults.Print "Oakley makes the shades"
    picResults.Print "That transform a tool"
    picResults.Print "You'd hate"
    picResults.Print "For the kids to think"
    picResults.Print "That you lost your cool"
    picResults.Print "I'mma do the things"
    picResults.Print "That I wanna do"
    picResults.Print "I ain't got a thing"
    picResults.Print "To prove to you"
    picResults.Print "I'll eat my candy"
    picResults.Print "With the pork and beans"
    picResults.Print "Excuse my manners"
    picResults.Print "If I make a scene"
    picResults.Print "I ain't gonna wear"
    picResults.Print "The clothes that you like"
    picResults.Print "I'm finally dandy"
    picResults.Print "With the me inside"
    picResults.Print "One look in the mirror"
    picResults.Print "And I'm tickled pink"
    picResults.Print "I don't give a hoot"
    picResults.Print "About what you think"

End Sub

Private Sub cmdQuit_Click()
MsgBox "Thanks for rocking out with Weezer, " & UserName & "! See you later!", , "Exit"
End
End Sub

Private Sub cmdScorcho_Click()
    picResults.Cls
    picResults.Print "Goddamn you half-Japanese girls"
    picResults.Print "Do it to me every time"
    picResults.Print "Oh, the redhead said you shred the cello "
    picResults.Print "And I'm jello, baby"
    picResults.Print "You won't talk, won't look, won't think of me"
    picResults.Print "I'm the epitome of Public Enemy"
    picResults.Print "Why you wanna go and do me like that?"
    picResults.Print "Come down on the street and dance with me"
    picResults.Print "I'm a lot like you so please"
    picResults.Print "Hello, I'm here, I'm waiting"
    picResults.Print "I think I'd be good for you"
    picResults.Print "And you'd be good for me"
End Sub

Private Sub cmdStart_Click()
    frmSongs.Hide
    frmBeginning.Show
End Sub

Private Sub cmdTest_Click()
    Dim Test As String, Test2 As String, Test3 As String, Test4 As String, Test5 As String, Test6 As String
    Dim CTR As Integer, N As Integer
    Dim AnswerNumber(1 To 6) As Integer
    Dim LyricAnswer(1 To 6) As String
    
    picResults.Cls
    cmdBuddy.Enabled = False
    cmdIsland.Enabled = False
    cmdBeverly.Enabled = False
    cmdPork.Enabled = False
    cmdKeep.Enabled = False
    cmdScorcho.Enabled = False
    CTR = 0
    Open App.Path & "\LyricAnswers.txt" For Input As #1
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, AnswerNumber(CTR), LyricAnswer(CTR)
    Loop
    Close #1
    
    N = 0

    Test = InputBox("Now for a test of your memory. Which song goes: 'I look just like [Song Title here], Uh-oh and you're Mary Tyler, I don't care what they say about us anyway, I don't care about that.'", "Question 2")
    If LCase(Test) = LCase("Buddy Holly") Then
        N = N + 1
    End If
    Test2 = InputBox("Which song goes: 'I'm a lot like you so please, Hello, I'm here, I'm waiting. I think I'd be good for you, And you'd be good for me'?", "Question 2")
    If LCase(Test2) = LCase("El Scorcho") Then
        N = N + 1
    End If
    Test3 = InputBox("Which song goes: 'You'll never be a better kind, if you don't leave the world behind. Waste my days, drown aways. It's just the thought of you in love with someone else. It breaks my heart to see you hangin' from your shelf.'", "Question 3")
    If LCase(Test3) = LCase("Keep Fishin'") Then
        N = N + 1
    End If
    Test4 = InputBox("Which song goes: 'When you're on a holiday, you can't find the words to say, all the things that come to you, and I wanna feel it too. On an [Song Title Here], we'll be playing and having fun and it makes me feel so fine, I can't control my brain.'", "Question 4")
    If LCase(Test4) = LCase("Island In the Sun") Then
        N = N + 1
    End If
    Test5 = InputBox("Which song goes: 'I'mma do the things that I wanna do, I ain't got a thing to prove to you, I'll eat my candy with the [Song Title Here], excuse my manners if I make a scene.'", "Question 5")
    If LCase(Test5) = LCase("Pork and Beans") Then
        N = N + 1
    End If
    Test6 = InputBox("Which song goes: '[Song Title Here], That's where I want to be, livin' in [Song Title Here]. [Song Title Here], rolling like a celebrity, livin' in [Song Title Here]'", "Question 6")
    If LCase(Test6) = LCase("Beverly Hills") Then
        N = N + 1
    End If
    
    Select Case N
        Case Is = 6
            MsgBox "Congratulations! You really know your stuff! You got them ALL correct!", , "6 Correct"
        Case Is = 5
            MsgBox "Good job! 5 out of 6 is really good!", , "5 Correct"
        Case Is = 4
            MsgBox "Way to go! Still more than half correct!", , "4 Correct"
        Case Is = 3
            MsgBox "You got half of the song titles right! Not too bad!", , "3 Correct"
        Case Is = 2
            MsgBox "Ooooo, It looks like you did not do so well. Try looking at the lyrics again and try again!", , "2 Correct"
        Case Is = 1
            MsgBox "You really need to look these lyrics over! Try again!", , "1 Correct"
        Case Else
            MsgBox "You got none right! What would Rivers say?", , "None Correct!"
    End Select
    
    cmdBuddy.Enabled = True
    cmdIsland.Enabled = True
    cmdBeverly.Enabled = True
    cmdPork.Enabled = True
    cmdKeep.Enabled = True
    cmdScorcho.Enabled = True
End Sub

Private Sub cmdTour_Click()
    frmSongs.Hide
    frmDisco.Show
End Sub
