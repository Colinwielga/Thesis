VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdLetsplaytrivia_Click()
    Dim qteam As String, ateam As String, qplayers As Integer, aplayers As Integer, qstrikes As Integer, astrikes As Integer, qbonds As String, abonds As String, qsox As String, asox As String, qleagues As String, aleagues As String, qchamps As Single, achamps As Single, qdiamond As Integer, adiamond As Integer, numMatches
    
    ' here are my values
    ateam = "Twins"
    aplayers = 9
    astrikes = 3
    abonds = "yes"
    asox = "Red"
    aleagues = "American"
    achamps = "1991"
    adiamonds = 4
 
    numMatches = 0
    
    ' ask user the questions, keep track of matches
    qteam = InputBox("What is the nickname of Minnesota's MLB team?", "Question 1")
    ' I'm doing a case-insensitve match here - not a requirement
    If LCase(ateam) = LCase(qteam) Then
        numMatches = numMatches + 1
        MsgBox "Correct!", , "Answer!"
    End If
    
    qplayers = InputBox("How many players players play in the field at one time?", "Question 2")
    If LCase(aplayers) = LCase(qplayers) Then
        numMatches = numMatches + 1
        MsgBox "Correct!", , "Answer!"
    End If
    
    qstrikes = InputBox("How many strikes does a batter get?", "Question 3")
    If astrikes = qstrikes Then
        numMatches = numMatches + 1
        MsgBox "Correct!", , "Answer!"
    End If
    
    qbonds = InputBox("Does Barry Bonds use steriods?", "Question 4")
    If abonds = qbonds Then
        numMatches = numMatches + 1
        MsgBox "Correct!", , "Answer!"
    End If
    
    qsox = InputBox("What color Sox play in Boston ?", "Question 5")
    If asox = qsox Then
        numMatches = numMatches + 1
        MsgBox "Correct!", , "Answer!"
    End If
    
    qleagues = InputBox("There are two leagues; the Nationals and what?", "Question 6")
    If aleagues = qleagues Then
        numMatches = numMatches + 1
        MsgBox "Correct!", , "Answer!"
    End If
    
     qchamps = InputBox("The Minnesota Twins were World Champions in 1987 and?", "Question 7")
    If achamps = qchamps Then
        numMatches = numMatches + 1
        MsgBox "Correct!", , "Answer!"
    End If
    
     qdiamonds = InputBox("How many bases are on a baseball diamond?", "Question 8")
    If adiamonds = qdiamonds Then
        numMatches = numMatches + 1
        MsgBox "Correct!", , "Answer!"
    End If
    
    picResults.Cls
    
    Select Case numMatches
        Case 0
            picResults.Print "Are you from Mars!"
        Case 1
            picResults.Print "Baseball is our national pastime .  One right??? You know nothing!!"
        Case 2
            picResults.Print "You need to brush up on your basball!"
        Case 3
            picResults.Print "Not very good, maybe you should try again!"
        Case 4
            picResults.Print "Could have done better!"
        Case 5
            picResults.Print "Half right, if you were in school you would get an F!!!"
        Case 6
            picResults.Print "Pretty good, someone watches baseball"
        Case 7
            picResults.Print "Wow, way to go!"
        Case 8
            picResults.Print "You are a basball GOD!!!"
        Case Else
            picResults.Print "Think we have an error here!"
    End Select
End Sub



Private Sub cmdQuit_Click()
    MsgBox ("Thanks for playing trivia!")
    End
End Sub

