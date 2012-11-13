VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   10785
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form4"
   Picture         =   "quiz form.frx":0000
   ScaleHeight     =   10785
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdmainpage 
      Caption         =   "Go back to Main Page"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   960
      TabIndex        =   7
      Top             =   8280
      Width           =   1815
   End
   Begin VB.CheckBox chk_4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Answer 4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6000
      TabIndex        =   4
      Top             =   7080
      Width           =   3975
   End
   Begin VB.CheckBox chk_3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Answer 3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6000
      TabIndex        =   3
      Top             =   6360
      Width           =   3975
   End
   Begin VB.CheckBox chk_2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Answer 2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6000
      TabIndex        =   2
      Top             =   5640
      Width           =   3975
   End
   Begin VB.CheckBox chk_1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Answer 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6000
      TabIndex        =   1
      Top             =   4920
      Width           =   3975
   End
   Begin VB.CommandButton btn_start 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Take Quiz Now!!!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   6000
      TabIndex        =   0
      Top             =   1920
      Width           =   4455
   End
   Begin VB.Label lbl_results 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   4680
      TabIndex        =   6
      Top             =   4560
      Width           =   6735
   End
   Begin VB.Label lbl_question 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   4680
      TabIndex        =   5
      Top             =   3360
      Width           =   7095
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project name: Gilligan's Island
'Form name:  quiz form
'Author:  Emily Olson
'Date written:  March 30, 2008
'Form Objective: have user take a quiz to determine which "Gilligan's Island character they are most like


'declare variables
Dim QuestionNumber As Integer
Dim Professor As Integer, MaryAnn As Integer, Gilligan As Integer, Howells As Integer, Skipper As Integer, MrsHowell As Integer, MrHowell As Integer, Ginger As Integer
Private Sub cmdmainpage_Click()
'load main page
    Form1.Show
    Form4.Hide
End Sub

Public Sub Form_Load()
'make question and answers not visible until quiz is started

    lbl_question.Visible = False
    chk_1.Visible = False
    chk_2.Visible = False
    chk_3.Visible = False
    chk_4.Visible = False
    chk_1 = False
    chk_2 = False
    chk_3 = False
    chk_4 = False
    lbl_results.Visible = False

'start answer counters so the results can be calculated
'start question number counters so the next question will be displayed

    QuestionNumber = 1
    Gilligan = 0
    Professor = 0
    Skipper = 0
    MaryAnn = 0
    Howells = 0
    Ginger = 0

End Sub
Public Sub btn_start_Click()

'show first question and possible answers

    If QuestionNumber = 1 Then

        btn_start.Caption = "Next Question"
        lbl_question.Visible = True
        chk_1.Visible = True
        chk_2.Visible = True
        chk_3.Visible = True
        chk_4.Visible = True
    
        lbl_question = "If you could choose one luxury item to take to your desert island, what would it be?"
        chk_1.Caption = "battery powered radio"
        chk_2.Caption = "video camera"
        chk_3.Caption = "your manicurist"
        chk_4.Caption = "idiot's guide to boat repair"
        
'question counter so will either go on to next question or display message to try again
        QuestionNumber = QuestionNumber + 1
'exit sub so it will loop to next question
        Exit Sub
    End If

    If QuestionNumber = 2 Then
'answer counters so results can be calculated from prior question
        If chk_1 = 1 Then
            Professor = Professor + 1
            Skipper = Skipper + 1
        End If
        If chk_2 = 1 Then
            Howells = Howells + 1
        End If
        If chk_3 = 1 Then
            Ginger = Ginger + 1
        End If
        If chk_4 = 1 Then
            Gilligan = Gilligan + 1
        End If
'clear checkmarks from prior question
        chk_1 = False
        chk_2 = False
        chk_3 = False
        chk_4 = False

        lbl_question = "If you were rescued by a foreign fishing vessel, and had to work for a ride back to civilization, what would your job be?"
        chk_1.Caption = "Slop Boy"
        chk_2.Caption = "Galley cook"
        chk_3.Caption = "Science officer"
        chk_4.Caption = "I don't work"

        QuestionNumber = QuestionNumber + 1
        Exit Sub
    End If


    If QuestionNumber = 3 Then
        If chk_1 = 1 Then Gilligan = Gilligan + 1
        If chk_2 = 1 Then MaryAnn = MaryAnn + 1
        If chk_3 = 1 Then Professor = Professor + 1
        If chk_4 = 1 Then
            Howells = Howells + 1
            Ginger = Ginger + 1
        End If

        chk_1 = False
        chk_2 = False
        chk_3 = False
        chk_4 = False

        lbl_question = "If you could take a single picture with you to your tropical isle to keep you from getting lonely, what would it be?"
        chk_1.Caption = "A picture of your pet pig, Elvis"
        chk_2.Caption = "A picture of yourself, digitally enhanced"
        chk_3.Caption = "A picture of your new Rolls Royce"
        chk_4.Caption = "A picture of your mom"
    
        QuestionNumber = QuestionNumber + 1
        Exit Sub
    End If

    If QuestionNumber = 4 Then

        If chk_1 = 1 Then Gilligan = Gilligan + 1
        If chk_2 = 1 Then Ginger = Ginger + 1
        If chk_3 = 1 Then Howells = Howells + 1
        If chk_4 = 1 Then MaryAnn = MaryAnn + 1


        chk_1 = False
        chk_2 = False
        chk_3 = False
        chk_4 = False

        lbl_question = "If you could take a vacation to anywhere in the world, where would you go?"
        chk_1.Caption = "Disneyland"
        chk_2.Caption = "An ocean cruise"
        chk_3.Caption = "Paris, just like every year"
        chk_4.Caption = "Space Camp"
    
        QuestionNumber = QuestionNumber + 1
        Exit Sub
    End If

    If QuestionNumber = 5 Then

        If chk_1 = 1 Then Gilligan = Gilligan + 1
        If chk_2 = 1 Then Skipper = Skipper + 1
        If chk_3 = 1 Then Howells = Howells + 1
        If chk_4 = 1 Then Professor = Professor + 1

        chk_1 = False
        chk_2 = False
        chk_3 = False
        chk_4 = False

        lbl_question = "What do you think is your best attribute?"
        chk_1.Caption = "My brains"
        chk_2.Caption = "My brawn"
        chk_3.Caption = "My beauty"
        chk_4.Caption = "My bucks"

        QuestionNumber = QuestionNumber + 1
        Exit Sub
    End If

    If QuestionNumber = 6 Then

        If chk_1 = 1 Then Professor = Professor + 1
        If chk_2 = 1 Then MaryAnn = MaryAnn + 1
        If chk_3 = 1 Then Ginger = Ginger + 1
        If chk_4 = 1 Then Howells = Howells + 1

        chk_1 = False
        chk_2 = False
        chk_3 = False
        chk_4 = False

        lbl_question = "Which of the following celebs would you most like to hang out with?"
        chk_1.Caption = "The Olsen Twins"
        chk_2.Caption = "Donald Trump"
        chk_3.Caption = "Shania Twain"
        chk_4.Caption = "Ivana Trump"

        QuestionNumber = QuestionNumber + 1
        Exit Sub
    End If

    If QuestionNumber = 7 Then

        If chk_1 = 1 Then Gilligan = Gilligan + 1
        If chk_2 = 1 Then Howells = Howells + 1
        If chk_3 = 1 Then MaryAnn = MaryAnn + 1
        If chk_4 = 1 Then Ginger = Ginger + 1

        chk_1 = False
        chk_2 = False
        chk_3 = False
        chk_4 = False

        lbl_question = "If you could only eat one of the following foods for the rest of your life, which would it be?"
        chk_1.Caption = "Macaroni and Cheese"
        chk_2.Caption = "Eggs and Bacon"
        chk_3.Caption = "Fish and Chips"
        chk_4.Caption = "Champagne and Caviar "

        QuestionNumber = QuestionNumber + 1
        Exit Sub
    End If

    If QuestionNumber = 8 Then

        If chk_1 = 1 Then Gilligan = Gilligan + 1
        If chk_2 = 1 Then MaryAnn = MaryAnn + 1
        If chk_3 = 1 Then Skipper = Skipper + 1
        If chk_4 = 1 Then
            Howells = Howells + 1
            Ginger = Ginger + 1
        End If

        chk_1 = False
        chk_2 = False
        chk_3 = False
        chk_4 = False

        lbl_question = "If we forgive you for your Sloth, Lust, and Anger, which of the remaining Seven Deadly Sins are you most guilty of?"
        chk_1.Caption = "Greed"
        chk_2.Caption = "Gluttony"
        chk_3.Caption = "Pride"
        chk_4.Caption = "Envy"

        QuestionNumber = QuestionNumber + 1
        Exit Sub
    End If

    If QuestionNumber = 9 Then

        If chk_1 = 1 Then Howells = Howells + 1
        If chk_2 = 1 Then Skipper = Skipper + 1
        If chk_3 = 1 Then Ginger = Ginger + 1
        If chk_4 = 1 Then Gilligan = Gilligan + 1

        chk_1 = False
        chk_2 = False
        chk_3 = False
        chk_4 = False

        lbl_question = "What type of music do you like?"
        chk_1.Caption = "Rock and Roll"
        chk_2.Caption = "Country"
        chk_3.Caption = "Classical"
        chk_4.Caption = "Drinking Songs"

        QuestionNumber = QuestionNumber + 1
        Exit Sub
    End If

    If QuestionNumber = 10 Then

        If chk_1 = 1 Then Gilligan = Gilligan + 1
        If chk_2 = 1 Then MaryAnn = MaryAnn + 2
        If chk_3 = 1 Then
            Howells = Howells + 1
            Professor = Professor + 1
        End If
        If chk_4 = 1 Then
            Ginger = Ginger + 1
            Skipper = Skipper + 1
        End If

        chk_1 = False
        chk_2 = False
        chk_3 = False
        chk_4 = False

        lbl_question = "What political party do you consider yourself a member of?"
        chk_1.Caption = "Republican"
        chk_2.Caption = "Democrat"
        chk_3.Caption = "Green Party"
        chk_4.Caption = "Party at my place!"

        QuestionNumber = QuestionNumber + 1
        Exit Sub
    End If

    If QuestionNumber = 11 Then

        If chk_1 = 1 Then Howells = Howells + 1
        If chk_2 = 1 Then Skipper = Skipper + 1
        If chk_3 = 1 Then MaryAnn = MaryAnn + 1
        If chk_4 = 1 Then Ginger = Ginger + 1

        chk_1 = False
        chk_2 = False
        chk_3 = False
        chk_4 = False

        lbl_question = "What's your favorite animal?"
        chk_1.Caption = "Dog"
        chk_2.Caption = "Cat"
        chk_3.Caption = "Horse"
        chk_4.Caption = "Helper Monkey"

        QuestionNumber = QuestionNumber + 1
        Exit Sub
    End If

    If QuestionNumber = 12 Then

        If chk_1 = 1 Then MaryAnn = MaryAnn + 1
        If chk_2 = 1 Then Ginger = Ginger + 1
        If chk_3 = 1 Then Howells = Howells + 1
        If chk_4 = 1 Then Gilligan = Gilligan + 1

        chk_1 = False
        chk_2 = False
        chk_3 = False
        chk_4 = False

        lbl_question = "Do you like children?"
        chk_1.Caption = "Yes, I want to have a bunch of them!"
        chk_2.Caption = "No, they're too messy."
        chk_3.Caption = "Sure, as long as the nanny's with them."
        chk_4.Caption = "I like 'All My Children,' does that count?"

        QuestionNumber = QuestionNumber + 1
        Exit Sub
    End If

    If QuestionNumber = 13 Then

        If chk_1 = 1 Then
            Gilligan = Gilligan + 1
            MaryAnn = MaryAnn + 1
        End If
        If chk_2 = 1 Then
            Professor = Professor + 1
            Skipper = Skipper + 1
        End If
        If chk_3 = 1 Then Howells = Howells + 1
        If chk_4 = 1 Then Ginger = Ginger + 1

        chk_1 = False
        chk_2 = False
        chk_3 = False
        chk_4 = False

        lbl_question = "Not that you'd have any on your desert island, but how do you like your eggs?"
        chk_1.Caption = "Over Easy"
        chk_2.Caption = "Poached"
        chk_3.Caption = "Hard boiled"
        chk_4.Caption = "Scrambled "

        QuestionNumber = QuestionNumber + 1
        Exit Sub
    End If

    If QuestionNumber = 14 Then

        If chk_1 = 1 Then
            Ginger = Ginger + 1
            MaryAnn = MaryAnn + 1
        End If
        If chk_2 = 1 Then Howells = Howells + 1
        If chk_3 = 1 Then
            Professor = Professor + 1
            Skipper = Skipper + 1
        End If
        If chk_4 = 1 Then Gilligan = Gilligan + 1

        chk_1 = False
        chk_2 = False
        chk_3 = False
        chk_4 = False

        lbl_question = "What do you live for?"
        chk_1.Caption = "Fame"
        chk_2.Caption = "Fortune"
        chk_3.Caption = "Fun"
        chk_4.Caption = "Family "

        QuestionNumber = QuestionNumber + 1
        Exit Sub
    End If

    If QuestionNumber = 15 Then

        If chk_1 = 1 Then Ginger = Ginger + 1
        If chk_2 = 1 Then Howells = Howells + 1
        If chk_3 = 1 Then Skipper = Skipper + 1
        If chk_4 = 1 Then MaryAnn = MaryAnn + 1

        chk_1 = False
        chk_2 = False
        chk_3 = False
        chk_4 = False

        lbl_question = "Are you in a relationship?"
        chk_1.Caption = "Yes, it rocks!"
        chk_2.Caption = "Yes, save me!"
        chk_3.Caption = "No, I like being alone."
        chk_4.Caption = "No, and I cry myself to sleep every night. "

        QuestionNumber = QuestionNumber + 1
        Exit Sub
    End If

    If QuestionNumber = 16 Then

        If chk_1 = 1 Then Howells = Howells + 1
        If chk_2 = 1 Then
            MaryAnn = MaryAnn + 1
            Skipper = Skipper + 1
            Gilligan = Gilligan + 1
        End If
        If chk_3 = 1 Then Professor = Professor + 1
        If chk_4 = 1 Then Ginger = Ginger + 1

        chk_1 = False
        chk_2 = False
        chk_3 = False
        chk_4 = False

        lbl_question = "What do you look for in a mate?"
        chk_1.Caption = "Good looks"
        chk_2.Caption = "Intelligence"
        chk_3.Caption = "A sense of humor"
        chk_4.Caption = "Money, baby"

        QuestionNumber = QuestionNumber + 1
        Exit Sub
    End If

    If QuestionNumber = 17 Then

        If chk_1 = 1 Then Ginger = Ginger + 1
        If chk_2 = 1 Then Professor = Professor + 1
        If chk_3 = 1 Then Gilligan = Gilligan + 1
        If chk_4 = 1 Then Howells = Howells + 1

        chk_1 = False
        chk_2 = False
        chk_3 = False
        chk_4 = False

        lbl_question = "What did you like most about school?"
        chk_1.Caption = "Social life"
        chk_2.Caption = "Sports"
        chk_3.Caption = "Cafeteria food"
        chk_4.Caption = "Education"

        QuestionNumber = QuestionNumber + 1
        Exit Sub
    End If

    If QuestionNumber = 18 Then

        If chk_1 = 1 Then Ginger = Ginger + 1
        If chk_2 = 1 Then Howells = Howells + 1
        If chk_3 = 1 Then Skipper = Skipper + 1
        If chk_4 = 1 Then Professor = Professor + 1

        chk_1 = False
        chk_2 = False
        chk_3 = False
        chk_4 = False

        lbl_question = "What type of TV show do you watch most?"
        chk_1.Caption = "Reality TV"
        chk_2.Caption = "News"
        chk_3.Caption = "Cartoons"
        chk_4.Caption = "Comedies"

        QuestionNumber = QuestionNumber + 1
        Exit Sub
    End If

    If QuestionNumber = 19 Then

        If chk_1 = 1 Then Ginger = Ginger + 1
        If chk_2 = 1 Then Howells = Howells + 1
        If chk_3 = 1 Then Gilligan = Gilligan + 1
        If chk_4 = 1 Then
            Skipper = Skipper + 1
            MaryAnn = MaryAnn + 1
        End If

        chk_1 = False
        chk_2 = False
        chk_3 = False
        chk_4 = False

        lbl_question = "If you had a time machine, what historical figure would you visit?"
        chk_1.Caption = "Galileo"
        chk_2.Caption = "Howard Hughes"
        chk_3.Caption = "Marilyn Monroe"
        chk_4.Caption = "John Deere "

        QuestionNumber = QuestionNumber + 1
        Exit Sub
    End If

    If QuestionNumber = 20 Then

        If chk_1 = 1 Then Professor = Professor + 1
        If chk_2 = 1 Then Howells = Howells + 1
        If chk_3 = 1 Then Ginger = Ginger + 1
        If chk_4 = 1 Then MaryAnn = MaryAnn + 1

        chk_1 = False
        chk_2 = False
        chk_3 = False
        chk_4 = False

        lbl_question = "Are you male or female"
        chk_1.Caption = "Male"
        chk_2.Caption = "Female"
        chk_3.Visible = False
        chk_4.Visible = False

        QuestionNumber = QuestionNumber + 1
        Exit Sub
    End If
'to determine if the user is mrs or mr howell
    If chk_1 = 1 Then
        MrHowell = Howells
    ElseIf chk_2 = 1 Then
        MrsHowell = Howells
    End If

'make results box visible
    lbl_results.Visible = True
    lbl_question.Visible = False
    chk_1.Visible = False
    chk_2.Visible = False
    chk_3.Visible = False
    chk_4.Visible = False


    Dim x As Integer
    Dim sorted As Boolean
    Dim temp As Integer
    Dim emily(0 To 6) As Integer
        emily(0) = Gilligan
        emily(1) = Skipper
        emily(2) = Ginger
        emily(3) = Professor
        emily(4) = MaryAnn
        emily(5) = MrHowell
        emily(6) = MrsHowell

'bubble sort the results from lowest to highest
    sorted = False
        Do While Not sorted
        sorted = True
    For x = 0 To UBound(emily) - 1
        If emily(x) > emily(x + 1) Then
            temp = emily(x + 1)
            emily(x + 1) = emily(x)
            emily(x) = temp
            sorted = False
        End If
    Next x
    Loop

'show results according to the highest number (which is also in position 6)
    Select Case emily(6)
        Case Gilligan
            lbl_results = "Well, " & user & ", it seems you are the most like Gilligan!  Ya goof!"
        Case Skipper
            lbl_results = "Well, " & user & ", it seems you are the most like Skipper!"
        Case Ginger
            lbl_results = "Well, " & user & ", it seems you are the most like Ginger! Can I have your autograph?"
        Case Professor
            lbl_results = "Well, " & user & ", it seems you are the most like Professor! Smarty pants!"
        Case MaryAnn
            lbl_results = "Well, " & user & ", it seems you are the most like Mary Ann! Ain't you a sweetie!"
        Case MrHowell
            lbl_results = "Well, " & user & ", it seems you are the most like Mr. Howell! Share the wealth!"
        Case MrsHowell
            lbl_results = "Well, " & user & ", it seems you are the most like Mrs. Howell! Share the wealth!"
        Case Else
            lbl_results = "You aren't anyone!"
    End Select


End Sub










