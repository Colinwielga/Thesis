VERSION 5.00
Begin VB.Form frmHusband 
   BackColor       =   &H00008000&
   Caption         =   "Find your Dream Husband!"
   ClientHeight    =   9885
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   ForeColor       =   &H00008000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   9885
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H0000C000&
      Caption         =   "Continue to Pick a Career"
      BeginProperty Font 
         Name            =   "Mathematica5"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5400
      Width           =   2895
   End
   Begin VB.CommandButton cmdHomepage 
      Caption         =   "Home"
      BeginProperty Font 
         Name            =   "Mathematica5"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12240
      TabIndex        =   4
      Top             =   8400
      Width           =   1695
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Mathematica5"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10440
      TabIndex        =   3
      Top             =   8400
      Width           =   1695
   End
   Begin VB.CommandButton cmdHairColor 
      Caption         =   "Sort  Husband Options By Age and  Hair Color"
      BeginProperty Font 
         Name            =   "Mathematica5"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3000
      TabIndex        =   2
      Top             =   3480
      Width           =   2895
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   5775
      Left            =   6600
      ScaleHeight     =   5715
      ScaleWidth      =   4995
      TabIndex        =   1
      Top             =   1440
      Width           =   5055
   End
   Begin VB.CommandButton cmdSlideShow 
      Caption         =   "View Ideal Candidates"
      BeginProperty Font 
         Name            =   "Mathematica5"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3000
      TabIndex        =   0
      Top             =   1440
      Width           =   2895
   End
   Begin VB.Shape shapeBackground 
      BorderColor     =   &H00004000&
      FillColor       =   &H00C0FFC0&
      FillStyle       =   7  'Diagonal Cross
      Height          =   7335
      Left            =   6120
      Shape           =   2  'Oval
      Top             =   720
      Width           =   6015
   End
End
Attribute VB_Name = "frmHusband"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: The Game Of Life
'Form Name: frmHusband
'Authors: Pam Arndt and Alisa Davis
'Date Written: 3/4/09
'Objective: User views slideshow of several famous men, sorts them by age or hair color, and choses husband to be saved for summary page
Option Explicit
Dim MenList(1 To 6) As Integer, age(1 To 6) As Integer, haircolor(1 To 6) As String, ctr As Integer

Private Sub cmdHairColor_Click()
Dim ColorInput As Integer, AgeInput As Integer
'This button lets a user pick their ideal hair color and age preference for their spouse via an input box

ColorInput = InputBox("Enter 1 if you prefer brunette, enter 2 if you prefer blonde.", "Hair Color")

If ColorInput = 1 Then 'The user inputs number between 1 for brunette
    AgeInput = InputBox("Type a 1 if you're looking for a husband between the ages of 20 and 30, 2 if you're looking for a husband between 30 and 40, or 3 if you prefer a husband over 40 years of age.", "Husband")
        If AgeInput = 1 Then 'User designates brunette preference of 20 to 30 age range
            Spouse = "Zac Efron"
            MsgBox "Your top choice is Zack Efron.", , "Spouse"
            ChildName = Left(UserFirstName, 3) + Right("Zack Efron", 3)
            MsgBox "You will name your first child " & ChildName, , "Child"
        ElseIf AgeInput = 2 Then 'User designates brunette preference of 30 to 40 age range
            Spouse = "Will Smith"
            MsgBox "Your top choice is Will Smith.", , "Spouse"
            ChildName = Left(UserFirstName, 3) + Right("Will Smith", 4)
            MsgBox "You will name your first child " & ChildName, , "Child"
        ElseIf AgeInput = 3 Then 'User input prefers brunette age of 40 and over
            Spouse = "George Clooney"
            MsgBox "Your top choice is George Clooney.", , "Spouse"
            ChildName = Left(UserFirstName, 3) + Right("George Clooney", 5)
            MsgBox "You will name your first child " & ChildName, , "Child"
        Else: MsgBox "Sorry, you entered an invalid option plese enter either a 1,2, or 3.", , "Oopsiedaisy!"
        End If
ElseIf ColorInput = 2 Then 'The user inputs number designating blonde preference
    AgeInput = InputBox("Type a 1 if you're looking for a husband between the ages of 20 and 30, 2 if you're looking for a husband between 30 and 40, or 3 if you prefer a husband over 40 years of age.", "Husband")
        If AgeInput = 1 Then  'User designates preference of blonde 20 to 30 age range
            Spouse = "Chad Michael Murray"
            MsgBox "Your top choice is Chad Michael Murray.", , "Spouse"
            ChildName = Left(UserFirstName, 3) + " " + Mid("Chad Michael Murray", 6, 7)
            MsgBox "You will name your first child " & ChildName, , "Child"
        ElseIf AgeInput = 2 Then 'User designates preference of blonde 30 to 40 age range
            Spouse = "Jude Law"
            MsgBox "Your top choice is Jude Law.", , "Spouse"
            ChildName = Right(UserFirstName, 3) + " " + Left("Jude Law", 4)
            MsgBox "You will name your first child " & ChildName, , "Child"
        ElseIf AgeInput = 3 Then 'User input prefers blonde age of 40 and over
            Spouse = "Brad Pitt"
            MsgBox "Your top choice is Brad Pitt.", , "Spouse"
            ChildName = Left(UserFirstName, 3) + Right("Brad Pitt", 3)
            MsgBox "You will name your first child " & ChildName, , "Child"
        Else: MsgBox "Sorry, you entered an invalid option plese enter either a 1,2, or 3.", , "Oopsiedaisy!"
        End If
Else: MsgBox "Sorry, you entered an invalid option plese enter either a 1 or 2.", , "Oopsiedaisy!"
End If
        
End Sub


Private Sub cmdQuit_Click()
End 'Quit program
End Sub

Private Sub cmdHomepage_Click()
frmBeginning.Show 'Return to beginning page
frmHusband.Hide
End Sub


Private Sub cmdSlideShow_Click()
'this subroutine is used to cycle through the array of filenames and display
'each picture for a short amount of time.

Dim name(1 To 6) As String, age(1 To 6) As Integer, haircolor(1 To 6) As String, PicCTR As Integer, stopper As Integer, t As Double, ctr2 As Double

ctr = 0

Open App.Path & "\Men.txt" For Input As #1

        Do While Not EOF(1)
            ctr = ctr + 1
            Input #1, name(ctr), age(ctr), haircolor(ctr)
        Loop
Close #1

' start slideshow with first picture
PicCTR = 1

stopper = 0 'this is a counter used to stop the slide show after a number of slides
   
'Open the file of picture names and put them in an array called names.

    
'load men picture names into array
Do While (stopper < 6)
    ctr = 0
    'load a picture using the picture name from the array of names
    picResults.Picture = LoadPicture(App.Path & "\" & name(PicCTR))
    
    ' use the Timer function to dispay each picture for about 2 seconds
    'empty loop ... just letting the timer run for 2 seconds
    t = Timer
    Do While (Timer - t) < 2
        ctr2 = ctr2 + 1
        If ctr2 = 1000000 Then
            'picResults.Print t, Timer
            'so you can see what this is doing
            ctr2 = 0
        End If
    Loop
    
    'stopper is used to stop the loop after showing 'stopper' pictures
    stopper = stopper + 1
    
    'the mod function gives the integer remainder of stopper/ctr,
    ' generating 1,2,3,1,2,3,...  in this case where ctr is three
    'whichOne will then always be the position in the array of a valid picture
    PicCTR = (stopper Mod ctr2) + 1
Loop

End Sub

Private Sub cmdNext_Click()
Dim Choice As String
'automatically takes user to career choice forms

Choice = InputBox("Type a 1 if you've completed or planning on completing college, or 2 if you've completed little or no college.", "Career Options")

'An else-if-then statement directs the user to the correct form of their career path.
    If Choice = 1 Then
        frmHusband.Hide
        frmCompletedCollege.Show
    ElseIf Choice = 2 Then
        frmHusband.Hide
        frmNoCollege.Show
    Else: MsgBox "Sorry, you entered an invalid option please enter either a 1 or 2.", , "Error"
End If


End Sub
