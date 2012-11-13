VERSION 5.00
Begin VB.Form frmWife 
   BackColor       =   &H00400040&
   Caption         =   "Find Your Dream Wife!"
   ClientHeight    =   10410
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   10410
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Continue to Pick a Career"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   360
      TabIndex        =   6
      Top             =   6240
      Width           =   2655
   End
   Begin VB.PictureBox picA 
      BorderStyle     =   0  'None
      Height          =   9735
      Left            =   9000
      Picture         =   "frmWife.frx":0000
      ScaleHeight     =   9735
      ScaleWidth      =   6135
      TabIndex        =   5
      Top             =   360
      Width           =   6135
   End
   Begin VB.CommandButton cmdHomepage 
      Caption         =   "Home"
      BeginProperty Font 
         Name            =   "Mathematica5"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   7320
      TabIndex        =   4
      Top             =   9600
      Width           =   1095
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
      Height          =   540
      Left            =   6000
      TabIndex        =   3
      Top             =   9600
      Width           =   1095
   End
   Begin VB.CommandButton cmdHairColor 
      Caption         =   "Sort  Wife Options By Age and  Hair Color"
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
      Left            =   240
      TabIndex        =   2
      Top             =   3840
      Width           =   2775
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00000000&
      Height          =   5775
      Left            =   3360
      ScaleHeight     =   5715
      ScaleWidth      =   5475
      TabIndex        =   1
      Top             =   1800
      Width           =   5535
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
      Left            =   240
      TabIndex        =   0
      Top             =   2280
      Width           =   2775
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FF80FF&
      FillStyle       =   5  'Downward Diagonal
      Height          =   6615
      Left            =   3120
      Shape           =   2  'Oval
      Top             =   1320
      Width           =   6015
   End
End
Attribute VB_Name = "frmWife"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: The Game Of Life
'Form Name: frmWife
'Authors: Pam Arndt and Alisa Davis
'Date Written: 3/4/09
'Objective: User views slideshow of several famous women, and can sort them by hair color and age.  If user picks a wife via a textbox, their
'choice is saved for the summary page
Option Explicit

Private Sub cmdEnter_Click()
Dim Choice As String
'automatically takes user to career choice forms

Choice = InputBox("Type a 1 if you've completed or planning on completing college, or 2 if you've completed little or no college.")

'An else-if-then statement directs the user to the correct form of their career path.
    If Choice = 1 Then
        frmWife.Hide
        frmCompletedCollege.Show
    ElseIf Choice = 2 Then
        frmWife.Hide
        frmNoCollege.Show
    Else: MsgBox ("Sorry, you entered an invalid option plese enter either a 1 or 2.")
End If

End Sub

Private Sub cmdHairColor_Click()
Dim ColorInput As Integer, AgeInput As Integer
'This button lets a user pick their ideal hair color and age preference for their spouse via an input box

ColorInput = InputBox("Enter 1 if you prefer brunette, enter 2 if you prefer blonde.", "Hair Color")

If ColorInput = 1 Then 'The user inputs number between 1 for brunette
    AgeInput = InputBox("Type a 1 if you're looking for a husband between the ages of 20 and 30, 2 if you're looking for a husband between 30 and 40, or 3 if you prefer a husband over 40 years of age.", "Husband")
        If AgeInput = 1 Then 'User designates preference of 20 to 30 age range
            Spouse = "Jessica Alba"
            ChildName = Left(UserFirstName, 3) + Mid("Jessica Alba", 5, 3)
            MsgBox "Your top choice is Jessica Alba.", , "Spouse" 'brunette aged 20-30
            MsgBox "You will name your first child " & ChildName, , "Child"
        ElseIf AgeInput = 2 Then 'User designates preference of 30 to 40 age range
            Spouse = "Angelina Jolie"
            ChildName = Left(UserFirstName, 3) + " " + Right("Angelina Jolie", 5)
            MsgBox "Your top choice is Angelina Jolie.", , "Spouse" 'brunette aged 30-40
            MsgBox "You will name your first child " & ChildName, , "Child"
        ElseIf AgeInput = 3 Then 'User input prefers age of 40 and over
            Spouse = "Halle Berry"
            ChildName = Left("Halle Berry", 3) + Right(UserFirstName, 3)
            MsgBox "Your top choice is Halle Berry.", , "Spouse" 'brunette aged 40+
            MsgBox "You will name your first child " & ChildName, , "Child"
        Else: MsgBox "Sorry, you entered an invalid option plese enter either a 1,2, or 3.", , "Oopsiedaisy!"
        End If
ElseIf ColorInput = 2 Then 'The user inputs number designating blonde preference
    AgeInput = InputBox("Type a 1 if you're looking for a husband between the ages of 20 and 30, 2 if you're looking for a husband between 30 and 40, or 3 if you prefer a husband over 40 years of age.", "Husband")
        If AgeInput = 1 Then  'User designates preference of 20 to 30 age range
            Spouse = "Carrie Underwood"
            ChildName = Left(UserFirstName, 3) + Mid("Carrie Underwood", 2, 4) + Right(UserFirstName, 2)
            MsgBox "Your top choice is Carrie Underwood.", , "Spouse"
            MsgBox "You will name your first child " & ChildName, , "Child"
        ElseIf AgeInput = 2 Then 'User designates preference of 30 to 40 age range
            Spouse = "Reese Witherspoon"
            ChildName = Left(UserFirstName, 3) + Right("Reese Witherspoon", 5)
            MsgBox "Your top choice is Reese Witherspoon.", , "Spouse"
            MsgBox "You will name your first child " & ChildName, , "Child"
        ElseIf AgeInput = 3 Then 'User input prefers age of 40 and over
            Spouse = "Nicole Kidman"
            ChildName = Left(UserFirstName, 3) + Mid("Nicole Kidman", 3, 4) + Right(UserFirstName, 2)
            MsgBox "Your top choice is Nicole Kidman.", , "Spouse"
            MsgBox "You will name your first child " & ChildName, , "Child"
        Else: MsgBox "Sorry, you entered an invalid option plese enter either a 1,2, or 3.", , "Oopsiedaisy!"
        End If
Else: MsgBox "Sorry, you entered an invalid option plese enter either a 1 or 2.", , "Oopsiedaisy!"
End If
        
End Sub

Private Sub cmdHomepage_Click() 'directs user to beginning page
frmBeginning.Show
frmWife.Hide
End Sub

Private Sub cmdQuit_Click() 'ends program
End
End Sub

Private Sub cmdSlideShow_Click()
'this subroutine is used to cycle through the array of filenames and display
'each picture for a short amount of time.

Dim name(1 To 6) As String, age(1 To 6) As Integer, haircolor(1 To 6) As String, PicCTR As Integer, stopper As Integer, t As Double, ctr2 As Double

Open App.Path & "\Women.txt" For Input As #1

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

