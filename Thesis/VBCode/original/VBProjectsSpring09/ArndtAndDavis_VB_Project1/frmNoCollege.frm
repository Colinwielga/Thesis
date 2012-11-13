VERSION 5.00
Begin VB.Form frmNoCollege 
   BackColor       =   &H00808000&
   Caption         =   "What Does The Future Say For Your Career?"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   BeginProperty Font 
      Name            =   "Mathematica5"
      Size            =   12
      Charset         =   2
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdCareer 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Pick Your Career and Continue"
      Height          =   855
      Left            =   8280
      MaskColor       =   &H00FFC0C0&
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   8520
      Width           =   6375
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   615
      Left            =   12000
      TabIndex        =   15
      Top             =   9960
      Width           =   1215
   End
   Begin VB.CommandButton cmdHome 
      Caption         =   "Home"
      Height          =   660
      Left            =   10440
      TabIndex        =   14
      Top             =   9960
      Width           =   1215
   End
   Begin VB.CommandButton cmdSalesPerson 
      Caption         =   "Sales Person"
      Height          =   1455
      Left            =   11640
      TabIndex        =   13
      Top             =   5880
      Width           =   2415
   End
   Begin VB.PictureBox Picture7 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   8880
      Picture         =   "frmNoCollege.frx":0000
      ScaleHeight     =   2475
      ScaleWidth      =   2355
      TabIndex        =   12
      Top             =   5520
      Width           =   2415
   End
   Begin VB.CommandButton cmdPoliceOfficer 
      Caption         =   "Police Officer"
      Height          =   1335
      Left            =   11640
      TabIndex        =   11
      Top             =   3120
      Width           =   2415
   End
   Begin VB.PictureBox Picture6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   9000
      Picture         =   "frmNoCollege.frx":2D97
      ScaleHeight     =   2235
      ScaleWidth      =   2235
      TabIndex        =   10
      Top             =   2760
      Width           =   2295
   End
   Begin VB.CommandButton cmdAthlete 
      Caption         =   "Athlete"
      Height          =   1215
      Left            =   11640
      TabIndex        =   9
      Top             =   720
      Width           =   2535
   End
   Begin VB.PictureBox Picture5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   8760
      Picture         =   "frmNoCollege.frx":8020
      ScaleHeight     =   1995
      ScaleWidth      =   2475
      TabIndex        =   8
      Top             =   240
      Width           =   2535
   End
   Begin VB.CommandButton cmdMechanic 
      Caption         =   "Mechanic"
      Height          =   1335
      Left            =   4200
      TabIndex        =   7
      Top             =   8880
      Width           =   2415
   End
   Begin VB.PictureBox Picture4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   2520
      Picture         =   "frmNoCollege.frx":A5A1
      ScaleHeight     =   2235
      ScaleWidth      =   1275
      TabIndex        =   6
      Top             =   8520
      Width           =   1335
   End
   Begin VB.CommandButton cmdEntertainer 
      Caption         =   "Entertainer"
      Height          =   1335
      Left            =   4200
      TabIndex        =   5
      Top             =   6360
      Width           =   2415
   End
   Begin VB.PictureBox Picture3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   720
      Picture         =   "frmNoCollege.frx":B65C
      ScaleHeight     =   2115
      ScaleWidth      =   3315
      TabIndex        =   4
      Top             =   6000
      Width           =   3375
   End
   Begin VB.CommandButton cmdArtist 
      Caption         =   "Artist"
      Height          =   1455
      Left            =   4200
      TabIndex        =   3
      Top             =   3600
      Width           =   2415
   End
   Begin VB.PictureBox Picture2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   1320
      Picture         =   "frmNoCollege.frx":DC17
      ScaleHeight     =   2475
      ScaleWidth      =   2595
      TabIndex        =   2
      Top             =   3120
      Width           =   2655
   End
   Begin VB.CommandButton cmdHairStylist 
      Caption         =   " Hair Stylist"
      Height          =   1455
      Left            =   4200
      TabIndex        =   1
      Top             =   720
      Width           =   2415
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   1080
      Picture         =   "frmNoCollege.frx":FE65
      ScaleHeight     =   2835
      ScaleWidth      =   2235
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmNoCollege"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: The Game Of Life
'Form Name: frmNoCollege
'Authors: Pam Arndt and Alisa Davis
'Date Written: 3/4/09
'Objective: User compares different job options for those with no college education.  This explores buttons with message boxes, input boxes, and If-ElseIf-Else statements
Option Explicit
Private Sub cmdArtist_Click()
Dim Art As String, whatart As String

Art = InputBox("Enter a 1 if you want to learn about being a painter, 2 for sculptor, 3 for graphic designer", "What kind of artist?")
If Art = 1 Then
        MsgBox "Maybe you will be good enough to join such famous names as Monet, Michaelangelo, Van Gogh, and Picasso!  Or maybe you will sell your artwork out of your garage.", , "Painter"
    ElseIf Art = 2 Then
        MsgBox "Better get used to having clay on all of your clothing!", , "Sculptor"
    ElseIf Art = 3 Then
        'Source:http://www.collegeboard.com/csearch/majors_careers/profiles/careers/106355.html
        MsgBox "You're in luck! The average yearly income of graphic designers was $45,340 in 2007, according to the U.S. Bureau of Labor Statistics.", , "Graphic Designer "
    Else
        MsgBox "Please enter a number 1 through 3 only", , "Error"

End If

End Sub

Private Sub cmdAthlete_Click()
Dim BestGuess As Single
'source: http://www.forbes.com/athletes2004/LIRWR6D.html?passListId=2&passYear=2004&passListType=Person&uniqueId=WR6D&datatype=Person
BestGuess = InputBox("How many millions of dollars do you think Tiger Woods earned for endorsements alone in 2004?", "Take a Guess")
    Select Case BestGuess
    Case Is < 70
        MsgBox "Too low!", , "Try Again"
    Case Is > 70
        MsgBox "Not quite that much!", , "Try Again"
    Case Is = 70
        MsgBox "Did you look that up?", , "Correct!"
    Case Else
        MsgBox "He earns way too much", , "Wow!"
    End Select

    
    
End Sub

Private Sub cmdCareer_Click()  'Asks the user to input their career choice, and takes them on to the next form "The Finer Things In Life"
Career = InputBox("Which career did you decide on?", "Decision time!")
frmTheFinerThingsInLife.Show
frmNoCollege.Hide

End Sub

Private Sub cmdEntertainer_Click()
MsgBox "We'll see you on TV someday!", , "Entertainer"

End Sub

Private Sub cmdHairStylist_Click()
'Disclaimer for hairstylist
MsgBox "You need a pretty artistic eye for this career. Duties may include cutting, coloring, and styling hair.  You need good customer service skills.  One bad haircut could ruin you!", , "Hairstylist"

End Sub

Private Sub cmdHome_Click()
frmBeginning.Show
frmNoCollege.Hide

End Sub

Private Sub cmdMechanic_Click()
MsgBox "You'll need to be good with your hands, and not afraid to get a bit dirty for this career!", , "Mechanic"

End Sub

Private Sub cmdPoliceOfficer_Click()
'Source: http://www.bls.gov/news.release/cfoi.nr0.htm
MsgBox "Did you know?  The number of fatal workplace injuries among protective service occupations rose 19 percent in 2007 to 337, led by an increase in the number of police officers fatally injured on the job..... but don't let me discourage you.", , "Police"

End Sub

Private Sub cmdQuit_Click()
End
End Sub


Private Sub cmdSalesPerson_Click()
Dim salesitem As String
salesitem = InputBox("What do you want to sell?", "Salesperson")
MsgBox "Sell whatever you want. Just don't call me up at dinner time!", , "Salesperson"

End Sub
