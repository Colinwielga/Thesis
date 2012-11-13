VERSION 5.00
Begin VB.Form frmStart 
   BackColor       =   &H00000080&
   Caption         =   "A Game of Thrones"
   ClientHeight    =   8340
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11445
   LinkTopic       =   "Form1"
   ScaleHeight     =   8340
   ScaleWidth      =   11445
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNaming 
      BackColor       =   &H8000000A&
      Caption         =   "Who are you, sir?"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7560
      TabIndex        =   4
      Top             =   600
      Width           =   2895
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00C0FFFF&
      FillColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1695
      Left            =   6120
      ScaleHeight     =   1635
      ScaleWidth      =   5115
      TabIndex        =   3
      Top             =   2160
      Width           =   5175
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Capitulate"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   2
      Top             =   6240
      Width           =   2775
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Forward!"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8280
      TabIndex        =   0
      Top             =   6240
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Created by ACR Schmelzer"
      BeginProperty Font 
         Name            =   "Mistral"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7680
      TabIndex        =   5
      Top             =   7800
      Width           =   3735
   End
   Begin VB.Image Image1 
      Height          =   4005
      Left            =   0
      Picture         =   "GTForm1.frx":0000
      Top             =   0
      Width           =   6000
   End
   Begin VB.Label lblWelcome 
      BackColor       =   &H00C0FFFF&
      Caption         =   $"GTForm1.frx":10A14
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   4440
      Width           =   10335
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Objective: to create a user name (Lord name); to initialize key boolean variables
''skill' and 'personality'; to introduce the user to the game
'Explaination: the user is asked to input string answers that serve as variables
'that work to create a user name and initialize to key public boolean variables:
''skill' and 'personality'
'the for also sets a public variable that is used to either restart the game and
'make sure the user enters correct string variables
'the input of certain variables also affects the battlepoints public variable


Private Sub cmdNaming_Click()
Dim Firstname As String
Dim Middlename As String
Dim Surname As String
Dim Street As String
Dim Mname As String
Dim Season As String
Dim Direction As String
Dim Sigil As String
Dim Weapon As String
Dim Skill As String
Dim Personality As String
my1variable = False
my2variable = False

Courage = False
Strength = False
Cunning = False
Persuasion = False
Orator = False
Looks = False
Lover = False
Scholar = False
Warrior = False


picResults.Cls
    Firstname = InputBox("Enter Your First Name", "Your Name, sir.")
    Middlename = InputBox("Enter Your Middle Name", "Your Middle Name, sir.")
    Surname = InputBox("Enter the name of your line", "Your Family Name, sir")
    Street = InputBox("Enter your street of residence", "Your street name, sir")
    Mname = InputBox("Enter your noble mother's family name", "Your mother's maiden name, sir")
    Season = InputBox("Enter your season of favor", "The season of your birth, sir")
    Direction = InputBox("Enter the direction of your lands (North, South, East, or West)", "The location of your lands, sir")
    Sigil = InputBox("Enter the animal you favor the most: Lion, Dragon, Wolf, Leviathan, Falcon, Stag,", "Your favorite animal, sir")
    Weapon = InputBox("Enter your weapon of choice: Sword, Axe, Spear, Bow, Knife, Mace, Warhammer, Morning Star,", "Your weapon, sir")
    Skill = InputBox("What is your paramount skill: Courage, Cunning, Strength, Persuasion, Looks", "Your skill, sir")
    Personality = InputBox("What type of person are you: Warrior, Scholar, Orator, Lover?", "Your personality, sir")
    
If Skill = "Courage" Or Skill = "courage" Or Skill = "Cunning" Or Skill = "cunning" Or Skill = "strength" Or Skill = "Strength" Or Skill = "persuasion" Or Skill = "Persuasion" Or Skill = "looks" Or Skill = "Looks" Then
    my1variable = True
End If
If Personality = "warrior" Or Personality = "Warrior" Or Personality = "scholar" Or Personality = "Scholar" Or Personality = "strength" Or Personality = "Strength" Or Personality = "orator" Or Personality = "Orator" Or Personality = "Lover" Or Personality = "lover" Then
    my2variable = True
End If
'set boolean variables
If Skill = "Courage" Or Skill = "courage" Then
    Courage = True
    Battlepoints = Battlepoints + 750
End If
If Skill = "Cunning" Or Skill = "cunning" Then
    Cunning = True
    Battlepoints = Battlepoints + 750
End If
If Skill = "strength" Or Skill = "Strength" Then
    Strength = True
    Battlepoints = Battlepoints + 750
End If
If Skill = "persuasion" Or Skill = "Persuasion" Then
    Persuasion = True
    Battlepoints = Battlepoints + 500
End If
If Skill = "looks" Or Skill = "Looks" Then
    Looks = True
    Battlepoints = Battlepoints + 100
End If
If Personality = "warrior" Or Personality = "Warrior" Then
    Warrior = True
    Battlepoints = Battlepoints + 750
End If
If Personality = "scholar" Or Personality = "Scholar" Then
    Scholar = True
    Battlepoints = Battlepoints + 1000
End If
If Personality = "orator" Or Personality = "Orator" Then
    Orator = True
    Battlepoints = Battlepoints + 1000
End If
If Personality = "lover" Or Personality = "Lover" Then
    Lover = True
    Battlepoints = Battlepoints + 0
End If

    
picResults.Print "Hail Lord "; Middlename; " "; Street; " of House "; Surname; "."
picResults.Print " Defender of the "; Direction; ", Son of "; Season
picResults.Print " and the great house of "; Mname
Lordname = Middlename & Street & Surname

End Sub

Private Sub cmdQuit_Click()
End
End Sub
Private Sub cmdStart_Click()

If my1variable = True And my2variable = True Then
    frmStart.Hide
    frmDeclaration.Show
Else
    MsgBox "You must let yourself be known!", , "Please declare your name and title, sir."
End If
End Sub

