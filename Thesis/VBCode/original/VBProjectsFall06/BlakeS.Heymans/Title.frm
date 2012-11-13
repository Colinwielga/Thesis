VERSION 5.00
Begin VB.Form frmTitle 
   BackColor       =   &H80000009&
   Caption         =   "Form1"
   ClientHeight    =   5730
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7890
   LinkTopic       =   "Form1"
   Picture         =   "Title.frx":0000
   ScaleHeight     =   5730
   ScaleWidth      =   7890
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdChere 
      Caption         =   "Click Here"
      BeginProperty Font 
         Name            =   "Rockwell Condensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3240
      TabIndex        =   5
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Rockwell Condensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6360
      TabIndex        =   3
      Top             =   4920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdProfiles 
      Caption         =   "See SJU Player Profiles"
      BeginProperty Font 
         Name            =   "Rockwell Condensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   2
      Top             =   4440
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdOpponentsearch 
      Caption         =   "Search for an Opponent by School "
      BeginProperty Font 
         Name            =   "Rockwell Condensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   1
      Top             =   3120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdNamesearch 
      Caption         =   "Search for Player Scores by Name"
      BeginProperty Font 
         Name            =   "Rockwell Condensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   1800
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Image imgTennis 
      Height          =   1305
      Left            =   2280
      Picture         =   "Title.frx":AA82
      Top             =   2400
      Visible         =   0   'False
      Width           =   3750
   End
   Begin VB.Label lbltitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Player Profiles and General Information For The 2006 MIAC Tournament!"
      BeginProperty Font 
         Name            =   "Rockwell Condensed"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   960
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   6015
   End
End
Attribute VB_Name = "frmTitle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
'2006 MIAC Tennis Tournament Distribution
'Title Form
'Blake Heymans
'10/25/06
'Overall Project Purpose
    'The purpose of this project is to provide coaches, player, parents, or just any
    'person interested in Saint John's Tennis with the results of
    'individual players and teams in the 2006 MIAC tennis tournament
    'as well as Saint John's player profiles. Often a person watching a
    'tennis tournament with so many players and matches in progress at a given
    'time will not be able to see every match but would like to know the
    'outcome of each match. This program is valuable for strategic reasons
    'but it is also valuable for sentimental reasons. A strategic reason
    'would be if a coach may want to see which players win or score better against
    'certain opponents. But the highest value of this program is its sentimental value.
    'Most tennis fans buy things that remind them of these tournaments later in life.
    'This would be the best way to market this program.
    'In the program there are efficient ways to search and find players or opponents.
    'It is important that the user be able to find both
    'opponents and individual player results.
'***Pictures were taken from Google image search as well as the Saint John's Univerity web site.
'Title Form Objective
    'This form is intended to lead the user into the program.  First it
    'displays a welcome message promting the user to click the buttons on the
    'Title Form. Then the Title Form displays the title of the program as well
    'as a directional set of buttons nesecessary to navigate the program.
    'The directional buttons send the user to various Forms with different funtctions.
'Pictures were taken from Google image search as well as the Saint John's Univerity web site.
Private Sub cmdChere_Click()
    
    'Promts the user with a welcome message then displays the various buttons and images.
    MsgBox "Welcome to the 2006 MIAC Tennis Tournament Distribution!  We Hope You Enjoyed the Tournament.                                           To Navigate Through the Distribution Please Click on the Buttons Provided.", , "2006 MIAC Tennis Tournament Distribution"
    
    cmdChere.Visible = False
    
    cmdNamesearch.Visible = True
    cmdOpponentsearch.Visible = True
    cmdProfiles.Visible = True
    cmdQuit.Visible = True
    imgTennis.Visible = True
    lbltitle.Visible = True

End Sub

Private Sub cmdNamesearch_Click()
    'Hides Title Form and shows Player Search Form
    frmPlayersearch.Show
    frmTitle.Hide
    
    MsgBox "Please Enter the Saint John's Player's Full Name Provided on the Roster.", , "Player Search"
End Sub

Private Sub cmdOpponentsearch_Click()
    'Hides Title Form and shows Opponent Search Form
    frmOpponent.Show
    frmTitle.Hide
    'Opens Input Box to be used on the Opponent Search Form
    Opp = InputBox("Your Choices Are UST, CC, and GAC.                                                                                                                      Be Sure to Capitalize the Letters!", "Enter An Opponent")
    MsgBox "Please Click on the MIAC Symbol of the Image of Minnesota to See The Results of the Match."
End Sub

Private Sub cmdProfiles_Click()
    'Hides Title Form and shows Player Profile Search Form
    frmProfiles.Show
    frmTitle.Hide
    
    MsgBox "To See a Player Profile, Enter the Corresponding Number of the Player Provided. Then Press the Profile Button", , "Profile Prompt"
End Sub

Private Sub cmdQuit_Click()
'Terminates Program
End
End Sub
