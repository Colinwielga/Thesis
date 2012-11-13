VERSION 5.00
Begin VB.Form frmmain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   4875
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10080
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4875
   ScaleWidth      =   10080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdmath 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Lucky Numbers"
      BeginProperty Font 
         Name            =   "Berlin Sans FB"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   8640
      Picture         =   "frmmain.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdhealth 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Health Partners"
      DisabledPicture =   "frmmain.frx":062F
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   8640
      Picture         =   "frmmain.frx":0E82
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   480
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.CommandButton cmdwhowon 
      Caption         =   "Click here to guess what team won it all and find out if you're right"
      Height          =   855
      Left            =   5520
      TabIndex        =   6
      Top             =   2400
      Width           =   1815
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      Height          =   615
      Left            =   360
      TabIndex        =   4
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdschoolsranks 
      BackColor       =   &H00800000&
      Caption         =   "View schools with their ranks"
      Height          =   855
      Left            =   360
      MaskColor       =   &H00800000&
      TabIndex        =   3
      Top             =   1320
      Width           =   1815
   End
   Begin VB.CommandButton cmdMVP 
      Caption         =   "View MVP of any Game"
      Height          =   855
      Left            =   360
      TabIndex        =   2
      Top             =   2400
      Width           =   1815
   End
   Begin VB.CommandButton cmdgotoscores 
      Caption         =   "View Scores and Points"
      Height          =   855
      Left            =   5520
      TabIndex        =   1
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Line Line1 
      X1              =   8160
      X2              =   8160
      Y1              =   0
      Y2              =   4920
   End
   Begin VB.Label lblsponsors 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sponsors"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8640
      TabIndex        =   9
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lbljeff 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Created By: Jeff Amble"
      BeginProperty Font 
         Name            =   "NSimSun"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6000
      TabIndex        =   5
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Image imgncaa 
      Height          =   3000
      Left            =   2520
      Picture         =   "frmmain.frx":16D5
      Top             =   1200
      Width           =   2445
   End
   Begin VB.Label lblheader 
      BackColor       =   &H00FFFFFF&
      Caption         =   "March Madness 2005"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   855
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   7215
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
                         'MARCH MADNESS(VBProject)'
                                   'BY'
                               'JEFF AMBLE'
                            'CSCI 130, 3/22/06'
'This program offers a variety of information on the 2005 NCAA Men's National'
'Basketball Tournament (March Madness).  The first thing my program does is' let the'
'user view the teams in each region.  The user is able to view the teams in order of'
'their rank in their region.  The next thing my program does is enable to user to'
'the MVP of any team, in any game they have.  You are able to search this by round.'
'The user is able to check the scores of any game.  This is also done by round.  The'
'user can also check the average points scored for any team throughout the'
'tournament.  My program has the option of guessing what rank the team that won the'
'championship is.  They can guess the rank until they are right.  There is a section'
'of my main page, under sponsors that enables the user to do 2 things.  One thing'
'the user can do is check there heart rate and see if they are healthy.  The other'
'thing the user is able to do is enter their age and receive their five lucky'
'numbers.
                        
Option Explicit
'This button enables to user to exit the program'
Private Sub cmdexit_Click()
    End
End Sub
'This button enables the user to go to the scores form'
Private Sub cmdgotoscores_Click()
    frmmain.Visible = False
    frmscores.Visible = True
End Sub
'This button enables the user to go to the health form'
Private Sub cmdhealth_Click()
    frmmain.Visible = False
    frmhealth.Visible = True
End Sub

'This button enables the user to go to the main MVP form'
Private Sub cmdMVP_Click()
    frmmain.Visible = False
    frmMVP.Visible = True
End Sub
'This button enables the user to go to the main rank form'
Private Sub cmdschoolsranks_Click()
    frmmain.Visible = False
    frmschoolsranksmain.Visible = True
End Sub
'This button enables the user to go to the guess form where they can guess the rank'
'of the champion'
Private Sub cmdwhowon_Click()
    frmmain.Visible = False
    frmguess.Visible = True
End Sub
'This button enables the user to go to the math form where they can get their lucky'
'numbers'
Private Sub cmdmath_Click()
    frmmain.Visible = False
    frmmath.Visible = True
End Sub


