VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00808000&
   Caption         =   "About Team Manager Pro"
   ClientHeight    =   7020
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10995
   LinkTopic       =   "Form1"
   ScaleHeight     =   7020
   ScaleWidth      =   10995
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAbout 
      Caption         =   "What Team Manager Pro Can Do for YOU!"
      Height          =   855
      Left            =   8520
      TabIndex        =   6
      Top             =   4800
      Width           =   2175
   End
   Begin VB.PictureBox picResults 
      Height          =   5295
      Left            =   3480
      ScaleHeight     =   5235
      ScaleWidth      =   4275
      TabIndex        =   5
      Top             =   1320
      Width           =   4335
   End
   Begin VB.PictureBox Picture2 
      Height          =   2655
      Left            =   360
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   2595
      ScaleWidth      =   2715
      TabIndex        =   4
      Top             =   3120
      Width           =   2775
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C000&
      Height          =   2895
      Left            =   8160
      Picture         =   "frmAbout.frx":6137
      ScaleHeight     =   2835
      ScaleWidth      =   2355
      TabIndex        =   3
      Top             =   240
      Width           =   2415
   End
   Begin VB.CommandButton cmdNavigateMainMenu 
      Caption         =   "Main Menu"
      Height          =   495
      Left            =   8400
      TabIndex        =   1
      Top             =   6360
      Width           =   1335
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   9840
      TabIndex        =   0
      Top             =   6360
      Width           =   975
   End
   Begin VB.Label lbl2 
      BackColor       =   &H00808000&
      Caption         =   "Double Click Me!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   8160
      TabIndex        =   9
      Top             =   3120
      Width           =   2535
   End
   Begin VB.Label lblClick 
      BackColor       =   &H00808000&
      Caption         =   "Double Click Me!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   480
      TabIndex        =   8
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Label lblDesigner 
      BackColor       =   &H00808000&
      Caption         =   "By: Erik Gamradt"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   6720
      Width           =   1575
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Team Manager Pro"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5775
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Team Manager Pro (ErikGamradtVBProject.vbp)
'frmAbout (frmAbout.frm)
'Designed By: Erik Gamradt
'22 March 2006
'The user is able to find more information on what the programs capabilities are, including testimonials from successful users.
Option Explicit

Private Sub cmdAbout_Click()
    picResults.Cls 'clears picture box of text
    picResults.Picture = LoadPicture("") 'clears picture box of pictures
    picResults.Print "     Team Manager Pro allows coaches of all levels"
    picResults.Print "to maintain a competitive edge throughout the season."
    picResults.Print "From high school basketball tournaments, to the NCAA "
    picResults.Print "basketball tournament and the NBA playoffs, coaches "
    picResults.Print "from all levels are able to organize basketball "
    picResults.Print "statistics for their own team and all their competitors"
    picResults.Print "using Team Manager Pro.  Just a basketball fan?  Team "
    picResults.Print "Manager Pro is an excellent resource in deciding which "
    picResults.Print "team you want to pick in your March Madness bracket.  "
    picResults.Print "With so many teams to keep straight, this program will"
    picResults.Print "take the stress out of tournament time by organizing "
    picResults.Print "all the hard statistics and displaying them in an easy "
    picResults.Print "to use format. "
    picResults.Print
    picResults.Print "    As a member of Team Manager Pro, you will be able "
    picResults.Print "to enter in statistical data, compute the average, compare"
    picResults.Print "individual player averages, and compare team averages as "
    picResults.Print "well.  This will give you a heads up on what players and "
    picResults.Print "what teams to watch out for, and allow your coaching staff"
    picResults.Print "to do initial size up with unfamiliar competitors, "
    picResults.Print "especially during tournament time.  When tournament time "
    picResults.Print "roles around, no coach will want to be caught without"
    picResults.Print "Team Manager Pro and its many benefits it has to offer."
End Sub

Private Sub cmdNavigateMainMenu_Click()
    frmAbout.Hide 'brings you to the main menu
    frmMainMenu.Show
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub Picture1_DblClick()
    picResults.Cls
    picResults.Picture = LoadPicture(App.Path & "\CoachJackson.jpg") 'loads desired picture into picture box
    picResults.Print
    picResults.Print
    picResults.Print
    picResults.Print
    picResults.Print
    picResults.Print
    picResults.Print
    picResults.Print
    picResults.Print
    picResults.Print
    picResults.Print "Coach Phil Jackson"
    picResults.Print
    picResults.Print "     I love Team Manager Pro!"
End Sub

Private Sub Picture2_DblClick()
    picResults.Cls
    picResults.Picture = LoadPicture(App.Path & "\CoachK.jpg")
    picResults.Print
    picResults.Print
    picResults.Print
    picResults.Print
    picResults.Print
    picResults.Print
    picResults.Print
    picResults.Print
    picResults.Print
    picResults.Print
    picResults.Print
    picResults.Print
    picResults.Print "Coach Mike Krzyzewski"
    picResults.Print
    picResults.Print "I don't know where I would be without Team Manager Pro!"
End Sub
