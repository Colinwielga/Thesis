VERSION 5.00
Begin VB.Form frmFirst 
   BackColor       =   &H000000FF&
   Caption         =   "Welcome to the SJU Lacrosse Program"
   ClientHeight    =   10740
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   10740
   ScaleWidth      =   15240
   Begin VB.CommandButton cmdCoaches 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      Caption         =   "COACHES"
      BeginProperty Font 
         Name            =   "Bell Gothic Std Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4680
      MaskColor       =   &H80000000&
      Picture         =   "frmFirst.frx":0000
      TabIndex        =   3
      Top             =   9120
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.CommandButton cmdName 
      Caption         =   "What's your name?"
      Height          =   855
      Left            =   4680
      TabIndex        =   2
      Top             =   360
      Width           =   4095
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "I""M FINISHED (but let me see my quiz score just one more time)"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   9240
      TabIndex        =   1
      Top             =   9120
      Width           =   4815
   End
   Begin VB.CommandButton cmdPlayers 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      Caption         =   "PLAYERS"
      BeginProperty Font 
         Name            =   "Bell Gothic Std Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   -120
      MaskColor       =   &H80000000&
      Picture         =   "frmFirst.frx":7A10E
      TabIndex        =   0
      Top             =   9120
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Image Image1 
      Height          =   10140
      Left            =   -720
      Picture         =   "frmFirst.frx":F421C
      Top             =   -120
      Width           =   14820
   End
End
Attribute VB_Name = "frmFirst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Title: SJU Lacrosse
'Author: Adam Benney
    'This project is an informational program for the St. John's Lacrosse club.  It is designed to offer the user a chance to learn more about the SJU lacrosse program
    ' and its players and coaches.

'Form: frmFirst, the homepage for the project
'This form first asks for the user to enter their name for later use.  Only after the users enters their name are they allowed to navigate to the other forms
    'via buttons at the bottom of the form.  The "I'm Finished..." button is available to the user at any time, regardless of whether they have entered their name or not.
   
'The information used in this project regarding players and coaches was found at www.csbsju.edu/sjulacrosseclub.


Option Explicit

Private Sub cmdCoaches_Click()                  'This button allows the user to navigate to the Coaches page (frmCoaches)
frmCoaches.Show
frmFirst.Hide
frmPlayers.Hide
frmQuiz.Hide
End Sub

Private Sub cmdName_Click()                     'This button prompt the user to input their name (the name is stored in a module so that it can be accessed throughout the project.
user_name = InputBox("Enter your name:", Name)
cmdName.Visible = False
cmdPlayers.Visible = True
cmdCoaches.Visible = True
End Sub

Private Sub cmdPlayers_Click()                  'This button allows the user to navigate to the Players page (frmPlayers)
frmFirst.Hide
frmPlayers.Show
frmCoaches.Hide
frmQuiz.Hide
End Sub

Private Sub cmdQuit_Click()                     'This button allows the user to exit the program and show their score on the coaches quiz from frmQuiz.
If points > 0 Then
    MsgBox "Your quiz score was: " & points & " of 7 questions answered correctly.", , "Quiz Score"
    MsgBox "Thanks for stopping by " & user_name, , "GOODBYE :-)"
Else
     MsgBox "You have not taken the Coaches Quiz so you do not have a score, but thanks for stopping by " & user_name, , "GOODBYE :-)"
End If
End
End Sub


