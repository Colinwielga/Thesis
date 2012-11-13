VERSION 5.00
Begin VB.Form frmLoad 
   Caption         =   "DINOSAURS!"
   ClientHeight    =   4020
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5970
   LinkTopic       =   "Form1"
   Picture         =   "frmLoad.frx":0000
   ScaleHeight     =   4020
   ScaleWidth      =   5970
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDinoTrivia 
      Caption         =   "Let's Find Out About Dinosaurs!"
      Height          =   735
      Left            =   2280
      TabIndex        =   2
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit the Program"
      Height          =   735
      Left            =   4080
      TabIndex        =   1
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton cmdBeginAdventure 
      BackColor       =   &H8000000E&
      Caption         =   "Begin the Adventure!"
      Height          =   735
      Left            =   480
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3240
      Width           =   1575
   End
End
Attribute VB_Name = "frmLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This program gives two options for the user.
'The first is to take part in an "Choose Your Own Adventure" Story
'The second is learn some information about popular dinosaurs

Private Sub cmdBeginAdventure_Click()
    'The purpose of this part of the program is to let the user
    'choose their own ending to the story.
    'There are multiple endings, and each choice the user makes will
    'Have different consequences, and lead to different endings.
    
    'This button begins the "Choose Your Own Adventure" Story
    'By clicking on the Begin Adventure Button, the player advances to the next form
    frmLoad.Visible = False
    frmStep1.Visible = True
    
    'Ask for player's name
    YourName = InputBox("Please enter your name for our future reference.", "User's Name")
    'Confirms the player's name
    MsgBox "Thank you " & YourName & ", now we may begin the adventure.", , "Name Confirmation"
    
End Sub

Private Sub cmdDinoTrivia_Click()
    'By clicking the button, the player advances to the trivia
    frmLoad.Visible = False
    frmTrivia.Visible = True
End Sub

Private Sub cmdQuit_Click()
    'Ends program
    End
End Sub
