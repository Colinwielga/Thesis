VERSION 5.00
Begin VB.Form FrmHowTo 
   Caption         =   "Directions"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6450
   FillColor       =   &H00C00000&
   LinkTopic       =   "Form1"
   Picture         =   "FrmHowTo.frx":0000
   ScaleHeight     =   5775
   ScaleWidth      =   6450
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cmdtwo 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   2
      Top             =   5160
      Width           =   2535
   End
   Begin VB.CommandButton Cmdone 
      Caption         =   "Go Back to Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   5160
      Width           =   2535
   End
   Begin VB.PictureBox PicHowTo 
      BackColor       =   &H00FFFFC0&
      Height          =   4335
      Left            =   360
      ScaleHeight     =   4275
      ScaleWidth      =   5475
      TabIndex        =   0
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "FrmHowTo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Hang Man
'Form Name: FrmHowTo
'Authors: Breanna Riskey and Heidi Golliet
'Date Completed: Monday, March 31st
'Objective: The purpose of this form is to tell the user how to operate the program

Option Explicit

Private Sub Cmdone_Click()
    FrmOptions.Visible = True
    FrmHowTo.Visible = False
    FrmPlayGame.Visible = False
    FrmHome.Visible = False
End Sub

Private Sub Cmdtwo_Click()
    End
End Sub

'this displays the directions for the game
Private Sub Form_Activate()
    PicHowTo.Print Tab(25); "Directions For Play"
    PicHowTo.Print "******************************************************************************************"
    PicHowTo.Print "Welcome to Hangman! We hope you have a great time playing our game!"
    PicHowTo.Print "In case you aren't familiar with the game we've provided directions"
    PicHowTo.Print "on how to play. The main objective is to solve the word before the"
    PicHowTo.Print "entire hangman appears. Begin by clicking on the button that asks you"
    PicHowTo.Print "to enter a letter. A box will appear and then you enter a letter. If "
    PicHowTo.Print "the letter is correct, it will appear within the blanks. If it is "
    PicHowTo.Print "incorrect, a part of the hangman will appear. The game continues on in "
    PicHowTo.Print "this manner until the word is solved or you have a complete hangman "
    PicHowTo.Print "(in which case you lose the game). You can also click on hint"
    PicHowTo.Print "if you are having problems solving the word."
End Sub




