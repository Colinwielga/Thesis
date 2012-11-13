VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   Picture         =   "StartScreen.frx":0000
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play Oregon Trail! "
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   0
      MaskColor       =   &H000000FF&
      TabIndex        =   1
      Top             =   0
      Width           =   4575
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Don't Play Oregon Trail (Because you just don't have the spirit of adventure today.)"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   360
      TabIndex        =   0
      Top             =   5040
      Width           =   4095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Oregon Trail
'Drew Madson & Sam Erickson Oct 2008
'This is the starting page of Oregon trail. It's pretty simple.

'Project Description
'This program seeks to mirror and satarize some of the more memorable parts of the old school video game, Oregon Trail. How can we simulate a hunting game? How can we create a store to purchase the goods for travel? And how can we select the characters and the course the journey takes? These were are most pressing questions.
'We chose this project because Oregon Trail is such a simple video game and we thought it'd be fun to emulate.
'The user engages in the game by entering information (such as user name and items to be purchased) and choosing options among buttons. Most of the buttons are on the home page where the user repeatedly returns. There is no real strategy to this program-just simple, fun engagement.


Private Sub cmdPlay_Click()

Dim UserName As String

    UserName = InputBox("What's your name partner? If you reckon' to hit the the trail, we better know what name to put on your grave.", "Welcome!") 'Retrieves and stores UserName in module/public
    Form1.Hide 'hides start page from user
    Form2.Show 'shows main page to user
    MsgBox "Welcome to the greatest thing you'll ever do, " & UserName & ".", , "Salutations."
End Sub

    
Private Sub cmdQuit_Click()

    End

End Sub

