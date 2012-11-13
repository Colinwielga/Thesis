VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H0000C000&
   Caption         =   "Form5"
   ClientHeight    =   8655
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11055
   LinkTopic       =   "Form5"
   ScaleHeight     =   8655
   ScaleWidth      =   11055
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3720
      TabIndex        =   4
      Top             =   5400
      Width           =   3735
   End
   Begin VB.CommandButton cmdQuit3 
      Caption         =   "Back to Selection Screen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7680
      TabIndex        =   1
      Top             =   7560
      Width           =   2775
   End
   Begin VB.Label lblnumbers 
      BackColor       =   &H0000C000&
      Caption         =   "58017"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1215
      Left            =   4200
      TabIndex        =   3
      Top             =   3360
      Width           =   2775
   End
   Begin VB.Label lblInstructions 
      BackColor       =   &H0000C000&
      Caption         =   $"Project5.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1095
      Left            =   1920
      TabIndex        =   2
      Top             =   1920
      Width           =   7815
   End
   Begin VB.Label lbl5Chances 
      BackColor       =   &H0000C000&
      Caption         =   "FIVE CHANCES"
      BeginProperty Font 
         Name            =   "@BatangChe"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1095
      Left            =   2640
      TabIndex        =   0
      Top             =   240
      Width           =   6015
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit     'In this game, the player is presented with five random numbers on the screen.
Dim Price As Double 'When ordered correctly, they form the price of a car.
Dim num As Double   'The player clicks the play button and is presented with an input box instructing them to enter the numbers in a different order, trying to form the price of the car.
                    'The player will have 5 chances to form the correct price.  If they can do it in five chances, they win the car; if not, they lose.
                    'Under Option Explicit I declared my variables.
                    '"num" acts as a counter, allowing the player only five chances.

Private Sub cmdPlay_Click()
Do Until num = 5
Price = InputBox("Enter your guess.  Remember to use each of the numbers 5,8,0,1, and 7 only once each.  Do not include commas or spaces in your answer.")
num = num + 1
If Price = 18570 Then   'As soon as the player enters the correct guess, they "win" the car.
    MsgBox "Congratulations!  The correct price is $18,570!  You win the car!"
    Form2.Show  'Takes the player back to the selection screen if they win.
    Form5.Hide
Else        'If they guess incorrectly, they can continue trying until the five chances are used.
    MsgBox "I'm sorry, that is incorrect.  Please try again."
    End If
Loop
MsgBox "The actual retail price was $18,570.  I'm sorry, you lose." 'If the chances run out before the correct price is guessed, the game ends.

End Sub

Private Sub cmdQuit3_Click()
Form2.Show  'Allows the player to go back to the selection screen.
Form5.Hide
End Sub

