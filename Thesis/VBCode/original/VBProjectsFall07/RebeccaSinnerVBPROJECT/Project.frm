VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000008&
   Caption         =   "Form1"
   ClientHeight    =   8475
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10920
   LinkTopic       =   "Form1"
   ScaleHeight     =   8475
   ScaleWidth      =   10920
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdComeOnDown 
      BackColor       =   &H80000008&
      Caption         =   "Come on down!"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3360
      MaskColor       =   &H00000000&
      TabIndex        =   1
      Top             =   6240
      Width           =   4575
   End
   Begin VB.PictureBox picResults 
      BorderStyle     =   0  'None
      Height          =   4695
      Left            =   2520
      Picture         =   "Project.frx":0000
      ScaleHeight     =   4695
      ScaleWidth      =   5895
      TabIndex        =   0
      Top             =   1200
      Width           =   5895
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H80000007&
      Caption         =   "It's Time For..."
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   855
      Left            =   3840
      TabIndex        =   2
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit         'This project is based on the game show "The Price is Right."  Players must first go through an opening bidding round.  If they are successful in this round, they have the chance to choose from three different "pricing games."  The games are: Eazy as 1-2-3, Punch a Bunch, and Five Chances.
Dim Bid As Double       'This is the first form in the project and the opening screen.
Dim Bid2 As Double      'I designed it to be played similarily to the bidding round in the Price is Right.
Dim Bid3 As Double      'The player is presented with an item to "bid" on.  The player is instucted to guess as close to the retail price of the product without going over.
Dim Playing As Boolean  'If they guess incorrectly, they have a chance to bid on more products.  As soon as they bid correctly, they "win" that product and the chance to move onto the next form and choose from a list of "pricing games" to play.
                        'Under Option Explicit, I declared my variables.

Private Sub cmdComeOnDown_Click()
MsgBox "Remember, whoever comes closest to the actual retail price without going over wins the bidding round and the chance to play a pricing game.  All actual retail prices will be whole numbers."
Playing = True      'This variable stays True to indicate that the game is still going on, but switches to False as soon as a correct bid is made because then this part of the game is over.

Bid = InputBox("Our first item up for bid is the iPod Shuffle.  Please enter your bid.", "Please enter your bid.")
If Bid >= 70 And Bid <= 79 Then
    MsgBox "The actual retail price is $79.  Congratulations!  You win the shuffle plus and the chance to play a pricing game."
    Form2.Show      'As soon as a correct bid is made, the first form closes and a Form2 opens.
    Form1.Hide
    Playing = False 'Playing changes to False when a correct bid is made to end this part of the game.
    End If
If Bid > 79 Or Bid < 70 Then
    MsgBox "I'm sorry.  The actual retail price is $79.  You did not enter a winning bid."
End If
If Playing Then
Bid2 = InputBox("Our next item up for bid is a year of education at CSB/SJU.  Please enter your bid.", "Please enter your bid.")
If Bid2 <= 34000 And Bid2 >= 32000 Then
    MsgBox "The actual reatail price is $34,000.  Congratulations!  You win a year at CSB/SJU and the chance to play a pricing game."
    Form2.Show      'As soon as a correct bid is made, the first form closes and a Form2 opens.
    Form1.Hide
    Playing = False
    End If
    If Bid2 < 32000 Or Bid2 > 34000 Then
        MsgBox "I'm sorry.  The actual retail price is $34,000.  You did not enter a winning bid."
            End If
            End If
            If Playing Then
Bid3 = InputBox("Our third and final item up for bid is a pair of American Eagle jeans.  Please enter your bid.", "Please enter your bid.")
If Bid3 >= 39 And Bid3 <= 50 Then
    MsgBox "The actual Retail price is $50.  Congratulations!  You win the jeans and the chance to play a pricing game."
    Form2.Show      'As soon as a correct bid is made, the first form closes and a Form2 opens.
    Form1.Hide
    Playing = False 'Playing changes to False when a correct bid is made to end this part of the game.
    End If
    If Bid3 < 39 Or Bid3 > 50 Then
    MsgBox "I'm sorry.  The actual Retail price is $50.  You did not enter a winning bid.  GAME OVER!"
    End If
    End If
                
End Sub


