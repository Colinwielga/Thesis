VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H80000008&
   Caption         =   "Form2"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   15240
   ScaleWidth      =   25080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Back to Game"
      Height          =   1575
      Left            =   10800
      TabIndex        =   2
      Top             =   13440
      Width           =   8175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Display Rules"
      Height          =   1335
      Left            =   2160
      TabIndex        =   1
      Top             =   13560
      Width           =   6015
   End
   Begin VB.PictureBox PicResults 
      Height          =   12855
      Left            =   480
      ScaleHeight     =   12795
      ScaleWidth      =   20355
      TabIndex        =   0
      Top             =   240
      Width           =   20415
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
'Command button shows displays the rules of the game in full detail in a picture box.

PicResults.Print "Game___ Instructions / Rules"
PicResults.Print "-There will be two different teams playing against each other; Team 1 and Team 2."
PicResults.Print "The game begins with a list of 6 words only showing the first and last words in the list."
PicResults.Print "The middle four words will be blank before the round begins.  The object of the game is to try and guess more missing words than your opponent."
PicResults.Print "All of the words will relate to the words above it and below it.  Each word is part of a two word phrase.  "
PicResults.Print ""
PicResults.Print "A sample Chain would be:"
PicResults.Print ""
PicResults.Print "Super"
PicResults.Print ""
PicResults.Print "Charge"
PicResults.Print ""
PicResults.Print "Card"
PicResults.Print ""
PicResults.Print "Shark"
PicResults.Print ""
PicResults.Print "Fin"
PicResults.Print ""
PicResults.Print "Soup"
PicResults.Print ""
PicResults.Print "The words are revealed one letter at a time by the user selecting the corresponding button."
PicResults.Print ""
PicResults.Print "For the example above :"
PicResults.Print " 1)  The user will select the button (Letter below) and the letter ( C ) will be displayed under the word (Super.)"
PicResults.Print ""
PicResults.Print "2)  A team's turn then consists, of typing a guess, of the word above or below one of the already revealed."
PicResults.Print ""
PicResults.Print "3)  A correct response wins the money for each word and the team remains guessing to complete the chain until incorrect."
PicResults.Print ""
PicResults.Print "4)   If the team in control guesses incorrectly, or gives no answer, control goes back to the other team."
PicResults.Print ""
PicResults.Print "5)  The game continues until the chain is finished."
PicResults.Print ""
PicResults.Print "6)  After this chain is finished the team who completes the last word of the chain qualifies for the bonus round to gain more money."
PicResults.Print ""
PicResults.Print "7)   In the bonus round, the team has to complete a four word chain with the first and last words showing and only the first letter of the middle two words given. An example could be:"
PicResults.Print ""
PicResults.Print "HALF"
PicResults.Print "B_______ (BAKED)"
PicResults.Print "A_______ (ALASKA)"
PicResults.Print "PIPELINE"
PicResults.Print ""
PicResults.Print "If answered correctly the team will earn bonus money.   If answered incorrectly the team will lose NO money but will remain in control to begin the next round."
PicResults.Print "8)  There will be 3 full rounds.  The first two rounds will consist of each team attempting to win the round and the bonus rounds."
PicResults.Print ""
PicResults.Print "9)   At the end of the third round, the team with the most money will move on to the final bonus round.  The losing team is done for the game."
PicResults.Print "The final bonus round, is similar to the previous bonus rounds, except that there is one extra incomplete word to be completed in the chain."
PicResults.Print "The winner will have to win the final bonus to keep their previous earnings. But as an incentive, they will have the chance to earn an extra $25,000.00 in a successful completion of the final bonus round."
PicResults.Print ""
PicResults.Print "At the end of this bonus round the game is over."
PicResults.Print ""
PicResults.Print "**Remember when typing in your guesses into the input boxes, all answers begin with a capital letter and must be spelled correctly.**"
Command1.Enabled = False
Command2.Enabled = True

End Sub

Private Sub Command2_Click()
'Command button brings the user in control back to form1 to play the rest of the game.
Form2.Hide
Form1.Show
End Sub

