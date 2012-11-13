VERSION 5.00
Begin VB.Form Rules 
   BackColor       =   &H0000C0C0&
   Caption         =   "Form1"
   ClientHeight    =   11370
   ClientLeft      =   6645
   ClientTop       =   4395
   ClientWidth     =   12885
   LinkTopic       =   "Form1"
   ScaleHeight     =   11370
   ScaleWidth      =   12885
   Begin VB.PictureBox picHowie 
      Height          =   4335
      Left            =   1800
      Picture         =   "DoND3.frx":0000
      ScaleHeight     =   4275
      ScaleWidth      =   7395
      TabIndex        =   6
      Top             =   5160
      Width           =   7455
   End
   Begin VB.CommandButton cmdGuess 
      Caption         =   "Guess Howie Mandel's Age"
      Height          =   1215
      Left            =   7560
      TabIndex        =   5
      Top             =   3840
      Width           =   3495
   End
   Begin VB.CommandButton cmdPlayers 
      Caption         =   "Who's Involved?"
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Top             =   3840
      Width           =   3495
   End
   Begin VB.PictureBox picResults2 
      Height          =   2415
      Left            =   120
      ScaleHeight     =   2355
      ScaleWidth      =   9195
      TabIndex        =   3
      Top             =   1200
      Width           =   9255
   End
   Begin VB.CommandButton cmdRulesDescription 
      Caption         =   "Get a Description of the Rules"
      Height          =   1215
      Left            =   3840
      TabIndex        =   2
      Top             =   3840
      Width           =   3495
   End
   Begin VB.CommandButton cmdBacktoStart 
      Caption         =   "Back to Main Menu"
      Height          =   2415
      Left            =   9600
      TabIndex        =   1
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label lblRules 
      Caption         =   "Rules of Deal or No Deal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10935
   End
End
Attribute VB_Name = "Rules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Deal or No Deal Introduction
'Form Name: Start
'Authors: Chris Bergstrom and Brady King
'Date Written: March 27th, 2008

'Objective of Form: To give an indepth explanation of the rules of the show.
'Also, the user will learn the roles of the people they see on the screen.
'And for fun, guess the age of the host of the American version, Howie Mandel.


Private Sub cmdBacktoStart_Click() 'Goes back to greeting form.
Rules.Hide 'Hides the Rules form.
Start.Show 'Displays the Greeting form.
End Sub

Private Sub cmdGuess_Click() 'This is a game meant to be fun yet informative about the host.
Dim Age As Single

Age = InputBox("Enter in you guess as to how old howie is?") 'Takes number from user.
If Age < 1 Then                                               'Then runs a check to see which category the number falls into.
MsgBox ("Come on, Get real. A human can't be " & Age & " years old.")
ElseIf Age <= 25 And Age >= 1 Then
MsgBox ("Nope, Keep Trying. " & Age & " is too young.")        'The computer then displays the appropriate message for that age group.
ElseIf Age <= 50 Then
MsgBox ("Closer...")
ElseIf Age = 52 Then
MsgBox ("BINGO! Howie is " & Age & " years old!")
ElseIf Age >= 53 And Age <= 75 Then
MsgBox (Age & " is too old! Try again.")
ElseIf Age > 75 And Age <= 100 Then
MsgBox ("Really? He's no young buck, but now you're just being mean!")
ElseIf Age > 100 Then
MsgBox ("We're talking about Howie Mandel... not Moses.")

End If ' This ends the check system.

End Sub

Private Sub cmdPlayers_Click() 'This informs the user of who is involved with the on-screen
                                'facilitation of the show.
picResults2.Cls
picResults2.Print "Deal or No Deal involves:" 'This displays the actors involved in Deal or No Deal in the picture box.
picResults2.Print
picResults2.Print "1) A contestant."
picResults2.Print
picResults2.Print "2) A host (In the American version actor Howie Mandel acts as the host/presenter)."
picResults2.Print
picResults2.Print "3) A banker, who is unknown to the audience."
picResults2.Print
picResults2.Print "4) And a group of models (In the American, Austrailian, Malaysian, and New Zealand Versions).**"
picResults2.Print
picResults2.Print "**In the original version models were not part of the show. Other contestants were used to help open briefcases.**"
End Sub

Private Sub cmdRulesDescription_Click() 'This describes the rules and procedures of the game.
picResults2.Cls 'This displays a description of the rules in the picture box.
picResults2.Print "The game begins with the contestant choosing a case (1 of 26) which they think will have the highest value."
picResults2.Print "Next the contestant chooses the remaining 25 cases one at a time for rejection."
picResults2.Print "The value of each rejected case is revealed to the contestant after it has been discarded."
picResults2.Print "Pressure mounts each round. After a specific portion of the cases are opened, the banker offers an amount of money.*"
picResults2.Print "If the contestant is content with one of the banker's offers, the game is over and the player wins the offered amount."
picResults2.Print "If unsatisfied with all the offers, they finish with the unknown amount from the first case chosen, which is then revealed.**"
picResults2.Print
picResults2.Print "*Prompting Howie Mandel to ask the all-important question: Deal or No Deal?"
picResults2.Print "**Unless they decide to swap their case with the last one in the gallery, then that chosen case is later revealed."



End Sub

