VERSION 5.00
Begin VB.Form frmTitle 
   Caption         =   "Form1"
   ClientHeight    =   8985
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11865
   LinkTopic       =   "Form1"
   Picture         =   "frmTitle.frx":0000
   ScaleHeight     =   8985
   ScaleWidth      =   11865
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picAnswer 
      Height          =   1935
      Left            =   6960
      ScaleHeight     =   1875
      ScaleWidth      =   4635
      TabIndex        =   5
      Top             =   3720
      Width           =   4695
   End
   Begin VB.CommandButton cmdGrammar 
      Caption         =   "Grammar Questions???"
      Height          =   735
      Left            =   8640
      TabIndex        =   4
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit!"
      Height          =   735
      Left            =   9360
      TabIndex        =   3
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton cmdWar 
      Caption         =   "Wartime Letter"
      Height          =   1215
      Left            =   6480
      TabIndex        =   2
      Top             =   7440
      Width           =   1335
   End
   Begin VB.CommandButton cmdWarming 
      Caption         =   "Fighting Global Warming"
      Height          =   1215
      Left            =   4560
      TabIndex        =   1
      Top             =   7440
      Width           =   1335
   End
   Begin VB.CommandButton cmdBaby 
      Caption         =   "Naming the Baby"
      Height          =   1215
      Left            =   2760
      Picture         =   "frmTitle.frx":9BBF
      TabIndex        =   0
      Top             =   7440
      Width           =   1335
   End
End
Attribute VB_Name = "frmTitle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Title:  VaB-Libs
'Form:  Title Screen
'Author:  Jed Anderson
'Date Written: March 20th, 2008
'Objective:  It's the title screen.  You start and end here.



Private Sub cmdBaby_Click()
'This changes forms when the Baby Naming button is clicked.
frmTitle.Hide
frmBaby.Show

End Sub

'This part is to give instruction on the parts of speech if needed.

Private Sub cmdGrammar_Click()
Dim Grammar As String

Grammar = InputBox("What part of speech is confusing you?", "Questions")

Select Case Grammar
    Case Is = LCase("verb")
        picAnswer.Cls
        picAnswer.Print "The verb is perhaps the most important part of the sentence."
        picAnswer.Print "A verb or compound verb asserts something about the subject "
        picAnswer.Print "of the sentence and express actions, events, or states of"
        picAnswer.Print " being.  The verb or compound verb is the critical element"
        picAnswer.Print "of the predicate of a sentence."
    Case Is = LCase("adjective")
        picAnswer.Cls
        picAndwer.Print "An adjective modifies a noun or a pronoun by describing, "
        picAndwer.Print "identifying, or quantifying words. An adjective usually "
        picAndwer.Print "precedes the noun or the pronoun which it modifies."
    Case Is = LCase("noun")
        picAnswer.Cls
        picAnswer.Print "A noun is a word used to name a person, animal, place, "
        picAnswer.Print "thing, and abstract idea. Nouns are usually the first words "
        picAnswer.Print "which small children learn."
    Case Is = LCase("adverb")
        picAnswer.Cls
        picAnswer.Print "An adverb can modify a verb, an adjective, another adverb, a "
        picAnswer.Print "phrase, or a clause. An adverb indicates manner, time, "
        picAnswer.Print "place, cause, or degree and answers questions such as how, "
        picAnswer.Print "when, Where, and how much."
    Case Is = LCase("pronoun")
        picAnswer.Cls
        picAnswer.Print "A pronoun can replace a noun or another pronoun. You use "
        picAnswer.Print "pronouns like he, which, none, and you, "
        picAnswer.Print "to make your sentences less cumbersome and less repetitive."
    Case Else
        picAnswer.Cls
        picAnswer.Print "I'm sorry, but either you made a typing error or that is"
        picAnswer.Print "a part of speech I am not familiar with."
End Select

End Sub

Private Sub cmdWar_Click()
'This changes forms when the War Letter button is clicked.
frmTitle.Hide
frmWarLetter.Show

End Sub

Private Sub cmdWarming_Click()
'This changes forms when the Global Warming button is clicked.
frmTitle.Hide
frmGlobal.Show

End Sub
Private Sub cmdQuit_Click()
'This ends the program when the Quit button is clicked.
End

End Sub

