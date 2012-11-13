VERSION 5.00
Begin VB.Form frmBaby 
   Caption         =   "Form2"
   ClientHeight    =   9135
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11955
   LinkTopic       =   "Form2"
   Picture         =   "frmBaby.frx":0000
   ScaleHeight     =   9135
   ScaleWidth      =   11955
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMakeBaby 
      Caption         =   "Generate VaB-Lib!"
      Height          =   735
      Left            =   1920
      TabIndex        =   18
      Top             =   5160
      Width           =   1815
   End
   Begin VB.PictureBox picResults 
      Height          =   6015
      Left            =   3960
      ScaleHeight     =   5955
      ScaleWidth      =   7395
      TabIndex        =   17
      Top             =   120
      Width           =   7455
   End
   Begin VB.CommandButton cmdMain 
      Caption         =   "Main Menu"
      Height          =   735
      Left            =   600
      TabIndex        =   16
      Top             =   5160
      Width           =   1095
   End
   Begin VB.TextBox txtTV 
      Height          =   285
      Left            =   2520
      TabIndex        =   15
      Top             =   4440
      Width           =   1335
   End
   Begin VB.TextBox txtPet 
      Height          =   285
      Left            =   2520
      TabIndex        =   14
      Top             =   3840
      Width           =   1335
   End
   Begin VB.TextBox txtBrand 
      Height          =   285
      Left            =   2520
      TabIndex        =   13
      Top             =   3240
      Width           =   1335
   End
   Begin VB.TextBox txtOffice 
      Height          =   285
      Left            =   2520
      TabIndex        =   12
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox txtChild 
      Height          =   285
      Left            =   2520
      TabIndex        =   11
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox txtFav 
      Height          =   285
      Left            =   2520
      TabIndex        =   10
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox txtNum 
      Height          =   285
      Left            =   2520
      TabIndex        =   9
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   2520
      TabIndex        =   8
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label9 
      Caption         =   "Your first name:"
      Height          =   255
      Left            =   960
      TabIndex        =   7
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Television charachter:"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Label Label7 
      Caption         =   "The name of your first pet:"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Label Label6 
      Caption         =   "Brand name:"
      Height          =   255
      Left            =   1200
      TabIndex        =   4
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Office supply:"
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Memorable childhood item:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Your favorite name:"
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Pick a number."
      Height          =   255
      Left            =   960
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "frmBaby"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Title:  VaB-Libs
'Form:  Title Screen
'Author:  Jed Anderson
'Date Written: March 20th, 2008
'Objective:  To provide a Baby Nameing scenario for a MadLib

Option Explicit



Private Sub cmdMakeBaby_Click()

'Declaration of all my variables for this form.
Dim Name As String
Dim Labor As Integer
Dim Favorite As String
Dim Office As String
Dim TV As String
Dim Childhood As String
Dim Parent As String
Dim Pet As String
Dim Gender As String
Dim Brand As String
Dim Pronoun As String

'This input box asked the user if they would like to be the mother
'or father within the story and uses the input to determine gender.

Parent = InputBox("Are you the Mother or the Father?", "Parent")
If Parent = LCase("mother") Then
    Gender = "husband"
Else
    Gender = "wife"
End If

'Depending upon the answer for the question will determine which
'pronoun will be used.

If Parent = LCase("mother") Then
    Pronoun = "he"
Else
    Pronoun = "she"
End If

'Here is the labling of all my text boxes from the form.
Name = txtName.Text
Labor = txtNum.Text
Favorite = txtFav.Text
Office = txtOffice.Text
TV = txtTV.Text
Childhood = txtChild.Text
Pet = txtPet.Text
Brand = txtBrand.Text

'This is how I decided to print the stroy.  It looks messy and
'unecessary, but the other ways I tried were too buggy.
picResults.Cls
picResults.Print "The baby had finally arrived after "; Labor; " hours of intense labor.  As you hold your beautiful "
picResults.Print "baby in your arms, you begin to discuss the name of the child.  'What do you think we should name it?'"
picResults.Print "you ask aloud.  'I was hoping the name could be "; Favorite; ". you say in reply to yourself.  'That is "
picResults.Print "the single ugliest name I've ever heard.' your "; Gender; " says.  'If you think for one instant that I'll "
picResults.Print "agree to that hideous thing you call a name, we're getting a divorce and I'll grant you supervised "
picResults.Print "visitation.  I mean, seriously.'  'Fine,' you say.  'What do you think about "; Childhood; "?'"
picResults.Print "'Who are you?' your "; Gender; " shouts back at you.  'Michael Jackson?  We can't name a child "
picResults.Print Childhood; ".  Think about what the kids at school will say; 'Hey "; Childhood
picResults.Print "wanna come over to my house for a sleep-over?'  You have got to be the most obtuse person on the "
picResults.Print "face of the Earth.'  'Alright,' you say.  'I've got it.  "; Brand; ".  Your "; Gender; " rolls "; Pronoun; " eyes."
picResults.Print "'That name was thought up in a board room by consultants and focus groups.  Try again genius.'"
picResults.Print "You sit back and ponder a little bit.  'How about "; Pet; "?' You immediately regret opening your mouth that time."
picResults.Print "'Why, good one honey.  From the very first day I found out we were pregnant, I also thought we should "
picResults.Print "condemn our child to a life in the porn industry.  "; Pet; ", good lord, the child won't even have to change "
picResults.Print "name when it begins dancing for dollar bills.  Have you got any more good ones lined "
picResults.Print "up?'  You feel stung by "; Pronoun; " harsh words but you keep trying.  'How about "; Office; "?'"
picResults.Print "'You're joking, right?  Are you thinking of names or are you just looking around the room and naming "
picResults.Print "the objects you see at random?  Are you looking at the "; Office; " over there or did you really want "
picResults.Print "to name our child "; Office; "?  Let's focus here!'  Darn it.  You can't seem to do anything right.  You "
picResults.Print "really put your mind to work now.  You've got to impress "; Pronoun; " with this one or else "; Pronoun; " may not let "
picResults.Print "you pick any more.  'Ooooh, how about "; TV; "?'  'That's it,' your "
picResults.Print "spouse says.  'I'm taking away your TV privileges and throwing out all of your movies!  Your mind has "
picResults.Print "clearly been warped and damaged by the radiation.'  'I give up then!' you say.  'I'm tired of your "
picResults.Print "insults!  You don't like "; Favorite; ", "; Brand; ", "; TV; ", or "; Name; "  Your "; Gender; " looks thoughtful for a moment.  "
picResults.Print Name; ".  I like that.  Maybe you're OK after all.'"



End Sub

Private Sub cmdMain_Click()
'This button brings you back to the title screen.
frmBaby.Hide
frmTitle.Show


End Sub

