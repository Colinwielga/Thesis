VERSION 5.00
Begin VB.Form frmWarLetter 
   Caption         =   "Form1"
   ClientHeight    =   9435
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11220
   LinkTopic       =   "Form1"
   Picture         =   "frmWarLetter.frx":0000
   ScaleHeight     =   9435
   ScaleWidth      =   11220
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Main Menu"
      Height          =   735
      Left            =   8400
      TabIndex        =   52
      Top             =   240
      Width           =   975
   End
   Begin VB.PictureBox picOutput 
      Height          =   7815
      Left            =   3480
      ScaleHeight     =   7755
      ScaleWidth      =   6915
      TabIndex        =   51
      Top             =   1320
      Width           =   6975
   End
   Begin VB.CommandButton cmdMakeLetter 
      Caption         =   "Generate VaB-Lib!"
      Height          =   735
      Left            =   5040
      TabIndex        =   50
      Top             =   240
      Width           =   1815
   End
   Begin VB.TextBox txtNoun9 
      Height          =   285
      Left            =   1800
      TabIndex        =   49
      Top             =   8520
      Width           =   1095
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1800
      TabIndex        =   48
      Top             =   8880
      Width           =   1335
   End
   Begin VB.TextBox txtNoun4 
      Height          =   285
      Left            =   1800
      TabIndex        =   47
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox txtVerb1 
      Height          =   285
      Left            =   1800
      TabIndex        =   46
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox txtAnimal 
      Height          =   285
      Left            =   1800
      TabIndex        =   45
      Top             =   3480
      Width           =   1335
   End
   Begin VB.TextBox txtVerb2 
      Height          =   285
      Left            =   1800
      TabIndex        =   44
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox txtNoun5 
      Height          =   285
      Left            =   1800
      TabIndex        =   43
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox txtNoun6 
      Height          =   285
      Left            =   1800
      TabIndex        =   42
      Top             =   4560
      Width           =   1335
   End
   Begin VB.TextBox txtVerb3 
      Height          =   285
      Left            =   1800
      TabIndex        =   41
      Top             =   4920
      Width           =   1335
   End
   Begin VB.TextBox txtVerb4 
      Height          =   285
      Left            =   1800
      TabIndex        =   40
      Top             =   5280
      Width           =   1215
   End
   Begin VB.TextBox txtGod 
      Height          =   285
      Left            =   1800
      TabIndex        =   39
      Top             =   5640
      Width           =   1455
   End
   Begin VB.TextBox txtCartoon 
      Height          =   285
      Left            =   1800
      TabIndex        =   38
      Top             =   6000
      Width           =   1455
   End
   Begin VB.TextBox txtJob 
      Height          =   285
      Left            =   1800
      TabIndex        =   37
      Top             =   6360
      Width           =   1455
   End
   Begin VB.TextBox txtNoun7 
      Height          =   285
      Left            =   1800
      TabIndex        =   36
      Top             =   6720
      Width           =   1215
   End
   Begin VB.TextBox txtVerb5 
      Height          =   285
      Left            =   1800
      TabIndex        =   35
      Top             =   7080
      Width           =   1095
   End
   Begin VB.TextBox txtNoun8 
      Height          =   285
      Left            =   1800
      TabIndex        =   34
      Top             =   7440
      Width           =   1215
   End
   Begin VB.TextBox txtFeeling 
      Height          =   285
      Left            =   1800
      TabIndex        =   33
      Top             =   7800
      Width           =   1455
   End
   Begin VB.TextBox txtVerb6 
      Height          =   285
      Left            =   1800
      TabIndex        =   32
      Top             =   8160
      Width           =   1455
   End
   Begin VB.TextBox txtNoun2 
      Height          =   285
      Left            =   1800
      TabIndex        =   31
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox txtNoun3 
      Height          =   285
      Left            =   1800
      TabIndex        =   30
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox txtAdj2 
      Height          =   285
      Left            =   1800
      TabIndex        =   29
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox txtAdj1 
      Height          =   285
      Left            =   1800
      TabIndex        =   28
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtNoun1 
      Height          =   285
      Left            =   1800
      TabIndex        =   27
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox txtFlower 
      Height          =   285
      Left            =   1800
      TabIndex        =   26
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox txtNum 
      Height          =   285
      Left            =   1800
      TabIndex        =   25
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label25 
      Caption         =   "Name"
      Height          =   255
      Left            =   960
      TabIndex        =   24
      Top             =   8880
      Width           =   615
   End
   Begin VB.Label Label24 
      Caption         =   "Noun"
      Height          =   255
      Left            =   960
      TabIndex        =   23
      Top             =   8520
      Width           =   615
   End
   Begin VB.Label Label23 
      Caption         =   "Verb + ing"
      Height          =   255
      Left            =   600
      TabIndex        =   22
      Top             =   8160
      Width           =   975
   End
   Begin VB.Label Label22 
      Caption         =   "Emotion + ness"
      Height          =   255
      Left            =   360
      TabIndex        =   21
      Top             =   7800
      Width           =   1215
   End
   Begin VB.Label Label21 
      Caption         =   "Noun"
      Height          =   255
      Left            =   1080
      TabIndex        =   20
      Top             =   7440
      Width           =   495
   End
   Begin VB.Label Label20 
      Caption         =   "Verb"
      Height          =   255
      Left            =   1080
      TabIndex        =   19
      Top             =   7080
      Width           =   495
   End
   Begin VB.Label Label19 
      Caption         =   "Noun"
      Height          =   255
      Left            =   1080
      TabIndex        =   18
      Top             =   6720
      Width           =   495
   End
   Begin VB.Label Label18 
      Caption         =   "Occupation"
      Height          =   255
      Left            =   600
      TabIndex        =   17
      Top             =   6360
      Width           =   975
   End
   Begin VB.Label Label17 
      Caption         =   "Cartoon Character"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   6000
      Width           =   1455
   End
   Begin VB.Label Label16 
      Caption         =   "Deity"
      Height          =   255
      Left            =   1080
      TabIndex        =   15
      Top             =   5640
      Width           =   495
   End
   Begin VB.Label Label15 
      Caption         =   "Verb (past tense)"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Label Label14 
      Caption         =   "Verb (past tense)"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Label Label13 
      Caption         =   "Noun (plural)"
      Height          =   255
      Left            =   600
      TabIndex        =   12
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label Label12 
      Caption         =   "Noun"
      Height          =   255
      Left            =   1080
      TabIndex        =   11
      Top             =   4200
      Width           =   495
   End
   Begin VB.Label Label11 
      Caption         =   "Verb"
      Height          =   255
      Left            =   1080
      TabIndex        =   10
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label Label10 
      Caption         =   "Animal"
      Height          =   255
      Left            =   960
      TabIndex        =   9
      Top             =   3480
      Width           =   615
   End
   Begin VB.Label Label9 
      Caption         =   "Verb"
      Height          =   255
      Left            =   1080
      TabIndex        =   8
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label Label8 
      Caption         =   "Noun (proper)"
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "Adjective"
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "Noun"
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "Noun (plural)"
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Adjective"
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Noun (plural)"
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Flower"
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Number"
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmWarLetter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Title:  VaB-Libs
'Form:  Title Screen
'Author:  Jed Anderson
'Date Written: March 20th, 2008
'Objective:  To provide a War Letter formula for a MadLib

Option Explicit

Private Sub cmdMakeLetter_Click()

'This turned out to be a big mistake of a project.
Dim Num As Integer
Dim Flower As String
Dim Noun1 As String
Dim Adj1 As String
Dim Noun2 As String
Dim Noun3 As String
Dim Adj2 As String
Dim Noun4 As String
Dim Verb1 As String
Dim Animal As String
Dim Verb2 As String
Dim Noun5 As String
Dim Noun6 As String
Dim Verb3 As String
Dim Verb4 As String
Dim Verb5 As String
Dim Noun7 As String
Dim God As String
Dim Cartoon As String
Dim Job As String
Dim Noun8 As String
Dim Verb6 As String
Dim Noun9 As String
Dim Feeling As String
Dim Noun10 As String
Dim Name As String

'Ouch.  I never want to declare another variable again as long as I live.
Num = txtNum.Text
Flower = txtFlower.Text
Adj1 = txtAdj1.Text
Adj2 = txtAdj2.Text
Noun1 = txtNoun1.Text
Noun2 = txtNoun2.Text
Noun3 = txtNoun3.Text
Noun4 = txtNoun4.Text
Noun5 = txtNoun5.Text
Noun6 = txtNoun6.Text
Noun7 = txtNoun7.Text
Noun8 = txtNoun8.Text
Noun9 = txtNoun9.Text
Verb1 = txtVerb1.Text
Verb2 = txtVerb2.Text
Verb3 = txtVerb3.Text
Verb4 = txtVerb4.Text
Verb5 = txtVerb5.Text
Verb6 = txtVerb6.Text
Animal = txtAnimal.Text
God = txtGod.Text
Cartoon = txtCartoon.Text
Job = txtJob.Text
Feeling = txtFeeling.Text
Name = txtName.Text

'Here goes the story printing again--the only way I know how.
picOutput.Print "My Dearest "; Name; ","
picOutput.Print ""
picOutput.Print "The platoon has been in the trenches for "; Num; " days now and I do not wish to cause you "
picOutput.Print "alarm, but I fear for my "; Noun1; ".  "; Flower; ", that's the nickname for our platoon leader, does not"
picOutput.Print "allow me to give details in my writing that my be useful to the enemy if intercepted, but rest"
picOutput.Print "assured I will "; Verb1; " with you again, my friend."
picOutput.Print ""
picOutput.Print "We've all been eating "; Noun2; " for every meal.  I grow "; Adj1; " of the taste as the weeks"
picOutput.Print "go by.  However, I find that if I mix up the "; Noun3; " with handfuls of "; Noun4; " and close"
picOutput.Print "my eyes and make believe, it tastes a little like potatoes mixed with dirt.  It is like "; Adj2
picOutput.Print Noun5; " in my mouth."
picOutput.Print ""
picOutput.Print "I miss "; Noun6; " and the morning sunshine that kissed the mountain tops at home.  I long"
picOutput.Print "to hear the welcoming "; Verb1; " of the "; Animal; " and the "; Verb2; " of "; Noun7; " hooves.  The desire to cast"
picOutput.Print "off these "; Noun8; " and put aside my bayonet is strong indeed."
picOutput.Print ""
picOutput.Print "Yesterday I was slotted for patrol and was "; Verb3; " at while I stopped in a"
picOutput.Print "wooded grove to rest.  As I "; Verb4; " behind a "; Noun9; " for cover as I prayed to "; God
picOutput.Print "to save me.  When that didn't work I used the entire capacity of my mind to fantasize about"
picOutput.Print Cartoon; " dressed in a(n) "; Job; " uniform.  I think I have found a new religion"
picOutput.Print "my friend."
picOutput.Print ""
picOutput.Print "With each passing day I strive with every "; Noun9; " in my body to "; Verb5; ", but secretly I yearn for "
picOutput.Print Noun5; " to take me away from this evil mess with swift wings.  I lie awake at night in"
picOutput.Print Feeling; " of what may be "; Verb5; " nearby in the darkness.  I sleep during the daylight "
picOutput.Print "while the droning "; Noun7; " and mortar rounds hypnotize me into rest.  I know not the meaning "
picOutput.Print "of safety."
picOutput.Print ""
picOutput.Print "My advice to you is this:  Live everyday you are given for peace.  Take no action in haste or "
picOutput.Print "anger.  Instead focus every breath you take, every move you make towards the lofty goal of "
picOutput.Print "SAVING ME!"
picOutput.Print ""
picOutput.Print "Do not weep for me.  I will always be near."
picOutput.Print "Love,"
picOutput.Print ""
picOutput.Print Name



End Sub

Private Sub cmdMenu_Click()
'This button brings the user back to the main menu.
frmWarLetter.Hide
frmTitle.Show

End Sub

