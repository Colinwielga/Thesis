VERSION 5.00
Begin VB.Form frmGlobal 
   ClientHeight    =   10440
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12240
   LinkTopic       =   "Form1"
   Picture         =   "frmGobal.frx":0000
   ScaleHeight     =   10440
   ScaleWidth      =   12240
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtNoun6 
      Height          =   285
      Left            =   9120
      TabIndex        =   79
      Top             =   4320
      Width           =   1335
   End
   Begin VB.TextBox txtName3 
      Height          =   285
      Left            =   9120
      TabIndex        =   77
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Main Menu"
      Height          =   735
      Left            =   9960
      TabIndex        =   76
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton cmdGlobal 
      Caption         =   "Create VaB-Lib!"
      Height          =   735
      Left            =   9480
      TabIndex        =   75
      Top             =   4920
      Width           =   1695
   End
   Begin VB.PictureBox picStory 
      Height          =   10335
      Left            =   3120
      ScaleHeight     =   10275
      ScaleWidth      =   5835
      TabIndex        =   74
      Top             =   120
      Width           =   5895
   End
   Begin VB.TextBox txtNoun5 
      Height          =   285
      Left            =   9120
      TabIndex        =   73
      Top             =   3600
      Width           =   1335
   End
   Begin VB.TextBox txtBody3 
      Height          =   285
      Left            =   9120
      TabIndex        =   72
      Top             =   3240
      Width           =   1335
   End
   Begin VB.TextBox txtVerb6 
      Height          =   285
      Left            =   9120
      TabIndex        =   71
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox txtNoun4 
      Height          =   285
      Left            =   9120
      TabIndex        =   70
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox txtVerb5 
      Height          =   285
      Left            =   9120
      TabIndex        =   69
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox txtAdj5 
      Height          =   285
      Left            =   9120
      TabIndex        =   68
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox txtSong2 
      Height          =   285
      Left            =   9120
      TabIndex        =   67
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox txtNoun3 
      Height          =   285
      Left            =   9120
      TabIndex        =   66
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox txtAdj4 
      Height          =   285
      Left            =   9120
      TabIndex        =   65
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox txtAdj3 
      Height          =   285
      Left            =   9120
      TabIndex        =   64
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox txtNoun2 
      Height          =   285
      Left            =   1680
      TabIndex        =   63
      Top             =   9360
      Width           =   1335
   End
   Begin VB.TextBox txtMusic 
      Height          =   285
      Left            =   1680
      TabIndex        =   62
      Top             =   9720
      Width           =   1335
   End
   Begin VB.TextBox txtPlace3 
      Height          =   285
      Left            =   1680
      TabIndex        =   61
      Top             =   9000
      Width           =   1335
   End
   Begin VB.TextBox txtBody2 
      Height          =   285
      Left            =   1680
      TabIndex        =   60
      Top             =   8640
      Width           =   1335
   End
   Begin VB.TextBox txtVerb4 
      Height          =   285
      Left            =   1680
      TabIndex        =   59
      Top             =   8280
      Width           =   1335
   End
   Begin VB.TextBox txtFood 
      Height          =   285
      Left            =   1680
      TabIndex        =   58
      Top             =   7920
      Width           =   1335
   End
   Begin VB.TextBox txtAdj2 
      Height          =   285
      Left            =   1680
      TabIndex        =   57
      Top             =   7560
      Width           =   1335
   End
   Begin VB.TextBox txtName2 
      Height          =   285
      Left            =   1680
      TabIndex        =   56
      Top             =   7200
      Width           =   1335
   End
   Begin VB.TextBox txtVerb3 
      Height          =   285
      Left            =   1680
      TabIndex        =   55
      Top             =   6840
      Width           =   1335
   End
   Begin VB.TextBox txtVehicle2 
      Height          =   285
      Left            =   1680
      TabIndex        =   54
      Top             =   6480
      Width           =   1335
   End
   Begin VB.TextBox txtStructure2 
      Height          =   285
      Left            =   1680
      TabIndex        =   53
      Top             =   6120
      Width           =   1335
   End
   Begin VB.TextBox txtDrink 
      Height          =   285
      Left            =   1680
      TabIndex        =   52
      Top             =   5760
      Width           =   1335
   End
   Begin VB.TextBox txtName1 
      Height          =   285
      Left            =   1680
      TabIndex        =   51
      Top             =   5400
      Width           =   1335
   End
   Begin VB.TextBox txtBody1 
      Height          =   285
      Left            =   1680
      TabIndex        =   50
      Top             =   5040
      Width           =   1335
   End
   Begin VB.TextBox txtVerb2 
      Height          =   285
      Left            =   1680
      TabIndex        =   49
      Top             =   4680
      Width           =   1335
   End
   Begin VB.TextBox txtShoe 
      Height          =   285
      Left            =   1680
      TabIndex        =   48
      Top             =   4320
      Width           =   1335
   End
   Begin VB.TextBox txtSong1 
      Height          =   285
      Left            =   1680
      TabIndex        =   47
      Top             =   3960
      Width           =   1335
   End
   Begin VB.TextBox txtVerb1 
      Height          =   285
      Left            =   1680
      TabIndex        =   46
      Top             =   3600
      Width           =   1335
   End
   Begin VB.TextBox txtNation 
      Height          =   285
      Left            =   1680
      TabIndex        =   45
      Top             =   3240
      Width           =   1335
   End
   Begin VB.TextBox txtAnimal2 
      Height          =   285
      Left            =   1680
      TabIndex        =   44
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox txtVehicle1 
      Height          =   285
      Left            =   1680
      TabIndex        =   43
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox txtStructure1 
      Height          =   285
      Left            =   1680
      TabIndex        =   42
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox txtNoun1 
      Height          =   285
      Left            =   1680
      TabIndex        =   41
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox txtPlace2 
      Height          =   285
      Left            =   1680
      TabIndex        =   40
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox txtAdj1 
      Height          =   285
      Left            =   1680
      TabIndex        =   39
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox txtPlace1 
      Height          =   285
      Left            =   1680
      TabIndex        =   38
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox txtAnimal1 
      Height          =   285
      Left            =   1680
      TabIndex        =   37
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label38 
      Caption         =   "Noun"
      Height          =   255
      Left            =   10680
      TabIndex        =   80
      Top             =   4320
      Width           =   855
   End
   Begin VB.Label Name 
      Caption         =   "Name"
      Height          =   255
      Left            =   10680
      TabIndex        =   78
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label37 
      Caption         =   "Noun"
      Height          =   255
      Left            =   10680
      TabIndex        =   36
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label Label36 
      Caption         =   "Body Part"
      Height          =   255
      Left            =   10680
      TabIndex        =   35
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label Label35 
      Caption         =   "Verb"
      Height          =   255
      Left            =   10680
      TabIndex        =   34
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label Label34 
      Caption         =   "Noun"
      Height          =   255
      Left            =   10680
      TabIndex        =   33
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label33 
      Caption         =   "Verb (past tense)"
      Height          =   255
      Left            =   10680
      TabIndex        =   32
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label32 
      Caption         =   "Adjective"
      Height          =   255
      Left            =   10680
      TabIndex        =   31
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label31 
      Caption         =   "Song Title"
      Height          =   255
      Left            =   10680
      TabIndex        =   30
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label30 
      Caption         =   "Noun"
      Height          =   255
      Left            =   10680
      TabIndex        =   29
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label29 
      Caption         =   "Adjective"
      Height          =   255
      Left            =   10680
      TabIndex        =   28
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label28 
      Caption         =   "Adjective"
      Height          =   255
      Left            =   10680
      TabIndex        =   27
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label27 
      Caption         =   "Body Part"
      Height          =   255
      Left            =   720
      TabIndex        =   26
      Top             =   8640
      Width           =   735
   End
   Begin VB.Label Label26 
      Caption         =   "Noun"
      Height          =   255
      Left            =   960
      TabIndex        =   25
      Top             =   9360
      Width           =   495
   End
   Begin VB.Label Label25 
      Caption         =   "Musical Instument"
      Height          =   255
      Left            =   240
      TabIndex        =   24
      Top             =   9720
      Width           =   1335
   End
   Begin VB.Label Label24 
      Caption         =   "Place"
      Height          =   255
      Left            =   960
      TabIndex        =   23
      Top             =   9000
      Width           =   495
   End
   Begin VB.Label Label23 
      Caption         =   "Song Title"
      Height          =   255
      Left            =   600
      TabIndex        =   22
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label22 
      Caption         =   "Footwear"
      Height          =   255
      Left            =   720
      TabIndex        =   21
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label Label21 
      Caption         =   "Verb"
      Height          =   255
      Left            =   1080
      TabIndex        =   20
      Top             =   4680
      Width           =   375
   End
   Begin VB.Label Label20 
      Caption         =   "Body Part"
      Height          =   255
      Left            =   600
      TabIndex        =   19
      Top             =   5040
      Width           =   855
   End
   Begin VB.Label Label19 
      Caption         =   "Name"
      Height          =   255
      Left            =   960
      TabIndex        =   18
      Top             =   5400
      Width           =   495
   End
   Begin VB.Label Label18 
      Caption         =   "Liquid"
      Height          =   255
      Left            =   960
      TabIndex        =   17
      Top             =   5760
      Width           =   495
   End
   Begin VB.Label Label17 
      Caption         =   "Structure"
      Height          =   255
      Left            =   720
      TabIndex        =   16
      Top             =   6120
      Width           =   735
   End
   Begin VB.Label Label16 
      Caption         =   "Vehicle"
      Height          =   255
      Left            =   840
      TabIndex        =   15
      Top             =   6480
      Width           =   615
   End
   Begin VB.Label Label15 
      Caption         =   "Verb"
      Height          =   255
      Left            =   1080
      TabIndex        =   14
      Top             =   6840
      Width           =   375
   End
   Begin VB.Label Label14 
      Caption         =   "Name"
      Height          =   255
      Left            =   960
      TabIndex        =   13
      Top             =   7200
      Width           =   495
   End
   Begin VB.Label Label13 
      Caption         =   "Adjective"
      Height          =   255
      Left            =   720
      TabIndex        =   12
      Top             =   7560
      Width           =   735
   End
   Begin VB.Label Label12 
      Caption         =   "Food"
      Height          =   255
      Left            =   960
      TabIndex        =   11
      Top             =   7920
      Width           =   495
   End
   Begin VB.Label Label11 
      Caption         =   "Verb (past tense)"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   8280
      Width           =   1335
   End
   Begin VB.Label Label10 
      Caption         =   "Verb"
      Height          =   255
      Left            =   1080
      TabIndex        =   9
      Top             =   3600
      Width           =   375
   End
   Begin VB.Label Label9 
      Caption         =   "Nationality"
      Height          =   255
      Left            =   600
      TabIndex        =   8
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label Label8 
      Caption         =   "Animal (plural)"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "Vehicle"
      Height          =   255
      Left            =   840
      TabIndex        =   6
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label Label6 
      Caption         =   "Structure"
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Noun"
      Height          =   255
      Left            =   960
      TabIndex        =   4
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "Place"
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Adjective"
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Place"
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Animal (plural)"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "frmGlobal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Title:  VaB-Libs
'Form:  Title Screen
'Author:  Jed Anderson
'Date Written: March 20th, 2008
'Objective:  To provide a fictional parody on solving Global Warming formula for
'a MadLib

Option Explicit

Private Sub cmdGlobal_Click()
'Yikes!  Look as all those variables!
Dim Animal1 As String
Dim Place1 As String
Dim Adj1 As String
Dim Place2 As String
Dim Noun1 As String
Dim Structure1 As String
Dim Vehicle1 As String
Dim Animal2 As String
Dim Nation As String
Dim Verb1 As String
Dim Song1 As String
Dim Shoe As String
Dim Verb2 As String
Dim Body1 As String
Dim Name1 As String
Dim Liquid As String
Dim Structure2 As String
Dim Vehicle2 As String
Dim Verb3 As String
Dim Name2 As String
Dim Adj2 As String
Dim Food As String
Dim Verb4 As String
Dim Body2 As String
Dim Noun2 As String
Dim Music As String
Dim Adj3 As String
Dim Adj4 As String
Dim Noun3 As String
Dim Song2 As String
Dim Adj5 As String
Dim Verb5 As String
Dim Noun4 As String
Dim Verb6 As String
Dim Body3 As String
Dim Noun5 As String
Dim Name3 As String
Dim Noun6 As String

'There has got to be a better way...
Name1 = txtName1.Text
Name2 = txtName2.Text
Name3 = txtName3.Text
Noun1 = txtNoun1.Text
Noun2 = txtNoun2.Text
Noun3 = txtNoun3.Text
Noun4 = txtNoun4.Text
Noun5 = txtNoun5.Text
Noun6 = txtNoun6.Text
Body1 = txtBody1.Text
Body2 = txtBody2.Text
Body3 = txtBody3.Text
Verb1 = txtVerb1.Text
Verb2 = txtVerb2.Text
Verb3 = txtVerb3.Text
Verb4 = txtVerb4.Text
Verb5 = txtVerb5.Text
Verb6 = txtVerb6.Text
Adj1 = txtAdj1.Text
Adj2 = txtAdj2.Text
Adj3 = txtAdj3.Text
Adj4 = txtAdj4.Text
Adj5 = txtAdj5.Text
Place1 = txtPlace1.Text
Place2 = txtPlace2.Text
Song1 = txtSong1.Text
Song2 = txtSong2.Text
Animal1 = txtAnimal1.Text
Animal2 = txtAnimal2.Text
Structure1 = txtStructure1.Text
Structure2 = txtStructure2.Text
Vehicle1 = txtVehicle1.Text
Vehicle2 = txtVehicle2.Text
Nation = txtNation.Text
Shoe = txtShoe.Text
Liquid = txtDrink.Text
Food = txtFood.Text
Music = txtMusic.Text

'Once again my story printing format.
picStory.Print "It all started back when I was hunting "; Animal1; " in the "; Place1; "."; Animal1; ","
picStory.Print "as I'm sure you all know, tend to make their lairs where it's "; Adj1; ".  They"
picStory.Print "have a searing internal temperature which needs to be kept cool by the glaciers found in "
picStory.Print "the areas above "; Place2; ".  I had a theory of what was causing all this global warming"
picStory.Print "hub-bub and I was determined to remedy the problem by my own means."
picStory.Print ""
picStory.Print "After kayaking for miles through a thin glacial-stream, I came upon a(n) "; Noun1; " stationed"
picStory.Print "just outside the mouth of great "; Structure1; ".  After portaging my "; Vehicle1; ", I cautiously"
picStory.Print "approached the "; Structure1; " in case it may be inhabited by flesh-eating "; Animal2; ", "
picStory.Print "waving my arms in the anti-jinxing circles which "; Animal2; " feared beyond all "
picStory.Print "else.  I was also taught by an old "; Nation; " shaman to sing an old folk tune that would "
picStory.Print "lull them to sleep in case I was face to face with a(n) "; Animal2; " charge."
picStory.Print ""
picStory.Print "As I "; Verb1; " closer, singing the words to, "; Song1; ", the three-foot snow crunching "
picStory.Print "beneath my "; Shoe; ", a man emerged from the "; Structure1; " carrying a guitar.  We "
picStory.Print "finished the tune and "; Verb2; Body1; " in greeting.  His name was "; Name1; " and "
picStory.Print "he just so happened to be hunting "; Animal1; " to stop global warming."
picStory.Print ""
picStory.Print "We spent the next few hours in his "; Structure1; ", drinking "; Liquid; " and planning our first "
picStory.Print "move against the "; Animal1; " which dwelt just inside a nearby "; Structure2; ".  Just "
picStory.Print "as we were arguing about who was going to have to be the distraction, we heard the "
picStory.Print "sound of a "; Vehicle2; " engine.  We crawled from the "; Structure1; " and looked to the sky to see "
picStory.Print "a person "; Verb3; " down, headed right for our vicinity."
picStory.Print ""
picStory.Print "The person landed and told us that his name was "; Name2; ".  "; Name2; " explained to us about "
picStory.Print "the "; Animal1; ", which were to be the causes of global warming.  This person was a "
picStory.Print "big fan of Earth, as were "; Name1; " and I, so we quickly found unity in our quest to save "
picStory.Print "the world."
picStory.Print ""
picStory.Print "After a few more sips of "; Liquid; " and some nice, "; Adj2; Food; ", we heard a great roar "
picStory.Print "blast from inside the "; Structure2; ".  It "; Verb4; " the very foundation of our "; Body2; ""
picStory.Print "and sent a tremor resonating to the "; Place2; ".  The roof of "; Name1; "'s fine "
picStory.Print "home cracked and began to fall on our heads.  Quickly, we finished our "; Food; ", and "
picStory.Print "scrambled from the "; Structure1; ".  The roaring continued to belt from the "; Structure2;
picStory.Print "and we were very afraid.  Through the fear cut a fell sound on the air.  A clear, audible "
picStory.Print "event also came from "; Structure2; " and we looked at each other in amazement.  It was "
picStory.Print "the sound of a Noun2"
picStory.Print ""
picStory.Print "Without another word between us, "; Name1; " grabbed a(n) "; Music; " and we "
picStory.Print "three ran into the "; Structure2; ".  The "; Noun2; " grew louder and louder as "; Structure2
picStory.Print "grew "; Adj3; " and "; Adj4; ".  Around the final bend we saw the "; Noun3; "shining like "
picStory.Print "a(n) "; Noun4; ".  It called to us and told us its name was "; Name3; ", and explained that it was "
picStory.Print "here to destroy the "; Animal1; " that were causing all the global warming.  'The"
picStory.Print "only way to defeat them is with music,' he said, 'but I don't know how to play.'  This "
picStory.Print Animal1; " is about to conquer me, so get out whilst you still can!"
picStory.Print ""
picStory.Print "We three stepped forward to join the "; Noun4; "'s side and from us, "; Song2; " rang out from "; Name1; ""
picStory.Print "and his "; Music; ".  "; Name2; " and I sang along with "; Adj4; " voices raised.  "
picStory.Print "The "; Animal1; Verb5; " in agony, spewing "; Noun5; " in all directions.  The "
picStory.Print Animal1; "'s head began to "; Verb6; " and bulge forth in odd places.  We rang the final chord "
picStory.Print "of the song with everything we had and the "; Animal1; "'s "; Body3; " exploded into a "
picStory.Print "spray of "; Noun6; " and blood.  It was over…"


End Sub

Private Sub cmdMenu_Click()
'This button brings the user back to the main menu.
frmGlobal.Hide
frmTitle.Show

End Sub

