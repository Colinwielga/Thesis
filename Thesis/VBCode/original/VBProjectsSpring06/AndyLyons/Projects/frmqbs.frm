VERSION 5.00
Begin VB.Form frmqbs 
   BackColor       =   &H000000FF&
   Caption         =   "Quarterbacks"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10875
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   10875
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdreturn 
      Caption         =   "Return to Offense Profiles"
      Height          =   855
      Left            =   4080
      TabIndex        =   1
      Top             =   3120
      Width           =   1815
   End
   Begin VB.PictureBox picDisplay 
      Height          =   3975
      Left            =   600
      ScaleHeight     =   3915
      ScaleWidth      =   8475
      TabIndex        =   0
      Top             =   4080
      Width           =   8535
   End
   Begin VB.Label lblqbs 
      BackColor       =   &H0000FFFF&
      Caption         =   "Quarterbacks"
      BeginProperty Font 
         Name            =   "Broadway"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   7
      Top             =   2640
      Width           =   3015
   End
   Begin VB.Label lblcroyle 
      Caption         =   "Brodie Croyle"
      Height          =   255
      Left            =   7560
      TabIndex        =   6
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label lbljacobs 
      Caption         =   "Omar Jacobs"
      Height          =   255
      Left            =   6000
      TabIndex        =   5
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label lblcutler 
      Caption         =   "Jay Cutler"
      Height          =   255
      Left            =   4680
      TabIndex        =   4
      Top             =   240
      Width           =   735
   End
   Begin VB.Label lblyoung 
      Caption         =   "Vince Young"
      Height          =   255
      Left            =   3000
      TabIndex        =   3
      Top             =   240
      Width           =   975
   End
   Begin VB.Label lblleinhart 
      Caption         =   "Matt Leinart"
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   240
      Width           =   855
   End
   Begin VB.Image imgcroyle 
      Height          =   1590
      Left            =   7560
      Picture         =   "frmqbs.frx":0000
      Top             =   600
      Width           =   1170
   End
   Begin VB.Image imgomar 
      Height          =   1530
      Left            =   6000
      Picture         =   "frmqbs.frx":61FA
      Top             =   600
      Width           =   1185
   End
   Begin VB.Image imgcutler 
      Height          =   1860
      Left            =   4320
      Picture         =   "frmqbs.frx":C1DC
      Top             =   600
      Width           =   1485
   End
   Begin VB.Image imgyoung 
      Height          =   1980
      Left            =   2880
      Picture         =   "frmqbs.frx":1536E
      Top             =   600
      Width           =   1260
   End
   Begin VB.Image imgleinhart 
      Height          =   1905
      Left            =   840
      Picture         =   "frmqbs.frx":1D5A0
      Top             =   600
      Width           =   1350
   End
End
Attribute VB_Name = "frmqbs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'2006 NFL Draft Simulator (Draft.vbp)
'frmquarterbacks(frmquarterbacks.frm)
'Andy Lyons
'March 24, 2006
'The purpose of this form is to show the user the top quarterbacks that are eligible for the 2006 draft. By clicking on thier image, it allows you to read their profile.
'this button returns user to main menu
Private Sub cmdreturn_Click()
    frmoffpositions.Show
    frmqbs.Hide
End Sub

Private Sub imgcroyle_Click()
    picDisplay.Cls
    picDisplay.Print "Height 6'2""", "Weight 205"
    picDisplay.Print "Positives: Strong arm, nice touch. Quick release that is enhanced by exceptionally fast throwing motion."; Tab(11); "Reads defenses well. Makes good decisions throwing from the pocket. Leadership."
    picDisplay.Print "Negatives:  At 6-2, faces a challenge seeing over linemen, and 205-pound frame could make him susceptible"; Tab(11); "to injuries. NFL stock already has been hurt by injury-filled collegiate career. Accuracy and touch on shorter"; Tab(11); "routes. Pocket awareness and ability to avoid pressure. "
End Sub

Private Sub imgcutler_Click()
    picDisplay.Cls
    picDisplay.Print "Height 6'2""", "Weight 223"
    picDisplay.Print "Positives: Strong arm and overall strength emanating from a solid frame. Makes every kind of throw, but is"; Tab(11); "especially impressive on deep outs and squeezing the ball through small openings. Intelligence, patience in"; Tab(11); "the pocket, and progression reads. Took advantage of having stage all to himself during Scouting Combine"; Tab(11); "workouts, in which Leinart and Young did not participate, with mostly impressive performance. "
    picDisplay.Print "Negatives: Occasional overconfidence with arm strength can lead to gambles that backfire. Has some work to"; Tab(11); "do on overall mechanics and getting rid of the ball quicker. Wasn't tremendously accurate during Combine"; Tab(11); "throwing drills."
End Sub

Private Sub imgleinhart_Click()
    picDisplay.Cls
    picDisplay.Print "Height 6'4""", "Weight 224"
    picDisplay.Print "Positives: Perfect size for an NFL quarterback because he is tall enough to easily see over linemen and has a"; Tab(11); "chance to hold up to physical punishment resulting from poor protection and rookie mistakes in picking up"; Tab(11); "the blitz, etc. Large hands and long arms are also plusses. Very good accuracy, especially on touch"; Tab(11); "passes. Excels at reading coverages and quickly finding open receiver. Better footwork than one might"; Tab(11); "expect for his large frame, and is quick to set up for throws. Strong leadership; doesn't rattle easily vs."; Tab(11); "pressure or when faced with adversity. "
    picDisplay.Print "Negatives: Arm strength unspectacular, but good enough that it shouldn't prove a major problem. As a lefty, will"; Tab(11); "force team that drafts him to focus on putting best pass-blocking tackle on the right, which is uncommon."; Tab(11); "Long throwing motion could give defensive backs an edge on anticipating where his passes go. Poor"; Tab(11); "mobility, but is a talented enough passer to overcome that and be the first quarterback selected. "

End Sub

Private Sub imgomar_Click()
    picDisplay.Cls
    picDisplay.Print "Height 6'4""", "Weight 224"
    picDisplay.Print "Positives: Size, strength, and athleticism. Strong arm, large hands. Puts plenty of zip on deep outs and is able"; Tab(11); "to get the ball into tight spaces. Wonderful accuracy and touch on short and intermediate throws. "
    picDisplay.Print "Negatives: Tends to pull down the ball and run too soon; needs to become more comfortable working from the pocket."; Tab(12); "Throwing mechanics require work. Field vision, ability to read defenses. "
End Sub

Private Sub imgyoung_Click()
    picDisplay.Cls
    picDisplay.Print "Height 6'5""", "Weight 230"
    picDisplay.Print "Positives: Exceptional combination of size and off-the-charts athleticism make him an extremely rare talent."; Tab(11); "Outstanding speed and mobility make him a constant threat to run, and will keep defenses occupied with"; Tab(11); "mainly trying to minimize the damage he can do with his feet. Quick enough to avoid pressure in the"; Tab(11); "pocket and keep a play alive, but strong enough to break tackles when he does run. Showed tremendous"; Tab(11); "poise in leading the Longhorns and scoring his winning touchdown run for the national championship in the"; Tab(11); "Rose Bowl, the biggest football stage this side of the Super Bowl."
    picDisplay.Print "Negatives: Stories of poor results on intelligence test at the Scouting Combine were overblown, but could prove"; Tab(11); "damaging to what once seemed like a surefire top-three (or top) pick. Sidearm delivery makes throws"; Tab(11); "vulnerable to being knocked down. Overall passing mechanics are raw and need work. Patience and"; Tab(11); "decision-making in the pocket. "

End Sub
