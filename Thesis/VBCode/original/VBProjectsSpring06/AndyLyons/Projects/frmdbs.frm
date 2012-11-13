VERSION 5.00
Begin VB.Form frmdbs 
   BackColor       =   &H000040C0&
   Caption         =   "Defensive Backs"
   ClientHeight    =   10335
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10785
   LinkTopic       =   "Form1"
   ScaleHeight     =   10335
   ScaleWidth      =   10785
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdreturn 
      Caption         =   "Return to Defensive Players"
      Height          =   855
      Left            =   4800
      TabIndex        =   1
      Top             =   3000
      Width           =   2175
   End
   Begin VB.PictureBox picDisplay 
      Height          =   4935
      Left            =   840
      ScaleHeight     =   4875
      ScaleWidth      =   9195
      TabIndex        =   0
      Top             =   3960
      Width           =   9255
   End
   Begin VB.Label lblphillips 
      Caption         =   "Anwar Phillips"
      Height          =   255
      Left            =   8520
      TabIndex        =   6
      Top             =   360
      Width           =   1455
   End
   Begin VB.Image imgphillips 
      Height          =   1665
      Left            =   8640
      Picture         =   "frmdbs.frx":0000
      Top             =   720
      Width           =   1320
   End
   Begin VB.Label lblashton 
      Caption         =   "Ashton Youbty"
      Height          =   255
      Left            =   6840
      TabIndex        =   5
      Top             =   360
      Width           =   1575
   End
   Begin VB.Image imgashton 
      Height          =   1545
      Left            =   6840
      Picture         =   "frmdbs.frx":72BA
      Top             =   720
      Width           =   1470
   End
   Begin VB.Label lblwilliams 
      Caption         =   "Jimmy Williams"
      Height          =   255
      Left            =   4920
      TabIndex        =   4
      Top             =   360
      Width           =   1695
   End
   Begin VB.Image imgwilliams 
      Height          =   1800
      Left            =   4920
      Picture         =   "frmdbs.frx":EA14
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label lblcromar 
      Caption         =   "Antonio Cromartie"
      Height          =   255
      Left            =   2880
      TabIndex        =   3
      Top             =   360
      Width           =   1455
   End
   Begin VB.Image imgcromartio 
      Height          =   1230
      Left            =   2640
      Picture         =   "frmdbs.frx":189B6
      Top             =   720
      Width           =   1860
   End
   Begin VB.Label lblhill 
      Caption         =   "Tye Hill"
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   360
      Width           =   735
   End
   Begin VB.Image imghill 
      Height          =   1620
      Left            =   1080
      Picture         =   "frmdbs.frx":20120
      Top             =   720
      Width           =   1305
   End
End
Attribute VB_Name = "frmdbs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'2006 NFL Draft Simulator (Draft.vbp)
'frmdbs(frmdbs.frm)
'Andy Lyons
'March 24, 2006
'This form allows the user to view Defensive Ends eligible for the 2006 draft. By clicking on their picture, it uploads the profile.
'returns user to main menu
Private Sub cmdreturn_Click()
    frmdefpositions.Show
    frmdbs.Hide
End Sub

Private Sub imgashton_Click()
    picDisplay.Cls
    picDisplay.Print "Height 5'11""", "Weight 189"
    picDisplay.Print "Positives:  Good size. Natural playmaker. Excellent speed and athleticism. Good intelligence and instincts. Change-of-direction"; Tab(11); "ability. Solid in run support. "
    picDisplay.Print "Negatives: Lacks bulk and strength. … Needs to develop better play-recognition skills, which should happen as he gains "; Tab(12); "experience and coaching. "
End Sub

Private Sub imgcromartio_Click()
    picDisplay.Cls
    picDisplay.Print "Height 6'2""", "Weight 203"
    picDisplay.Print "Positives: Great size. Outstanding speed, athleticism, and leaping ability. Superb instincts. Long arms and strong hands."; Tab(11); "Provides strong run support. Versatility with background as a receiver and kick-return skills."
    picDisplay.Print "Negatives:  Missed the 2005 season after undergoing major knee surgery, although he ran well in his individual workout for"; Tab(12); "scouts. Needs to work on anticipation of pass routes. "

End Sub

Private Sub imghill_Click()
    picDisplay.Cls
    picDisplay.Print "Height 5'9""", "Weight 185"
    picDisplay.Print "Positives: Tremendous speed. Good instincts. Excels in bump-and-run coverage. Handles himself well in man-to-man and zone"; Tab(11); "coverage. Tough, physical player who is aggressive against the run. Outstanding leaping ability."
    picDisplay.Print "Negatives: At 5-foot-9 and 185 pounds, he lacks ideal size and doesn't have much bulk. Must work on improving his ability to"; Tab(12); "read the eyes of the quarterback and overall technique. "

End Sub

Private Sub imgphillips_Click()
    picDisplay.Cls
    picDisplay.Print "Height 5'11""", "Weight 190"
    picDisplay.Print "Positives: Good size and long arms. Excellent leaping ability and ball skills. Strength and aggressiveness allow him to"; Tab(11); "hold up well against larger receivers. Solid in run support. "
    picDisplay.Print "Negatives:  Must work on play-recognition skills and get out of the habit of focusing too much on the quarterback. Needs"; Tab(12); "to develop better tackling skills in the open field. "
End Sub

Private Sub imgwilliams_Click()
    picDisplay.Cls
    picDisplay.Print "Height 6'2""", "Weight 216"
    picDisplay.Print "Positives: Excellent size. Outstanding speed and change-of-direction skills. Physical player with considerable strength."; Tab(11); "Good hands and a threat to go the distance with every interception. Could excel at safety, as well as at cornerback. "
    picDisplay.Print "Negatives: Poor individual workout for scouts leaves lingering negative impression given that he opted not to work out at"; Tab(12); "the Scouting Combine. Must become more focused on receiver he covers rather than devoting too much attention"; Tab(12); "to the quarterback."
End Sub
