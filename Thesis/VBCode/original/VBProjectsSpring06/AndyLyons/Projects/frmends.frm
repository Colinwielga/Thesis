VERSION 5.00
Begin VB.Form frmends 
   BackColor       =   &H0000FF00&
   Caption         =   "Defensive Ends"
   ClientHeight    =   8670
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10785
   LinkTopic       =   "Form1"
   ScaleHeight     =   8670
   ScaleWidth      =   10785
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdreturn 
      BackColor       =   &H0000FF00&
      Caption         =   "Return to Defenisve Positions"
      Height          =   615
      Left            =   4680
      TabIndex        =   7
      Top             =   4320
      Width           =   1575
   End
   Begin VB.PictureBox picDisplay 
      Height          =   3375
      Left            =   480
      ScaleHeight     =   3315
      ScaleWidth      =   9435
      TabIndex        =   6
      Top             =   5040
      Width           =   9495
   End
   Begin VB.Label lbldefends 
      Caption         =   "DEFENSIVE ENDS"
      BeginProperty Font 
         Name            =   "Juice ITC"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3840
      TabIndex        =   5
      Top             =   3240
      Width           =   3375
   End
   Begin VB.Label lbllawson 
      Caption         =   "Manny Lawson"
      Height          =   255
      Left            =   8520
      TabIndex        =   4
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label lbltapp 
      Caption         =   "Darryl Tapp"
      Height          =   255
      Left            =   6840
      TabIndex        =   3
      Top             =   720
      Width           =   855
   End
   Begin VB.Label lblhali 
      Caption         =   "Tamba Hali"
      Height          =   255
      Left            =   4680
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lblmathias 
      Caption         =   "Mathias Kiwanuka"
      Height          =   495
      Left            =   2640
      TabIndex        =   1
      Top             =   720
      Width           =   855
   End
   Begin VB.Label lblwilliams 
      Caption         =   "Mario Williams"
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   1095
   End
   Begin VB.Image imglawson 
      Height          =   1830
      Left            =   8520
      Picture         =   "frmends.frx":0000
      Top             =   1200
      Width           =   1275
   End
   Begin VB.Image imgtapp 
      Height          =   1665
      Left            =   6480
      Picture         =   "frmends.frx":7A42
      Top             =   1200
      Width           =   1665
   End
   Begin VB.Image imghali 
      Height          =   1740
      Left            =   4440
      Picture         =   "frmends.frx":10C34
      Top             =   1200
      Width           =   1485
   End
   Begin VB.Image imgmathias 
      Height          =   1605
      Left            =   2520
      Picture         =   "frmends.frx":19466
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Image imgwilliams 
      Height          =   2055
      Left            =   480
      Picture         =   "frmends.frx":1FAA4
      Top             =   960
      Width           =   1440
   End
End
Attribute VB_Name = "frmends"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'2006 NFL Draft Simulator (Draft.vbp)
'frmends(frmends.frm)
'Andy Lyons
'March 24, 2006
'This form allows the user to view Defensive Ends eligible for the 2006 draft. By clicking on their picture, it uploads the profile.
'returns user to main menu
Private Sub cmdreturn_Click()
    frmdefpositions.Show
    frmends.Hide
End Sub

Private Sub imghali_Click()
    picDisplay.Cls
    picDisplay.Print "Height 6'3""", "Weight 275"
    picDisplay.Print "Positives: Explosiveness off the snap. Battles non-stop to the ball. Good pass-rushing moves. Upper-body strength, which allows"; Tab(11); "him to escape blocks and hold his own at the point of attack. Shows great ball instincts. Impressive performance in the"; Tab(11); "Senior Bowl. "
    picDisplay.Print "Negatives: At 6-3, he does not have ideal height for the position. Must develop better use of his hands to get separation"; Tab(12); "from blockers. "

End Sub

Private Sub imglawson_Click()
    picDisplay.Cls
    picDisplay.Print "Height 6'6""", "Weight 245"
    picDisplay.Print "Positives: Excellent athleticism and speed. Consistently zips around offensive tackles and has outstanding ability to close on the ball."; Tab(11); "Impressive strength. Great instincts. Works at building a repertoire of pass-rush moves. "
    picDisplay.Print "Negatives: Smallish frame, although has the skills that could allow him to be an outside linebacker in 3-4 alignment. Lower-body"; Tab(12); "strength. Must improve his leverage."
End Sub

Private Sub imgmathias_Click()
    picDisplay.Cls
    picDisplay.Print "Height 6'5""", "Weight 266"
    picDisplay.Print "Positives: Good size and speed. A former basketball player, he has plenty of athleticism. Nice variety of moves. Constant hustle allows"; Tab(11); "him to often make plays in pursuit. Shows excellent instincts that he enhances with considerable study that allows him to"; Tab(11); "make accurate pre-snap reads of offensive linemen. Solid containment on outside runs. "
    picDisplay.Print "Negatives: Tendency to play too tall and allow blockers to get under him. Needs to work on utilizing more of his quickness to get"; Tab(12); "a better jump off the ball. "

End Sub

Private Sub imgtapp_Click()
    picDisplay.Cls
    picDisplay.Print "Height 6'1""", "Weight 252"
    picDisplay.Print "Positives: Outstanding strength that helps him to consistently get penetration. Plays with good leverage. Physical player who shows"; Tab(11); "great intensity. Superb instincts help him to often make plays in the backfield."
    picDisplay.Print "Negatives: At 6-1 and 252 pounds, doesn't have ideal size for the position. Limited athleticism and speed. "

End Sub

Private Sub imgwilliams_Click()
    picDisplay.Cls
    picDisplay.Print "Height 6'7""", "Weight 295"
    picDisplay.Print "Positives: Great size. Superior speed and athleticism. Excellent anticipation of the snap and quickness off the ball allow him to win"; Tab(11); "most battles on his first step. Outstanding closing speed. Excellent footwork and change-of-direction skills. Superb ball"; Tab(11); "instincts. Shows equal dominance rushing the passer or stopping the run. Big hitter."
    picDisplay.Print "Negatives: Despite tremendously long arms, needs to work at getting better separation from blockers. Must improve technique rather"; Tab(12); "than mostly relying on running around blockers. "

End Sub
