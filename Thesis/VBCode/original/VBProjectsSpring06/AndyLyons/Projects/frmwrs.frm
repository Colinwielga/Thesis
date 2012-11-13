VERSION 5.00
Begin VB.Form frmwrs 
   BackColor       =   &H00004000&
   Caption         =   "Wide Receivers"
   ClientHeight    =   8760
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11070
   LinkTopic       =   "Form1"
   ScaleHeight     =   8760
   ScaleWidth      =   11070
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picDisplay 
      Height          =   2655
      Left            =   600
      ScaleHeight     =   2595
      ScaleWidth      =   9195
      TabIndex        =   2
      Top             =   4680
      Width           =   9255
   End
   Begin VB.CommandButton cmdreturn 
      Caption         =   "Return to Offensive Players"
      Height          =   855
      Left            =   4320
      TabIndex        =   0
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Wide Receivers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      TabIndex        =   7
      Top             =   2760
      Width           =   2655
   End
   Begin VB.Label lblhagan 
      Caption         =   "Derek Hagan"
      Height          =   255
      Left            =   7560
      TabIndex        =   6
      Top             =   480
      Width           =   1215
   End
   Begin VB.Image imghagan 
      Height          =   1635
      Left            =   7560
      Picture         =   "frmwrs.frx":0000
      Top             =   720
      Width           =   1155
   End
   Begin VB.Label lblavant 
      Caption         =   "Jason Avant"
      Height          =   255
      Left            =   6000
      TabIndex        =   5
      Top             =   480
      Width           =   1455
   End
   Begin VB.Image imgavant 
      Height          =   1635
      Left            =   6000
      Picture         =   "frmwrs.frx":630A
      Top             =   720
      Width           =   1305
   End
   Begin VB.Label lblmoss 
      Caption         =   "Sinorice Moss"
      Height          =   255
      Left            =   4680
      TabIndex        =   4
      Top             =   480
      Width           =   1215
   End
   Begin VB.Image imgmoss 
      Height          =   1920
      Left            =   4680
      Picture         =   "frmwrs.frx":D3B4
      Top             =   720
      Width           =   1125
   End
   Begin VB.Label lbljackson 
      Caption         =   "Chad Jackson"
      Height          =   255
      Left            =   3360
      TabIndex        =   3
      Top             =   480
      Width           =   1215
   End
   Begin VB.Image imgjackson 
      Height          =   1740
      Left            =   3480
      Picture         =   "frmwrs.frx":145F6
      Top             =   720
      Width           =   930
   End
   Begin VB.Label lblholmes 
      Caption         =   "Santonio Holmes"
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   480
      Width           =   1575
   End
   Begin VB.Image imgholmes 
      Height          =   1575
      Left            =   1920
      Picture         =   "frmwrs.frx":19B68
      Top             =   720
      Width           =   1155
   End
End
Attribute VB_Name = "frmwrs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'2006 NFL Draft Simulator (Draft.vbp)
'frmwrs(frmwrs.frm)
'Andy Lyons
'March 24, 2006
'This form allows the user to view the Wide Receivers eligible for the 2006 NFL Draft. By clicking on their photo, it will upload their profile for the user to read.
'returns user to main menu
Private Sub cmdreturn_Click()
    frmoffpositions.Show
    frmwrs.Hide
End Sub

Private Sub imgavant_Click()
    picDisplay.Cls
    picDisplay.Print "Height 6'0""", "Weight 209 "
    picDisplay.Print "Positives: Highly productive college career that probably would have received greater attention had he not been overshadowed"; Tab(11); "by Braylon Edwards, whom the Browns made the third overall pick of the 2005 draft. Outstanding hands. Able to"; Tab(11); "make every catch an NFL receiver needs to make long, short, over the middle. Runs excellent routes and shows"; Tab(11); "considerable burst in and out of cuts. Good size and strength. Superior athleticism. Reads coverages well. Shows"; Tab(11); " tremendous determination to get open. Aggressive blocker. "
    picDisplay.Print "Negatives: Doesn't have top-level speed. Durability issues."
End Sub

Private Sub imghagan_Click()
    picDisplay.Cls
    picDisplay.Print "Height 6'3""", "Weight 200"
    picDisplay.Print "Positives: Intelligent route-runner with more savvy than one would expect from a young receiver when it comes to getting open"; Tab(11); "against zone and man-to-man coverage. Superb footwork and change-of-direction ability. Excels at recognizing"; Tab(11); "coverages. Strong leadership skills. "
    picDisplay.Print "Negatives: Tends to avoid going over the middle for tough catch. Doesn't have great speed, which limits ability to get separation."; Tab(11); "Needs to work on getting release from the line. "
End Sub

Private Sub imgholmes_Click()
    picDisplay.Cls
    picDisplay.Print "Height 5'10""", "Weight 198"
    picDisplay.Print "Positives: Dependable, sure-handed player who easily is the draft's best player at his position. Will make tough catches over"; Tab(11); "the middle and can grab passes thrown over his head in full stride. Runs crisp, precise routes. Gets good"; Tab(11); "separation once he is in his pattern. Outstanding footwork that allows him to make quick starts and stops. Has"; Tab(11); "explosiveness to quickly get upfield after the catch and be a legitimate threat to go the distance whenever the ball"; Tab(11); "is in his hands. "
    picDisplay.Print "Negatives: Must reduce number of passes that he catches with his body, rather than strictly with his hands. Smallish frame"; Tab(11); "creates some difficulties when it comes to releasing from the line of scrimmage. Needs to learn how to make better"; Tab(11); "use of hands to escape defenders trying to jam him. "
 
End Sub

Private Sub imgjackson_Click()
    picDisplay.Cls
    picDisplay.Print "Height 6'1""", "Weight 201"
    picDisplay.Print "Positives: Superb concentration on the ball. Makes difficult catches in traffic. Excels at running underneath routes, but has"; Tab(11); "speed and burst to turn them into long gains. Has size and strength to fight off jams at the line. Runs precise routes. "
    picDisplay.Print "Negatives: NFL teams don't know how effective he can be on deep routes because he didn't run many of them in college. Must"; Tab(11); "improve blocking technique and develop a more aggressive attitude toward making contact. "

End Sub

Private Sub imgmoss_Click()
    picDisplay.Cls
    picDisplay.Print "Height 5'7""", "Weight 183"
    picDisplay.Print "Positives: Great speed and a threat to score on every catch. Makes good adjustments to poorly thrown balls. Works particularly"; Tab(11); "well against zone coverage. Comes out of breaks quickly and gets good separation once he is in his routes."
    picDisplay.Print "Negatives: Inconsistency when it comes to catching the ball. Won't always make the tough catch in traffic. Smallish frame makes"; Tab(11); "him vulnerable to be knocked off routes by larger and more physical defensive backs. "

End Sub

