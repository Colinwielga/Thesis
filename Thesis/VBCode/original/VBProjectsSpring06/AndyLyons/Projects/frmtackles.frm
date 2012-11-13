VERSION 5.00
Begin VB.Form frmtackles 
   BackColor       =   &H00800000&
   Caption         =   "Tackles"
   ClientHeight    =   7515
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8340
   LinkTopic       =   "Form1"
   ScaleHeight     =   7515
   ScaleWidth      =   8340
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      Height          =   3135
      Left            =   480
      ScaleHeight     =   3075
      ScaleWidth      =   7155
      TabIndex        =   6
      Top             =   3960
      Width           =   7215
   End
   Begin VB.CommandButton cmdtackles 
      BackColor       =   &H80000010&
      Caption         =   "Return to Offensive Players"
      Height          =   615
      Left            =   3360
      TabIndex        =   0
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tackles"
      BeginProperty Font 
         Name            =   "Castellar"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   7
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Label lbltrueblood 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Jeremy Trueblood"
      Height          =   495
      Left            =   6360
      TabIndex        =   5
      Top             =   360
      Width           =   975
   End
   Begin VB.Image imgtrueblood 
      Height          =   1275
      Left            =   6360
      Picture         =   "frmtackles.frx":0000
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label lblwinston 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Eric Winston"
      Height          =   255
      Left            =   5040
      TabIndex        =   4
      Top             =   360
      Width           =   1215
   End
   Begin VB.Image imgwinston 
      Height          =   1695
      Left            =   5040
      Picture         =   "frmtackles.frx":494E
      Top             =   840
      Width           =   1065
   End
   Begin VB.Label lblmcneill 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Marcus McNeill"
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label lbljustice 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Winston Justice"
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
   Begin VB.Image imgmcneill 
      Height          =   1605
      Left            =   3600
      Picture         =   "frmtackles.frx":A8E8
      Top             =   840
      Width           =   1245
   End
   Begin VB.Image imgjustice 
      Height          =   1710
      Left            =   2160
      Picture         =   "frmtackles.frx":1127E
      Top             =   840
      Width           =   1140
   End
   Begin VB.Label lblbrick 
      BackColor       =   &H00FFFFFF&
      Caption         =   "D'Brickashaw Ferguson"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.Image imgbrickshaw 
      Height          =   1605
      Left            =   600
      Picture         =   "frmtackles.frx":17848
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "frmtackles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'2006 NFL Draft Simulator (Draft.vbp)
'frmtackles(frmtackles.frm)
'Andy Lyons
'March 24, 2006
'This form is used for searching the Offensive Tackles eligible for the 2006 NFL Draft. By clicking on each persons photo, it allows the user to view each players profile.
'returns user to main menu
Private Sub cmdtackles_Click()
    frmoffpositions.Show
    frmtackles.Hide
End Sub


Private Sub imgbrickshaw_Click()
    picResults.Cls
    picResults.Print "Height 6'6""", "Weight 312"
    picResults.Print "Positives: Universally regarded as one of the most dominant tackles to emerge from the college"; Tab(11); "ranks in recent years. Tremendous size, upper-body strength, arm length and punch."; Tab(11); "Amazingly quick out of his stance and in setting up for pass protection.  Shows"; Tab(11); "considerable athleticism for his size. Great hustle and balance."
    picResults.Print "Negatives: Lacks lower-body power and explosiveness, but should be able to improve that by"; Tab(11); " adding bulk (he played at 290 pounds last season) and being committed to an NFL"; Tab(11); " team's strength-and-conditioning program."

End Sub

Private Sub imgjustice_Click()
    picResults.Cls
    picResults.Print "Height 6'6""", "Weight 312"
    picResults.Print "Positives: Excellent size, upper-body strength, and arm length. Superior athleticism, especially"; Tab(11); "when it comes to adjusting to an inside rush. Has had advanced development"; Tab(11); "of pass-protection skills in a pro-style offense. Takes good blocking angles. Superb"; Tab(11); "footwork, especially when blocking for the run."
    picResults.Print "Negatives: Despite being athletically gifted, he occasionally has problems handling ultra-fast"; Tab(11); "outside rushers. Needs to work on developing greater explosion and taking better"; Tab(11); "advantage of his massive frame."

End Sub

Private Sub imgmcneill_Click()
    picResults.Cls
    picResults.Print "Height 6'7""", "Weight 336"
    picResults.Print "Positives: Amazing combination of abundant size and speed. Ability to consistently dominant"; Tab(11); "defenders. Does a good job of utilizing long arms and strong hands. Sets feet quickly"; Tab(11); "and has footwork to handle speed rushers. Bulk and strength allow him to take on"; Tab(11); "bull rushers with little problem."
    picResults.Print "Negatives: Must work on improving technical aspects of his game, especially maintaining good"; Tab(11); "leverage. Concerns over his ability to keep his weight under control."

End Sub

Private Sub imgtrueblood_Click()
    picResults.Cls
    picResults.Print "Height 6'8""", "Weight 316"
picResults.Print "Positives: Ultra-large frame, arm length (34½ inches), and hand size (10½ inches). Better athlete"; Tab(11); "than size might indicate. Good quickness and footwork. Takes good angles on blocks"; Tab(11); "and finishes well."
picResults.Print "Negatives: Plenty of work needed on techniques, especially when it comes to preventing defenders"; Tab(11); "from getting under his pads and maintaining balance. Needs to develop better lateral"; Tab(11); "movement and body control in the open field. Could add some lower-body strength."

End Sub

Private Sub imgwinston_Click()
    picResults.Cls
    picResults.Print "Height 6'6""", "Weight 310"
    picResults.Print "Positives: Former collegiate tight end brings considerable athleticism to the position."; Tab(11); "Impressive performances in Scouting Combine drills and at the Hurricanes' Pro Day on"; Tab(11); " May 4 figure to do wonders for his draft stock. Shows good football temperament and"; Tab(11); "physical style of play that matches the mentality of his opponents. Good size  and strength. Long arms and strong hands. Superb body control and footwork."
    picResults.Print "Negatives: Suffered a torn knee ligament that ended his season in 2004, although pre-draft"; Tab(11); "workouts have put to rest many of the lingering concerns. Pass-protection technique"; Tab(11); " needs work. Must improve lower-body strength."
 
End Sub
