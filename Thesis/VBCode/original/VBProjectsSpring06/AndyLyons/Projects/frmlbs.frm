VERSION 5.00
Begin VB.Form frmlbs 
   BackColor       =   &H00000000&
   Caption         =   "Linebackers"
   ClientHeight    =   8895
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9345
   LinkTopic       =   "Form1"
   ScaleHeight     =   8895
   ScaleWidth      =   9345
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picDisplay 
      Height          =   3735
      Left            =   480
      ScaleHeight     =   3675
      ScaleWidth      =   8595
      TabIndex        =   6
      Top             =   4800
      Width           =   8655
   End
   Begin VB.CommandButton cmdreturn 
      Caption         =   "Return to Defensive Positions"
      Height          =   975
      Left            =   3720
      TabIndex        =   5
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label lbllbs 
      Caption         =   "Linebackers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      TabIndex        =   7
      Top             =   2880
      Width           =   2655
   End
   Begin VB.Label lblhodge 
      Caption         =   "Abdul Hodge"
      Height          =   255
      Left            =   7440
      TabIndex        =   4
      Top             =   360
      Width           =   1215
   End
   Begin VB.Image imghodge 
      Height          =   1665
      Left            =   7440
      Picture         =   "frmlbs.frx":0000
      Top             =   720
      Width           =   1245
   End
   Begin VB.Image imgparham 
      Height          =   2130
      Left            =   5760
      Picture         =   "frmlbs.frx":6D86
      Top             =   600
      Width           =   1410
   End
   Begin VB.Label lblparham 
      Caption         =   "Kai Parham"
      Height          =   255
      Left            =   5880
      TabIndex        =   3
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label lbljackson 
      Caption         =   "D'Qwell Jackson"
      Height          =   255
      Left            =   3720
      TabIndex        =   2
      Top             =   360
      Width           =   1695
   End
   Begin VB.Image imgjackson 
      Height          =   1800
      Left            =   3720
      Picture         =   "frmlbs.frx":10B50
      Top             =   720
      Width           =   1680
   End
   Begin VB.Image imggreenway 
      Height          =   1605
      Left            =   2280
      Picture         =   "frmlbs.frx":1A912
      Top             =   720
      Width           =   1185
   End
   Begin VB.Label lblgreenway 
      Caption         =   "Chad Greenway"
      Height          =   255
      Left            =   2280
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label lblhawk 
      Caption         =   "A.J. Hawk"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
   Begin VB.Image imghawk 
      Height          =   1530
      Left            =   600
      Picture         =   "frmlbs.frx":20DA4
      Top             =   720
      Width           =   1380
   End
End
Attribute VB_Name = "frmlbs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'2006 NFL Draft Simulator (Draft.vbp)
'frmlbs(frmlbs.frm)
'Andy Lyons
'March 24, 2006
'This form allows the user to view Linebackers eligible for the 2006 draft. By clicking on their picture, it uploads the profile.

'returns user to main menu
Private Sub cmdreturn_Click()
    frmdefpositions.Show
    frmlbs.Hide
End Sub

Private Sub imggreenway_Click()
    picDisplay.Cls
    picDisplay.Print "Height 6'2""", "Weight 242"
    picDisplay.Print "Positives: Tough, aggressive, explosive player. Excellent speed and lateral pursuit. Great body control. Superb vision"; Tab(11); "and intelligence. Big hitter who delivers a blow with authority."
    picDisplay.Print "Negatives: Must add lower-body strength. Needs to work on man-to-man pass-coverage skills. "
End Sub

Private Sub imghawk_Click()
    picDisplay.Cls
    picDisplay.Print "Height 6'1""", "Weight 248"
    picDisplay.Print "Positives: Tremendous speed and athleticism. True throwback player who constantly hustles and is extremely aggressive"; Tab(11); "at the point of attack. Makes plays sideline-to-sideline. Great instincts. Outstanding explosiveness and closing "; Tab(11); "speed on blitzes. Strong pass-coverage skills. Uses hands well to separate from blockers. Impressive workout"; Tab(11); "at the Scouting Combine. "
    picDisplay.Print "Negatives: Could add some bulk to be better able to handle offensive linemen."
End Sub

Private Sub imghodge_Click()
    picDisplay.Cls
    picDisplay.Print "Height 6'2""", "Weight 235"
    picDisplay.Print "Positives: Highly productive college player. Hard hitter and exceptionally good tackler. Toughness, strength, and speed."; Tab(11); "Shows good technique and leverage. Intelligence and instincts. "
    picDisplay.Print "Negatives: Change-of-direction skills. Man-to-man and zone-coverage ability. "

End Sub

Private Sub imgjackson_Click()
    picDisplay.Cls
    picDisplay.Print "Height 6'0""", "Weight 230"
    picDisplay.Print "Positives: Physical at point of attack. Superb change-of-direction skills. Good instincts and reads keys well. Strong"; Tab(11); "tackler and has a knack for forcing fumbles. Shows natural playmaking ability in pass coverage. "
    picDisplay.Print "Negatives: At 6-foot and 230 pounds, he does not have great size for the position, nor is he particularly fast. Must"; Tab(11); "add lower-body strength. "

End Sub

Private Sub imgparham_Click()
    picDisplay.Cls
    picDisplay.Print "Height 6'3""", "Weight 256"
    picDisplay.Print "Positives:  Good size. Excellent speed, body control and lateral movement. Plays with considerable power. Big hitter who"; Tab(11); "can stuff the run. Shows good pursuit. Steady improvement in overall instincts, but especially in awareness"; Tab(11); "on blitzes."
    picDisplay.Print "Negatives: Must improve leverage. Change-of-direction skills."
End Sub
