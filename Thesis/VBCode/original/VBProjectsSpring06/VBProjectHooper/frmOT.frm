VERSION 5.00
Begin VB.Form frmOT 
   Caption         =   "Form1"
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8520
   LinkTopic       =   "Form1"
   Picture         =   "frmOT.frx":0000
   ScaleHeight     =   8520
   ScaleWidth      =   8520
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdback 
      Caption         =   "Back"
      Height          =   615
      Left            =   480
      TabIndex        =   5
      Top             =   7440
      Width           =   1695
   End
   Begin VB.CommandButton cmdDF 
      Caption         =   "D'Brickashaw Ferguson"
      Height          =   735
      Left            =   4920
      TabIndex        =   4
      Top             =   360
      Width           =   2535
   End
   Begin VB.CommandButton cmdJT 
      Caption         =   "Jeremy Trueblood"
      Height          =   735
      Left            =   3240
      TabIndex        =   3
      Top             =   4200
      Width           =   2535
   End
   Begin VB.CommandButton cmdWJ 
      Caption         =   "Winston Justice"
      Height          =   735
      Left            =   2040
      TabIndex        =   2
      Top             =   600
      Width           =   2535
   End
   Begin VB.CommandButton cmdEW 
      Caption         =   "Eric Winston"
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   3360
      Width           =   2535
   End
   Begin VB.CommandButton cmdMM 
      Caption         =   "Marcus McNeil"
      Height          =   735
      Left            =   720
      TabIndex        =   0
      Top             =   1800
      Width           =   2535
   End
End
Attribute VB_Name = "frmOT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'show profiles via message boxes
Option Explicit
Private Sub cmdBack_Click()
    frmOT.Hide
    frmAthletes.Show
End Sub

Private Sub cmdDF_Click()
    MsgBox "Positives: Universally regarded as one of the most dominant tackles to emerge from the college ranks in recent years. ... Tremendous size (6-foot-6 and 312 pounds), upper-body strength, arm length and punch. ... Amazingly quick out of his stance and in setting up for pass protection. ... Shows considerable athleticism for his size. ... Great hustle and balance.                     Negatives: Lacks lower-body power and explosiveness, but should be able to improve that by adding bulk (he played at 290 pounds last season) and being committed to an NFL team's strength-and-conditioning program.", , "D'Brickshaw Ferguson (Virginia)"
End Sub

Private Sub cmdEW_Click()
    MsgBox "Positives: Former collegiate tight end brings considerable athleticism to the position. ... Impressive performances in Scouting Combine drills and at the Hurricanes' Pro Day on May 4 figure to do wonders for his draft stock. ... Shows good football temperament and physical style of play that matches the mentality of his opponents. ... Good size (6-6 and 310 pounds) and strength. ... Long arms and strong hands. ... Superb body control and footwork.                     Negatives: Suffered a torn knee ligament that ended his season in 2004, although pre-draft workouts have put to rest many of the lingering concerns. ... Pass-protection technique needs work. ... Must improve lower-body strength.", "Eric Winston (Miami)"
End Sub

Private Sub cmdJT_Click()
    MsgBox "Positives: Ultra-large frame (6-8 and 316 pounds), arm length (34½ inches), and hand size (10½ inches). ... Better athlete than size might indicate. ... Good quickness and footwork. ... Takes good angles on blocks and finishes well.                     Negatives: Plenty of work needed on techniques, especially when it comes to preventing defenders from getting under his pads and maintaining balance. ... Needs to develop better lateral movement and body control in the open field. ... Could add some lower-body strength.", , "Jeremy Trueblood (Boston College)"
End Sub

Private Sub cmdMM_Click()
    MsgBox "Positives: Amazing combination of abundant size (6-7 and 336 pounds) and speed. ... Ability to consistently dominant defenders. ... Does a good job of utilizing long arms and strong hands. ... Sets feet quickly and has footwork to handle speed rushers. ... Bulk and strength allow him to take on bull rushers with little problem.                       Negatives: Must work on improving technical aspects of his game, especially maintaining good leverage. ... Concerns over his ability to keep his weight under control.", , "Marcus McNeil (Auburn)"
End Sub

Private Sub cmdWJ_Click()
    MsgBox "Positives: Excellent size (6-6 and 319 pounds), upper-body strength, and arm length. ... Superior athleticism, especially when it comes to adjusting to an inside rush. ... Has had advanced development of pass-protection skills in a pro-style offense. ... Takes good blocking angles. ... Superb footwork, especially when blocking for the run.                       Negatives: Despite being athletically gifted, he occasionally has problems handling ultra-fast outside rushers. ... Needs to work on developing greater explosion and taking better advantage of his massive frame.", , "Winston Justice (USC)"
End Sub
