VERSION 5.00
Begin VB.Form frmguards 
   Caption         =   "Guards"
   ClientHeight    =   8610
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9405
   LinkTopic       =   "Form2"
   ScaleHeight     =   8610
   ScaleWidth      =   9405
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picDisplay 
      Height          =   3855
      Left            =   240
      ScaleHeight     =   3795
      ScaleWidth      =   8835
      TabIndex        =   1
      Top             =   3840
      Width           =   8895
   End
   Begin VB.CommandButton cmdreturn 
      Caption         =   "Return to Offensive Positions"
      Height          =   975
      Left            =   3720
      TabIndex        =   0
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label lblboothe 
      Caption         =   "Kevin Boothe"
      Height          =   255
      Left            =   7560
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.Image imgboothe 
      Height          =   1590
      Left            =   7680
      Picture         =   "frmguards.frx":0000
      Top             =   600
      Width           =   945
   End
   Begin VB.Label lblsims 
      Caption         =   "Rob Sims"
      Height          =   255
      Left            =   5640
      TabIndex        =   5
      Top             =   120
      Width           =   735
   End
   Begin VB.Image imgsims 
      Height          =   1200
      Left            =   5400
      Picture         =   "frmguards.frx":4FC2
      Top             =   600
      Width           =   1905
   End
   Begin VB.Label lblspencer 
      Caption         =   "Charles Spencer"
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.Image imgspencer 
      Height          =   1545
      Left            =   3840
      Picture         =   "frmguards.frx":C804
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label lbljoseph 
      Caption         =   "Davin Joseph"
      Height          =   255
      Left            =   2280
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.Image imgjoseph 
      Height          =   1665
      Left            =   2400
      Picture         =   "frmguards.frx":1341A
      Top             =   720
      Width           =   1245
   End
   Begin VB.Label lblgilles 
      Caption         =   "Max Jean-Gilles"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.Image imggilles 
      Height          =   2055
      Left            =   360
      Picture         =   "frmguards.frx":1A1A0
      Top             =   480
      Width           =   1740
   End
End
Attribute VB_Name = "frmguards"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'2006 NFL Draft Simulator (Draft.vbp)
'frmguards(frmguards.frm)
'Andy Lyons
'March 24, 2006
'This form allows the user to view Offensive Guards eligible for the 2006 draft. By clicking on their picture, it uploads the profile.

'Clicking this button brings the user back to the Offensive Player screen
'returns user to main menu
Private Sub cmdreturn_Click()
    frmoffpositions.Show
    frmguards.Hide
End Sub


Private Sub imgboothe_Click()
    picDisplay.Cls
    picDisplay.Print "Height 6'4""", "Weight 316"
    picDisplay.Print "Positives: Good size. Nice combination of bulk and athleticism. Long arms and powerful hands that he uses well to knock"; Tab(11); "back and lock onto defenders. Superior lateral movement. Intelligence and blitz/stunt awareness. "
    picDisplay.Print "Negatives: Durability after suffering ankle and hand injuries in high school and again in college. Lack of top-level speed"; Tab(11); "is a concern, especially in terms of getting to linebackers and safeties in the open field. "
End Sub

Private Sub imggilles_Click()
    picDisplay.Cls
    picDisplay.Print "Height 6'3""", "Weight 355"
    picDisplay.Print "Positives: Packs tremendous bulk in a 6-foot-3, 355-pound frame. Won't often be overpowered. Long arms and good use"; Tab(11); "of hands to lock onto defenders. Shows good body control and balance, especially in the open field. Impressive"; Tab(11); "footwork and quickness off the snap as well as laterally. Establishes leverage quickly. "
    picDisplay.Print "Negatives: Athleticism leaves something to be desired and stamina and conditioning need work. Needs to improve speed"; Tab(11); "on pulls and traps. "
End Sub

Private Sub imgjoseph_Click()
    picDisplay.Cls
    picDisplay.Print "Height 6'2""", "Weight 311"
    picDisplay.Print "Positives: Plays with considerable power that allows him to consistently rock defenders with hand punches. Does a good"; Tab(11); "job of maintaining leverage. Shows an aggressive mentality, especially in pass protection. Superior mobility"; Tab(11); "and body control allow him to pull and trap well, and make blocks in the open field."
    picDisplay.Print "Negatives: At 6-2, lacks ideal height, although 311-pound frame is plenty wide for an NFL guard. Blitz and stunt recognition."
End Sub

Private Sub imgsims_Click()
    picDisplay.Cls
    picDisplay.Print "Height 6'2""", "Weight 307"
    picDisplay.Print "Positives: Toughness and strength, especially when it comes to locking onto defenders. Plenty of width in"; Tab(11); "307-pound frame. Quick feet. Takes good angles on blocks. Recognition of blitz and stunts. Versatility, with"; Tab(11); " collegiate experience at tackle. "
    picDisplay.Print "Negatives: At 6-2, lacks ideal height. Agility. Must improve physical conditioning. "
End Sub

Private Sub imgspencer_Click()
    picDisplay.Cls
    picDisplay.Print "Height 6'4""", "Weight 352"
    picDisplay.Print "Positives: Good size. Highly athletic, which allows him to be effective in small areas as well as in the open field."; Tab(11); "Explodes out of stance and gets to linebackers and safeties quickly. Good arm extension and powerful enough to"; Tab(11); "rock defenders with hand punches. Remarkable progress since switching from defensive tackle as a junior, and"; Tab(11); "impressed many NFL scouts during Senior Bowl workouts. "
    picDisplay.Print "Negatives: Lack of offensive line experience. Must work at taking better advantage of size and strength when"; Tab(12); "run-blocking at point of attack. "

End Sub
