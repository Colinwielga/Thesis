VERSION 5.00
Begin VB.Form frmsdts 
   BackColor       =   &H000080FF&
   Caption         =   "Defensive Tackles"
   ClientHeight    =   7320
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   ScaleHeight     =   7320
   ScaleWidth      =   8985
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picDisplay 
      Height          =   3255
      Left            =   360
      ScaleHeight     =   3195
      ScaleWidth      =   8115
      TabIndex        =   6
      Top             =   3840
      Width           =   8175
   End
   Begin VB.CommandButton cmdreturn 
      Caption         =   "Return to Defensive Positions"
      Height          =   855
      Left            =   3600
      TabIndex        =   5
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label lbldts 
      Caption         =   "Defensive Tackles"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   24
         Charset         =   1
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   7
      Top             =   2160
      Width           =   3735
   End
   Begin VB.Label lblwright 
      Caption         =   "Rodrique Wright"
      Height          =   255
      Left            =   6960
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.Image imgwright 
      Height          =   1665
      Left            =   6960
      Picture         =   "frmsdts.frx":0000
      Top             =   480
      Width           =   1245
   End
   Begin VB.Label lblwroten 
      Caption         =   "Claude Wroten"
      Height          =   255
      Left            =   5520
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.Image imgwroten 
      Height          =   1485
      Left            =   5520
      Picture         =   "frmsdts.frx":6D86
      Top             =   480
      Width           =   1290
   End
   Begin VB.Label lblharris 
      Caption         =   "Orien Harris"
      Height          =   255
      Left            =   4200
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.Image imgharris 
      Height          =   1590
      Left            =   4200
      Picture         =   "frmsdts.frx":D254
      Top             =   360
      Width           =   945
   End
   Begin VB.Label lblwatson 
      Caption         =   "Gabe Watson"
      Height          =   255
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.Image imgwatson 
      Height          =   1395
      Left            =   2400
      Picture         =   "frmsdts.frx":12216
      Top             =   360
      Width           =   1440
   End
   Begin VB.Label lblngata 
      Caption         =   "Haloti Ngata"
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Image imgngata 
      Height          =   1995
      Left            =   600
      Picture         =   "frmsdts.frx":18AF8
      Top             =   360
      Width           =   1440
   End
End
Attribute VB_Name = "frmsdts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'2006 NFL Draft Simulator (Draft.vbp)
'frmsdts(frmsdts.frm)
'Andy Lyons
'March 24, 2006
'This form allows the user to view Defensive Tackles eligible for the 2006 draft. By clicking on their picture, it uploads the profile.
'returns user to main menu
Private Sub cmdreturn_Click()
    frmdefpositions.Show
    frmsdts.Hide
End Sub

Private Sub imgharris_Click()
    picDisplay.Cls
    picDisplay.Print "Height 6'3""", "Weight 301"
    picDisplay.Print "Positives: Initial quickness. Upper-body strength. Plays with good leverage. Separates from blockers well."; Tab(11); "Strong pursuit and knack for forcing fumbles. Shows good polish and readiness for a promising"; Tab(11); "NFL career. "
    picDisplay.Print "Negatives: At 6-3 and 301 pounds, doesn't have ideal size for the position. Needs to work on overall"; Tab(11); "techniques. Knee and elbow injuries in 2003 raise some durability concerns. "

End Sub


Private Sub imgngata_Click()
    picDisplay.Cls
    picDisplay.Print "Height 6'4""", "Weight 338"
    picDisplay.Print "Positives: Great size. Amazing combination of phenomenal strength and great quickness off the snap. Extremely"; Tab(11); "hard to budge. Had consistent success despite facing almost constant double- and triple-team blocking."; Tab(11); "Besides tying up blockers, also has remarkable ability to shake free and make tackles. Nice agility and"; Tab(11); "athleticism that allow him to change direction well in open field and be effective when dropping into"; Tab(11); "coverage in a zone-blitz scheme. "
    picDisplay.Print "Negatives: Occasionally too dependant on brute strength rather than techniques, which he still needs to develop."; Tab(12); "Needs to work on ability to anticipate draws and screens. "

End Sub

Private Sub imgwatson_Click()
    picDisplay.Cls
    picDisplay.Print "Height 6'3""", "Weight 339"
    picDisplay.Print "Positives: Good size. Ultra-powerful force that routinely commands double teams. Quickness off the ball."; Tab(11); "Instincts for the ball. Shows good discipline by not over-pursuing against the run and by staying in"; Tab(11); "rushing lane against the pass. Made strong impression as a run-stuffer in the Senior Bowl. "
    picDisplay.Print "Negatives: Athleticism and speed. Reputation for occasionally taking plays off."
End Sub

Private Sub imgwright_Click()
    picDisplay.Cls
    picDisplay.Print "Height 6'5""", "Weight 300"
    picDisplay.Print "Positives: Superb combination of large frame, excellent mobility, and considerable power. Arguably the best"; Tab(11); "athlete among players at his position. Shows a wide variety of pass-rush moves and runs with greater"; Tab(11); "fluidity than one would expect from a defensive tackle. Enough versatility to be a good fit in an odd- or"; Tab(11); "even-man front. "
    picDisplay.Print "Negatives: Must learn to use hands and take better advantage of long arms to escape blockers. Reputation for"; Tab(12); "not performing to full potential. A 2004 ankle injury raises some concerns about durability. "

End Sub
    
Private Sub imgwroten_Click()
    picDisplay.Cls
    picDisplay.Print "Height 6'2""", "Weight 302"
    picDisplay.Print "Positives: Upper-body strength. Explosiveness off the ball. Body control and ability to make plays on the"; Tab(11); "move. Inside pass-rush skills. Good ball instincts. "
    picDisplay.Print "Negatives: At 6-2 and 302 pounds, doesn't have ideal size for the position. Needs to play with more"; Tab(11); "leverage. Must make better use of hands to separate from blockers against the run. "
End Sub
