VERSION 5.00
Begin VB.Form frmcenters 
   BackColor       =   &H00000080&
   Caption         =   "Centers"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8355
   LinkTopic       =   "Form1"
   ScaleHeight     =   6210
   ScaleWidth      =   8355
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picDisplay 
      Height          =   2655
      Left            =   240
      ScaleHeight     =   2595
      ScaleWidth      =   7875
      TabIndex        =   1
      Top             =   3240
      Width           =   7935
   End
   Begin VB.CommandButton cmdreturn 
      Caption         =   "Return to Offensive Players"
      Height          =   855
      Left            =   3360
      TabIndex        =   0
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label lblmangold 
      Caption         =   "Nick Mangold"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lbldegory 
      Caption         =   "Mike Degory"
      Height          =   255
      Left            =   1680
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lbleslinger 
      Caption         =   "Greg Eslinger"
      Height          =   255
      Left            =   3240
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblspitz 
      Caption         =   "Jason Spitz"
      Height          =   255
      Left            =   4920
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblcook 
      Caption         =   "Ryan Cook"
      Height          =   255
      Left            =   6720
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.Image imgcook 
      Height          =   1590
      Left            =   6840
      Picture         =   "frmcenters.frx":0000
      Top             =   360
      Width           =   945
   End
   Begin VB.Image imgspitz 
      Height          =   1590
      Left            =   5040
      Picture         =   "frmcenters.frx":4FC2
      Top             =   360
      Width           =   945
   End
   Begin VB.Image imgeslinger 
      Height          =   1665
      Left            =   3240
      Picture         =   "frmcenters.frx":9F84
      Top             =   360
      Width           =   1215
   End
   Begin VB.Image imgdegory 
      Height          =   1530
      Left            =   1800
      Picture         =   "frmcenters.frx":10992
      Top             =   480
      Width           =   1110
   End
   Begin VB.Image imgmangold 
      Height          =   1590
      Left            =   480
      Picture         =   "frmcenters.frx":16314
      Top             =   480
      Width           =   1050
   End
End
Attribute VB_Name = "frmcenters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'2006 NFL Draft Simulator (Draft.vbp)
'frmcenters(frmcenters.frm)
'Andy Lyons
'March 24, 2006
'This form allows the user to look at the profiles of the center position available in the draft.
'Returns user to main menu
Private Sub cmdreturn_Click()
    frmoffpositions.Show
    frmcenters.Hide
End Sub

Private Sub imgcook_Click()
    picDisplay.Cls
    picDisplay.Print " Height 6'6""", "Weight 328"
    picDisplay.Print "Positives: Tremendous size. Good upper-body strength. Considerable arm length allows him to consistently lock"; Tab(11); "up pass rushers."
    picDisplay.Print "Negatives: Needs to add lower-body strength. Has work to do in gaining better stunt/blitz awareness. Must do a"; Tab(11); "better job of maintaining leverage and adjusting to make blocks in the open field."
End Sub

Private Sub imgdegory_Click()
    picDisplay.Cls
    picDisplay.Print "Height 6'5""", "Weight 305"
    picDisplay.Print "Positives: Excellent size and arm length. Superior technique in pass protection. Excellent overall strength."; Tab(11); "Great desire to improve his game, as reflected by steady development throughout his collegiate"; Tab(11); "career. Shows a great deal of intelligence and savvy in stunt and blitz recognition, and consistently"; Tab(11); "makes proper line calls. Willing and coachable enough to work at guard and center, which will"; Tab(11); "make him a more valuable backup early in his NFL career."
    picDisplay.Print "Negatives: Athleticism. Needs to improve body control and balance when making blocks in the open field. "

End Sub

Private Sub imgeslinger_Click()
    picDisplay.Cls
    picDisplay.Print "Height 6'3""", "Weight 292"
    picDisplay.Print "Positives: Good size and athleticism. Superior techniques. Outstanding footwork that allows him to get"; Tab(12); "consistently good leverage. Makes excellent use of hands to separate from defenders. Shows"; Tab(12); "superb recognition of stunts and blitzes and consistently makes proper line calls. Perfect fit for any"; Tab(12); "zone-blocking running scheme. "
    picDisplay.Print "Negatives: Needs to develop more lower-body strength. Although he is built just right for lateral movement"; Tab(12); "in a zone-blocking scheme, he could benefit from adding some bulk and strength, especially"; Tab(12); "when he has to take on a power-oriented inside rusher. "

End Sub

Private Sub imgmangold_Click()
    picDisplay.Cls
    picDisplay.Print "Height 6'3""", "Weight 300"
    picDisplay.Print "Positives: Good size and upper-body strength.Shows considerable toughness and will battle to the whistle."; Tab(11); "Possesses intelligence to recognize stunts and blitzes, and consistently makes proper line calls."; Tab(11); "Takes good angles on blocks. Does a nice job of pulling and trapping. Sets up quickly in pass"; Tab(11); "protection, which gives him an edge vs. speedy inside rushers."
    picDisplay.Print "Negatives: Must enhance lower-body strength to deal with massive, bull-rushing tackles. Needs to work on"; Tab(11); "change-of-direction skills to become better able to block defenders on the move."

End Sub


Private Sub imgspitz_Click()
    picDisplay.Cls
    picDisplay.Print "Height 6'3""", "Weight 313"
    picDisplay.Print "Positives: Good size, speed, and toughness. Makes good use of hands to gain separation. Solid run blocker."; Tab(11); "Good recognition of stunts and blitzes. Versatility, with more experience at guard than center."
    picDisplay.Print "Negatives: A converted guard who still is learning the center position. Must work on improving change-of-"; Tab(11); "direction ability and maintaining balance when blocking on the move. "

End Sub
