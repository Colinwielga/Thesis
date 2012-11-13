VERSION 5.00
Begin VB.Form frmDE 
   BackColor       =   &H0000FFFF&
   Caption         =   "Form1"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   Picture         =   "frmDE.frx":0000
   ScaleHeight     =   6495
   ScaleWidth      =   7455
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdback 
      Caption         =   "Back"
      Height          =   375
      Left            =   6240
      TabIndex        =   5
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton cmdDT 
      Caption         =   "Darryl Tapp"
      Height          =   855
      Left            =   4080
      TabIndex        =   4
      Top             =   2640
      Width           =   2535
   End
   Begin VB.CommandButton cmdML 
      Caption         =   "Manny Lawson"
      Height          =   855
      Left            =   4080
      TabIndex        =   3
      Top             =   5160
      Width           =   2535
   End
   Begin VB.CommandButton cmdTH 
      Caption         =   "Tamba Hali"
      Height          =   855
      Left            =   2160
      TabIndex        =   2
      Top             =   3840
      Width           =   2535
   End
   Begin VB.CommandButton cmdMK 
      Caption         =   "Mathias Kiwanuka"
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   5280
      Width           =   2535
   End
   Begin VB.CommandButton cmdMW 
      Caption         =   "Mario Williams"
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   2520
      Width           =   2535
   End
End
Attribute VB_Name = "frmDE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'messageboxes containing profiles
Option Explicit

Private Sub cmdBack_Click()
    frmDE.Hide
    frmAthletes.Show
End Sub

Private Sub cmdDT_Click()
    MsgBox "Positives: Outstanding strength that helps him to consistently get penetration. ... Plays with good leverage. ... Physical player who shows great intensity. ... Superb instincts help him to often make plays in the backfield.                        Negatives: At 6-1 and 252 pounds, doesn't have ideal size for the position. ... Limited athleticism and speed.", , "Darryl Tapp (Virginia Tech)"
End Sub

Private Sub cmdMK_Click()
    MsgBox "Positives: Good size (6-5 and 266 pounds) and speed. ... A former basketball player, he has plenty of athleticism. ... Nice variety of moves. ... Constant hustle allows him to often make plays in pursuit. ... Shows excellent instincts that he enhances with considerable study that allows him to make accurate pre-snap reads of offensive linemen. ... Solid containment on outside runs.                        Negatives: Tendency to play too tall and allow blockers to get under him. ... Needs to work on utilizing more of his quickness to get a better jump off the ball.", , "Mathias Kiwanuka (Boston College)"
End Sub

Private Sub cmdML_Click()
    MsgBox "Positives: Excellent athleticism and speed. ... Consistently zips around offensive tackles and has outstanding ability to close on the ball. ... Impressive strength. ... Great instincts. ... Works at building a repertoire of pass-rush moves.                       Negatives: Smallish (241 pounds) frame, although has the skills that could allow him to be an outside linebacker in 3-4 alignment. ... Lower-body strength. ... Must improve his leverage.", , "Manny Lawson (NC State)"
End Sub

Private Sub cmdMW_Click()
    MsgBox "Positives: Great size (6-foot-7 and 295 pounds). ... Superior speed and athleticism. ... Excellent anticipation of the snap and quickness off the ball allow him to win most battles on his first step. ... Outstanding closing speed. ... Excellent footwork and change-of-direction skills. ... Superb ball instincts. ... Shows equal dominance rushing the passer or stopping the run. ... Big hitter.                      Negatives: Despite tremendously long arms, needs to work at getting better separation from blockers. ... Must improve technique rather than mostly relying on running around blockers.", , "Mario Williams (NC State)"
End Sub

Private Sub cmdTH_Click()
    MsgBox "Positives: Explosiveness off the snap. ... Battles non-stop to the ball. ... Good pass-rushing moves. ... Upper-body strength, which allows him to escape blocks and hold his own at the point of attack. ... Shows great ball instincts. ... Impressive performance in the Senior Bowl.                        Negatives: At 6-3, he does not have ideal height for the position. ... Must develop better use of his hands to get separation from blockers.", , "Tamba Hali (Penn State)"
End Sub
