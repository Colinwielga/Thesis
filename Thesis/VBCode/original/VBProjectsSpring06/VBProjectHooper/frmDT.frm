VERSION 5.00
Begin VB.Form frmDT 
   Caption         =   "Form1"
   ClientHeight    =   7515
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   Picture         =   "frmDT.frx":0000
   ScaleHeight     =   7515
   ScaleWidth      =   6030
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdback 
      Caption         =   "Back"
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   6960
      Width           =   1935
   End
   Begin VB.CommandButton cmdCW 
      Caption         =   "Claude Wroten"
      Height          =   735
      Left            =   3120
      TabIndex        =   4
      Top             =   4680
      Width           =   2295
   End
   Begin VB.CommandButton cmdOH 
      Caption         =   "Orien Harris"
      Height          =   735
      Left            =   480
      TabIndex        =   3
      Top             =   4680
      Width           =   2295
   End
   Begin VB.CommandButton cmdHN 
      Caption         =   "Haloti Ngata"
      Height          =   735
      Left            =   1800
      TabIndex        =   2
      Top             =   2640
      Width           =   2295
   End
   Begin VB.CommandButton cmdGW 
      Caption         =   "Gabe Watson"
      Height          =   735
      Left            =   480
      TabIndex        =   1
      Top             =   3600
      Width           =   2295
   End
   Begin VB.CommandButton cmdRW 
      Caption         =   "Rodrique Wright"
      Height          =   735
      Left            =   3120
      TabIndex        =   0
      Top             =   3600
      Width           =   2295
   End
End
Attribute VB_Name = "frmDT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'show profiles via message boxes
Option Explicit

Private Sub cmdBack_Click()
    frmDT.Hide
    frmAthletes.Show
End Sub

Private Sub cmdCW_Click()
    MsgBox "Positives: Upper-body strength. ... Explosiveness off the ball. ... Body control and ability to make plays on the move. ... Inside pass-rush skills. ... Good ball instincts.                       Negatives: At 6-2 and 302 pounds, doesn't have ideal size for the position. ... Needs to play with more leverage. ... Must make better use of hands to separate from blockers against the run.", , "Claude Wroten (LSU)"
End Sub

Private Sub cmdGW_Click()
    MsgBox "Positives: Good size (6-3 and 339 pounds). ... Ultra-powerful force that routinely commands double teams. ... Quickness off the ball. ... Instincts for the ball. ... Shows good discipline by not over-pursuing against the run and by staying in rushing lane against the pass. ... Made strong impression as a run-stuffer in the Senior Bowl.                       Negatives: Athleticism and speed. ... Reputation for occasionally taking plays off.", , "Gabe Watson (Michigan)"
End Sub

Private Sub cmdHN_Click()
    MsgBox "Positives: Great size (6-foot-4 and 338 pounds). ... Amazing combination of phenomenal strength and great quickness off the snap. ... Extremely hard to budge. ... Had consistent success despite facing almost constant double- and triple-team blocking. ... Besides tying up blockers, also has remarkable ability to shake free and make tackles. ... Nice agility and athleticism that allow him to change direction well in open field and be effective when dropping into coverage in a zone-blitz scheme.                       Negatives: Occasionally too dependant on brute strength rather than techniques, which he still needs to develop. ... Needs to work on ability to anticipate draws and screens.", , "Haloti Ngata (Oregon)"
End Sub

Private Sub cmdOH_Click()
    MsgBox "Positives: Initial quickness. ... Upper-body strength. ... Plays with good leverage. ... Separates from blockers well. ... Strong pursuit and knack for forcing fumbles. ... Shows good polish and readiness for a promising NFL career.                        Negatives: At 6-3 and 301 pounds, doesn't have ideal size for the position. ... Needs to work on overall techniques. ... Knee and elbow injuries in 2003 raise some durability concerns.", , "Orien Harris (Miami)"
End Sub

Private Sub cmdRW_Click()
    MsgBox "Positives: Superb combination of large frame (6-5 and 300 pounds), excellent mobility, and considerable power. ... Arguably the best athlete among players at his position. ... Shows a wide variety of pass-rush moves and runs with greater fluidity than one would expect from a defensive tackle. ... Enough versatility to be a good fit in an odd- or even-man front.                     Negatives: Must learn to use hands and take better advantage of long arms to escape blockers. ... Reputation for not performing to full potential. ... A 2004 ankle injury raises some concerns about durability.", , "Rodrique Wright (Texas)"
End Sub
