VERSION 5.00
Begin VB.Form frmQB 
   Caption         =   "Form1"
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5235
   LinkTopic       =   "Form1"
   Picture         =   "frmQB.frx":0000
   ScaleHeight     =   6780
   ScaleWidth      =   5235
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdback 
      Caption         =   "Back"
      Height          =   615
      Left            =   3360
      TabIndex        =   5
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CommandButton cmdOJ 
      Caption         =   "Omar Jacobs"
      Height          =   495
      Left            =   3600
      TabIndex        =   4
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton cmdJC 
      Caption         =   "Jay Cutler"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton cmdBC 
      Caption         =   "Brodie Croyle"
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CommandButton cmdVY 
      Caption         =   "Vince Young"
      Height          =   495
      Left            =   3360
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton cmdML 
      Caption         =   "Matt Leinart"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "frmQB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'show profiles via message boxes
Option Explicit

Private Sub cmdBack_Click()
    frmQB.Hide
    frmAthletes.Show
End Sub

Private Sub cmdBC_Click()
    MsgBox "Positives: Strong arm, nice touch. ... Quick release that is enhanced by exceptionally fast throwing motion. ... Reads defenses well. ... Makes good decisions throwing from the pocket. ... Leadership.                        Negatives: At 6-2, faces a challenge seeing over linemen, and 205-pound frame could make him susceptible to injuries. NFL stock already has been hurt by injury-filled collegiate career. ... Accuracy and touch on shorter routes. ... Pocket awareness and ability to avoid pressure.", , "Brodie Croyle (Alabama)"
End Sub

Private Sub cmdJC_Click()
    MsgBox "Positives: Strong arm and overall strength emanating from solid, 6-2, 223-pound frame. ... Makes every kind of throw, but is especially impressive on deep outs and squeezing the ball through small openings. ... Intelligence, patience in the pocket, and progression reads. ... Took advantage of having stage all to himself during Scouting Combine workouts, in which Leinart and Young did not participate, with mostly impressive performance.                     Negatives: Occasional overconfidence with arm strength can lead to gambles that backfire. ... Has some work to do on overall mechanics and getting rid of the ball quicker. ... Wasn't tremendously accurate during Combine throwing drills.", "Jay Cutler (Vanderbilt)"
End Sub

Private Sub cmdML_Click()
    MsgBox "Positives: Perfect size (6-foot-4 and 224 pounds) for an NFL quarterbackbecause he is tall enough to easily see over linemen and has a chance to hold up to physical punishment resulting from poor protection and rookie mistakes in picking up the blitz, etc. ... Large hands and long arms are also plusses. ... Very good accuracy, especially on touch passes. ... Excels at reading coverages and quickly finding open receiver. ... Better footwork than one might expect for his large frame, and is quick to set up for throws. ... Strong leadership; doesn't rattle easily vs. pressure or when faced with adversity.                       Negatives: Arm strength unspectacular, but good enough that it shouldn't prove a major problem ... as a lefty he will force the team that drafts him to put the best pass blocking tackle on the right, which is uncommon ... long throwing motion could give db's the edge on where he wants to go ... poor mobility", , "Matt Leinart (USC)"
 End Sub

Private Sub cmdOJ_Click()
    MsgBox "Positives: Size (6-4 and 224 pounds), strength, and athleticism. ... Strong arm, large hands. ... Puts plenty of zip on deep outs and is able to get the ball into tight spaces ... Wonderful accuracy and touch on short and intermediate throws.                      Negatives: Tends to pull down the ball and run too soon; needs to become more comfortable working from the pocket. ... Throwing mechanics require work. ... Field vision, ability to read defenses.", , "Omar Jacobs (Bowling Green)"
End Sub

Private Sub cmdVY_Click()
    MsgBox "Positives: Exceptional combination of size (6-5 and 230 pounds) and off-the-charts athleticism make him an extremely rare talent. ... Outstanding speed and mobility make him a constant threat to run, and will keep defenses occupied with mainly trying to minimize the damage he can do with his feet. ... Quick enough to avoid pressure in the pocket and keep a play alive, but strong enough to break tackles when he does run. ... Showed tremendous poise in leading the Longhorns and scoring his winning touchdown run for the national championship in the Rose Bowl, the biggest football stage this side of the Super Bowl.                      Negatives: Stories of poor results on intelligence test at the Scouting Combine were overblown, but could prove damaging to what once seemed like a surefire top-three (or top) pick. ... Sidearm delivery makes throws vulnerable. ... Overall passing mechanics are raw and need work. ... Patience and decision-making in the pocket.", , "Vince Young (Texas)"
End Sub
