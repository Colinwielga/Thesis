VERSION 5.00
Begin VB.Form frmTE 
   Caption         =   "Form1"
   ClientHeight    =   5610
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7470
   LinkTopic       =   "Form1"
   Picture         =   "frmTE.frx":0000
   ScaleHeight     =   5610
   ScaleWidth      =   7470
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdback 
      Caption         =   "Back"
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdLP 
      Caption         =   "Leonard Pope"
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   480
      Width           =   1935
   End
   Begin VB.CommandButton cmdVD 
      Caption         =   "Vernon Davis"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton cmdML 
      Caption         =   "Mercedes Lewis"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   1815
   End
   Begin VB.CommandButton cmdAF 
      Caption         =   "Anthony Fasano"
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton cmdJK 
      Caption         =   "Joe Klopfenstein"
      Height          =   375
      Left            =   5280
      TabIndex        =   0
      Top             =   1320
      Width           =   1815
   End
End
Attribute VB_Name = "frmTE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'show profiles via message boxes
Option Explicit

Private Sub cmdAF_Click()
    MsgBox "Positives: Catches the ball well in traffic and on the run. … Tremendous intelligence allows him to consistently find openings against zone and man-to-man coverage. … Worked well in a pro-style offense that should enhance his readiness for the NFL. … Good initial burst helps compensate for general lack of speed. … Displays solid blocking technique. … Superior hustle and work ethic.                        Negatives: Lack of speed and athleticism. … Doesn't make many would-be tacklers miss and is not particularly effective on seam routes, where tight ends usually make their greatest impact.", , "Anthony Fasano (Notre Dame)"
End Sub

Private Sub cmdBack_Click()
    frmTE.Hide
    frmAthletes.Show
End Sub

Private Sub cmdJK_Click()
    MsgBox "Positives: Nice combination of height (6-5) and speed. … Runs good, crisp routes. … Will battle for the ball. … Intelligence. … Durability. … Strong work ethic.                        Negatives: Doesn 't consistently get good release from the line. … Must add bulk and strength to 250-pound frame to become a more effective blocker.", , "Joe Klopfenstein (Colorado)"
End Sub

Private Sub cmdLP_Click()
    MsgBox "Positives: Tall target (6-7 and 250 pounds) with large hands and long arms that he uses well to keep defenders at bay. … Impressive athletic display at Scouting Combine. … Superior speed that is capable of stretching coverage. … Shows strong explosion off the line of scrimmage. … Excels at finding soft spots in zone coverage.                     Negatives: Needs to add some bulk, which figures to be fairly easy given his frame. … Has problems sometimes with defenders who are able to get their hands on him near the line.", , "Leonard Pope (Georgia)"
End Sub

Private Sub cmdML_Click()
    MsgBox "Positives: Concentration when catching the ball. … Makes good use of 6-6 frame, particularly on jump balls. … Good route-runner. … Solid blocker.                       Negatives: Doesn 't get much explosion off the line. … Has problems releasing on the snap. … Lacks a natural running stride.", , "Mercedes Lewis (UCLA)"
End Sub

Private Sub cmdVD_Click()
    MsgBox "Positives: Off-the-charts performance in Scouting Combine drills only enhanced the status of a player already widely regarded as the draft's best at his position. … Has a staggering combination of outstanding size (6-foot-3 and 254 pounds), speed (ran 40-yard dash at Combine in 4.38 seconds), strength (best bench press showing of any tight end at the Combine), and athleticism (impressive vertical leap of 40 inches). … Dependable hands. … Excellent route-runner with knowledge to find openings against zone- and man-to-man coverage, and necessary burst to create separation in his patterns. … Capable of consistently turning short catches into long gains.                      Negatives: Needs to work on his blocking skills. … Despite strength, doesn't seem to know how to use hands properly when engaging with linebackers.", , "Vernon Davis (Maryland)"
End Sub

Private Sub Form_Load()
    frmTE.Hide
    frmAthletes.Show
End Sub
