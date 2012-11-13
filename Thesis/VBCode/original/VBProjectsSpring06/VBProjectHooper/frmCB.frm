VERSION 5.00
Begin VB.Form frmCB 
   Caption         =   "Form1"
   ClientHeight    =   7515
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5130
   LinkTopic       =   "Form1"
   Picture         =   "frmCB.frx":0000
   ScaleHeight     =   7515
   ScaleWidth      =   5130
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAP 
      Caption         =   "Anwar Phillips"
      Height          =   495
      Left            =   1440
      TabIndex        =   5
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton cmdAY 
      Caption         =   "Ashton Youboty"
      Height          =   495
      Left            =   960
      TabIndex        =   4
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CommandButton cmdJW 
      Caption         =   "Jimmy Williams"
      Height          =   495
      Left            =   3000
      TabIndex        =   3
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton cmdAC 
      Caption         =   "Antonio Cromartie"
      Height          =   495
      Left            =   840
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdTH 
      Caption         =   "Tye Hill"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   6840
      Width           =   2055
   End
End
Attribute VB_Name = "frmCB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'message boxes containing profiles
Option Explicit

Private Sub cmdAC_Click()
    MsgBox "Positives: Great size (6-2 and 203 pounds). … Outstanding speed, athleticism, and leaping ability. … Superb instincts. … Long arms and strong hands. … Provides strong run support. … Versatility with background as a receiver and kick-return skills.                     Negatives: Missed the 2005 season after undergoing major knee surgery, although he ran well in his individual workout for scouts. … Needs to work on anticipation of pass routes.", , "Antonio Cromartie (Florida State)"
End Sub

Private Sub cmdAP_Click()
    MsgBox "Positives: Good size (5-11 and 190 pounds) and long arms. … Excellent leaping ability and ball skills. … Strength and aggressiveness allow him to hold up well against larger receivers. … Solid in run support                     Negatives: Must work on play-recognition skills and get out of the habit of focusing too much on the quarterback. … Needs to develop better tackling skills in the open field.", , "Anwar Phillips (Penn State)"
End Sub

Private Sub cmdAY_Click()
    MsgBox "Positives: Good size (5-11 and 189 pounds). … Natural playmaker. … Excellent speed and athleticism. … Good intelligence and instincts. … Change-of-direction ability. … Solid in run support                        Negatives: Lacks bulk and strength. … Needs to develop better play-recognition skills, which should happen as he gains experience and coaching.", , "Ashton Youboty (Ohio State)"
End Sub

Private Sub cmdBack_Click()
    frmCB.Hide
    frmAthletes.Show
End Sub

Private Sub cmdJW_Click()
    MsgBox "Positives: Excellent size (6-2 and 216 pounds). … Outstanding speed and change-of-direction skills. … Physical player with considerable strength. … Good hands and a threat to go the distance with every interception. … Could excel at safety, as well as at cornerback.                      Negatives: Poor individual workout for scouts leaves lingering negative impression given that he opted not to work out at the Scouting Combine. … Must become more focused on receiver he covers rather than devoting too much attention to the quarterback.", , "Jimmy Williams (Virginia Tech)"
End Sub

Private Sub cmdTH_Click()
    MsgBox "Positives: Tremendous speed. … Good instincts. … Excels in bump-and-run coverage. … Handles himself well in man-to-man and zone coverage. … Tough, physical player who is aggressive against the run. … Outstanding leaping ability.                        Negatives: At 5-foot-9 and 185 pounds, he lacks ideal size and doesn't have much bulk. … Must work on improving his ability to read the eyes of the quarterback and overall technique.", , "Tye Hill (Clemson)"
End Sub
