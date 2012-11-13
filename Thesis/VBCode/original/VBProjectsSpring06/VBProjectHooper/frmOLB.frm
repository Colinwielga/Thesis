VERSION 5.00
Begin VB.Form frmOLB 
   Caption         =   "Form1"
   ClientHeight    =   5625
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   Picture         =   "frmOLB.frx":0000
   ScaleHeight     =   5625
   ScaleWidth      =   7440
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdback 
      Caption         =   "Back"
      Height          =   495
      Left            =   5760
      TabIndex        =   5
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton cmdAH 
      Caption         =   "A.J. Hawk"
      Height          =   495
      Left            =   2400
      TabIndex        =   4
      Top             =   240
      Width           =   2055
   End
   Begin VB.CommandButton cmdTH 
      Caption         =   "Thomas Howard"
      Height          =   495
      Left            =   4560
      TabIndex        =   3
      Top             =   960
      Width           =   2055
   End
   Begin VB.CommandButton cmdDR 
      Caption         =   "DeMeco Ryans"
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   2520
      Width           =   2055
   End
   Begin VB.CommandButton cmdCG 
      Caption         =   "Chad Greenway"
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   1200
      Width           =   2055
   End
   Begin VB.CommandButton cmdBC 
      Caption         =   "Bobby Carpenter"
      Height          =   495
      Left            =   5040
      TabIndex        =   0
      Top             =   2040
      Width           =   2055
   End
End
Attribute VB_Name = "frmOLB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'show profiles via message boxes
Option Explicit

Private Sub cmdAH_Click()
    MsgBox "Positives: Tremendous speed and athleticism. … True throwback player who constantly hustles and is extremely aggressive at the point of attack. … Makes plays sideline-to-sideline. … Great instincts. … Outstanding explosiveness and closing speed on blitzes. … Strong pass-coverage skills. … Uses hands well to separate from blockers. … Impressive workout at the Scouting Combine.                      Negatives: Could add some bulk to be better able to handle offensive linemen.", , "A.J. Hawk (Ohio State)"
End Sub

Private Sub cmdBack_Click()
    frmOLB.Hide
    frmAthletes.Show
End Sub

Private Sub cmdBC_Click()
    MsgBox "Positives: Good size (6-2 and 256 pounds) and speed. … Tough and instinctive. … Man-to-man pass-coverage skills. … Versatility to play any linebacker spot. … Strong work ethic.                        Negatives: Must add strength and learn how to play with greater leverage. … Suffered a broken ankle in mid-November.", "Bobby Carpenter (Ohio State)"
End Sub

Private Sub cmdCG_Click()
    MsgBox "Positives: Tough, aggressive, explosive player. … Excellent speed and lateral pursuit. … Great body control. … Superb vision and intelligence. … Big hitter who delivers a blow with authority.                     Negatives: Must add lower-body strength. … Needs to work on man-to-man pass-coverage skills.", , "Chad Greenway (Iowa)"
End Sub

Private Sub cmdDR_Click()
    MsgBox "Positives: Impressive display of techniques and instincts. … Shows good leverage and discipline in avoiding over-pursuit. … Quickness. … Use of hands to shed blockers.                     Negatives: At 6-1 and 236 pounds, doesn't have ideal size to hold his own at the point of attack. … Needs to add strength and bulk.", , "DeMeco Ryans (Alabama)"
End Sub

Private Sub cmdTH_Click()
    MsgBox "Positives: Good size (6-foot-3 and 239 pounds), speed, and athleticism. … Outside pass-rushing skills. … Sideline-to-sideline playmaking. … Change-of-direction ability. … Plenty of room for growth physically and in general knowledge and awareness of playing the position at a much higher level of football.                      Negatives: Must add upper-body strength and make better use of hands to separate from blockers. … Too much reliance on speed and athleticism at mid-major-level college competition.", , "Thomas Howard (UTEP)"
End Sub
