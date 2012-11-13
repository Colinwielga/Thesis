VERSION 5.00
Begin VB.Form frmWR 
   Caption         =   "Form1"
   ClientHeight    =   3360
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   Picture         =   "frmWR.frx":0000
   ScaleHeight     =   3360
   ScaleWidth      =   5985
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdback 
      Caption         =   "Back"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton cmdJA 
      Caption         =   "Jason Avant"
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   1320
      Width           =   1935
   End
   Begin VB.CommandButton cmdSH 
      Caption         =   "Santonio Holmes"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton cmdSM 
      Caption         =   "Sinorice Moss"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton cmdDH 
      Caption         =   "Derek Hagan"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   1560
      Width           =   1935
   End
   Begin VB.CommandButton cmdCJ 
      Caption         =   "Chad Jackson"
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
End
Attribute VB_Name = "frmWR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdback_Click()
    frmWR.Hide
    frmAthletes.Show
End Sub

Private Sub cmdCJ_Click()
    MsgBox "Positives: Superb concentration on the ball. ... Makes difficult catches in traffic. ... Excels at running underneath routes, but has speed and burst to turn them into long gains. ... Has size (6-1 and 201 pounds) and strength to fight off jams at the line. ... Runs precise routes.                      Negatives: NFL teams don't know how effective he can be on deep routes because he didn't run many of them in college. ... Must improve blocking technique and develop a more aggressive attitude toward making contact.", , "Chad Jackson (Florida)"
End Sub

Private Sub cmdDH_Click()
    MsgBox "Positives: Intelligent route-runner with more savvy than one would expect from a young receiver when it comes to getting open against zone and man-to-man coverage. ... Superb footwork and change-of-direction ability. ... Excels at recognizing coverages. ... Strong leadership skills.                      Negatives: Tends to avoid going over the middle for tough catch. ... Doesn't have great speed, which limits ability to get separation. ... Needs to work on getting release from the line."
",,"Derek Hagan (Arizona State)"
End Sub

Private Sub cmdJA_Click()
    MsgBox "Positives: Highly productive college career that probably would have received greater attention had he not been overshadowed by Braylon Edwards, whom the Browns made the third overall pick of the 2005 draft. ... Outstanding hands. ... Able to make every catch an NFL receiver needs to make -- long, short, over the middle. Runs excellent routes and shows considerable burst in and out of cuts. ... Good size (6-0 and 209 yards) and strength. ... Superior athleticism ... Reads coverages well. ... Shows tremendous determination to get open. ... Aggressive blocker.                        Negatives: Doesn't have top-level speed. ... Durability issues.", , "Jason Avant (Michigan)"
End Sub

Private Sub cmdSH_Click()
    MsgBox "Positives: Dependable, sure-handed player who easily is the draft's best player at his position. ... Will make tough catches over the middle and can grab passes thrown over his head in full stride. ... Runs crisp, precise routes. ... Gets good separation once he is in his pattern. ... Outstanding footwork that allows him to make quick starts and stops. ... Has explosiveness to quickly get upfield after the catch and be a legitimate threat to go the distance whenever the ball is in his hands.                        Negatives: Must reduce number of passes that he catches with his body, rather than strictly with his hands. ... Smallish frame (5-foot-10 and 198 pounds) creates some difficulties when it comes to releasing from the line of scrimmage. ... Needs to learn how to make better use of hands to escape defenders trying to jam him.", , "Santonio Holmes (Ohio State)"
End Sub

Private Sub cmdSM_Click()
    MsgBox "Positives: Great speed and a threat to score on every catch. ... Makes good adjustments to poorly thrown balls. ... Works particularly well against zone coverage. ... Comes out of breaks quickly and gets good separation once he is in his routes.                       Negatives: Inconsistency when it comes to catching the ball. ... Won't always make the tough catch in traffic. ... Smallish frame (5-7 and 183 pounds) makes him vulnerable to be knocked off routes by larger and more physical defensive backs.", , "Sinorice Moss (Miami)"
End Sub
