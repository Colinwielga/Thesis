VERSION 5.00
Begin VB.Form frmILB 
   Caption         =   "Form1"
   ClientHeight    =   5625
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8235
   LinkTopic       =   "Form1"
   Picture         =   "frmILB.frx":0000
   ScaleHeight     =   5625
   ScaleWidth      =   8235
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdback 
      Caption         =   "Back"
      Height          =   615
      Left            =   3240
      TabIndex        =   5
      Top             =   4800
      Width           =   1575
   End
   Begin VB.CommandButton cmdAH 
      Caption         =   "Abdul Hodge"
      Height          =   495
      Left            =   2880
      TabIndex        =   4
      Top             =   600
      Width           =   2055
   End
   Begin VB.CommandButton cmdDR 
      Caption         =   "Dale Robinson"
      Height          =   495
      Left            =   5640
      TabIndex        =   3
      Top             =   960
      Width           =   2055
   End
   Begin VB.CommandButton cmdFR 
      Caption         =   "Freddie Roach"
      Height          =   495
      Left            =   5640
      TabIndex        =   2
      Top             =   240
      Width           =   2055
   End
   Begin VB.CommandButton cmdKP 
      Caption         =   "Kai Parham"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   2055
   End
   Begin VB.CommandButton cmdDJ 
      Caption         =   "D'Qwell Jackson"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "frmILB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'show profiles via message boxes
Option Explicit

Private Sub cmdAH_Click()
    MsgBox "Positives: Highly productive college player. ... Hard hitter and exceptionally good tackler. ... Toughness, strength, and speed. ... Shows good technique and leverage. ... Intelligence and instincts.                     Negatives: Change-of-direction skills. ... Man-to-man and zone-coverage ability.", , "Abdul Hodge (Iowa)"
End Sub

Private Sub cmdBack_Click()
    frmILB.Hide
    frmAthletes.Show
End Sub

Private Sub cmdDJ_Click()
    MsgBox "Positives: Physical at point of attack. ... Superb change-of-direction skills. ... Good instincts and reads keys well. ... Strong tackler and has a knack for forcing fumbles. ... Shows natural playmaking ability in pass coverage.                       Negatives: At 6-foot and 230 pounds, he does not have great size for the position, nor is he particularly fast. ... Must add lower-body strength.", , "D'Qwell Jackson (Maryland)"
End Sub

Private Sub cmdDR_Click()
    MsgBox "Positives: Big playmaker who always hustles to the whistle. ... Hard hitter. ... Lower-body strength. ... Intelligence and instincts. ... Good recognition of offensive keys and quick ability to process what he sees. ... Toughness.                      Negatives: Doesn't have great size (6-0 and 231 pounds) or speed. ... Struggles to shed blocks.", , "Dale Robinson (Arizona State)"
End Sub

Private Sub cmdFR_Click()
    MsgBox "Positives: Strength and ability to shed blockers. ... Excellent toughness. ... Hard hitter. ... Outstanding instincts and good technique. ... Good footwork. ... Versatile enough to play inside or outside.                        Negatives: Doesn't have great speed or athleticism. ... Change-of-direction skills.", , "Freddie Roach (Alabama)"
End Sub

Private Sub cmdKP_Click()
    MsgBox "Positives: Good size (6-3 and 256 pounds). ... Excellent speed, body control and lateral movement. ... Plays with considerable power. ... Big hitter who can stuff the run. ... Shows good pursuit. ... Steady improvement in overall instincts, but especially in awareness on blitzes.                        Negatives: Must improve leverage. ... Change-of-direction skills.", , "Kai Parham (Virginia)"
End Sub
