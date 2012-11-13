VERSION 5.00
Begin VB.Form frmFB 
   Caption         =   "Form1"
   ClientHeight    =   5595
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   Picture         =   "frmFB.frx":0000
   ScaleHeight     =   5595
   ScaleWidth      =   4500
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdback 
      Caption         =   "Back"
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton cmdLV 
      Caption         =   "Lawrence Vickers"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   3120
      Width           =   1935
   End
   Begin VB.CommandButton cmdNT 
      Caption         =   "Naufahu Tahi"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   1935
   End
   Begin VB.CommandButton cmdGH 
      Caption         =   "Gilbert Harris"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   1935
   End
   Begin VB.CommandButton cmdGM 
      Caption         =   "Garrett Mills"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton cmdMB 
      Caption         =   "Matt Bernstein"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   1935
   End
End
Attribute VB_Name = "frmFB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'show profiles via message boxes
Option Explicit
Private Sub cmdBack_Click()
    frmFB.Hide
    frmAthletes.Show
End Sub

Private Sub cmdGH_Click()
    MsgBox "Positives: Despite a lack of size (6-1 and 223 pounds) and bulk, is a remarkably powerful inside runner. ... Excellent in short-yardage and goal-line situations. ... Natural runner with superior vision and cutback skills. ... Solid receiver out of the backfield.                      Negatives: Too small to be a traditional fullback and hard to project him as an H-back. ... Durability. ... Ability to protect the ball while running.", , "Gilbert Harris (Arizona)"
End Sub

Private Sub cmdGM_Click()
    MsgBox "Positives: Former standout pass-catching tight end in college. ... Considerable athleticism, speed and play-making ability. ... Blocks well in the open field. ... Has all of the qualities necessary to develop into an effective H-back rather than a traditional fullback.                       Negatives: Doesn 't have enough size (6-foot-1 and 248 pounds) to fill the traditional fullback role as a powerful lead blocker or short-yardage/goal-line runner. ... Needs to develop skills as a ball carrier.", , "Garrett Mills (Tulsa)"
End Sub

Private Sub cmdLV_Click()
    MsgBox "Positives: Powerful runner who will take on tacklers, but has enough athleticism to avoid them. ... Effective in goal-line/short-yardage situations. ... Has good patience to allow blocks to develop and hits the hole quickly. ... Solid pass catcher.                        Negatives: At 5-11 and 239 pounds, relatively small to be a traditional fullback. ... Doesn't have the bulk or strength to hold his own in one-on-one blocking situations.", , "Lawrence Vickers (Colorado)"
End Sub

Private Sub cmdMB_Click()
    MsgBox "Positives: Probably the best blocking fullback in this year's class. ... Runs with power and keeps himself low, which is how fullbacks are supposed to run. ... Will battle for tough yards and continue to make progress after initial contact.                        Negatives: Needs to work on pass-protection and receiving skills. ... Almost no big-play threat due to limited speed and athleticism.", , "Matt Bernstein (Wisconsin)"
End Sub

Private Sub cmdNT_Click()
    MsgBox "Positives: Straight-line runner who packs a fairly good punch despite his relative lack of size (6-0 and 230 pounds). ... Should be able to add bulk and strength that would allow him to potentially make a decent impact in short-yardage and goal-line situations. ... A solid receiver out of the backfield, he is adept at running both short and intermediate routes. ... Good blocker at the point of attack                     Negatives: Lacks speed and athleticism. ... Needs to work on making blocks in the open field.", , "Naufahu Tahi (BYU)"
End Sub
