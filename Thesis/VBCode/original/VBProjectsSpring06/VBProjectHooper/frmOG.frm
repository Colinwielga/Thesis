VERSION 5.00
Begin VB.Form frmOG 
   Caption         =   "Form1"
   ClientHeight    =   8745
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   Picture         =   "frmOG.frx":0000
   ScaleHeight     =   8745
   ScaleWidth      =   8985
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdback 
      Caption         =   "Back"
      Height          =   615
      Left            =   1320
      TabIndex        =   5
      Top             =   7560
      Width           =   1815
   End
   Begin VB.CommandButton cmdMJ 
      Caption         =   "Max Jean-Gilles"
      Height          =   615
      Left            =   4680
      TabIndex        =   4
      Top             =   960
      Width           =   2295
   End
   Begin VB.CommandButton cmdDJ 
      Caption         =   "Davin Joseph"
      Height          =   615
      Left            =   6240
      TabIndex        =   3
      Top             =   1920
      Width           =   2295
   End
   Begin VB.CommandButton cmdCS 
      Caption         =   "Charles Spencer"
      Height          =   615
      Left            =   6480
      TabIndex        =   2
      Top             =   3120
      Width           =   2295
   End
   Begin VB.CommandButton cmdRS 
      Caption         =   "Rob Sims"
      Height          =   615
      Left            =   6000
      TabIndex        =   1
      Top             =   4320
      Width           =   2295
   End
   Begin VB.CommandButton cmdKB 
      Caption         =   "Kevin Boothe"
      Height          =   615
      Left            =   5280
      TabIndex        =   0
      Top             =   5520
      Width           =   2175
   End
End
Attribute VB_Name = "frmOG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'show profiles via message boxes
Option Explicit
Private Sub cmdBack_Click()
    frmOG.Hide
    frmAthletes.Show
End Sub

Private Sub cmdCS_Click()
    MsgBox "Positives: Good size (6-4 and 352 pounds). ... Highly athletic, which allows him to be effective in small areas as well as in the open field. ... Explodes out of stance and gets to linebackers and safeties quickly. ... Good arm extension and powerful enough to rock defenders with hand punches. ... Remarkable progress since switching from defensive tackle as a junior, and impressed many NFL scouts during Senior Bowl workouts.                        Negatives: Lack of offensive line experience. ... Must work at taking better advantage of size and strength when run-blocking at point of attack.", , "Charles Spencer (Pittsburgh)"
End Sub

Private Sub cmdDJ_Click()
    MsgBox "Positives: Plays with considerable power that allows him to consistently rock defenders with hand punches. ... Does a good job of maintaining leverage. ... Shows an aggressive mentality, especially in pass protection. ... Superior mobility and body control allow him to pull and trap well, and make blocks in the open field.                        Negatives: At 6-2, lacks ideal height, although 311-pound frame is plenty wide for an NFL guard. ... Blitz and stunt recognition.", , "Davin Joseph (Oklahoma)"
End Sub

Private Sub cmdKB_Click()
    MsgBox "Positives: Good size (6-4 and 316 pounds). ... Nice combination of bulk and athleticism. ... Long arms and powerful hands that he uses well to knock back and lock onto defenders. ... Superior lateral movement. ... Intelligence and blitz/stunt awareness.                       Negatives: Durability after suffering ankle and hand injuries in high school and again in college. ... Lack of top-level speed is a concern, especially in terms of getting to linebackers and safeties in the open field.", , "Kevin Boothe (Cornell)"
End Sub

Private Sub cmdMJ_Click()
    MsgBox "Positives: Packs tremendous bulk in a 6-foot-3, 355-pound frame. ... Won't often be overpowered. ... Long arms and good use of hands to lock onto defenders. ...Shows good body control and balance, especially in the open field. ... Impressive footwork and quickness off the snap as well as laterally. ... Establishes leverage quickly.                       Negatives: Athleticism leaves something to be desired and stamina and conditioning need work. ... Needs to improve speed on pulls and traps.", , "Max Jean-Gilles (Georgia)"
End Sub

Private Sub cmdRS_Click()
    MsgBox "Positives: Toughness and strength, especially when it comes to locking onto defenders. ... Plenty of width in 307-pound frame. ... Quick feet. ... Takes good angles on blocks. ... Recognition of blitz and stunts. ... Versatility, with collegiate experience at tackle.                     Negatives: At 6-2, lacks ideal height. ... Agility. ... Must improve physical conditioning.", , "Rob Sims (Ohio State)"
End Sub
