VERSION 5.00
Begin VB.Form frmRB 
   Caption         =   "Form1"
   ClientHeight    =   8130
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8745
   LinkTopic       =   "Form1"
   Picture         =   "frmRB.frx":0000
   ScaleHeight     =   8130
   ScaleWidth      =   8745
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   615
      Left            =   6000
      TabIndex        =   5
      Top             =   7320
      Width           =   1335
   End
   Begin VB.CommandButton cmdLM 
      Caption         =   "Laurence Maroney"
      Height          =   615
      Left            =   1800
      TabIndex        =   4
      Top             =   480
      Width           =   2055
   End
   Begin VB.CommandButton cmdDW 
      Caption         =   "DeAngelo Williams"
      Height          =   615
      Left            =   1200
      TabIndex        =   3
      Top             =   3600
      Width           =   2055
   End
   Begin VB.CommandButton cmdJA 
      Caption         =   "Joseph Addai"
      Height          =   615
      Left            =   6240
      TabIndex        =   2
      Top             =   2160
      Width           =   2055
   End
   Begin VB.CommandButton cmdRB 
      Caption         =   "Reggie Bush"
      Height          =   615
      Left            =   6360
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
   Begin VB.CommandButton cmdLW 
      Caption         =   "LenDale White"
      Height          =   615
      Left            =   1080
      TabIndex        =   0
      Top             =   1920
      Width           =   2055
   End
End
Attribute VB_Name = "frmRB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'show profiles via message boxes
Option Explicit
Private Sub cmdBack_Click()
    frmRB.Hide
    frmAthletes.Show
End Sub

Private Sub cmdDW_Click()
    MsgBox "Positives: Good burst and quickness. ... Superb change-of-direction skills, body control and balance. ... Despite smallish frame (5-8 and 208 pounds), can power through defenders and be effective in short-yardage and goal-line situations. ... Reliable hands and ability to make catches in full stride makes him a strong threat as a receiver out of the backfield.                      Negatives: Lack of height, which causes problems in picking up the blitz. ... Although fast enough to outrun linebackers, does not have genuine breakaway speed. ... Durability; has a troubling history of injuries in college.", , "DeAngelo Williams (Memphis)"
End Sub

Private Sub cmdJA_Click()
    MsgBox "Positives: Exceptional speed, giving him the ability to go the distance on any carry. ... Tougher and more physical runner than one would expect from a back that stands 5-11 and weighs 215 pounds. ... Superb footwork. ... Natural receiver who not only runs precise routes and can catch the ball in full stride, but also has the ability to recognize coverages. ... Size is no impediment when it comes to picking up the blitz; willing to take on defensive linemen and results usually are favorable.                        Negatives: Although he was healthy the past two seasons, has a history of knee injuries. ... Not particularly elusive. ... Lack of size could be a drawback when needing to move the pile in short-yardage and goal-line situations.", , "Joseph Addai (LSU)"
End Sub

Private Sub cmdLM_Click()
    MsgBox "Positives: Shows superb vision and patience to allow blocks to develop. ... Decisive, usually runs with authority and won't dance too much. ... Despite relatively small frame (5-11 and 217 pounds), running style reflects surprisingly good power and explosiveness. ... Breakaway speed in the open field and outstanding body control. ... Excellent ball security.                        Negatives: Could develop greater toughness as a blocker and become more willing to engage in contact. ... Must have better awareness in picking up blitzes. ... Receiving skills need plenty of work in every respect.", , "Laurence Maroney (Minnesota)"
End Sub

Private Sub cmdLW_Click()
    MsgBox "Positives: Excellent size (6-1 and 238 pounds) to be a steady, every-down, power back. ... Won't hesitate to go through defenders rather than around them, but has enough body control and quickness to be highly effective running outside. ... Shows good patience to allow blocks to form and is a decisive runner. ... Excels in short-yardage and goal-line situations. ... Excellent ball security.                       Negatives: Not very elusive and doesn't have breakaway speed. ... Needs to improve blitz recognition and blocking technique in pass protection. ... Must work on receiving skills.", , "LenDale White (USC)"
End Sub

Private Sub cmdRB_Click()
    MsgBox "Positives: Has enough physical gifts to become one of the very best and most complete players ever at the position. ... Tremendous speed makes him a threat to go the distance each time he touches the ball, and almost impossible to catch from behind once he reaches the open field. ... Superb body control and elusiveness allow him to instantly bounce outside or find cutback lane. ... Great vision... Superior receiving skills, including ability to recognize coverages, run routes, create separation quickly, and make difficult catches. ... Although smallish frame (5-foot-10 and 201 pounds) limits ability to push the pile, still has enough power to be effective running between the tackles and can add bulk through diet and weight-training. Negatives: Lack of size and sharing rushing load at USC with LenDale White raise questions about his durability. ... Needs work on pass-protection techniques", , "Reggie Bush (USC)"
End Sub
