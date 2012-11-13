VERSION 5.00
Begin VB.Form frmC 
   BackColor       =   &H00004040&
   Caption         =   "Form1"
   ClientHeight    =   4905
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6480
   LinkTopic       =   "Form1"
   Picture         =   "frmC.frx":0000
   ScaleHeight     =   4905
   ScaleWidth      =   6480
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRC 
      Caption         =   "Ryan Cook"
      Height          =   615
      Left            =   4560
      TabIndex        =   5
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton cmdJS 
      Caption         =   "Jason Spitz"
      Height          =   615
      Left            =   360
      TabIndex        =   4
      Top             =   2160
      Width           =   1935
   End
   Begin VB.CommandButton cmdGE 
      Caption         =   "Greg Eslinger"
      Height          =   615
      Left            =   360
      TabIndex        =   3
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton cmdMD 
      Caption         =   "Mike DeGory"
      Height          =   615
      Left            =   2400
      TabIndex        =   2
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton cmdNM 
      Caption         =   "Nick Mangold "
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   615
      Left            =   4560
      TabIndex        =   0
      Top             =   4200
      Width           =   615
   End
End
Attribute VB_Name = "frmC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'display message boxes containing profile
Option Explicit
Private Sub cmdBack_Click()
    frmC.Hide
    frmAthletes.Show
End Sub

Private Sub cmdGE_Click()
    MsgBox "Positives: Good size (6-3 and 292 pounds) and athleticism. ... Superior techniques. ... Outstanding footwork that allows him to get consistently good leverage. ... Makes excellent use of hands to separate from defenders. ... Shows superb recognition of stunts and blitzes and consistently makes proper line calls. ... Perfect fit for any zone-blocking running scheme.                     Negatives: Needs to develop more lower-body strength. ... Although he is built just right for lateral movement in a zone-blocking scheme, he could benefit from adding some bulk and strength, especially when he has to take on a power-oriented inside rusher.", , "Greg Eslinger (Minnesota)"
End Sub

Private Sub cmdJS_Click()
    MsgBox "Positives: Good size (6-3 and 313 pounds), speed and toughness. ... Makes good use of hands to gain separation. ... Solid run blocker. ... Good recognition of stunts and blitzes. ... Versatility, with more experience at guard than center.                      Negatives: A converted guard who still is learning the center position. ... Must work on improving change-of-direction ability and maintaining balance when blocking on the move.", , "Jason Spitz (Lousville)"
End Sub

Private Sub cmdMD_Click()
    MsgBox "Positives: Excellent size (6-5 and 305 pounds) and arm length. ... Superior technique in pass protection. ... Excellent overall strength. ... Great desire to improve his game, as reflected by steady development throughout his collegiate career. ... Shows a great deal of intelligence and savvy in stunt and blitz recognition, and consistently makes proper line calls. ... Willing and coachable enough to work at guard and center, which will make him a more valuable backup early in his NFL career.                       Negatives: Athleticism. ... Needs to improve body control and balance when making blocks in the open field.", , "Mike DeGory (Florida)"
End Sub

Private Sub cmdNM_Click()
    MsgBox "Positives: Good size (6-foot-3 and 300 pounds) and upper-body strength. ... Shows considerable toughness and will battle to the whistle. ... Possesses intelligence to recognize stunts and blitzes, and consistently makes proper line calls. ... Takes good angles on blocks. ... Does a nice job of pulling and trapping. ... Sets up quickly in pass protection, which gives him an edge vs. speedy inside rushers.                     Negatives: Must enhance lower-body strength to deal with massive, bull-rushing tackles. ... Needs to work on change-of-direction skills to become better able to block defenders on the move.", , "Nick Mangold (Ohio State)"
End Sub
Private Sub cmdRC_Click()
    MsgBox "Positives: Tremendous size (6-6 and 328 pounds). ... Good upper-body strength. ... Considerable arm length allows him to consistently lock up pass rushers.                     Negatives: Needs to add lower-body strength. ... Has work to do in gaining better stunt/blitz awareness. ... Must do a better job of maintaining leverage and adjusting to make blocks in the open field.", , "Ryan Cook (New Mexico)"
End Sub
