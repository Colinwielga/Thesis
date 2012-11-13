VERSION 5.00
Begin VB.Form frmfbs 
   BackColor       =   &H00404040&
   Caption         =   "Fullbacks"
   ClientHeight    =   8430
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10785
   LinkTopic       =   "Form1"
   ScaleHeight     =   8430
   ScaleWidth      =   10785
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picDisplay 
      Height          =   2895
      Left            =   600
      ScaleHeight     =   2835
      ScaleWidth      =   9555
      TabIndex        =   1
      Top             =   4560
      Width           =   9615
   End
   Begin VB.CommandButton cmdreturn 
      Caption         =   "Return to Offensive Players"
      Height          =   975
      Left            =   4200
      TabIndex        =   0
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label lblfbs 
      BackColor       =   &H000080FF&
      Caption         =   "Full Backs"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   7
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label lblvickers 
      Caption         =   "Lawrence Vickers"
      Height          =   375
      Left            =   8040
      TabIndex        =   6
      Top             =   480
      Width           =   1335
   End
   Begin VB.Image imgvickers 
      Height          =   1680
      Left            =   8040
      Picture         =   "frmfbs.frx":0000
      Top             =   840
      Width           =   1200
   End
   Begin VB.Label lbltahi 
      Caption         =   "Naufahu Tahi"
      Height          =   255
      Left            =   6000
      TabIndex        =   5
      Top             =   480
      Width           =   1335
   End
   Begin VB.Image imgtahi 
      Height          =   1680
      Left            =   5880
      Picture         =   "frmfbs.frx":6942
      Top             =   840
      Width           =   1890
   End
   Begin VB.Label lblharris 
      Caption         =   "Gilbert Harris"
      Height          =   255
      Left            =   4320
      TabIndex        =   4
      Top             =   600
      Width           =   1335
   End
   Begin VB.Image imgharris 
      Height          =   1170
      Left            =   4320
      Picture         =   "frmfbs.frx":10FC4
      Top             =   960
      Width           =   1320
   End
   Begin VB.Label lblbernstein 
      Caption         =   "Matt Bernstein"
      Height          =   255
      Left            =   2880
      TabIndex        =   3
      Top             =   480
      Width           =   1215
   End
   Begin VB.Image imgbernstein 
      Height          =   1695
      Left            =   2880
      Picture         =   "frmfbs.frx":16076
      Top             =   720
      Width           =   1170
   End
   Begin VB.Label lblmills 
      Caption         =   "Garrett Mills"
      Height          =   255
      Left            =   1560
      TabIndex        =   2
      Top             =   480
      Width           =   855
   End
   Begin VB.Image imgmills 
      Height          =   1485
      Left            =   1440
      Picture         =   "frmfbs.frx":1C8E4
      Top             =   840
      Width           =   1110
   End
End
Attribute VB_Name = "frmfbs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'2006 NFL Draft Simulator (Draft.vbp)
'frmfbs(frmfbs.frm)
'Andy Lyons
'March 24, 2006
'This form allows the user to look at the profiles of fullbacks that are eligible for the 2006 draft. By clicking on their photo it allows you to see their profile.

'returns user to main menu
Private Sub cmdreturn_Click()
    frmoffpositions.Show
    frmfbs.Hide
End Sub

Private Sub imgbernstein_Click()
    picDisplay.Cls
    picDisplay.Print "Height 6'1""", "Weight 260 "
    picDisplay.Print "Positives: Probably the best blocking fullback in this year's class. Runs with power and keeps himself low, which is how fullbacks"; Tab(11); "are supposed to run. Will battle for tough yards and continue to make progress after initial contact."
    picDisplay.Print "Negatives: Needs to work on pass-protection and receiving skills. Almost no big-play threat due to limited speed and athleticism. "

End Sub

Private Sub imgharris_Click()
    picDisplay.Cls
    picDisplay.Print "Height 6'1""", "Weight; 223 "
    picDisplay.Print "Positives: Despite a lack of size and bulk, is a remarkably powerful inside runner. Excellent in short-yardage and goal-line situations."; Tab(11); "Natural runner with superior vision and cutback skills. Solid receiver out of the backfield. "
    picDisplay.Print "Negatives: Too small to be a traditional fullback and hard to project him as an H-back. Durability. Ability to protect the ball while running."
        
End Sub

Private Sub imgmills_Click()
    picDisplay.Cls
    picDisplay.Print "Height 6'1""", "Weight 248"
    picDisplay.Print "Positives: Former standout pass-catching tight end in college. Considerable athleticism, speed and play-making ability. Blocks well in"; Tab(11); "the open field. Has all of the qualities necessary to develop into an effective H-back rather than a traditional fullback."
    picDisplay.Print "Negatives: Doesn't have enough size to fill the traditional fullback role as a powerful lead blocker or short-yardage/goal-line runner."; Tab(11); "Needs to develop skills as a ball carrier."
End Sub

Private Sub imgtahi_Click()
    picDisplay.Cls
    picDisplay.Print "Height 6'0""", "Weight 230"
    picDisplay.Print "Positives: Straight-line runner who packs a fairly good punch despite his relative lack of size. Should be able to add bulk and strength"; Tab(11); "that would allow him to potentially make a decent impact in short-yardage and goal-line situations. A solid receiver out of the"; Tab(11); "backfield, he is adept at running both short and intermediate routes. Good blocker at the point of attack. "
    picDisplay.Print "Negatives: Lacks speed and athleticism. Needs to work on making blocks in the open field."
End Sub

Private Sub imgvickers_Click()
    picDisplay.Cls
    picDisplay.Print "Height 5'11""", "Weight 239"
    picDisplay.Print "Positives: Powerful runner who will take on tacklers, but has enough athleticism to avoid them. Effective in goal-line/short-yardage"; Tab(11); "situations. Has good patience to allow blocks to develop and hits the hole quickly. Solid pass catcher. "
    picDisplay.Print "Negatives: At 5-11 and 239 pounds, relatively small to be a traditional fullback. Doesn't have the bulk or strength to hold his own in"; Tab(11); "one-on-one blocking situations."
End Sub
