VERSION 5.00
Begin VB.Form frmtes 
   BackColor       =   &H00FFFF00&
   Caption         =   "Tight Ends"
   ClientHeight    =   8745
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11070
   LinkTopic       =   "Form1"
   ScaleHeight     =   8745
   ScaleWidth      =   11070
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picDisplay 
      Height          =   3375
      Left            =   480
      ScaleHeight     =   3315
      ScaleWidth      =   9555
      TabIndex        =   1
      Top             =   4200
      Width           =   9615
   End
   Begin VB.CommandButton cmdtes 
      Caption         =   "Return to Offensive Players"
      Height          =   855
      Left            =   4560
      TabIndex        =   0
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Tight Ends"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   7
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label lblklopfenstein 
      Caption         =   "Joe Klopfenstein"
      Height          =   255
      Left            =   8400
      TabIndex        =   6
      Top             =   360
      Width           =   1215
   End
   Begin VB.Image imgjoe 
      Height          =   1665
      Left            =   8520
      Picture         =   "frmtes.frx":0000
      Top             =   720
      Width           =   1020
   End
   Begin VB.Label lbllewis 
      Caption         =   "Marcedes Lewis"
      Height          =   255
      Left            =   6480
      TabIndex        =   5
      Top             =   360
      Width           =   1215
   End
   Begin VB.Image imglewis 
      Height          =   1665
      Left            =   6480
      Picture         =   "frmtes.frx":58B6
      Top             =   720
      Width           =   1245
   End
   Begin VB.Label lblfasano 
      Caption         =   "Anthony Fasano"
      Height          =   255
      Left            =   4560
      TabIndex        =   4
      Top             =   360
      Width           =   1335
   End
   Begin VB.Image imgfasano 
      Height          =   1710
      Left            =   4560
      Picture         =   "frmtes.frx":C63C
      Top             =   720
      Width           =   1305
   End
   Begin VB.Label lblpope 
      Caption         =   "Leonard Pope"
      Height          =   255
      Left            =   2880
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
   Begin VB.Image imgpope 
      Height          =   1530
      Left            =   3000
      Picture         =   "frmtes.frx":13C0E
      Top             =   720
      Width           =   1110
   End
   Begin VB.Label lbldavis 
      Caption         =   "Vernon Davis"
      Height          =   255
      Left            =   1320
      TabIndex        =   2
      Top             =   360
      Width           =   1095
   End
   Begin VB.Image imgdavis 
      Height          =   1590
      Left            =   1440
      Picture         =   "frmtes.frx":19590
      Top             =   720
      Width           =   945
   End
End
Attribute VB_Name = "frmtes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'2006 NFL Draft Simulator (Draft.vbp)
'frmtes(frmtes.frm)
'Andy Lyons
'March 24, 2006
'This form allows the user to view the Tight Ends that are eligible for the 2006 draft. By clicking on their photo, it allows the user to view their profile.
'returns user to main menu
Private Sub cmdtes_Click()
    frmoffpositions.Show
    frmtes.Hide
End Sub
Private Sub imgdavis_Click()
    picDisplay.Cls
    picDisplay.Print "Height 6'3""", "Weight 254"
    picDisplay.Print "Positives: Off-the-charts performance in Scouting Combine drills only enhanced the status of a player already widely regarded as"; Tab(11); "the draft's best at his position. Has a staggering combination of outstanding size, speed, strength, and athleticism."; Tab(11); "Dependable hands.Excellent route-runner with knowledge to find openings against zone- and man-to-man coverage,"; Tab(11); "and necessary burst to create separation in his patterns. Capable of consistently turning short catches into long gains. "
    picDisplay.Print "Negatives: Needs to work on his blocking skills. Despite strength, doesn't seem to know how to use hands properly when engaging"; Tab(11); "with linebackers. "
End Sub

Private Sub imgfasano_Click()
    picDisplay.Cls
    picDisplay.Print "Height 6'5""", "Weight 255"
    picDisplay.Print "Positives: Catches the ball well in traffic and on the run. Tremendous intelligence allows him to consistently find openings"; Tab(11); "against zone and man-to-man coverage. Worked well in a pro-style offense that should enhance his readiness for the NFL. "; Tab(11); "Goodinitial burst helps compensate for general lack of speed. Displays solid blocking technique. Superior hustle and work"; Tab(11); "ethic. "
    picDisplay.Print "Negatives: Lack of speed and athleticism. Doesn't make many would-be tacklers miss and is not particularly effective on seam"; Tab(12); "routes, where tight ends usually make their greatest impact."
End Sub

Private Sub imgjoe_Click()
    picDisplay.Cls
    picDisplay.Print "Height 6'5""", "Weight 250"
    picDisplay.Print "Positives: Nice combination of height and speed. Runs good, crisp routes. Will battle for the ball. Intelligence. Durability."; Tab(11); "Strong work ethic."
    picDisplay.Print "Negatives: Doesn't consistently get good release from the line. Must add bulk and strength to 250-pound frame to become a more"; Tab(12); "effective blocker."
End Sub

Private Sub imglewis_Click()
    picDisplay.Cls
    picDisplay.Print "Height 6'6""", "Weight 256"
    picDisplay.Print "Positives: Concentration when catching the ball. Makes good use of 6-6 frame, particularly on jump balls. Good route-runner."; Tab(11); "Solid blocker"
    picDisplay.Print "Negatives: Doesn't get much explosion off the line. Has problems releasing on the snap. Lacks a natural running stride. "

End Sub

Private Sub imgpope_Click()
    picDisplay.Cls
    picDisplay.Print "Height 6'7""", "Weight 250"
    picDisplay.Print "Positives: Tall target with large hands and long arms that he uses well to keep defenders at bay. Impressive athletic display at"; Tab(11); "Scouting Combine. Superior speed that is capable of stretching coverage. Shows strong explosion off the line of scrimmage."; Tab(11); "Excels at finding soft spots in zone coverage. "
    picDisplay.Print "Negatives: Needs to add some bulk, which figures to be fairly easy given his frame. Has problems sometimes with defenders who"; Tab(11); "are able to get their hands on him near the line."
End Sub
