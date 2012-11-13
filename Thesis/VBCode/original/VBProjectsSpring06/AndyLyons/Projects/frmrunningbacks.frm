VERSION 5.00
Begin VB.Form frmrunningbacks 
   BackColor       =   &H0000C0C0&
   Caption         =   "Runningbacks"
   ClientHeight    =   8295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10665
   LinkTopic       =   "Form1"
   ScaleHeight     =   8295
   ScaleWidth      =   10665
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      Height          =   3615
      Left            =   480
      ScaleHeight     =   3555
      ScaleWidth      =   9675
      TabIndex        =   2
      Top             =   4440
      Width           =   9735
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "Back to Offensive Players"
      Height          =   975
      Left            =   4680
      TabIndex        =   0
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label lblrbs 
      Caption         =   "Running Backs"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   7
      Top             =   2760
      Width           =   2895
   End
   Begin VB.Label lbladdai 
      Caption         =   "Joseph Addai"
      Height          =   255
      Left            =   8520
      TabIndex        =   6
      Top             =   600
      Width           =   1335
   End
   Begin VB.Image imgaddai 
      Height          =   1545
      Left            =   8400
      Picture         =   "frmrunningbacks.frx":0000
      Top             =   960
      Width           =   1500
   End
   Begin VB.Label lblwilliams 
      Caption         =   "Deangelo Williams"
      Height          =   255
      Left            =   6360
      TabIndex        =   5
      Top             =   600
      Width           =   1695
   End
   Begin VB.Image imgwilliams 
      Height          =   1635
      Left            =   6360
      Picture         =   "frmrunningbacks.frx":78F6
      Top             =   960
      Width           =   1860
   End
   Begin VB.Label lblwhite 
      Caption         =   "Lendale White"
      Height          =   255
      Left            =   4800
      TabIndex        =   4
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Laurence Maroney"
      Height          =   255
      Left            =   3000
      TabIndex        =   3
      Top             =   480
      Width           =   1335
   End
   Begin VB.Image imgmaroney 
      Height          =   1620
      Left            =   2880
      Picture         =   "frmrunningbacks.frx":1179C
      Top             =   1080
      Width           =   1635
   End
   Begin VB.Image imgbush 
      Height          =   1905
      Left            =   720
      Picture         =   "frmrunningbacks.frx":1A23E
      Top             =   840
      Width           =   1350
   End
   Begin VB.Image imgwhite 
      Height          =   1815
      Left            =   4800
      Picture         =   "frmrunningbacks.frx":22970
      Top             =   840
      Width           =   1320
   End
   Begin VB.Label lblbush 
      Caption         =   "Reggie Bush"
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "frmrunningbacks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'2006 NFL Draft Simulator (Draft.vbp)
'frmrunningbacks(frmrunningbacks.frm)
'Andy Lyons
'March 24, 2006
'This form allows the user to look at all the profiles for the runningbacks eligible for the 2006 draft. By clicking on each players image, the user can read it.
'returns user to main menu
Private Sub cmdback_Click()
    frmoffpositions.Show
    frmrunningbacks.Hide
End Sub
Private Sub imgbush_Click()
    picResults.Cls
    picResults.Print "Height 5'10""", "Weight 201"
    picResults.Print "Positives: Has enough physical gifts to become one of the very best and most complete players ever at the position."; Tab(11); "Tremendous speed makes him a threat to go the distance each time he touches the ball, and almost impossible to catch"; Tab(11); "from behind once he reaches the open field. Superb body control and elusiveness allow him to instantly bounce outside or"; Tab(11); "find cutback lane.Great vision to find the right hole. Superior receiving skills, including ability to recognize coverages, run"; Tab(11); "routes, create separation quickly, and make difficult catches."
    picResults.Print "Negatives: Not many for the presumptive top overall choice, but you can always find flaws even with the most"; Tab(11); "talented players. Lack of size and sharing rushing load at USC with LenDale White raise questions about his durability. Needs"; Tab(11); "work on pass-protection techniques, especially gaining leverage when taking on larger defensive linemen. "

End Sub

Private Sub imgmaroney_Click()
    picResults.Cls
    picResults.Print "Height 5'11""", "Weight 217"
    picResults.Print "Positives: Shows superb vision and patience to allow blocks to develop.Decisive, usually runs with authority and"; Tab(11); "won 't dance too much. Despite relatively small frame, running style reflects surprisingly good power and"; Tab(11); "explosiveness.Breakaway speed in the open field and outstanding body control.Excellent ball security"
    picResults.Print "Negatives:Could develop greater toughness as a blocker and become more willing to engage in contact.Must have better"; Tab(11); "awareness in picking up blitzes.Receiving skills need plenty of work in every respect. "
End Sub

Private Sub imgwhite_Click()
    picResults.Cls
    picResults.Print "Height 6'1""", "Weight 238"
    picResults.Print "Positives: Excellent size to be a steady, every-down, power back. Won't hesitate to go through defenders rather than around"; Tab(11); "them, but has enough body control and quickness to be highly effective running outside.Shows good"; Tab(11); "patience to allow blocks to form and is a decisive runner.Excels in short-yardage and goal-line situations."; Tab(11); "Excellent ball security."
    picResults.Print "Negatives: Not very elusive and doesn't have breakaway speed. Needs to improve blitz recognition and blocking technique in pass"; Tab(11); "protection. Must work on receiving skills. "
End Sub

Private Sub imgwilliams_Click()
    picResults.Cls
    picResults.Print "Height 5'8""", "Weight 208 "
    picResults.Print "Positives: Good burst and quickness. Superb change-of-direction skills, body control and balance. Despite smallish frame, can power"; Tab(11); "through defenders and be effective in short-yardage and goal-line situations. Reliable hands"; Tab(11); "and ability to make catches in full stride makes him a strong threat as a receiver out of the backfield. "
    picResults.Print "Negatives:Lack of height, which causes problems in picking up the blitz. Although fast enough to outrun linebackers, does not have"; Tab(11); "genuine breakaway speed. Durability: has a troubling history of injuries in college."
End Sub
Private Sub imgaddai_Click()
    picResults.Cls
    picResults.Print "Height 5'11""", "Weight 215"
    picResults.Print "Positives: Exceptional speed, giving him the ability to go the distance on any carry. Tougher and more physical runner than one would"; Tab(11); "expect from a back his size. Superb footwork. Natural receiver who not only runs precise routes"; Tab(11); "and can catch the ball in full stride, but also has the ability to recognize coverages. Size is no"; Tab(11); "impediment when it comes to picking up the blitz; willing to take on defensive linemen and results usually are favorable. "
    picResults.Print "Negatives: Although he was healthy the past two seasons, has a history of knee injuries. Not particularly elusive. Lack of size could be"; Tab(11); "a drawback when needing to move the pile in short-yardage and goal-line situations."
End Sub
