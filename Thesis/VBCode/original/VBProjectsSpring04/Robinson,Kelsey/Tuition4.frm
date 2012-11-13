VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00C0C000&
   Caption         =   "Form4"
   ClientHeight    =   8760
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10950
   LinkTopic       =   "Form4"
   ScaleHeight     =   8760
   ScaleWidth      =   10950
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Click on the college of your choice"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   120
      Width           =   4815
   End
   Begin VB.PictureBox picRichU 
      Height          =   975
      Left            =   840
      Picture         =   "Tuition4.frx":0000
      ScaleHeight     =   915
      ScaleWidth      =   3675
      TabIndex        =   9
      Top             =   4320
      Width           =   3735
   End
   Begin VB.PictureBox picCC 
      Height          =   1575
      Left            =   5280
      Picture         =   "Tuition4.frx":4B54
      ScaleHeight     =   1515
      ScaleWidth      =   2595
      TabIndex        =   8
      Top             =   4080
      Width           =   2655
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Go Back One Slide"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6840
      Width           =   2535
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Go to the Next Slide"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6840
      Width           =   2655
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6840
      Width           =   2295
   End
   Begin VB.PictureBox picStOlaf 
      Height          =   1335
      Left            =   5040
      Picture         =   "Tuition4.frx":7F87
      ScaleHeight     =   1275
      ScaleWidth      =   2235
      TabIndex        =   4
      Top             =   1200
      Width           =   2295
   End
   Begin VB.PictureBox picUST 
      Height          =   1095
      Left            =   600
      Picture         =   "Tuition4.frx":8E99
      ScaleHeight     =   1035
      ScaleWidth      =   6915
      TabIndex        =   3
      Top             =   2880
      Width           =   6975
   End
   Begin VB.PictureBox picUC 
      Height          =   855
      Left            =   4800
      Picture         =   "Tuition4.frx":C3AD
      ScaleHeight     =   795
      ScaleWidth      =   2355
      TabIndex        =   2
      Top             =   5760
      Width           =   2415
   End
   Begin VB.PictureBox picGus 
      Height          =   1095
      Left            =   240
      Picture         =   "Tuition4.frx":CB53
      ScaleHeight     =   1035
      ScaleWidth      =   3195
      TabIndex        =   1
      Top             =   5520
      Width           =   3255
   End
   Begin VB.PictureBox picSJUCSB 
      Height          =   1815
      Left            =   480
      Picture         =   "Tuition4.frx":D4CF
      ScaleHeight     =   1755
      ScaleWidth      =   3435
      TabIndex        =   0
      Top             =   960
      Width           =   3495
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Deciding on a College
' Form 2 (Tuition2)
' Kelsey Robinson
' March 10th, 2004
' This form lets the user get some "quick facts" about each college by clicking on its logo.

Private Sub picSJUCSB_Click()
MsgBox ("The University of St. John's and the College of Saint Benedict's is located just north of St. Cloud.")
MsgBox ("St. Ben's is known for their strong Nursing and Education programs, while St. John's is known for their strong Business program.")
End Sub

Private Sub picStOlaf_Click()
MsgBox ("The College of St. Olaf is located in Northfield, Minnesota, south of the Twin Cities.")
MsgBox ("St. Olaf is known for their strong programs in Music and Science.")
End Sub



Private Sub picUST_Click()
MsgBox ("The University of St. Thomas is located in St. Paul, Minnesota")
MsgBox ("St. Thomas is known for their well developed Business and Pre-Law majors.")
End Sub

Private Sub picRichU_Click()
MsgBox ("Rich University is located in the heart of New York City")
MsgBox ("Rich is recognized around the world for their strong Medical, Business, Science, Political Science and Engineering fields.")
End Sub

Private Sub picCC_Click()
MsgBox ("Clown College is located in Brooklyn Park, Minnesota")
MsgBox ("Clown is known for their homemaking and entertainment programs.")
End Sub

Private Sub picGus_Click()
MsgBox ("Gustavus Adolphus College is located in St. Peter, Minnesota, just north of Mankato")
MsgBox ("Gustavus is known for their strong Political Science programs.")
End Sub

Private Sub picUC_Click()
MsgBox ("The University of Colorado-Boulder is located 60 miles north of Denver, Colorado.")
MsgBox ("UC-Boulder offers 85 majors at the bachelor's level, 70 at the master's level and 50 at the doctoral level.")
End Sub

Private Sub cmdBack_Click()
Form4.Hide
Form3.Show
End Sub
Private Sub cmdNext_Click()
Form4.Hide
Form5.Show
End Sub
Private Sub cmdQuit_Click()
End
End Sub

