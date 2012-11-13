VERSION 5.00
Begin VB.Form startform 
   Caption         =   "Form1"
   ClientHeight    =   8775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14865
   LinkTopic       =   "Form1"
   Picture         =   "startform.frx":0000
   ScaleHeight     =   8775
   ScaleMode       =   0  'User
   ScaleWidth      =   14865
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdbib 
      Caption         =   "Bibliography"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   5
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton cmdquit 
      Caption         =   "Leave the alley"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12480
      TabIndex        =   4
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton cmdscoring 
      Caption         =   "Learn how to score a game of bowling"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   10800
      Picture         =   "startform.frx":1EE99
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4800
      Width           =   2415
   End
   Begin VB.CommandButton cmdterms 
      Caption         =   "Learn some bowling terms"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   7560
      Picture         =   "startform.frx":212F4
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4800
      Width           =   2175
   End
   Begin VB.CommandButton cmdteam 
      Caption         =   "Calculate team information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   4440
      Picture         =   "startform.frx":2374F
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4800
      Width           =   2175
   End
   Begin VB.CommandButton cmdavg 
      Caption         =   "Calculate your average"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   1320
      Picture         =   "startform.frx":25BAA
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4800
      Width           =   2175
   End
End
Attribute VB_Name = "startform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'BowlingProdject
'startform
'Zach Neumann
'3/30/2008
'This form helps navigate between the other forms. The over all purpose
'was to have a bowler be able to calculate their average and compare their average against pros,
'also to find some information about their team that can be updated weekly. Finally it allows
'a bowler to learn how to score a game and also score a game that they are playing, and allows them
'to learn some bowling terms

Private Sub cmdavg_Click()
    startform.Hide
    avgform.Show
    teamform.Hide
    termsform.Hide
End Sub

Private Sub cmdbib_Click()
startform.Hide
    scoringform.Hide
    teamform.Hide
    avgform.Hide
    termsform.Hide
    bibform.Show
End Sub

Private Sub cmdquit_Click()
End
End Sub

Private Sub cmdscoring_Click()
    startform.Hide
    scoringform.Show
    teamform.Hide
    avgform.Hide
    termsform.Hide
End Sub

Private Sub cmdteam_Click()
    teamform.Show
    avgform.Hide
    startform.Hide
    termsform.Hide
    scoringform.Hide
End Sub

Private Sub cmdterms_Click()
    teamform.Hide
    avgform.Hide
    startform.Hide
    termsform.Show
    scoringform.Hide
End Sub

Private Sub Command1_Click()
    teamform.Hide
    avgform.Hide
    startform.Hide
    termsform.Hide
    scoringform.Show
End Sub

