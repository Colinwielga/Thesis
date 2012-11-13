VERSION 5.00
Begin VB.Form Predictionform 
   BackColor       =   &H00C0C000&
   Caption         =   "Best Shooter"
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9825
   LinkTopic       =   "Form1"
   ScaleHeight     =   7500
   ScaleWidth      =   9825
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton shooter 
      BackColor       =   &H00C000C0&
      Caption         =   "Display if Player selected is the best shooter"
      Height          =   1095
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4680
      Width           =   1335
   End
   Begin VB.OptionButton Kevin 
      BackColor       =   &H00FF0000&
      Caption         =   "Kevin"
      Height          =   495
      Left            =   7920
      TabIndex        =   16
      Top             =   2760
      Width           =   1215
   End
   Begin VB.OptionButton Magic 
      BackColor       =   &H0000FFFF&
      Caption         =   "Magic"
      Height          =   495
      Left            =   6120
      TabIndex        =   15
      Top             =   2760
      Width           =   1215
   End
   Begin VB.OptionButton Kareem 
      BackColor       =   &H0000FFFF&
      Caption         =   "Kareem"
      Height          =   495
      Left            =   4320
      TabIndex        =   14
      Top             =   2760
      Width           =   1335
   End
   Begin VB.OptionButton Shaq 
      BackColor       =   &H0000FFFF&
      Caption         =   "Shaquelle"
      Height          =   495
      Left            =   2280
      TabIndex        =   13
      Top             =   2760
      Width           =   1335
   End
   Begin VB.PictureBox picresults6 
      BackColor       =   &H0080FF80&
      Height          =   2175
      Left            =   720
      ScaleHeight     =   2115
      ScaleWidth      =   6435
      TabIndex        =   12
      Top             =   5160
      Width           =   6495
   End
   Begin VB.OptionButton Michael 
      BackColor       =   &H000000FF&
      Caption         =   "Michael"
      Height          =   495
      Left            =   360
      TabIndex        =   11
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton futureform 
      BackColor       =   &H0080C0FF&
      Caption         =   "Click to Move to Next Form"
      Height          =   975
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6240
      Width           =   1335
   End
   Begin VB.PictureBox picresults1 
      Height          =   1575
      Left            =   240
      Picture         =   "Predictionform.frx":0000
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   4
      Top             =   480
      Width           =   1575
   End
   Begin VB.PictureBox picresults2 
      Height          =   1575
      Left            =   2160
      Picture         =   "Predictionform.frx":4638
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   3
      Top             =   480
      Width           =   1575
   End
   Begin VB.PictureBox picresults3 
      Height          =   1575
      Left            =   4200
      Picture         =   "Predictionform.frx":9376
      ScaleHeight     =   1515
      ScaleWidth      =   1275
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.PictureBox picresults4 
      Height          =   1575
      Left            =   5880
      Picture         =   "Predictionform.frx":A99A
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   1
      Top             =   480
      Width           =   1575
   End
   Begin VB.PictureBox picresults5 
      BackColor       =   &H00FF0000&
      Height          =   1575
      Left            =   7920
      Picture         =   "Predictionform.frx":C04B
      ScaleHeight     =   1515
      ScaleWidth      =   1275
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label instructions 
      BackColor       =   &H000080FF&
      Caption         =   $"Predictionform.frx":DA1E
      Height          =   1215
      Left            =   1920
      TabIndex        =   17
      Top             =   3600
      Width           =   4695
   End
   Begin VB.Label player1 
      BackColor       =   &H000000FF&
      Caption         =   "Michael Jordan"
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label player2 
      BackColor       =   &H0000FFFF&
      Caption         =   "Shaquelle O'Neal"
      Height          =   375
      Left            =   2160
      TabIndex        =   9
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label player3 
      BackColor       =   &H0000FFFF&
      Caption         =   "Kareem Abdul Jabbar"
      Height          =   375
      Left            =   4080
      TabIndex        =   8
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label player4 
      BackColor       =   &H0000FFFF&
      Caption         =   "Magic Johnson"
      Height          =   375
      Left            =   6000
      TabIndex        =   7
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label player5 
      BackColor       =   &H00FF0000&
      Caption         =   "Kevin Garnett"
      Height          =   375
      Left            =   7920
      TabIndex        =   6
      Top             =   2280
      Width           =   1455
   End
End
Attribute VB_Name = "Predictionform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : Basketball Game Statistics (BasketballPlayersInput.vbp)
'Form Name : Predictionform (Preditcionform.frm)
'Author : Anthony Mayerhofer
'Date Written : March 15, 2004
'Purpose of Project : To read the file for the statistics
                      'from the basketball game
                      'then have the user predict
                      'the most efficient shooter and rebounder
                      'and finally display the results
                      'of several comparisons
                      'to evaluate players on several criteria
'Purpose of Form : ' have user click a particular player
                   ' and see if that player
                   ' is the best shooter on the team
                   
Private Sub Form_Load()
'choose options available to user
    shooter.Enabled = False
    futureform.Enabled = False
End Sub

Private Sub Kareem_Click()
    'allow program to determine if michael is best shooter
        shooter.Enabled = True
End Sub

Private Sub Kevin_Click()
    'allow program to determine if michael is best shooter
        shooter.Enabled = True
End Sub

Private Sub futureform_Click()
'change which forms can be seen by user
    Predictionform.Hide
    Predictionform2.Show
End Sub

Private Sub Magic_Click()
    'allow program to determine if michael is best shooter
        shooter.Enabled = True
End Sub

Private Sub Michael_Click()
    'allow program to determine if michael is best shooter
        shooter.Enabled = True
End Sub

Private Sub Shaq_Click()
    'allow program to determine if michael is best shooter
        shooter.Enabled = True
End Sub

Private Sub shooter_Click()
'see if shooter selected is best shooter based on shooting percentage

If Michael = True Then
    X = 1
  ElseIf Shaq = True Then
    X = 2
  ElseIf Kareem = True Then
    X = 3
  ElseIf Magic = True Then
    X = 4
  ElseIf Kevin = True Then
    X = 5
End If

'print header for picture box
    picresults6.Print "Best Shooter by Shooting Percentage."
    picresults6.Print "************************************************************************************************"
'determine if player user selected is best shooter
'and print results (program knows person in position CTR is best shooter)
  CTR = 5
    If X = CTR Then
        picresults6.Print "You are correct.  Kevin Garnett is the best shooter on the team."
      ElseIf X <> CTR Then
        picresults6.Print "Sorry, you are incorrect.  Kevin Garnett is the best shooter on the team."
    End If
    
'switch options avaialable to user
    shooter.Enabled = False
    futureform.Enabled = True
    
End Sub
