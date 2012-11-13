VERSION 5.00
Begin VB.Form Predictionform2 
   BackColor       =   &H00FFFF80&
   Caption         =   "Best Rebounder"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9810
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   9810
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picresults5 
      BackColor       =   &H00FF0000&
      Height          =   1575
      Left            =   8040
      Picture         =   "Predictionform2.frx":0000
      ScaleHeight     =   1515
      ScaleWidth      =   1275
      TabIndex        =   12
      Top             =   240
      Width           =   1335
   End
   Begin VB.PictureBox picresults4 
      Height          =   1575
      Left            =   6000
      Picture         =   "Predictionform2.frx":19D3
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   11
      Top             =   240
      Width           =   1575
   End
   Begin VB.PictureBox picresults3 
      Height          =   1575
      Left            =   4320
      Picture         =   "Predictionform2.frx":3084
      ScaleHeight     =   1515
      ScaleWidth      =   1275
      TabIndex        =   10
      Top             =   240
      Width           =   1335
   End
   Begin VB.PictureBox picresults2 
      Height          =   1575
      Left            =   2280
      Picture         =   "Predictionform2.frx":46A8
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   9
      Top             =   240
      Width           =   1575
   End
   Begin VB.PictureBox picresults1 
      Height          =   1575
      Left            =   360
      Picture         =   "Predictionform2.frx":93E6
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   8
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton followingform 
      BackColor       =   &H0080C0FF&
      Caption         =   "Click to Move to Next Form"
      Height          =   975
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6000
      Width           =   1335
   End
   Begin VB.OptionButton Michael 
      BackColor       =   &H000000FF&
      Caption         =   "Michael"
      Height          =   495
      Left            =   480
      TabIndex        =   6
      Top             =   2520
      Width           =   1455
   End
   Begin VB.PictureBox picresults6 
      BackColor       =   &H008080FF&
      Height          =   2175
      Left            =   840
      ScaleHeight     =   2115
      ScaleWidth      =   6435
      TabIndex        =   5
      Top             =   4920
      Width           =   6495
   End
   Begin VB.OptionButton Shaq 
      BackColor       =   &H0000FFFF&
      Caption         =   "Shaquelle"
      Height          =   495
      Left            =   2400
      TabIndex        =   4
      Top             =   2520
      Width           =   1335
   End
   Begin VB.OptionButton Kareem 
      BackColor       =   &H0000FFFF&
      Caption         =   "Kareem"
      Height          =   495
      Left            =   4440
      TabIndex        =   3
      Top             =   2520
      Width           =   1335
   End
   Begin VB.OptionButton Magic 
      BackColor       =   &H0000FFFF&
      Caption         =   "Magic"
      Height          =   495
      Left            =   6240
      TabIndex        =   2
      Top             =   2520
      Width           =   1215
   End
   Begin VB.OptionButton Kevin 
      BackColor       =   &H00FF0000&
      Caption         =   "Kevin"
      Height          =   495
      Left            =   8040
      TabIndex        =   1
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton rebounder 
      BackColor       =   &H00C000C0&
      Caption         =   "Display if Player selected is the best rebounder"
      Height          =   1095
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label player5 
      BackColor       =   &H00FF0000&
      Caption         =   "Kevin Garnett"
      Height          =   375
      Left            =   8040
      TabIndex        =   18
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label player4 
      BackColor       =   &H0000FFFF&
      Caption         =   "Magic Johnson"
      Height          =   375
      Left            =   6120
      TabIndex        =   17
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label player3 
      BackColor       =   &H0000FFFF&
      Caption         =   "Kareem Abdul Jabbar"
      Height          =   375
      Left            =   4200
      TabIndex        =   16
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label player2 
      BackColor       =   &H0000FFFF&
      Caption         =   "Shaquelle O'Neal"
      Height          =   375
      Left            =   2280
      TabIndex        =   15
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label player1 
      BackColor       =   &H000000FF&
      Caption         =   "Michael Jordan"
      Height          =   375
      Left            =   360
      TabIndex        =   14
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label instructions 
      BackColor       =   &H000080FF&
      Caption         =   $"Predictionform2.frx":DA1E
      Height          =   1215
      Left            =   2040
      TabIndex        =   13
      Top             =   3360
      Width           =   4695
   End
End
Attribute VB_Name = "Predictionform2"
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
                   ' is the best rebounder on the team
                   
Private Sub followingform_Click()
'change forms that can be seen by user
    Predictionform2.Hide
    tableform.Show
End Sub

Private Sub Form_Load()
'select options available to user
    followingform.Enabled = False
    rebounder.Enabled = False
End Sub

Private Sub rebounder_Click()
'see if shooter selected is best rebounder based on number of rebounds

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
    picresults6.Print "Best Rebounder based on number of rebounds."
    picresults6.Print "************************************************************************************************"
'determine if player user selected is best shooter
'and print results (program knows person in position CTR is best shooter)
  CTR = 2
    If X = CTR Then
        picresults6.Print "You are correct.  Shaquelle O'Neal is the best rebounder on the team."
      ElseIf X <> CTR Then
        picresults6.Print "Sorry, you are incorrect.  Shaquelle O'Neal is the best rebounder on the team."
    End If
    
'switch options avaialable to user
    rebounder.Enabled = False
    followingform.Enabled = True
End Sub
Private Sub Kareem_Click()
    'allow program to determine if michael is best rebounder
        rebounder.Enabled = True
End Sub

Private Sub Kevin_Click()
    'allow program to determine if michael is best rebounder
        rebounder.Enabled = True
End Sub
Private Sub Magic_Click()
    'allow program to determine if michael is best rebounder
        rebounder.Enabled = True
End Sub

Private Sub Michael_Click()
    'allow program to determine if michael is best rebounder
        rebounder.Enabled = True
End Sub

Private Sub Shaq_Click()
    'allow program to determine if michael is best rebounder
        rebounder.Enabled = True
End Sub
