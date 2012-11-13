VERSION 5.00
Begin VB.Form BasketballPlayersInput 
   BackColor       =   &H000080FF&
   Caption         =   "Basketball Players Statistics Input"
   ClientHeight    =   6690
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   9420
   LinkTopic       =   "Form1"
   ScaleHeight     =   6690
   ScaleWidth      =   9420
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton readfile 
      BackColor       =   &H00FF00FF&
      Caption         =   "Click here to read the file of statistics for the basketball game."
      Height          =   1335
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3840
      Width           =   5295
   End
   Begin VB.PictureBox picresults5 
      BackColor       =   &H00FF0000&
      Height          =   1575
      Left            =   7680
      Picture         =   "BasketballPlayersInput.frx":0000
      ScaleHeight     =   1515
      ScaleWidth      =   1275
      TabIndex        =   5
      Top             =   600
      Width           =   1335
   End
   Begin VB.PictureBox picresults4 
      Height          =   1575
      Left            =   5760
      Picture         =   "BasketballPlayersInput.frx":19D3
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   4
      Top             =   600
      Width           =   1575
   End
   Begin VB.PictureBox picresults3 
      Height          =   1575
      Left            =   4080
      Picture         =   "BasketballPlayersInput.frx":3084
      ScaleHeight     =   1515
      ScaleWidth      =   1275
      TabIndex        =   3
      Top             =   600
      Width           =   1335
   End
   Begin VB.PictureBox picresults2 
      Height          =   1575
      Left            =   2040
      Picture         =   "BasketballPlayersInput.frx":46A8
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   2
      Top             =   600
      Width           =   1575
   End
   Begin VB.PictureBox picresults1 
      Height          =   1575
      Left            =   120
      Picture         =   "BasketballPlayersInput.frx":93E6
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton Nextform 
      BackColor       =   &H0000FF00&
      Caption         =   "Click to Move to Prediction Form."
      Height          =   975
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5880
      Width           =   3015
   End
   Begin VB.Label player5 
      BackColor       =   &H00FF0000&
      Caption         =   "Kevin Garnett"
      Height          =   375
      Left            =   7680
      TabIndex        =   10
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label player4 
      BackColor       =   &H0000FFFF&
      Caption         =   "Magic Johnson"
      Height          =   375
      Left            =   5880
      TabIndex        =   9
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label player3 
      BackColor       =   &H0000FFFF&
      Caption         =   "Kareem Abdul Jabbar"
      Height          =   375
      Left            =   3960
      TabIndex        =   8
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label player2 
      BackColor       =   &H0000FFFF&
      Caption         =   "Shaquelle O'Neal"
      Height          =   375
      Left            =   2040
      TabIndex        =   7
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label player1 
      BackColor       =   &H000000FF&
      Caption         =   "Michael Jordan"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Width           =   1575
   End
End
Attribute VB_Name = "BasketballPlayersInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : Basketball Game Statistics (BasketballPlayersInput.vbp)
'Form Name : BasketballPlayersInput (BasketballPlayersInput.frm)
'Author : Anthony Mayerhofer
'Date Written : March 15, 2004
'Purpose of Project : To read the file for the statistics
                      'from the basketball game
                      'then have the user predict
                      'the most efficient shooter and rebounder
                      'and finally display the results
                      'of several comparisons
                      'to evaluate players on several criteria
'Purpose of Form : To read the statistics from the game
                   'and arrange that information into arrays
                   'then proceed to prediction form
                   

Option Explicit
'is a command to forces the programmer to declare all variables
'for that form before the variables can be used
  
Private Sub Form_Load()
'set options available to user
    readfile.Enabled = True
    Nextform.Enabled = False
End Sub

Private Sub Nextform_Click()
'determine which forms can be seen by user
    BasketballPlayersInput.Hide
    Predictionform.Show
    tableform.Hide
    Predictionform2.Hide
End Sub

Private Sub readfile_Click()
'reads statistics from game into 5 parrallel arrays
Dim Path As String

'set path
    Path = "N:\CS130\handin\Mayerhofer, Anthony\"
'open file
    Open Path & "statistics.txt" For Input As #1
'create arrays
    For CTR = 1 To 5
        Input #1, names(CTR), points(CTR), shootingpercentage(CTR), rebounds(CTR)
    Next CTR
    
    'close file
        Close #1
'disable option
    readfile.Enabled = False
    Nextform.Enabled = True
    
End Sub
