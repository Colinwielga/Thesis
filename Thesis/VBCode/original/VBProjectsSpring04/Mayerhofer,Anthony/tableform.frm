VERSION 5.00
Begin VB.Form tableform 
   BackColor       =   &H00FF00FF&
   Caption         =   "Statistics Table Form"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10125
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   10125
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton favoriteplayer 
      BackColor       =   &H0080FFFF&
      Caption         =   "Click here to see if your favorite player is on the team."
      Height          =   495
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5640
      Width           =   2535
   End
   Begin VB.PictureBox picresults1 
      Height          =   1575
      Left            =   240
      Picture         =   "tableform.frx":0000
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   12
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox picresults2 
      Height          =   1575
      Left            =   2280
      Picture         =   "tableform.frx":4638
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   11
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox picresults3 
      Height          =   1575
      Left            =   4320
      Picture         =   "tableform.frx":9376
      ScaleHeight     =   1515
      ScaleWidth      =   1275
      TabIndex        =   10
      Top             =   120
      Width           =   1335
   End
   Begin VB.PictureBox picresults4 
      Height          =   1575
      Left            =   6240
      Picture         =   "tableform.frx":A99A
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   9
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox picresults5 
      BackColor       =   &H00FF0000&
      Height          =   1575
      Left            =   8280
      Picture         =   "tableform.frx":C04B
      ScaleHeight     =   1515
      ScaleWidth      =   1275
      TabIndex        =   8
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton review 
      BackColor       =   &H00FFFF80&
      Caption         =   "Review Tables (Makes tables options available again)"
      Height          =   735
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6240
      Width           =   1695
   End
   Begin VB.CommandButton allstatistics 
      BackColor       =   &H0000C000&
      Caption         =   "Table of Players with all statistics from the game displayed"
      Height          =   855
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton cmdrebounds 
      BackColor       =   &H000080FF&
      Caption         =   "Table of Players from highest to lowest based on rebounds"
      Height          =   855
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CommandButton cmdshootingpercentage 
      BackColor       =   &H0000FFFF&
      Caption         =   "Table of Players from highest to lowest based on shooting percentage"
      Height          =   855
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton cmdpoint 
      BackColor       =   &H00FF8080&
      Caption         =   "Table of Players from highest to lowest based on points"
      Height          =   855
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3720
      Width           =   1695
   End
   Begin VB.PictureBox picresults7 
      BackColor       =   &H0080FFFF&
      Height          =   4575
      Left            =   120
      ScaleHeight     =   4515
      ScaleWidth      =   6195
      TabIndex        =   1
      Top             =   2400
      Width           =   6255
   End
   Begin VB.CommandButton Quit 
      BackColor       =   &H000000FF&
      Caption         =   "Quit"
      Height          =   735
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6240
      Width           =   1455
   End
   Begin VB.Label player1 
      BackColor       =   &H000000FF&
      Caption         =   "Michael Jordan"
      Height          =   375
      Left            =   240
      TabIndex        =   17
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label player2 
      BackColor       =   &H0000FFFF&
      Caption         =   "Shaquelle O'Neal"
      Height          =   375
      Left            =   2280
      TabIndex        =   16
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label player3 
      BackColor       =   &H0000FFFF&
      Caption         =   "Kareem Abdul Jabbar"
      Height          =   375
      Left            =   4200
      TabIndex        =   15
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label player4 
      BackColor       =   &H0000FFFF&
      Caption         =   "Magic Johnson"
      Height          =   375
      Left            =   6240
      TabIndex        =   14
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label player5 
      BackColor       =   &H00FF0000&
      Caption         =   "Kevin Garnett"
      Height          =   375
      Left            =   8160
      TabIndex        =   13
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label lastinstructions 
      BackColor       =   &H0080FF80&
      Caption         =   $"tableform.frx":DA1E
      Height          =   1215
      Left            =   6600
      TabIndex        =   2
      Top             =   2400
      Width           =   3375
   End
End
Attribute VB_Name = "tableform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : Basketball Game Statistics (BasketballPlayersInput.vbp)
'Form Name : tableform (tableform.frm)
'Author : Anthony Mayerhofer
'Date Written : March 15, 2004
'Purpose of Project : To read the file for the statistics
                      'from the basketball game
                      'then have the user predict
                      'the most efficient shooter and rebounder
                      'and finally display the results
                      'of several comparisons
                      'to evaluate players on several criteria
'Purpose of Form : to have coach (user) evaluate players
                    'based on several tables of statistics
                    'from the game
Option Explicit
'forces user to declare all variables for this form
Dim Pass As Integer, Comp As Integer, K As Integer
Dim tempnames As String, temppoints As Integer
Dim tempshooting As Single, temprebounds As Integer
                                       
Private Sub allstatistics_Click()
'clear picture box of previous results
    picresults7.Cls
    
'print a table with each player and all statistics
    picresults7.Print "All players will all statistics."
    picresults7.Print "************************************************************************************************"
    picresults7.Print "Player", Tab(30); "Points", "Shooting Percentage", "Rebounds"
    
    For K = 1 To 5
        picresults7.Print names(K), Tab(30); points(K), FormatPercent(shootingpercentage(K)), , rebounds(K)
    Next K

'change options available to user
    allstatistics.Enabled = False
    favoriteplayer.Enabled = True
    
End Sub

Private Sub cmdrebounds_Click()
'clear picture box of previous results
    picresults7.Cls
'produce table of players ranking them on shootingpercentage
'from highest to lowest
'set counter to last person in list
    CTR = 5

For Pass = 1 To CTR - 1
    For Comp = 1 To CTR - Pass
        If rebounds(Comp) < rebounds(Comp + 1) Then
        
            'switch names
            tempnames = names(Comp)
            names(Comp) = names(Comp + 1)
            names(Comp + 1) = tempnames
            
            'and also switch points
            temppoints = points(Comp)
            points(Comp) = points(Comp + 1)
            points(Comp + 1) = temppoints
            
            'and also switch shooting percentage
            tempshooting = shootingpercentage(Comp)
            shootingpercentage(Comp) = shootingpercentage(Comp + 1)
            shootingpercentage(Comp + 1) = tempshooting
            
            'and also switch rebounds
            temprebounds = rebounds(Comp)
            rebounds(Comp) = rebounds(Comp + 1)
            rebounds(Comp + 1) = temprebounds
            
            
        End If
    Next Comp
Next Pass

'print table with players and points in descending order
    picresults7.Print "Players in descending order by rebounds from the game."
    picresults7.Print "************************************************************************************************"
    picresults7.Print "Player", Tab(40); "Rebounds"
    
    For K = 1 To 5
        picresults7.Print names(K), Tab(40); rebounds(K)
    Next K

'switch options available to user
    cmdrebounds.Enabled = False
    allstatistics.Enabled = True
End Sub

Private Sub cmdshootingpercentage_Click()
'clear picture box of previous results
    picresults7.Cls
'produce table of players ranking them on shootingpercentage
'from highest to lowest
'set counter to last person in list
CTR = 5

For Pass = 1 To CTR - 1
    For Comp = 1 To CTR - Pass
        If shootingpercentage(Comp) < shootingpercentage(Comp + 1) Then
        
            'switch names
            tempnames = names(Comp)
            names(Comp) = names(Comp + 1)
            names(Comp + 1) = tempnames
            
            'and also switch points
            temppoints = points(Comp)
            points(Comp) = points(Comp + 1)
            points(Comp + 1) = temppoints
            
            'and also switch shooting percentage
            tempshooting = shootingpercentage(Comp)
            shootingpercentage(Comp) = shootingpercentage(Comp + 1)
            shootingpercentage(Comp + 1) = tempshooting
            
            'and also switch rebounds
            temprebounds = rebounds(Comp)
            rebounds(Comp) = rebounds(Comp + 1)
            rebounds(Comp + 1) = temprebounds
            
            
        End If
    Next Comp
Next Pass

'print table with players and points in descending order
    picresults7.Print "Players in descending order by shooting percentages from the game."
    picresults7.Print "************************************************************************************************"
    picresults7.Print "Player", Tab(40); "Shooting Percentage"
    
    For K = 1 To 5
        picresults7.Print names(K), Tab(40); FormatPercent(shootingpercentage(K))
    Next K

'switch options available to user
    cmdshootingpercentage.Enabled = False
    cmdrebounds.Enabled = True
End Sub

Private Sub favoriteplayer_Click()
'clear results of previous button in picture box
    picresults7.Cls

'ask user for favorite player and see if player is on team

Dim favoriteplayer As String, notfound As Boolean, J As Integer

'get player from user
    favoriteplayer = InputBox("Enter the first and last name of your favorite player to see if they are on the team.")

'compare against names array and search for player entered by user
notfound = True
J = 0
    Do While J < 5
        J = J + 1
            If names(J) = favoriteplayer Then
                notfound = False
            End If
    Loop

'print results
    If notfound Then
        picresults7.Print ; favoriteplayer; " is not on the team."
      Else
        picresults7.Print ; favoriteplayer; " is on the team."
    End If
    
'change options available to user
    
    Quit.Enabled = True
    review.Enabled = True
End Sub

Private Sub Form_Load()
    Quit.Enabled = False
    review.Enabled = False
    cmdpoint.Enabled = True
    cmdshootingpercentage.Enabled = False
    cmdrebounds.Enabled = False
    allstatistics.Enabled = False
    favoriteplayer.Enabled = False
    
End Sub

Private Sub cmdpoint_Click()
'produce table of players ranking them on points scored
'from highest to lowest
'set counter to last person in list
CTR = 5

For Pass = 1 To CTR - 1
    For Comp = 1 To CTR - Pass
        If points(Comp) < points(Comp + 1) Then
        
            'switch names
            tempnames = names(Comp)
            names(Comp) = names(Comp + 1)
            names(Comp + 1) = tempnames
            
            'and also switch points
            temppoints = points(Comp)
            points(Comp) = points(Comp + 1)
            points(Comp + 1) = temppoints
            
            'and also switch shooting percentage
            tempshooting = shootingpercentage(Comp)
            shootingpercentage(Comp) = shootingpercentage(Comp + 1)
            shootingpercentage(Comp + 1) = tempshooting
            
            'and also switch rebounds
            temprebounds = rebounds(Comp)
            rebounds(Comp) = rebounds(Comp + 1)
            rebounds(Comp + 1) = temprebounds
            
            
        End If
    Next Comp
Next Pass

'print table with players and points in descending order
    picresults7.Print "Players in descending order by point totals from the game."
    picresults7.Print "************************************************************************************************"
    picresults7.Print "Player", Tab(40); "Points"
    
    For K = 1 To 5
        picresults7.Print names(K), Tab(40); points(K)
    Next K

'switch options available to user
    cmdpoint.Enabled = False
    cmdshootingpercentage.Enabled = True
    
End Sub

Private Sub Quit_Click()
'end program
    End
End Sub


Private Sub review_Click()
'make table options available again and clear results
    picresults7.Cls
    Quit.Enabled = False
    cmdpoint.Enabled = True
    review.Enabled = False
    favoriteplayer.Enabled = False
    
    
End Sub
