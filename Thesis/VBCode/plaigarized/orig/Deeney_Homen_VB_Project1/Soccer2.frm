VERSION 5.00
Begin VB.Form Stats 
   BackColor       =   &H00800000&
   Caption         =   "Form2"
   ClientHeight    =   8415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12615
   LinkTopic       =   "Form2"
   Picture         =   "Soccer2.frx":0000
   ScaleHeight     =   8415
   ScaleWidth      =   12615
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   495
      Left            =   11880
      TabIndex        =   16
      Top             =   7080
      Width           =   615
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   11880
      TabIndex        =   15
      Top             =   7800
      Width           =   615
   End
   Begin VB.TextBox txtInput 
      Height          =   495
      Left            =   7800
      TabIndex        =   13
      Top             =   2640
      Width           =   2175
   End
   Begin VB.CommandButton cmdDisplay 
      BackColor       =   &H000000C0&
      Caption         =   "Top 20 scorers from World Cup"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton cmdGoalstats 
      BackColor       =   &H0000C0C0&
      Caption         =   "Match your own goal stats against the best!"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CommandButton cmdAvgshots 
      BackColor       =   &H0000C0C0&
      Caption         =   "What was the average number of shots taken?"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6600
      Width           =   1695
   End
   Begin VB.CommandButton cmdAccuracy 
      BackColor       =   &H0000C0C0&
      Caption         =   "Calculate each player's shot accuracy"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6600
      Width           =   1695
   End
   Begin VB.CommandButton cmdLeastgoals 
      BackColor       =   &H0000C0C0&
      Caption         =   "Who scored the least number of goals?"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5520
      Width           =   1695
   End
   Begin VB.CommandButton cmdOver10 
      BackColor       =   &H0000C0C0&
      Caption         =   "Who were the ball hogs?"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4440
      Width           =   1695
   End
   Begin VB.CommandButton cmdComment 
      BackColor       =   &H0000C0C0&
      Caption         =   "Did these top players shoot too much or too little?"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CommandButton cmdReadall 
      BackColor       =   &H000000C0&
      Caption         =   "Click here to Begin"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   360
      Width           =   1695
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFFFF&
      Height          =   4815
      Left            =   600
      ScaleHeight     =   4755
      ScaleWidth      =   6675
      TabIndex        =   1
      Top             =   1560
      Width           =   6735
   End
   Begin VB.CommandButton cmdNextForm 
      BackColor       =   &H000000C0&
      Caption         =   "We've seen the players from 2006; Now gear up for the powerhouse of teams taking the stage in 2010"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6600
      Width           =   3495
   End
   Begin VB.Label Label4 
      Caption         =   $"Soccer2.frx":41E72
      Height          =   1095
      Left            =   7800
      TabIndex        =   14
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H000000C0&
      Caption         =   "Finished?==>"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   12
      Top             =   6720
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000C0&
      Caption         =   "Next click here for results ===>"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6600
      TabIndex        =   11
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Top Player statistics from 2006 World Cup play"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   720
      TabIndex        =   10
      Top             =   360
      Width           =   3735
   End
End
Attribute VB_Name = "Stats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: WorldCup
'Form Name: Stats
'Author: Brian Deeney and Nick Homen
'Date written: 2-20-10
'Objective: The purpose of this form is to use a data file in an array of player statistics that allows the user to explore the information in many different ways, all the while learning more about some important soccer players from around the globe.

'This is a great button that allows the user to see just how accurate these top scorers were by calculating each player's percent of goals vs. attempts.
Private Sub cmdAccuracy_Click()
Dim I As Integer, Accuracy(1 To 100) As Single, Percent(1 To 100) As Single
    picResults.Cls
    picResults.Print "Team"; Tab(20); "LastName"; Tab(40); "FirstName"; Tab(52); "Goals"; Tab(60); "Shots"; Tab(70); "% of Shots Made"
    picResults.Print "--------------------------------------------------------------------------------------------------------------------------------------------"
    picResults.Print
    'This code calculates the decimal of accuracy by dividing shots taken and goals made
    For I = 1 To CTR
        Accuracy(I) = Shots(I) / Goals(I)
    Next I
    'This code puts the number into a percent format (out of 100)
    For I = 1 To CTR
        Percent(I) = 100 / Accuracy(I)
    Next I
    For I = 1 To CTR
        picResults.Print Team(I); Tab(20); LastName(I); Tab(40); FirstName(I); Tab(52); Goals(I); Tab(60); Shots(I); Tab(75); FormatNumber((Percent(I)), 1)
    Next I
End Sub

'This button calculates the average number of shots taken by a player by adding up the total number of shots and dividing by the amount of players
Private Sub cmdAvgshots_Click()
Dim I As Integer, Total As Integer, Average As Integer
    For I = 1 To CTR
        Total = Total + Shots(I)
    Next I
    
    Average = Total / CTR
    MsgBox ("The average amount of shots taken by a player during the Cup was " & Average & " shots.")
End Sub

'This button allows you to go back and access the welcome screen
Private Sub cmdBack_Click()
Stats.Hide
StartUp.Show
End Sub

'This button is entertaining because it throws labels on each player judging how many shots they took themselves.
Private Sub cmdComment_Click()
picResults.Cls
    picResults.Print "Last Name"; Tab(20); "Shots"; Tab(30); "Comment"
    picResults.Print "------------------------------------------------------------------------"
    
    Dim Comment As String, K As Integer
    
    'Too many shots by a player and they're a ball hog, too few and they're too timid, just right and they're a good team player!
    For K = 1 To CTR
    Select Case Shots(K)
        Case 10 To 15
            Comment = "Certified Ball hog"
        Case 8 To 9
            Comment = "Should pass more often"
        Case 5 To 7
            Comment = "Great team player"
        Case 1 To 5
            Comment = "Take more shots!"
        Case Is < 1
            Comment = "Too timid for the pros.."
        Case Else
            Comment = "What are you thinking?!"
    End Select
    picResults.Print LastName(K); Tab(20); Shots(K); Tab(30); Comment
    Next K
End Sub

'This button takes all of the data that was put into arrays and displays it in the picture box with the appropriate headings
Private Sub cmdDisplay_Click()
Dim J As Integer
    
    picResults.Print "Team"; Tab(20); "LastName"; Tab(40); "FirstName"; Tab(52); "Goals"; Tab(60); "Shots"
    picResults.Print "---------------------------------------------------------------------------------------------------------------"
    picResults.Print
    
    For J = 1 To CTR
        picResults.Print Team(J); Tab(20); LastName(J); Tab(40); FirstName(J); Tab(52); Goals(J); Tab(60); Shots(J)
    Next J
  'This code allows the buttons that will utilize the now accessible data to be used
  cmdGoalstats.Enabled = True
  cmdComment.Enabled = True
  cmdOver10.Enabled = True
  cmdLeastgoals.Enabled = True
  cmdAvgshots.Enabled = True
  cmdAccuracy.Enabled = True
  cmdNextForm.Enabled = True
End Sub

'This button utilizes a textbox to compare data from the user with data from the array
Private Sub cmdGoalstats_Click()
picResults.Cls
    picResults.Print "Team"; Tab(20); "LastName"; Tab(40); "FirstName"; Tab(52); "Goals"; Tab(60); "Shots"
    picResults.Print "---------------------------------------------------------------------------------------------------------------"
    picResults.Print
     
    Dim InputGoals As Integer
    InputGoals = txtInput.Text
        
    Dim J As Integer, Found As Boolean, I As Integer
        J = 0
        Found = False
        
        'This code looks for players who surpass the number of goals inputed by the user
        For J = 1 To CTR
            If Goals(J) > InputGoals Then
                Found = True
                I = I + 1
                picResults.Print Team(J); Tab(20); LastName(J); Tab(40); FirstName(J); Tab(52); Goals(J); Tab(60); Shots(J)
            End If
        Next J
        
    MsgBox "There are " & I & " players who scored more than " & InputGoals & " goals", , "message"
End Sub

'This button allows the user to view the data from a reverse point of view, organizing from least goals to most
Private Sub cmdLeastgoals_Click()
Dim Pos As Integer, pass As Integer, Tempfirstname As String, TempShots As Integer, J As Integer, Templastname As String, TempGoals As String, TempTeam As String

'This chunk of code sorts the array using a bubble sort from least to most goals
For pass = 1 To CTR - 1
    For Pos = 1 To CTR - pass
        If Goals(Pos) > Goals(Pos + 1) Then
            TempGoals = Goals(Pos)
            Goals(Pos) = Goals(Pos + 1)
            Goals(Pos + 1) = TempGoals
            
            TempShots = Shots(Pos)
            Shots(Pos) = Shots(Pos + 1)
            Shots(Pos + 1) = TempShots
            
            Tempfirstname = FirstName(Pos)
            FirstName(Pos) = FirstName(Pos + 1)
            FirstName(Pos + 1) = Tempfirstname
            
            Templastname = LastName(Pos)
            LastName(Pos) = LastName(Pos + 1)
            LastName(Pos + 1) = Templastname
            
            TempTeam = Team(Pos)
            Team(Pos) = Team(Pos + 1)
            Team(Pos + 1) = TempTeam
        End If
    Next Pos
Next pass

picResults.Cls
    picResults.Print "Team"; Tab(20); "LastName"; Tab(40); "FirstName"; Tab(52); "Goals"; Tab(60); "Shots"
    picResults.Print "-----------------------------------------------------------------------------------------------------------------"
    picResults.Print

For J = 1 To CTR
             picResults.Print Team(J); Tab(20); LastName(J); Tab(40); FirstName(J); Tab(52); Goals(J); Tab(60); Shots(J)
    Next J
End Sub

'This button allows you to go to the form that lets the user check out team jerseys
Private Sub cmdNextForm_Click()
Stats.Hide
Jerseys.Show
End Sub

'Quit
Private Sub cmdQuit_Click()
End
End Sub

'This button searches the array for any player who shot more than 10 times and displays the data in the picture box and message box
Private Sub cmdOver10_Click()
Dim J As Integer, Found As Boolean, I As Integer
        J = 0
        Found = False
        picResults.Cls
        picResults.Print "Team"; Tab(20); "LastName"; Tab(40); "FirstName"; Tab(52); "Goals"; Tab(60); "Shots"
        picResults.Print "---------------------------------------------------------------------------------------------------------------"
        picResults.Print
        For J = 1 To CTR
            If Shots(J) > 10 Then
                Found = True
                I = I + 1
                picResults.Print Team(J); Tab(20); LastName(J); Tab(40); FirstName(J); Tab(52); Goals(J); Tab(60); Shots(J)
            End If
        Next J
            If Found = True Then
                MsgBox "There are " & I & " players who took more than 10 shots on goal"
            Else
                MsgBox "There are no players who took more than 10 shots on goal"
            End If
End Sub

'This button starts the form by arranging all of the data in the text file into arrays.  It is the only button accessible at this time.
Private Sub cmdReadall_Click()
 Open App.Path & "\Soccer.txt" For Input As #1
    
    Do While Not EOF(1)
        CTR = CTR + 1
        Input #1, Team(CTR), LastName(CTR), FirstName(CTR), Goals(CTR), Shots(CTR)
    Loop
  
    MsgBox ("Statistical data retrieval complete")
    cmdDisplay.Enabled = True
    Close #1
End Sub

