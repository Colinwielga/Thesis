VERSION 5.00
Begin VB.Form ProjectFace2 
   BackColor       =   &H80000001&
   Caption         =   "Player Statistics"
   ClientHeight    =   8010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10695
   LinkTopic       =   "Form2"
   ScaleHeight     =   8010
   ScaleWidth      =   10695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPointsPerMin 
      Caption         =   "Sort By Most Points Per Minute In A Game"
      Enabled         =   0   'False
      Height          =   735
      Left            =   2160
      TabIndex        =   18
      Top             =   3840
      Width           =   1815
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   735
      Left            =   2160
      TabIndex        =   16
      Top             =   6720
      Width           =   1815
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "Return to Previous Screen"
      Height          =   735
      Left            =   2160
      TabIndex        =   15
      Top             =   5760
      Width           =   1815
   End
   Begin VB.CommandButton cmdFavorite 
      Caption         =   "Find your Favorite Timberwolves'    Statistics"
      Enabled         =   0   'False
      Height          =   735
      Left            =   2160
      TabIndex        =   14
      Top             =   4800
      Width           =   1815
   End
   Begin VB.CommandButton cmdPoints 
      Caption         =   "Sort By Most Points Per Game"
      Enabled         =   0   'False
      Height          =   735
      Left            =   2160
      TabIndex        =   13
      Top             =   2880
      Width           =   1815
   End
   Begin VB.CommandButton cmdAssists 
      Caption         =   "Sort By Most Assists Per Game"
      Enabled         =   0   'False
      Height          =   735
      Left            =   120
      TabIndex        =   12
      Top             =   6720
      Width           =   1815
   End
   Begin VB.CommandButton cmdRebounds 
      Caption         =   "Sort By Most  Rebounds Per Game"
      Enabled         =   0   'False
      Height          =   735
      Left            =   120
      TabIndex        =   11
      Top             =   5760
      Width           =   1815
   End
   Begin VB.CommandButton cmdFieldGoal 
      Caption         =   "Sort By Highest Field Goal %"
      Enabled         =   0   'False
      Height          =   735
      Left            =   120
      TabIndex        =   10
      Top             =   4800
      Width           =   1815
   End
   Begin VB.CommandButton cmdMinutes 
      Caption         =   "Sort By Most Minutes Per Game"
      Enabled         =   0   'False
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   3840
      Width           =   1815
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "Show Player Statistics"
      Height          =   735
      Left            =   120
      TabIndex        =   8
      Top             =   2880
      Width           =   1815
   End
   Begin VB.PictureBox picresults 
      Height          =   4095
      Left            =   4080
      ScaleHeight     =   4035
      ScaleWidth      =   6435
      TabIndex        =   7
      Top             =   2760
      Width           =   6495
   End
   Begin VB.PictureBox Picture6 
      Height          =   1935
      Left            =   8880
      Picture         =   "ProjectFace2.frx":0000
      ScaleHeight     =   1875
      ScaleWidth      =   1515
      TabIndex        =   6
      Top             =   720
      Width           =   1575
   End
   Begin VB.PictureBox Picture5 
      Height          =   1935
      Left            =   7200
      Picture         =   "ProjectFace2.frx":8534
      ScaleHeight     =   1875
      ScaleWidth      =   1515
      TabIndex        =   5
      Top             =   720
      Width           =   1575
   End
   Begin VB.PictureBox Picture4 
      Height          =   1935
      Left            =   5280
      Picture         =   "ProjectFace2.frx":EA9E
      ScaleHeight     =   1875
      ScaleWidth      =   1635
      TabIndex        =   4
      Top             =   720
      Width           =   1695
   End
   Begin VB.PictureBox Picture3 
      Height          =   1935
      Left            =   3480
      Picture         =   "ProjectFace2.frx":155F6
      ScaleHeight     =   1875
      ScaleWidth      =   1515
      TabIndex        =   3
      Top             =   720
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      Height          =   1935
      Left            =   1800
      Picture         =   "ProjectFace2.frx":1E4B5
      ScaleHeight     =   1875
      ScaleWidth      =   1515
      TabIndex        =   2
      Top             =   720
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   1935
      Left            =   120
      Picture         =   "ProjectFace2.frx":24BC8
      ScaleHeight     =   1875
      ScaleWidth      =   1515
      TabIndex        =   1
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Designer: Ben Lewis"
      Height          =   255
      Left            =   9000
      TabIndex        =   19
      Top             =   7800
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   $"ProjectFace2.frx":2C518
      Height          =   615
      Left            =   4080
      TabIndex        =   17
      Top             =   7080
      Width           =   6495
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000008&
      Caption         =   "        Player  Statistics                                          Player Statistics"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   7935
   End
End
Attribute VB_Name = "ProjectFace2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Minnesota Timberwolves Basketball Season 2003-2004
'Project 1 (Project1.vbp)
'ProjectFace2 (ProjectFace2.Frm)
'Ben Lewis
'March 13, 2004
'The Purpose of this project is to gain knowledge of the statistics from various players of the Minnesota Basketball team for 2003-2004 season
'The Purpose of this form is to visually show the user what the program executes
'Also, this form is a guide for debugging problems with variables, and making sure all variables used in computations are declared
Option Explicit




Private Sub cmdAssists_Click()
'This command will sort Timberwolves statistics from Most Assists Per Game to Least Assists Per Game
'Open "N:\CS130\handin\Lewis, Ben\twolves.txt" For Input As #1
Open PATH & "twolves.txt" For Input As #1
    For I = 1 To 13
        Input #1, Names(I), Numbers(I), MinutesPerGame(I), FieldGoals(I), ReboundsPerGame(I), AssistsPerGame(I), PointsPerGame(I)
    Next I
Close
'This will clear Picture Box of any previous data
picresults.Cls
'Print Header
picresults.Print "Statistics Sorted by Most Assists Per Game"
picresults.Print
picresults.Print "Player"; Tab(45); "Assists/Game"
picresults.Print "************************************************************************************************************"
'This sorts the array and displays the results in the picture box by most Assists Per Game
I = 13
    For Pass = 1 To I - 1
        For Comp = 1 To I - Pass
            If AssistsPerGame(Comp) < AssistsPerGame(Comp + 1) Then
            
                TempAssistsPerGame = AssistsPerGame(Comp)
                AssistsPerGame(Comp) = AssistsPerGame(Comp + 1)
                AssistsPerGame(Comp + 1) = TempAssistsPerGame
                
                TempNames = Names(Comp)
                Names(Comp) = Names(Comp + 1)
                Names(Comp + 1) = TempNames
            
            End If
        Next Comp
    Next Pass
picresults.Print
'Prints Data that is desired for command
    For I = 1 To 13
        picresults.Print Names(I); Tab(45); FormatNumber(AssistsPerGame(I))
    Next I
            
End Sub





Private Sub cmdFavorite_Click()
'This Command will allow the user to find favorite players statistics
'Open "N:\CS130\handin\Lewis, Ben\twolves.txt" For Input As #1
Open PATH & "twolves.txt" For Input As #1
    For I = 1 To 13
        Input #1, Names(I), Numbers(I), MinutesPerGame(I), FieldGoals(I), ReboundsPerGame(I), AssistsPerGame(I), PointsPerGame(I)
    Next I
Close
'This will clear Picture Box of any previous data
picresults.Cls
I = 0
'Allows user to type favorite Timberwolves player in input box that pops up
N = InputBox("Enter your favorite Timberwolves Name, ", "Player's Name")
Found = False
Do While (Not Found) And (I <= 12)
    I = I + 1
    If N = Names(I) Then
        'Prints Header
        picresults.Print
        picresults.Print "Your Favorite player is, "; N
        picresults.Print
        picresults.Print "His Statistics are:"
        picresults.Print
        picresults.Print " Player"; Tab(27); "#"; Tab(34); "MPG"; Tab(49); "FG%"; Tab(59); "RPG"; Tab(71); "APG"; Tab(81); "PPG"
        picresults.Print "************************************************************************************************************"
        picresults.Print
        'Prints Data that is desired for command
        picresults.Print Tab(1); Names(I); Tab(26); Numbers(I); Tab(34); FormatNumber(MinutesPerGame(I)); Tab(47); FormatPercent(FieldGoals(I)); Tab(59); FormatNumber(ReboundsPerGame(I)); Tab(71); FormatNumber(AssistsPerGame(I)); Tab(81); FormatNumber(PointsPerGame(I))
        Found = True
    End If
Loop
    If Not Found Then
        MsgBox "Sorry but the name you entered is not a member of the team.", , "Error"
    End If

End Sub





Private Sub cmdFieldGoal_Click()
'This command will sort Timberwolves statistics from Highest Field Goal % to lowest Field Goal %
'Open "N:\CS130\handin\Lewis, Ben\twolves.txt" For Input As #1
Open PATH & "twolves.txt" For Input As #1
    For I = 1 To 13
        Input #1, Names(I), Numbers(I), MinutesPerGame(I), FieldGoals(I), ReboundsPerGame(I), AssistsPerGame(I), PointsPerGame(I)
    Next I
Close
'This will clear Picture Box of any previous data
picresults.Cls
'Prints Header
picresults.Print "Statistics Sorted by Highest Field Goal %"
picresults.Print
picresults.Print "Player"; Tab(45); "Field Goal %"
picresults.Print "************************************************************************************************************"
'This sorts the array and displays the results in the picture box by highest Field goal %
I = 13
    For Pass = 1 To I - 1
        For Comp = 1 To I - Pass
            If FieldGoals(Comp) < FieldGoals(Comp + 1) Then
            
                TempFieldGoals = FieldGoals(Comp)
                FieldGoals(Comp) = FieldGoals(Comp + 1)
                FieldGoals(Comp + 1) = TempFieldGoals
                
                TempNames = Names(Comp)
                Names(Comp) = Names(Comp + 1)
                Names(Comp + 1) = TempNames
                
            End If
        Next Comp
    Next Pass
picresults.Print
'Prints Data that is desired for command
    For I = 1 To 13
        picresults.Print Names(I); Tab(45); FormatPercent(FieldGoals(I))
    Next I
End Sub





Private Sub cmdMinutes_Click()
'This command will sort Timberwolves statistics from most MPG to least MPG
'Open "N:\CS130\handin\Lewis, Ben\twolves.txt" For Input As #1
Open PATH & "twolves.txt" For Input As #1
    For I = 1 To 13
        Input #1, Names(I), Numbers(I), MinutesPerGame(I), FieldGoals(I), ReboundsPerGame(I), AssistsPerGame(I), PointsPerGame(I)
    Next I
Close
'This will clear Picture Box of any previous data
picresults.Cls
'Prints Header
picresults.Print "Statistics Sorted by Most Minutes Played Per Game"
picresults.Print
picresults.Print "Player"; Tab(41); "Minutes/Game"
picresults.Print "************************************************************************************************************"
'This sorts the array and displays the results in the picture box by Minutes played from largest to smallest
I = 13
    For Pass = 1 To I - 1
        For Comp = 1 To I - Pass
            If MinutesPerGame(Comp) < MinutesPerGame(Comp + 1) Then
                
                TempMinutesPerGame = MinutesPerGame(Comp)
                MinutesPerGame(Comp) = MinutesPerGame(Comp + 1)
                MinutesPerGame(Comp + 1) = TempMinutesPerGame
                
                TempNames = Names(Comp)
                Names(Comp) = Names(Comp + 1)
                Names(Comp + 1) = TempNames
                
            End If
        Next Comp
    Next Pass
picresults.Print
'Prints Data that is desired for command
    For I = 1 To 13
        picresults.Print Names(I); Tab(45); FormatNumber(MinutesPerGame(I))
    Next I
            
End Sub





Private Sub cmdPoints_Click()
'This command will sort Timberwolves Statistics from Most Points Per Game to Least Points Per Game
'Open "N:\CS130\handin\Lewis, Ben\twolves.txt" For Input As #1
Open PATH & "twolves.txt" For Input As #1
    For I = 1 To 13
        Input #1, Names(I), Numbers(I), MinutesPerGame(I), FieldGoals(I), ReboundsPerGame(I), AssistsPerGame(I), PointsPerGame(I)
    Next I
Close
'This will clear the Picture box of any previous data
picresults.Cls
'Prints Header
picresults.Print "Statistics Sorted by Most Points Per Game"
picresults.Print
picresults.Print "Player"; Tab(41); "Points/Game"
picresults.Print "************************************************************************************************************"
'This sorts the array and displays the results in the picture box by Most Points Per Game to Least Points Per Game
I = 13
    For Pass = 1 To I - 1
        For Comp = 1 To I - Pass
            If PointsPerGame(Comp) < PointsPerGame(Comp + 1) Then
            
                TempPointsPerGame = PointsPerGame(Comp)
                PointsPerGame(Comp) = PointsPerGame(Comp + 1)
                PointsPerGame(Comp + 1) = TempPointsPerGame
                
                TempNames = Names(Comp)
                Names(Comp) = Names(Comp + 1)
                Names(Comp + 1) = TempNames
            
            End If
        Next Comp
    Next Pass
picresults.Print
'Prints Data that is desired for command
    For I = 1 To 13
        picresults.Print Names(I); Tab(45); FormatNumber(PointsPerGame(I))
    Next I
    
End Sub






Private Sub cmdPointsPerMin_Click()
'This command will sort Timberwolves Statistics from Most Points Per Minute Per Game to Least Points Per Minute Per Game
'Open "N:\CS130\handin\Lewis, Ben\twolves.txt" For Input As #1
Open PATH & "twolves.txt" For Input As #1
    For I = 1 To 13
        Input #1, Names(I), Numbers(I), MinutesPerGame(I), FieldGoals(I), ReboundsPerGame(I), AssistsPerGame(I), PointsPerGame(I)
    Next I
Close
'This will clear the Picture box of any previous data
picresults.Cls
'Prints Header
picresults.Print "Statistics Sorted by Most Points Per Minute Per Game"
picresults.Print
picresults.Print "Player"; Tab(41); "Points/Minute/Game"
picresults.Print "************************************************************************************************************"
'This sorts the array and displays the results in the picture box by Most Points Per Game to Least Points Per Game
    For I = 1 To 13
        PointsPerMinute(I) = PointsPerGame(I) / MinutesPerGame(I)
    Next I
I = 13
    For Pass = 1 To I - 1
        For Comp = 1 To I - Pass
            If PointsPerMinute(Comp) < PointsPerMinute(Comp + 1) Then
            
                TempPointsPerMinute = PointsPerMinute(Comp)
                PointsPerMinute(Comp) = PointsPerMinute(Comp + 1)
                PointsPerMinute(Comp + 1) = TempPointsPerMinute
                
                TempNames = Names(Comp)
                Names(Comp) = Names(Comp + 1)
                Names(Comp + 1) = TempNames
            
            End If
        Next Comp
    Next Pass
picresults.Print
'Prints Data that is desired for command
    For I = 1 To 13
        picresults.Print Names(I); Tab(45); FormatNumber(PointsPerMinute(I))
    Next I
            
End Sub





Private Sub cmdPrevious_Click()
'Enables User to return to previous screen
ProjectFace2.Hide
ProjectFace1.Show
cmdMinutes.Enabled = False
cmdFieldGoal.Enabled = False
cmdRebounds.Enabled = False
cmdAssists.Enabled = False
cmdPoints.Enabled = False
cmdPointsPerMin.Enabled = False
cmdFavorite.Enabled = False
End Sub



Private Sub cmdQuit_Click()
'Enables User to Stop Program
End
End Sub





Private Sub cmdRebounds_Click()
'This Command will sort Timberwolves Statistics from Most Reb. Per Game to Least Reb. Per Game
'Open "N:\CS130\handin\Lewis, Ben\twolves.txt" For Input As #1
Open PATH & "twolves.txt" For Input As #1
    For I = 1 To 13
        Input #1, Names(I), Numbers(I), MinutesPerGame(I), FieldGoals(I), ReboundsPerGame(I), AssistsPerGame(I), PointsPerGame(I)
    Next I
Close
'This will clear Picture Box of any previous data
picresults.Cls
'Prints Header
picresults.Print "Statistics Sorted by Most Rebounds Per Game"
picresults.Print
picresults.Print "Player"; Tab(41); "Rebounds/Game"
picresults.Print "************************************************************************************************************"
'This sorts the array and displays the results in the picture box by Rebounds/Game from largest to smallest
I = 13
    For Pass = 1 To I - 1
        For Comp = 1 To I - Pass
            If ReboundsPerGame(Comp) < ReboundsPerGame(Comp + 1) Then
                
                TempReboundsPerGame = ReboundsPerGame(Comp)
                ReboundsPerGame(Comp) = ReboundsPerGame(Comp + 1)
                ReboundsPerGame(Comp + 1) = TempReboundsPerGame
                
                TempNames = Names(Comp)
                Names(Comp) = Names(Comp + 1)
                Names(Comp + 1) = TempNames
                
            End If
        Next Comp
    Next Pass
picresults.Print
'Prints Data that is desired for command
    For I = 1 To 13
         picresults.Print Names(I); Tab(45); FormatNumber(ReboundsPerGame(I))
    Next I
    

End Sub





Private Sub cmdShow_Click()
'This command will load all of the individual statistics and display them
'Clears picture box of any data
picresults.Cls
'Prints Titles for data
picresults.Print " Player"; Tab(27); "#"; Tab(34); "MPG"; Tab(49); "FG%"; Tab(59); "RPG"; Tab(71); "APG"; Tab(81); "PPG"
picresults.Print "************************************************************************************************************"
'Open "N:\CS130\handin\Lewis, Ben\twolves.txt" For Input As #1
Open PATH & "twolves.txt" For Input As #1
    For I = 1 To 13
        Input #1, Names(I), Numbers(I), MinutesPerGame(I), FieldGoals(I), ReboundsPerGame(I), AssistsPerGame(I), PointsPerGame(I)
        picresults.Print Tab(1); Names(I); Tab(26); Numbers(I); Tab(34); FormatNumber(MinutesPerGame(I)); Tab(47); FormatPercent(FieldGoals(I)); Tab(59); FormatNumber(ReboundsPerGame(I)); Tab(71); FormatNumber(AssistsPerGame(I)); Tab(81); FormatNumber(PointsPerGame(I))
    Next I
Close
'Allows buttons that were unavailable to now be available
cmdMinutes.Enabled = True
cmdFieldGoal.Enabled = True
cmdRebounds.Enabled = True
cmdAssists.Enabled = True
cmdPoints.Enabled = True
cmdPointsPerMin.Enabled = True
cmdFavorite.Enabled = True
End Sub


