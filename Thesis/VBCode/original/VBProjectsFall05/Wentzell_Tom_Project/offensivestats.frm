VERSION 5.00
Begin VB.Form frmoffense 
   BackColor       =   &H00800000&
   Caption         =   "Offensive Statistics--2005 Houston Astros"
   ClientHeight    =   7560
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9645
   FillColor       =   &H00800000&
   LinkTopic       =   "Form2"
   ScaleHeight     =   7560
   ScaleWidth      =   9645
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdcalc 
      Caption         =   "Calculate Your Own Batting Average!"
      Height          =   1095
      Left            =   5280
      TabIndex        =   13
      Top             =   5880
      Width           =   1935
   End
   Begin VB.PictureBox picResults2 
      BackColor       =   &H000000FF&
      Height          =   1335
      Left            =   1800
      ScaleHeight     =   1275
      ScaleWidth      =   2955
      TabIndex        =   12
      Top             =   5760
      Width           =   3015
   End
   Begin VB.PictureBox picPlayer 
      BackColor       =   &H000000FF&
      Height          =   1335
      Left            =   480
      ScaleHeight     =   1275
      ScaleWidth      =   915
      TabIndex        =   10
      Top             =   5760
      Width           =   975
   End
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit"
      Height          =   615
      Left            =   8040
      TabIndex        =   9
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton cmdmain 
      Caption         =   "Return to Main Menu"
      Height          =   735
      Left            =   6600
      TabIndex        =   8
      Top             =   360
      Width           =   2415
   End
   Begin VB.CommandButton cmdpitching 
      Caption         =   "View Pitching Statistics"
      Height          =   735
      Left            =   3720
      TabIndex        =   7
      Top             =   360
      Width           =   2415
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H000000FF&
      Height          =   3975
      Left            =   3480
      ScaleHeight     =   3915
      ScaleWidth      =   5595
      TabIndex        =   6
      Top             =   1440
      Width           =   5655
   End
   Begin VB.CommandButton cmdhomeruns 
      Caption         =   "Sort Statistics by Home Runs (HR)"
      Height          =   615
      Left            =   360
      TabIndex        =   5
      Top             =   3120
      Width           =   2655
   End
   Begin VB.CommandButton cmdrbi 
      Caption         =   "Sort Statistics by Runs Batted In (RBI)"
      Height          =   615
      Left            =   360
      TabIndex        =   4
      Top             =   3960
      Width           =   2655
   End
   Begin VB.CommandButton cmdbattingaverage 
      Caption         =   "Sort Statistics by Batting Average (BA)"
      Height          =   615
      Left            =   360
      TabIndex        =   3
      Top             =   4800
      Width           =   2655
   End
   Begin VB.CommandButton cmdruns 
      Caption         =   "Sort Statistics by Runs Scored (R)"
      Height          =   615
      Left            =   360
      TabIndex        =   2
      Top             =   1440
      Width           =   2655
   End
   Begin VB.CommandButton cmdhits 
      Caption         =   "Sort Statistics by Hits (H)"
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   2280
      Width           =   2655
   End
   Begin VB.CommandButton cmdload 
      Caption         =   "Load 2005 Houston Astros Offensive Statistics"
      Height          =   1095
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label lbltom 
      BackColor       =   &H000000FF&
      Caption         =   "Tom Wentzell"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   7200
      Width           =   1095
   End
End
Attribute VB_Name = "frmoffense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'2005 Houston Astros Statistics(Wentzell_Tom_Project)
'frmoffense (frmoffense.frm)
'Tom Wentzell
'October 30, 2005
'The purpose of this form is to display individual offensive statistics for each player
'on the Houston Astros who recorded at least 100 at bats in 2005.  It can sort the players
'by production in different statistical categories, allowing the user to view the leaders
'of each statistical category, and it also gives the user an opportunity to calculate his
'or her own batting average.

'Declare form level variables.  All of these variables will be used in multiple subroutines.
Option Explicit
Dim Player(1 To 20) As String, Runs(1 To 20) As Integer, Hits(1 To 20) As Integer
Dim HomeRuns(1 To 20) As Integer, RBI(1 To 20) As Integer, Batting(1 To 20) As Double
Dim CTR As Integer, Pass As Integer, Comp As Integer, J As Integer, tempname As String

'This button loads offensive data from a data file into the program in the form of an
'array.  The data is given for 13 Astros players and it covers five statistical categories.
'This data will be displayed in its original form along with headings displaying
'the information  given in each column.
Private Sub cmdload_Click()
picResults.Cls
picResults.Print "Player"; Tab(26); "R"; Tab(33); "H"; Tab(41); "HR"; Tab(48); "RBI"; Tab(58); "BA"
picResults.Print "----------"; Tab(25); "-----"; Tab(32); "-----"; Tab(41); "-----"; Tab(48); "-----"; Tab(58); "-----"
picResults.Print ""
Open App.Path & "\Offensive Stats.txt" For Input As #1
CTR = 0
Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, Player(CTR), Runs(CTR), Hits(CTR), HomeRuns(CTR), RBI(CTR), Batting(CTR)
    picResults.Print Player(CTR); Tab(25); Runs(CTR); Tab(32); Hits(CTR); Tab(41); HomeRuns(CTR); Tab(48); RBI(CTR); Tab(58); FormatNumber(Batting(CTR), 3)
Loop
Close
End Sub


'This button uses a bubble sort to rearrange the players according to highest number of
'runs scored.  It then reprints the data in this new sequence.  It also displays the
'picture of the player who lead the team in this statistical category as well as the
'amount of runs he scored in leading the team.
Private Sub cmdruns_Click()
picResults.Cls
picResults.Print "Player"; Tab(26); "R"; Tab(33); "H"; Tab(41); "HR"; Tab(48); "RBI"; Tab(58); "BA"
picResults.Print "----------"; Tab(25); "-----"; Tab(32); "-----"; Tab(41); "-----"; Tab(48); "----"; Tab(58); "---"
picResults.Print ""
For Pass = 1 To CTR - 1
    For Comp = 1 To CTR - Pass
        If Runs(Comp) < Runs(Comp + 1) Then
            tempname = Runs(Comp)
            Runs(Comp) = Runs(Comp + 1)
            Runs(Comp + 1) = tempname
            tempname = Player(Comp)
            Player(Comp) = Player(Comp + 1)
            Player(Comp + 1) = tempname
            tempname = Hits(Comp)
            Hits(Comp) = Hits(Comp + 1)
            Hits(Comp + 1) = tempname
            tempname = HomeRuns(Comp)
            HomeRuns(Comp) = HomeRuns(Comp + 1)
            HomeRuns(Comp + 1) = tempname
            tempname = RBI(Comp)
            RBI(Comp) = RBI(Comp + 1)
            RBI(Comp + 1) = tempname
            tempname = Batting(Comp)
            Batting(Comp) = Batting(Comp + 1)
            Batting(Comp + 1) = tempname
        End If
    Next Comp
Next Pass

For J = 1 To CTR
    picResults.Print Player(J); Tab(25); Runs(J); Tab(32); Hits(J); Tab(41); HomeRuns(J); Tab(48); RBI(J); Tab(58); FormatNumber(Batting(J), 3)
Next J
picPlayer.Picture = LoadPicture(App.Path & "\Biggio.jpg")
picResults2.Cls
picResults2.Print "Craig Biggio lead the 2005 Houston"
picResults2.Print "Astros in runs scored with 94."
End Sub

'This button uses a bubble sort to rearrange the players according to highest number of
'hits.  It then reprints the data in this new sequence.  It also displays the
'picture of the player who lead the team in this statistical category as well as the
'amount of hits he recorded in leading the team.
Private Sub cmdhits_Click()
picResults.Cls
picResults.Print "Player"; Tab(26); "R"; Tab(33); "H"; Tab(41); "HR"; Tab(48); "RBI"; Tab(58); "BA"
picResults.Print "----------"; Tab(25); "-----"; Tab(32); "-----"; Tab(41); "-----"; Tab(48); "----"; Tab(58); "---"
picResults.Print ""
For Pass = 1 To CTR - 1
    For Comp = 1 To CTR - Pass
        If Hits(Comp) < Hits(Comp + 1) Then
            tempname = Hits(Comp)
            Hits(Comp) = Hits(Comp + 1)
            Hits(Comp + 1) = tempname
            tempname = Player(Comp)
            Player(Comp) = Player(Comp + 1)
            Player(Comp + 1) = tempname
            tempname = Runs(Comp)
            Runs(Comp) = Runs(Comp + 1)
            Runs(Comp + 1) = tempname
            tempname = HomeRuns(Comp)
            HomeRuns(Comp) = HomeRuns(Comp + 1)
            HomeRuns(Comp + 1) = tempname
            tempname = RBI(Comp)
            RBI(Comp) = RBI(Comp + 1)
            RBI(Comp + 1) = tempname
            tempname = Batting(Comp)
            Batting(Comp) = Batting(Comp + 1)
            Batting(Comp + 1) = tempname
        End If
    Next Comp
Next Pass

For J = 1 To CTR
    picResults.Print Player(J); Tab(25); Runs(J); Tab(32); Hits(J); Tab(41); HomeRuns(J); Tab(48); RBI(J); Tab(58); FormatNumber(Batting(J), 3)
Next J
picPlayer.Picture = LoadPicture(App.Path & "\Taveras.jpg")
picResults2.Cls
picResults2.Print "Willy Taveras lead the 2005 Houston"
picResults2.Print "Astros in total hits with 172."
End Sub

'This button uses a bubble sort to rearrange the players according to highest number of
'home runs hit.  It then reprints the data in this new sequence.  It also displays the
'picture of the player who lead the team in this statistical category as well as the
'amount of home runs he hit in leading the team.
Private Sub cmdhomeruns_Click()
picResults.Cls
picResults.Print "Player"; Tab(26); "R"; Tab(33); "H"; Tab(41); "HR"; Tab(48); "RBI"; Tab(58); "BA"
picResults.Print "----------"; Tab(25); "-----"; Tab(32); "-----"; Tab(41); "-----"; Tab(48); "----"; Tab(58); "---"
picResults.Print ""
For Pass = 1 To CTR - 1
    For Comp = 1 To CTR - Pass
        If HomeRuns(Comp) < HomeRuns(Comp + 1) Then
            tempname = HomeRuns(Comp)
            HomeRuns(Comp) = HomeRuns(Comp + 1)
            HomeRuns(Comp + 1) = tempname
            tempname = Player(Comp)
            Player(Comp) = Player(Comp + 1)
            Player(Comp + 1) = tempname
            tempname = Runs(Comp)
            Runs(Comp) = Runs(Comp + 1)
            Runs(Comp + 1) = tempname
            tempname = Hits(Comp)
            Hits(Comp) = Hits(Comp + 1)
            Hits(Comp + 1) = tempname
            tempname = RBI(Comp)
            RBI(Comp) = RBI(Comp + 1)
            RBI(Comp + 1) = tempname
            tempname = Batting(Comp)
            Batting(Comp) = Batting(Comp + 1)
            Batting(Comp + 1) = tempname
        End If
    Next Comp
Next Pass

For J = 1 To CTR
    picResults.Print Player(J); Tab(25); Runs(J); Tab(32); Hits(J); Tab(41); HomeRuns(J); Tab(48); RBI(J); Tab(58); FormatNumber(Batting(J), 3)
Next J
picPlayer.Picture = LoadPicture(App.Path & "\Ensberg.jpg")
picResults2.Cls
picResults2.Print "Morgan Ensberg lead the 2005 Houston"
picResults2.Print "Astros in home runs with 36."
End Sub

'This button uses a bubble sort to rearrange the players according to highest number of
'runs batted in.  It then reprints the data in this new sequence.  It also displays the
'picture of the player who lead the team in this statistical category as well as the
'amount of runs batted in he recorded in leading the team.
Private Sub cmdrbi_Click()
picResults.Cls
picResults.Print "Player"; Tab(26); "R"; Tab(33); "H"; Tab(41); "HR"; Tab(48); "RBI"; Tab(58); "BA"
picResults.Print "----------"; Tab(25); "-----"; Tab(32); "-----"; Tab(41); "-----"; Tab(48); "----"; Tab(58); "---"
picResults.Print ""
For Pass = 1 To CTR - 1
    For Comp = 1 To CTR - Pass
        If RBI(Comp) < RBI(Comp + 1) Then
            tempname = RBI(Comp)
            RBI(Comp) = RBI(Comp + 1)
            RBI(Comp + 1) = tempname
            tempname = Player(Comp)
            Player(Comp) = Player(Comp + 1)
            Player(Comp + 1) = tempname
            tempname = Runs(Comp)
            Runs(Comp) = Runs(Comp + 1)
            Runs(Comp + 1) = tempname
            tempname = Hits(Comp)
            Hits(Comp) = Hits(Comp + 1)
            Hits(Comp + 1) = tempname
            tempname = HomeRuns(Comp)
            HomeRuns(Comp) = HomeRuns(Comp + 1)
            HomeRuns(Comp + 1) = tempname
            tempname = Batting(Comp)
            Batting(Comp) = Batting(Comp + 1)
            Batting(Comp + 1) = tempname
        End If
    Next Comp
Next Pass

For J = 1 To CTR
    picResults.Print Player(J); Tab(25); Runs(J); Tab(32); Hits(J); Tab(41); HomeRuns(J); Tab(48); RBI(J); Tab(58); FormatNumber(Batting(J), 3)
Next J
picPlayer.Picture = LoadPicture(App.Path & "\Ensberg.jpg")
picResults2.Cls
picResults2.Print "Morgan Ensberg lead the 2005 Houston"
picResults2.Print "Astros in runs batted in with 101."
End Sub

'This button uses a bubble sort to rearrange the players according to highest batting
'average.  It then reprints the data in this new sequence.  It also displays the
'picture of the player who lead the team in this statistical category as well as his
'his final batting average in leading the team.
Private Sub cmdbattingaverage_Click()
picResults.Cls
picResults.Print "Player"; Tab(26); "R"; Tab(33); "H"; Tab(41); "HR"; Tab(48); "RBI"; Tab(58); "BA"
picResults.Print "----------"; Tab(25); "-----"; Tab(32); "-----"; Tab(41); "-----"; Tab(48); "----"; Tab(58); "---"
picResults.Print ""
For Pass = 1 To CTR - 1
    For Comp = 1 To CTR - Pass
        If Batting(Comp) < Batting(Comp + 1) Then
            tempname = Batting(Comp)
            Batting(Comp) = Batting(Comp + 1)
            Batting(Comp + 1) = tempname
            tempname = Player(Comp)
            Player(Comp) = Player(Comp + 1)
            Player(Comp + 1) = tempname
            tempname = Runs(Comp)
            Runs(Comp) = Runs(Comp + 1)
            Runs(Comp + 1) = tempname
            tempname = Hits(Comp)
            Hits(Comp) = Hits(Comp + 1)
            Hits(Comp + 1) = tempname
            tempname = HomeRuns(Comp)
            HomeRuns(Comp) = HomeRuns(Comp + 1)
            HomeRuns(Comp + 1) = tempname
            tempname = RBI(Comp)
            RBI(Comp) = RBI(Comp + 1)
            RBI(Comp + 1) = tempname
        End If
    Next Comp
Next Pass

For J = 1 To CTR
    picResults.Print Player(J); Tab(25); Runs(J); Tab(32); Hits(J); Tab(41); HomeRuns(J); Tab(48); RBI(J); Tab(58); FormatNumber(Batting(J), 3)
Next J
picPlayer.Picture = LoadPicture(App.Path & "\Berkman.jpg")
picResults2.Cls
picResults2.Print "Lance Berkman lead the 2005 Houston"
picResults2.Print "Astros in batting average at 0.293."
End Sub

'This button allows the user to figure out his or her own batting average through the
'use of an input box.  There are  conditional statements relating to the input box
'so that the data the user enters is applicable to batting average statistics.  It then
'displays the user's at bats, hits, and batting average.

Private Sub cmdcalc_Click()
Dim average As Double, Atbats As Integer, Userhits As Integer
Atbats = InputBox("Enter your total number of at bats:", "At Bats")

Do While Atbats <= 0
    MsgBox "Sorry, you must enter a positive number", , Error
    Atbats = InputBox("Enter your total number of at bats:", "At Bats")
Loop

Userhits = InputBox("Enter the amount of hits you got in those at bats:", "Hits")
    
Do While Userhits > Atbats Or Userhits < 0
    MsgBox "Your total number of hits cannot be a negative number.  It must also less than or equal to your total number of at bats.", , Error
    Userhits = InputBox("Enter the amount of hits you got in those at bats:", "Hits")
Loop

average = Userhits / Atbats
picPlayer.Picture = LoadPicture("")
picResults2.Cls
picResults.Cls

average = Userhits / Atbats
picResults2.Print "You hit safely"; Userhits; "times"
picResults2.Print "out of"; Atbats; "at bats"
picResults2.Print "for a batting average of "; FormatNumber(average, 3)

End Sub

'This command button directs the user to the pitching statistics form.  It also clears
'the picture of the currently displayed statistical leader so that upon returning to the
'offensive form at a later time, all picture boxes will be clear.
Private Sub cmdpitching_Click()
    frmpitching.Show
    frmoffense.Hide
    picPlayer.Picture = LoadPicture("")
End Sub

'This command button directs the user to the main menu.  It also clears
'the picture of the currently displayed statistical leader so that upon returning to the
'offensive form at a later time, all picture boxes will be clear.
Private Sub cmdmain_Click()
    frmmain.Show
    frmoffense.Hide
    picPlayer.Picture = LoadPicture("")
End Sub

'This command button allows the user to exit the program.
Private Sub cmdquit_Click()
End
End Sub

